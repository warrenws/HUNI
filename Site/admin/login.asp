<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
On Error Resume Next

Dim strRole, strSiteVersion, strSourcePage, strReturnLink

'Redirect the user the SSL version if required
If Application("ForceSSL") Then
	If Request.ServerVariables("SERVER_PORT")=80 Then
		If Request.ServerVariables("QUERY_STRING") = "" Then
			Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
		Else
			Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
		End If
	End If
End If

'Delete the old session if they are logging out.
If Request.QueryString("Action") = "logout" Then

	strUserName = GetUser

	strSQL = "DELETE FROM Sessions WHERE SessionID='" & Replace(Request.Cookies("SessionID"),"'","''") & "'"
	Application("Connection").Execute(strSQL)
	
	'Clear the cookies
	Response.Cookies("SessionID") = ""
	Response.Cookies("Role") = ""

	UpdateLog "UserLogout","",strUserName,"",Request.ServerVariables("REMOTE_ADDR"),""

	Response.Redirect("login.asp")
	
End If

'Get the information from the form
strUserName = Request.Form("UserName")
strPassword = Request.Form("Password")
strLogin = Request.Form("Login")
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
strIPAddress = Request.ServerVariables("REMOTE_ADDR")

'Remove all old sessions
DeleteOldSessions

'Build return string
If Request.ServerVariables("QUERY_STRING") = "" Then
	strReturnLink = ""
Else
	strReturnLink = "?" & Request.ServerVariables("QUERY_STRING")
End If

'If they are already logged in send them back
If Request.Cookies("SessionID") <> "" Then
	strSQL = "SELECT UserAgent FROM Sessions WHERE SessionID='" & Replace(Request.Cookies("SessionID"),"'","''") & "'"
	Set objActiveSession = Application("Connection").Execute(strSQL)
	
	If Not objActiveSession.EOF Then
	
		'This line fixes a redirect loop, the user would be sent back and fourth if the useragent changed since last time
		If Left(Replace(objActiveSession(0),"'","''"),250) = Left(Replace(strUserAgent,"'","''"),250) Then

			'Redirect the user to the page they came from, or to the default page
			strSourcePage = Request.QueryString("SourcePage")
			If InStr(strSourcePage,",") Then
				strSourcePage = Left(strSourcePage,InStr(strSourcePage,",") - 1)
			End If
			If strSourcePage = "" Then
				Response.Redirect("index.asp")
			Else
				Response.Redirect(strSourcePage & BuildReturnLink)
			End If
			
		Else
			
			'The user agent has changed since the last login, we're going to delete the old session
			strSQL = "DELETE FROM Sessions WHERE SessionID='" & Replace(Request.Cookies("SessionID"),"'","''") & "'"
			Application("Connection").Execute(strSQL)
			
		End If
		
	End If
End If

'If they hit the login button
If strLogin = "Login" Then

	'Make sure they entered a username
	If Trim(strUserName) = "" Then
	
		strMessage = "Username Missing"
		strMessageType = "Error"
	
	Else

		'Fix the username if it's an email address
		If InStr(strUserName,"@") Then
			strUserName = Left(strUserName,InStr(strUserName,"@") - 1)
		End If

		'Fix the username if it's in legacy form
		If InStr(strUserName,"\") Then
			strUserName = Right(strUserName,Len(strUserName) - InStr(strUserName,("\")))
		End If

		'Create objects required to connect to AD
		Set objConnection = CreateObject("ADODB.Connection")
		Set objCommand = CreateObject("ADODB.Command")
		Set objRootDSE = GetObject("LDAP://" & Application("Domain") & "/rootDSE")

		'Create a connection to AD
		objConnection.Provider = "ADSDSOObject"
		
		'Try to connect to Active Directory using the user name and password they provided
		objConnection.Open "Active Directory Provider",strUserName & "@" & Application("Domain"), strPassword
		objCommand.ActiveConnection = objConnection
		strDNSDomain = objRootDSE.Get("DefaultNamingContext")
		objCommand.CommandText = "<LDAP://" & Application("DomainController") & "/" & strDNSDomain & ">;(&(objectCategory=person)(objectClass=user)(samaccountname=" & strUserName & ")); GivenName,SN,name,memberOf ;subtree"

		'Initiate the LDAP query and return results to a RecordSet object.
		Set objRecordSet = objCommand.Execute
	
		'If the connection works then we have the correct username and password
		If Err.Number = 0 Then
		
			'Now that they have authenticated, see if they are authorized 
			bolAuthorized = False
			For Each strGroup in objRecordSet.Fields(3).Value
				If InStr(strGroup,Application("DomainGroupUsers")) Then
					bolAuthorized = True
					If strRole = "" Then
						strRole = "User"
					End If
				End If
				If InStr(strGroup,Application("DomainGroupAdmins")) Then
					bolAuthorized = True
					strRole = "Admin"
				End If
			Next
		
			'If they were a member of the right group let them in
			If bolAuthorized Then
			
				'See if they created a session in the past 10 seconds.
				strSQL = "SELECT SessionID,LoginTime FROM Sessions WHERE Username='" & strUserName & "' AND UserAgent='" & strUserAgent &"' And LoginDate=Date()"
				Set objActiveSession = Application("Connection").Execute(strSQL)
				
				If objActiveSession.EOF Then
					CreateNewSession
					SendEmail strUserName
				Else
					intSeconds = DateDiff("s",objActiveSession(1),Time())
					If intSeconds > 10 Then
						CreateNewSession
						SendEmail strUserName
					Else
						'Redirect the user to the page they came from, or to the default page
						strSourcePage = Request.QueryString("SourcePage")
						If InStr(strSourcePage,",") Then
							strSourcePage = Left(strSourcePage,InStr(strSourcePage,",") - 1)
						End If
						If strSourcePage = "" Then
							Response.Redirect("index.asp")
						Else
							Response.Redirect(strSourcePage & BuildReturnLink)
						End If
					End If
				End If
				
				'Redirect the user to the page they came from, or to the default page
				strSourcePage = Request.QueryString("SourcePage")
				If InStr(strSourcePage,",") Then
					strSourcePage = Left(strSourcePage,InStr(strSourcePage,",") - 1)
				End If
				If strSourcePage = "" Then
					Response.Redirect("index.asp")
				Else
					Response.Redirect(strSourcePage & BuildReturnLink)
				End If
			
			Else
				strMessage = "Access Denied"
				strMessageType = "Error"
			End If
			
		Else
			Err.Clear
			strMessage = "Incorrect Password"
			strMessageType = "Error"
		End If
		
	End If
	
End If

'Get the User Agent from the client so we know what browser they are using
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

intControlSize = 30

If InStr(strUserAgent,"iPhone") Then
	intControlSize = 25
End If

If InStr(strUserAgent,"Android") Then
	intControlSize = 15
End if

If InStr(strUserAgent,"Windows Phone") Then
	intControlSize = 15
End If

If InStr(strUserAgent,"CrOS") Then
	intControlSize = 20
End If
	
	If InStr(strUserAgent,"Windows") Then
		intControlSize = 20
	End If

If IsMobile Then
	strSiteVersion = "Mobile"
Else
	strSiteVersion = "Full"
End If
%>

<html>

<head>
	<title><%=Application("SiteName")%></title>
		<link rel="stylesheet" type="text/css" href="../style.css" /> 
		<link rel="apple-touch-icon" href="../images/inventory.png" /> 
		<link rel="shortcut icon" href="../images/inventory.ico" />
		<meta name="viewport" content="width=device-width,user-scalable=0" /> 
		<meta name="theme-color" content="#333333">
		<link rel="stylesheet" href="../assets/css/jquery-ui.css">
		<script src="../assets/js/jquery.js"></script>
		<script src="../assets/js/jquery-ui.js"></script>
		<script type="text/javascript">	
			$(document).ready( function () {
			<%	If Not IsMobile And Not IsiPad Then%>
					$( document ).tooltip({track: true});
			<%	End If %>	
			}
		</script>
		
</head>

<body class="<%=strSiteVersion%>">
	<div class="Header"><%=Application("SiteName")%></div> 
	<br />
	<br /> 
	<div class="Card NormalCard">
		<form method="post" action="login.asp<%=strReturnLink%>">
		<div class="CardTitle">Enter your username and password</div>
		<div Class="CardColumn1">Username </div>
		<div Class="CardColumn2"><input class="Card" type="text" name="Username" /></div>
		<div Class="CardColumn1">Password </div>
		<div Class="CardColumn2"><input class="Card" type="password" name="Password" /></div>
		<div class="Button"><input type="submit" value="Login" name="Login" /></div>
		<div class="<%=strMessageType%>"><%=strMessage%></div>
		</form>
	</div> 
	<div class="Version">Version <%=Application("Version")%></div> 
</body>

</html>

<%
Sub CreateNewSession

	intSessionID = GenerateSessionID
	strSQL = "INSERT INTO Sessions " & _
	"(Username,SessionID,IPAddress,UserAgent,LoginDate,LoginTime,ExpirationDate,Role) VALUES " & _
	"('" & strUserName & "','" & intSessionID & "','" & strIPAddress & "','" & _
	Left(Replace(strUserAgent,"'","''"),250) & "',Date(),Time(),#" & Date() + Application("LogInDays") & "#,'" & strRole & "')"
	Application("Connection").Execute(strSQL)
	Response.Cookies("SessionID") = intSessionID
	Response.Cookies("SessionID").Expires = Date() + Application("LogInDays")
	Response.Cookies("Role") = strRole 
	
	UpdateLog "UserLoginAdmin","",strUserName,"",strIPAddress,""
	
End Sub 
%>

<%
Sub DeleteOldSessions
	
	Dim objOldSessions
	
	strSQL = "SELECT Username,IPAddress FROM Sessions WHERE DATE() >= ExpirationDate"
	Set objOldSessions = Application("Connection").Execute(strSQL)
	
	If Not objOldSessions.EOF Then
		Do Until objOldSessions.EOF
			UpdateLogAuto "AutoLogOut","",objOldSessions(0),"",objOldSessions(1),""
			objOldSessions.MoveNext
		Loop
	End If
	
	strSQL = "DELETE FROM Sessions WHERE Date() >= ExpirationDate"
	Application("Connection").Execute(strSQL)

End Sub
%>

<%
Function GenerateSessionID
	
	'Get a random number 
	GenerateSessionID = GetRandomNumber(1000000000,9999999999)

	'See if it's already in use in the database
	strSQL = "SELECT ID FROM Sessions WHERE SessionID ='" & GenerateSessionID & "'"
	Set objSessionCheck = Application("Connection").Execute(strSQL)
	If Not objSessionCheck.EOF Then
		GenerateSessionID = GenerateSessionID()
	End If
	
End Function
%>

<%
Function GetRandomNumber(intLow,intHigh)
	Randomize
	GetRandomNumber = (Int(RND * (intHigh - intLow + 1))) + intLow
End Function
%>

<%Function BuildReturnLink

	BuildReturnLink = "?" & Request.ServerVariables("QUERY_STRING")

End Function%>

<%Sub UpdateLog(EntryType,DeviceTag,UserName,OldValue,NewValue,EventNumber)

	Dim strType, strOldNotes, strNewNotes, strOldValue, strNewValue, intTag, strUserName, datDate, datTime, strSQL, intEventNumber

	'Get the type
	If EntryType <> "" Then
		strType = EntryType
	Else 
		Exit Sub
	End If
	
	'If a notes field was updated then the data needs to be stored in the notes field of the log
	If InStr(strType,"Notes") > 0 Then
		strOldNotes = OldValue
		strNewNotes = NewValue
		strOldValue = ""
		strNewValue = ""
	Else
		strOldNotes = ""
		strNewNotes = ""
		strOldValue = OldValue
		strNewValue = NewValue
	End If
	
	If EventNumber = "" Then
		intEventNumber = 0
	Else
		intEventNumber = EventNumber
	End If
	
	'Get the other things needed for the log
	intTag = DeviceTag
	strUserName = UserName
	datDate = Date()
	datTime = Time()
	
	strSQL = "INSERT INTO Log (LGTag,UserName,Type,OldValue,NewValue,OldNotes,NewNotes,UpdatedBy,LogDate,LogTime,Active,Deleted,EventNumber)" & vbCRLF
	strSQL = strSQL & "VALUES ('"
	strSQL = strSQL & intTag & "','"
	strSQL = strSQL & Replace(strUserName,"'","''") & "','"
	strSQL = strSQL & Replace(strType,"'","''") & "','"
	strSQL = strSQL & Replace(strOldValue,"'","''") & "','"
	strSQL = strSQL & Replace(strNewValue,"'","''") & "','"
	strSQL = strSQL & Replace(strOldNotes,"'","''") & "','"
	strSQL = strSQL & Replace(strNewNotes,"'","''") & "','"
	strSQL = strSQL & Replace(strUserName,"'","''") & "',#"
	strSQL = strSQL & datDate & "#,#"
	strSQL = strSQL & datTime & "#,True,False," & intEventNumber & ")"
	Application("Connection").Execute(strSQL)
	
End Sub%>

<%Sub UpdateLogAuto(EntryType,DeviceTag,UserName,OldValue,NewValue,EventNumber)

	Dim strType, strOldNotes, strNewNotes, strOldValue, strNewValue, intTag, strUserName, datDate, datTime, strSQL, intEventNumber

	'Get the type
	If EntryType <> "" Then
		strType = EntryType
	Else 
		Exit Sub
	End If
	
	'If a notes field was updated then the data needs to be stored in the notes field of the log
	If InStr(strType,"Notes") > 0 Then
		strOldNotes = OldValue
		strNewNotes = NewValue
		strOldValue = ""
		strNewValue = ""
	Else
		strOldNotes = ""
		strNewNotes = ""
		strOldValue = OldValue
		strNewValue = NewValue
	End If
	
	If EventNumber = "" Then
		intEventNumber = 0
	Else
		intEventNumber = EventNumber
	End If
	
	'Get the other things needed for the log
	intTag = DeviceTag
	strUserName = UserName
	datDate = Date()
	datTime = Time()
	
	strSQL = "INSERT INTO Log (LGTag,UserName,Type,OldValue,NewValue,OldNotes,NewNotes,UpdatedBy,LogDate,LogTime,Active,Deleted,EventNumber)" & vbCRLF
	strSQL = strSQL & "VALUES ('"
	strSQL = strSQL & intTag & "','"
	strSQL = strSQL & Replace(strUserName,"'","''") & "','"
	strSQL = strSQL & Replace(strType,"'","''") & "','"
	strSQL = strSQL & Replace(strOldValue,"'","''") & "','"
	strSQL = strSQL & Replace(strNewValue,"'","''") & "','"
	strSQL = strSQL & Replace(strOldNotes,"'","''") & "','"
	strSQL = strSQL & Replace(strNewNotes,"'","''") & "','"
	strSQL = strSQL & Replace("Automated","'","''") & "',#"
	strSQL = strSQL & datDate & "#,#"
	strSQL = strSQL & datTime & "#,True,False," & intEventNumber & ")"
	Application("Connection").Execute(strSQL)
	
End Sub%>

<%Sub SendEmail(strUserName)

	'This will send out an email

	Dim strSMTPPickupFolder, objMessage, objConf, strMessage, strSubject, bolHTMLMEssage
	Dim bolSendAsAdmin

	Const cdoSendUsingPickup = 1

	strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"

	'Create the objects required to send the mail.
	Set objMessage = CreateObject("CDO.Message")
	Set objConf = objMessage.Configuration
	With objConf.Fields
		.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
		.item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
		.Update
	End With
		
	'Send the message
	objMessage.From = "inventory@lkgeorge.org"
	objMessage.To = "hullm@lkgeorge.org"
	objMessage.Subject = "Inventory Site Login"
	objMessage.TextBody = strUserName & " just logged into the inventory site."
	objMessage.Send
	
	'Close objects
	Set objMessage = Nothing
	Set objConf = Nothing
	
End Sub%>

<%
Function IsMobile

	'Get the User Agent from the client so we know what browser they are using
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Check the user agent for signs they are on a mobile browser
	If InStr(strUserAgent,"iPhone") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"iPad") Then
		IsMobile = False
	ElseIf InStr(strUserAgent,"Android") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"Windows Phone") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"BlackBerry") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"Nintendo") Then
		IsMobile = True 
	ElseIf InStr(strUserAgent,"PlayStation Vita") Then
		IsMobile = True
	Else
		IsMobile = False
	End If 
	
	If InStr(strUserAgent,"Nexus 9") Then  
		IsMobile = False
	End If
	
End Function 
%>

<%Function IsiPad

	Dim strUserAgent

	'Get the User Agent from the client so we know what browser they are using
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Check the user agent for signs they are on a mobile browser
	If InStr(strUserAgent,"iPad") Then
		IsiPad = True
	Else
		IsiPad = False
	End If
	
End Function %>

<%Function GetUser

	Const USERNAME = 1

	Dim strUserAgent, strSessionID, objSessionLookup, strSQL
	
	'Get some needed data
	strSessionID = Request.Cookies("SessionID")
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
	
	'Send them to the logon screen if they don't have a Session ID
	If strSessionID = "" Then
		GetUser = ""

	'Get the username from the database
	Else
	
		strSQL = "SELECT ID,UserName,SessionID,IPAddress,UserAgent,ExpirationDate FROM Sessions "
		strSQL = strSQL & "WHERE UserAgent='" & Left(Replace(strUserAgent,"'","''"),250) & "' And SessionID='" & Replace(strSessionID,"'","''") & "'"
		strSQL = strSQL & " And ExpirationDate > Date()"
		Set objSessionLookup = Application("Connection").Execute(strSQL)
		
		'If a session isn't found for then kick them out
		If objSessionLookup.EOF Then
			GetUser = ""
		Else
			GetUser = objSessionLookup(USERNAME)
		End If
	End If  
	
End Function%>