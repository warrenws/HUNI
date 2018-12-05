<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/16/14
'Last Updated 1/14/18

'This is the main admin page for the inventory website

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim intTag, strSerial, strDeviceMessage, objGuideRooms, intBOCESTag, objClassOf, objSites
Dim strUserMessage, strSite, strAUPYes, strAUPNo, strSites, objFirstNames, objLastNames

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

   Dim strSQL, intYear

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Lookup Device"
         LookupDevice
      Case "Lookup User"
         LookupUser
   End Select
   
   'Get the data for the sites drop down menu
   strSQL = "SELECT Site FROM Sites WHERE Active=True ORDER BY Site"
   Set objSites = Application("Connection").Execute(strSQL)
   
   'Get the data for the guide rooms drop down menu
   strSQL = "SELECT DISTINCT HomeRoom FROM People WHERE Homeroom <> '' AND Active=True ORDER BY HomeRoom"
   Set objGuideRooms = Application("Connection").Execute(strSQL)
   
   'Get the data for the role drop down menu
   If Month(Date) >= 7 Then
   	intYear = Year(Date) + 1
   Else
   	intYear = Year(Date)
   End If
   strSQL = "SELECT DISTINCT ClassOf From People WHERE ClassOf >= " & intYear & " ORDER BY ClassOf DESC"
   Set objClassOf = Application("Connection").Execute(strSQL)
 
      
   'Display the error message if they were sent back to this page
   Select Case Request.QueryString("Error")
      Case "DeviceNotFound"
         strDeviceMessage = "<div Class=""Error"">Device not found</div>"
      Case "NoDevicesFound"
         strDeviceMessage = "<div Class=""Error"">No devices found</div>"
      Case "UserNotFound"
         strDeviceMessage = "<div Class=""Error"">User not found</div>"
      Case "NoUsersFound"
         strDeviceMessage = "<div Class=""Error"">No users found</div>"
   End Select
   
   'Get the list of firstnames for the auto complete
   strSQL = "SELECT DISTINCT FirstName FROM People WHERE Active=True AND Role='Student'"
   Set objFirstNames = Application("Connection").Execute(strSQL)
   
   'Get the list of lastnames for the auto complete
   strSQL = "SELECT DISTINCT LastName FROM People WHERE Active=True AND Role='Student'"
   Set objLastNames = Application("Connection").Execute(strSQL)
   
   'Set up the variables needed for the site then load it
   SetupSite
   DisplaySite
   
End Sub%>

<%Sub DisplaySite%>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title><%=Application("SiteName")%></title>
      <link rel="stylesheet" type="text/css" href="style.css" /> 
      <link rel="apple-touch-icon" href="images/inventory.png" /> 
      <link rel="shortcut icon" href="images/inventory.ico" />
      <meta name="viewport" content="width=device-width" />
      <meta name="theme-color" content="#333333">
      <link rel="stylesheet" href="assets/css/jquery-ui.css">
		<script src="assets/js/jquery.js"></script>
		<script src="assets/js/jquery-ui.js"></script>
		<script>
			$(document).ready(function () { 
			
			<% If Not IsMobile Then %>
					$( document ).tooltip({track: true});
			<% End If %>
			
			<% If Not objFirstNames.EOF Then %>
				var possibleFirstNames = [
			<% Do Until objFirstNames.EOF %>	
					"<%=objFirstNames(0)%>",
				<%	objFirstNames.MoveNext
				Loop %>
			];
				$( "#FirstNames" ).autocomplete({
					source: possibleFirstNames
				});
		<% End If %>
		
		<% If Not objLastNames.EOF Then %>
				var possibleLastNames = [
			<% Do Until objLastNames.EOF %>	
					"<%=objLastNames(0)%>",
				<%	objLastNames.MoveNext
				Loop %>
			];
				$( "#LastNames" ).autocomplete({
					source: possibleLastNames
				});
		<% End If %>
			
			})
		</script>
   </head>

   <body class="<%=strSiteVersion%>">
   
      <div class="Header"><%=Application("SiteName")%></div>
      <div>
         <ul class="NavBar" align="center">
            <li><a href="index.asp"><img src="images/home.png" title="Home" height="32" width="32"/></a></li>
            <li><a href="log.asp"><img src="images/log.png" title="System Log" height="32" width="32"/></a></li>
            <li><a href="login.asp?action=logout"><img src="images/logout.png" title="Log Out" height="32" width="32"/></a></li>
         </ul>
      </div>   

		<%
		UserSearchCard
		DeviceSearchCard
		%>
      
      <div class="Version">Version <%=Application("Version")%></div>
   </body>
   </html>

<%End Sub%>

<%Sub UserSearchCard%>

	<div class="Card NormalCard">
		<form method="POST" action="index.asp">
		<div class="CardTitle">Search for a User</div>
		<div>
			<div Class="CardColumn1">First name: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="FirstName" id="FirstNames" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Last name: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="LastName" id="LastNames" />
			</div>
		</div>
		
		<div>
			<div Class="CardColumn1">Site: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="Site">
					<option value=""></option>
				<% Do Until objSites.EOF %>
							<option value="<%=objSites(0)%>"><%=objSites(0)%></option>
				<%	 objSites.MoveNext
					Loop %>
				</select>
			</div>
		</div>
		
		
		<div>
			<div Class="CardColumn1">Grade: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="Role">
					<option value=""></option>
				<% Do Until objClassOf.EOF %>
							<option value="<%=objClassOf(0)%>"><%=GetRole(objClassOf(0))%></option>
				<%    objClassOf.MoveNext
					Loop %>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Guide room: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="GuideRoom">
					<option value=""></option>
					<% Do Until objGuideRooms.EOF %>
							<option value="<%=objGuideRooms(0)%>"><%=objGuideRooms(0)%></option>
					<%    objGuideRooms.MoveNext
						Loop %>
				</select>
			</div>
		</div>
	<% If Application("ShowPasswords") Then %>   
		<div>
			<div Class="CardColumn1">AUP: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="AUP">
					<option></option>
					<option value="Yes" <%=strAUPYes%>>Yes</option>
					<option value="No" <%=strAUPNo%>>No</option>
				</select>
			</div>
		</div>
	<% End If %>
		<div>
			<%=strUserMessage%>
			<div class="Button"><input type="submit" value="Lookup User" name="Submit" /></div>
		</div>
		</form>
	</div>  
      
<%End Sub %>

<%Sub DeviceSearchCard%>

	<div class="Card NormalCard">
		<form method="POST" action="index.asp">
		<div class="CardTitle">Search for a Device</div>
		<div>
			<div Class="CardColumn1">Asset tag: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthSmall" type="text" name="Tag" value="<%=intTag%>" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">BOCES tag: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthSmall" type="text" name="BOCESTag" value="<%=intBOCESTag%>" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Serial #: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="Serial" value="<%=strSerial%>"/>
			</div>
		</div>
		<div>
			<%=strDeviceMessage%>
			<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		</div>
		</form>
	</div>
        
<%End Sub%> 

<%Sub LookupDevice

   Dim strSQLWhere, strSQL, objDeviceLookup, intDeviceCount, strURL

   If Application("UseLeadingZeros") Then
		intTag = Request.Form("Tag")
	Else
		If IsNumeric(Request.Form("Tag")) Then
			intTag = Int(Request.Form("Tag"))
		Else
			intTag = Request.Form("Tag")
		End If
	End If
   
   If IsNumeric(Request.Form("BOCESTag")) Then
      intBOCESTag = Int(Request.Form("BOCESTag"))
   Else
      intBOCESTag = Request.Form("BOCESTag")
   End If
   
   strSerial = Request.Form("Serial")
   strSite = Request.Form("Site")
   intDeviceCount = 0
   
   If intTag = "" And intBOCESTag = "" And strSerial = "" And strSite = "" Then
      strDeviceMessage = "<div Class=""Error"">Device not found</div>"
   
   Else
   
      strSQLWhere = "WHERE "
      If intTag <> "" Then
         strSQLWhere = strSQLWhere & "LGTag='" & intTag & "' AND "
      End If
      
      If intBOCESTag <> "" Then
         strSQLWhere = strSQLWhere & "BOCESTag='" & intBOCESTag & "' AND "
      End If
   
      If strSerial <> "" Then
         strSQLWhere = strSQLWhere & "SerialNumber='" & strSerial & "' AND "
      End If
      
      If strSite <> "" Then
         strSQLWhere = strSQLWhere & "Site='" & strSite & "' AND "
         strURL = strURL & "&Site=" & Replace(strSite," ","%20")
      End If

      strSQLWhere = Left(strSQLWhere,Len(strSQLWhere) - 5)
   
      strSQL = "SELECT ID,LGTag FROM Devices " & strSQLWhere
   
      Set objDeviceLookup = Application("Connection").Execute(strSQL)

      If Not objDeviceLookup.EOF Then
         Do Until objDeviceLookup.EOF
            intDeviceCount = intDeviceCount + 1
            objDeviceLookup.MoveNext
         Loop
         objDeviceLookup.MoveFirst
      End If
      
      If strURL <> "" Then
         strURL = Right(strURL, Len(strURL) - 1)
      End If
      
      Select Case intDeviceCount
         Case 0         
            strDeviceMessage = "<div Class=""Error"">Device not found</div>"
         Case 1
            Response.Redirect("device.asp?Tag=" & objDeviceLookup(1))
         'Case Else
          '  Response.Redirect("devices.asp?" & strURL)
      End Select
   End If
   
End Sub%>

<%Sub LookupUser

   Dim strFirstName, strLastName, strGuideRoom, strSQL, objUserLookup, intUserCount, strURL, strRole
   Dim strAUP
   
   Select Case Request.Form("AUP")
   	Case "Yes"
   		strAUPYes = "selected=""selected"""
   		strAUP = "True"
   	Case "No"
   		strAUPNo = "selected=""selected"""
   		strAUP = "False"
   End Select
   
   strFirstName = Request.Form("FirstName")
   strLastName = Request.Form("LastName")
   strGuideRoom = Request.Form("GuideRoom")
   strRole = Request.Form("Role")
   strSite = Request.Form("Site")
   intUserCount = 0
   
   If strFirstName = "" And strLastName = "" And strGuideRoom = "" And strRole = "" And strAUP = "" And strSite = "" Then
      strUserMessage = "<div Class=""Error"">User not found</div>"
   Else
   
      strSQL = "SELECT ID, UserName FROM People WHERE ClassOf > 2000 AND "
      If strFirstName <> "" Then
         strSQL = strSQL & "FirstName Like '%" & Replace(strFirstName,"'","''") & "%' AND "
         strURL = strURL & "&FirstName=" & strFirstName
      End If
      If strLastName <> "" Then
         strSQL = strSQL & "LastName Like '%" & Replace(strLastName,"'","''") & "%' AND "
         strURL = strURL & "&LastName=" & strLastName
      End If
      If strSite <> "" Then
         strSQL = strSQL & "Site='" & Replace(strSite,"'","''") & "' AND "
         strURL = strURL & "&UserSite=" & Replace(strSite," ","%20")

      End If
      If strGuideRoom <> "" Then
         strSQL = strSQL & "HomeRoom Like '%" & Replace(strGuideRoom,"'","''") & "%' AND "
         strURL = strURL & "&GuideRoom=" & strGuideRoom
      End If
      If strAUP <> "" Then
      	strSQL = strSQL & "AUP=" & strAUP & " AND "
      	strURL = strURL & "&AUP=" & strAUP
      End If
      If strRole <> "" Then
         strSQL = strSQL & "ClassOf=" & strRole & "     "
         strURL = strURL & "&Role=" & strRole
      End If
      strSQL = Left(strSQL,Len(strSQL) - 5) & " AND Active=True"
      Response.Write("*" & strSQL)
      Set objUserLookup = Application("Connection").Execute(strSQL)
   
      If Not objUserLookup.EOF Then
         Do Until objUserLookup.EOF
            intUserCount = intUserCount + 1
            objUserLookup.MoveNext
         Loop
         objUserLookup.MoveFirst
      End If 
       
      If strURL <> "" Then
         strURL = Right(strURL, Len(strURL) - 1)
      End If
   
      Select Case intUserCount
         Case 0 
            strUserMessage = "<div Class=""Error"">User not found</div>"
         'Case 1
         '   Response.Redirect("user.asp?UserName=" & objUserLookup(1))
         Case Else
            Response.Redirect("users.asp?" & strURL)
      End Select
      
   End If
   
End Sub%>

<%Function GetRole(intYear)

   Dim datToday, intMonth, intCurrentYear, intGrade
   
   'Set the role to the correct staff group
   Select Case intYear
      Case 10
         GetRole = "Tech Staff"   
      Case 20
         GetRole = "Teacher"
      Case 30
         GetRole = "Staff"
      Case 40
         GetRole = "TA"
      Case 50
         GetRole = "Sub"
      Case 60
         GetRole = "Student Teacher"
   End Select 
      
   'Convert the graduating year to a grade
   datToday = Date
   intMonth = DatePart("m",datToday)
   intCurrentYear = Right(DatePart("yyyy",datToday),2)
   intYear = Right(intYear,2)

   If intMonth >= 7 And intMonth <= 12 Then
      intCurrentYear = intCurrentYear + 1
   End If

   intGrade = 12 - (intYear - intCurrentYear)

	If GetRole = "" Then
	
		'Change the grade number into text
		Select Case intGrade
			Case 0
				GetRole = "Kindergarten"
			Case 1
				GetRole = "1st"
			Case 2
				GetRole = "2nd"
			Case 3
				GetRole = "3rd"
			Case 4
				GetRole = "4th"   
			Case 5
				GetRole = "5th"
			Case 6
				GetRole = "6th"
			Case 7
				GetRole = "7th"
			Case 8
				GetRole = "8th"
			Case 9
				GetRole = "9th"
			Case 10
				GetRole = "10th"
			Case 11
				GetRole = "11th"
			Case 12
				GetRole = "12th"
			Case Else
				GetRole = "Graduated"
		End Select

	End If

End Function%>
<%
' Anything below here should exist on all pages 
%>

<%Sub DenyAccess 

   'If we're not using basic authentication then send them to the login screen
   If bolShowLogout Then
      Response.Redirect("login.asp?action=logout")
   Else
   
   SetupSite
   
   %>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title><%=Application("SiteName")%></title>
      <link rel="stylesheet" type="text/css" href="style.css" /> 
      <link rel="apple-touch-icon" href="images/inventory.png" /> 
      <link rel="shortcut icon" href="images/inventory.ico" />
      <meta name="viewport" content="width=device-width" />
   </head>
   <body>
      <center><b>Access Denied</b></center>
   </body>
   </html>
   
<% End If

End Sub%>

<%Function AccessGranted

   Dim objNetwork, strUserAgent, strSQL, strRole, objNameCheckSet

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

   'Get the users logon name
   Set objNetwork = CreateObject("WSCRIPT.Network")   
   strUser = objNetwork.UserName
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

   'Check and see if anonymous access is enabled
   If LCase(Left(strUser,4)) = "iusr" Then
      strUser = GetUser
      bolShowLogout = True
   Else
      bolShowLogout = False
   End If

   'Build the SQL string, this will check the userlevel of the user.
   strSQL = "Select Role" & vbCRLF
   strSQL = strSQL & "From Sessions" & vbCRLF
   strSQL = strSQL & "WHERE UserName='" & strUser & "' And SessionID='" & Request.Cookies("SessionID") & "'"
   Set objNameCheckSet = Application("Connection").Execute(strSQL)
   strRole = objNameCheckSet(0)

   If strRole = "Admin" or strRole = "User" Then
      AccessGranted = True
   Else
      AccessGranted = False
   End If

End Function%>

<%Function GetUser

   Const USERNAME = 1

   Dim strUserAgent, strSessionID, objSessionLookup, strSQL
   
   'Get some needed data
   strSessionID = Request.Cookies("SessionID")
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
   
   'Send them to the logon screen if they don't have a Session ID
   If strSessionID = "" Then
      SendToLogonScreen

   'Get the username from the database
   Else
   
      strSQL = "SELECT ID,UserName,SessionID,IPAddress,UserAgent,ExpirationDate FROM Sessions "
      strSQL = strSQL & "WHERE UserAgent='" & Left(Replace(strUserAgent,"'","''"),250) & "' And SessionID='" & Replace(strSessionID,"'","''") & "'"
      strSQL = strSQL & " And ExpirationDate > Date()"
      Set objSessionLookup = Application("Connection").Execute(strSQL)
      
      'If a session isn't found for then kick them out
      If objSessionLookup.EOF Then
         SendToLogonScreen
      Else
         GetUser = objSessionLookup(USERNAME)
      End If
   End If  
   
End Function%>

<%Function IsMobile

   Dim strUserAgent

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
End Function%>

<%Sub SendToLogonScreen

   Dim strReturnLink, strSourcePage
      
   'Build the return link before sending them away.
   strReturnLink =  "?" & Request.ServerVariables("QUERY_STRING")
   strSourcePage = Request.ServerVariables("SCRIPT_NAME")
   strSourcePage = Right(strSourcePage,Len(strSourcePage) - InStrRev(strSourcePage,"/"))
   If strReturnLink = "?" Then
      strReturnLink = "?SourcePage=" & strSourcePage
   Else
      strReturnLink = strReturnLink & "&SourcePage=" & strSourcePage
   End If
   
   Response.Redirect("login.asp" & strReturnLink)
   
End Sub %>

<%Sub SetupSite
   
   If IsMobile Then
      strSiteVersion = "Mobile"
   Else
      strSiteVersion = "Full"
   End If
   
End Sub%>