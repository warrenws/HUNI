<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/29/15
'Last Updated 1/14/18

'This page shows the details for a single user in the inventory website

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim strUserName, objDeviceList, objUser, strSubmitTo
Dim strCardType, intActiveAssignmentCount, intOldAssignmentCount, strBackLink, intUserID

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

	Dim strSQL

	'Get the variables from the URL
	strUserName = Request.QueryString("UserName")
	strBackLink = BackLink
	
	'If nothing was submitted send them back to the index page
   If strUserName = "" Then
      Response.Redirect("index.asp?Error=UserNotFound")
   End If

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Return"
         ReturnMissingItem
   End Select
   
   'Get the user's information
	strSQL = "SELECT ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,Active,Warning,PWord,AUP,Notes" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE UserName='" & Replace(strUserName,"'","''") & "'" & vbCRLF
	Set objUser = Application("Connection").Execute(strSQL)
	
	'Send them back to the index page if the user isn't found
   If objUser.EOF Then
   	Response.Redirect("index.asp?Error=UserNotFound")
   End If
   
   intUserID = objUser(0)
   
   'Get the list of devices assigned to the user
   strSQL = "SELECT Assignments.LGTag,DateIssued,DateReturned,Assignments.Active,Assignments.Notes,Model" & vbCRLF
	strSQL = strSQL & "FROM Devices INNER JOIN (People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo) ON Devices.LGTag = Assignments.LGTag" & vbCRLF
	strSQL = strSQL & "WHERE People.UserName='" & strUserName & "' AND Assignments.Active=True" & vbCRLF
	strSQL = strSQL & "ORDER BY DateIssued DESC"
	Set objDeviceList =  Application("Connection").Execute(strSQL)
	
	'Count the number of active and old assignments 
	intActiveAssignmentCount = 0
	intOldAssignmentCount = 0
	If Not objDeviceList.EOF Then
		Do Until objDeviceList.EOF
			If objDeviceList(3) Then
				intActiveAssignmentCount = intActiveAssignmentCount + 1
			Else
				intOldAssignmentCount = intOldAssignmentCount + 1
			End If
			objDeviceList.MoveNext
		Loop
		objDeviceList.MoveFirst
	End If
   
   'Get the URL used to submit forms
   If Request.ServerVariables("QUERY_STRING") = "" Then
      strSubmitTo = "user.asp"
   Else   
      strSubmitTo = "user.asp?" & Request.ServerVariables("QUERY_STRING")
   End If
   
   'Set up the variables needed for the site then load it
   SetupSite
   DisplaySite
   
End Sub%>

<%Sub DisplaySite %>

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
   		$(function() {
   		<% If Not IsMobile Then %>
					$( document ).tooltip({track: true});
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
		UserCard
		ActiveAssignments 
		%>

		<div class="Version">Version <%=Application("Version")%></div>
   </body>

   </html>

<%End Sub%>

<%Sub UserCard 

	Dim objFSO, strDeviceList

   Set objFSO = CreateObject("Scripting.FileSystemObject")

	If Not objUser.EOF Then
		
		Do Until objUser.EOF 
         
      	If objUser(8) Then 
				strCardType = "NormalCard"
			Else
				strCardType = "DisabledCard"
			End If %>
			<div class="Card <%=strCardType%>">
				<div class="CardTitle">
				<% If objUser(7) = "Student" Then %>
					<% If Application("ShowPasswords") Then %>
						<% If objUser(11) Then %>
								<image src="images/yes.png" width="15" height="15" title="AUP Signed" />
						<% Else %>
								<image src="images/no.png" width="15" height="15" title="AUP Not Signed" />
						<% End If %>
					<% End If %>
				<% End If %>
					<%=objUser(1) & " " & objUser(2)%>	
				</div>
				<div Class="ImageSectionInCard">
			<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objUser(7) & "s\" & objUser(4) & ".jpg") Then %>      
					<img class="PhotoCard" src="/photos/<%=objUser(7)%>s/<%=objUser(4)%>.jpg" width="96" />
			<% Else %>
					<img class="PhotoCard" src="/photos/<%=objUser(7)%>s/missing.png" width="96" />
			<% End If %>
				</div>
				<div Class="RightOfImageInCard">
					<div>
						<div Class="PhotoCardColumn1">Grade: </div>
						<div Class="PhotoCardColumn2Long">
							<a href="users.asp?Role=<%=objUser(5)%>"><%=GetRole(objUser(5))%></a>
						</div>
					</div>
				<% If objUser(7) = "Student" Then %>	
							<div>
								<div Class="PhotoCardColumn1">Guide: </div>
								<div Class="PhotoCardColumn2Long">
									<a href="users.asp?GuideRoom=<%=objUser(6)%>"><%=objUser(6)%></a>
								</div>
							</div>
						<%	If Application("ShowPasswords") Then %>
							<div>
								<div Class="CardMerged">Username: <%=strUserName%> </div>
							</div>
							<div>
								<div Class="CardMerged">Password: <%=objUser(10)%> </div>
							</div>
						<% End If %>
					<% End If %>

				<% If Not objDeviceList.EOF Then
						strDeviceList = "" %>
						<div Class="TwoColumnCardMerged">Assigned: 
				
					<%	Do Until objDeviceList.EOF
							strDeviceList = strDeviceList & "<a href=""device.asp?Tag=" & objDeviceList(0) & """>" & objDeviceList(0) & "</a>, "
							objDeviceList.MoveNext
						Loop 
						objDeviceList.MoveFirst

						strDeviceList = Left(strDeviceList,Len(strDeviceList) - 2)%>
						<%=strDeviceList%>
						</div>
				<%	End If %>
				</div>
			</div>
   
      <% objUser.MoveNext
      Loop 
   End If 

End Sub%>

<%Sub ActiveAssignments

	Dim objFSO, intLoopCounter, strSQL, objModel
	
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	intLoopCounter = 0 %>

<% If Not objDeviceList.EOF Then %>

	<% If intActiveAssignmentCount >= 1 Then %>

			<div class="Card NormalCard">
		
			<% If intActiveAssignmentCount = 1 Then %>
					<div class="CardTitle">Active Assignment</div>
			<% Else %>
					<div class="CardTitle">Active Assignments</div>
			<% End If %>

		<%	Do Until objDeviceList.EOF 
			
				'Active assignment 
				If objDeviceList(3) Then 
			
					intLoopCounter = intLoopCounter + 1
					
					strSQL = "Select Model FROM Devices WHERE LGTag='" & objDeviceList(0) & "'"
					Set objModel = Application("Connection").Execute(strSQL)
					
					%> 
				<div Class="ImageSectionInAssignmentCard">
				<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(5)," ","") & ".png") Then %>
						<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">      
							<img class="PhotoCard" src="images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="96" />
						</a>
				<% Else %>
						<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
							<img class="PhotoCard" src="images/devices/missing.png" width="96" />
						</a>
				<% End If %>
				</div>
				<div Class="RightOfImageInAssignmentCard">
							<div>
								<div Class="PhotoCardColumn1">Tag:</div>
								<div Class="PhotoCardColumn2">
									<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>"><%=objDeviceList(0)%></a>
								</div>
							</div>
						<% If Not objModel.EOF Then %>
								<div>
									<div Class="PhotoCardColumn1">Model: </div>
									<div Class="PhotoCardColumn2"><%=objModel(0)%></div>
								</div>
						<% End If %>
							<div>
								<div Class="PhotoCardColumn1">Date: </div>
								<div Class="PhotoCardColumn2"><%=ShortenDate(objDeviceList(1))%></div>
							</div>
						</div>
					

					<%	If Int(intActiveAssignmentCount) <> Int(intLoopCounter) Then %>
							<hr />
					<% End If %>
								
			<% End If
			
				objDeviceList.MoveNext
			Loop 
			objDeviceList.MoveFirst%>
			
			</div>
		<% End If %>
<%	End If %>

<%End Sub%>

<%Sub ReturnMissingItem

	Dim intAssignmentID, bolAdapterReturned, bolCaseReturned, strSQL

	intAssignmentID = Request.Form("AssignmentID")
	bolAdapterReturned = Request.Form("Adapter")
	bolCaseReturned = Request.Form("Case")
	
	If bolAdapterReturned Then
		strSQL = "UPDATE Assignments SET AdapterReturned=True WHERE ID=" & intAssignmentID
		Application("Connection").Execute(strSQL)
	End If
	
	If bolCaseReturned Then
		strSQL = "UPDATE Assignments SET CaseReturned=True WHERE ID=" & intAssignmentID
		Application("Connection").Execute(strSQL)
	End If
	
End Sub%>

<%Function BackLink

	Dim strPage
	
	If Request.QueryString("Back") = "" Then
		BackLink = "&Back=" & Server.UrlEncode(Request.ServerVariables("QUERY_STRING"))
		strPage = Right(Request.ServerVariables("SCRIPT_NAME"),Len(Request.ServerVariables("SCRIPT_NAME")) - InStrRev(Request.ServerVariables("SCRIPT_NAME"),"/"))
		BackLink = BackLink & "&Page=" & strPage
	Else
		BackLink = "&Back=" & Server.UrlEncode(Request.QueryString("Back")) & "&Page=" & Request.QueryString("Page")
	End If
	 
End Function%>

<%Function ShortenDate(strDate)
	
	If Not IsNull(strDate) Then
		If strDate <> "" Then
			ShortenDate = Left(strDate,Len(strDate) - 4) & Right(strDate,2)
		End If
	End If

End Function %>

<%Function ShortenTime(strTime)
	
	If Not IsNull(strTime) Then
		If strTime <> "" Then
			ShortenTime = Left(strTime,Len(strTime) - 6) & " " & Right(strTime,2)
		End If
	End If

End Function %>

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
End Function %>

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