<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/22/15
'Last Updated 1/14/18

'This page shows a list of users as a result of a search.

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim strFirstName, strLastName, strGuideRoom, objUserList, strRole, intUserCount, strAUP, strSite
Dim strDisplayColumns, strUserType, strView

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

	Dim strSQL

	'Get the variables from the URL
	strFirstName = Request.QueryString("FirstName")
	strLastName = Request.QueryString("LastName")
	strGuideRoom = Request.QueryString("GuideRoom")
	strRole = Request.QueryString("Role")
	strAUP = Request.QueryString("AUP")
	strView = Request.QueryString("View")
	strSite = Request.QueryString("UserSite")
	
	'If nothing was submitted send them back to the index page
   If strFirstName = "" Then
      If strLastName = "" Then
      	If strGuideRoom = "" Then
      		If strRole = "" Then
      			If strAUP = "" Then
      				If strSite = "" Then
      				Response.Redirect("index.asp?Error=NoUsersFound")
      				End If
      			End If
      		End If
      	End If
      End If
   End If
	
   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case ""
         
   End Select
	
	'Get the list of users
	strSQL = "SELECT ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,Active,Warning,Loaned,PWord,AUP,Site,Notes FROM People WHERE ClassOf > 2000 AND " & vbCRLF
	If strFirstName <> "" Then
      strSQL = strSQL & "FirstName Like '%" & Replace(strFirstName,"'","''") & "%' AND "
   End If
   If strLastName <> "" Then
      strSQL = strSQL & "LastName Like '%" & Replace(strLastName,"'","''") & "%' AND "
   End If
   If strSite <> "" Then
   	  strSQL = strSQL & "Site = '" & Replace(strSite,"'","''") & "' AND "
   End If
   If strGuideRoom <> "" Then
      strSQL = strSQL & "HomeRoom Like '%" & Replace(strGuideRoom,"'","''") & "%' AND "
   End If
   If strAUP <> "" Then
   	strSQL = strSQL & "AUP=" & strAUP & " AND "
   End If
   If strRole <> "" Then
   	strSQL = strSQL & "ClassOf=" & strRole & "     "
   End If
   
   strSQL = Left(strSQL,Len(strSQL) - 5) & " AND Active=True" & vbCRLF
   strSQL = strSQL & "ORDER BY LastName, FirstName"

	Set objUserList = Application("Connection").Execute(strSQL)

	'If no user is found send them back to the index page.
   If objUserList.EOF Then
    	Response.Redirect("index.asp?Error=NoUsersFound")
   Else
   	intUserCount = 0
   	Do Until objUserList.EOF
   		intUserCount = intUserCount + 1
   		objUserList.MoveNext
   	Loop
   	objUserList.MoveFirst
   End If	
   
   If LCase(strView) = "table" Then
   	strSiteVersion = "Full"
   End If
	
	'0		0	Photo
	'1		1	Name
	'2		2	First Name
	'3		3	Last Name
	'4		4	User Name
	'5		5	EMail
	'		6	Password
	'		7	AUP
	'6		8	Site
	'7		9	Guide Room
	'8		10	Class Of
	'9		11	Role
	'10	12	Assigned Device
	'11	13	Assigned Tag
	
	'Find out if they are a student or adult
   If strRole <> "" Then
		If IsNumeric(strRole) Then
			If strRole >= 2000 Then
				strUserType = "Student"
			Else 
				strUserType = "Adult"
			End If
		Else
			strUserType = strRole
		End If
	End If
	
	'Set the visible columns
	If IsMobile Then
		If Application("ShowPasswords") Then
			strDisplayColumns = "0,2,3,4,5,7,8,9,10,11,12,13"
		Else
			strDisplayColumns = "0,2,3,5,6,7,8,9,10,11"
		End If
	Else
		If Application("ShowPasswords") Then
			strDisplayColumns = "0,2,3,5,10,12"
		Else
			strDisplayColumns = "0,2,3,5,11"
		End If
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
		<link rel="stylesheet" href="assets/css/jquery.dataTables.min.css">
		<link rel="stylesheet" href="assets/css/buttons.dataTables.min.css">
		<script src="assets/js/jquery.dataTables.min.js"></script>
		<script src="assets/js/dataTables.buttons.min.js"></script>
		<script src="assets/js/buttons.colVis.min.js"></script>
		<script src="assets/js/buttons.html5.min.js"></script>
		<script src="assets/js/jszip.min.js"></script>
		<script type="text/javascript">
			$(document).ready( function () {
			
			<% If Not IsMobile Then %>
					$( document ).tooltip({track: true});
			<% End If %>
			
    			var table = $('#ListView').DataTable( {
    				paging: false,
    				"info": false,
    				"autoWidth": false,
    				"order": [[ 1, "asc" ]],
    				dom: 'Bfrtip',
    				// stateSave: true,
    				buttons: [
						{
							extend: 'colvis',
							text: 'Show/Hide Columns'
						}
				<% If Not IsMobile Then %>		
						,
						{
							extend: 'csvHtml5',
							text: 'Download CSV',
							title: 'Inventory - Users'
						}
				<% End If %>
        			]
        		
    			});
					table.columns([<%=strDisplayColumns%>]).visible(false);	
    		
    		} );
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
      <div align="center">Number of users found: <%=intUserCount%></div>
         
		<% If Not objUserList.EOF Then
	
		Select Case LCase(strView)
			Case "table"
				ShowUserTable
			Case "card"
				ShowUserCards
			Case Else
				ShowUserCards
		End Select
   End If %>
		
		<div class="Version">Version <%=Application("Version")%></div>
   </body>

   </html>

<%End Sub%>

<%Sub ShowUserCards%>

<% If Not objUserList.EOF Then
		
		Dim objFSO, strSQL, objDeviceList, strDeviceList
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")%>
		
		<div class="ViewButtonUser">
			<a href="<%=SwitchView("Table")%>"><img src="images/table.png" title="Table View" height="32" width="32"/></a>
		</div>
		<div>
	<%	Do Until objUserList.EOF %>
		
			<div class="Card NormalCard">
				<div class="CardTitle">
				<% If objUserList(7) = "Student" Then %>
					<% If Application("ShowPasswords") Then %>
						<% If objUserList(13) Then %>
								<image src="images/yes.png" width="15" height="15" title="AUP Signed" />
						<% Else %>
								<image src="images/no.png" width="15" height="15" title="AUP Not Signed" />
						<% End If %>
					<% End If %>
				<% End If %>
					<a href="user.asp?UserName=<%=objUserList(3)%>"><%=objUserList(1) & " " & objUserList(2)%></a>
				</div>
				<div Class="ImageSectionInCard">
				<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objUserList(7) & "s\" & objUserList(4) & ".jpg") Then %>   
						<a href="user.asp?UserName=<%=objUserList(3)%>">   
							<img class="PhotoCard" src="/photos/<%=objUserList(7)%>s/<%=objUserList(4)%>.jpg" width="96" />
						</a>
				<% Else %>
						<a href="user.asp?UserName=<%=objUserList(3)%>">
							<img class="PhotoCard" src="/photos/<%=objUserList(7)%>s/missing.png" width="96" />
						</a>
				<% End If %>
				</div>
				<div Class="RightOfImageInCard">
					<div>
						<div Class="PhotoCardColumn1">Grade: </div>
						<div Class="PhotoCardColumn2Long">
							<a href="users.asp?Role=<%=objUserList(5)%>"><%=GetRole(objUserList(5))%></a>
						</div>
					</div>
				<% If objUserList(7) = "Student" Then %>	
						<div>
							<div Class="PhotoCardColumn1">Guide: </div>
							<div Class="PhotoCardColumn2Long">
								<a href="users.asp?GuideRoom=<%=objUserList(6)%>"><%=objUserList(6)%></a>
							</div>
						</div>
					<%	If Application("ShowPasswords") Then %>
						<div>
							<div Class="CardMerged">Username: <%=objUserList(3)%></div>
						</div>
						<div>
							<div Class="CardMerged">Password: <%=objUserList(12)%></div>
						</div>
					
				<% End If
				End If %>
			
			<% If objUserList(8) > 0 Then
				
				strSQL = "SELECT Assignments.LGTag,Devices.Model" & vbCRLF
				strSQL = strSQL & "FROM Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag" & vbCRLF
				strSQL = strSQL & "WHERE Assignments.AssignedTo=" & objUserList(0) & " AND Assignments.Active=True"
				Set objDeviceList = Application("Connection").Execute(strSQL)
				
				If Not objDeviceList.EOF Then
					strDeviceList = "" %>
					<div Class="TwoColumnCardMerged">Assigned: 
				
				<%	Do Until objDeviceList.EOF
						strDeviceList = strDeviceList & "<a href=""device.asp?Tag=" & objDeviceList(0) & """>" & objDeviceList(0) & "</a>, "
						objDeviceList.MoveNext
					Loop 

					strDeviceList = Left(strDeviceList,Len(strDeviceList) - 2)%>
					<%=strDeviceList%>
					</div>
			<%	End If
			End If %>
				</div>
			</div>
   
      <% objUserList.MoveNext
      Loop %>
      
      </div>
<% End If %>
   
<%End Sub%>

<%Sub ShowUserTable 

	Dim strSQL, objDeviceList, strDeviceAssetTagList, strDeviceList, strRowClass, objFSO, intDaysRemaining, arrPWordLastSet
	
	Set objFSO = CreateObject("Scripting.FileSystemObject") %>
	
	<div class="ViewButtonUser">
		<a href="<%=SwitchView("Card")%>"><img src="images/card.png" title="Card View" height="32" width="32"/></a>
	</div>

	<div>
		<table align="center" Class="ListView" id="ListView">
			<thead>
			<th>Photo</th>
			<th>Name</th>
			<th>First Name</th>
			<th>Last Name</th>
			<th>User Name</th>
			<th>Email</th>
		<% If Application("ShowPasswords") Then %>
				<th>Password</th>
				<th>AUP</th>
		<% End If %>
			<th>Site</th>
			<th>Guide Room</th>
			<th>Class Of</th>
			<th>Grade</th>
			<th>Assigned Device</th>
			<th>Assigned Tag</th>
			</thead>
			<tbody>
	<% Do Until objUserList.EOF 
	
			If objUserList(10) Then
				'strRowClass = " Class=""Warning"""
			ElseIf objUserList(11) Then
				'strRowClass = " Class=""Loaned"""
			ElseIf objUserList(9) Then
				strRowClass = ""
			Else 
				strRowClass = " Class=""Disabled"""
			End If %>
			
			<tr <%=strRowClass%>>
			<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objUserList(7) & "s\" & objUserList(4) & ".jpg") Then %>   
					<td <%=strRowClass%> width="1px">
						<a href="user.asp?UserName=<%=objUserList(3)%>">   
							<img src="/photos/<%=objUserList(7)%>s/<%=objUserList(4)%>.jpg" title="<%=objUserList(4)%>" width="72" />
						</a>
					</td>
			<% Else %>
					<td <%=strRowClass%> width="1px">
						<a href="user.asp?UserName=<%=objUserList(3)%>">
							<img src="/photos/<%=objUserList(7)%>s/missing.png" title="<%=objUserList(4)%>" width="72" />
						</a>
					</td>
			<% End If %>
				<td <%=strRowClass%>><a href="user.asp?UserName=<%=objUserList(3)%>"><%=objUserList(2)%>, <%=objUserList(1)%></a></td>
				<td <%=strRowClass%>><a href="user.asp?UserName=<%=objUserList(3)%>"><%=objUserList(1)%></a></td>
				<td <%=strRowClass%>><a href="user.asp?UserName=<%=objUserList(3)%>"><%=objUserList(2)%></a></td>
				<td <%=strRowClass%>><%=objUserList(3)%></td>
				<td <%=strRowClass%>><%=objUserList(3)%>@<%=Application("Domain")%></td>
			<% If Application("ShowPasswords") Then %>	
					<td <%=strRowClass%>><%=objUserList(12)%></td>
				<% If objUserList(13) Then %>
						<td <%=strRowClass%> id="center">Yes</td>
				<% Else %>
						<td <%=strRowClass%> id="center">No</td>
				<% End If %>
			<% End If %>
				
				<td <%=strRowClass%>><%=objUserList(14)%></td>
				<td <%=strRowClass%>><a href="users.asp?GuideRoom=<%=objUserList(6)%>&View=Table"><%=objUserList(6)%></a></td>
			<% If objUserList(5) > 2000 Then %>	
					<td <%=strRowClass%> id="center"><%=objUserList(5)%></td>
			<% Else %>
					<td <%=strRowClass%> id="center"></td>
			<% End If %>
				<td <%=strRowClass%>><a href="users.asp?Role=<%=objUserList(5)%>&View=Table"><%=GetRole(objUserList(5))%></a></td>
				
			<% strDeviceList = ""
				strDeviceAssetTagList = ""
				If objUserList(8) > 0 Then
			
					strSQL = "SELECT Assignments.LGTag,Devices.Model" & vbCRLF
					strSQL = strSQL & "FROM Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag" & vbCRLF
					strSQL = strSQL & "WHERE Assignments.AssignedTo=" & objUserList(0) & " AND Assignments.Active=True"
					Set objDeviceList = Application("Connection").Execute(strSQL)
					
			
					If Not objDeviceList.EOF Then 
						Do Until objDeviceList.EOF
							strDeviceList = strDeviceList & "<a href=""device.asp?Tag=" & objDeviceList(0) & """>" & objDeviceList(1) & "</a>, "
							strDeviceAssetTagList = strDeviceAssetTagList & "<a href=""device.asp?Tag=" & objDeviceList(0) & """>" & objDeviceList(0) & "</a>, "
							objDeviceList.MoveNext
						Loop 
						strDeviceList = Left(strDeviceList,Len(strDeviceList) - 2)
						strDeviceAssetTagList = Left(strDeviceAssetTagList,Len(strDeviceAssetTagList) - 2)
					End If
				End If %>
				<td <%=strRowClass%>><%=strDeviceList%></td>
				<td <%=strRowClass%>><%=strDeviceAssetTagList%></td>
			</tr>
		<% objUserList.MoveNext
		Loop %>
			</tbody>
		</table>
<%End Sub%>

<%Function SwitchView(strView)

	Dim strURL
	
	If LCase(Request.QueryString("View")) = "" Then
		If Request.ServerVariables("QUERY_STRING") = "" Then
			SwitchView = "users.asp?View=" & strView
		Else
			SwitchView = "users.asp?" & Request.ServerVariables("QUERY_STRING") & "&View=" & strView
		End If	
	Else
		Select Case LCase(strView)
			Case "card"  
				SwitchView = "users.asp?" & Replace(Request.ServerVariables("QUERY_STRING"),"Table","Card")
			Case "table"
				SwitchView = "users.asp?" & Replace(Request.ServerVariables("QUERY_STRING"),"Card","Table")
		End Select
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