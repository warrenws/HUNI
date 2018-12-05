<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/22/15
'Last Updated 1/14/18

'This page shows a list of users as a result of a search.

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim strFirstName, strLastName, strGuideRoom, objUserList, strRole, intUserCount, strWithDevice
Dim strStatus, strSite, strMissing, strView, strSearchMessage, strBackLink, strLoanedOut
Dim strCardType, strColumns, strAUP, strNotes, strDisplayColumns, strUserType, objLastNames
Dim strDescription, strCurrentPage, strOwes, strCustomDisplay, intOrderColumn, strInternetAccess

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

	Dim strSQL, strSQLWhere, strDisabled

	'Get the variables from the URL
	strSite = Request.QueryString("UserSite")
	strFirstName = Request.QueryString("FirstName")
	strLastName = Request.QueryString("LastName")
	strGuideRoom = Request.QueryString("GuideRoom")
	strRole = Request.QueryString("Role")
	strWithDevice = Request.QueryString("WithDevice")
	strStatus = Request.QueryString("UserStatus")
	strMissing = Request.QueryString("Missing")
	strView = Request.QueryString("View")
	strLoanedOut = Request.QueryString("LoanedOut")
	strNotes = Request.QueryString("UserNotes")
	strDescription = Request.QueryString("Description")
	strOwes = Request.QueryString("Owes")
	strInternetAccess = Request.QueryString("InternetAccess")
	strAUP = Request.QueryString("AUP")
	
	strCustomDisplay = Request.QueryString("Display")
	strBackLink = BackLink
	strCurrentPage = LCase(Right(Request.ServerVariables("URL"),Len(Request.ServerVariables("URL")) - InStrRev(Request.ServerVariables("URL"),"/")))

	'If nothing was submitted send them back to the index page
   If strFirstName = "" Then
      If strLastName = "" Then
      	If strGuideRoom = "" Then
      		If strRole = "" Then
      			If strSite = "" Then
      				If strWithDevice = "" Then
      					If strStatus = "" Then
      						If strLoanedOut = "" Then
									If strMissing = "" Then
										If strDescription = "" Then
											If strOwes = "" Then
												If strAUP = "" Then
													If strNotes = "" Then
														If strInternetAccess = "" Then
															If Request.QueryString("Source") <> "" Then
																Response.Redirect("search.asp?Error=NoUsersFound")
															ElseIf strCurrentPage = "users.asp" Then
																Response.Redirect("index.asp?Error=NoUsersFound")
															Else
																Response.Redirect(Request.QueryString("Source") & "?Error=NoUsersFound")
															End If
														End If
													End If 
												End If
											End If
										End If
									End If
								End If
      					End If
      				End If
      			End If
      		End If
      	End If
      End If
   End If
	
   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Save"
      	SaveSearch
         
   End Select
	
	'Get the list of users
	strSQLWhere = "WHERE "
	If strSite <> "" Then
      strSQLWhere = strSQLWhere & "People.Site = '" & Replace(strSite,"'","''") & "' AND "
   End If
	If strFirstName <> "" Then
      strSQLWhere = strSQLWhere & "FirstName Like '%" & Replace(strFirstName,"'","''") & "%' AND "
   End If
   If strLastName <> "" Then
      strSQLWhere = strSQLWhere & "LastName Like '%" & Replace(strLastName,"'","''") & "%' AND "
   End If
   If strGuideRoom <> "" Then
      strSQLWhere = strSQLWhere & "HomeRoom Like '%" & Replace(strGuideRoom,"'","''") & "%' AND "
   End If
   If strNotes <> "" Then
      strSQLWhere = strSQLWhere & "People.Notes Like '%" & Replace(strNotes,"'","''") & "%' AND "
   End If
   If strDescription <> "" Then
      strSQLWhere = strSQLWhere & "People.Description Like '%" & Replace(strDescription,"'","''") & "%' AND "
   End If
   If strOwes <> "" Then
   	strSQLWhere = strSQLWhere & "Warning=True AND "
   End If
   If strAUP <> "" Then
   	strSQLWhere = strSQLWhere & "AUP=" & strAUP & " AND "
   End If
   If strInternetAccess <> "" Then
   	If strInternetAccess = "Unknown" Then
   		strSQLWhere = strSQLWhere & "(InternetAccess Is Null OR InternetAccess='') AND "
   	Else
   		strSQLWhere = strSQLWhere & "InternetAccess='" & Replace(strInternetAccess,"'","''") & "' AND "
   	End If
   End If
   If strRole <> "" Then
   	Select Case strRole
            Case "Adult", "Student"
               strSQLWhere = strSQLWhere & "People.Role='" & Replace(Replace(strRole,"'","''"),"Adult","Teacher") & "' AND "
            Case Else
               strSQLWhere = strSQLWhere & "People.ClassOf=" & Replace(strRole,"'","''") & " AND "
         End Select
   End If
   
   Select Case strWithDevice
   	Case "Yes"
   		strSQLWhere = strSQLWhere & "HasDevice=True AND "
   	Case "No"
   		strSQLWhere = strSQLWhere & "HasDevice=False AND "
   End Select
   
   Select Case strStatus
		Case "Enabled"
			strSQLWhere = strSQLWhere & "People.Active=True AND "
		Case "Disabled"
			strSQLWhere = strSQLWhere & "People.Active=False AND "
		Case "All"
		Case Else
			strSQLWhere = strSQLWhere & "People.Active=True AND "
   End Select
   
   Select Case strWithDevice
		Case "Yes"
			strSQLWhere = strSQLWhere & "People.HasDevice=True AND "
		Case "No"
			strSQLWhere = strSQLWhere & "People.HasDevice=False AND "
   End Select
   
   strSQLWhere = strSQLWhere & "People.Deleted=False AND "
	
	strSQLWhere = Left(strSQLWhere,Len(strSQLWhere) - 5)
	
	If strMissing <> "" Then
		Select Case strMissing
			Case "Anything"
				strSQL = "SELECT People.ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,People.Active,Warning,Loaned,PWord,AUP,Site,People.Notes,PWordLastSet,PhoneNumber,RoomNumber,Description,PWordNeverExpires,LastExternalCheckIn,LastInternalCheckIn,Birthday,InternetAccess" & vbCRLF
				strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
				strSQL = strSQL & strSQLWhere & " AND (CaseReturned=False OR AdapterReturned=False) AND Assignments.Active=False"
			Case "Case"
				strSQL = "SELECT People.ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,People.Active,Warning,Loaned,PWord,AUP,Site,People.Notes,PWordLastSet,PhoneNumber,RoomNumber,Description,PWordNeverExpires,LastExternalCheckIn,LastInternalCheckIn,Birthday,InternetAccess" & vbCRLF
				strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
				strSQL = strSQL & strSQLWhere & " AND CaseReturned=False AND Assignments.Active=False"
			Case "Power Supply"
				strSQL = "SELECT People.ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,People.Active,Warning,Loaned,PWord,AUP,Site,People.Notes,PWordLastSet,PhoneNumber,RoomNumber,Description,PWordNeverExpires,LastExternalCheckIn,LastInternalCheckIn,Birthday,InternetAccess" & vbCRLF
				strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
				strSQL = strSQL & strSQLWhere & " AND AdapterReturned=False AND Assignments.Active=False"
		End Select
	ElseIf strLoanedOut <> "" Then
		Select Case strLoanedOut
			Case "Anything"
				strSQL = "SELECT DISTINCT People.ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,People.Active,Warning,Loaned,PWord,AUP,Site,People.Notes,PWordLastSet,PhoneNumber,RoomNumber,Description,PWordNeverExpires,LastExternalCheckIn,LastInternalCheckIn,Birthday,InternetAccess" & vbCRLF
				strSQL = strSQL & "FROM People INNER JOIN Loaned ON People.ID = Loaned.AssignedTo" & vbCRLF
				strSQL = strSQL & strSQLWhere & " AND (Loaned.Returned=False)"
			Case Else
				strSQL = "SELECT DISTINCT People.ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,People.Active,Warning,Loaned,PWord,AUP,Site,People.Notes,PWordLastSet,PhoneNumber,RoomNumber,Description,PWordNeverExpires,LastExternalCheckIn,LastInternalCheckIn,Birthday,InternetAccess" & vbCRLF
				strSQL = strSQL & "FROM People INNER JOIN Loaned ON People.ID = Loaned.AssignedTo" & vbCRLF
				strSQL = strSQL & strSQLWhere & " AND (Loaned.Returned=False) AND Loaned.Item='" & Replace(strLoanedOut,"'","''") & "'"
		End Select
	Else
		strSQL = "SELECT ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,Active,Warning,Loaned,PWord,AUP,Site,Notes,PWordLastSet,PhoneNumber,RoomNumber,Description,PWordNeverExpires,LastExternalCheckIn,LastInternalCheckIn,Birthday,InternetAccess FROM People " & strSQLWhere
	End If

   strSQL = strSQL & " ORDER BY LastName, FirstName"
   Set objUserList = Application("Connection").Execute(strSQL)

	'If no user is found send them back to the index page.
   If objUserList.EOF Then
   	If Request.QueryString("Source") = "" Then
   		Response.Redirect("search.asp?Error=NoUsersFound")
   	Else
   		Response.Redirect(Request.QueryString("Source") & "?Error=NoUsersFound")
   	End If
   Else
   	intUserCount = 0
   	Do Until objUserList.EOF
   		intUserCount = intUserCount + 1
   		objUserList.MoveNext
   	Loop
   	objUserList.MoveFirst
   End If	
   
   'Get the list of lastnames for the auto complete
   strSQL = "SELECT DISTINCT LastName FROM People WHERE Active=True"
   Set objLastNames = Application("Connection").Execute(strSQL)
	
   'Set up the variables needed for the site then load it
   SetupSite
   If LCase(strView) = "table" Then
   	strSiteVersion = "Full"
   End If
   
	'0		0	Photo
	'1		1	ID
	'2		2	Name
	'3		3	First Name
	'4		4	Last Name
	'5		5	User Name
	'6		6	EMail
	'7    7  Description
	'		8	Password
	'		9	AUP
	'8		10	Password Changed
	'9		11	Days Remaining
	'10	12	Site
	'11	13	Room
	'12	14	Phone
	'13	15	Guide Room
	'14	16	Class Of
	'15	17	Role
	'16	18	Assigned Device
	'17	19	Assigned Tag
	'18	20	User Notes

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
	
	If strCustomDisplay = "" Then
	
		'Set the column to sort by
		intOrderColumn = 2
	
		'Set the hidden columns
		Select Case strUserType
		
			Case "Student"
				If IsMobile Then
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,5,6,7,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26"  
					Else
						strDisplayColumns = "0,1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,18,19,20,21,22,23,24"  
					End If
				Else
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,6,7,9,10,11,13,14,16,17,19,21,22,23,24,25,26"  
					Else
						strDisplayColumns = "0,1,3,4,6,7,8,9,10,11,12,14,17,19,20,21,22,23,24"  
					End If
				End If

			Case "Adult"
				If IsMobile Then
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,5,6,7,8,9,10,12,13,15,16,17,18,19,20,21,22,23,24,25,26"  
					Else
						strDisplayColumns = "0,1,3,4,5,6,7,8,10,11,13,14,15,16,17,18,19,20,21,22,23,24"  
					End If
				Else
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,6,7,8,9,10,12,15,16,17,18,19,21,22,23,24,25,26"  
					Else
						strDisplayColumns = "0,1,3,4,6,7,8,10,13,14,15,16,17,19,20,21,22,23,24"  
					End If
				End If		
		
			Case Else
				If IsMobile Then
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,5,6,7,9,10,11,12,13,14,15,16,17,18,19,21,22,23,24,25,26"  
					Else
						strDisplayColumns = "0,1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,18,19,20,21,22,23,24"  
					End If
				Else
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,6,7,9,10,11,13,14,15,18,19,21,22,23,24,25,26"  
					Else
						strDisplayColumns = "0,1,3,4,6,7,8,9,11,12,13,17,19,20,21,22,23,24"  
					End If
				End If

		End Select
	
	Else
	
		Select Case strCustomDisplay
			Case "Owed"
				intOrderColumn = 2
				If IsMobile Then
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,23,25,26"  
					Else
						strDisplayColumns = "0,1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,21,23,24"  
					End If
				Else
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,6,7,9,10,11,13,14,16,17,19,20,21,23,25,26"  
					Else
						strDisplayColumns = "0,1,3,4,6,7,8,9,10,11,12,14,17,20,22,23"  
					End If
				End If
			Case "CheckIn"
				If IsMobile Then
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24"  
						If Request.QueryString("Internal") = "True" Then
							intOrderColumn = 24
						Else
							intOrderColumn = 23
						End If
					Else
						strDisplayColumns = "0,1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22"  
						If Request.QueryString("Internal") = "True" Then
							intOrderColumn = 22
						Else
							intOrderColumn = 21
						End If
					End If
				Else
					If Application("ShowPasswords") Then
						strDisplayColumns = "0,1,3,4,6,7,9,10,11,13,14,16,17,19,21,22,23,24"  
						If Request.QueryString("Internal") = "True" Then
							intOrderColumn = 24
						Else
							intOrderColumn = 23
						End If
					Else
						strDisplayColumns = "0,1,3,4,6,7,8,9,10,11,12,14,17,19,20,21"  
						If Request.QueryString("Internal") = "True" Then
							intOrderColumn = 22
						Else
							intOrderColumn = 21
						End If
					End If
				End If
		End Select
	
	End If
	   
   DisplaySite
   
End Sub%>

<%Sub DisplaySite %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
		<link rel="stylesheet" href="../assets/css/jquery.dataTables.min.css">
		<link rel="stylesheet" href="../assets/css/buttons.dataTables.min.css">
		<script src="../assets/js/jquery.dataTables.min.js"></script>
		<script src="../assets/js/dataTables.buttons.min.js"></script>
		<script src="../assets/js/buttons.colVis.min.js"></script>
		<script src="../assets/js/buttons.html5.min.js"></script>
		<script src="../assets/js/jszip.min.js"></script>
		<script type="text/javascript">
			$(document).ready( function () {
			
			<%	If Not IsMobile And Not IsiPad Then%>
					$( document ).tooltip({track: true});
			<% End If %>	
				
    			var table = $('#ListView').DataTable( {
    				paging: false,
    				"info": false,
    				"autoWidth": false,
    				"order": [[ <%=intOrderColumn%>, "asc" ]],
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
    				$('#body').show();
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
    		
    		} );
    	</script>
   </head>

   <body class="<%=strSiteVersion%>" id="body" style="display:none;" >
   
      <div class="Header"><%=Application("SiteName")%> (<%=intUserCount%>)</div>
      <div>
         <ul class="NavBar" align="center">
            <li><a href="index.asp"><img src="../images/home.png" title="Home" height="32" width="32"/></a></li>
            <li><a href="search.asp"><img src="../images/search.png" title="Search" height="32" width="32"/></a></li>
            <li><a href="stats.asp"><img src="../images/stats.png" title="Stats" height="32" width="32"/></a></li>
            <li><a href="log.asp"><img src="../images/log.png" title="System Log" height="32" width="32"/></a></li>
            <li><a href="add.asp"><img src="../images/add.png" title="Add Person or Device" height="32" width="32"/></a></li>
            <li><a href="login.asp?action=logout"><img src="../images/logout.png" title="Log Out" height="32" width="32"/></a></li>
         </ul>
      </div>
      <%JumpToDevice%>
      
<% If Not objUserList.EOF Then
	
		Select Case LCase(strView)
			Case "table"
				ShowUserTable
			Case "card"
				ShowUserCards
			Case Else
			
				If IsMobile Then
				
					If LCase(Application("DefaultViewMobile")) = "table" Then
						If intUserCount < Application("CardThreshold") Then
							ShowUserCards
						Else
							ShowUserTable
						End If
					Else
						ShowUserCards
					End If
				
				Else
			
					If LCase(Application("DefaultView")) = "table" Then
						If intUserCount < Application("CardThreshold") Then
							ShowUserCards
						Else
							ShowUserTable
						End If
					Else
						ShowUserCards
					End If
					
				End If
		End Select
      SaveAsSearch
   End If %>
		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
   </body>

   </html>

<%End Sub%>

<%Sub ShowUserCards 

	Dim arrPWordLastSet, intAge

	If IsMobile Then %>
		<div class="ViewButtonMobile">
<% Else %>
		<div class="ViewButton">
<% End If %>
		<a href="<%=SwitchView("Table")%>"><img src="../images/table.png" title="Table View" height="32" width="32"/></a>
	</div>
	
	<div class="center"><%=FilterBar%></div>
	<div Class="<%=strColumns%>">
<%	If Not objUserList.EOF Then
		
		Dim objFSO, objDeviceList, strSQL, strDeviceList, intDaysRemaining, strUserInfo
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		Do Until objUserList.EOF
		
			If objUserList(10) Then 
				strCardType = "WarningCard"
		 	ElseIf objUserList(11) Then 
				strCardType = "LoanedCard"
			ElseIf objUserList(9) Then 
				strCardType = "NormalCard"
		 	Else 
				strCardType = "DisabledCard"
			End If 
			
			'Build the user info popup
			strUserInfo = ""
			If objUserList(22) <> "" Then
				strUserInfo = "Internal Access: " & objUserList(22) & " &#013 "
			End If
			If objUserList(21) Then
				strUserInfo = strUserInfo & "External Access: " & objUserList(21) & " &#013 "
			End If
			If objUserList(23) <> "" Then
				intAge = DateDiff("yyyy",objUserList(23),Date)
				If Date < DateSerial(Year(Date), Month(objUserList(23)), Day(objUserList(23))) Then
					intAge = intAge - 1 
				End If
				strUserInfo = strUserInfo & "Birthday: " & objUserList(23) & " &#013 "
   			strUserInfo = strUserInfo & "Age: " & intAge
			End If %>
			
			<div class="Card <%=strCardType%>">
				<div class="CardTitle">
				<% If objUserList(7) = "Student" Then %>
					<% If Application("ShowPasswords") Then %>
						<% If objUserList(13) Then %>
								<image src="../images/yes.png" width="15" height="15" title="AUP Signed" />
						<% Else %>
								<image src="../images/no.png" width="15" height="15" title="AUP Not Signed" />
						<% End If %>
					<% End If %>
				<% End If %>
					<a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>"><%=objUserList(1) & " " & objUserList(2)%></a>
				<% If strUserInfo <> "" Then %>
						<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strUserInfo%>"  />&nbsp;</div>
				<% End If %>
				</div>
				<div Class="ImageSectionInCard">
				<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objUserList(7) & "s\" & objUserList(4) & ".jpg") Then %>   
						<a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>">   
							<img class="PhotoCard" src="/photos/<%=objUserList(7)%>s/<%=objUserList(4)%>.jpg" title="<%=objUserList(4)%>" width="96" />
						</a>
				<% Else %>
						<a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>">
							<img class="PhotoCard" src="/photos/<%=objUserList(7)%>s/missing.png" title="<%=objUserList(4)%>" width="96" />
						</a>
				<% End If %>
			
				</div>
				<div Class="RightOfImageInCard">
				<div>
					<div Class="PhotoCardColumn1">Role: </div>
					<div Class="PhotoCardColumn2Long">
						<a href="users.asp?Role=<%=objUserList(5)%>"><%=GetRole(objUserList(5))%></a>
					</div>
				</div>
			<% If objUserList(7) = "Student" Then 
					If objUserList(6) <> "" Then %>	
						<div>
							<div Class="PhotoCardColumn1"><%=Application("HomeroomName")%>: </div>
							<div Class="PhotoCardColumn2Long">
								<a href="users.asp?GuideRoom=<%=objUserList(6)%>"><%=objUserList(6)%></a>
							</div>
						</div>
				<%	End If
					If Application("ShowPasswords") Then %>
				   <div>
						<div Class="CardMerged">Username: <%=objUserList(3)%></div>
					</div>
				   <div>
						<div Class="CardMerged">Password: <%=objUserList(12)%></div>
					</div>
				<% End If
				Else 
					
					If Not objUserList(20) Then 'Password doesn't expire
					
						If Not IsNull(objUserList(16)) Then
							arrPWordLastSet = Split(objUserList(16)," ")
							If CDate(arrPWordLastSet(0)) > #1/1/80# Then 
					
								intDaysRemaining = DateDiff("d",Date(),DateAdd("d",Application("PasswordsExpire"),arrPWordLastSet(0))) %>
					
								<div Class="CardMerged">Password Changed: <%=ShortenDate(arrPWordLastSet(0))%></div>
						
							<% If intDaysRemaining > 10 Then %>
									<div Class="CardMerged">Days Remaining: <%=intDaysRemaining%></div>
							<% ElseIf intDaysRemaining >= 1 Then %>
									<div Class="CardMerged Error">Days Remaining: <%=intDaysRemaining%></div>
							<% Else %>
									<div Class="CardMerged Error">Days Remaining: Expired</div>
							<% End If 
						
							Else %>
						
								<div Class="CardMerged">Password Changed: ---</div>
								<div Class="CardMerged Error">Days Remaining: Expired</div>
							
						<%	End If %>
					<% Else %>
							<div Class="CardMerged">Password Changed: ---</div>
							<div Class="CardMerged">Days Remaining: ---</div>
					<%	End If %>
				<% Else 
						If Not IsNull(objUserList(16)) Then
							arrPWordLastSet = Split(objUserList(16)," ") %>
							<div Class="CardMerged">Password Changed: <%=ShortenDate(arrPWordLastSet(0))%></div>
							<div Class="CardMerged">Days Remaining: ---</div>
					<% Else %>
							<div Class="CardMerged">Password Changed: ---</div>
							<div Class="CardMerged">Days Remaining: ---</div>
					<% End If %>
				<%	End If%>
			<%	End If %>
				
			<% If objUserList(17) <> "" Then %>
				<div Class="CardMerged">Phone: <%=objUserList(17)%></div>
			<% End If %>
			
			<% If objUserList(8) > 0 Then
					
					strSQL = "SELECT Assignments.LGTag,Devices.Model" & vbCRLF
					strSQL = strSQL & "FROM Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag" & vbCRLF
					strSQL = strSQL & "WHERE Assignments.AssignedTo=" & objUserList(0) & " AND Assignments.Active=True"
					Set objDeviceList = Application("Connection").Execute(strSQL)
					
					If Not objDeviceList.EOF Then
						strDeviceList = "" %>
						<div Class="CardMerged">Assigned: 
					
					<%	Do Until objDeviceList.EOF
							strDeviceList = strDeviceList & "<a href=""device.asp?Tag=" & objDeviceList(0) & strBackLink & """>" & objDeviceList(1) & "</a>, "
							objDeviceList.MoveNext
						Loop 
	
						strDeviceList = Left(strDeviceList,Len(strDeviceList) - 2)%>
						<%=strDeviceList%>
						</div>
				<%	End If
			   End If %>
			   	</div>
			
			<% If objUserList(19) <> "" Then %>
					<div><%=Replace(objUserList(19),vbCRLF,"<br />")%> </div>
			<% End If %>
			
			<% If objUserList(15) <> "" Then %>
					<div><b>User Notes</b>: <%=Replace(objUserList(15),vbCRLF,"<br />")%> </div>
			<% End If %>
			</div>
			
   
      <% objUserList.MoveNext
      Loop %>
  		</div>     
<% End If 
   
End Sub%>

<%Sub ShowUserTable 

	Dim strSQL, objDeviceList, strDeviceAssetTagList, strDeviceList, strRowClass, objFSO, intDaysRemaining, arrPWordLastSet
	Dim objOwedList, strOwedList, strOwedDate, strLatestBilled
	
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	
	If IsMobile Then %>
		<div class="ViewButtonMobile">
<% Else %>
		<div class="ViewButton">
<% End If %>
		<a href="<%=SwitchView("Card")%>"><img src="../images/card.png" title="Card View" height="32" width="32"/></a>
	</div>
	<div class="center"><%=Replace(FilterBar,"?","?View=Table&")%></div>
	<div>
		<table align="center" Class="ListView" id="ListView">
			<thead>
			<th>Photo</th>
			<th>ID</th>
			<th>Name</th>
			<th>First Name</th>
			<th>Last Name</th>
			<th>User Name</th>
			<th>Email</th>
			<th>Description</th>
		<% If Application("ShowPasswords") Then %>
				<th>Password</th>
				<th>AUP</th>
		<% End If %>
			<th>Password Changed</th>
			<th>Days Remaining</th>
			<th>Site</th>
			<th>Room</th>
			<th>Phone</th>
			<th><%=Application("HomeroomNameLong")%></th>
			<th>Class Of</th>
			<th>Role</th>
			<th>Assigned Device</th>
			<th>Assigned Tag</th>
			<th>User Notes</th>
			<th>Internet Access</th>
			<th>Billed Date</th>	
			<th>Latest Billed</th>
			<th>Owes</th>
			<th>External Access</th>
			<th>Internal Access</th>
			</thead>
			<tbody>
	<% Do Until objUserList.EOF 
	
			If objUserList(10) Then
				strRowClass = " Class=""Warning"""
			ElseIf objUserList(11) Then
				strRowClass = " Class=""Loaned"""
			ElseIf objUserList(9) Then
				strRowClass = ""
			Else 
				strRowClass = " Class=""Disabled"""
			End If %>
			
			<tr <%=strRowClass%>>
			<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objUserList(7) & "s\" & objUserList(4) & ".jpg") Then %>   
					<td <%=strRowClass%> width="1px">
						<a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>">   
							<img src="/photos/<%=objUserList(7)%>s/<%=objUserList(4)%>.jpg" title="<%=objUserList(4)%>" width="72" />
						</a>
					</td>
			<% Else %>
					<td <%=strRowClass%> width="1px">
						<a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>">
							<img src="/photos/<%=objUserList(7)%>s/missing.png" title="<%=objUserList(4)%>" width="72" />
						</a>
					</td>
			<% End If %>
				<td <%=strRowClass%> id="center"><%=objUserList(4)%></td>
				<td <%=strRowClass%>><a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>"><%=objUserList(2)%>, <%=objUserList(1)%></a></td>
				<td <%=strRowClass%>><a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>"><%=objUserList(1)%></a></td>
				<td <%=strRowClass%>><a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>"><%=objUserList(2)%></a></td>
				<td <%=strRowClass%>><%=objUserList(3)%></td>
				<td <%=strRowClass%>><%=objUserList(3)%>@<%=Application("Domain")%></td>
				<td <%=strRowClass%>><%=objUserList(19)%></td>
			<% If Application("ShowPasswords") Then %>	
					<td <%=strRowClass%>><%=objUserList(12)%></td>
				<% If objUserList(13) Then %>
						<td <%=strRowClass%> id="center">Yes</td>
				<% Else %>
						<td <%=strRowClass%> id="center">No</td>
				<% End If %>
			<% End If 
				If Not IsNull(objUserList(16)) Then
					arrPWordLastSet = Split(objUserList(16)," ")
					If CDate(arrPWordLastSet(0)) > #1/1/80# Then %>
						<td <%=strRowClass%>><%=objUserList(16)%></td>
				<% Else %>
						<td <%=strRowClass%>>---</td>
				<% End If 
				Else %>
					<td <%=strRowClass%>>---</td>
			<%	End If %>
			<% If Not objUserList(20) Then 'Password doesn't expire
				   If Not IsNull(objUserList(16)) Then
						intDaysRemaining = DateDiff("d",Date(),DateAdd("d",Application("PasswordsExpire"),objUserList(16)))
				
						If intDaysRemaining > 10 Then %>
							<td id="center" <%=strRowClass%>><%=intDaysRemaining%></td>
					<% Else 
							If intDaysRemaining <= 0 Then %>
								<td id="center" Class="Disabled">Expired</td>
						<% Else %>
								<td id="center" Class="Disabled"><%=intDaysRemaining%></td>
						<% End If %>
					<% End If 
					Else %>
						<td id="center">---</td>
				<% End If %>
			<% Else %>
					<td id="center">---</td>
			<% End If %>	
				
				<td <%=strRowClass%>><%=objUserList(14)%></td>
				<td <%=strRowClass%>><%=objUserList(18)%></td>
				<td <%=strRowClass%>><%=objUserList(17)%></td>
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
							strDeviceList = strDeviceList & "<a href=""device.asp?Tag=" & objDeviceList(0) & strBackLink & """>" & objDeviceList(1) & "</a>, "
							strDeviceAssetTagList = strDeviceAssetTagList & "<a href=""device.asp?Tag=" & objDeviceList(0) & strBackLink & """>" & objDeviceList(0) & "</a>, "
							objDeviceList.MoveNext
						Loop 
						strDeviceList = Left(strDeviceList,Len(strDeviceList) - 2)
						strDeviceAssetTagList = Left(strDeviceAssetTagList,Len(strDeviceAssetTagList) - 2)
					End If
				End If %>
				<td <%=strRowClass%>><%=strDeviceList%></td>
				<td <%=strRowClass%>><%=strDeviceAssetTagList%></td>
			<% If NOT IsNull(objUserList(15)) Then %>
					<td <%=strRowClass%>><%=Replace(objUserList(15),vbCRLF,"<br />")%></td>
			<% Else %>
					<td <%=strRowClass%>><%=objUserList(15)%></td>
			<% End If%>	
			<% If NOT IsNull(objUserList(24)) Then %>
					<td <%=strRowClass%>><%=objUserList(24)%></td>
			<% Else %>
					<td <%=strRowClass%>></td>
			<% End If%>	
			
			<% If objUserList(10) Then 
					strSQL = "Select Item,Price,RecordedDate FROM Owed WHERE Active=True AND OwedBy=" & objUserList(0)
					Set objOwedList = Application("Connection").Execute(strSQL)
					strOwedDate = ""
					strOwedList = ""
					strLatestBilled = ""
					Do Until objOwedList.EOF
						strOwedDate = strOwedDate & ShortenDate(objOwedList(2)) & " <br />"
						strOwedList = strOwedList & objOwedList(0) & " - $" & objOwedList(1) & " <br />"
						strLatestBilled =  ShortenDate(objOwedList(2))
						objOwedList.MoveNext
					Loop
					strOwedList = Left(strOwedList,Len(strOwedList) - 7)%>
					<td <%=strRowClass%>><%=strOwedDate%></td>
					<td <%=strRowClass%>><%=strLatestBilled%></td>
					<td <%=strRowClass%>><%=strOwedList%></td>
			<% Else %>
					<td <%=strRowClass%>>&nbsp;</td>
					<td <%=strRowClass%>>&nbsp;</td>
					<td <%=strRowClass%>>&nbsp;</td>
		 	<% End If %>
		 		<td <%=strRowClass%> id="center"><%=objUserList(21)%></td>
		 		<td <%=strRowClass%> id="center"><%=objUserList(22)%></td>
			</tr>
		<% objUserList.MoveNext
		Loop %>
			</tbody>
		</table>
<%End Sub%>

<%Sub DrawIcon(strIconType,intPosition,strURLData)

	Dim strPosition

	Select Case intPosition
		Case 0
			strPosition = "Left"
		Case 1
			strPosition = "One"
		Case 2
			strPosition = "Two"
		Case 3
			strPosition = "Three"
	End Select

	Select Case strIconType
		Case "HelpDesk" %>
			<div class="Button<%=strPosition%>">
				<a href="<%=Application("HelpDeskURL")%>/index.asp?UserName=<%=strURLData%>">
					<image src="../images/helpdesk.png" width="20" height="20" title="Enter Help Desk Ticket" />
				</a>
			</div>
	<%	Case "Back" %>
			<div class="Button<%=strPosition%>">
				<a href="<%=Request.QueryString("Page")%>?<%=Request.QueryString("Back")%>">
					<image src="../images/back.png" width="20" height="20" title="Return to Search Results"/>
				</a>
			</div>
<%	End Select

End Sub %>

<%Sub JumpToDevice%>

	<div Class="HeaderCard">
		<form method="POST" action="search.asp">
		Asset tag: <input class="Card InputWidthSmall" type="text" name="SmartBox" id="LastNames" />
		<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		<input type="hidden" value="" name="BOCESTag" />
		</form>
	</div>

<%End Sub%>

<%Function FilterBar
   
   If strSite <> "" Then
   	FilterBar = FilterBar & "Site = <a href=""users.asp?UserSite=" & strSite & """>" & strSite & "</a> | "
   End If
   
   If strFirstName <> "" Then
   	FilterBar = FilterBar & "First Name = <a href=""users.asp?FirstName=" & strFirstName & """>" & strFirstName & "</a> | "
   End If
   
   If strLastName <> "" Then
   	FilterBar = FilterBar & "Last Name = <a href=""users.asp?LastName=" & strLastName & """>" & strLastName & "</a> | "
   End If
   
   If strGuideRoom <> "" Then
   	FilterBar = FilterBar & Application("HomeroomNameLong") & " = <a href=""users.asp?GuideRoom=" & strGuideRoom & """>" & strGuideRoom & "</a> | "
   End If
   
   If strRole <> "" Then
		Select Case strRole
			Case "Adult", "Student"
				FilterBar = FilterBar & "Role = <a href=""users.asp?Role=" & strRole & """>" & strRole & "</a> | "
			Case Else
				FilterBar = FilterBar & "Role = <a href=""users.asp?Role=" & strRole & """>" & GetRole(strRole) & "</a> | "
		End Select
	End If
   
   If strWithDevice <> "" Then
   	FilterBar = FilterBar & "With Device = <a href=""users.asp?WithDevice=" & strWithDevice & """>" & strWithDevice & "</a> | "
   End If
   
   If strMissing <> "" Then
   	FilterBar = FilterBar & "Missing = <a href=""users.asp?WithDevice=" & strMissing & """>" & strMissing & "</a> | "
   End If
   	
   If strLoanedOut <> "" Then
   	FilterBar = FilterBar & "Loaned = <a href=""users.asp?LoanedOut=" & strLoanedOut & """>" & strLoanedOut & "</a> | "
   End If 
   
   If strStatus <> "" Then
   	FilterBar = FilterBar & "Status = <a href=""users.asp?UserStatus=" & strStatus & """>" & strStatus & "</a> | "
   End If
   
   If strOwes <> "" Then
   	FilterBar = FilterBar & "Owes = <a href=""users.asp?Owes=" & strOwes & """>" & strOwes & "</a> | "
   End If
   
   If strAUP <> "" Then
   	FilterBar = FilterBar & "AUP = <a href=""users.asp?AUP=" & strAUP & """>" & strAUP & "</a> | "
   End If
   
   If strInternetAccess <> "" Then
   	FilterBar = FilterBar & "Internet = <a href=""users.asp?InternetAccess=" & strInternetAccess & """>" & strInternetAccess & "</a> | "
   End If
   
   If strDescription <> "" Then
   	FilterBar = FilterBar & "Description = <a href=""users.asp?Description=" & strNotes & """>" & strDescription & "</a> | "
   End If
   
   If strNotes <> "" Then
   	FilterBar = FilterBar & "Notes = <a href=""users.asp?UserNotes=" & strNotes & """>" & strNotes & "</a> | "
   End If
   
   If FilterBar <> "" Then
   	FilterBar = Left(FilterBar,Len(FilterBar) - 3)
   End If

End Function %>

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

<%Sub SaveAsSearch 

	Dim strPage

	strPage = Right(Request.ServerVariables("SCRIPT_NAME"),Len(Request.ServerVariables("SCRIPT_NAME")) - InStrRev(Request.ServerVariables("SCRIPT_NAME"),"/"))
	strPage = strPage & "?" & Request.ServerVariables("QUERY_STRING")%>
	
	<br />
	
	<div class="Card NormalCard">
      <form method="POST" action="<%=strPage%>">
     	 	<div class="CardTitle">Save Search</div>
      	<div Class="CardColumn1">Search Name:</div>
         <div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="SearchName""/>
			</div>
         <div>
            <div class="Button"><input type="submit" value="Save" name="Submit" /></div>
         </div>
      <% If strSearchMessage <> "" Then %>
   		<div>
   			<%=strSearchMessage%>
   		</div>
   <% End If %> 
      </form>
      </div>

<%End Sub%>

<%Sub SaveSearch

	Dim strPage, strSearchName, strQueryString, strSQL
	
	strPage = Right(Request.ServerVariables("SCRIPT_NAME"),Len(Request.ServerVariables("SCRIPT_NAME")) - InStrRev(Request.ServerVariables("SCRIPT_NAME"),"/"))
	strQueryString = Replace(Replace(Request.ServerVariables("QUERY_STRING"),"'","''")," ","%20")
	strSearchName = Request.Form("SearchName")
	
	strSQL = "INSERT INTO SavedSearches (SearchName,Page,QueryString,UserName,Active) VALUES (" 
	strSQL = strSQL & "'" & Replace(strSearchName,"'","''") & "',"
	strSQL = strSQL & "'" & strPage & "',"
	strSQL = strSQL & "'" & strQueryString & "',"
	strSQL = strSQL & "'" & strUser & "',"
	strSQL = strSQL & "True)"
	Application("Connection").Execute(strSQL)
	
	UpdateLog "SearchSaved","","","",strSearchName,""
	
	strSearchMessage = "<div Class=""Information"">Saved</div>"

End Sub%>

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

<%Function GetUserName(UserID)

	Dim strSQL, objUserInfo
	
	If UserID <> "" Then
		
		strSQL = "SELECT UserName FROM People WHERE ID=" & UserID
		Set objUserInfo = Application("Connection").Execute(strSQL)
		
		If Not objUserInfo.EOF Then
			GetUserName = objUserInfo(0)
		Else
			GetUserName = UserID
		End If
	Else
		GetUserName = UserID
	End If

End Function%>

<%Function GetDisplayName(Username)

	Dim strSQL, objUserInfo

	If Username <> "" Then
		
		strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & Username & "'"
		Set objUserInfo = Application("Connection").Execute(strSQL)
		
		If Not objUserInfo.EOF Then
			GetDisplayName = objUserInfo(1) & ", " & objUserInfo(0)
		Else
			GetDisplayName = Username
		End If
	Else
		GetDisplayName = Username
	End If

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
	strSQL = strSQL & Replace(strUser,"'","''") & "',#"
	strSQL = strSQL & datDate & "#,#"
	strSQL = strSQL & datTime & "#,True,False," & intEventNumber & ")"
	Application("Connection").Execute(strSQL)
	
End Sub%>

<%Function GetRole(intYear)

   Dim datToday, intMonth, intCurrentYear, intGrade, strSQL, objRole
   
   'If they're an adult then get their role from the database
   If intYear <= 1000 Then
   	strSQL = "SELECT Role FROM Roles WHERE RoleID=" & intYear
   	Set objRole = Application("Connection").Execute(strSQL)
   	
   	If Not objRole.EOF Then
   		GetRole = objRole(0)
   	End If
   End If 
      
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
				GetRole = "Kindergarten Student"
			Case 1
				GetRole = "1st Grade Student"
			Case 2
				GetRole = "2nd Grade Student"
			Case 3
				GetRole = "3rd Grade Student"
			Case 4
				GetRole = "4th Grade Student"   
			Case 5
				GetRole = "5th Grade Student"
			Case 6
				GetRole = "6th Grade Student"
			Case 7
				GetRole = "7th Grade Student"
			Case 8
				GetRole = "8th Grade Student"
			Case 9
				GetRole = "9th Grade Student"
			Case 10
				GetRole = "10th Grade Student"
			Case 11
				GetRole = "11th Grade Student"
			Case 12
				GetRole = "12th Grade Student"
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
      <link rel="stylesheet" type="text/css" href="../style.css" /> 
      <link rel="apple-touch-icon" href="../images/inventory.png" /> 
      <link rel="shortcut icon" href="../images/inventory.ico" />
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

   If strRole = "Admin" Then
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
   
   If Application("MultiColumn") Then
  		strColumns = "MultiColumn"
  	End If
   
End Sub%>