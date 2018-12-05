<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/29/15
'Last Updated 1/14/18

'This page shows the details for a single user in the inventory website

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim strUserName, objDeviceList, objUser, objMissingStuff, strSubmitTo, strBackLink
Dim objLoanedOut, objItems, intUserID, strCardType, strColumns, strNewAssignmentMessage
Dim intActiveAssignmentCount, intOldAssignmentCount, strUserMessage, objEvents, objLog
Dim intActiveEventCount, intOldEventCount, strViewAllToggle, intViewAllCounter, objLastNames
Dim objSites, objRoles, objRooms, objDescriptions, strDeviceInfo, objPurchasableItems
Dim objOwes, strUserInfo, objInternetTypes
Dim deviceOn, WshShell, PINGFlag, ipAddress, status, strNotifyMessage

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions

	Dim strSQL, intAge

	'Get the variables from the URL
	strUserName = Request.QueryString("UserName")
	strBackLink = BackLink

	'If nothing was submitted send them back to the index page
   If strUserName = "" Then
      Response.Redirect("index.asp?Error=UserNotFound")
   End If

	'Get the user's information
	strSQL = "SELECT ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,Active,Warning,PWord,AUP,Notes,PWordLastSet,PhoneNumber,RoomNumber,Description,Site,PWordNeverExpires,LastExternalCheckIn,LastInternalCheckIn,Birthday,InternetAccess" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE UserName='" & Replace(strUserName,"'","''") & "' AND NOT Deleted" & vbCRLF
	Set objUser = Application("Connection").Execute(strSQL)

	'Build the user info popup
	If objUser(20) <> "" Then
		strUserInfo = "Internal Access: " & objUser(20) & " &#013 "
	End If
   	If objUser(19) Then
   		strUserInfo = strUserInfo & "External Access: " & objUser(19) & " &#013 "
   	End If
   	If objUser(21) <> "" Then
		intAge = DateDiff("yyyy",objUser(21),Date)
		If Date < DateSerial(Year(Date), Month(objUser(21)), Day(objUser(21))) Then
			intAge = intAge - 1
		End If
   		strUserInfo = strUserInfo & "Birthday: " & objUser(21) & " &#013 "
   		strUserInfo = strUserInfo & "Age: " & intAge
	End If

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Return"
         Return
      Case "Loan Out Item","Loan"
      	LoanOutItem
      Case "Update User"
      	UpdateUser
      Case "Assign Device"
      	AssignDevice
      Case "Restore User"
      	RestoreUser
      Case "Disable User"
      	DisableUser
      Case "Bill"
      	BillUser
      Case "Paid"
      	UserPaid
      Case "PasswordExpire"
      	SetUserPasswordToExpire
      Case "PasswordNotExpire"
      	SetUserPasswordNotToExpire
      Case "Change Password"
      	UpdateUserPassword
      Case "Notify"
      	If IsEmailValid(Request.Form("NotifyEmail")) Then
      		EMailGuardian "OwesMoney"
      		strNotifyMessage = "<div Class=""Information"">Message Sent</div>"
      		UpdateLog "NotificationEMailSent","",strUserName,"",Request.Form("NotifyEmail"),""
      	Else
      		strNotifyMessage = "<div Class=""Error"">Invalid EMail Address</div>"
      	End If
   End Select

   'Send them back to the index page if the user isn't found
   If objUser.EOF Then
   	Response.Redirect("index.asp?Error=UserNotFound")
   End If

   intUserID = objUser(0)

   'Get the list of devices assigned to the user
   strSQL = "SELECT Assignments.LGTag,DateIssued,DateReturned,Assignments.Active,Assignments.Notes,Model,Devices.Active,IssuedBy,ReturnedBy,Devices.Deleted,Devices.HasInsurance,Devices.DatePurchased,SerialNumber,Manufacturer,InternalIP,LastUser,ComputerName,OSVersion,LastCheckInDate,LastCheckInTime,Assigned" & vbCRLF
	strSQL = strSQL & "FROM Devices INNER JOIN (People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo) ON Devices.LGTag = Assignments.LGTag" & vbCRLF
	strSQL = strSQL & "WHERE People.UserName='" & Replace(strUserName,"'","''") & "'" & vbCRLF
	strSQL = strSQL & "ORDER BY DateIssued"
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

	'Get the list of events associated with the user
	strSQL = "SELECT ID,Type,Notes,EventDate,EventTime,Resolved,ResolvedDate,ResolvedTime,Category,Warranty,LGTag,UserID,Site,Model,EnteredBy,CompletedBy " &_
		"FROM Events WHERE UserID=" & intUserID & " ORDER BY ID DESC"
   Set objEvents = Application("Connection").Execute(strSQL)

   'Count the number of active and old assignments
	intActiveEventCount = 0
	intOldEventCount = 0
	If Not objEvents.EOF Then
		Do Until objEvents.EOF
			If Not objEvents(5) Then
				intActiveEventCount = intActiveEventCount + 1
			Else
				intOldEventCount = intOldEventCount + 1
			End If
			objEvents.MoveNext
		Loop
		objEvents.MoveFirst
	End If

   'Get the list of loaned out items
   strSQL = "SELECT ID, Item, LoanDate FROM Loaned WHERE AssignedTo=" & objUser(0) & " AND Returned=False ORDER By LoanDate"
   Set objLoanedOut =  Application("Connection").Execute(strSQL)

   'Get the list of items to loan out
   strSQL = "SELECT Item FROM Items WHERE Active=True ORDER BY Item"
   Set objItems =  Application("Connection").Execute(strSQL)

   'Get the list of things they owe money for
   strSQL = "SELECT ID,Item,Price,RecordedDate,Returnable FROM Owed WHERE Active=True AND OwedBy=" & objUser(0) & " ORDER BY RecordedDate"
   Set objOwes = Application("Connection").Execute(strSQL)

   'Get the list of things they could owe money for
   strSQL = "SELECT Item, Price FROM Purchasable WHERE Active=True ORDER BY Item"
	Set objPurchasableItems = Application("Connection").Execute(strSQL)

   'Get the log items for the user
   strSQL = "SELECT LGTag,UserName,EventNumber,Type,OldValue,NewValue,UpdatedBy,LogDate,LogTime,OldNotes,NewNotes" & vbCRLF
   strSQL = strSQL & "FROM Log WHERE Active=True AND Deleted=False And UserName='" & Replace(strUserName,"'","''") & "' ORDER BY ID DESC"
   Set objLog = Application("Connection").Execute(strSQL)

   'Get the list of lastnames for the auto complete
   strSQL = "SELECT DISTINCT LastName FROM People WHERE Active=True"
   Set objLastNames = Application("Connection").Execute(strSQL)

   'Get the list of rooms for the auto complete
   strSQL = "SELECT DISTINCT RoomNumber FROM People WHERE Active=True And RoomNumber<>'' And Role='Teacher'"
   Set objRooms = Application("Connection").Execute(strSQL)

   'Get the list of descriptions for the auto complete
   strSQL = "SELECT DISTINCT Description FROM People WHERE Active=True And Description<>'' And Role='Teacher'"
   Set objDescriptions = Application("Connection").Execute(strSQL)

   'Get the data for the sites drop down menu
   strSQL = "SELECT Site FROM Sites WHERE Active=True ORDER BY Site"
   Set objSites = Application("Connection").Execute(strSQL)

   'Get the list of roles for the drop down menu
   strSQL = "SELECT Role,RoleID FROM Roles WHERE Active=True ORDER BY Role"
   Set objRoles = Application("Connection").Execute(strSQL)
   
   'Get the list of Internet types for the drop down menu
   strSQL = "SELECT InternetType FROM InternetTypes WHERE Active=True"
   Set objInternetTypes = Application("Connection").Execute(strSQL)

   'Set the if condition for the view all icon
   strViewAllToggle = ""
   intViewAllCounter = 0
	If intOldAssignmentCount > 0 Then
		strViewAllToggle = strViewAllToggle & "$(oldAssignments).is("":visible"") && "
		intViewAllCounter = intViewAllCounter + 1
	End If

	If Not objEvents.EOF Then
		strViewAllToggle = strViewAllToggle & "$(events).is("":visible"") && "
		intViewAllCounter = intViewAllCounter + 1
	End If

	If Not objLog.EOF Then
		strViewAllToggle = strViewAllToggle & "$(userlog).is("":visible"") && "
		intViewAllCounter = intViewAllCounter + 1
	End If
	If Len(strViewAlltoggle) > 0 Then
		strViewAllToggle = Left(strViewAllToggle,Len(strViewAllToggle) - 4)
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
		<script src="../assets/js/jquery.flip.min.js"></script>
		<script type="text/javascript">
			$(document).ready( function () {

			<%	If Not IsMobile And Not IsiPad Then%>
					$( document ).tooltip({track: true});
			<% End If %>

				document.getElementById("ChangePasswordButton").disabled = true;

			//	$("#card").flip({
			//		axis: 'y'
			//	});

				var oldAssignments = document.getElementById("OldAssignments");
				var events = document.getElementById("Events");
				var userlog = document.getElementById("UserLog");
				var disableUser = document.getElementById("disableUser");
				var roomView = document.getElementById("RoomView");
				var roomEdit = document.getElementById("RoomEdit");
				var phoneView = document.getElementById("PhoneView");
				var phoneEdit = document.getElementById("PhoneEdit");
				var descriptionView = document.getElementById("DescriptionView");
				var descriptionEdit = document.getElementById("DescriptionEdit");
				var rolesView = document.getElementById("RolesView");
				var rolesEdit = document.getElementById("RolesEdit");
				var siteEdit = document.getElementById("SiteEdit");
				var studentID = document.getElementById("StudentID");
				var passwordInfo = document.getElementById("PasswordInfo");
				var nameView = document.getElementById("NameView");
				var nameEdit = document.getElementById("NameEdit");
				var passwordView = document.getElementById("PasswordView");
				var passwordEdit = document.getElementById("PasswordEdit");
				var internetView = document.getElementById("InternetView");
				var internetEdit = document.getElementById("InternetEdit");
				var notifyArea = document.getElementById("Notify");
				var showHideEffect = "blind";
				var effectSpeed = 200;


				$(notifyArea).hide();
				$("#emailToggle").click(function(){
					if ($(notifyArea).is(":visible")) {
						$(notifyArea).hide(showHideEffect,{},effectSpeed);
					} else {
						$(notifyArea).show(showHideEffect,{},effectSpeed);
					}
					return false;
				});


				$("#assignmentsToggle").click(function(){
					if ($(oldAssignments).is(":visible")) {
						$(oldAssignments).hide(showHideEffect,{},effectSpeed);
					} else {
						$(events).hide(showHideEffect,{},effectSpeed);
						$(userlog).hide(showHideEffect,{},effectSpeed);
						$(oldAssignments).show(showHideEffect,{},effectSpeed);
					}
					return false;
				});

				$("#eventsToggle").click(function(){
					if ($(events).is(":visible")) {
						$(events).hide(showHideEffect,{},effectSpeed);
					} else {
						$(oldAssignments).hide(showHideEffect,{},effectSpeed);
						$(userlog).hide(showHideEffect,{},effectSpeed);
						$(events).show(showHideEffect,{},effectSpeed);
					}
					return false;
				});

				$("#logToggle").click(function(){
					if ($(userlog).is(":visible")) {
						$(userlog).hide(showHideEffect,{},effectSpeed);
					} else {
						$(oldAssignments).hide(showHideEffect,{},effectSpeed);
						$(events).hide(showHideEffect,{},effectSpeed);
						$(userlog).show(showHideEffect,{},effectSpeed);
					}
					return false;
				});
			<% If Len(strViewAlltoggle) > 0 Then %>
					$("#viewAllToggle").click(function(){
						if (<%=strViewAllToggle%>) {
							$(oldAssignments).hide(showHideEffect,{},effectSpeed);
							$(events).hide(showHideEffect,{},effectSpeed);
							$(userlog).hide(showHideEffect,{},effectSpeed);
						} else {
							$(oldAssignments).show(showHideEffect,{},effectSpeed);
							$(events).show(showHideEffect,{},effectSpeed);
							$(userlog).show(showHideEffect,{},effectSpeed);
						}
						return false;
					});
			<% End If %>

				$(oldAssignments).hide();
				$(events).hide();
				$(userlog).hide();
				$(disableUser).hide();
				$(roomEdit).hide();
				$(phoneEdit).hide();
				$(descriptionEdit).hide();
				$(rolesEdit).hide();
				$(siteEdit).hide();
				$(studentID).hide();
				$(nameEdit).hide();
				$(passwordEdit).hide();
				$(internetEdit).hide();
				$('#body').show();

				$("#editToggle").click(function(){
					if ($(disableUser).is(":visible")) {
						$(disableUser).hide();
						$(roomEdit).hide();
						$(phoneEdit).hide();
						$(descriptionEdit).hide();
						$(rolesEdit).hide();
						$(siteEdit).hide();
						$(studentID).hide();
						$(nameEdit).hide();
						$(internetEdit).hide();

						$(roomView).show();
						$(phoneView).show();
						$(descriptionView).show();
						$(rolesView).show();
						$(passwordInfo).show();
						$(nameView).show();
						$(internetView).show();
					} else {
						$(disableUser).show();
						$(roomEdit).show();
						$(phoneEdit).show();
						$(descriptionEdit).show();
						$(rolesEdit).show();
						$(siteEdit).show();
						$(studentID).show();
						$(nameEdit).show();
						$(internetEdit).show();

						$(roomView).hide();
						$(phoneView).hide();
						$(descriptionView).hide();
						$(rolesView).hide();
						$(passwordInfo).hide();
						$(nameView).hide();
						$(internetView).hide();
					}

					return false;
				});

				$("#editStudentToggle").click(function(){
					if ($(passwordEdit).is(":visible")) {
						$(passwordEdit).hide();
						$(internetEdit).hide();

						$(passwordView).show();
						$(internetView).show();

					} else {
						$(passwordEdit).show();
						$(internetEdit).show();

						$(passwordView).hide();
						$(internetView).hide();
					}

					return false;
				});

    			var table = $('#ListView').DataTable( {
    				paging: false,
    				"info": false,
    				"autoWidth": false,
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
							title: 'Old Assignments - <%=Replace(strUserName,"'","\'")%>'
						}
				<% End If %>
        			]

    			});

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

				<% If Not objRooms.EOF Then %>
						var possibleRooms = [
					<% Do Until objRooms.EOF %>
							"<%=objRooms(0)%>",
						<%	objRooms.MoveNext
						Loop %>
					];
						$( "#Rooms" ).autocomplete({
							source: possibleRooms
						});
				<% End If %>

				<% If Not objDescriptions.EOF Then %>
						var possibleDescriptions = [
					<% Do Until objDescriptions.EOF %>
							"<%=objDescriptions(0)%>",
						<%	objDescriptions.MoveNext
						Loop %>
					];
						$( "#Descriptions" ).autocomplete({
							source: possibleDescriptions
						});
				<% End If %>

    	<% If IsMobile Then %>
    			table.columns([0,2,5,6,7]).visible(false);
    	<% Else %>
				table.columns([0]).visible(false);
		<% End If %>

    		} );

			$(document).ready( function () {
    			var eventTable = $('#EventTable').DataTable( {
    				paging: false,
    				"info": false,
    				"autoWidth": false,
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
							title: 'Events - <%=Replace(strUserName,"'","\'")%>'
						}
				<% End If %>
        			]
    			})

    	<% If IsMobile Then %>
    			eventTable.columns([0,4,5,6,7,8,9,10,11,12]).visible(false);
    	<% Else %>
				eventTable.columns([4,5,7,9,10,11]).visible(false);
		<% End If %>

    		} );

			$(document).ready( function () {
    			var logTable = $('#LogTable').DataTable( {
    				paging: false,
    				"info": false,
    				"autoWidth": false,
    				dom: 'Bfrtip',
    				"order": [],
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
							title: 'Log - <%=Replace(strUserName,"'","\'")%>'
						}
				<% End If %>
        			]
    			})

    	<% If IsMobile Then %>
				logTable.columns([0,1,3,4,6,7]).visible(false);
    	<% Else %>
				logTable.columns([4,7]).visible(false);
		<% End If %>

    		} );

    		$(document).ready( function () {

				var overlay = document.getElementById('changePasswordOverlay');
				var changePassword = document.getElementById("changePassword");
				var span = document.getElementsByClassName("close")[0];
				var password = document.getElementById("NewPassword");
		  		var confirmPassword = document.getElementById("ConfirmNewPassword");
		  		var passwordValid = document.getElementById("ValidIcon");
		  		var passwordMatch = document.getElementById("MatchIcon");
		  		var adminUsername = document.getElementById("AdminUserName");
		  		var adminPassword = document.getElementById("AdminPassword");
		  		var requireChange = document.getElementById("RequireChange");

				changePassword.onclick = function() {
					 overlay.style.display = "block";
					 return false;
				}

				span.onclick = function() {
					 overlay.style.display = "none";
					 password.value = "";
					 confirmPassword.value = "";
					 adminUsername.value = "";
					 adminPassword.value = "";
					 requireChange.checked = false;
					 passwordValid.src="../images/notgood.png";
					 passwordMatch.src="../images/notgood.png";
					 document.getElementById("ChangePasswordButton").disabled = true;
				}

				window.onclick = function(event) {
					 if (event.target == overlay) {
						  overlay.style.display = "none";
					 }
				}

    		} );

    		function checkNewPasswordForm() {

		  		var adminUsername = document.getElementById("AdminUserName");
		  		var adminPassword = document.getElementById("AdminPassword");
		  		var password = document.getElementById("NewPassword");
		  		var confirmPassword = document.getElementById("ConfirmNewPassword");
		  		var requireChange = document.getElementById("RequireChange");
		  		var passwordValid = document.getElementById("ValidIcon");
		  		var passwordMatch = document.getElementById("MatchIcon");
		  		var changePasswordButton = document.getElementById("ChangePasswordButton");
		  		var lengthGood;
         	var passwordsMatch;

         	if ((password.value.length >= 8) && (password.type == "password")) {
					passwordValid.src="../images/good.png";
					lengthGood = true;
				} else {
					passwordValid.src="../images/notgood.png";
					lengthGood = false;
				}

				if ((password.value == confirmPassword.value) && (password.value != "")) {
					passwordMatch.src="../images/good.png";
					passwordsMatch = true;
				} else {
					passwordMatch.src="../images/notgood.png";
					passwordsMatch = false;
				}

				if ((lengthGood) && (passwordsMatch)) {
					if ((adminUsername.value != "") && (adminPassword.value != "")) {
						changePasswordButton.disabled = false;
					} else {
						changePasswordButton.disabled = true;
					}
				}

		  	};

    	</script>
    	<script type="text/javascript">

			function setSubmitValue(value) {
					jQuery('#mouseOnValue').val(value);
				}
		</script>
   </head>

   <body class="<%=strSiteVersion%>" id="body" style="display:none;">

      <div class="Header"><%=Application("SiteName")%></div>
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
      <div Class="<%=strColumns%>">

		<%
		JumpToDevice
		UserCard

		If intActiveAssignmentCount >= 1 Then
			ActiveAssignments
		Else
			If objUser(8) Then
				AddAssignmentCard
			End If
		End If

		If objUSer(8) Then
			EquipmentCard
		Else
			If Not objOwes.EOF Or Not objLoanedOut.EOF Then
				EquipmentCard
			End If
		End If
		OldAssignements
		EventsTable
		ShowLog
		ChangePasswordCard

		%>
		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
   </body>

   </html>

<%End Sub%>

<%Sub ChangePasswordCard%>
	<form method="POST" action="<%=strSubmitTo%>">
		<input type="hidden" name="UserID" value="<%=intUserID%>" />
		<div id="changePasswordOverlay" class="overlay">

			<div class="overlay-content">
				<div class="CardTitle">
					<span class="close">&times;</span>
					Change Password
					</div>
				<div class="overlay-body">
					<div Class="CardMerged">Enter administrative credentials.</div>
					<div Class="CardColumn1">Username:</div>
					<div Class="CardColumn2">
						<input class="Card InputWidthLarge" type="text" name="AdminUserName" id="AdminUserName" oninput="checkNewPasswordForm()" />
					</div>
					<div Class="CardColumn1">Password:</div>
					<div Class="CardColumn2">
						<input class="Card InputWidthLarge" type="password" name="AdminPassword" id="AdminPassword" oninput="checkNewPasswordForm()" />
					</div>
					<div Class="CardMerged">Enter a new password for the user.</div>
					<div Class="CardColumn1">Password:</div>
					<div Class="CardColumn2">
						<input class="Card InputWidthLarge" type="password" name="Password" id="NewPassword" oninput="checkNewPasswordForm()" />
					</div>
					<div Class="CardColumn1">Confirm:</div>
					<div Class="CardColumn2">
						<input class="Card InputWidthLarge" type="password" name="ConfirmPassword" id="ConfirmNewPassword" oninput="checkNewPasswordForm()" />
					</div>
					<div Class="CardColumn1">Require Change:</div>
					<div Class="CardColumn2">
						<input class="Card" type="checkbox" name="RequireChange" id="RequireChange" oninput="checkNewPasswordForm()" />
					</div>

					<div>
						<div Class="CardMerged Center">Password Valid <image src="../images/notgood.png" class="Icon" id="ValidIcon" ></div>
					</div>
					<div>
						<div Class="CardMerged Center">Passwords Match <image src="../images/notgood.png" class="Icon" id="MatchIcon" ></div>
					</div>
					<div>
						<div class="Button"><input type="submit" value="Change Password" name="Submit" id="ChangePasswordButton" /></div>
					</div>
				</div>
			</div>
		</div>
	</form>

<%End Sub%>

<%Sub UserCard

	Dim objFSO, strCardType, arrPWordLastSet, intDaysRemaining, strAUPChecked, strSelected, strPasswordLastReset, strChangePassword
	Dim strPasswordChangeText, strPasswordResetText, strPasswordTextClass

	Set objFSO = CreateObject("Scripting.FileSystemObject")

   If Not objUser.EOF Then

		Do Until objUser.EOF

			If objUser(11) Then
				strAUPChecked = "checked=""checked"""
			End If

      	If objUser(9) Then
      		strCardType = "WarningCard"
         ElseIf Not objLoanedOut.EOF Then
      		strCardType = "LoanedCard"
         ElseIf objUser(8) Then
				strCardType = "NormalCard"
		   Else
				strCardType = "DisabledCard"
		   End If %>

		   <div class="Card <%=strCardType%>">
		   	<form method="POST" action="<%=strSubmitTo%>">
				<input type="hidden" name="UserID" value="<%=intUserID%>" />
		   	<button style="overflow: visible !important; height: 0 !important; width: 0 !important; margin: 0 !important; border: 0 !important; padding: 0 !important; display: block !important;" type="submit" name="Submit" value="Update User" /></button>
		   	<div class="CardTitle" id="NameView">
		   	<% If objUser(7) = "Student" Then %>
					<% If Application("ShowPasswords") Then %>
						<% If objUser(11) Then %>
								<image src="../images/yes.png" width="15" height="15" title="AUP Signed" />
						<% Else %>
								<image src="../images/no.png" width="15" height="15" title="AUP Not Signed" />
						<% End If %>
					<% End If %>
				<% End If %>
					<%=objUser(1) & " " & objUser(2)%>
				<% If strUserInfo <> "" Then %>
						<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strUserInfo%>"  />&nbsp;</div>
				<% End If %>
				</div>

				<div class="CardTitle" id="NameEdit">
		   	<% If objUser(7) = "Student" Then %>
					<% If Application("ShowPasswords") Then %>
						<% If objUser(11) Then %>
								<image src="../images/yes.png" width="15" height="15" title="AUP Signed" />
						<% Else %>
								<image src="../images/no.png" width="15" height="15" title="AUP Not Signed" />
						<% End If %>
					<% End If %>
				<% Else %>
						<input type="image" src="../images/disable.png" value="Disable User" id="disableUser" width="15" height="15" title="Disable User" onmouseover="setSubmitValue('Disable User')" />
				<% End If %>
						<input Class="Card InputWidthSmall" type="text" name="FirstName" value="<%=objUser(1)%>">
						<input Class="Card InputWidthSmall" type="text" name="LastName" value="<%=objUser(2)%>">
						<input Class="Card InputWidthSmall" type="text" name="NewUserName" value="<%=strUserName%>">
				</div>

				<div Class="ImageSectionInCard" id="card">
					<div class="front">
					<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objUser(7) & "s\" & objUser(4) & ".jpg") Then %>
							<img class="PhotoCard" src="/photos/<%=objUser(7)%>s/<%=objUser(4)%>.jpg" title="<%=objUser(4)%>" width="96" />
					<% Else %>
							<img class="PhotoCard" src="/photos/<%=objUser(7)%>s/missing.png" title="<%=objUser(4)%>" width="96" />
					<% End If %>
					</div>
<!--					<div class="back">
						Photo ID: <%=objUser(4)%>
					</div> -->
				</div>
				<div Class="RightOfImageInCard">
					<div>
						<div Class="PhotoCardColumn1">Role: </div>
						<div Class="PhotoCardColumn2Long" id="RolesView">
							<a href="users.asp?Role=<%=objUser(5)%>"><%=GetRole(objUser(5))%></a>
						</div>
					<% If CInt(objUser(5)) > 1000 Then %>
							<input type="hidden" name="Role" value="<%=CInt(objUser(5))%>" />
					<% Else %>
							<div Class="PhotoCardColumn2Long" id="RolesEdit">
								<select Class="Card" name="Role">
									<option value=""></option>
								<% Do Until objRoles.EOF
										If objUser(5) <> "" Then
											If CInt(objUser(5)) = CInt(objRoles(1)) Then
												strSelected = "selected=""selected"""
											Else
												strSelected = ""
											End If
										End If %>
											<option <%=strSelected%> value="<%=objRoles(1)%>"><%=objRoles(0)%></option>
								<%    objRoles.MoveNext
									Loop
									objRoles.MoveFirst%>
								</select>
							</div>
					<% End If %>
						<div id="SiteEdit">
							<div Class="PhotoCardColumn1">Site: </div>
								<div Class="PhotoCardColumn2Long" id="SiteEdit">
									<select Class="Card" name="Site">
										<option value=""></option>
									<% Do Until objSites.EOF
											If objUser(17) = objSites(0) Then
												strSelected = "selected=""selected"""
											Else
												strSelected = ""
											End If %>
												<option <%=strSelected%> value="<%=objSites(0)%>"><%=objSites(0)%></option>
									<%    objSites.MoveNext
										Loop
										objSites.MoveFirst%>
									</select>
								</div>
							</div>
							<div id="StudentID">
								 <div Class="PhotoCardColumn1">Photo:</div>
								 <div Class="PhotoCardColumn2Long"><input class="Card InputWidthSmall" type="text" name="PhotoID" Value="<%=objUser(4)%>" id="Makes" /></div>
							</div>

				<% If objUser(7) = "Student" Then
						If objUser(6) <> "" Then%>
							<div>
								<div Class="PhotoCardColumn1"><%=Application("HomeroomName")%>: </div>
								<div Class="PhotoCardColumn2Long">
									<a href="users.asp?GuideRoom=<%=objUser(6)%>"><%=objUser(6)%></a>
								</div>
							</div>
					<% End If
						If Application("ShowPasswords") Then %>
						<div>
							<div Class="CardMerged">Username: <%=strUserName%> </div>
						</div>
						<div>
							<div Class="CardMerged" id="PasswordView">Password: <%=objUser(10)%> </div>
							<div Class="CardMerged" id="PasswordEdit">Password: <input Class="Card InputWidthSmall" type="text" name="UserPassword" value="<%=objUser(10)%>"> </div>
						</div>

						<div>
							<div Class="CardMerged">AUP Signed: <input type="checkbox" name="AUP" value="True" <%=strAUPChecked%> /></div>
						</div>
					<% End If %>
					
				<% Else %>
						<div id="PasswordInfo">
					<% strPasswordChangeText = ""
						strPasswordResetText = ""
						strPasswordTextClass = ""
						strChangePassword = "<a href="""" class=""Button"" id=""changePassword""><image src=""../images/changepword.png"" width=""15"" height=""15"" title=""Change Password"" /></a>"

						If Not objUser(18) Then 'Password does expire; 18 = PWordNeverExpires

							If IsNull(objUser(13)) Then 'Password last set is empty, new user; 13 = PWordLastSet
								strPasswordLastReset = "6/16/78"
								strPasswordChangeText = "---"
								strPasswordResetText = "---"
								strPasswordTextClass = "CardMerged"

						   Else 'Password last set has a valid value
								strPasswordLastReset = objUser(13)
								arrPWordLastSet = Split(strPasswordLastReset," ")

								If Not IsNull(strPasswordLastReset) Then 'Not sure if this check is needed, oh well
									If CDate(arrPWordLastSet(0)) > #1/1/80# Then

										intDaysRemaining = DateDiff("d",Date(),DateAdd("d",Application("PasswordsExpire"),arrPWordLastSet(0)))
										strPasswordChangeText = ShortenDate(CDate(arrPWordLastSet(0)))
										strPasswordTextClass = "CardMerged"

									   If intDaysRemaining > 10 Then 'Display the days remaining with the right color
											strPasswordTextClass = "CardMerged"
									   ElseIf intDaysRemaining >= 1 Then
											strPasswordTextClass = "CardMerged Error"
									   Else
											strPasswordTextClass = "CardMerged Error"
											strPasswordResetText = "Expired"
									   End If

									Else 'Their password has probably expired
										strPasswordChangeText = "---"
										strPasswordResetText = "Expired"
										strPasswordTextClass = "CardMerged Error"

									End If
								Else 'Password last reset is null
									strPasswordChangeText = "---"
									strPasswordResetText = "---"
									strPasswordTextClass = "CardMerged"
								End If
							End If
						Else 'The password doesn't expire
							strPasswordLastReset = objUser(13)
							If Not IsNull(strPasswordLastReset) Then
								arrPWordLastSet = Split(strPasswordLastReset," ")

								If CDate(arrPWordLastSet(0)) > #1/1/80# Then
									strPasswordChangeText = ShortenDate(CDate(arrPWordLastSet(0)))
									strPasswordResetText = "---"
									strPasswordTextClass = "CardMerged"
							   Else
									strPasswordChangeText = "---"
									strPasswordResetText = "---"
									strPasswordTextClass = "CardMerged"
							   End If
						   Else
								strPasswordChangeText = "---"
								strPasswordResetText = "---"
								strPasswordTextClass = "CardMerged"
						   End If

					   End If

					   If strPasswordResetText = "" Then
					   	strPasswordResetText = intDaysRemaining
					   End If

					   %>

							<div Class="CardMerged">PWord Changed: <%=strPasswordChangeText%> <%=strChangePassword%></div>
							<div Class="<%=strPasswordTextClass%>">Days Remaining: <%=strPasswordResetText%></div>
						</div>
						<% If objUser(15) <> "" Then %>
								<div id="RoomView">
									<div Class="CardMerged">Room: <a href="devices.asp?Room=<%=objUser(15)%>&DeviceSite=<%=objUser(17)%>&View=Card"><%=objUser(15)%></a> </div>
								</div>
						<% End If %>
								<div id="RoomEdit">
									<div Class="CardMerged">Room: <input Class="Card InputWidthSmall" type="text" name="Room" id="Rooms" value="<%=objUser(15)%>"></div>
								</div>

						<% If objUser(14) <> "" Then %>
							<div id="PhoneView">
								<div Class="CardMerged">Phone: <%=objUser(14)%> </div>
							</div>
						<% End If %>
							<div id="PhoneEdit">
								<div Class="CardMerged">Phone: <input Class="Card InputWidthSmall" type="text" name="Phone" value="<%=objUser(14)%>"></div>
							</div>

					<% End If %>
				</div>
			</div>
			
			<% If objUser(7) = "Student" Then 
					If objUser(22) <> "" And Not IsNull(objUser(22)) Then %>
						<div id="InternetView">
							<div Class="CardMerged" >Internet: <%=objUser(22)%> </div>
						</div>
				<% End If %>
					
						<div id="InternetEdit">
							<div Class="CardMerged">Internet: 
								<select Class="Card" name="InternetAccess">
									<option value=""></option>
								<% Do Until objInternetTypes.EOF
										If objUser(22) = objInternetTypes(0) Then
											strSelected = "selected=""selected"""
										Else
											strSelected = ""
										End If %>
											<option <%=strSelected%> value="<%=objInternetTypes(0)%>"><%=objInternetTypes(0)%></option>
								<%    objInternetTypes.MoveNext
									Loop
									objInternetTypes.MoveFirst%>
								</select>
							</div>
						</div>
				
			<%	End If %>
			
			<% If objUser(16) <> "" Then %>
					<div id="DescriptionView">
						<div Class="CardMerged"><%=objUser(16)%> </div>
					</div>
			<% End If %>
					<div id="DescriptionEdit">
						<div Class="CardMerged"><input Class="Card InputWidthFull" type="text" name="Description" value="<%=objUser(16)%>" id="Descriptions"></div>
					</div>
			<div>User Notes: </div>
			<div>
				<textarea class="Card" rows="5" name="Notes" cols="90" style="width: 99%;"><%=objUser(12)%></textarea>
			</div>
			<br />
				<input type="hidden" name="Submit" value="" id="mouseOnValue" >
			<div>
				<div class="Button"><input type="image" src="../images/save.png" width="20" height="20" title="Update User" onmouseover="setSubmitValue('Update User')" /></div>
			</div>


		<% If objUser(8) Then %>
			<% If objUser(7) = "Student" Then %>
					<a href="" class="Button" id="editStudentToggle">
						<image src="../images/edit.png" height="20" width="20" title="Toggle Edit Mode">
					</a>
			<% Else %>
					<a href="" class="Button" id="editToggle">
						<image src="../images/edit.png" height="20" width="20" title="Toggle Edit Mode">
					</a>
			<% End If %>
		<% Else %>
				<div>
					<div class="Button"><input type="image" src="../images/restore.png" width="20" height="20" title="Restore User" onmouseover="setSubmitValue('Restore User')" /></div>
				</div>
		<% End If %>

		<% If objUser(5) <= 1000 Then 'Only display the icon if they are an adult%>
			<% If objUser(8) Then 'Only display the icon if they are active%>
				<% If objUser(18) Then 'Password doesn't expire%>
						<div>
							<div class="Button"><input type="image" src="../images/pwordnotexpire.png" value="PasswordExpire" width="20" height="20" title="Toggle Password Expires" onmouseover="setSubmitValue('PasswordExpire')" /></div>
						</div>
				<% Else %>
						<div>
							<div class="Button"><input type="image" src="../images/pwordexpire.png" value="PasswordNotExpire" width="20" height="20" title="Toggle Password Expires" onmouseover="setSubmitValue('PasswordNotExpire')" /></div>
						</div>
				<% End If %>
			<% End If %>
		<% End If %>

		<% If Request.QueryString("Back") <> "" Then
				DrawIcon "Back",0,""
			End If

			If intViewAllCounter > 1 Then
				DrawIcon "ViewAll",0,""
			End If

			If Application("HelpDeskURL") <> "" Then
				DrawIcon "HelpDesk",0,objUser(3)
			End If

		   If intOldAssignmentCount > 0 Then
		   	DrawIcon "Assignments",0,""
		   	Response.Write "<div class=""ButtonText"">" & intOldAssignmentCount & "</div>"
		   End If

		   If Not objEvents.EOF Then
		   	DrawIcon "Events",0,""
		   	Response.Write "<div class=""ButtonText"">" & intActiveEventCount & "/" & intOldEventCount & "</div>"
		   End If

			If Not objLog.EOF Then
				DrawIcon "Log",0,""
		   End If %>

			<% If strUserMessage <> "" Then %>
					<%=strUserMessage%>
			<% End If %>

				</form>
			</div>

      <% objUser.MoveNext
      Loop
      objUser.MoveFirst
   End If %>

<%End Sub%>

<%Sub AssignmentCards

	Dim objFSO

	Set objFSO = CreateObject("Scripting.FileSystemObject") %>

<% If Not objDeviceList.EOF Then

		Do Until objDeviceList.EOF

			'Active assignment
			If objDeviceList(3) Then
				If objDeviceList(6) Then
					strCardType = "NormalCard"
				Else
					strCardType = "DisabledCard"
				End If %>

			<div class="Card <%=strCardType%>">
				<div class="CardTitle">Active Assignment</div>
			<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(5)," ","") & ".png") Then %>
					<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
						<% If InStr(LCase(objDeviceList(5)),"ipad") Then %>
								<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="70" />
						<% Else %>
								<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="96" />
						<% End If %>
					</a>
			<% Else %>
					<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
						<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
					</a>
			<% End If %>


					<div>
						<div Class="PhotoCardColumn1">Tag </div>
						<div Class="PhotoCardColumn2">
							<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>"><%=objDeviceList(0)%></a>
						</div>
					</div>
					<div>
						<div Class="PhotoCardColumn1">Date: </div>
						<div Class="PhotoCardColumn2"><%=ShortenDate(objDeviceList(1))%></div>
					</div>
				</div>

		<% Else %>

			<% If objDeviceList(6) Then
					strCardType = "OldAssignmentCard"
				Else
					strCardType = "DisabledCard"
				End If %>
				<div class="Card <%=strCardType%>">
				<div class="CardTitle">Old Assignment</div>
			<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(5)," ","") & ".png") Then %>
					<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
						<% If InStr(LCase(objDeviceList(5)),"ipad") Then %>
								<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="70" />
						<% Else %>
								<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="96" />
						<% End If %>
					</a>
			<% Else %>
					<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
						<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
					</a>
			<% End If %>
					<div>
						<div Class="PhotoCardColumn1">Tag: </div>
						<div Class="PhotoCardColumn2">
							<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>"><%=objDeviceList(0)%></a>
						</div>
					</div>
					<div>
						<div Class="PhotoCardColumn1">Date: </div>
						<div Class="PhotoCardColumn2"><%=ShortenDate(objDeviceList(1)) & " - " & ShortenDate(objDeviceList(2))%></div>
					</div>
			<% If objDeviceList(4) <> "" Then %>
					<div>&nbsp;</div>
					<div>Notes: </div>
					<div><%=objDeviceList(4)%></div>
			<% End If %>
				</div>

		<% End If

			objDeviceList.MoveNext
		Loop

	End If %>

<%End Sub%>

<%Sub ActiveAssignments

	Dim objFSO, intLoopCounter, strSQL, objModel, objEventLookup, bolOpenEvent

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	intLoopCounter = 0 %>

<% If Not objDeviceList.EOF Then %>

	<% If intActiveAssignmentCount >= 1 Then %>

			<div class="Card NormalCard">

			<% If intActiveAssignmentCount = 1 Then %>
					<div class="CardTitle"><image src="../images/assignment.png" width="15" height="15" title="Assignments" /> Active Assignment</div>
			<% Else %>
					<div class="CardTitle"><image src="../images/assignment.png" width="15" height="15" title="Assignments" /> Active Assignments</div>
			<% End If %>

		<%	Do Until objDeviceList.EOF

				'Active assignment
				If objDeviceList(3) Then

					strDeviceInfo = ""
					If objDeviceList(16) <> "" Then
						strDeviceInfo = "Name: " & objDeviceList(16) & " &#013 "
					End If
					If objDeviceList(15) <> "" Then
						strDeviceInfo = strDeviceInfo & "Last User: " & objDeviceList(15) & " &#013 "
					End If
					If objDeviceList(17) <> "" Then
						strDeviceInfo = strDeviceInfo & "OS Version: " & objDeviceList(17) & " &#013 "
					End If
					If objDeviceList(18) <> "" Then
						strDeviceInfo = strDeviceInfo & "Last Checkin: " & objDeviceList(18) & " - " & objDeviceList(19)
					End If

					intLoopCounter = intLoopCounter + 1

					strSQL = "SELECT Model, Active FROM Devices WHERE LGTag='" & objDeviceList(0) & "'"
					Set objModel = Application("Connection").Execute(strSQL)
					
					strSQL = "SELECT ID FROM Events WHERE Resolved=False AND LGTag='" & objDeviceList(0) & "'"
					Set objEventLookup = Application("Connection").Execute(strSQL)
					
					If objEventLookup.EOF Then
						bolOpenEvent = False
					Else
						bolOpenEvent = True
					End If
					
					If objModel(1) = False Then %>
						<div Class="DisabledSectionInCard">
				<% ElseIf bolOpenEvent Then %>
						<div Class="WarningSectionInCard">
				<% Else %>
						<div>
				<% End If %>	
				
				<div Class="ImageSectionInAssignmentCard">
				<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(5)," ","") & ".png") Then %>
						<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
							<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="96" />
						</a>
				<% Else %>
						<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
							<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
						</a>
				<% End If %>


				<%	If objDeviceList(14) <> "" Then
						DrawIcon "Remote",0,""
					End If %>

				<% If strDeviceInfo <> "" Then
						If Application("MunkiReportServer") = "" Then %>
							<image src="../images/info.png" class="ButtonLeftAssignment" height="22" width="22" title="<%=strDeviceInfo%>">
					<% Else %>
							<a href="<%=Application("MunkiReportServer")%>/index.php?/clients/detail/<%=objDeviceList(12)%>" class="ButtonLeftAssignment" target="_blank">
								<image src="../images/info.png" width="20" height="20" title="<%=strDeviceInfo%>"  />
							</a>
					<% End If
					End If %>

				</div>
				<div Class="RightOfImageInAssignmentCard">
							<div>
								<div Class="PhotoCardColumn1">Tag:</div>
								<div Class="PhotoCardColumn2">
									<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>"><%=objDeviceList(0)%></a>
									<% If objDeviceList(10) Then %>
										&nbsp;<image src="../images/yes.png" width="15" height="15" title="Insured" />
									<% End If %>

								</div>
							</div>
						<% If Not objModel.EOF Then %>
								<div>
									<div Class="PhotoCardColumn1">Model: </div>
									<div Class="PhotoCardColumn2"><%=objModel(0)%></div>
								</div>
						<% End If %>
							<div>
								<div Class="CardMerged">Serial:

						<% Select Case objDeviceList(13)
								Case "Apple" %>
									<a href="https://checkcoverage.apple.com/us/en/?sn=<%=objDeviceList(12)%>" target="_blank"><%=objDeviceList(12)%></a>
							<% Case "Dell" %>
									<a href="http://www.dell.com/support/home/us/en/19/product-support/servicetag/<%=objDeviceList(12)%>" target="_blank"><%=objDeviceList(12)%></a>
							<% Case Else %>
									<%=objDeviceList(12)%>
						 <% End Select %>

								</div>
							</div>
							<div>
								<div Class="CardMerged">Assigned: <%=ShortenDate(objDeviceList(1))%></div>
							</div>
						<% If objDeviceList(11) <> "" Then %>
							<div>
								<div Class="CardMerged">Purchased: <%=ShortenDate(objDeviceList(11))%> - Year <%=GetAge(objDeviceList(11))%></div>
							</div>

						<% End If %>
						</div>
					</div>
					<%	If Int(intActiveAssignmentCount) <> Int(intLoopCounter) Then %>
							<hr />
					<% End If %>

			<% End If

				objDeviceList.MoveNext
			Loop
			objDeviceList.MoveFirst%>

			<hr />
			<div Class=Center>Assign Device</div>
			<br />
			<form method="POST" action="<%=strSubmitTo%>">
			<input type="hidden" name="StudentID" value="<%=intUserID%>" />
         <div Class="CardColumn1">Asset Tag: </div>
         <div Class="CardColumn2"><input Class="Card InputWidthSmall" type="text" name="Tag"></div>
      <% If strNewAssignmentMessage <> "" Then %>
      		<%=strNewAssignmentMessage%>
      <% End If %>
      	<div class="Button"><input type="submit" value="Assign Device" name="Submit" /></div>
      </form>

			</div>
		<% End If %>
<%	End If %>

<%End Sub%>

<%Sub AddAssignmentCard%>

	<div class="Card NormalCard">
		<div class="CardTitle">Assign Device</div>
		<form method="POST" action="<%=strSubmitTo%>">
			<input type="hidden" name="StudentID" value="<%=intUserID%>" />
         <div Class="CardColumn1">Asset Tag: </div>
         <div Class="CardColumn2"><input Class="Card InputWidthSmall" type="text" name="Tag"></div>
      <% If strNewAssignmentMessage <> "" Then %>
      		<%=strNewAssignmentMessage%>
      <% End If %>
      	<div class="Button"><input type="submit" value="Assign Device" name="Submit" /></div>
      </form>
	</div>

<%End Sub%>

<%Sub OldAssignements

	Dim objFSO, objPersonLookup, strSQL, strRowClass

	Set objFSO = CreateObject("Scripting.FileSystemObject")

   If Not objDeviceList.EOF Then

		If intOldAssignmentCount >=1 Then %>
			<div id="OldAssignments">
				<br />
				<image src="../images/assignment.png" height="15" width="15" title="Old Assignments"> Old Assignments
				<table align="center" Class="ListView" id="ListView">
					<thead>
						<th>Photo</th>
						<th>Tag</th>
						<th>Model</th>
						<th>Start Date</th>
						<th>End Date</th>
						<th>Assigned By</th>
						<th>Returned By</th>
						<th>Assignment Notes</th>
					</thead>
					<tbody>
			<% Do Until objDeviceList.EOF

					If Not objDeviceList(3) Then

						If objDeviceList(6) Then
							strRowClass = ""
						Else
							strRowClass = " Class=""Disabled"""
						End If %>

						<tr <%=strRowClass%>>
							<td <%=strRowClass%> width="1px">
						<% If objDeviceList(9) Then %>
							<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(5)," ","") & ".png") Then %>
										<% If InStr(LCase(objDeviceList(5)),"ipad") Then %>
												<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="70" />
										<% Else %>
												<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="96" />
										<% End If %>
							<% Else %>
										<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
							<% End If %>
						<% Else %>
							<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(5)," ","") & ".png") Then %>
									<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
										<% If InStr(LCase(objDeviceList(5)),"ipad") Then %>
												<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="70" />
										<% Else %>
												<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(5)," ","")%>.png" width="96" />
										<% End If %>
									</a>
							<% Else %>
									<a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>">
										<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
									</a>
							<% End If %>
						<% End If %>
							</td>

						<% If objDeviceList(9) Then %>
								<td <%=strRowClass%> id="center"><%=objDeviceList(0)%></td>
						<% Else %>
								<td <%=strRowClass%> id="center"><a href="device.asp?Tag=<%=objDeviceList(0)%><%=strBackLink%>"><%=objDeviceList(0)%></a></td>
						<% End If %>

							<td <%=strRowClass%>><%=objDeviceList(5)%></td>
							<td <%=strRowClass%> id="center"><%=ShortenDate(objDeviceList(1))%></td>

						<% If objUser(8) Then %>
							<% If objDeviceList(20) Then %>
									<td <%=strRowClass%> id="center"><%=ShortenDate(objDeviceList(2))%></td>
							<% Else %>
								<% If objDeviceList(6) Then %>
										<form method="POST" action="<%=strSubmitTo%>">
											<input type="hidden" name="StudentID" value="<%=intUserID%>" />
											<input type="hidden" name="Tag" value="<%=objDeviceList(0)%>" />
											<td <%=strRowClass%> id="center">
												<input type="hidden" value="Assign Device" name="Submit">
												<%=ShortenDate(objDeviceList(2))%> <input type="image" src="../images/assignment.png" width="15" height="15" title="Reassign Device" />
											</td>
										</form>
								<% Else %>
										<td <%=strRowClass%> id="center"><%=ShortenDate(objDeviceList(2))%></td>
								<% End If %>
							<% End If %>
						<% Else %>
								<td <%=strRowClass%> id="center"><%=ShortenDate(objDeviceList(2))%></td>
						<% End If %>

						<% If objDeviceList(7) = "" Then %>
								<td <%=strRowClass%>>&nbsp;</td>
						<% Else

								strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objDeviceList(7) & "'"
								Set objPersonLookup = Application("Connection").Execute(strSQL) %>

							<%	If Not objPersonLookup.EOF Then %>
									<td <%=strRowClass%>><%=objPersonLookup(1)%>, <%=objPersonLookup(0)%></td>
							<% Else %>
									<td <%=strRowClass%>><%=objDeviceList(7)%></td>
							<% End If%>

						<% End If %>

						<% If objDeviceList(8) = "" Then %>
								<td <%=strRowClass%>>&nbsp;</td>
						<% Else

								strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objDeviceList(8) & "'"
								Set objPersonLookup = Application("Connection").Execute(strSQL) %>

							<%	If Not objPersonLookup.EOF Then %>
									<td <%=strRowClass%>><%=objPersonLookup(1)%>, <%=objPersonLookup(0)%></td>
							<% Else %>
									<td <%=strRowClass%>><%=objDeviceList(8)%></td>
							<% End If%>

						<% End If %>

						<% If NOT IsNull(objDeviceList(4)) Then %>
								<td <%=strRowClass%>><%=Replace(objDeviceList(4),vbCRLF,"<br />")%></td>
						<% Else %>
								<td <%=strRowClass%>><%=objDeviceList(4)%></td>
						<% End If%>
						</tr>
				<%	End If
					objDeviceList.MoveNext
				Loop
				objDeviceList.MoveFirst %>
					</tbody>
				</table>
			</div>
	<% End If
	End If

End Sub%>

<%Sub EquipmentCard

	Dim strSQL, objItemLookup, objTotal, objParents %>

	<div class="Card NormalCard">

		<div class="CardTitle">Equipment</div>
		<% If objUSer(8) Then %>
			<form method="POST" action="<%=strSubmitTo%>">
			<input type="hidden" name="UserID" value="<%=intUserID%>" />
				<div Class="CardMerged">Loan:
					<select Class="Card" name="Item">
							<option value=""></option>
					<% Do Until objItems.EOF %>
								<option value="<%=objItems(0)%>"><%=objItems(0)%></option>
					<%    objItems.MoveNext
						Loop
						objItems.MoveFirst%>
					</select>
				<div class="Button"><input type="submit" value="Loan" name="Submit" /></div>
			</div>
			</form>
			<form method="POST" action="<%=strSubmitTo%>">
			<input type="hidden" name="UserID" value="<%=intUserID%>" />
			<div Class="CardMerged">
				<div>Bill:
					<select Class="Card" name="Item">
							<option value=""></option>
					<% Do Until objPurchasableItems.EOF %>
								<option value="<%=objPurchasableItems(0)%>"><%=objPurchasableItems(0)%> - $<%=objPurchasableItems(1)%></option>
					<%    objPurchasableItems.MoveNext
						Loop
						objPurchasableItems.MoveFirst%>
					</select>
				<div class="Button"><input type="submit" value="Bill" name="Submit" /></div>
			</div>
			</form>
			<% If Not objLoanedOut.EOF Or Not objOwes.EOF Then%>
					<br />
					<hr />
			<% End If %>
		<% End If %>
		<% If Not objLoanedOut.EOF Then %>

			<div Class="Center">Borrowed Items</div>
			<br />
			<% Do Until objLoanedOut.EOF %>
					<form method="POST" action="<%=strSubmitTo%>">
					<input type="hidden" name="LoanID" value="<%=objLoanedOut(0)%>" />
					<div>
						<a href="users.asp?LoanedOut=<%=Replace(objLoanedOut(1)," ","%20")%>"><%=objLoanedOut(1)%></a>&nbsp;&nbsp;&nbsp;
					<% strSQL = "SELECT ID FROM Purchasable WHERE Item='" & objLoanedOut(1) & "'"

						Set objItemLookup = Application("Connection").Execute(strSQL) %>

						<div class="Button"><input type="submit" value="Return" name="Submit" /></div>

					<% If Not objItemLookup.EOF Then	%>
							<div class="Button"><input type="submit" value="Bill" name="Submit" /></div>
					<% End If %>
					</div>
					</form>
				<% objLoanedOut.MoveNext
				Loop %>
			<% If Not objOwes.EOF Then %>
					<br />
					<hr />
			<% End If %>
		<% End If %>

		<% If Not objOwes.EOF Then

				'Get the total amount they owe
				strSQL = "SELECT Sum(Price) AS SumOfPrice FROM Owed GROUP BY OwedBy,Active Having Active=True AND OwedBy=" & objUser(0)
				Set objTotal = Application("Connection").Execute(strSQL)
				
				'Get their parents email addresses
				strSQL = "SELECT EMail, RelationShip, FirstName, LastName FROM Parents WHERE StudentID=" & objUser(4)
				Set objParents = Application("Connection").Execute(strSQL)
				
				%>

					<div Class="Center"><a href="" id="emailToggle"><image src="../images/email.png" height="20" width="20" title="Email Parents"></a> Owes $<%=objTotal(0)%></div><br/>
			<% Do Until objOwes.EOF %>
				   <form method="POST" action="<%=strSubmitTo%>">
						<input type="hidden" name="OwedID" value="<%=objOwes(0)%>" />
						<%=objOwes(1)%> - $<%=objOwes(2)%>
						<div class="Button"><input type="submit" value="Paid" name="Submit" /></div>
					<% If objOwes(4) Then %>
							<div class="Button"><input type="submit" value="Return" name="Submit" /></div>
					<% End If %>
					</form>
				<% objOwes.MoveNext
				Loop %>
			<div id="Notify">
				<br />
				<hr />
				
				<%
				
					If Not objParents.EOF Then
						Do Until objParents.EOF 
							If objParents(0) <> "" Then%>
								
								<form method="POST" action="<%=strSubmitTo%>">
									<div class="CardMerged">
										<%=objParents(1)%>: <input Class="Card InputWidthLarge" value="<%=objParents(0)%>" type="text" name="NotifyEmail">
										<div class="Button"><input type="submit" value="Notify" name="Submit" /></div>
									</div>
								</form>
	
						<%	End If
							objParents.MoveNext
						Loop
					End If
				%>
				<form method="POST" action="<%=strSubmitTo%>">
					<div class="CardMerged">
						EMail: <input Class="Card InputWidthLarge" type="text" name="NotifyEmail">
						<div class="Button"><input type="submit" value="Notify" name="Submit" /></div>
					</div>
				</form>
				<% If strNotifyMessage <> "" Then %>
						<%=strNotifyMessage%>
				<% End If %>
			</div>
		<%	End If %>
		</div>
	</div>

<%End Sub%>

<%Sub EventsTable

	Dim strWarrantyInfo, strCompleteInfo, strSQL, objName, objDeletedCheck

	If Not objEvents.EOF Then %>
		<div id="Events">
			<br />
			<image src="../images/event.png" height="15" width="15" Title="Events"> Events
			<table align="center" Class="ListView" id="EventTable">
				<thead>
					<th>Event</th>
					<th>Asset Tag</th>
					<th>Type</th>
					<th>Category</th>
					<th>Model</th>
					<th>Site</th>
					<th>Start Date</th>
					<th>End Date</th>
					<th>Warranty</th>
					<th>Entered By</th>
					<th>Completed By</th>
					<th>Complete</th>
					<th>Event Notes</th>
				</thead>

				<tbody>
		<% Do Until objEvents.EOF

		 		'Check and see if the devices has been deleted
		 		strSQL = "SELECT Deleted FROM Devices WHERE LGTag='" & objEvents(10) & "'"
		 		Set objDeletedCheck = Application("Connection").Execute(strSQL)

				If objEvents(9) Then
					strWarrantyInfo = "Yes"
				Else
					strWarrantyInfo = "No"
				End If

				If objEvents(5) Then
					strCompleteInfo = "Yes"
				Else
					strCompleteInfo = "No"
				End If%>

				<tr>
					<td id="center"><%=objEvents(0)%></th>

				<% If objDeletedCheck(0) Then %>
						<td id="center"><%=objEvents(10)%></td>
				<% Else %>
						<td id="center"><a href="device.asp?Tag=<%=objEvents(10)%><%=strBackLink%>"><%=objEvents(10)%></a></td>
				<% End If %>

					<td><a href="events.asp?EventType=<%=objEvents(1)%>&View=Table"><%=objEvents(1)%></a></td>
					<td><a href="events.asp?Category=<%=objEvents(8)%>&View=Table"><%=objEvents(8)%></a></td>
					<td><a href="events.asp?EventModel=<%=objEvents(13)%>&View=Table"><%=objEvents(13)%></a></td>
					<td><a href="events.asp?EventSite=<%=objEvents(12)%>&View=Table"><%=objEvents(12)%></a></td>
					<td><%=ShortenDate(objEvents(3))%></td>
					<td><%=ShortenDate(objEvents(6))%></td>
					<td id="center"><a href="events.asp?Warranty=<%=strWarrantyInfo%>&View=Table"><%=strWarrantyInfo%></a></td>

				<% If objEvents(14) <> "" Then

						strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objEvents(14) & "'"
						Set objName = Application("Connection").Execute(strSQL)

						If Not objName.EOF Then %>
							<td><%=objName(1)%>, <%=objName(0)%></td>
					<% Else %>
							<td></td>
					<%	End If %>

				<% Else %>
						<td></td>
				<% End If %>

				<% If objEvents(15) <> "" Then

						strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objEvents(15) & "'"
						Set objName = Application("Connection").Execute(strSQL)

						If Not objName.EOF Then %>
							<td><%=objName(1)%>, <%=objName(0)%></td>
					<% Else %>
							<td></td>
					<%	End If %>

				<% Else %>
						<td></td>
				<% End If %>

					<td id="center"><%=strCompleteInfo%></td>

				<% If NOT IsNull(objEvents(2)) Then %>
					<td><%=Replace(objEvents(2),vbCRLF,"<br />")%></td>
				<% Else %>
					<td><%=objEvents(2)%></td>
				<% End If %>
				</tr>
			<%	objEvents.MoveNext
			Loop
			objEvents.MoveFirst %>
				</tbody>
			</table>
		</div>
<% End If %>

<%End Sub%>

<%Sub ShowLog

	If Not objLog.EOF Then %>
		<div id="UserLog">
		<br />
		<image src="../images/log.png" height="15" width="15" title="Log"> Log
		<table align="center" Class="ListView" id="LogTable">
			<thead>
				<th>Date</th>
				<th>Time</th>
				<th>Type</th>
				<th>Asset Tag</th>
				<th>Event</th>
				<th>Performed By</th>
				<th>New Value</th>
				<th>Old Value</th>
			</thead>
			<tbody>
	<% Do Until objLog.EOF %>
				<tr>
					<td id="center"><%=ShortenDate(objLog(7))%></td>
					<td id="center"><%=ShortenTime(objLog(8))%></td>
					<td><%=LogEntryType(objLog(3))%></td>
					<td id="center"><a href="device.asp?Tag=<%=objLog(0)%>"><%=objLog(0)%></a></td>

				<%	If objLog(2) = 0 Then %>
						<td></td>
				<% Else %>
						<td id="center"><%=objLog(2)%></td>
				<% End If %>

					<td id="center"><%=LCase(objLog(6))%></td>

				<% If InStr(objLog(3),"Notes") = 0 And InStr(objLog(3),"DeviceReturned") = 0 Then %>
						<td><%=objLog(5)%></td>
						<td><%=objLog(4)%></td>
				<% Else %>
					<% If NOT IsNull(objLog(10)) Then %>
							<td><%=Replace(objLog(10),vbCRLF,"<br />")%></td>
					<% Else %>
							<td><%=objLog(10)%></td>
					<% End If%>

					<% If NOT IsNull(objLog(9)) Then %>
							<td><%=Replace(objLog(9),vbCRLF,"<br />")%></td>
					<% Else %>
							<td><%=objLog(9)%></td>
					<% End If%>

				<% End If %>

				</tr>
		<% objLog.MoveNext
		Loop %>
			</tbody>
		</table>
	</div>
<%	End If

End Sub%>

<%Function LogEntryType(EntryType)

	Select Case EntryType
		Case "AccountDisabledAUP"
			LogEntryType = "Account Disabled - AUP Not Turned In"
		Case "AccountDisabledSchoolTool"
			LogEntryType = "Account Disabled - Student Left District"
		Case "AccountEnabledAUP"
			LogEntryType = "Account Enabled - AUP Turned In"
		Case "AccountEnabledSchoolTool"
			LogEntryType = "Account Enabled - Student Returned To District"
		Case "AUPDisabled"
			LogEntryType = "AUP Not Turned In"
		Case "AUPEnabled"
			LogEntryType = "AUP Turned In"
		Case "AutoLogOut"
			LogEntryType = "Logout"
		Case "ComputerNameChange"
			LogEntryType = "Computer Name Changed"
		Case "ComputerImaged"
			LogEntryType = "Computer Imaged"
		Case "DatabaseUpgraded"
			LogEntryType = "Database Upgraded"
		Case "DeviceAssigned"
			LogEntryType = "Device Assigned"
		Case "DeviceAdded"
			LogEntryType = "New Device Added"
		Case "DeviceDecommissioned"
			LogEntryType = "Device Decommissioned"
		Case "DeviceDisabled"
			LogEntryType = "Device Decommissioned"
		Case "DeviceRestored"
			LogEntryType = "Device Restored to Active Status"
		Case "DeviceReturned"
			LogEntryType = "Device Returned"
		Case "DeviceReturnedAdapterMissing"
			LogEntryType = "Device Returned without Adapter"
		Case "DeviceReturnedCaseMissing"
			LogEntryType = "Device Returned without Case"
		Case "DeviceReturnedDamaged"
			LogEntryType = "Device Returned Damaged"
		Case "DeviceUpdatedAssetTag"
			LogEntryType = "Device Asset Tag Updated"
		Case "DeviceUpdatedBOCESTag"
			LogEntryType = "Device BOCES Tag Updated"
		Case "DeviceUpdatedInsurance"
			LogEntryType = "Insurance Updated"
		Case "DeviceUpdatedMACAddress"
			LogEntryType = "MAC Address Updated"
		Case "DeviceUpdatedMake"
			LogEntryType = "Device Make Updated"
		Case "DeviceUpdatedModel"
			LogEntryType = "Device Model Updated"
		Case "DeviceUpdatedNotes"
			LogEntryType = "Device Notes Updated"
		Case "DeviceUpdatedPurchased"
			LogEntryType = "Device Purchase Date Updated"
		Case "DeviceUpdatedRoom"
			LogEntryType = "Device Room Updated"
		Case "DeviceUpdatedSerial"
			LogEntryType = "Device Serial Number Updated"
		Case "DeviceUpdatedSite"
			LogEntryType = "Device Site Updated"
		Case "DeviceUpdatedTagAdded"
			LogEntryType = "Device Tag Added"
		Case "DeviceUpdatedTagDeleted"
			LogEntryType = "Device Tag Deleted"
		Case "EmailedTeachersAssignment"
			LogEntryType = "EMail Sent about Assignment"
		Case "EventAdded"
			LogEntryType = "Event Added"
		Case "EventClosed"
			LogEntryType = "Event Closed"
		Case "EventUpdatedCategory"
			LogEntryType = "Event Category Updated"
		Case "EventUpdatedNotes"
			LogEntryType = "Event Notes Updated"
		Case "EventUpdatedType"
			LogEntryType = "Event Type Updated"
		Case "EventUpdatedWarranty"
			LogEntryType = "Event Warranty Updated"
		Case "ExternalIPChange"
			LogEntryType = "External IP Changed"
		Case "InternalIPChange"
			LogEntryType = "Internal IP Changed"
		Case "InternetAccessChanged"
			LogEntryType = "Internet Access Changed"
		Case "ItemPaidFor"
			LogEntryType = "Equipment Paid Off"
		Case "ItemReturned"
			LogEntryType = "Billed Equipment Returned"
		Case "LastUserChange"
			LogEntryType = "Different User Signed In"
		Case "LoanedItemBilled"
			LogEntryType = "Billed for Loaned Equipment"
		Case "LoanedOutItem"
			LogEntryType = "Equipment Loaned Out"
		Case "LoanedOutItemReturned"
			LogEntryType = "Equipment Returned"
		Case "MoneyOwed"
			LogEntryType = "Billed for Equipment"
		Case "NewADAccountCreated"
			LogEntryType = "Account Created in Active Directory"
		Case "NewStudentDetected"
			LogEntryType = "New Student Detected"
		Case "NewStudentPasswordSet"
			LogEntryType = "Password Entered for New Student"
		Case "NewStudentReady"
			LogEntryType = "New Student Activated"
		Case "NotificationEMailSent"
			LogEntryType = "Notification EMail Sent"
		Case "OSVersionChange"
			LogEntryType = "Operating System Updated"
		Case "PasswordExpiresSet"
			LogEntryType = "Password Expires"
		Case "PasswordNeverExpiresSet"
			LogEntryType = "Password Doesn't Expire"
		Case "ReturnedMissingAdapter"
			LogEntryType = "Returned Missing Adapter"
		Case "ReturnedMissingCase"
			LogEntryType = "Returned Missing Case"
		Case "SearchSaved"
			LogEntryType = "New Search Created"
		Case "StudentGradeChange"
			LogEntryType = "Student Changed Grade"
		Case "UserAdded"
			LogEntryType = "New User Added"
		Case "UserDisabled"
			LogEntryType = "User Disabled"
		Case "UserUpdatedAUP"
			LogEntryType = "User AUP Status Updated"
		Case "UserLogin"
			LogEntryType = "Login"
		Case "UserLoginAdmin"
			LogEntryType = "Admin Login"
		Case "UserLogout"
			LogEntryType = "Logout"
		Case "UserRestored"
			LogEntryType = "User Restored"
		Case "UserUpdatedDescription"
			LogEntryType = "User's Description Updated"
		Case "UserUpdatedFirstName"
			LogEntryType = "User's First Name Updated"
		Case "UserUpdatedLastName"
			LogEntryType = "User's Last Name Updated"
		Case "UserUpdatedNotes"
			LogEntryType = "User's Notes Updated"
		Case "UserUpdatedPassword"
			LogEntryType = "User's Password Updated"
		Case "UserUpdatedPhone"
			LogEntryType = "User's Phone Update"
		Case "UserUpdatedPhotoID"
			LogEntryType = "User's Photo ID Update"
		Case "UserUpdatedRole"
			LogEntryType = "User's Role Update"
		Case "UserUpdatedRoom"
			LogEntryType = "User's Room Update"
		Case "UserUpdatedSite"
			LogEntryType = "User's Site Update"
		Case "UserUpdatedUserName"
			LogEntryType = "User's Username Updated"
		Case Else
			LogEntryType = EntryType
	End Select

End Function%>

<%Sub JumpToDevice%>

	<div Class="HeaderCard">
		<form method="POST" action="search.asp">
		Asset tag: <input class="Card InputWidthSmall" type="text" name="SmartBox" id="LastNames" />
		<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		<input type="hidden" value="" name="BOCESTag" />
		</form>
	</div>

<%End Sub%>

<%Sub LoanOut %>
	<% If objLoanedOut.EOF Then
			strCardType ="NormalCard"
		Else
			strCardType = "LoanedCard"
		End If %>

		<div class="Card <%=strCardType%>">

			<div class="CardTitle">Loan Equipment</div>
			<form method="POST" action="<%=strSubmitTo%>">
			<input type="hidden" name="UserID" value="<%=intUserID%>" />
			<div>
				<div Class="CardColumn1">Item: </div>
				<div Class="CardColumn2">
					<select Class="Card" name="Item">
							<option value=""></option>
					<% Do Until objItems.EOF %>
								<option value="<%=objItems(0)%>"><%=objItems(0)%></option>
					<%    objItems.MoveNext
						Loop
						objItems.MoveFirst%>
					</select>
				</div>
			</div>
			<div>
				<div class="Button"><input type="submit" value="Loan Out Item" name="Submit" /></div>
			</div>
			</form>
		</div>

<%End Sub%>

<%Sub LoanedOut %>
		<% If Not objLoanedOut.EOF Then %>
				<div class="Card LoanedCard">
					<div class="CardTitle">Borrowing</div>
				<% Do Until objLoanedOut.EOF %>
						<form method="POST" action="<%=strSubmitTo%>">
						<input type="hidden" name="LoanID" value="<%=objLoanedOut(0)%>" />
						<div>
							<a href="users.asp?LoanedOut=<%=Replace(objLoanedOut(1)," ","%20")%>"><%=objLoanedOut(1)%></a>&nbsp;&nbsp;&nbsp;
							<div class="Button"><input type="submit" value="Return" name="Submit" /></div>
						</div>
						</form>
					<% objLoanedOut.MoveNext
					Loop %>
				</div>
		<% End If %>
<%End Sub%>

<%Sub MissingStuff %>
<% If Not objMissingStuff.EOF Then %>

		<div class="Card WarningCard">
			<div class="CardTitle">Missing Equipment</div>

	<% Do Until objMissingStuff.EOF %>

		<% If Not objMissingStuff(2) Then %>
				<form method="POST" action="<%=strSubmitTo%>">
				<input type="hidden" name="AssignmentID" value="<%=objMissingStuff(0)%>" />
				<input type="hidden" name="Adapter" value="True" />
				<div>
					Tag: <a href="device.asp?Tag=<%=objMissingStuff(1)%><%=strBackLink%>"><%=objMissingStuff(1)%></a> - Power Adapter
					<div class="Button"><input type="submit" value="Return" name="Submit" /></div>
				</div>
				</form>
		<% End If %>

		<% If Not objMissingStuff(3) Then %>
				<form method="POST" action="<%=strSubmitTo%>">
				<input type="hidden" name="AssignmentID" value="<%=objMissingStuff(0)%>" />
				<input type="hidden" name="Case" value="True" />
				<div>
					Tag: <a href="device.asp?Tag=<%=objMissingStuff(1)%><%=strBackLink%>"><%=objMissingStuff(1)%></a> - Laptop Case
					<div class="Button"><input type="submit" value="Return" name="Submit" /></div>
				</div>
				</form>
		<% End If %>

		<% objMissingStuff.MoveNext
		Loop %>
		</div>
<% End If %>
		</div>
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
		Case 4
			strPosition = "Four"
	End Select

	Select Case strIconType
		Case "HelpDesk" %>
			<a href="<%=Application("HelpDeskURL")%>/index.asp?UserName=<%=strURLData%>" class="Button<%=strPosition%>">
				<image src="../images/helpdesk.png" width="20" height="20" title="Enter Help Desk Ticket" />
			</a>
	<%	Case "Back" %>
			<a href="<%=Request.QueryString("Page")%>?<%=Request.QueryString("Back")%>" class="Button<%=strPosition%>">
				<image src="../images/back.png" width="20" height="20" title="Return to Search Results"/>
			</a>
	<% Case "Log" %>
			<a href="" class="Button<%=strPosition%>" id="logToggle">
				<image src="../images/log.png" height="20" width="20" title="View Log">
			</a>

	<% Case "Assignments" %>
			<a href="" class="Button<%=strPosition%>" id="assignmentsToggle">
				<image src="../images/assignment.png" height="20" width="20" title="View Old Assignments">
			</a>

	<% Case "Events" %>
			<a href="" class="Button<%=strPosition%>" id="eventsToggle">
				<image src="../images/event.png" height="20" width="20" title="View Events">
			</a>

	<% Case "ViewAll" %>
			<a href="" class="Button<%=strPosition%>" id="viewAllToggle">
				<image src="../images/viewall.png" height="20" width="20" title="View All">
			</a>
			
	<% Case "EMail" %>
			<a href="" class="Button<%=strPosition%>" id="emailToggle">
				<image src="../images/email.png" height="20" width="20" title="Email Parents">
			</a>

	<% Case "Remote" %>
			<a href="vnc://<%=objDeviceList(14)%>:5900" class="Button<%=strPosition%>" >
	<%
	Set WshShell = CreateObject("WScript.Shell")
   	PINGFlag = Not CBool(WshShell.run("ping -n 1 -w 1000 " & objDeviceList(14),0,True))
   	If PINGFlag = True Then
'    		deviceOn = "greendot.png"
   		status = "Remote Control Online"
   	%>	<image src="../images/remote.png" height="20" width="20" title="<%=status%>" class="ButtonLeftAssignment">
   <%Else
'    		deviceOn = "reddot.png"
   		status = "Remote Control Offline"
   %>
   		<image style="opacity:0.5;filter:alpha(opacity=50)" src="../images/remote.png" height="20" width="20" title="<%=status%>" class="ButtonLeftAssignment">
   <% End If %>
			</a>
			<a href="ssh://admin@<%=objDeviceList(14)%>" class="Button<%=strPosition%>" >
		<% If PINGFlag Then %>
			<image src="../images/ssh.png" height="20" width="20" title="SSH Online" class="ButtonLeftAssignment">
		<% Else %>
			<image style="opacity:0.5;filter:alpha(opacity=50)" src="../images/ssh.png" height="20" width="20" title="SSH Offline" class="ButtonLeftAssignment">

		<% End If %>
			</a>

	<% Case "Info"
			If Application("MunkiReportServer") = "" Then %>
				<image src="../images/info.png" class="ButtonLeftAssignment" height="22" width="22" title="<%=strDeviceInfo%>">
		<% Else %>
				<a href="<%=Application("MunkiReportServer")%>/index.php?/clients/detail/<%=objDevice(4)%>" target="_blank">
					<div class="ButtonLeftAssignment"><image src="../images/info.png" width="20" height="20" title="<%=strDeviceInfo%>"  />&nbsp;</div>
				</a>
		<% End If %>


<%	End Select

End Sub %>

<%Sub AssignDevice

   Dim intStudent, bolInsurance, strSQL, objDeviceCheck, objAssignmentCheck, objAssignedTo
   Dim objDeviceCount, intDeviceCount, intTag, objOpenEvents, objActiveCheck

   'Grade the data from the form
   intStudent = Request.Form("StudentID")
   bolInsurance = False

   'Get the variables from the URL or form and fix it
   intTag = Request.Form("Tag")
   If intTag = "" then
      intTag = Request.QueryString("Tag")
   End If
   If Not Application("UseLeadingZeros") Then
		If IsNumeric(intTag) Then
			intTag = Int(intTag)
		Else
			If Left(intTag,4) = "TECH" Then
		   	intTag = Replace(intTag,"TECH","")
				If IsNumeric(intTag) Then
		      	intTag = Int(intTag)
			  	End If
			End If
		End If
	End If

   'Make sure they submitted something
   If intStudent = "" Or intTag = "" Then
      strNewAssignmentMessage = "<div Class=""Error"">Missing Data</div>"
   Else

      'Check and see if the tag is in the database
      strSQL = "SELECT ID FROM Devices WHERE LGTag='" & intTag & "'"
      Set objDeviceCheck = Application("Connection").Execute(strSQL)

      If Not objDeviceCheck.EOF Then
      
      	'Check and see if device is active in the database
			strSQL = "SELECT ID FROM Devices WHERE LGTag='" & intTag & "' AND Active=True"
			Set objActiveCheck = Application("Connection").Execute(strSQL)

				
			If Not objActiveCheck.EOF Then
				
				'Check and see if the device is already assigned
				strSQL = "SELECT AssignedTo FROM Assignments WHERE LGTag='" & intTag & "' And Active=True"
				Set objAssignmentCheck = Application("Connection").Execute(strSQL)

				If objAssignmentCheck.EOF Then

					'Check and see if the device has an open event
					strSQL = "SELECT ID FROM Events WHERE LGTag='" & intTag & "' AND Resolved=False"
					Set objOpenEvents = Application("Connection").Execute(strSQL)

					If objOpenEvents.EOF Then

						'Make sure the insurance variable is ready
						If Not bolInsurance Then
							bolInsurance = False
						End If

						'Create the assignment in the database
						strSQL = "INSERT INTO Assignments (LGTag, DateIssued, TimeIssued, Active, AssignedTo, HasInsurance, IssuedBy)" & vbCRLF
						strSQL = strSQL & "VALUES ('"
						strSQL = strSQL & intTag & "',#"
						strSQL = strSQL & Date & "#,#"
						strSQL = strSQL & Time & "#,"
						strSQL = strSQL & True & ","
						strSQL = strSQL & intStudent & ","
						strSQL = strSQL & bolInsurance & ",'"
						strSQL = strSQL & strUser & "')"
						Application("Connection").Execute(strSQL)

						'Get the current number of devices the user has
						strSQL = "SELECT DeviceCount,FirstName,LastName,UserName FROM People WHERE ID=" & intStudent
						Set objDeviceCount = Application("Connection").Execute(strSQL)
						If IsNull(objDeviceCount(0)) Then
							intDeviceCount = 1
						Else
							intDeviceCount = objDeviceCount(0) + 1
						End If

						'Update the student to show they have a device
						strSQL = "UPDATE People" & vbCRLF
						strSQL = strSQL & "SET HasDevice = True, DeviceCount=" & intDeviceCount & vbCRLF
						strSQL = strSQL & "WHERE ID = " & intStudent
						Application("Connection").Execute(strSQL)

						'Update the device to show it's assigned
						strSQL = "UPDATE Devices SET "
						strSQL = strSQL & "Assigned=True,"
						strSQL = strSQL & "FirstName='" & Replace(objDeviceCount(1),"'","''") & "',"
						strSQL = strSQL & "LastName='" &  Replace(objDeviceCount(2),"'","''") & "',"
						strSQL = strSQL & "UserName='" &  Replace(objDeviceCount(3),"'","''") & "'" & vbCRLF
						strSQL = strSQL & "WHERE LGTag='" & intTag & "'"
						Application("Connection").Execute(strSQL)

						UpdateLog "DeviceAssigned",intTag,objDeviceCount(3),"",GetDisplayName(objDeviceCount(3)),""

						'Set the message to return
						strNewAssignmentMessage = "<div class=""Information"">Assigned</div>"

					Else
						strNewAssignmentMessage = "<div Class=""Error""><a href=""device.asp?Tag=" & intTag & """>" & intTag & "</a> Has an Open Event, Cannot Assign</div>"
					End If

				Else

					'Find out who the device is already assigned to
					'strSQL = "SELECT FirstName, LastName" & vbCRLF
					'strSQL = strSQL & "FROM People" &vbCRLF
					'strSQL = strSQL & "WHERE ID=" & objAssignmentCheck(0)
					'Set objAssignedTo = Application("Connection").Execute(strSQL)

					'strNewAssignmentMessage = "<div Class=""Error"">Device already assigned to " & objAssignedTo(0) & " " & objAssignedTo(1) & "</div>"

					strNewAssignmentMessage = "<div Class=""Error"">Device already assigned</div>"

				End If
			Else
				strNewAssignmentMessage = "<div Class=""Error"">Device not active</div>"
			End If
      Else
         strNewAssignmentMessage = "<div Class=""Error"">Device not found</div>"
      End If

   End If

End Sub%>

<%Sub UpdateUser

	Dim strSQL, strNotes, strOldNotes, objUserLookup, bolAUP, bolOldAUP, strFirstName, strLastName, intRole
	Dim strSite, intPhotoID, strRoom, strPhone, strDescription, strOldFirstName, strOldLastName, intOldRole
	Dim strOldSite, intOldPhotoID, strOldRoom, strOldPhone, strOldDescription, intIndex, strNewUserName
	Dim objUserCheck, bolUserNameChanged, strOldUserName, strPassword, strOldPassword, strInternetAccess
	Dim strOldInternetAccess

	'Get the values from the form
	strNotes = Request.Form("Notes")
	bolAUP = Request.Form("AUP")
	strFirstName = Request.Form("FirstName")
	strLastName = Request.Form("LastName")
	intRole = Request.Form("Role")
	strSite = Request.Form("Site")
	intPhotoID = Request.Form("PhotoID")
	strRoom = Request.Form("Room")
	strPhone = Request.Form("Phone")
	strDescription = Request.Form("Description")
	strNewUserName = Request.Form("NewUserName")
	strPassword = Request.Form("UserPassword")
	strInternetAccess = Request.Form("InternetAccess")
	bolUserNameChanged = False

	'Fix the AUP variable
	If bolAUP = "True" Then
		bolAUP = True
	Else
		bolAUP = False
	End If

	'Don't change the AUP status if the user is an adult
	If intRole < 1000 Then
		bolAUP = True
	End If

	'Make sure they didn't enter anything that's too long
	If Len(strDescription) > 250 Then
		strUserMessage = "<div Class=""Error"">Description Longer Than 250 Characters</div>"
		Exit Sub
	End If
	If Len(strPhone) > 25 Then
		strUserMessage = "<div Class=""Error"">Phone Number Longer Than 25 Characters</div>"
		Exit Sub
	End If
	If Len(strPassword) > 120 Then
		strUserMessage = "<div Class=""Error"">Password Longer Than 120 Characters</div>"
		Exit Sub
	End If
	If Len(strFirstName) > 50 Then
		strUserMessage = "<div Class=""Error"">First Name Longer Than 50 Characters</div>"
		Exit Sub
	End If
	If Len(strLastName) > 50 Then
		strUserMessage = "<div Class=""Error"">Last Name Longer Than 50 Characters</div>"
		Exit Sub
	End If
	If Len(strNewUserName) > 50 Then
		strUserMessage = "<div Class=""Error"">Username Longer Than 50 Characters</div>"
		Exit Sub
	End If
	If Len(strRoom) > 50 Then
		strUserMessage = "<div Class=""Error"">Room Longer Than 50 Characters</div>"
		Exit Sub
	End If

	'Make sure the purchased date is in the right format
	If intPhotoID <> "" Then
		If Not IsNumeric(intPhotoID) Then
			strUserMessage = "<div Class=""Error"">Invalid Photo ID</div>"
			Exit Sub
		End If
	End If

	'Update the username if the first or last name changed.
	If LCase(Trim(strUserName)) <> LCase(Trim(strNewUserName)) Then

		'Make sure the username isn't already in use.
		strSQL = "SELECT ID FROM People WHERE Username='" & Replace(strNewUserName,"'","''") & "'"
		Set objUserCheck = Application("Connection").Execute(strSQL)

		If objUserCheck.EOF Then
			bolUserNameChanged = True
		Else
			strUserMessage = "<div Class=""Error"">Username Already In Use</div>"
			Exit Sub
		End If

		'Make sure they didn't submit a blank username
		If strNewUserName = "" Then
			strUserMessage = "<div Class=""Error"">Username Can't Be Blank</div>"
			Exit Sub
		End If

	End If

	'Get the current values from the database
	strSQL = "SELECT Notes,AUP,FirstName,LastName,ClassOf,Site,StudentID,RoomNumber,PhoneNumber,Description,PWord,InternetAccess" &vbCRLF
	strSQL = strSQL & "FROM People Where UserName='" & Replace(strUserName,"'","''") & "'"
	Set objUserLookup = Application("Connection").Execute(strSQL)

	'The user should be found, but just in case we'll make sure
	If Not objUserLookup.EOF Then

		'Get the old values
		strOldNotes = objUserLookup(0)
		bolOldAUP = objUserLookup(1)
		strOldFirstName = objUserLookup(2)
		strOldLastName = objUserLookup(3)
		intOldRole = objUserLookup(4)
		strOldSite = objUserLookup(5)
		intOldPhotoID = objUserLookup(6)
		strOldRoom = objUserLookup(7)
		strOldPhone = objUserLookup(8)
		strOldDescription = objUserLookup(9)
		strOldPassword = objUserLookup(10)
		strOldInternetAccess = objUserLookup(11)
		strOldUserName = strUserName

		'Fix the old values if they were null
		If IsNull(strOldNotes) THen
			strOldNotes = ""
		End If
		If IsNull(strOldFirstName) THen
			strOldFirstName = ""
		End If
		If IsNull(strOldLastName) THen
			strOldLastName = ""
		End If
		If IsNull(intOldRole) THen
			intOldRole = 0
		End If
		If IsNull(strOldSite) THen
			strOldSite = ""
		End If
		If IsNull(intOldPhotoID) THen
			intOldPhotoID = 0
		End If
		If IsNull(strOldRoom) THen
			strOldRoom = ""
		End If
		If IsNull(strOldPhone) THen
			strOldPhone = ""
		End If
		If IsNull(strOldDescription) THen
			strOldDescription = ""
		End If
		If IsNull(strOldPassword) Then
			strOldPassword = ""
		End If
		If IsNull(strOldInternetAccess) Then
			strOldInternetAccess = ""
		End If

	End If
	
	'Fix the password if it's blank
	If strPassword = "" Then
		strPassword = strOldPassword
	End If

	'Don't change the AUP if we're not showing passwords
	If Not Application("ShowPasswords") Then
		bolAUP = bolOldAUP
	End If

	'Make sure a user with the same name doesn't already exist in the new site or role
	If strOldSite <> strSite Or CInt(intOldRole) <> CInt(intRole) Then
		strSQL = "SELECT ID FROM People Where FirstName='" & strFirstName & "' AND LastName='" & Replace(strLastName,"'","''") & "' AND "
		strSQL = strSQL & "ClassOf=" & intRole & " AND Site='" & strSite & "'"
		Set objUserCheck = Application("Connection").Execute(strSQL)

		If Not objUserCheck.EOF Then
			strUserMessage = "<div Class=""Error"">Name Conflict</div>"
			Exit Sub
		End If
	End If

	'Save the changes
	strSQL = "UPDATE People SET Notes='" & Replace(strNotes,"'","''") & "',"
	strSQL = strSQL & "FirstName='" & Replace(strFirstName,"'","''") & "',"
	strSQL = strSQL & "LastName='" & Replace(strLastName,"'","''") & "',"
	strSQL = strSQL & "ClassOf=" & intRole & ","
	strSQL = strSQL & "Site='" & Replace(strSite,"'","''") & "',"
	strSQL = strSQL & "StudentID=" & intPhotoID & ","
	strSQL = strSQL & "RoomNumber='" & Replace(strRoom,"'","''") & "',"
	strSQL = strSQL & "PhoneNumber='" & Replace(strPhone,"'","''") & "',"
	strSQL = strSQL & "Description='" & Replace(strDescription,"'","''") & "',"
	strSQL = strSQL & "UserName='" & Replace(strNewUserName,"'","''") & "',"
	strSQL = strSQL & "PWord='" & Replace(strPassword,"'","''") & "',"
	strSQL = strSQL & "InternetAccess='" & Replace(strInternetAccess,"'","''") & "',"
	strSQL = strSQL & "AUP=" & bolAUP & " "
	strSQL = strSQL & "WHERE UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)

	'Update the log
	If strOldNotes <> strNotes Then
		UpdateLog "UserUpdatedNotes","",strUserName,strOldNotes,strNotes,""
	End If
	If bolOldAUP <> bolAUP Then
		If bolAUP Then
			UpdateLog "UserUpdatedAUP","",strUserName,"No","Yes",""
		Else
			UpdateLog "UserUpdatedAUP","",strUserName,"Yes","No",""
		End If
	End If
	If strOldInternetAccess <> strInternetAccess Then
		UpdateLog "InternetAccessChanged","",strUserName,strOldInternetAccess,strInternetAccess,""
	End If
	
	
	If strOldPassword <> strPassword Then
		UpdateLog "UserUpdatedPassword","",strUserName,strOldPassword,strPassword,""

		'Disable any pending tasks related to changing the password
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='UpdatePassword' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "UpdatePassword','" & Replace(strUserName,"'","''") & "','" & Replace(strPassword,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If
	If strOldFirstName <> strFirstName Then
		UpdateLog "UserUpdatedFirstName","",strUserName,strOldFirstName,strFirstName,""
		strSQL = "UPDATE Devices SET FirstName='" & strFirstName & "' WHERE FirstName='" & strOldFirstName & "' AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Disable any pending tasks related to changing the first name
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='UpdateFirstName' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "UpdateFirstName','" & Replace(strUserName,"'","''") & "','" & Replace(strFirstName,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If
	If strOldLastName <> strLastName Then
		UpdateLog "UserUpdatedLastName","",strUserName,strOldLastName,strLastName,""
		strSQL = "UPDATE Devices SET LastName='" & Replace(strLastName,"'","''") & "' WHERE LastName='" & Replace(strOldLastName,"'","''") & "' AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Disable any pending tasks related to changing the last name
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='UpdateLastName' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "UpdateLastName','" & Replace(strUserName,"'","''") & "','" & Replace(strLastName,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If
	If CInt(intOldRole) <> CInt(intRole) Then 'This change means we have to move the user in AD as well.
		UpdateLog "UserUpdatedRole","",strUserName,GetRole(intOldRole),GetRole(intRole),""

		'Disable any pending tasks related to moving the user
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='MoveUser' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "MoveUser','" & Replace(strUserName,"'","''") & "','','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If
	If strOldSite <> strSite Then
		UpdateLog "UserUpdatedSite","",strUserName,strOldSite,strSite,""

		'Disable any pending tasks related to moving the user
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='MoveUser' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "MoveUser','" & Replace(strUserName,"'","''") & "','','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If
	If CLng(intOldPhotoID) <> CLng(intPhotoID) Then
		UpdateLog "UserUpdatedPhotoID","",strUserName,intOldPhotoID,intPhotoID,""
	End If
	If strOldRoom <> strRoom Then
		UpdateLog "UserUpdatedRoom","",strUserName,strOldRoom,strRoom,""

		'Disable any pending tasks related to changing the room
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='UpdateRoom' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "UpdateRoom','" & Replace(strUserName,"'","''") & "','" & Replace(strRoom,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If
	If strOldPhone <> strPhone Then
		UpdateLog "UserUpdatedPhone","",strUserName,strOldPhone,strPhone,""

		'Disable any pending tasks related to changing the phone number
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='UpdatePhone' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "UpdatePhone','" & Replace(strUserName,"'","''") & "','" & Replace(strPhone,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If
	If strOldDescription <> strDescription Then
		UpdateLog "UserUpdatedDescription","",strUserName,strOldDescription,strDescription,""

		'Disable any pending tasks related to changing the description
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='UpdateDescription' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "UpdateDescription','" & Replace(strUserName,"'","''") & "','" & Replace(strDescription,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If
	If bolUserNameChanged Then
		UpdateLog "UserUpdatedUserName","",strUserName,strUserName,strNewUserName,""
		strSQL = "UPDATE People SET UserName='" & Replace(strNewUserName,"'","''") & "' WHERE UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Log SET UserName='" & Replace(strNewUserName,"'","''") & "' WHERE UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Log SET UpdatedBy='" & Replace(strNewUserName,"'","''") & "' WHERE UpdatedBy='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Devices SET UserName='" & Replace(strNewUserName,"'","''") & "' WHERE UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Devices SET LastUser='" & Replace(strNewUserName,"'","''") & "' WHERE LastUser='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Assignments SET ReturnedBy='" & Replace(strNewUserName,"'","''") & "' WHERE ReturnedBy='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Assignments SET IssuedBy='" & Replace(strNewUserName,"'","''") & "' WHERE IssuedBy='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Events SET EnteredBy='" & Replace(strNewUserName,"'","''") & "' WHERE EnteredBy='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Events SET CompletedBy='" & Replace(strNewUserName,"'","''") & "' WHERE CompletedBy='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Sessions SET Username='" & Replace(strNewUserName,"'","''") & "' WHERE Username='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE PendingTasks SET Username='" & Replace(strNewUserName,"'","''") & "' WHERE Username='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
		strSQL = strSQL & "UpdateUserName','" & Replace(strUserName,"'","''") & "','" & Replace(strNewUserName,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
		Application("Connection").Execute(strSQL)

	End If

	'Add the AUP change to the pending tasks table if needed
	If bolOldAUP <> bolAUP Then

		'Disable any pending tasks related to changing the AUP
		strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='AUPEnable' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Add the tasks to the pending tasks table
		If bolAUP Then
			strSQL = "INSERT INTO PendingTasks (Task,UserName,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
			strSQL = strSQL & "AUPEnable','" & Replace(strUserName,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
			Application("Connection").Execute(strSQL)
		End If

	End If

	strUserMessage = "<div Class=""Information"">Updated</div>"

	If bolUserNameChanged Then
		Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?UserName=" & strNewUserName
	End If

End Sub%>

<%Sub RestoreUser

	Dim strSQL

	'Enable the User
	strSQL = "UPDATE People SET Active=True WHERE UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)

	'Disable any pending tasks related to enabling the account
	strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='EnableUser' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)

	'Add the tasks to the pending tasks table
	strSQL = "INSERT INTO PendingTasks (Task,UserName,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
	strSQL = strSQL & "EnableUser','" & Replace(strUserName,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
	Application("Connection").Execute(strSQL)

	'Update the log
	UpdateLog "UserRestored","",strUserName,"Disabled","Enabled",""

End Sub%>

<%Sub DisableUser

	Dim strSQL

	'Disable the User
	strSQL = "UPDATE People SET Active=False WHERE UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)

	'Disable any pending tasks related to disabling the account
	strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='DisableUser' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)

	'Add the tasks to the pending tasks table
	strSQL = "INSERT INTO PendingTasks (Task,UserName,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
	strSQL = strSQL & "DisableUser','" & Replace(strUserName,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
	Application("Connection").Execute(strSQL)

	'Update the log
	UpdateLog "UserDisabled","",strUserName,"Enabled","Disabled",""

End Sub%>

<%Sub Return

	If Request.Form("AssignmentID") <> "" Then
		ReturnMissingItem
	Else
		ReturnLoanedItem
	End If

End Sub%>

<%Sub ReturnMissingItem

	Dim intAssignmentID, bolAdapterReturned, bolCaseReturned, strSQL, objItemCheck, objDeviceLookup

	intAssignmentID = Request.Form("AssignmentID")
	bolAdapterReturned = Request.Form("Adapter")
	bolCaseReturned = Request.Form("Case")

	If bolAdapterReturned Then
		strSQL = "UPDATE Assignments SET AdapterReturned=True WHERE ID=" & intAssignmentID
		Application("Connection").Execute(strSQL)

		strSQL = "SELECT LGTag FROM Assignments WHERE ID=" & intAssignmentID
		Set objDeviceLookup = Application("Connection").Execute(strSQL)

		UpdateLog "ReturnedMissingAdapter",objDeviceLookup(0),strUserName,"","",""
	End If

	If bolCaseReturned Then
		strSQL = "UPDATE Assignments SET CaseReturned=True WHERE ID=" & intAssignmentID
		Application("Connection").Execute(strSQL)

		strSQL = "SELECT LGTag FROM Assignments WHERE ID=" & intAssignmentID
		Set objDeviceLookup = Application("Connection").Execute(strSQL)

		UpdateLog "ReturnedMissingCase",objDeviceLookup(0),strUserName,"","",""
	End If

	strSQL = "SELECT Assignments.ID FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
	strSQL = strSQL & "WHERE (UserName='" & Replace(strUserName,"'","''") & "' AND Assignments.Active=False) AND "
	strSQL = strSQL & "(AdapterReturned=False OR CaseReturned=False)"
	Set objItemCheck = Application("Connection").Execute(strSQL)

	If objItemCheck.EOF Then
		strSQL = "UPDATE People SET Warning=False WHERE UserName='" & Replace(strUserName,"'","''") & "'"
		Application("Connection").Execute(strSQL)
	End If

End Sub%>

<%Sub ReturnLoanedItem

		Dim intLoanID, strSQL, objUserID, objUserCheck, intOwedID

		intLoanID = Request.Form("LoanID")
		intOwedID = Request.Form("OwedID")

		If Not intLoanID = "" Then

			strSQL = "SELECT AssignedTo,Item FROM Loaned WHERE ID=" & intLoanID
			Set objUserID = Application("Connection").Execute(strSQL)

			strSQL = "UPDATE Loaned SET Returned=True,ReturnDate=#" & Date & "# WHERE ID=" & intLoanID
			Application("Connection").Execute(strSQL)

			strSQL = "SELECT ID FROM Loaned WHERE Returned=False AND AssignedTo=" & objUserID(0)
			Set objUserCheck = Application("Connection").Execute(strSQL)

			If objUserCheck.EOF Then
				strSQL = "UPDATE People SET Loaned=False WHERE ID=" & objUserID(0)
				Application("Connection").Execute(strSQL)
			End If

			UpdateLog "LoanedOutItemReturned","",strUserName,"",objUserID(1),""

		End If

		If Not intOwedID = "" Then

			strSQL = "SELECT OwedBy,Item,Price FROM Owed WHERE ID=" & intOwedID
			Set objUserID = Application("Connection").Execute(strSQL)

			strSQL = "UPDATE Owed SET Active=False,PaidDate=Date() WHERE ID=" & intOwedID
			Application("Connection").Execute(strSQL)

			strSQL = "SELECT ID FROM Owed WHERE Active=True AND OwedBy=" & objUserID(0)
			Set objUserCheck = Application("Connection").Execute(strSQL)

			If objUserCheck.EOF Then
				strSQL = "UPDATE People SET Warning=False WHERE ID=" & objUserID(0)
				Application("Connection").Execute(strSQL)
			End If

			UpdateLog "ItemReturned","",strUserName,"",objUserID(1),""

			EMailBusinessOffice "ItemReturned", objUserID(1), objUserID(2)

		End If

End Sub%>

<%Sub LoanOutItem

	Dim intUserID, strItem, strSQL

	intUserID = Request.Form("UserID")
	strItem = Request.Form("Item")

	If strItem <> "" Then
		strSQL = "INSERT INTO Loaned (Item,AssignedTo,LoanDate)" & vbCRLF
		strSQL = strSQL & "VALUES ("
		strSQL = strSQL & "'" & Replace(strItem,"'","''") & "',"
		strSQL = strSQL & intUserID & ","
		strSQL = strSQL & "#" & Date & "#)"
		Application("Connection").Execute(strSQL)

		strSQL = "UPDATE People SET Loaned=True WHERE ID=" & intUserID
		Application("Connection").Execute(strSQL)

		UpdateLog "LoanedOutItem","",strUserName,"",strItem,""

	End If

End Sub%>

<%Sub BillUser

	Dim intUserID, strItem, strSQL, objPrice,intLoanID, objLoanedItem, bolReturnable, objUserCheck

	intUserID = Request.Form("UserID")
	strItem = Request.Form("Item")
	intLoanID = Request.Form("LoanID")

	If Not intLoanID = "" Then
		strSQL = "SELECT AssignedTo,Item FROM Loaned WHERE ID=" & intLoanID
		Set objLoanedItem = Application("Connection").Execute(strSQL)

		If Not objLoanedItem.EOF Then
			intUserID = objLoanedItem(0)
			strItem = objLoanedItem(1)
		End If

		strSQL = "UPDATE Loaned SET Returned=True,ReturnDate=Date() WHERE ID=" & intLoanID
		Application("Connection").Execute(strSQL)

		strSQL = "SELECT ID FROM Loaned WHERE Returned=False AND AssignedTo=" & intUserID
		Set objUserCheck = Application("Connection").Execute(strSQL)

		If objUserCheck.EOF Then
			strSQL = "UPDATE People SET Loaned=False WHERE ID=" & intUserID
			Application("Connection").Execute(strSQL)
		End If

	End If

	If strItem <> "" Then

		strSQL = "SELECT Price FROM Purchasable WHERE Item='" & strItem & "'"
		Set objPrice = Application("Connection").Execute(strSQL)

		Select Case strItem
			Case "Insurance Copay"
				bolReturnable = False
			Case Else
				bolReturnable = True
		End Select

		strSQL = "INSERT INTO Owed (Item,OwedBy,Price,RecordedDate,Active,Returnable)" & vbCRLF
		strSQL = strSQL & "VALUES ("
		strSQL = strSQL & "'" & Replace(strItem,"'","''") & "',"
		strSQL = strSQL & intUserID & "," & objPrice(0) & ","
		strSQL = strSQL & "#" & Date & "#,True," & bolReturnable & ")"
		Application("Connection").Execute(strSQL)

		strSQL = "UPDATE People SET Warning=True WHERE ID=" & intUserID
		Application("Connection").Execute(strSQL)

		If intLoanID = "" Then
			UpdateLog "MoneyOwed","",strUserName,"",strItem & " - $" & objPrice(0),""
		Else
			UpdateLog "LoanedItemBilled","",strUserName,"",strItem & " - $" & objPrice(0),""
		End If

		EMailBusinessOffice "BillUser", strItem, objPrice(0)

	End If

End Sub%>

<%Sub UserPaid

	Dim intOwedID, strSQL, objUserID, objUserCheck

	intOwedID = Request.Form("OwedID")

	If Not intOwedID = "" Then

		strSQL = "SELECT OwedBy,Item,Price FROM Owed WHERE ID=" & intOwedID
		Set objUserID = Application("Connection").Execute(strSQL)

		strSQL = "UPDATE Owed SET Active=False,PaidDate=Date() WHERE ID=" & intOwedID
		Application("Connection").Execute(strSQL)

		strSQL = "SELECT ID FROM Owed WHERE Active=True AND OwedBy=" & objUserID(0)
		Set objUserCheck = Application("Connection").Execute(strSQL)

		If objUserCheck.EOF Then
			strSQL = "UPDATE People SET Warning=False WHERE ID=" & objUserID(0)
			Application("Connection").Execute(strSQL)
		End If

		UpdateLog "ItemPaidFor","",strUserName,"",objUserID(1) & " - $" & objUserID(2),""

	End If

End Sub%>

<%Sub SetUserPasswordNotToExpire

	Dim intUserID, strSQL

	intUserID = Request.Form("UserID")

	'Update the database
	strSQL = "UPDATE People SET PWordNeverExpires=True WHERE ID=" & intUserID
	Application("Connection").Execute(strSQL)

	'Disable any pending tasks related to changing the password status
	strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='PasswordExpires' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)
	strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='PasswordNeverExpires' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)

	'Add the tasks to the pending tasks table
	strSQL = "INSERT INTO PendingTasks (Task,UserName,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
	strSQL = strSQL & "PasswordNeverExpires','" & Replace(strUserName,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
	Application("Connection").Execute(strSQL)

	'Record the action to the log
	UpdateLog "PasswordNeverExpiresSet","",strUserName,"Password Expires","Password Doesn't Expire",""

End Sub%>

<%Sub SetUserPasswordToExpire

	Dim intUserID, strSQL

	intUserID = Request.Form("UserID")

	'Update the database
	strSQL = "UPDATE People SET PWordNeverExpires=False WHERE ID=" & intUserID
	Application("Connection").Execute(strSQL)

	'Disable any pending tasks related to changing the password status
	strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='PasswordExpires' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)
	strSQL = "UPDATE PendingTasks SET Active=False WHERE Task='PasswordNeverExpires' AND Active=True AND UserName='" & Replace(strUserName,"'","''") & "'"
	Application("Connection").Execute(strSQL)

	'Add the tasks to the pending tasks table
	strSQL = "INSERT INTO PendingTasks (Task,UserName,UpdatedBy,TaskDate,TaskTime,Active) VALUES ('"
	strSQL = strSQL & "PasswordExpires','" & Replace(strUserName,"'","''") & "','" & strUser & "',#" & Date() & "#,#" & Time() & "#,True)"
	Application("Connection").Execute(strSQL)

	'Record the action to the log
	UpdateLog "PasswordExpiresSet","",strUserName,"Password Doesn't Expires","Password Expires",""

End Sub%>

<%Sub UpdateUserPassword

	'This will update the password on a user's account

	Const ADS_SCOPE_SUBTREE = 2

	Dim objRootDSE, objUserLookUp, objUser, arrUserData, strSQL, objOldPassword, strOldPassword, objADCommand
	Dim objADConnection, strAdminUsername, strAdminPassword, strNewPassword, intUserID, strUserName
	Dim objDirectory, bolRequireChange

	'Get the variables from the form
	intUserID = Request.Form("UserID")
	strAdminUserName = Request.Form("AdminUserName")
	strAdminPassword = Request.Form("AdminPassword")
	strNewPassword = Request.Form("Password")
	bolRequireChange = Request.Form("RequireChange")

	'Fix the admin username if it's an email address
	If InStr(strAdminUserName,"@") Then
		strAdminUserName = Left(strAdminUserName,InStr(strAdminUserName,"@") - 1)
	End If

	'Fix the admin username if it's in legacy form
	If InStr(strAdminUserName,"\") Then
		strAdminUserName = Right(strAdminUserName,Len(strAdminUserName) - InStr(strAdminUserName,("\")))
	End If

	'Get the username from the database
	strSQL = "SELECT Username, PWordNeverExpires FROM People WHERE ID=" & intUserID
	Set objUserLookUp = Application("Connection").Execute(strSQL)

	'Get the username from the returned object
	If Not objUserLookUp.EOF Then
		strUserName = objUserLookUp(0)
	Else
		strUserMessage = "<div Class=""Error"">User Not Found</div>"
		Exit Sub
	End If

	'This is where we have to try and catch errors
	On Error Resume Next

	'Create a RootDSE object for the domain
   Set objRootDSE = GetObject("LDAP://RootDSE")

   'Establish a connection to Active Directory using ActiveX Data Object
   Set objADConnection = CreateObject("ADODB.Connection")
   Set objADCommand = CreateObject("ADODB.Command")
   objADConnection.Provider = "ADsDSOObject"

   'Create the command object and attach it to the connection object
	objADConnection.Properties("User ID") = strAdminUserName & "@" & Application("Domain")
	objADConnection.Properties("Password") = strAdminPassword
	objADConnection.Properties("Encrypt Password") = TRUE
	objADConnection.Properties("ADSI Flag") = 1
	objADConnection.Open "Active Directory Provider"
	Set objADCommand.ActiveConnection = objADConnection
	objADCommand.Properties("Page Size") = 1000
	objADCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE

	'Create a RootDSE object for the domain
	Set objRootDSE = GetObject("LDAP://RootDSE")

	'Get the user object from Active Directory
	objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
	">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
	Set objUserLookup = objADCommand.Execute

	'If there was an erro then exit without doing anything
	If Err Then
		Select Case Err.Number
			Case -2147217911
				strUserMessage = "<div Class=""Error"">Invalid Credentials</div>"
			Case Else
				strUserMessage = "<div Class=""Error"">Error Connecting to Active Directory</div>"
		End Select
		Err.Clear
		Exit Sub
	End If

	'This is a new way to do it, kind of neat.
	'objADCommand.CommandText = "SELECT distinguishedName FROM 'LDAP://" & Application("DomainController") & "/" & objRootDSE.Get("DefaultNamingContext") & _
	'	"' WHERE objectCategory='user' AND samAccountName='" & Replace(strUserName,"'","''") & "'"

	'We have to use this different way to connect to the user object so we can pass the username and password
	Set objDirectory=getobject("LDAP:")
	Set objUser=objDirectory.OpenDSObject("LDAP://" &  objUserLookup(0), strAdminUserName & "@" & Application("Domain"), strAdminPassword, 1)

	'Set the password and save it
	objUser.SetPassword(strNewPassword)
	objUser.SetInfo

	'Since the password change went well we need to do the work to set have the user change the password if needed.
	'If the password needs to be changed then we have to turn off the password never expires if it's on.
	If bolRequireChange Then
		If objUserLookUp(1) Then
			SetUserPasswordToExpire
		End If
		objUser.Put "userAccountControl", 512
		objUser.Put "PwdLastSet", 0
		objUser.SetInfo
	End If

	strUserMessage = "<div Class=""Information"">Password Changed</div>"

	'Update the log
   UpdateLog "UserUpdatedPassword","",strUserName,"","",""

   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing

End Sub%>

<%Sub EMailBusinessOffice(strType,strItem,intPrice)

	Const cdoSendUsingPickup = 1

   Dim strSMTPPickupFolder, strFrom, objMessage, objConf, strMessage, strBCC
   Dim strSQL, intUserID, strUserName, objUserID, strSubject, strUser, objNetwork
   Dim objFSO, strEMailPath, txtEMailMessage

	strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"
	strFrom = Application("EMailNotifications")
	
	'Get the current user's email address
	Set objNetwork = CreateObject("WSCRIPT.Network")
   strUser = objNetwork.UserName
	If LCase(Left(strUser,4)) = "iusr" Then
      strUser = GetUser
   End If
   strUser = strUser & "@" & Application("Domain")

	'Create the objects required to send the mail.
	Set objMessage = CreateObject("CDO.Message")
	Set objConf = objMessage.Configuration
	
	With objConf.Fields
		.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
		.item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
		.Update
	End With

	'Get the email message
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strEMailPath = Left(Request.ServerVariables("PATH_TRANSLATED"),Len(Request.ServerVariables("PATH_TRANSLATED"))-8)
	strEMailPath = strEMailPath & "..\..\Scripts\EMail\"
	Set txtEMailMessage = objFSO.OpenTextFile(strEMailPath & strType & ".txt")

	'Read in the stored email message
	strMessage = txtEMailMessage.ReadAll
	txtEMailMessage.Close

	'Merge the variables into the message
	strMessage = Replace(strMessage,"#USERURL#","<a href=""http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?UserName=" & objUser(3) & """>" & _
		objUser(1) & " " & objUser(2) & "</a>")
	strMessage = Replace(strMessage,"#ITEM#",strItem)
	strMessage = Replace(strMessage,"#PRICE#",intPrice)

	'Set the subject of the message
	Select Case strType
		Case "BillUser"
			strSubject = "User Owes $" & intPrice & " for " & strItem

		Case "ItemReturned"
			strSubject = "Equipment has been Returned"

	End Select

	'Send the message
	With objMessage
		.To = Application("BusinessOfficeEMail")
		.CC = strUser
		.From = strFrom
		.Subject = strSubject
		.HTMLBody = strMessage
		If strBCC <> "" Then
			.BCC = strBCC
		End If
	  .Send
	End With

	'Close objects
	Set objMessage = Nothing
	Set objConf = Nothing
	Set objFSO = Nothing

End Sub%>

<%Sub EMailGuardian(strType)

	Const cdoSendUsingPickup = 1

   Dim strSMTPPickupFolder, objMessage, objConf, strMessage, strBCC
   Dim strSQL, intUserID, strUserName, objUserID, strSubject, strUser, objNetwork
   Dim objOwes, objTotal, objFSO, strEMailPath, txtEMailMessage, strItemList

	strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"

	'Get the current user's email address
	Set objNetwork = CreateObject("WSCRIPT.Network")
   strUser = objNetwork.UserName
	If LCase(Left(strUser,4)) = "iusr" Then
      strUser = GetUser
   End If
   strUser = strUser & "@" & Application("Domain")

	'Create the objects required to send the mail.
	Set objMessage = CreateObject("CDO.Message")
	Set objConf = objMessage.Configuration
	With objConf.Fields
		.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
		.item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
		.Update
	End With

	'Get the email message
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strEMailPath = Left(Request.ServerVariables("PATH_TRANSLATED"),Len(Request.ServerVariables("PATH_TRANSLATED"))-8)
	strEMailPath = strEMailPath & "..\..\Scripts\EMail\"
	Set txtEMailMessage = objFSO.OpenTextFile(strEMailPath & strType & ".txt")
	
	'Read in the stored email message
	strMessage = txtEMailMessage.ReadAll
	txtEMailMessage.Close
	
	Select Case strType
		Case "OwesMoney"
		
			'Set the subject of the email
			strSubject = "Owes Money"
		
			'Get the total amount owed
			strSQL = "SELECT Sum(Price) AS SumOfPrice FROM Owed GROUP BY OwedBy,Active Having Active=True AND OwedBy=" & objUser(0)
			Set objTotal = Application("Connection").Execute(strSQL)
		
			'Get the list of things they owe money for
			strSQL = "SELECT ID,Item,Price,RecordedDate,Returnable FROM Owed WHERE Active=True AND OwedBy=" & objUser(0) & " ORDER BY RecordedDate"
			Set objOwes = Application("Connection").Execute(strSQL)
		
			'Build the html list of items they owe money for
			strItemList = "<ul style=""margin: 0;"">"
			If Not objOwes.EOF Then
				Do Until objOwes.EOF
					strItemList = strItemList & "<li style=""margin: 0;"">" & objOwes(1) & ": $" & objOwes(2) & "</li>"
					ObjOwes.MoveNext
				Loop
			End If
			strItemList = strItemList & "</ul>"
			
			'Merge the variables into the message
			strMessage = Replace(strMessage,"#FIRSTNAME#",objUser(1))
			strMessage = Replace(strMessage,"#LASTNAME#",objUser(2))
			strMessage = Replace(strMessage,"#ITEMLIST#",strItemList)
			strMessage = Replace(strMessage,"#TOTAL#",objTotal(0))
			strMessage = Replace(strMessage,vbCRLF,"<br />")

	End Select

	'Send the email
	With objMessage
		.To = Request.Form("NotifyEmail")
		.BCC = strUser
		.From = strUser
		.Subject = strSubject
		.HTMLBody = strMessage
		If strBCC <> "" Then
			.BCC = strBCC
		End If
	  .Send
	End With

	'Close objects
	Set objMessage = Nothing
	Set objConf = Nothing
	Set objFSO = Nothing

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

<%Function GetAge(strDate)

   Dim strMonth, strDay, strYear, intIndex, datStartofYear, datEndofYear

   strMonth = Month(Date)
   strDay = Day(Date)
   strYear = Year(Date)

   For intIndex = 0 to 100

      datStartofYear = strMonth & "/" & strDay & "/" & strYear - intIndex - 1
      datEndofYear = strMonth & "/" & strDay & "/" & strYear - intIndex
      If CDate(strDate) >= CDate(datStartofYear) And CDate(strDate) <= CDate(datEndofYear)  Then
         GetAge = intIndex + 1
         Exit For
      End If

   Next

End Function %>

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

		strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & Replace(strUserName,"'","''") & "'"
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
' Source http://www.aspfree.com/c/a/ASP-Code/VBScript-function-to-validate-Email-Addresses/
' Function IsEmailValid(strEmail)
' Action: checks if an email is correct.
' Parameter: strEmail - the Email address
' Returned value: on success it returns True, else False.
Function IsEmailValid(strEmail)
 
    Dim strArray
    Dim strItem
    Dim i
    Dim c
    Dim blnIsItValid
 
    ' assume the email address is correct 
    blnIsItValid = True
   
    ' split the email address in two parts: name@domain.ext
    strArray = Split(strEmail, "@")
 
    ' if there are more or less than two parts 
    If UBound(strArray) <> 1 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If
 
    ' check each part
    For Each strItem In strArray
        ' no part can be void
        If Len(strItem) <= 0 Then
            blnIsItValid = False
            IsEmailValid = blnIsItValid
            Exit Function
        End If
       
        ' check each character of the part
        ' only following "abcdefghijklmnopqrstuvwxyz_-.'"
        ' characters and the ten digits are allowed
        For i = 1 To Len(strItem)
               c = LCase(Mid(strItem, i, 1))
               ' if there is an illegal character in the part
               If InStr("abcdefghijklmnopqrstuvwxyz_-.'", c) <= 0 And Not IsNumeric(c) Then
                   blnIsItValid = False
                   IsEmailValid = blnIsItValid
                   Exit Function
               End If
        Next
  
      ' the first and the last character in the part cannot be . (dot)
        If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then
           blnIsItValid = False
           IsEmailValid = blnIsItValid
           Exit Function
        End If
    Next
 
    ' the second part (domain.ext) must contain a . (dot)
    If InStr(strArray(1), ".") <= 0 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If
 
    ' check the length oh the extension 
    i = Len(strArray(1)) - InStrRev(strArray(1), ".")
    ' the length of the extension can be only 2, 3, or 4
    ' to cover the new "info" extension
    If i <> 2 And i <> 3 And i <> 4 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If

    ' after . (dot) cannot follow a . (dot)
    If InStr(strEmail, "..") > 0 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If
 
    ' finally it's OK 
    IsEmailValid = blnIsItValid
   
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
