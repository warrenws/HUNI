<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 2/16/16
'Last Updated 1/14/18

'This page is used to add devices and users

'Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser, objReports, strReport, strSubmitTo, strColumns
Dim objSites, objMakes, objModels, objRooms
Dim strNewDeviceMessage, intTag, intBOCESTag, strSerial, strMake, strModel, strDeviceSite
Dim strRoom, datPurchased, strMACAddress, strAppleID, objRoles, strFirstName, strLastName
Dim strUsername, intPhotoID, intRole, intClassOf, strNewUserMessage, strUserSite
Dim objPendingUsers, objDeviceTypes, strDeviceType, objLastNames, strPassword, objDescriptions
Dim strDescription, bolRequireChange, strTags

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions

	GetVariablesFromForm

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
		Case "Add Device"
         strNewDeviceMessage = GetNewDeviceMessage(AddDevice(True))

   	Case "Add User"
   		strNewUserMessage = GetNewUserMessage(AddUser(True),strUserName)
   		strUsername = ""

   	Case "Activate Student"
   		ActivateStudent

   End Select

	If strNewDeviceMessage = "" Then
		strNewDeviceMessage = "* Required"
	End If

'	If strNewUserMessage = "" Then
'		strNewUserMessage = "* Required"
'	End If

   'Get the URL used to submit forms
   If Request.ServerVariables("QUERY_STRING") = "" Then
      strSubmitTo = "add.asp"
   Else
      strSubmitTo = "add.asp?" & Request.ServerVariables("QUERY_STRING")
   End If

   'Get the data for the sites drop down menu
   strSQL = "SELECT Site FROM Sites WHERE Active=True ORDER BY Site"
   Set objSites = Application("Connection").Execute(strSQL)

   'Get the list of roles for the drop down menu
   strSQL = "SELECT Role,RoleID FROM Roles WHERE Active=True ORDER BY Role"
   Set objRoles = Application("Connection").Execute(strSQL)

   'Get the list of device types for the auto complete
   strSQL = "SELECT DISTINCT DeviceType FROM Devices WHERE Active=True And DeviceType<>''"
   Set objDeviceTypes = Application("Connection").Execute(strSQL)

   'Get the list of makes for the auto complete
   strSQL = "SELECT DISTINCT Manufacturer FROM Devices WHERE Active=True And Manufacturer<>''"
   Set objMakes = Application("Connection").Execute(strSQL)

   'Get the list of models for the auto complete
   strSQL = "SELECT DISTINCT Model FROM Devices WHERE Active=True And Model<>''"
   Set objModels = Application("Connection").Execute(strSQL)

   'Get the list of rooms for the auto complete
   strSQL = "SELECT DISTINCT Room FROM Devices WHERE Active=True And Room<>''"
   Set objRooms = Application("Connection").Execute(strSQL)

   'Get the list of pending accounts
   strSQL = "SELECT StudentID,FirstName,LastName,UserName,HomeRoom FROM People WHERE PWord='NewAccount'"
   Set objPendingUsers = Application("Connection").Execute(strSQL)

   'Get the list of lastnames for the auto complete
   strSQL = "SELECT DISTINCT LastName FROM People WHERE Active=True"
   Set objLastNames = Application("Connection").Execute(strSQL)

   'Get the list of descriptions for the auto complete
   strSQL = "SELECT DISTINCT Description FROM People WHERE Active=True And Description<>'' And Role='Teacher'"
   Set objDescriptions = Application("Connection").Execute(strSQL)

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
		<script>
			$(function() {

			<% If Not IsMobile Then %>
					$( document ).tooltip({track: true});
			<% End If %>

				document.getElementById("Tag").focus();
				document.getElementById("AddUserButton").disabled = true;
				document.getElementById("AddDeviceButton").disabled = true;

			<%	If Not objPendingUsers.EOF Then
					Do Until objPendingUsers.EOF %>
						document.getElementById("<%=objPendingUsers(3)%>Button").disabled = true;
					<%	objPendingUsers.MoveNext
					Loop
					objPendingUsers.MoveFirst
				End If %>

				$( "#PurchasedDate" ).datepicker({
					changeMonth: true,
					changeYear: true,
					showOtherMonths: true,
					selectOtherMonths: true,
					onClose: function( selectedDate ) {
				   	$( "#to" ).datepicker( "option", "minDate", selectedDate );
					}
				});

			<% If Not objDeviceTypes.EOF Then %>
				var possibleDeviceTypes = [
   		<% Do Until objDeviceTypes.EOF %>
					"<%=objDeviceTypes(0)%>",
				<%	objDeviceTypes.MoveNext
   			Loop
   			objDeviceTypes.MoveFirst%>
   		];
				$( "#DeviceTypes" ).autocomplete({
					source: possibleDeviceTypes
				});
		<% End If %>

			<% If Not objMakes.EOF Then %>
					var possibleMakes = [
				<% Do Until objMakes.EOF %>
						"<%=objMakes(0)%>",
					<%	objMakes.MoveNext
					Loop %>
				];
					$( "#Makes" ).autocomplete({
						source: possibleMakes
					});
			<% End If %>

			<% If Not objModels.EOF Then %>
					var possibleModels = [
				<% Do Until objModels.EOF %>
						"<%=objModels(0)%>",
					<%	objModels.MoveNext
					Loop %>
				];
					$( "#Models" ).autocomplete({
						source: possibleModels
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

		  });
		<% If Not objPendingUsers.EOF Then
				Do Until objPendingUsers.EOF %>

					function check<%=objPendingUsers(3)%>Password() {
						var password = document.getElementById("<%=objPendingUsers(3)%>Password");
						var lengthGood;

						if (password.value.length >= 8) {
							document.getElementById("<%=objPendingUsers(3)%>Icon").src="../images/good.png";
							lengthGood = true;
						} else {
							document.getElementById("<%=objPendingUsers(3)%>Icon").src="../images/notgood.png";
							lengthGood = false;
						}

						if (lengthGood) {
							document.getElementById("<%=objPendingUsers(3)%>Button").disabled = false;
						} else {
							document.getElementById("<%=objPendingUsers(3)%>Button").disabled = true;
						}
					};
				<% objPendingUsers.MoveNext
				Loop
				objPendingUsers.MoveFirst
			End If %>

		  	function buildUserName() {

		  		var firstName = document.getElementById("NewFirstName");
		  		var lastName = document.getElementById("NewLastName");
		  		var username = document.getElementById("NewUserName");
		  		username.value = lastName.value.toLowerCase() + firstName.value.substring(0,1).toLowerCase();
		  		checkNewUserForm();
		  	};

		  	function checkNewUserForm() {

		  		var firstName = document.getElementById("NewFirstName");
		  		var lastName = document.getElementById("NewLastName");
		  		var username = document.getElementById("NewUserName");
		  		var password = document.getElementById("NewPassword");
		  		var confirmPassword = document.getElementById("ConfirmNewPassword");
		  		var role = document.getElementById("Role");
		  		var userSite = document.getElementById("UserSite");
		  		var addUserButton = document.getElementById("AddUserButton");
		  		var passwordValid = document.getElementById("ValidIcon");
		  		var passwordMatch = document.getElementById("MatchIcon");
		  		var requireChange = document.getElementById("RequireChange");
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
					if ((firstName.value != "") && (lastName.value != "") && (username.value != "")) {
						if ((role.value != "") && (userSite.value != "")) {
							addUserButton.disabled = false;
						} else {
							addUserButton.disabled = true;
						}
					}
				}

		  	};

		  	function checkNewDeviceForm() {

		  		var tag = document.getElementById("Tag");
		  		var serial = document.getElementById("Serial");
		  		var deviceType = document.getElementById("DeviceTypes");
		  		var make	= document.getElementById("Makes");
		  		var model = document.getElementById("Models");
		  		var site = document.getElementById("DeviceSite");
		  		var purchaseDate = document.getElementById("PurchasedDate");
		  		var addDeviceButton = document.getElementById("AddDeviceButton");

		  		if ((tag.value != "") && (serial.value != "") && (deviceType.value != "")) {
		  			if ((make.value != "") && (model.value != "") && (site.value != "") && (purchaseDate.value != "")) {
		  				addDeviceButton.disabled = false;
		  			} else {
		  				addDeviceButton.disabled = true;
		  			}
		  		} else {
		  			addDeviceButton.disabled = true;
		  		}

		  	};

  		</script>
   </head>
   <body class="<%=strSiteVersion%>">

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

	<%
		JumpToDevice
		AddDeviceCard
		AddUserCard
		PendingAccounts
	%>

		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
   </body>

   </html>

<%End Sub%>

<%Sub AddDeviceCard %>

	<div Class="Card NormalCard">
      <form method="POST" action="<%=strSubmitTo%>">
      <div Class="CardTitle">Add Device</div>
      <div Class="CardColumn1">Asset Tag:* </div>
      <div Class="CardColumn2"><input class="Card InputWidthSmall" type="text" name="Tag" Value="<%=intTag%>" id="Tag" oninput="checkNewDeviceForm()" /></div>
      <div Class="CardColumn1">BOCES Tag: </div>
      <div Class="CardColumn2"><input class="Card InputWidthSmall" type="text" name="BOCESTag" Value="<%=intBOCESTag%>" /></div>
      <div Class="CardColumn1">Serial Number:* </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="Serial" Value="<%=strSerial%>" id="Serial" oninput="checkNewDeviceForm()" /></div>
      <div Class="CardColumn1">Device Type:* </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="DeviceType" Value="<%=strDeviceType%>" id="DeviceTypes" oninput="checkNewDeviceForm()" /></div>
      <div Class="CardColumn1">Make:* </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="Make" Value="<%=strMake%>" id="Makes" oninput="checkNewDeviceForm()" /></div>
      <div Class="CardColumn1">Model:* </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="Model" Value="<%=strModel%>" id="Models" oninput="checkNewDeviceForm()" /></div>
      <div>
			<div Class="CardColumn1">Site:* </div>
			<div Class="CardColumn2">
				<select Class="Card" name="DeviceSite" id="DeviceSite"oninput="checkNewDeviceForm()" >
					<option value=""></option>
				<% Do Until objSites.EOF
						If strDeviceSite = objSites(0) Then
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
		<div Class="CardColumn1">Room: </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="Room" Value="<%=strRoom%>" id="Rooms" /></div>
      <div Class="CardColumn1">Purchased:* </div>
      <div Class="CardColumn2"><input class="Card InputWidthSmall" type="text" name="Purchased" Value="<%=datPurchased%>" id="PurchasedDate" onchange="checkNewDeviceForm()" /></div>

      <div Class="CardColumn1">Tag: </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="GroupTag" Value="<%=strTags%>" oninput="checkNewDeviceForm()" /></div>
		<div>
         <div class="Button"><input type="submit" value="Add Device" name="Submit" id="AddDeviceButton" /></div>
      </div>
	<% If strNewDeviceMessage <> "" Then %>
		<div>
			<%=strNewDeviceMessage%>
		</div>
	<% End If %>
   </div>

<%End Sub%>

<%Sub AddUserCard %>

	<div Class="Card NormalCard">
      <form method="POST" action="<%=strSubmitTo%>">
      <div Class="CardTitle">Add User</div>
      <div Class="CardColumn1">First Name: </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="FirstName" Value="<%=strFirstName%>" id="NewFirstName" oninput="buildUserName()" /></div>
      <div Class="CardColumn1">Last Name: </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="LastName" Value="<%=strLastName%>" id="NewLastName" oninput="buildUserName()"/></div>
      <div Class="CardColumn1">Username: </div>
      <div Class="CardColumn2"><input class="Card InputWidthLarge" type="text" name="UserName" Value="<%=strUsername%>" id="NewUserName" oninput="checkNewUserForm()" /></div>
      <div Class="CardColumn1">Password:</div>
      <div Class="CardColumn2">
      	<input class="Card InputWidthSmall" type="password" name="Password" id="NewPassword" oninput="checkNewUserForm()" />
			<input class="Card" type="checkbox" name="RequireChange" id="RequireChange" oninput="checkNewUserForm()" /> Require
      </div>
      <div Class="CardColumn1">Confirm:</div>
      <div Class="CardColumn2">
      	<input class="Card InputWidthSmall" type="password" name="ConfirmPassword" id="ConfirmNewPassword" oninput="checkNewUserForm()" />
      	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Change
      </div>
      <div>
			<div Class="CardColumn1">Role: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="Role" id="Role" oninput="checkNewUserForm()">
					<option value=""></option>
				<% Do Until objRoles.EOF
						If intRole <> "" Then
							If CInt(intRole) = CInt(objRoles(1)) Then
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
		</div>
      <div>
			<div Class="CardColumn1">Site: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="UserSite" id="UserSite" oninput="checkNewUserForm()">
					<option value=""></option>
				<% Do Until objSites.EOF
						If strUserSite = objSites(0) Then
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
		<div>
			<div Class="CardColumn1">Description: </div>
			<div Class="CardColumn2"><input Class="Card InputWidthLarge" type="text" name="Description" id="Descriptions"></div>
		</div>
		<div>
			<div Class="CardMerged Center">Password Valid <image src="../images/notgood.png" class="Icon" id="ValidIcon" ></div>
		</div>
		<div>
			<div Class="CardMerged Center">Passwords Match <image src="../images/notgood.png" class="Icon" id="MatchIcon" ></div>
		</div>

      <div>

         <div class="Button"><input type="submit" value="Add User" name="Submit" id="AddUserButton" /></div>
      </div>
	<% If strNewUserMessage <> "" Then %>
		<div>
			<%=strNewUserMessage%>
		</div>
	<% End If %>
   </div>

<%End Sub%>

<%Sub PendingAccounts

	Dim objFSO,intStudentID, strFirstName, strLastName, strUserName, strHomeRoom

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If Not objPendingUsers.EOF Then %>
		<br />
	<%	Do Until objPendingUsers.EOF

			intStudentID = objPendingUsers(0)
			strFirstName = objPendingUsers(1)
			strLastName = objPendingUsers(2)
			strUserName = objPendingUsers(3)
			strHomeRoom = objPendingUsers(4)
			%>

			<div class="Card NormalCard">
				<form method="POST" action="<%=strSubmitTo%>">
				<input type="hidden" name="StudentUserName" value="<%=strUserName%>" />
				<div class="CardTitle"><%=strFirstName%>&nbsp;<%=strLastName%></div>
				<div Class="ImageSectionInCard">
				<% If objFSO.FileExists(Application("PhotoLocation") & "\" & intStudentID & "s\" & intStudentID & ".jpg") Then %>
						<img class="PhotoCard" src="/photos/students/<%=intStudentID%>.jpg" title="<%=intStudentID%>" width="96" />
				<% Else %>
						<img class="PhotoCard" src="/photos/students/missing.png" title="<%=intStudentID%>" width="96" />
				<% End If %>
				</div>
				<div Class="RightOfImageInCard">
					<div>
						<div Class="CardMerged">Guide: <%=objPendingUsers(4)%></div>
					</div>
					<div>
						<div Class="CardMerged">Username: <%=strUserName%> </div>
					</div>
					<div>
						<div Class="CardMerged">
							Password: <input class="Card InputWidthSmall" type="text" name="PWord" id="<%=strUserName%>Password" oninput="check<%=strUserName%>Password()" />
							&nbsp;&nbsp;<image src="../images/notgood.png" class="Icon" id="<%=strUserName%>Icon" >
						</div>
					</div>
					<div>
						<div Class="CardMerged">AUP Signed: <input type="checkbox" name="AUP" value="True" checked="checked" /></div>
						<div class="Button"><input type="submit" value="Activate Student" name="Submit" id="<%=strUserName%>Button" /></div>
					</div>
				</div>
				</form>
			</div>
		<% objPendingUsers.MoveNext
		Loop
	End If

	Set objFSO = Nothing

End Sub%>

<%Sub JumpToDevice%>

	<div Class="HeaderCard">
		<form method="POST" action="search.asp">
		Asset tag: <input class="Card InputWidthSmall" type="text" name="SmartBox" id="LastNames" />
		<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		<input type="hidden" value="" name="BOCESTag" />
		</form>
	</div>

<%End Sub%>

<%Function AddDevice(bolAddToInventory)

	Dim strSQL, objDeviceCheck, arrTags, objTagCheck

	'See if the needed values are present
	If intTag = "" Then
		AddDevice = 1
		Exit Function
	End If

	If strSerial = "" Then
		AddDevice = 2
		Exit Function
	End If

	If strMake = "" Then
		AddDevice = 3
		Exit Function
	End If

	If strModel = "" Then
		AddDevice = 4
		Exit Function
	End If

	If strDeviceSite = "" Then
		AddDevice = 5
		Exit Function
	End If

	If datPurchased = "" Then
		AddDevice = 6
		Exit Function
	End If

	If strDeviceType = "" Then
		AddDevice = 13
		Exit Function
	End If

	'Make sure the asset tag doesn't already exist
	strSQL = "SELECT ID FROM Devices WHERE LGTag='" & intTag & "'"
	Set objDeviceCheck = Application("Connection").Execute(strSQL)
	If Not objDeviceCheck.EOF Then
		AddDevice = 7
		Exit Function
	End If

	'Make sure the BOCES tag doesn't already exist
	If intBOCESTag <> "" Then
		strSQL = "SELECT ID FROM Devices WHERE BOCESTag='" & intBOCESTag & "'"
		Set objDeviceCheck = Application("Connection").Execute(strSQL)
		If Not objDeviceCheck.EOF Then
			AddDevice = 8
			Exit Function
		End If
	End If

	'Make sure the serial number doesn't already exist
	strSQL = "SELECT ID FROM Devices WHERE SerialNumber='" & strSerial & "'"
	Set objDeviceCheck = Application("Connection").Execute(strSQL)
	If Not objDeviceCheck.EOF Then
		AddDevice = 9
		Exit Function
	End If

	'Make sure the date is in date format
	If Not IsDate(datPurchased) Then
		AddDevice = 10
		Exit Function
	End If

	'Make sure the MACAddress is in the right format
	If strMACAddress <> "" Then
		If Not ValidMACAddress(strMACAddress) Then
			AddDevice = 11
			Exit Function
		End If
	End If

	'Make sure the AppleID is in the right format
	If strAppleID <> "" Then
		If Not ValidEMailAddress(strAppleID) Then
			AddDevice = 12
			Exit Function
		End If
	End If

	If bolAddToInventory Then

		'Add the device to the inventory
		strSQL = "INSERT INTO Devices (LGTag,BOCESTag,SerialNumber,Manufacturer,Model,Site,Room,DatePurchased,MACAddress,AppleID,Active,DateAdded,DeviceType)" & vbCRLF
		strSQL = strSQL & "VALUES ('"
		strSQL = strSQL & Replace(intTag,"'","''") & "','"
		strSQL = strSQL & intBOCESTag & "','"
		strSQL = strSQL & Replace(strSerial,"'","''") & "','"
		strSQL = strSQL & Replace(strMake,"'","''") & "','"
		strSQL = strSQL & Replace(strModel,"'","''") & "','"
		strSQL = strSQL & Replace(strDeviceSite,"'","''") & "','"
		strSQL = strSQL & Replace(strRoom,"'","''") & "',#"
		strSQL = strSQL & datPurchased & "#,'"
		strSQL = strSQL & FixMACAddress(strMACAddress) & "','"
		strSQL = strSQL & Replace(strAppleID,"'","''") & "',True,#" & Date() & "#,'" & Replace(strDeviceType,"'","''") & "')"
		Application("Connection").Execute(strSQL)

		UpdateLog "DeviceAdded",intTag,"","",strMake & " - " & strModel,""

		'Add the tags to the database
		arrTags = Split(strTags,",")
		For intIndex = 0 to UBound(arrTags)
			strTag = Trim(arrTags(intIndex))

			If strTag <> "" Then

				strSQL = "SELECT Tag FROM Tags WHERE Tag='" & Replace(strTag,"'","''") & "' AND LGTag='" & intTag & "'"
				Set objTagCheck = Application("Connection").Execute(strSQL)

				If objTagCheck.EOF Then
					strSQL = "INSERT INTO Tags (LGTag, Tag) VALUES ('" & intTag & "','" & Replace(strTag,"'","''") & "')"
					Application("Connection").Execute(strSQL)

					UpdateLog "DeviceUpdatedTagAdded",intTag,"","",strTag,""

				End If

			End If

			Set objTagCheck = Nothing

		Next

		'Reset the unique values
		intTag = ""
		intBOCESTag = ""
		strSerial = ""
		strMACAddress = ""
		strAppleID = ""

		AddDevice = 0
	Else
		AddDevice = -1
	End If

End Function%>

<%Function AddUser(bolAddToInventory)

	Dim strSQL, objUserCheck, strUserRole, intUserClassOf, bolAUP

	'See if the needed values are present
	If strFirstName = "" Then
		AddUser = 1
		Exit Function
	End If

	If strLastName = "" Then
		AddUser = 2
		Exit Function
	End If

	If strUsername = "" Then
		AddUser = 3
		Exit Function
	End If

	If intRole = "" Then
		AddUser = 4
		Exit Function
	End If

	If strUserSite = "" Then
		AddUser = 5
		Exit Function
	End If

	'Make sure the user doesn't already exist
	strSQL = "SELECT ID FROM People WHERE Username='" & strUserName & "'"
	Set objUserCheck = Application("Connection").Execute(strSQL)
	If Not objUserCheck.EOF Then
		AddUser = 6
		Exit Function
	End If

	'Make sure a user with the same first name and last name don't already exist in the location and role.
	'This helps with syncing to Active Directory
	strSQL = "SELECT ID FROM People Where FirstName='" & strFirstName & "' AND LastName='" & strLastName & "' AND "
	strSQL = strSQL & "ClassOf=" & intRole & " AND Site='" & strUserSite & "'"
	Set objUserCheck = Application("Connection").Execute(strSQL)

	If Not objUserCheck.EOF Then
		AddUser = 6
		Exit Function
	End If

	'Prepare the Role and Class Of information for the database
	If intRole >= 2000 Then
		If intClassOf <> "" Then
			If IsNumeric(intClassOf) Then
				If CInt(intClassOf) < Year(Date) Then
					AddUser = 9
					Exit Function
				End If
				If CInt(intClassOf) > Year(Date) + 13 Then
					AddUser = 10
					Exit Function
				End If
				intUserClassOf = intClassOf
				strUserRole = "Student"
				bolAUP = False
			Else
				AddUser = 7
				Exit Function
			End If
		Else
			AddUser = 8
			Exit Function
		End If
	Else
		intUserClassOf = intRole
		strUserRole = "Teacher"
		bolAUP = True
	End If

	If intPhotoID <> "" Then
		If IsNumeric(intPhotoID) Then

			strSQL = "SELECT ID FROM People WHERE Role='" & strUserRole & "' AND StudentID=" & intPhotoID
			Set objUserCheck = Application("Connection").Execute(strSQL)

			If Not objUserCheck.EOF Then
				AddUser = 12
				Exit Function
			End If

		Else
			AddUser = 11
			Exit Function
		End If
	Else
		intPhotoID = 0
	End If

	If bolRequireChange = "on" Then
		bolRequireChange = True
	Else
		bolRequireChange = False
	End If

	If bolAddToInventory Then

		'Add the user to the inventory
		strSQL = "INSERT INTO People (FirstName,LastName,Username,Role,Description,ClassOf,Site,StudentID,Active,PWordNeverExpires,AUP,DateAdded)" & vbCRLF
		strSQL = strSQL & "VALUES ('"
		strSQL = strSQL & Replace(strFirstName,"'","''") & "','"
		strSQL = strSQL & Replace(strLastName,"'","''") & "','"
		strSQL = strSQL & Replace(strUsername,"'","''") & "','"
		strSQL = strSQL & Replace(strUserRole,"'","''") & "','"
		strSQL = strSQL & Replace(strDescription,"'","''") & "','"
		strSQL = strSQL & intUserClassOf & "','"
		strSQL = strSQL & Replace(strUserSite,"'","''") & "',"
		strSQL = strSQL & intPhotoID & ",True," & bolRequireChange & ","
		strSQL = strSQL & bolAUP & ",#" & Date() & "#)"
		Application("Connection").Execute(strSQL)

		'Add the task to the database
		strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ("
		strSQL = strSQL & "'CreateAccount','"
		strSQL = strSQL & strUserName & "','"
		strSQL = strSQL & strPassword & "','"
		strSQL = strSQL & strUser & "',#"
		strSQL = strSQL & Date() & "#,#"
		strSQL = strSQL & Time() & "#,True)"
		Application("Connection").Execute(strSQL)



		UpdateLog "UserAdded","",strUserName,"",GetDisplayName(strUserName),""

		'Reset the unique values
		strFirstName = ""
		strLastName = ""
		intPhotoID = ""

		AddUser = 0
	Else
		AddUser = -1
	End If

End Function%>

<%Sub ActivateStudent

	Dim strSQL, strPassword, bolAUP, strUserName

	'Get the variables from the form
	strPassword = CStr(Request.Form("PWord"))
	bolAUP = Request.Form("AUP")
	strUserName = Request.Form("StudentUserName")

	'Set the information in the database
	If bolAUP = "True" Then
		strSQL = "UPDATE People SET Deleted=False,AUP=True,Active=False,"
	Else
		strSQL = "UPDATE People SET Deleted=False,AUP=False,Active=False,"
	End If
	strSQL = strSQL & "PWord='" & strPassword & "',Notes='Waiting to be Activated - " & Date() & "' WHERE UserName='" & strUserName & "'"
	Application("Connection").Execute(strSQL)

	'Add the task to the database
	strSQL = "INSERT INTO PendingTasks (Task,UserName,NewValue,UpdatedBy,TaskDate,TaskTime,Active) VALUES ("
	strSQL = strSQL & "'ActivateStudent','"
	strSQL = strSQL & strUserName & "','"
	strSQL = strSQL & strPassword & "','"
	strSQL = strSQL & strUser & "',#"
	strSQL = strSQL & Date() & "#,#"
	strSQL = strSQL & Time() & "#,True)"
	Application("Connection").Execute(strSQL)

	UpdateLog "NewStudentPasswordSet","",strUserName,"","",""

End Sub%>

<%Sub GetVariablesFromForm

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
   
   If intTag = 0 Then
   	intTag = ""
   End If

	intBOCESTag = Request.Form("BOCESTag")
	strSerial = Request.Form("Serial")
	strDeviceType = Request.Form("DeviceType")
	strMake = Request.Form("Make")
	strModel = Request.Form("Model")
	strDeviceSite = Request.Form("DeviceSite")
	strRoom = Request.Form("Room")
	datPurchased = Request.Form("Purchased")
	strMACAddress = Request.Form("MACAddress")
	strAppleID = Request.Form("AppleID")
	strFirstName = Request.Form("FirstName")
	strLastName = Request.Form("LastName")
	strUsername = Request.Form("Username")
	intClassOf = Request.Form("ClassOf")
	intPhotoID = Request.Form("PhotoID")
	intRole = Request.Form("Role")
	strUserSite = Request.Form("UserSite")
	strPassword = Request.Form("Password")
	strDescription = Request.Form("Description")
	bolRequireChange = Request.Form("RequireChange")
	strTags = Request.Form("GroupTag")

	'Fix the serial number if it's an Apple
	If strMake = "Apple" And LCase(Left(strSerial,1)) = "s" Then
		strSerial = Right(strSerial,Len(strSerial) - 1)
	End If

	'Fix the username if it's an email address
	If InStr(strUserName,"@") Then
		strUserName = Left(strUserName,InStr(strUserName,"@") - 1)
	End If

	'Fix the username if it's in legacy form
	If InStr(strUserName,"\") Then
		strUserName = Right(strUserName,Len(strUserName) - InStr(strUserName,("\")))
	End If

End Sub%>

<%Function GetNewDeviceMessage(intError)

	Select Case intError
		Case -1
			GetNewDeviceMessage = "<div Class=""Information"">Looks Good</div>"
		Case 0
			GetNewDeviceMessage = "<div Class=""Information"">Device Added</div>"
		Case 1
			GetNewDeviceMessage = "<div Class=""Error"">Asset Tag Missing</div>"
		Case 2
			GetNewDeviceMessage = "<div Class=""Error"">Serial Number Missing</div>"
		Case 3
			GetNewDeviceMessage = "<div Class=""Error"">Make Missing</div>"
		Case 4
			GetNewDeviceMessage = "<div Class=""Error"">Model Missing</div>"
		Case 5
			GetNewDeviceMessage = "<div Class=""Error"">Site Missing</div>"
		Case 6
			GetNewDeviceMessage = "<div Class=""Error"">Data Purchased Missing</div>"
		Case 7
			GetNewDeviceMessage = "<div Class=""Error"">Duplicate Asset Tag</div>"
		Case 8
			GetNewDeviceMessage = "<div Class=""Error"">Duplicate BOCES Tag</div>"
		Case 9
			GetNewDeviceMessage = "<div Class=""Error"">Duplicate Serial Number</div>"
		Case 10
			GetNewDeviceMessage = "<div Class=""Error"">Purchased not in date form</div>"
		Case 11
			GetNewDeviceMessage = "<div Class=""Error"">Invalid MAC Address</div>"
		Case 12
			GetNewDeviceMessage = "<div Class=""Error"">Invalid Apple ID</div>"
		Case 13
			GetNewDeviceMessage = "<div Class=""Error"">Device Type Missing</div>"
		Case Else
			GetNewDeviceMessage = "<div Class=""Error"">Unknown Error #" & intError & "</div>"
	End Select

End Function%>

<%Function GetNewUserMessage(intError,strUserName)

	Dim strPage

	strPage = Left(Request.ServerVariables("SCRIPT_NAME"),InStrRev(Request.ServerVariables("SCRIPT_NAME"),"/")) & "user.asp?UserName="

	Select Case intError
		Case -1
			GetNewUserMessage = "<div Class=""Information"">Looks Good</div>"
		Case 0
			GetNewUserMessage = "<div Class=""Information""><a href=""" & strPage & strUserName & """>User Added</a></div>"
		Case 1
			GetNewUserMessage = "<div Class=""Error"">First Name Missing</div>"
		Case 2
			GetNewUserMessage = "<div Class=""Error"">Last Name Missing</div>"
		Case 3
			GetNewUserMessage = "<div Class=""Error"">User Name Missing</div>"
		Case 4
			GetNewUserMessage = "<div Class=""Error"">Role Missing</div>"
		Case 5
			GetNewUserMessage = "<div Class=""Error"">Site Missing</div>"
		Case 6
			GetNewUserMessage = "<div Class=""Error"">User Already Exists</div>"
		Case 7
			GetNewUserMessage = "<div Class=""Error"">Class Of Must Be a Year</div>"
		Case 8
			GetNewUserMessage = "<div Class=""Error"">Class Of Missing</div>"
		Case 9
			GetNewUserMessage = "<div Class=""Error"">Class Of Too Low</div>"
		Case 10
			GetNewUserMessage = "<div Class=""Error"">Class Of Too High</div>"
		Case 11
			GetNewUserMessage = "<div Class=""Error"">Photo ID Must Be a Number</div>"
		Case 12
			GetNewUserMessage = "<div Class=""Error"">Duplicate Photo ID</div>"
		Case Else
			GetNewUserMessage = "<div Class=""Error"">Unknown Error #" & intError & "</div>"
	End Select

End Function%>

<%Function ValidMACAddress(strMAC)

   Dim intIndex, strChar, bolValidChar, strTempMAC

   strTempMAC = FixMACAddress(strMAC)

   'Check and see if the MAC address is the correct length
   If Len(strTempMAC) = 17 Then

      'Loop through each character in the MAC address
      For intIndex = 1 to Len(strTempMAC)
         strChar = Mid(strTempMAC,intIndex,1)

         Select Case intIndex

            'If the separator isn't a ; then it's not right
            Case 3,6,9,12,15
               If strChar <> ":" Then
                  ValidMACAddress = False
                  Exit Function
               End If

            Case Else

               'If it's not a number make sure it's a character between a-f
               If Not IsNumeric(strChar) Then
                  bolValidChar = False
                  Select Case LCase(strChar)
                     Case "a", "b", "c", "d", "e", "f"
                        bolValidChar = True
                  End Select
                  If Not bolValidChar Then
                     ValidMACAddress = False
                     Exit Function
                  End If
               End If

         End Select

      Next

      ValidMACAddress = True

   Else

      ValidMACAddress = False

   End If

End Function%>

<%Function FixMACAddress(strMAC)

   Dim intIndex, strChar

   'Check and see if it's the correct length with or without separator
   If Len(strMAC) = 12 Or Len(strMAC) = 17 Then

      'Loop through each character
      For intIndex = 1 to Len(strMAC)

         'Add in the : to the result"
         Select Case Len(FixMACAddress)
            Case 2,5,8,11,14
               FixMACAddress = FixMACAddress & ":"
         End Select

         'Add the characer to the result if it's a valid character
         strChar = Mid(strMAC,intIndex,1)
         If Not IsNumeric(strChar) Then
            Select Case LCase(strChar)
               Case "a", "b", "c", "d", "e", "f"
                  FixMACAddress = FixMACAddress & LCase(strChar)
            End Select
         Else
            FixMACAddress = FixMACAddress & strChar
         End If

      Next

   Else

      'Return the orignal MAC address if it can't be fixed.
      FixMACAddress = strMAC

   End If

End Function%>

<%Function ValidEMailAddress(strEMail)

   Dim strUserName, strDomain

   If InStr(strEMail,"@") <> 0 Then

      strUserName = Left(strEMail,InStr(strEMail,"@") - 1)
      strDomain = Mid(strEMail,InStr(strEMail,"@") + 1,Len(strEMail)-InStr(strEMail,"@") + 1)

      If InStr(strDomain,"@") <> 0 Then
         ValidEMailAddress = False
         Exit Function
      End If

      If InStr(strDomain,".") = 0 Then
         ValidEMailAddress = False
         Exit Function
      End If

      ValidEMailAddress = True
   Else
      ValidEMailAddress = False
   End If

End Function%>

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
