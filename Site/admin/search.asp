<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 4/19/16
'Last Updated 1/14/18

'This is the search page for the inventory website

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim intTag, strSerial, objGuideRooms, intBOCESTag, objClassOf, objSites, strUserSite, strDeviceSite
Dim objOldestDevice, datOldestDevice, intYears, intIndex, strModel, strTags, strRoom, objRooms
Dim strSelected, intDeviceYear, strFirstName, strLastName, strGuideRoom, strWithDevice
Dim strRole, strAssignedYes, strAssignedNo, strWithDeviceYes, strWithDeviceNo, strUserStatusAll
Dim strDeviceMessage, strUserMessage, strUserStatusEnabled, strUserStatusDisabled
Dim objTags, Key, objTagList, strTagList, strNotes, strOwesYes, strOwesNo, strOwes
Dim strDeviceStatusEnabled, strDeviceStatusDisabled, strDeviceStatusAll
Dim objEventTypes, objCategories, strEventType, strCategory, strWarranty, strComplete
Dim strWarrantyYes, strWarrantyNo, strCompleteYes, strCompleteNo, strEventMessage, strCompleteAll
Dim datStartDate, datEndDate, objSavedSearches, intSavedSearch, strDeviceViewCard
Dim strDeviceViewTable, strUserViewCard, strUserViewTable, strEventViewCard, strEventViewTable
Dim strDeviceView, strUserView, strEventView, objItems, strUserLoaned, strLoanedAnything
Dim strColumns, strAUPYes, strAUPNo, strAUP, strUserNotes, strDeviceNotes, objDisabledUsers
Dim intDisabledUsersCount, intEventNumber, strEventSite, strEventModel, strSmartBox
Dim objMakes, objModels, objStudentsPerGrade, intHighestValue, strMake, objFirstNames, objLastNames
Dim objDeviceTypes, strDeviceType, strDescription, objRoles, strAvailableTags, objInternetTypes
Dim strInternetAccess, strIPAddress

'See if the user has the rights to visit this page
If AccessGranted Then
	ProcessSubmissions
Else
	DenyAccess
End If %>

<%Sub ProcessSubmissions

	Dim strSQL

	GetVariablesFromURL

	If intTag = 0 Then
		intTag = ""
	End If

	'Check and see if anything was submitted to the site
	Select Case Request.Form("Submit")
		Case "Lookup Device"
			LookupDevice
		Case "Lookup User"
			LookupUser
		Case "Lookup Event"
			LookupEvent
		Case "Search"
			Search
		Case "Load"
			LoadSearch
	End Select

	'Get the data for the guide rooms drop down menu
	strSQL = "SELECT DISTINCT HomeRoom FROM People WHERE Homeroom <> '' AND Active=True ORDER BY HomeRoom"
	Set objGuideRooms = Application("Connection").Execute(strSQL)

	'Get the data for the role drop down menu
	strSQL = "SELECT DISTINCT ClassOf From People WHERE ClassOf>1000 ORDER BY ClassOf DESC"
	Set objClassOf = Application("Connection").Execute(strSQL)

	'Get the list of roles for the drop down menu
	strSQL = "SELECT Role,RoleID FROM Roles WHERE Active=True ORDER BY Role"
	Set objRoles = Application("Connection").Execute(strSQL)

	'Get the data for the room drop down menu
	strSQL = "SELECT DISTINCT Room From Devices WHERE Active=True AND Room <> '' ORDER BY Room"
	Set objRooms = Application("Connection").Execute(strSQL)

	'Get the data for the sites drop down menu
	strSQL = "SELECT Site FROM Sites WHERE Active=True ORDER BY Site"
	Set objSites = Application("Connection").Execute(strSQL)

	'Get the data for the makes drop down menu
	strSQL = "SELECT DISTINCT Manufacturer From Devices WHERE Active=True AND Manufacturer <> '' ORDER BY Manufacturer"
	Set objMakes = Application("Connection").Execute(strSQL)

	'Get the list of event types for the event types drop down menu
	strSQL = "SELECT EventType FROM EventTypes WHERE Active=True ORDER BY EventType"
	Set objEventTypes = Application("Connection").Execute(strSQL)

	'Get the list of categories from the category drop down menu
	strSQL = "SELECT Category FROM Categories WHERE Active=True ORDER BY Category"
	Set objCategories = Application("Connection").Execute(strSQL)

	'Get the list of device types for the device type drop down menu
	strSQL = "SELECT DISTINCT DeviceType FROM Devices WHERE Active=True And DeviceType<>''"
	Set objDeviceTypes = Application("Connection").Execute(strSQL)

	'Get the list of models for the auto complete
	strSQL = "SELECT DISTINCT Model FROM Devices WHERE Active=True And Model<>''"
	Set objModels = Application("Connection").Execute(strSQL)

	'Get the list of items that are loaned out.
	strSQL = "SELECT DISTINCT Item FROM Loaned WHERE Returned=False"
	Set objItems = Application("Connection").Execute(strSQL)

	'Get the list of firstnames for the auto complete
	strSQL = "SELECT DISTINCT FirstName FROM People WHERE Active=True"
	Set objFirstNames = Application("Connection").Execute(strSQL)

	'Get the list of lastnames for the auto complete
	strSQL = "SELECT DISTINCT LastName FROM People WHERE Active=True"
	Set objLastNames = Application("Connection").Execute(strSQL)
	
	'Get the list of Internet types for the drop down menu
	strSQL = "SELECT InternetType FROM InternetTypes WHERE Active=True"
	Set objInternetTypes = Application("Connection").Execute(strSQL)

	'Get the oldest device from the inventory
	strSQL = "SELECT DatePurchased FROM Devices WHERE DatePurchased Is Not Null ORDER BY DatePurchased"
	Set objOldestDevice = Application("Connection").Execute(strSQL)
	If Not objOldestDevice.EOF Then
		datOldestDevice = objOldestDevice(0)
		intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice) + 1
	End If

	'Get number of devices assigned to disabled students
	strSQL = "SELECT ID, UserName FROM People WHERE People.Active=False AND People.HasDevice=True"
	Set objDisabledUsers = Application("Connection").Execute(strSQL)
	If Not objDisabledUsers.EOF Then
		Do Until objDisabledUsers.EOF
			intDisabledUsersCount = intDisabledUsersCount + 1
			objDisabledUsers.MoveNext
		Loop
	End If

	'Get all the tags from the database and process them
	strSQL = "SELECT DISTINCT Tag FROM Tags ORDER BY Tag"
	Set objTagList = Application("Connection").Execute(strSQL)

	If Not objTagList.EOF Then
		Do Until objTagList.EOF
			strTagList = strTagList & "<a href=""devices.asp?Tags=" & objTagList(0) & """>" & objTagList(0) & "</a>, "
			objTagList.MoveNext
		Loop
		strTagList = Left(strTagList,Len(strTagList) - 2)
	End If

	'Get all the saved searches from the database
	strSQL = "SELECT ID,SearchName FROM SavedSearches WHERE Active=True ORDER BY SearchName"
	Set objSavedSearches = Application("Connection").Execute(strSQL)

	'Display the error message if they were sent back to this page
	Select Case Request.QueryString("Error")
		Case "DeviceNotFound"
			strDeviceMessage = "<div Class=""Error"">Device not found</div>"
		Case "NoDevicesFound"
			strDeviceMessage = "<div Class=""Error"">No devices found</div>"
		Case "UserNotFound"
			strUserMessage = "<div Class=""Error"">User not found</div>"
		Case "NoUsersFound"
			strUserMessage = "<div Class=""Error"">No users found</div>"
		Case "NoEventsFound"
			strEventMessage = "<div Class=""Error"">No events found</div>"
	End Select

	Select Case Application("SiteName")

		Case "Lake George Inventory", "Lake George Inventory - Dev", "Inventory Demo"

			strAvailableTags = "Interactive Tags:" & " &#013 "
			strAvailableTags = strAvailableTags & "Missing - Speaks to User" & " &#013 "
			strAvailableTags = strAvailableTags & "Lost - Speaks to User" & " &#013 "
			strAvailableTags = strAvailableTags & "Disable - Make Unusable" & " &#013 "
			strAvailableTags = strAvailableTags & "Screenshot - Grabs Screnshots" & " &#013 "
			strAvailableTags = strAvailableTags & "Notify - Emails Techs" & " &#013 "
			strAvailableTags = strAvailableTags & "See Me - Notifies User" & " &#013 "

		Case "Schuylerville Inventory"

			strAvailableTags = "Spare - Shows on Chart" & " &#013 "
			strAvailableTags = strAvailableTags & "Replacement - Shows on Chart"

		Case Else

			strAvailableTags = "No Extra Information"

	End Select

	'Set up the variables needed for the site then load it
	SetupSite
	DisplaySite

End Sub%>

<%Sub DisplaySite

	Dim strSQL %>

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

			<%	If Not IsMobile And Not IsiPad Then%>
					$( document ).tooltip({track: true});
			<%	End If %>

			$( "#from" ).datepicker({
				changeMonth: true,
				changeYear: true,
				showOtherMonths: true,
				selectOtherMonths: true,
				onClose: function( selectedDate ) {
					$( "#to" ).datepicker( "option", "minDate", selectedDate );
				}
			});
			$( "#to" ).datepicker({
				changeMonth: true,
				changeYear: true,
				showOtherMonths: true,
				selectOtherMonths: true,
				onClose: function( selectedDate ) {
					$( "#from" ).datepicker( "option", "maxDate", selectedDate );
				}
			});

		<%	If Not objModels.EOF Then %>
				var possibleModels = [
			<%	Do Until objModels.EOF %>
					"<%=objModels(0)%>",
				<%	objModels.MoveNext
				Loop
				objModels.MoveFirst%>
			];
				$( "#DeviceModels" ).autocomplete({
					source: possibleModels
				});
				$( "#EventModels" ).autocomplete({
					source: possibleModels
				});
		<%	End If %>

		<%	If Not objFirstNames.EOF Then %>
				var possibleFirstNames = [
			<%	Do Until objFirstNames.EOF %>
					"<%=objFirstNames(0)%>",
				<%	objFirstNames.MoveNext
				Loop %>
			];
				$( "#FirstNames" ).autocomplete({
					source: possibleFirstNames
				});
		<%	End If %>

		<%	If Not objLastNames.EOF Then %>
				var possibleLastNames = [
			<%	Do Until objLastNames.EOF %>
					"<%=objLastNames(0)%>",
				<%	objLastNames.MoveNext
				Loop %>
			];
				$( "#LastNames" ).autocomplete({
					source: possibleLastNames
				});
		<%	End If %>

		});

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

		<div Class="<%=strColumns%>">

		<%
		WarningCard
		UserSearchCard
		DeviceSearchCard
		EventSearchCard
		TagsCard
		SavedSearchesCard
		%>

		</div>
		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
	</body>
	</html>

<%End Sub%>

<%Sub WarningCard

	If intDisabledUsersCount >= 1 Then %>
		<div Class="SiteErrorCard">
			<div Class="Center">
				<a href="users.asp?UserStatus=Disabled&WithDevice=Yes">
				Devices are Assigned to Disabled Users
				</a>
			</div>
		</div>
<%	End If

End Sub%>

<%Sub UserSearchCard%>

	<div class="Card NormalCard">
		<form method="POST" action="search.asp">
		<div class="CardTitle">Search for a User</div>
		<div>
			<div Class="CardColumn1">First name: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="FirstName" value="<%=strFirstName%>" id="FirstNames" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Last name: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="LastName" value="<%=strLastName%>" id="LastNames" />
			</div>
		</div>

		<div>
			<div Class="CardColumn1">Role: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="Role">
					<option value=""></option>
				<%	If strRole = "Adult" Then %>
						<option value="Adult" selected="selected">Adult</option>
				<%	Else %>
						<option value="Adult">Adult</option>
				<%	End If %>
				<%	If strRole = "Student" Then %>
						<option value="Student" selected="selected">Student</option>
				<%	Else %>
						<option value="Student">Student</option>
				<%	End If %>
				<%	Do Until objClassOf.EOF
						If strRole <> "" Then
							If IsNumeric(strRole) Then
								If Int(strRole) = Int(objClassOf(0)) Then
									strSelected = "selected=""selected"""
								Else
									strSelected = ""
								End If
							End If
						End If %>
						<option <%=strSelected%> value="<%=objClassOf(0)%>"><%=GetRole(objClassOf(0))%></option>
					<%	objClassOf.MoveNext
					Loop %>

				<%	Do Until objRoles.EOF
						If strRole <> "" Then
							If IsNumeric(strRole) Then
								If Int(strRole) = Int(objRoles(1)) Then
									strSelected = "selected=""selected"""
								Else
									strSelected = ""
								End If
							End If
						End If %>
						<option <%=strSelected%> value="<%=objRoles(1)%>"><%=objRoles(0)%></option>
					<%	objRoles.MoveNext
					Loop %>

				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1"><%=Application("HomeroomNameLong")%>: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="GuideRoom">
					<option value=""></option>
				<%	Do Until objGuideRooms.EOF
						If strGuideRoom = objGuideRooms(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objGuideRooms(0)%>"><%=objGuideRooms(0)%></option>
					<%	objGuideRooms.MoveNext
					Loop %>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Site: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="UserSite">
					<option value=""></option>
				<%	Do Until objSites.EOF
						If strUserSite = objSites(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objSites(0)%>"><%=objSites(0)%></option>
					<%	objSites.MoveNext
					Loop
					objSites.MoveFirst%>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">With Device: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="WithDevice">
					<option></option>
					<option value="Yes" <%=strWithDeviceYes%>>Yes</option>
					<option value="No" <%=strWithDeviceNo%>>No</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Owes Money: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="Owes">
					<option></option>
					<option value="Yes" <%=strOwesYes%>>Yes</option>
					<option value="No" <%=strOwesNo%>>No</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Loaned: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="UserLoaned">
					<option></option>
					<option value="Anything" <%=strLoanedAnything%>>Anything</option>
			<%	If Not objItems.EOF Then %>
				<%	Do Until objItems.EOF
						If strUserLoaned = objItems(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objItems(0)%>"><%=objItems(0)%></option>
					<%	objItems.MoveNext
					Loop
					objItems.MoveFirst%>
			<%	End If %>
				</select>
			</div>
		</div>
<%	If Application("ShowPasswords") Then %>
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
<%	End If %>
		<div>
			<div Class="CardColumn1">Internet: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="InternetAccess">
					<option></option>
			<%	If Not objInternetTypes.EOF Then %>
				<%	Do Until objInternetTypes.EOF
						If strInternetAccess = objInternetTypes(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objInternetTypes(0)%>"><%=objInternetTypes(0)%></option>
					<%	objInternetTypes.MoveNext
					Loop
					objInternetTypes.MoveFirst%>
			<%	End If %>
					<option value="Unknown">Unknown</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Status: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="UserStatus">
					<option ></option>
					<option value="Enabled" <%=strUserStatusEnabled%>>Enabled</option>
					<option value="Disabled" <%=strUserStatusDisabled%>>Disabled</option>
					<option value="All" <%=strUserStatusAll%>>All</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Description: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="Description" value="<%=strDescription%>" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">User Notes: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="UserNotes" value="<%=strUserNotes%>" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">View: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="UserView">
					<option ></option>
					<option value="Card" <%=strUserViewCard%>>Cards</option>
					<option value="Table" <%=strUserViewTable%>>Table</option>
				</select>
			</div>
		</div>
		<div>
			<div class="Button"><input type="submit" value="Lookup User" name="Submit" /></div>
		</div>
	<%	If strUserMessage <> "" Then %>
			<div>
				<%=strUserMessage%>
			</div>
	<%	End If %>
		</form>
	</div>

<%End Sub%>

<%Sub DeviceSearchCard%>

	<div class="Card NormalCard">
		<form method="POST" action="search.asp">
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
			<div Class="CardColumn1">Device Type: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="DeviceType">
					<option value=""></option>
				<%	Do Until objDeviceTypes.EOF
						If strDeviceType = objDeviceTypes(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objDeviceTypes(0)%>"><%=objDeviceTypes(0)%></option>
					<%	objDeviceTypes.MoveNext
					Loop %>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Make: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="Make">
					<option value=""></option>
				<%	Do Until objMakes.EOF
						If strMake = objMakes(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objMakes(0)%>"><%=objMakes(0)%></option>
					<%	objMakes.MoveNext
					Loop %>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Model: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="Model" value="<%=strModel%>" id="DeviceModels" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Tags: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="Tags" value="<%=strTags%>"/>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Room: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="Room">
					<option value=""></option>
				<%	Do Until objRooms.EOF
						If strRoom = objRooms(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objRooms(0)%>"><%=objRooms(0)%></option>
					<%	objRooms.MoveNext
					Loop %>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Site: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="DeviceSite">
					<option value=""></option>
				<%	Do Until objSites.EOF
						If strDeviceSite = objSites(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objSites(0)%>"><%=objSites(0)%></option>
					<%	objSites.MoveNext
					Loop
					objSites.MoveFirst%>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Device Year: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="DeviceYear">
					<option value=""></option>
				<%	For intIndex = 1 to intYears
						strSelected = ""
						If intDeviceYear <> "" Then
							If Int(intDeviceYear) = Int(intIndex) Then
								strSelected = "selected=""selected"""
							Else
								strSelected = ""
							End If
						End If %>
						<option <%=strSelected%> value="<%=intIndex%>"><%="Year " & intIndex%></option>
				<%	Next %>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Assigned: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="Assigned">
					<option></option>
					<option value="Yes" <%=strAssignedYes%>>Yes</option>
					<option value="No" <%=strAssignedNo%>>No</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Status: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="DeviceStatus">
					<option ></option>
					<option value="Enabled" <%=strDeviceStatusEnabled%>>Enabled</option>
					<option value="Disabled" <%=strDeviceStatusDisabled%>>Disabled</option>
					<option value="All" <%=strDeviceStatusAll%>>All</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">IP Address: </div>
			<div Class="CardColumn2">
				<input class="Card" type="text" name="IPAddress" value="<%=strIPAddress%>" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Device Notes: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="DeviceNotes" value="<%=strDeviceNotes%>" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">View: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="DeviceView">
					<option ></option>
					<option value="Card" <%=strDeviceViewCard%>>Cards</option>
					<option value="Table" <%=strDeviceViewTable%>>Table</option>
				</select>
			</div>
		</div>
		<div>
			<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		</div>
	<%	If strDeviceMessage <> "" Then %>
		<div>
			<%=strDeviceMessage%>
		</div>
	<%	End If %>
		</form>
	</div>

<%End Sub%>

<%Sub EventSearchCard%>

	<div class="Card NormalCard">
		<form method="POST" action="search.asp">
		<div class="CardTitle">Search for an Event</div>
		<div>
			<div>
				<div Class="CardColumn1">Number: </div>
				<div Class="CardColumn2">
					<input class="Card InputWidthSmall" type="text" name="EventNumber" value="<%=intEventNumber%>" />
				</div>
			</div>

			<div Class="CardColumn1">Type: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="EventType">
					<option value=""></option>
				<%	Do Until objEventTypes.EOF
						If strEventType = objEventTypes(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objEventTypes(0)%>"><%=objEventTypes(0)%></option>
					<%	objEventTypes.MoveNext
					Loop
					objEventTypes.MoveFirst%>
					<option value="Decommission Device">Decommission Device</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Category: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="Category">
					<option value=""></option>
				<%	Do Until objCategories.EOF
						If strCategory = objCategories(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objCategories(0)%>"><%=objCategories(0)%></option>
					<%	objCategories.MoveNext
					Loop %>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Model: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="EventModel" value="<%=strEventModel%>" id="EventModels" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Site: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="EventSite">
					<option value=""></option>
				<%	Do Until objSites.EOF
						If strEventSite = objSites(0) Then
							strSelected = "selected=""selected"""
						Else
							strSelected = ""
						End If %>
						<option <%=strSelected%> value="<%=objSites(0)%>"><%=objSites(0)%></option>
					<%	objSites.MoveNext
					Loop
					objSites.MoveFirst%>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Device Year: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="DeviceYear">
					<option value=""></option>
				<%	For intIndex = 1 to intYears
						strSelected = ""
						If intDeviceYear <> "" Then
							If Int(intDeviceYear) = Int(intIndex) Then
								strSelected = "selected=""selected"""
							Else
								strSelected = ""
							End If
						End If %>
						<option <%=strSelected%> value="<%=intIndex%>"><%="Year " & intIndex%></option>
				<%	Next %>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Warranty: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="Warranty">
					<option></option>
					<option value="Yes" <%=strWarrantyYes%>>Yes</option>
					<option value="No" <%=strWarrantyNo%>>No</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Complete: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="Complete">
					<option></option>
					<option value="Yes" <%=strCompleteYes%>>Yes</option>
					<option value="No" <%=strCompleteNo%>>No</option>
					<option value="All" <%=strCompleteAll%>>All</option>
				</select>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Start Date: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthSmall" type="text" name="StartDate" value="<%=datStartDate%>" id="from" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">End Date: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthSmall" type="text" name="EndDate" value="<%=datEndDate%>" id="to" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Event Notes: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="Notes" value="<%=strNotes%>"/>
			</div>
		</div>
		<div>
			<div Class="CardColumn1">View: </div>
			<div Class="CardColumn2">
				<select Class="Card" Name="EventView">
					<option ></option>
					<option value="Card" <%=strEventViewCard%>>Cards</option>
					<option value="Table" <%=strEventViewTable%>>Table</option>
				</select>
			</div>
		</div>
		<div>
			<div class="Button"><input type="submit" value="Lookup Event" name="Submit" /></div>
		</div>
	<%	If strEventMessage <> "" Then %>
			<div>
				<%=strEventMessage%>
			</div>
	<%	End If %>
		</form>
	</div>

<%End Sub%>

<%Sub TagsCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">
			Available Tags
			<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strAvailableTags%>"  />&nbsp;</div>
		</div>
		<div>
			<%=strTagList%>
		</div>
	</div>
<%End Sub%>

<%Sub StudentsPerGradeCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">Students Per Grade</div>
		<div id="openPerTech"></div>
	</div>
<%End Sub%>

<%Sub SavedSearchesCard%>
	<div class="Card NormalCard">
	<form method="POST" action="search.asp">
		<div class="CardTitle">Saved Searches</div>
		<div Class="CardMerged">Search:
			<select Class="Card" name="SearchName">
				<option value=""></option>
			<%	Do Until objSavedSearches.EOF
					If Int(intSavedSearch) = Int(objSavedSearches(0)) Then
						strSelected = "selected=""selected"""
					Else
						strSelected = ""
					End If %>
					<option <%=strSelected%> value="<%=objSavedSearches(0)%>"><%=objSavedSearches(1)%></option>
				<%	objSavedSearches.MoveNext
				Loop %>
			</select>
		</div>
		<div>
			<div class="Button"><input type="submit" value="Search" name="Submit" /></div>
			<div class="Button"><input type="submit" value="Load" name="Submit" /></div>
		</div>
	</form>
	</div>
<%End Sub%>

<%Sub ShowCharts

	strSQL = "SELECT CourseName, Count(Students.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Students" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (HomeRoom Is Not Null And HomeRoom  <>'')" & vbCRLF
	strSQL = strSQL & "GROUP BY CourseName" & vbCRLF
	strSQL = strSQL & "ORDER BY CourseName"
	Set objStudentsPerGrade = Application("Connection").Execute(strSQL) %>

		<script type="text/javascript" src="https://www.google.com/jsapi"></script>
		<script type="text/javascript">
		google.load("visualization", "1", {packages:["corechart"]});
		google.setOnLoadCallback(drawStudentsPerGrade);

		function drawStudentsPerGrade() {
			var data = google.visualization.arrayToDataTable([
				['Grade', 'Students'],
		<%	intHighestValue = 0
			If Not objStudentsPerGrade.EOF Then
				Do Until objStudentsPerGrade.EOF
					If objStudentsPerGrade(1) > intHighestValue Then
						intHighestValue = objStudentsPerGrade(1)
					End If
				%>
					['<%=GetGrade(objStudentsPerGrade(0))%>', <%=objStudentsPerGrade(1)%>],
			<%	objStudentsPerGrade.MoveNext
				Loop
			End If%>

			]);

			var options = {
				title: 'Total Number of Students Per Grade',
				bar: {groupWidth: "90%"},
				vAxis: {viewWindow: {max : <%=intHighestValue+(intHighestValue*.2)%>},minValue: 0},
				titleTextStyle:{fontSize: 14},
				titlePosition: 'out',
				legend:{position: 'none'},
				chartArea:{width:'90%', height:'80%'},
			};

			var chart = new google.visualization.ColumnChart(document.getElementById('openPerTech'));
			chart.draw(data, options);
		}
		</script>

<%End Sub%>

<%Sub LookupDevice

	Dim strSQLWhere, strSQL, objDeviceLookup, intDeviceCount, strURL, arrTags, intIndex
	Dim strAssigned, intTagCount, intDeviceTag, strDeviceStatus, intUserCount, objUserLookup
	Dim strInternalIP, strExternalIP

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

	If Not Request.Form("SmartBox") = "" Then
		intTag = Request.Form("SmartBox")
		If Not Application("UseLeadingZeros") Then
			If IsNumeric(intTag) Then
				intTag = Int(intTag)
			Else
				If Left(intTag,4) = "TECH" Then
					intTag = Replace(intTag,"TECH","")
				End If
				If IsNumeric(intTag) Then
					intTag = Int(intTag)
				Else
					LookupUser
					intTag = ""
					Exit Sub
				End If
			End If
		End If
	End If
	
	If intTag = 0 Then
		intTag = ""
	End If

	If IsNumeric(Request.Form("BOCESTag")) Then
		intBOCESTag = Int(Request.Form("BOCESTag"))
	Else
		intBOCESTag = Request.Form("BOCESTag")
	End If

	Select Case Request.Form("Assigned")
		Case "Yes"
			strAssigned = "Yes"
			strAssignedYes = "selected=""selected"""
		Case "No"
			strAssigned = "No"
			strAssignedNo = "selected=""selected"""
		Case Else
			strAssigned = ""
	End Select

	Select Case Request.Form("DeviceStatus")
		Case "Enabled"
			strDeviceStatus = "Enabled"
			strDeviceStatusEnabled = "selected=""selected"""
		Case "Disabled"
			strDeviceStatus = "Disabled"
			strDeviceStatusDisabled = "selected=""selected"""
		Case "All"
			strDeviceStatus = "All"
			strDeviceStatusAll = "selected=""selected"""
		Case Else
			If intTag = "" Then
				strDeviceStatus = ""
			Else
				strDeviceStatus = "All"
			End If
	End Select

	Select Case Request.Form("DeviceView")
		Case "Card"
			strDeviceView = "Card"
			strDeviceViewCard = "selected=""selected"""
		Case "Table"
			strDeviceView = "Table"
			strDeviceViewTable = "selected=""selected"""
		Case Else
			strDeviceView = ""
	End Select

	strSerial = Request.Form("Serial")
	strDeviceType = Request.Form("DeviceType")
	strMake = Request.Form("Make")
	strModel = Request.Form("Model")
	strTags = Request.Form("Tags")
	strRoom = Request.Form("Room")
	strDeviceSite = Request.Form("DeviceSite")
	intDeviceYear = Request.Form("DeviceYear")
	strIPAddress = Request.Form("IPAddress")
	strDeviceNotes = Request.Form("DeviceNotes")
	intDeviceCount = 0
	
	'Determine if the IP address is internal or external.
	Select Case Left(strIPAddress,3)
		Case "10.", "172", "192"
			strInternalIP = strIPAddress
		Case Else 
			strExternalIP = strIPAddress
	End Select	

	If intTag = "" And intBOCESTag = "" And strSerial = "" And strMake = "" And strDeviceSite = "" _
		And intDeviceYear = "" And strTags = "" And strRoom = "" And strModel = "" And strAssigned = "" _
		And strDeviceStatus = "" And strDeviceNotes = "" And strDeviceType = "" And strIPAddress = "" Then
		strDeviceMessage = "<div Class=""Error"">Device not found</div>"

	Else

		strSQLWhere = "WHERE "

		If intTag <> "" Then
			strSQLWhere = strSQLWhere & "Devices.LGTag='" & Replace(intTag,"'","''") & "' AND "
		End If

		If intBOCESTag <> "" Then
			strSQLWhere = strSQLWhere & "Devices.BOCESTag='" & Replace(intBOCESTag,"'","''") & "' AND "
		End If

		If strSerial <> "" Then
			strSQLWhere = strSQLWhere & "Devices.SerialNumber='" & Replace(strSerial,"'","''") & "' AND "
		End If

		If strDeviceType <> "" Then
			strSQLWhere = strSQLWhere & "Devices.DeviceType='" & Replace(strDeviceType,"'","''") & "' AND "
			strURL = strURL & "&DeviceType=" & Replace(strDeviceType," ","%20")
		End If

		If strMake <> "" Then
			strSQLWhere = strSQLWhere & "Devices.Manufacturer='" & Replace(strMake,"'","''") & "' AND "
			strURL = strURL & "&Make=" & Replace(strMake," ","%20")
		End If

		If strModel <> "" Then
			strSQLWhere = strSQLWhere & "Devices.Model Like '%" & Replace(strModel,"'","''") & "%' AND "
			strURL = strURL & "&Model=" & Replace(strModel," ","%20")
		End If

		If strRoom <> "" Then
			strSQLWhere = strSQLWhere & "Devices.Room='" & Replace(strRoom,"'","''") & "' AND "
			strURL = strURL & "&Room=" & Replace(strRoom," ","%20")
		End If

		If strInternalIP <> "" Then
			strSQLWhere = strSQLWhere & "Devices.InternalIP Like '%" & Replace(strInternalIP,"'","''") & "%' AND "
			strURL = strURL & "&InternalIP=" & Replace(strInternalIP," ","%20")
		End If
		
		If strExternalIP <> "" Then
			strSQLWhere = strSQLWhere & "Devices.ExternalIP Like '%" & Replace(strExternalIP,"'","''") & "%' AND "
			strURL = strURL & "&ExternalIP=" & Replace(strExternalIP," ","%20")
		End If

		If strDeviceNotes <> "" Then
			strSQLWhere = strSQLWhere & "Devices.Notes Like '%" & Replace(strDeviceNotes,"'","''") & "%' AND "
			strURL = strURL & "&DeviceNotes=" & Replace(strDeviceNotes," ","%20")
		End If

		If strDeviceSite <> "" Then
			strSQLWhere = strSQLWhere & "Devices.Site='" & Replace(strDeviceSite,"'","''") & "' AND "
			strURL = strURL & "&DeviceSite=" & Replace(strDeviceSite," ","%20")
		End If

		If intDeviceYear <> "" Then
			strSQLWhere = strSQLWhere & _
			"Devices.DatePurchased>=#" & DateAdd("yyyy",intDeviceYear * -1,Date) & "# AND " & _
			"Devices.DatePurchased<=#" & DateAdd("yyyy",(intDeviceYear -1) * -1,Date) & "# AND "
			strURL = strURL & "&Year=" & intDeviceYear
		End If

		Select Case strAssigned
			Case "Yes"
				strSQLWhere = strSQLWhere & "Assigned=True AND "
				strURL = strURL & "&Assigned=Yes"
			Case "No"
				strSQLWhere = strSQLWhere & "Assigned=False AND "
				strURL = strURL & "&Assigned=No"
		End Select

		Select Case strDeviceStatus
			Case "Enabled"
				strSQLWhere = strSQLWhere & "Devices.Active=True AND "
				strURL = strURL & "&DeviceStatus=Enabled"
			Case "Disabled"
				strSQLWhere = strSQLWhere & "Devices.Active=False AND "
				strURL = strURL & "&DeviceStatus=Disabled"
			Case "All"
				strURL = strURL & "&DeviceStatus=All"
			Case Else
				strSQLWhere = strSQLWhere & "Devices.Active=True AND "
		End Select

		strSQLWhere = strSQLWhere & "Devices.Deleted=False AND "

		If strSQLWhere <> "WHERE " Then
			strSQLWhere = Left(strSQLWhere,Len(strSQLWhere) - 5)
		Else
			strSQLWhere = ""
		End If

		If strTags = "" Then
			intTagCount = 0
			strSQL = "SELECT ID, LGTag FROM Devices " & strSQLWhere

		Else
			strSQL = "SELECT Devices.LGTag,Tags.Tag FROM Tags INNER JOIN Devices ON Tags.LGTag = Devices.LGTag "

			arrTags = Split(strTags,",")
			intTagCount = UBound(arrTags) + 1

			strSQL = "SELECT COUNT(Tags.Tag) AS TagCount, Devices.LGTag FROM(" & strSQL & strSQLWhere & ") WHERE "

			For intIndex = 0 to UBound(arrTags)
				strSQL = strSQL & "Tags.Tag='" & Trim(Replace(arrTags(intIndex),"'","''")) & "' OR "
			Next
			strSQL = Left(strSQL,Len(strSQL) - 4) & " GROUP BY Devices.LGTag"
			strURL = strURL & "&Tags=" & Replace(strTags," ","%20")

		End If

		Set objDeviceLookup = Application("Connection").Execute(strSQL)

		If Not objDeviceLookup.EOF Then
			Do Until objDeviceLookup.EOF
				If intTagCount > 0 Then
					If objDeviceLookup(0) = intTagCount Then
						intDeviceTag = objDeviceLookup(1)
						intDeviceCount = intDeviceCount + 1
					End If
				Else
					intDeviceTag = objDeviceLookup(1)
					intDeviceCount = intDeviceCount + 1
				End If
				objDeviceLookup.MoveNext
			Loop
			objDeviceLookup.MoveFirst
		End If

		If strDeviceView <> "" Then
			strURL = strURL & "&View=" & strDeviceView
		End If

		If strURL <> "" Then
			strURL = Right(strURL, Len(strURL) - 1)
		End If

		'Look and see if any users are in the room, if so go to the
		intUserCount = 0
		strSQL = "SELECT ID, UserName FROM People WHERE Active=True AND "
		If strRoom <> "" Then
			strSQL = strSQL & "RoomNumber='" & Replace(strRoom,"'","''") & "'"

			If strDeviceSite <> "" Then
				strSQL = strSQL & " AND Site='" & Replace(strDeviceSite,"'","''") & "'"
			End If
		Else
			strSQL = strSQL & "RoomNumber='NO ROOM NUMBER GIVEN'"
		End If
		Set objUserLookup = Application("Connection").Execute(strSQL)

		If Not objUserLookup.EOF Then
			Do Until objUserLookup.EOF
				intUserCount = intUserCount + 1
				objUserLookup.MoveNext
			Loop
		End If

		Select Case intDeviceCount
			Case 0
				If intUserCount > 0 Then
					Response.Redirect("devices.asp?" & strURL)
				End If
				strDeviceMessage = "<div Class=""Error"">Device not found</div>"
			Case 1
				If intUserCount > 0 Then
					Response.Redirect("devices.asp?" & strURL)
				End If
				Response.Redirect("device.asp?Tag=" & intDeviceTag)
			Case Else
				Response.Redirect("devices.asp?" & strURL)
		End Select
	End If

End Sub%>

<%Sub LookupUser

	Dim  strSQL, objUserLookup, intUserCount, strURL, strSQLWhere, strUserStatus, strMissing, strLoaned

	Select Case Request.Form("WithDevice")
		Case "Yes"
			strWithDevice = "Yes"
			strWithDeviceYes = "selected=""selected"""
		Case "No"
			strWithDevice = "No"
			strWithDeviceNo = "selected=""selected"""
		Case Else
			strWithDevice = ""
	End Select

	Select Case Request.Form("UserLoaned")
		Case "Anything"
			strLoaned = "Anything"
			strLoanedAnything = "selected=""selected"""
	End Select

	Select Case Request.Form("Owes")
		Case "Yes"
			strOwesYes = "selected=""selected"""
			strOwes = "True"
		Case "No"
			strOwesNo = "selected=""selected"""
			strOwes = "False"
	End Select

	Select Case Request.Form("AUP")
		Case "Yes"
			strAUPYes = "selected=""selected"""
			strAUP = "True"
		Case "No"
			strAUPNo = "selected=""selected"""
			strAUP = "False"
	End Select

	Select Case Request.Form("UserStatus")
		Case "Enabled"
			strUserStatus = "Enabled"
			strUserStatusEnabled = "selected=""selected"""
		Case "Disabled"
			strUserStatus = "Disabled"
			strUserStatusDisabled = "selected=""selected"""
		Case "All"
			strUserStatus = "All"
			strUserStatusAll = "selected=""selected"""
		Case Else
			strUserStatus = ""
	End Select

	Select Case Request.Form("UserView")
		Case "Card"
			strUserView = "Card"
			strUserViewCard = "selected=""selected"""
		Case "Table"
			strUserView = "Table"
			strUserViewTable = "selected=""selected"""
		Case Else
			strUserView = ""
	End Select

	strFirstName = Request.Form("FirstName")
	strLastName = Request.Form("LastName")
	strGuideRoom = Request.Form("GuideRoom")
	strRole = Request.Form("Role")
	strUserSite = Request.Form("UserSite")
	strUserLoaned = Request.Form("UserLoaned")
	strDescription = Request.Form("Description")
	strUserNotes = Request.Form("UserNotes")
	strInternetAccess = Request.Form("InternetAccess")
	intUserCount = 0

	If Request.Form("SmartBox") <> "" Then
		strLastName = Request.Form("SmartBox")
	End If

	If strFirstName = "" And strLastName = "" And strGuideRoom = "" And strRole = "" And strOwes = "" _
		And strModel = "" And strUserSite = "" And strWithDevice = "" And strUserStatus = "" _
		And strMissing = "" And strUserLoaned = "" And strAUP = "" And strUserNotes = "" And strDescription = "" _
		And strInternetAccess = "" Then
		strUserMessage = "<div Class=""Error"">User not found</div>"
	Else

		strSQLWhere = "WHERE "
		If strFirstName <> "" Then
			strSQLWhere = strSQLWhere & "People.FirstName Like '%" & Replace(strFirstName,"'","''") & "%' AND "
			strURL = strURL & "&FirstName=" & strFirstName
		End If
		If strLastName <> "" Then
			strSQLWhere = strSQLWhere & "People.LastName Like '%" & Replace(strLastName,"'","''") & "%' AND "
			strURL = strURL & "&LastName=" & strLastName
		End If
		If strGuideRoom <> "" Then
			strSQLWhere = strSQLWhere & "People.HomeRoom Like '%" & Replace(strGuideRoom,"'","''") & "%' AND "
			strURL = strURL & "&GuideRoom=" & strGuideRoom
		End If
		If strUserSite <> "" Then
			strSQLWhere = strSQLWhere & "People.Site='" & Replace(strUserSite,"'","''") & "' AND "
			strURL = strURL & "&UserSite=" & Replace(strUserSite," ","%20")
		End If
		If strAUP <> "" Then
			strSQLWhere = strSQLWhere & "AUP=" & strAUP & " AND "
			strURL = strURL & "&AUP=" & strAUP
		End If
		If strUserNotes <> "" Then
			strSQLWhere = strSQLWhere & "People.Notes Like '%" & Replace(strUserNotes,"'","''") & "%' AND "
			strURL = strURL & "&UserNotes=" & strUserNotes
		End If
		If strDescription <> "" Then
			strSQLWhere = strSQLWhere & "People.Description Like '%" & Replace(strDescription,"'","''") & "%' AND "
			strURL = strURL & "&Description=" & strDescription
		End If
		If strOwes <> "" Then
			strSQLWhere = strSQLWhere & "Warning=True AND "
			strURL = strURL & "&Owes=" & strOwes
		End If
		If strInternetAccess <> "" Then
			If strInternetAccess = "Unknown" Then
				strSQLWhere = strSQLWhere & "(InternetAccess Is Null OR InternetAccess='') AND "
			Else
				strSQLWhere = strSQLWhere & "InternetAccess='" & Replace(strInternetAccess,"'","''") & "' AND "
			End If
			strURL = strURL & "&InternetAccess=" & strInternetAccess
		End If
		If strRole <> "" Then
			Select Case strRole
				Case "Adult", "Student"
					strSQLWhere = strSQLWhere & "People.Role='" & Replace(Replace(strRole,"'","''"),"Adult","Teacher") & "' AND "
				Case Else
					strSQLWhere = strSQLWhere & "People.ClassOf=" & Replace(strRole,"'","''") & " AND "
			End Select
			strURL = strURL & "&Role=" & strRole
		End If

		Select Case strUserStatus
			Case "Enabled"
				strSQLWhere = strSQLWhere & "People.Active=True AND "
				strURL = strURL & "&UserStatus=Enabled"
			Case "Disabled"
				strSQLWhere = strSQLWhere & "People.Active=False AND "
				strURL = strURL & "&UserStatus=Disabled"
			Case "All"
				strURL = strURL & "&UserStatus=All"
			Case Else
				strSQLWhere = strSQLWhere & "People.Active=True AND "
		End Select

		Select Case strWithDevice
			Case "Yes"
				strSQLWhere = strSQLWhere & "People.HasDevice=True AND "
				strURL = strURL & "&WithDevice=Yes"
			Case "No"
				strSQLWhere = strSQLWhere & "People.HasDevice=False AND "
				strURL = strURL & "&WithDevice=No"
		End Select

		strSQLWhere = strSQLWhere & "People.Deleted=False AND "

		If strSQLWhere <> "" Then
			strSQLWhere = Left(strSQLWhere,Len(strSQLWhere) - 5)
		End If

		If strMissing <> "" Then
			Select Case strMissing
				Case "Anything"
					strSQL = "SELECT People.ID, UserName" & vbCRLF
					strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
					strSQL = strSQL & strSQLWhere & " AND (CaseReturned=False OR AdapterReturned=False) AND Assignments.Active=False"
					strURL = strURL & "&Missing=Anything"
				Case "Case"
					strSQL = "SELECT People.ID, UserName" & vbCRLF
					strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
					strSQL = strSQL & strSQLWhere & " AND CaseReturned=False AND Assignments.Active=False"
					strURL = strURL & "&Missing=Case"
				Case "Power Supply"
					strSQL = "SELECT People.ID, UserName" & vbCRLF
					strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
					strSQL = strSQL & strSQLWhere & " AND AdapterReturned=False AND Assignments.Active=False"
					strURL = strURL & "&Missing=Power%20Supply"
			End Select
		ElseIf strUserLoaned <> "" Then
			Select Case strUserLoaned
				Case "Anything"
					strSQL = "SELECT People.ID, UserName" & vbCRLF
					strSQL = strSQL & "FROM People INNER JOIN Loaned ON People.ID = Loaned.AssignedTo" & vbCRLF
					strSQL = strSQL & strSQLWhere & " AND (Loaned.Returned=False)"
					strURL = strURL & "&LoanedOut=Anything"
				Case Else
					strSQL = "SELECT People.ID, UserName" & vbCRLF
					strSQL = strSQL & "FROM People INNER JOIN Loaned ON People.ID = Loaned.AssignedTo" & vbCRLF
					strSQL = strSQL & strSQLWhere & " AND (Loaned.Returned=False) AND Loaned.Item='" & Replace(strUserLoaned,"'","''") & "'"
					strURL = strURL & "&LoanedOut=" & Replace(strUserLoaned," ","%20")
			End Select
		Else
			strSQL = "SELECT ID, UserName FROM People " & strSQLWhere
		End If

		Set objUserLookup = Application("Connection").Execute(strSQL)

		If Not objUserLookup.EOF Then
			Do Until objUserLookup.EOF
				intUserCount = intUserCount + 1
				objUserLookup.MoveNext
			Loop
			objUserLookup.MoveFirst
		End If

		If strUserView <> "" Then
			strURL = strURL & "&View=" & strUserView
		End If

		If strURL <> "" Then
			strURL = Right(strURL, Len(strURL) - 1)
		End If

		Select Case intUserCount
			Case 0
				strUserMessage = "<div Class=""Error"">User not found</div>"
			Case 1
				Response.Redirect("user.asp?UserName=" & objUserLookup(1))
			Case Else
				Response.Redirect("users.asp?" & strURL)
		End Select

	End If

End Sub%>

<%Sub LookupEvent

	Dim strURL, strSQL, strSQLWhere, objEventLookup, intEventCount

	intEventNumber = Request.Form("EventNumber")
	strEventType = Request.Form("EventType")
	strCategory = Request.Form("Category")
	strEventModel = Request.Form("EventModel")
	strEventSite = Request.Form("EventSite")
	intDeviceYear = Request.Form("DeviceYear")
	strNotes = Request.Form("Notes")
	datStartDate = Request.Form("StartDate")
	datEndDate = Request.Form("EndDate")

	Select Case Request.Form("Warranty")
		Case "Yes"
			strWarranty = "Yes"
			strWarrantyYes = "selected=""selected"""
		Case "No"
			strWarranty = "No"
			strWarrantyNo = "selected=""selected"""
		Case Else
			strWarranty = ""
	End Select

	Select Case Request.Form("Complete")
		Case "Yes"
			strComplete = "Yes"
			strCompleteYes = "selected=""selected"""
		Case "No"
			strComplete = "No"
			strCompleteNo = "selected=""selected"""
		Case "All"
			strComplete = "All"
			strCompleteAll = "selected=""selected"""
		Case Else
			strComplete = ""
	End Select

	Select Case Request.Form("EventView")
		Case "Card"
			strEventView = "Card"
			strEventViewCard = "selected=""selected"""
		Case "Table"
			strEventView = "Table"
			strEventViewTable = "selected=""selected"""
		Case Else
			strEventView = ""
	End Select

	If strEventType = "Decommission Device" Then
		strComplete = "Yes"
	End If

	intEventCount = 0

	strSQLWhere = "WHERE "
	If intEventNumber <> "" Then
		If IsNumeric(intEventNumber) Then
			strSQLWhere = strSQLWhere & "Events.ID=" & intEventNumber & " AND "
			strURL = strURL & "&EventNumber=" & intEventNumber
		End If
	End If
	If strEventType <> "" Then
		strSQLWhere = strSQLWhere & "Events.Type='" & Replace(strEventType,"'","''") & "' AND "
		strURL = strURL & "&EventType=" & strEventType
	End If
	If strCategory <> "" Then
		strSQLWhere = strSQLWhere & "Events.Category='" & Replace(strCategory,"'","''") & "' AND "
		strURL = strURL & "&Category=" & strCategory
	End If
	If strNotes <> "" Then
		strSQLWhere = strSQLWhere & "Events.Notes Like '%" & Replace(strNotes,"'","''") & "%' AND "
		strURL = strURL & "&EventNotes=" & strNotes
	End If
	If strEventModel <> "" Then
		strSQLWhere = strSQLWhere & "Events.Model Like '%" & Replace(strEventModel,"'","''") & "%' AND "
		strURL = strURL & "&EventModel=" & Replace(strEventModel," ","%20")
	End If
	If strEventSite <> "" Then
		strSQLWhere = strSQLWhere & "Events.Site='" & Replace(strEventSite,"'","''") & "' AND "
		strURL = strURL & "&EventSite=" & Replace(strEventSite," ","%20")
	End If
	If intDeviceYear <> "" Then
		strSQLWhere = strSQLWhere & _
		"Devices.DatePurchased>=#" & DateAdd("yyyy",intDeviceYear * -1,Date) & "# AND " & _
		"Devices.DatePurchased<=#" & DateAdd("yyyy",(intDeviceYear -1) * -1,Date) & "# AND "
		strURL = strURL & "&Year=" & intDeviceYear
	End If

	Select Case strWarranty
		Case "Yes"
			strSQLWhere = strSQLWhere & "Events.Warranty=True AND "
			strURL = strURL & "&Warranty=Yes"
		Case "No"
			strSQLWhere = strSQLWhere & "Events.Warranty=False AND "
			strURL = strURL & "&Warranty=No"
	End Select

	Select Case strComplete
		Case "Yes"
			strSQLWhere = strSQLWhere & "Events.Resolved=True AND "
			strURL = strURL & "&Complete=Yes"
		Case "No"
			strSQLWhere = strSQLWhere & "Events.Resolved=False AND "
			strURL = strURL & "&Complete=No"
		Case "All"
			strURL = strURL & "&Complete=All"
		Case Else
			If intEventNumber = "" Then
				strSQLWhere = strSQLWhere & "Events.Resolved=False AND "
			End If
	End Select

	If datStartDate <> "" Then
		strSQLWhere = strSQLWhere & "Events.EventDate>=#" & datStartDate & "# AND "
		strURL = strURL & "&StartDate=" & datStartDate
	End If

	If datEndDate <> "" Then
		strSQLWhere = strSQLWhere & "Events.EventDate<=#" & datEndDate & "# AND "
		strURL = strURL & "&EndDate=" & datEndDate
	End If

	If strSQLWhere <> "WHERE " Then
		strSQLWhere = Left(strSQLWhere,Len(strSQLWhere) - 5)
	Else
		strSQLWhere = ""
	End If

	If intDeviceYear = "" Then
		strSQL = "SELECT ID, LGTag FROM Events " & strSQLWhere
	Else
		strSQL = "SELECT Events.ID, Events.LGTag FROM Devices INNER JOIN Events ON Devices.LGTag = Events.LGTag " & strSQLWhere
	End If

	Set objEventLookup = Application("Connection").Execute(strSQL)

	If Not objEventLookup.EOF Then
		Do Until objEventLookup.EOF
			intEventCount = intEventCount + 1
			objEventLookup.MoveNext
		Loop
		objEventLookup.MoveFirst
	End If

	If strEventView <> "" Then
		strURL = strURL & "&View=" & strEventView
	End If

	If strURL <> "" Then
		strURL = "?" & Right(strURL, Len(strURL) - 1)
	End If

	Select Case intEventCount
		Case 0
			strEventMessage = "<div Class=""Error"">No events found</div>"
		Case Else
			Response.Redirect("events.asp" & strURL)
	End Select

End Sub%>

<%Sub Search

	Dim strSQL, intSearchID, objSearch

	intSearchID = Request.Form("SearchName")

	If Not intSearchID = "" Then
		strSQL = "SELECT Page,QueryString FROM SavedSearches WHERE ID=" & intSearchID
		Set objSearch = Application("Connection").Execute(strSQL)

		If Not objSearch.EOF Then
			Response.Redirect(objSearch(0) & "?" & objSearch(1)) & "&Source=search.asp"
		End If
	End If

End Sub%>

<%Sub LoadSearch

	Dim strSQL, intSearchID, objSearch, strPage, strQueryString

	intSearchID = Request.Form("SearchName")

	If Not intSearchID = "" Then
		strSQL = "SELECT Page,QueryString FROM SavedSearches WHERE ID=" & intSearchID
		Set objSearch = Application("Connection").Execute(strSQL)

		If Not objSearch.EOF Then

			strPage = objSearch(0)
			strQueryString = objSearch(1)

			Select Case LCase(strPage)
				Case "users.asp"
					strQueryString = Replace(strQueryString,"&View=","&UserView=")
				Case "devices.asp"
					strQueryString = Replace(strQueryString,"&View=","&DeviceView=")
				Case "events.asp"
					strQueryString = Replace(strQueryString,"&View=","&EventView=")
			End Select

			Response.Redirect("search.asp?" & strQueryString & "&Search=" & intSearchID)
		End If
	End If

End Sub%>

<%Sub GetVariablesFromURL

	intSavedSearch = Request.QueryString("Search")

	strFirstName = Request.QueryString("FirstName")
	strLastName = Request.QueryString("LastName")
	strRole = Request.QueryString("Role")
	strGuideRoom = Request.QueryString("GuideRoom")
	strUserSite = Request.QueryString("UserSite")
	strWithDevice = Request.QueryString("WithDevice")
	strUserLoaned = Request.QueryString("LoanedOut")
	strUserNotes = Request.QueryString("UserNotes")
	strDescription = Request.QueryString("Description")
	strInternetAccess = Request.QueryString("InternetAccess")

	intBOCESTag = Request.QueryString("BOCESTag")
	strSerial = Request.QueryString("Serial")
	strDeviceType = Request.QueryString("DeviceType")
	strMake = Request.QueryString("Make")
	strModel = Request.QueryString("Model")
	strTags = Request.QueryString("Tags")
	strRoom = Request.QueryString("Room")
	strDeviceSite = Request.QueryString("DeviceSite")
	intDeviceYear = Request.QueryString("Year")
	strDeviceNotes = Request.QueryString("DeviceNotes")

	intEventNumber = Request.QueryString("EventNumber")
	strEventType = Request.QueryString("EventType")
	strCategory = Request.QueryString("Category")
	datStartDate = Request.QueryString("StartDate")
	datEndDate = Request.QueryString("EndDate")
	strNotes = Request.QueryString("EventNotes")
	strEventModel = Request.QueryString("EventModel")
	strEventSite = Request.QueryString("EventSite")

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

	Select Case Request.QueryString("WithDevice")
		Case "Yes"
			strWithDeviceYes = "selected=""selected"""
		Case "No"
			strWithDeviceNo = "selected=""selected"""
	End Select

	Select Case Request.QueryString("UserStatus")
		Case "Enabled"
			strUserStatusEnabled = "selected=""selected"""
		Case "Disabled"
			strUserStatusDisabled = "selected=""selected"""
		Case "All"
			strUserStatusAll = "selected=""selected"""
	End Select

	Select Case Request.QueryString("Owes")
		Case "True"
			strOwesYes = "selected=""selected"""
		Case "False"
			strOwesNo = "selected=""selected"""
	End Select

	Select Case Request.QueryString("AUP")
		Case "True"
			strAUPYes = "selected=""selected"""
		Case "False"
			strAUPNo = "selected=""selected"""
	End Select

	Select Case Request.QueryString("UserView")
		Case "Card"
			strUserViewCard = "selected=""selected"""
		Case "Table"
			strUserViewTable = "selected=""selected"""
	End Select

	Select Case Request.QueryString("Assigned")
		Case "Yes"
			strAssignedYes = "selected=""selected"""
		Case "No"
			strAssignedNo = "selected=""selected"""
	End Select

	Select Case Request.QueryString("DeviceStatus")
		Case "Enabled"
			strDeviceStatusEnabled = "selected=""selected"""
		Case "Disabled"
			strDeviceStatusDisabled = "selected=""selected"""
		Case "All"
			strDeviceStatusAll = "selected=""selected"""
	End Select

	Select Case Request.QueryString("DeviceView")
		Case "Card"
			strDeviceViewCard = "selected=""selected"""
		Case "Table"
			strDeviceViewTable = "selected=""selected"""
	End Select

	Select Case Request.QueryString("Warranty")
		Case "Yes"
			strWarrantyYes = "selected=""selected"""
		Case "No"
			strWarrantyNo = "selected=""selected"""
	End Select

	Select Case Request.QueryString("Complete")
		Case "Yes"
			strCompleteYes = "selected=""selected"""
		Case "No"
			strCompleteNo = "selected=""selected"""
		Case "All"
			strCompleteAll = "selected=""selected"""
	End Select

	Select Case Request.QueryString("EventView")
		Case "Card"
			strEventViewCard = "selected=""selected"""
		Case "Table"
			strEventViewTable = "selected=""selected"""
	End Select

	If strUserLoaned = "Anything" Then
		strLoanedAnything = "selected=""selected"""
	End If

End Sub%>

<%Function GetStartOfFiscalYear(datToday)

	If IsDate(datToday) Then
		If Month(datToday) >= 7 Then
			GetStartOfFiscalYear = "7/1/" & Year(datToday)
		Else
			GetStartOfFiscalYear = "7/1/" & Year(datToday) - 1
		End If
	Else
		GetStartOfFiscalYear = GetStartOfFiscalYear(Date)
	End If

End Function%>

<%Function GetGrade(intYear)

	Dim datToday, intMonth, intCurrentYear

	datToday = Date
	intMonth = DatePart("m",datToday)
	intCurrentYear = Right(DatePart("yyyy",datToday),2)
	intYear = Right(intYear,2)

	If intMonth >= 7 And intMonth <= 12 Then
		intCurrentYear = intCurrentYear + 1
	End If

	GetGrade = 12 - (intYear - intCurrentYear)

	If GetGrade = 0 Then
		GetGrade = "K"
	End If

End Function%>

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
				GetRole = "Graduated " & 2000 + intYear
		End Select

	End If

End Function%>

<%
'Anything below here should exist on all pages
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

<%	sEnd If

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
End Function%>

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