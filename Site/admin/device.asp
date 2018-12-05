<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/16/14
'Last Updated 1/14/18

'This page shows the details for a single device in the inventory website

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim objDevice, strActiveChecked, objAssignment, objOldAssignments, objSites, objEventTypes, objEvents
Dim intTag, strSubmitTo, strMessage, strAssignedTo, objClasses, intClassOf, intStudent, objCategories
Dim bolAdapterReturned, bolCaseReturned, strInsuredChecked, strLocationMessage, strNewAssignmentMessage
Dim strAddEventMessage, strOldAssignmentMessage, intEventID, objTags, strTags, strBackLink, bolClosedEvents
Dim objMissingStuff, objLoanedOut, strMissingStuff, strLoanedOut, strCardType, strColumns, bolOpenEvent
Dim objIssuedBy, objRooms, strMACAddress, strAppleID, objLog, intClosedEventsCount, intOldAssignmentCount
Dim strViewAllToggle, objLastNames, objMakes, objModels
Dim deviceOn, WshShell, PINGFlag, ipAddress, status

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions

   Dim strSQL

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

	'Get the information about the device
   strSQL = "SELECT Manufacturer,Site,Model,Room,SerialNumber,Cart,BOCESTag,HasInsurance,DatePurchased,Active,AppleID,MACAddress,Notes,DeviceType,InternalIP,ExternalIP,LastUser,OSVersion,ComputerName,LastCheckInDate,LastCheckInTime,LastUser,OSVersion,Assigned" & vbCRLF
   strSQL = strSQL & "FROM Devices" & vbCRLF
   strSQL = strSQL & "WHERE LGTag='" & intTag & "' AND Deleted=False"
   Set objDevice = Application("Connection").Execute(strSQL)

   'Ping device to see if it is currently on
   ipAddress = objDevice(14)
   Set WshShell = CreateObject("WScript.Shell")
   PINGFlag = Not CBool(WshShell.run("ping -n 1 -w 1000 " & ipAddress,0,True))
   If PINGFlag = True Then
'    	deviceOn = "greendot.png"
   	status = "Remote Control Online"
   Else
'    	deviceOn = "reddot.png"
   	status = "Remote Control Offline"
   End If

	'See if a new MAC address was submitted, if so use it
' 	If Request.Form("MACAddress") <> "" Then
' 		strMACAddress = Request.Form("MACAddress")
' 	Else
' 		strMACAddress = objDevice(11)
' 	End If

	'See if a new Apple ID was submitted, if so use it
' 	If Request.Form("AppleID") <> "" Then
' 		strAppleID = Request.Form("AppleID")
' 	Else
' 		strAppleID = objDevice(10)
' 	End If

	'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Add Event"
         AddSubmittedEvent
      Case "Assign"
      	AssignDevice
      Case "Return"
         ReturnDevice
      Case "Update Event"
      	UpdateEvent
      Case "Update Device"
      	UpdateDevice
      Case "EMail Teachers"
      	EMailTeachers
      Case "Restore Device"
      	RestoreDevice
      Case "Disable Device"
      	DisableDevice
   End Select

	'Setup the assign a device form
   intClassOf = Request.Form("ClassOf")
   intStudent = Request.Form("StudentID")
   If intClassOf = "" Then
      intClassOf = 0
   End If
   If intStudent = "" Then
      intStudent = 0
   End If

   'If the device isn't found send them back to the index page.
   If objDevice.EOF Then
      Response.Redirect("index.asp?Error=DeviceNotFound")
   End If

   'Set the status of the insured checkbox
   If objDevice(7) Then
      strInsuredChecked = "checked=""checked"""
   Else
      strInsuredChecked = ""
   End If

   'Set the status of the active checkbox
   If objDevice(9) Then
      strActiveChecked = "checked=""checked"""
   Else
      strActiveChecked = ""
   End If

  	'Get the current assignment
   strSQL = "SELECT FirstName,LastName,ClassOf,Assignments.Notes,HasInsurance,StudentID,UserName,Role,HomeRoom,People.Active,Warning,Loaned,People.ID" & _
   	",Username,PWord,AUP,DateIssued,IssuedBy,LastExternalCheckIn,LastInternalCheckIn,Birthday" & vbCRLF
   strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
   strSQL = strSQL & "WHERE LGTag='" & intTag & "' AND Assignments.Active=True"
   Set objAssignment = Application("Connection").Execute(strSQL)

   If Not objAssignment.EOF Then

		'Get the list of loaned out items
		strSQL = "SELECT ID, Item, LoanDate FROM Loaned WHERE AssignedTo=" & objAssignment(12) & " AND Returned=False ORDER By LoanDate"
		Set objLoanedOut =  Application("Connection").Execute(strSQL)

		If Not objLoanedOut.EOF Then
			Do Until objLoanedOut.EOF
				strLoanedOut = strLoanedOut & objLoanedOut(1) & ", "
				objLoanedOut.MoveNext
			Loop
			strLoanedOut = Left(strLoanedOut,Len(strLoanedOut) - 2)
		End If

		'Get the full name of the person who loaned it out to them.
		strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objAssignment(17) & "'"
		Set objIssuedBy = Application("Connection").Execute(strSQL)

	End If

   'Get the old assignments
   strSQL = "SELECT FirstName,LastName,ClassOf,DateIssued,DateReturned,Assignments.Notes,StudentID,UserName,Role,HomeRoom,People.Active," & _
   	"Warning,IssuedBy,ReturnedBy,Loaned,People.Deleted,People.ID" & vbCRLF
   strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
   strSQL = strSQL & "WHERE LGTag='" & intTag & "' AND Assignments.Active=False ORDER BY DateIssued"
   Set objOldAssignments = Application("Connection").Execute(strSQL)

   'Count the old assignments
   intOldAssignmentCount = 0
   If Not objOldAssignments.EOF Then
   	Do Until objOldAssignments.EOF
   		intOldAssignmentCount = intOldAssignmentCount + 1
   		objOldAssignments.MoveNext
   	Loop
   	objOldAssignments.MoveFirst
   End If

   'Get the list of classes for the assign a device drop down menu
   strSQL = "SELECT DISTINCT ClassOf FROM People WHERE Active=True ORDER BY ClassOf DESC"
   Set objClasses = Application("Connection").Execute(strSQL)

   'Get the list of rooms for the auto complete
   strSQL = "SELECT DISTINCT Room FROM Devices WHERE Active=True And Room<>''"
   Set objRooms = Application("Connection").Execute(strSQL)

   'Get the list of events for this device
   strSQL = "SELECT ID,Type,Notes,EventDate,EventTime,Resolved,ResolvedDate,ResolvedTime,Category,Warranty,UserID,Site,Model,EnteredBy,CompletedBy FROM Events WHERE LGTag='" & intTag & "'"
   Set objEvents = Application("Connection").Execute(strSQL)

   'Get the list of lastnames for the auto complete
   strSQL = "SELECT DISTINCT LastName FROM People WHERE Active=True"
   Set objLastNames = Application("Connection").Execute(strSQL)

   'Get the list of makes for the auto complete
   strSQL = "SELECT DISTINCT Manufacturer FROM Devices WHERE Active=True And Manufacturer<>''"
   Set objMakes = Application("Connection").Execute(strSQL)

   'Get the list of models for the auto complete
   strSQL = "SELECT DISTINCT Model FROM Devices WHERE Active=True And Model<>''"
   Set objModels = Application("Connection").Execute(strSQL)

   'Check and see if there are any open events
   bolOpenEvent = False
   bolClosedEvents = False
   intClosedEventsCount = 0
   If Not objEvents.EOF Then
   	Do Until objEvents.EOF
   		If Not objEvents(5) Then
   			bolOpenEvent = True
   		Else
   			bolClosedEvents = True
   			intClosedEventsCount = intClosedEventsCount + 1
   		End If
   		objEvents.MoveNext
   	Loop
   	objEvents.MoveFirst
   End If

   'Get the list of tags for this device
   strSQL = "SELECT Tag FROM Tags WHERE LGTag='" & intTag & "' ORDER BY Tag"
   Set objTags = Application("Connection").Execute(strSQL)
   strTags = ""
   If Not objTags.EOF Then
   	Do Until objTags.EOF
   		strTags = strTags & objTags(0) & ", "
   		objTags.MoveNext
   	Loop
   	strTags = Left(strTags,Len(strTags) - 2)
   End If

   'Get the log items for this device
   strSQL = "SELECT LGTag,UserName,EventNumber,Type,OldValue,NewValue,UpdatedBy,LogDate,LogTime,OldNotes,NewNotes" & vbCRLF
   strSQL = strSQL & "FROM Log WHERE Active=True AND Deleted=False And LGTag='" & intTag & "' ORDER BY ID DESC"
   Set objLog = Application("Connection").Execute(strSQL)

   'Get the list of sites for the site drop down menu
   strSQL = "SELECT Site FROM Sites WHERE Active=True ORDER BY Site"
   Set objSites = Application("Connection").Execute(strSQL)

   'Get the list of event types for the event types drop down menu
   strSQL = "SELECT EventType FROM EventTypes WHERE Active=True ORDER BY EventType"
   Set objEventTypes = Application("Connection").Execute(strSQL)

   'Get the list of categories from the category drop down menu
   strSQL = "SELECT Category FROM Categories WHERE Active=True ORDER BY Category"
   Set objCategories = Application("Connection").Execute(strSQL)

   'Set the if condition for the view all icon
   strViewAllToggle = ""
	If intOldAssignmentCount > 0 Then
		strViewAllToggle = strViewAllToggle & "$(oldAssignments).is("":visible"") && "
	End If

	If intClosedEventsCount > 0 Then
		strViewAllToggle = strViewAllToggle & "$(events).is("":visible"") && "
	End If

	If Not objLog.EOF Then
		strViewAllToggle = strViewAllToggle & "$(userlog).is("":visible"") && "
	End If
	If Len(strViewAlltoggle) > 0 Then
		strViewAllToggle = Left(strViewAllToggle,Len(strViewAllToggle) - 4)
	End If

   'Get the URL used to submit forms
   If Request.ServerVariables("QUERY_STRING") = "" Then
      strSubmitTo = "device.asp"
   Else
      strSubmitTo = "device.asp?" & Request.ServerVariables("QUERY_STRING")
   End If

   'Set up the variables needed for the site then load it
   SetupSite
   DisplaySite

End Sub%>

<%Sub DisplaySite

	Dim intCounter, datToday, strSQL, objNames %>

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

				var oldAssignments = document.getElementById("OldAssignments");
				var events = document.getElementById("Events");
				var userlog = document.getElementById("UserLog");
				var siteEdit = document.getElementById("siteEdit");
				var siteView = document.getElementById("siteView");
				var roomEdit = document.getElementById("roomEdit");
				var roomView = document.getElementById("roomView");
				var tagsEdit = document.getElementById("tagsEdit");
				var tagsView = document.getElementById("tagsView");
				var macAddressEdit = document.getElementById("macAddressEdit");
				var macAddressView = document.getElementById("macAddressView");
				var appleIDEdit = document.getElementById("appleIDEdit");
				var appleIDView = document.getElementById("appleIDView");
				var bocesTagEdit = document.getElementById("bocesTagEdit");
				var bocesTagView = document.getElementById("bocesTagView");
				var serialNumberEdit = document.getElementById("serialNumberEdit");
				var serialNumberView = document.getElementById("serialNumberView");
				var modelEdit = document.getElementById("modelEdit");
				var modelView = document.getElementById("modelView");
				var makeEdit = document.getElementById("makeEdit");
				var makeView = document.getElementById("makeView");
				var assetTagEdit = document.getElementById("assetTagEdit");
				var assetTagView = document.getElementById("assetTagView");
				var purchasedView = document.getElementById("purchasedView");
				var purchasedEdit = document.getElementById("purchasedEdit");
				$('#body').show();
		<% If objDevice(3) = "" Or IsNull(objDevice(3)) Then %>
				$(roomView).hide();
		<% Else %>
				$(roomEdit).hide();
		<% End If %>

				var showHideEffect = "blind"
				var effectSpeed = 200;

				$(siteView).hide();
				$(tagsView).hide();
				$(macAddressView).hide();
				$(appleIDView).hide();
				$(purchasedEdit).hide();
				$(bocesTagEdit).hide();
				$(serialNumberEdit).hide();
				$(modelEdit).hide();
				$(makeEdit).hide();
				$(assetTagEdit).hide();

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

			<% If strViewAllToggle <> "" Then %>
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

				$("#editToggle").click(function(){
					if ($(purchasedView).is(":visible")) {
					/*	$(siteEdit).show();
						$(tagsEdit).show();
						$(macAddressEdit).show();
						$(appleIDEdit).show(); */
						$(roomEdit).show();
						$(purchasedEdit).show();
						$(bocesTagEdit).show();
						$(serialNumberEdit).show();
						$(modelEdit).show();
						$(makeEdit).show();
						$(assetTagEdit).show();

					/*	$(siteView).hide();
						$(tagsView).hide();
						$(macAddressView).hide();
						$(appleIDView).hide(); */
						$(roomView).hide();
						$(purchasedView).hide();
						$(bocesTagView).hide();
						$(serialNumberView).hide();
						$(modelView).hide();
						$(makeView).hide();
						$(assetTagView).hide();

					} else {

					/*	$(siteView).show();
						$(tagsView).show();
						$(macAddressView).show();
						$(appleIDView).show(); */
						$(purchasedView).show();
						$(bocesTagView).show();
						$(serialNumberView).show();
						$(modelView).show();
						$(makeView).show();
						$(assetTagView).show();

				<% If objDevice(3) = "" Or IsNull(objDevice(3)) Then %>
						$(roomView).hide();
						$(roomEdit).show();
				<% Else %>
						$(roomView).show();
						$(roomEdit).hide();
				<% End If %>

					/*	$(siteEdit).hide();
						$(tagsEdit).hide();
						$(macAddressEdit).hide();
						$(appleIDEdit).hide(); */
						$(purchasedEdit).hide();
						$(bocesTagEdit).hide();
						$(serialNumberEdit).hide();
						$(modelEdit).hide();
						$(makeEdit).hide();
						$(assetTagEdit).hide();
					}
					return false;
				});

				$(oldAssignments).hide();
				$(events).hide();
				$(userlog).hide();

				$( "#PurchasedDate" ).datepicker({
					changeMonth: true,
					changeYear: true,
					showOtherMonths: true,
					selectOtherMonths: true,
					onClose: function( selectedDate ) {
				   	$( "#to" ).datepicker( "option", "minDate", selectedDate );
					}
				});

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
							title: 'Events - <%=intTag%>'
						}
				<% End If %>
        			]
    			})

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

    	<% If IsMobile Then %>
    			eventTable.columns([0,3,4,5,6,8,9,10,11]).visible(false);
		<% Else %>
				eventTable.columns([3,4,6,9,10]).visible(false);
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

    		} );
    	</script>
    	<script type="text/javascript">
			$(document).ready( function () {
    			var eventTable = $('#OldAssignmentTable').DataTable( {
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
							title: 'Old Assignments - <%=intTag%>'
						}
				<% End If %>
        			]
    			})

    	<% If IsMobile Then %>
    			eventTable.columns([0,4,5,6]).visible(false);
    	<% Else %>
    			eventTable.columns([0]).visible(false);
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
							title: 'Log - <%=intTag%>'
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

    	</script>

		<script type="text/javascript">

			function setSubmitValue(value) {
					jQuery('#mouseOnValue').val(value);
				}
		</script>

<% If objAssignment.EOF Then %>

		<script type="text/javascript">

			window.onload = function()
			{

				var ddlClassOf = document.getElementById("ClassOf");

				ddlClassOf.options[0]=new Option("","");

		<%	intCounter = 1
			Do Until objClasses.EOF
				If objClasses(0) <> "" Then %>
					ddlClassOf.options[<%=intCounter%>]=new Option("<%=GetRole(objClasses(0))%>","<%=objClasses(0)%>");
					if(<%=objClasses(0)%>==<%=intClassOf%>){ddlClassOf.options[<%=intCounter%>].selected=true;}
			<% End If
				intCounter = intCounter + 1
				objClasses.MoveNext
			Loop
			objClasses.MoveFirst%>
				UpdateStudents();
			}

			function UpdateStudents()
			{

				var ddlClassOf = document.getElementById("ClassOf");
				var ddlStudent = document.getElementById("StudentID");

				ddlStudent.options.length = 0;

				switch (ddlClassOf.value)
					{

				<% If Not objClasses.EOF Then
						Do Until objClasses.EOF

							strSQL = "SELECT ID, LastName, FirstName" & vbCRLF
							strSQL = strSQL & "FROM People" & vbCRLF
							strSQL = strSQL & "WHERE Active=True AND ClassOf = " & objClasses(0) & vbCRLF
							strSQL = strSQL & "ORDER BY LastName, FirstName"
							Set objNames = Application("Connection").Execute(strSQL) %>

							case "<%=objClasses(0)%>":
								ddlStudent.options[0]=new Option("","");
						<% intCounter = 1
							Do Until objNames.EOF %>
								ddlStudent.options[<%=intCounter%>]=new Option("<%=objNames(1)%>, <%=objNames(2)%>","<%=objNames(0)%>");
								if(<%=objNames(0)%>==<%=intStudent%>){ddlStudent.options[<%=intCounter%>].selected=true;}
							<% intCounter = intCounter + 1
								objNames.MoveNext
							Loop %>
								break;
						<% objClasses.MoveNext
						Loop
						objClasses.MoveFirst
					End If%>
					}
			}
		</script>
<% End If %>

   </head>

   <body class="<%=strSiteVersion%>" id="body" style="display:none;" >

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
		JumpToDevice %>
		<div Class="<%=strColumns%>">
		<%	DeviceInformationCombined
			ActiveAssignments

			If bolOpenEvent Then
				OpenEventsCards
			Else
				If objDevice(9) Then
					AddEvent
				End If
			End If

			OldAssignmentsTable
			OldEventsTable
			ShowLog

			%>
		</div>
	<%	'SearchForDevice
		%>
		<%=strMessage%>

		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
	</body>
	</html>

<%End Sub%>

<%Sub DeviceInformationCombined

	Dim objFSO, strDeviceInfo

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strDeviceInfo = "Name: " & objDevice(18) & " &#013 "
	strDeviceInfo = strDeviceInfo & "Last User: " & objDevice(21) & " &#013 "
	strDeviceInfo = strDeviceInfo & "OS Version: " & objDevice(22) & " &#013 "
	strDeviceInfo = strDeviceInfo & "Last Checkin: " & objDevice(19) & " - " & objDevice(20)
	%>

<% If objDevice(9) Then
		strCardType = "NormalCard"
	Else
		strCardType = "DisabledCard"
	End If 
	
	If bolOpenEvent Then
		strCardType = "WarningCard"
	End If %>
	
	<div class="Card <%=strCardType%>">
		<form method="POST" action="<%=strSubmitTo%>">
		<button style="overflow: visible !important; height: 0 !important; width: 0 !important; margin: 0 !important; border: 0 !important; padding: 0 !important; display: block !important;" type="submit" name="Submit" value="Update Device" /></button>
		<div class="CardTitle" id="assetTagView">
		<% If objDevice(19) <> "" Then %>
			<% If Application("MunkiReportServer") = "" Then %>
					<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strDeviceInfo%>"  />&nbsp;</div>
			<% Else %>
					<a href="<%=Application("MunkiReportServer")%>/index.php?/clients/detail/<%=objDevice(4)%>" target="_blank">
						<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strDeviceInfo%>"  />&nbsp;</div>
					</a>
			<% End If %>
		<% End If %>
		<% If objDevice(7) Then %>
				<image src="../images/yes.png" width="15" height="15" title="Insured" />
		<% End If %>
		<!--
<% If objDevice(14) <> "" Then %>
			<image src="../images/<%=deviceOn%>" width="8" height="8" title="<%=status%>" /> Asset Tag <%=intTag%>
		<% Else %>
			Asset Tag <%=intTag%>
		<% End If %>
 -->
 		Asset Tag <%=intTag%>
		</div>
		<div class="CardTitle" id="assetTagEdit">
		<input type="image" src="../images/disable.png" value="Disable Device" name="Submit" width="15" height="15" title="Decommissioned Device" onmouseover="setSubmitValue('Disable Device')"/>
		<% If objDevice(7) Then %>
				<image src="../images/yes.png" width="15" height="15" title="Insured" />
		<% End If %>
			Asset Tag <input Class="Card InputWidthSmall" type="text" name="AssetTag" value="<%=intTag%>">
		</div>
		<div Class="ImageSectionInCard">
			<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDevice(2)," ","") & ".png") Then %>
					<img class="PhotoCard" src="../images/devices/<%=Replace(objDevice(2)," ","")%>.png" width="96" />
			<% Else %>
					<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
			<% End If %>
		</div>
		<div Class="RightOfImageInCard">
			<div>
				<div Class="PhotoCardColumn1">Make: </div>
				<div Class="PhotoCardColumn2" id= "makeView"><%=objDevice(0)%></div>
				<div Class="PhotoCardColumn2" id= "makeEdit"><input Class="Card InputWidthMedium" type="text" name="Make" value="<%=objDevice(0)%>" id="Makes"></div>
			</div>
			<div>
				<div Class="PhotoCardColumn1">Model: </div>
				<div Class="PhotoCardColumn2" id= "modelView"><%=objDevice(2)%></div>
				<div Class="PhotoCardColumn2" id= "modelEdit"><input Class="Card InputWidthMedium" type="text" name="Model" value="<%=objDevice(2)%>" id="Models"></div>
			</div>
			<div>
				<div Class="PhotoCardColumn1">Serial: </div>
				<div Class="PhotoCardColumn2" id="serialNumberView">
			<% Select Case objDevice(0)
					Case "Apple" %>
						<a href="https://checkcoverage.apple.com?caller=sp&sn=<%=objDevice(4)%>" target="_blank"><%=objDevice(4)%></a>
				<% Case "Dell" %>
						<a href="http://www.dell.com/support/home/us/en/19/product-support/servicetag/<%=objDevice(4)%>" target="_blank"><%=objDevice(4)%></a>
				<% Case Else %>
						<%=objDevice(4)%>
			 <% End Select %>
				</div>
				<div Class="PhotoCardColumn2" id="serialNumberEdit">
					<input Class="Card InputWidthMedium" type="text" name="serialNumber" value="<%=objDevice(4)%>" >
				</div>
			</div>
			<% 'If objDevice(6) <> "" Then Turned this off so the BOCES field is always displayed.  Needed for edit mode.%>
				<div>
					<div Class="CardMerged" id="bocesTagEdit">BOCES Tag: <input Class="Card InputWidthSmall" type="text" name="BOCESTag" value="<%=objDevice(6)%>" ></div>
					<div Class="CardMerged" id="bocesTagView">BOCES Tag: <%=objDevice(6)%> </div>
				</div>
			<% 'End If %>
		<% If objDevice(8) <> "" Then %>
				<div>
					<div Class="CardMerged" id="purchasedEdit">Purchased: <input Class="Card InputWidthDate" type="text" name="Purchased" value="<%=ShortenDate(objDevice(8))%>" id="PurchasedDate" ></div>
					<div Class="CardMerged" id="purchasedView">Purchased: <%=ShortenDate(objDevice(8))%> - Year <%=GetAge(objDevice(8))%></div>
				</div>
		<% End If %>
	</div>
      <div>
         <div Class="CardColumn1">Site: </div>
         <div Class="CardColumn2">
            <select Class="Card" name="Site" id="siteEdit">
                  <option value=""></option>
            <% Do Until objSites.EOF
                  If objSites(0) = objDevice(1) Then %>
                     <option selected="selected" value="<%=objSites(0)%>"><%=objSites(0)%></option>
               <% Else %>
                     <option value="<%=objSites(0)%>"><%=objSites(0)%></option>
               <% End If %>
            <%    objSites.MoveNext
               Loop
               objSites.MoveFirst%>
            </select>
            <div id="siteView"><%=objDevice(1)%></div>
         </div>
      </div>
      <div>
         <div Class="CardColumn1">Room: </div>
         <div Class="CardColumn2" id="roomEdit"><input Class="Card InputWidthLarge" type="text" name="Room" value="<%=objDevice(3)%>" id="Rooms" ></div>
         <div Class="CardColumn2" id="roomView"><a href="devices.asp?Room=<%=objDevice(3)%>&DeviceSite=<%=objDevice(1)%>&View=Card"><%=objDevice(3)%></a></div>
      </div>
      <div>
         <div Class="CardColumn1">Tags: </div>
         <div Class="CardColumn2" id="tagsEdit"><input Class="Card InputWidthLarge" type="text" name="Tags" value="<%=strTags%>"></div>
         <div Class="CardColumn2" id="tagsView"><%=strTags%></div>
      </div>
<!--<% If objDevice(13) = "Laptop" Or objDevice(13) = "Desktop" Or objDevice(13) = "Access Point" Then %>
		<div>
			<div Class="CardColumn1">MAC Address: </div>
			<div Class="CardColumn2" id="macAddressEdit"><input Class="Card InputWidthLarge" type="text" name="MACAddress" value="<%=strMACAddress%>"></div>
			<div Class="CardColumn2" id="macAddressView"><%=strMACAddress%></div>
		</div>
  	<% End If %>
   <% If InStr(objDevice(2),"iPad") Then %>
			<div>
				<div Class="CardColumn1">Apple ID: </div>
				<div Class="CardColumn2" id="appleIDEdit"><input Class="Card InputWidthLarge" type="text" name="AppleID" value="<%=strAppleID%>"></div>
				<div Class="CardColumn2" id="appleIDView"><%=strAppleID%></div>
			</div>
   <% End If %> -->
   	<div>Device Notes: </div>
		<div>
			<textarea class="Card" rows="5" name="Notes" cols="90" style="width: 99%;"><%=objDevice(12)%></textarea>
		</div>
		<div>&nbsp;</div>
      <input type="hidden" name="Insured" value="<%=objDevice(7)%>" />
      <input type="hidden" name="Submit" value="Update Device" id="mouseOnValue" />
      <div>
         <div class="Button"><input type="image" src="../images/save.png" width="20" height="20" title="Update Device" onmouseover="setSubmitValue('Update Device')" /></div>
      </div>

	<% If objDevice(9) Then %>
			<a href="" class="Button" id="editToggle">
				<image src="../images/edit.png" height="20" width="20" title="Toggle Edit Mode">
			</a>
	<% Else %>
			<div>
				<div class="Button"><input type="image" src="../images/restore.png" value="Restore Device" name="Submit" width="20" height="20" title="Restore Device" onmouseover="setSubmitValue('Restore Device')" /></div>
			</div>
	<% End If %>

	<% If Request.QueryString("Back") <> "" Then
			DrawIcon "Back",0,""
	   End If

	   If Len(strViewAllToggle) > 0 Then
	   	DrawIcon "ViewAll",0,""
	   End If

	   If intOldAssignmentCount > 0 Then
			DrawIcon "Assignments",0,""
			Response.Write "<div class=""ButtonText"">" & intOldAssignmentCount & "</div>"
	   End If

		If intClosedEventsCount > 0 Then
			DrawIcon "Events",0,""
			Response.Write "<div class=""ButtonText"">" & intClosedEventsCount & "</div>"
		End If

		If Not objLog.EOF Then
			DrawIcon "Log",0,""
		End If

		If objDevice(14) <> "" Then
			DrawIcon "Remote",0,""
		End If
		%>

	<% If strLocationMessage <> "" Then %>
			<%=strLocationMessage%>
	<% End If %>

      </form>
   </div>

<%End Sub%>

<%Sub DeviceInformation

	Dim objFSO

   Set objFSO = CreateObject("Scripting.FileSystemObject")%>

<% If objDevice(9) Then
		strCardType = "NormalCard"
	Else
		strCardType = "DisabledCard"
	End If %>
	<div class="Card <%=strCardType%>">
		<div class="CardTitle">
		<% If objDevice(7) Then %>
				<image src="../images/yes.png" width="15" height="15" title="Insured" />
		<% End If %>
			Asset Tag <%=intTag%>
		</div>
	<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDevice(2)," ","") & ".png") Then %>
			<% If InStr(LCase(objDevice(2)),"ipad") Then %>
					<img class="PhotoCard" src="../images/devices/<%=Replace(objDevice(2)," ","")%>.png" width="70" />
			<% Else %>
					<img class="PhotoCard" src="../images/devices/<%=Replace(objDevice(2)," ","")%>.png" width="96" />
			<% End If %>
	<% Else %>
			<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
	<% End If %>

      <div>
         <div Class="PhotoCardColumn1">Make: </div>
         <div Class="PhotoCardColumn2"><%=objDevice(0)%></div>
      </div>
      <div>
         <div Class="PhotoCardColumn1">Model: </div>
         <div Class="PhotoCardColumn2"><%=objDevice(2)%></div>
      </div>
      <div>
         <div Class="PhotoCardColumn1">Serial: </div>
         <div Class="PhotoCardColumn2">
      <% Select Case objDevice(0)
      		Case "Apple" %>
         		<a href="https://checkcoverage.apple.com/us/en/?sn=<%=objDevice(4)%>" target="_blank"><%=objDevice(4)%></a>
       	<% Case "Dell" %>
       			<a href="http://www.dell.com/support/home/us/en/19/product-support/servicetag/<%=objDevice(4)%>" target="_blank"><%=objDevice(4)%></a>
       	<% Case Else %>
       			<%=objDevice(4)%>
       <% End Select %>
         </div>
      </div>
      <% If objDevice(6) <> "" Then %>
			<div>
				<div Class="CardMerged">BOCES Tag: <%=objDevice(6)%></div>
			</div>
		<% End If %>
	<% If objDevice(8) <> "" Then %>
			<div>
				<div Class="CardMerged">Purchased: <%=ShortenDate(objDevice(8))%> - Year <%=GetAge(objDevice(8))%></div>
			</div>
	<% End If %>
	<% If Request.QueryString("Back") <> "" Then
			DrawIcon "Back",1,""
		End If %>
   </div>

   <div class="Card NormalCard">
      <form method="POST" action="<%=strSubmitTo%>">
      <div class="CardTitle">Location</div>
      <div>
         <div Class="CardColumn1">Site: </div>
         <div Class="CardColumn2">
            <select Class="Card" name="Site">
                  <option value=""></option>
            <% Do Until objSites.EOF
                  If objSites(0) = objDevice(1) Then %>
                     <option selected="selected" value="<%=objSites(0)%>"><%=objSites(0)%></option>
               <% Else %>
                     <option value="<%=objSites(0)%>"><%=objSites(0)%></option>
               <% End If %>
            <%    objSites.MoveNext
               Loop
               objSites.MoveFirst%>
            </select>
         </div>
      </div>
      <div>
         <div Class="CardColumn1">Room: </div>
         <div Class="CardColumn2"><input Class="Card InputWidthLarge" type="text" name="Room" value="<%=objDevice(3)%>"></div>
      </div>
      <div>
         <div Class="CardColumn1">Tags: </div>
         <div Class="CardColumn2"><input Class="Card InputWidthLarge" type="text" name="Tags" value="<%=strTags%>"></div>
      </div>
   <% If InStr(objDevice(2),"iPad") Then %>
			<div>
				<div Class="CardColumn1">Apple ID: </div>
				<div Class="CardColumn2"><input Class="Card InputWidthLarge" type="text" name="AppleID" value="<%=objDevice(10)%>"></div>
			</div>
   <% End If %>
      <div>
         <div Class="CardColumn1">Insured: </div>
         <div Class="CardColumn2">
         	<input type="checkbox" name="Insured" <%=strInsuredChecked%> value="True" />
         </div>
      </div>
      <div>
         <div class="Button"><input type="submit" value="Update Device" name="Submit" /></div>
      </div>
   <% If strLocationMessage <> "" Then %>
   	<div>
   		<%=strLocationMessage%>
   	</div>
   <% End If %>
      </form>
   </div>

<% End Sub%>

<%Sub ActiveAssignments

	Dim intAge %>

<% If Not objAssignment.EOF Then

      Dim objFSO, strUserInfo

      Set objFSO = CreateObject("Scripting.FileSystemObject")

      'Build the user info popup
		strUserInfo = ""
		If objAssignment(19) <> "" Then
			strUserInfo = "Internal Access: " & objAssignment(19) & " &#013 "
		End If
		If objAssignment(18) Then
			strUserInfo = strUserInfo & "External Access: " & objAssignment(18) & " &#013 "
		End If
		If objAssignment(20) <> "" Then
			intAge = DateDiff("yyyy",objAssignment(20),Date)
			If Date < DateSerial(Year(Date), Month(objAssignment(20)), Day(objAssignment(20))) Then
				intAge = intAge - 1
			End If
			strUserInfo = strUserInfo & "Birthday: " & objAssignment(20) & " &#013 "
			strUserInfo = strUserInfo & "Age: " & intAge
		End If

      Do Until objAssignment.EOF %>

		<% If objAssignment(10) Then
				strCardType = "WarningCard"
			ElseIf objAssignment(11)Then
				strCardType = "LoanedCard"
			ElseIf objAssignment(9) Then
				strCardType = "NormalCard"
			Else
				strCardType = "DisabledCard"
			End If %>
			<div class="Card <%=strCardType%>">
				<div class="CardTitle">
				<% If objAssignment(7) = "Student" Then %>
					<% If Application("ShowPasswords") Then %>
						<% If objAssignment(15) Then %>
								<image src="../images/yes.png" width="15" height="15" title="AUP Signed" />
						<% Else %>
								<image src="../images/no.png" width="15" height="15" title="AUP Not Signed" />
						<% End If %>
					<% End If %>
				<% End If %>
					<image src="../images/assignment.png" width="15" height="15" title="Assignments" /> Active Assignment
				<% If strUserInfo <> "" Then %>
						<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strUserInfo%>"  />&nbsp;</div>
				<% End If %>
				</div>
				<form method="POST" action="<%=strSubmitTo%>">
				<div Class="ImageSectionInCard">
			<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objAssignment(7) & "s\" & objAssignment(5) & ".jpg") Then %>
					<a href="user.asp?UserName=<%=objAssignment(6)%><%=strBackLink%>">
						<img class="PhotoCard" src="/photos/<%=objAssignment(7)%>s/<%=objAssignment(5)%>.jpg" title="<%=objAssignment(5)%>" width="96" />
					</a>
			<% Else %>
					<a href="user.asp?UserName=<%=objAssignment(6)%><%=strBackLink%>">
						<img class="PhotoCard" src="/photos/<%=objAssignment(7)%>s/missing.png" title="<%=objAssignment(5)%>" width="96" />
					</a>
			<% End If %>
				</div>
				<div Class="RightOfImageInCard">
					<div Class="PhotoCardColumn1">Name: </div>
					<div Class="PhotoCardColumn2">
						<a href="user.asp?UserName=<%=objAssignment(6)%><%=strBackLink%>"><%=objAssignment(0) & " " & objAssignment(1)%></a>
					</div>
					<div>
						<div Class="PhotoCardColumn1">Role: </div>
						<div Class="PhotoCardColumn2Long">
							<a href="users.asp?Role=<%=objAssignment(2)%>"><%=GetRole(objAssignment(2))%></a>
						</div>
					</div>

					<div Class="CardMerged">
						Adapter Returned: <input Class="Card" type="checkbox" checked="checked" name="Adapter" value="True" />
					</div>
					<div Class="CardMerged">
						Case Returned: <input Class="Card" type="checkbox" checked="checked" name="Case" value="True" />
					</div>
					<div Class="CardMerged">
						Returned Damaged: <input Class="Card" type="checkbox" name="Damaged" value="True" />
					</div>
				</div>
			<% If objAssignment(7) = "Student" Then
					If Application("ShowPasswords") Then %>
						<div>
							<div Class="CardMerged">Username: <%=objAssignment(13)%></div>
						</div>
						<div>
							<div Class="CardMerged">Password: <%=objAssignment(14)%></div>
						</div>
				<% End If %>
				<% If objAssignment(8) <> "" Then %>
						<div Class="CardMerged">
							<div><%=Application("HomeroomNameLong")%>:
								<a href="users.asp?GuideRoom=<%=objAssignment(8)%>"><%=objAssignment(8)%></a>
							</div>
						</div>
				<% End If %>
			<% End If %>

			<% If objAssignment(16) <> ""  Then %>
				<% If Not objIssuedBy.EOF Then %>
						<div>
							<div Class="CardMerged">Assigned on <%=ShortenDate(objAssignment(16))%> by <%=objIssuedBy(0)%>&nbsp;<%=objIssuedBy(1)%></div>
						</div>
				<% Else %>
						<div>
							<div Class="CardMerged">Assigned on <%=ShortenDate(objAssignment(16))%></div>
						</div>
				<% End If %>
			<% End If %>

			<% If strMissingStuff <> "" Then %>
					<div Class="CardMerged">
						<div>
							Missing: <%=strMissingStuff%>
						</div>
					</div>
			<% End If %>

			<% If strLoanedOut <> "" Then %>
					<div Class="CardMerged">
						<div>
							Borrowing: <%=strLoanedOut%>
						</div>
					</div>
			<% End If %>

				<div>Assignment Notes: </div>
				<div>
					<textarea class="Card" rows="5" name="Notes" cols="90" style="width: 99%;"></textarea>
				</div>
				<div>&nbsp;</div>
				<div class="Button"><input type="submit" value="Return" name="Submit" /></div>
				<div class="Button"><input type="submit" value="EMail Teachers" name="Submit" />&nbsp;&nbsp;</div>
			<% If Application("HelpDeskURL") <> "" Then
					DrawIcon "HelpDesk",0,objAssignment(6)
				End If

			   If strNewAssignmentMessage <> "" Then %>
				<div>
					<%=strNewAssignmentMessage%>
				</div>
			<% End If %>
				</form>
			</div>

      <% objAssignment.MoveNext
      Loop

   Else
   	If objDevice(9) Then%>

		<div class="Card NormalCard">
         <form method="POST" action="<%=strSubmitTo%>">
         <div class="CardTitle"><image src="../images/assignment.png" width="15" height="15" title="Assignments" /> Assign Device</div>
         <div>
            <div Class="CardColumn1">Role: </div>
            <div Class="CardColumn2">
               <select class="Card" name="ClassOf" id="ClassOf" onchange="UpdateStudents();">
               </select>
            </div>
         </div>
         <div>
            <div Class="CardColumn1">Person: </div>
            <div Class="CardColumn2">
               <select class="Card" name="StudentID" id="StudentID">
               </select>
            </div>
         </div>
         <div>
         <% If Not bolOpenEvent Then %>
           		<div class="Button"><input type="submit" value="Assign" name="Submit" /></div>
         <% Else %>
         		<div class="Button"><input type="submit" value="Assign" disabled name="Submit" /></div>
         <% End If %>
         </div>
         <% If strOldAssignmentMessage <> "" Then %>
				<div>
					<%=strOldAssignmentMessage%>
				</div>
			<% End If %>
         </form>
      </div>
      <% End If %>

<% End If %>

<%End Sub%>

<%Sub OldAssignments%>

   <% If Not objOldAssignments.EOF Then

      Dim objFSO

      Set objFSO = CreateObject("Scripting.FileSystemObject")

      Do Until objOldAssignments.EOF

			If objOldAssignments(11) Then
				strCardType = "WarningCard"
			ElseIf objOldAssignments(10) Then
				strCardType = "OldAssignmentCard"
			Else
				strCardType = "DisabledCard"
			End If %>
		<div class="Card <%=strCardType%>">
				<div class="CardTitle">Old Assignment</div>
				<div>
			<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objOldAssignments(8) & "s\" & objOldAssignments(6) & ".jpg") Then %>
					<a href="user.asp?UserName=<%=objOldAssignments(7)%><%=strBackLink%>">
						<img class="PhotoCard" src="/photos/<%=objOldAssignments(8)%>s/<%=objOldAssignments(6)%>.jpg" title="<%=objOldAssignments(6)%>" width="96" />
					</a>
			<% Else %>
					<a href="user.asp?UserName=<%=objOldAssignments(7)%><%=strBackLink%>">
						<img class="PhotoCard" src="/photos/<%=objOldAssignments(8)%>s/missing.png" title="<%=objOldAssignments(6)%>" width="96" />
					</a>
			<% End If %>
				</div>

				<div>
					<div Class="PhotoCardColumn1">Name: </div>
					<div Class="PhotoCardColumn2">
						<a href="user.asp?UserName=<%=objOldAssignments(7)%><%=strBackLink%>"><%=objOldAssignments(0) & " " & objOldAssignments(1)%></a>
					</div>
				</div>
				<div>
					<div Class="PhotoCardColumn1">Date: </div>
					<div Class="PhotoCardColumn2"><%=ShortenDate(objOldAssignments(3)) & " - " & ShortenDate(objOldAssignments(4))%></div>
				</div>
		<% If objOldAssignments(8) = "Student" Then %>
				<div>
					<div Class="PhotoCardColumn1">Guide: </div>
					<div Class="PhotoCardColumn2Long">
						<a href="users.asp?GuideRoom=<%=objOldAssignments(9)%>"><%=objOldAssignments(9)%></a>
					</div>
				</div>
		<% End If %>
		<% If objOldAssignments(5) <> "" Then %>
			<% If objOldAssignments(8) = "Student" Then %>
					<div>&nbsp;</div>
			<% Else %>
					<div>&nbsp;</div>
					<div>&nbsp;</div>
			<% End If %>
				<div>Assignment Notes: </div>
			<% If objOldAssignments(8) = "Student" Then
					If Application("HelpDeskURL") <> "" Then
						DrawIcon "HelpDesk",2,objOldAssignments(7)
					End If
			   End If %>
			<div><%=objOldAssignments(5)%></div>
		<% Else %>
			<% If objOldAssignments(8) = "Student" Then
					DrawIcon "HelpDesk",1,objOldAssignments(7)
			   End If %>
		<% End If %>

			</div>

      <% objOldAssignments.MoveNext
      Loop
      objOldAssignments.MoveFirst %>
   <% End If %>

<%End Sub%>

<%Sub OldAssignmentsTable

	Dim objFSO, strSQL, strAssignedBy, strReturnedBy, objPersonLookup, strRowClass

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If Not objOldAssignments.EOF Then %>
		<div id="OldAssignments">
			<br />
			<image src="../images/assignment.png" height="15" width="15" title="Old Assignments"> Old Assignments
			<table align="center" Class="ListView" id="OldAssignmentTable">
				<thead>
					<th>Photo</th>
					<th>Name</th>
					<th>Start Date</th>
					<th>End Date</th>
					<th>Assigned By</th>
					<th>Returned By</th>
					<th>Assignment Notes</th>
				</thead>
				<tbody>
		<% Do Until objOldAssignments.EOF

				strAssignedBy = ""
				strReturnedBy = ""

				If objOldAssignments(11) Then
					strRowClass = " Class=""Warning"""
				ElseIf objOldAssignments(14) Then
					strRowClass = " Class=""Loaned"""
				ElseIf objOldAssignments(10) Then
					strRowClass = ""
				Else
					strRowClass = " Class=""Disabled"""
				End If %>

				<tr <%=strRowClass%>>

					<td <%=strRowClass%> width="1px">
				<% If objOldAssignments(15) Then %>
					<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objOldAssignments(8) & "s\" & objOldAssignments(6) & ".jpg") Then %>
							<img src="/photos/<%=objOldAssignments(8)%>s/<%=objOldAssignments(6)%>.jpg" title="<%=objOldAssignments(6)%>" width="96" />
					<% Else %>
							<img src="/photos/<%=objOldAssignments(8)%>s/missing.png" title="<%=objOldAssignments(6)%>" width="96" />
					<% End If %>
				<% Else %>
					<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objOldAssignments(8) & "s\" & objOldAssignments(6) & ".jpg") Then %>
							<a href="user.asp?UserName=<%=objOldAssignments(7)%><%=strBackLink%>">
								<img src="/photos/<%=objOldAssignments(8)%>s/<%=objOldAssignments(6)%>.jpg" title="<%=objOldAssignments(6)%>" width="96" />
							</a>
					<% Else %>
							<a href="user.asp?UserName=<%=objOldAssignments(7)%><%=strBackLink%>">
								<img src="/photos/<%=objOldAssignments(8)%>s/missing.png" title="<%=objOldAssignments(6)%>" width="96" />
							</a>
					<% End If %>
				<% End If %>
					</td>

				<% If objOldAssignments(15) Then %>
						<td <%=strRowClass%>><%=objOldAssignments(1) & ", " & objOldAssignments(0)%></td>
				<% Else %>
						<td <%=strRowClass%>><a href="user.asp?UserName=<%=objOldAssignments(7)%><%=strBackLink%>"><%=objOldAssignments(1) & ", " & objOldAssignments(0)%></a></td>
				<% End If %>

					<td <%=strRowClass%> id="center"><%=ShortenDate(objOldAssignments(3))%></td>

				<% If objDevice(9) Then %>
				   <% If Not objDevice(23) Then %>
					<% If objOldAssignments(10) Then %>
							<form method="POST" action="<%=strSubmitTo%>">
								<input type="hidden" name="StudentID" value="<%=objOldAssignments(16)%>" />
								<input type="hidden" name="Tag" value="<%=intTag%>" />
								<td <%=strRowClass%> id="center">
									<input type="hidden" value="Assign" name="Submit" />
									<%=ShortenDate(objOldAssignments(4))%> <input type="image" src="../images/assignment.png" width="15" height="15" title="Reassign Device" />
								</td>
							</form>
					<% Else %>
							<td <%=strRowClass%> id="center"><%=ShortenDate(objOldAssignments(4))%></td>
					<% End If %>
					<% Else %>
					      <td <%=strRowClass%> id="center"><%=ShortenDate(objOldAssignments(4))%></td>
					<% End If %>
				<% Else %>
						<td <%=strRowClass%> id="center"><%=ShortenDate(objOldAssignments(4))%></td>
				<% End If %>

				<% If objOldAssignments(12) = "" Then %>
						<td <%=strRowClass%>>&nbsp;</td>
				<% Else

						strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objOldAssignments(12) & "'"
						Set objPersonLookup = Application("Connection").Execute(strSQL) %>

					<%	If Not objPersonLookup.EOF Then %>
							<td <%=strRowClass%>><%=objPersonLookup(1)%>, <%=objPersonLookup(0)%></td>
					<% Else %>
							<td <%=strRowClass%>><%=objOldAssignments(12)%></td>
					<%End If%>

				<% End If %>

				<% If objOldAssignments(13) = "" Then %>
						<td <%=strRowClass%>>&nbsp;</td>
				<% Else

						strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objOldAssignments(13) & "'"
						Set objPersonLookup = Application("Connection").Execute(strSQL) %>

					<%	If Not objPersonLookup.EOF Then %>
							<td <%=strRowClass%>><%=objPersonLookup(1)%>, <%=objPersonLookup(0)%></td>
					<% Else %>
							<td <%=strRowClass%>><%=objOldAssignments(13)%></td>
					<%End If%>

				<% End If %>

				<% If NOT IsNull(objOldAssignments(5)) Then %>
						<td <%=strRowClass%>><%=Replace(objOldAssignments(5),vbCRLF,"<br />")%></td>
				<% Else %>
						<td <%=strRowClass%>><%=objOldAssignments(5)%></td>
				<% End If%>
				</tr>

			<%	objOldAssignments.MoveNext
			Loop
			objOldAssignments.MoveFirst %>
				</tbody>
			</table>
		</div>
<% End If %>

<%End Sub%>

<%Sub OpenEventsCards

	Dim strSelected, strWarrantyChecked

	If Not objEvents.EOF Then %>

   <% Do Until objEvents.EOF

   		If objEvents(9) Then
				strWarrantyChecked = "checked=""checked"""
			Else
				strWarrantyChecked = ""
			End If

   		If Not objEvents(5) Then %>

				<div class="Card NormalCard">
					<form method="POST" action="<%=strSubmitTo%>">
					<input type="hidden" name="EventID" value="<%=objEvents(0)%>" />
					<div class="CardTitle">Event <%=objEvents(0)%></div>
					<div>
						<div Class="CardColumn1">Event Type: </div>
						<div Class="CardColumn2">
							<select Class="Card" name="EventType">
								<option value=""></option>
						<% If Not objEventTypes.EOF Then
								Do Until objEventTypes.EOF
									If objEvents(1) = objEventTypes(0) Then
										strSelected = "selected=""selected"""
									Else
										strSelected = ""
									End If %>
									<option value="<%=objEventTypes(0)%>" <%=strSelected%>><%=objEventTypes(0)%></option>
								<% objEventTypes.MoveNext
								Loop
							End If %>
							</select>
						</div>
					</div>
					<div>
						<div Class="CardColumn1">Submitted: </div>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3)) & " " & ShortenTime(objEvents(4))%></div>
					</div>
					<div Class="CardColumn1">Category:</div>
					<div Class="CardColumn2">
						<select Class="Card" name="Category">
							<option value=""></option>
					<% If Not objCategories.EOF Then
							Do Until objCategories.EOF
								If objCategories(0) = objEvents(8) Then
									strSelected = "selected=""selected"""
								Else
									strSelected = ""
								End If %>
								<option value="<%=objCategories(0)%>" <%=strSelected%>><%=objCategories(0)%></option>
							<% objCategories.MoveNext
							Loop
							objCategories.MoveFirst
						End If %>
						</select>
					</div>
					<div Class="CardColumn1">Warranty:</div>
					<div Class="CardColumn2">
						<input type="checkbox" name="Warranty" value="True" <%=strWarrantyChecked%> />
					</div>
					<div>
						<div Class="CardColumn1">Complete: </div>
						<div Class="CardColumn2">
							<input Class="Card" type="checkbox" value="TRUE" name="Resolved" />
						</div>
					</div>
					<div>Event Notes: </div>
					<div>
						<textarea Class="Card" rows="5" name="Notes" cols="90" style="width: 99%;"><%=objEvents(2)%></textarea>
					</div>
					<div>&nbsp;</div>
					<div Class="Button"><input type="submit" value="Update Event" name="Submit" /></div>
			<% If CInt(intEventID) = CInt(objEvents(0)) Then %>
					<div>
						<div class="Information">Updated</div>
					</div>
			<% End If %>
					</form>
				</div>

      <% End If
      	objEvents.MoveNext
      Loop
      objEvents.MoveFirst
   End If %>


<%End Sub%>

<%Sub OldEventsTable

	Dim strWarrantyInfo, strSQL, objName

	If bolClosedEvents Then %>
		<div id="Events">
			<br />
			<image src="../images/event.png" height="15" width="15" title="Events"> Completed Events
			<table align="center" Class="ListView" id="EventTable">
				<thead>
					<th>Event</th>
					<th>Type</th>
					<th>Category</th>
					<th>Model</th>
					<th>Site</th>
					<th>Start Date</th>
					<th>End Date</th>
					<th>User</th>
					<th>Warranty</th>
					<th>Entered By</th>
					<th>Completed By</th>
					<th>Event Notes</th>
				</thead>
				<tbody>
		<% Do Until objEvents.EOF

				If objEvents(5) Then

					If objEvents(9) Then
						strWarrantyInfo = "Yes"
					Else
						strWarrantyInfo = "No"
					End If %>

					<tr>
						<td id="center"><%=objEvents(0)%></td>
						<td><a href="events.asp?EventType=<%=objEvents(1)%>&View=Table"><%=objEvents(1)%></a></td>
						<td><a href="events.asp?Category=<%=objEvents(8)%>&View=Table"><%=objEvents(8)%></a></td>
						<td><a href="events.asp?EventModel=<%=objEvents(12)%>&View=Table"><%=objEvents(12)%></a></td>
						<td><a href="events.asp?EventSite=<%=objEvents(11)%>&View=Table"><%=objEvents(11)%></a></td>
						<td><%=ShortenDate(objEvents(3))%></td>
						<td><%=ShortenDate(objEvents(6))%></td>
					<% If objEvents(10) <> "" Then

							strSQL = "SELECT FirstName,LastName,UserName FROM People WHERE ID=" & objEvents(10)
							Set objName = Application("Connection").Execute(strSQL)

							If Not objName.EOF Then %>
								<td>
									<a href="user.asp?UserName=<%=objName(2)%><%=strBackLink%>"><%=objName(1)%>, <%=objName(0)%></a>
								</td>
						<% Else %>
								<td>&nbsp;</td>
						<%	End If %>

					<% Else %>
							<td>&nbsp;</td>
					<% End If %>
						<td id="center"><a href="events.asp?Warranty=<%=strWarrantyInfo%>&View=Table"><%=strWarrantyInfo%></a></td>

					<% If objEvents(13) <> "" Then

							strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objEvents(13) & "'"
							Set objName = Application("Connection").Execute(strSQL)

							If Not objName.EOF Then %>
								<td><%=objName(1)%>, <%=objName(0)%></td>
						<% Else %>
								<td></td>
						<%	End If %>

					<% Else %>
							<td></td>
					<% End If %>

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


					<% If NOT IsNull(objEvents(2)) Then %>
						<td><%=Replace(objEvents(2),vbCRLF,"<br />")%></td>
					<% Else %>
						<td><%=objEvents(2)%></td>
					<% End If %>

					</tr>
			<%	End If
				objEvents.MoveNext
			Loop
			objEvents.MoveFirst %>
				</tbody>
			</table>
		</div>
<% End If %>

<%End Sub%>

<%Sub Events

	Dim strSelected, strWarrantyChecked

	If Not objEvents.EOF Then %>

   <% Do Until objEvents.EOF

   		If objEvents(9) Then
				strWarrantyChecked = "checked=""checked"""
			Else
				strWarrantyChecked = ""
			End If

   		If objEvents(5) Then %>

				<div class="Card NormalCard">
					<div class="CardTitle">Resolved Event</div>
					<div>
						<div Class="CardColumn1">Event Type: </div>
						<div Class="CardColumn2"><%=objEvents(1)%></div>
					</div>
					<div>
						<div Class="CardColumn1">Date: </div>
				<% If ShortenDate(objEvents(3)) = ShortenDate(objEvents(6)) Then %>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3))%></div>
				<% Else %>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3)) & " - " & ShortenDate(objEvents(6))%></div>
				<% End If %>
					</div>
					<div>
						<div Class="CardColumn1">Category: </div>
						<div Class="CardColumn2"><%=objEvents(8)%></div>
					</div>
					<div>
						<div Class="CardColumn1">Warranty: </div>
						<div Class="CardColumn2"><input type="checkbox" name="Warranty" value="True" <%=strWarrantyChecked%> /></div>
					</div>
				<% If objEvents(2) <> "" And Not IsNull(objEvents(2)) Then %>
						<div>Event Notes: </div>
						<div><%=objEvents(2)%></div>
				<% End If %>
				</div>

		<% Else %>

				<div class="Card NormalCard">
					<form method="POST" action="<%=strSubmitTo%>">
					<input type="hidden" name="EventID" value="<%=objEvents(0)%>" />
					<div class="CardTitle">Open Event</div>
					<div>
						<div Class="CardColumn1">Event Type: </div>
						<div Class="CardColumn2"><%=objEvents(1)%></div>
					</div>
					<div>
						<div Class="CardColumn1">Submitted: </div>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3)) & " " & ShortenTime(objEvents(4))%></div>
					</div>
					<div Class="CardColumn1">Category:</div>
					<div Class="CardColumn2">
						<select Class="Card" name="Category">
							<option value=""></option>
					<% If Not objCategories.EOF Then
							Do Until objCategories.EOF
								If objCategories(0) = objEvents(8) Then
									strSelected = "selected=""selected"""
								Else
									strSelected = ""
								End If %>
								<option value="<%=objCategories(0)%>" <%=strSelected%>><%=objCategories(0)%></option>
							<% objCategories.MoveNext
							Loop
							objCategories.MoveFirst
						End If %>
						</select>
					</div>
					<div Class="CardColumn1">Warranty:</div>
					<div Class="CardColumn2">
						<input type="checkbox" name="Warranty" value="True" <%=strWarrantyChecked%> />
					</div>
					<div>
						<div Class="CardColumn1">Complete: </div>
						<div Class="CardColumn2">
							<input Class="Card" type="checkbox" value="TRUE" name="Resolved" />
						</div>
					</div>
					<div>Event Notes: </div>
					<div>
						<textarea Class="Card" rows="5" name="Notes" cols="90" style="width: 99%;"><%=objEvents(2)%></textarea>
					</div>
					<div>&nbsp;</div>
					<div Class="Button"><input type="submit" value="Update Event" name="Submit" /></div>
			<% If CInt(intEventID) = CInt(objEvents(0)) Then %>
					<div>
						<div class="Information">Updated</div>
					</div>
			<% End If %>
					</form>
				</div>

      <% End If
      	objEvents.MoveNext
      Loop
   End If %>

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
				<th>Username</th>
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
					<td id="center"><a href="user.asp?Username=<%=objLog(1)%>"><%=objLog(1)%></a></td>

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

<%Sub SearchForDevice%>
	<div class="Card NormalCard">
		<form method="POST" action="index.asp">
		<div class="CardTitle">Search for a Device</div>
		<div>
			<div Class="CardColumn1">Asset tag: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthSmall" type="text" name="Tag" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">BOCES tag: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthSmall" type="text" name="BOCESTag" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Serial #: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="Serial" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Site: </div>
			<div Class="CardColumn2">
				<select Class="Card" name="DeviceSite">
					<option value=""></option>
				<% Do Until objSites.EOF %>
							<option value="<%=objSites(0)%>"><%=objSites(0)%></option>
				<%    objSites.MoveNext
					Loop %>
				</select>
			</div>
		</div>
		<div>
			<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		</div>
		</form>
	</div>
<%End Sub%>

<%Sub JumpToDevice%>

	<div Class="HeaderCard">
		<form method="POST" action="search.asp">
		Asset tag: <input class="Card InputWidthSmall" type="text" name="SmartBox" id="LastNames" />
		<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		<input type="hidden" value="" name="BOCESTag" />
		</form>
	</div>

<%End Sub%>

<%Sub AddEvent%>

   <div Class="Card NormalCard">
      <form method="POST" action="<%=strSubmitTo%>">
      <div Class="CardTitle"><image src="../images/event.png" height="15" width="15" title="Events">&nbsp;Add Event</div>
      <div Class="CardColumn1">Event Type: </div>
      <div Class="CardColumn2">
         <select Class="Card" name="EventType">
            <option value=""></option>
      <% If Not objEventTypes.EOF Then
				Do Until objEventTypes.EOF %>
					<option value="<%=objEventTypes(0)%>"><%=objEventTypes(0)%></option>
				<% objEventTypes.MoveNext
				Loop
			End If %>
         </select>
      </div>
      <div Class="CardColumn1">Category:</div>
      <div Class="CardColumn2">
      	<select Class="Card" name="Category">
            <option value=""></option>
      <% If Not objCategories.EOF Then
				Do Until objCategories.EOF %>
					<option value="<%=objCategories(0)%>"><%=objCategories(0)%></option>
				<% objCategories.MoveNext
				Loop
				objCategories.MoveFirst
			End If %>
         </select>
      </div>
      <div Class="CardColumn1">Warranty:</div>
      <div Class="CardColumn2">
			<input type="checkbox" name="Warranty" value="True" />
      </div>
      <div>
		<div Class="CardColumn1">Complete: </div>
			<div Class="CardColumn2">
				<input Class="Card" type="checkbox" value="TRUE" name="Resolved" />
			</div>
		</div>
      <div>Event Notes:</div>
      <div><textarea Class="Card" rows="5" name="Notes" cols="90" style="width: 99%;"></textarea></div>
      <div>&nbsp;</div>
      <div Class="Button"><input type="submit" value="Add Event" name="Submit" /></div>
   <% If strAddEventMessage <> "" Then %>
   		<div>
   			<%=strAddEventMessage%>
   		</div>
   <% End If %>
      </form>
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

	<% Case "Remote" %>
			<a href="vnc://<%=objDevice(14)%>:5900" class="Button<%=strPosition%>" >
		<% If PINGFlag Then %>
			<image src="../images/remote.png" height="20" width="20" title="<%=status%>">
		<% Else %>
			<image style="opacity:0.5;filter:alpha(opacity=50)" src="../images/remote.png" height="20" width="20" title="<%=status%>">

		<% End If %>
			</a>
			<a href="ssh://admin@<%=objDevice(14)%>" class="Button<%=strPosition%>" >
		<% If PINGFlag Then %>
			<image src="../images/ssh.png" height="20" width="20" title="SSH Online">
		<% Else %>
			<image style="opacity:0.5;filter:alpha(opacity=50)" src="../images/ssh.png" height="20" width="20" title="SSH Offline">

		<% End If %>
			</a>

<%	End Select

End Sub %>

<%Sub AddSubmittedEvent

   Dim strEventType, strNotes, strSQL, strCategory, bolWarranty, objUserID, intUserID
   Dim datDate, datTime, objEventLookup, strUserName, bolResolved

   'Check and see if the device already has an event.
   strSQL = "SELECT ID FROM Events WHERE Resolved=False AND LGTag='" & intTag & "'"
   Set objEventLookup = Application("Connection").Execute(strSQL)
   If Not objEventLookup.EOF Then
   	Exit Sub
   End If

   strEventType = Request.Form("EventType")
   strNotes = Request.Form("Notes")
   strCategory = Replace(Request.Form("Category"),"'","''")
   bolResolved = Request.Form("Resolved")
   datDate = Date()
   datTime = Time()

   If Request.Form("Warranty") = "True" Then
   	bolWarranty = True
   Else
   	bolWarranty = False
   End If

   If strEventType <> "" Then

   	If strCategory <> "" Then

			'Get the userID
			strSQL = "SELECT AssignedTo FROM Assignments WHERE Active=True AND LGTag='" & intTag & "'"
			Set objUserID = Application("Connection").Execute(strSQL)
			If Not objUserID.EOF Then
				intUserID = objUserID(0)
				strUserName = GetUserName(intUserID)
			Else
				If strEventType = "Insurance Claim" Then
					strSQL = "SELECT AssignedTo FROM Assignments WHERE Active=False AND LGTag='" & intTag & "' ORDER BY ID DESC"
					Set objUserID = Application("Connection").Execute(strSQL)
					If Not objUserID.EOF Then
						intUserID = objUserID(0)
						strUserName = GetUserName(intUserID)
					Else
						intUserID = 0
						strUserName = ""
					End If
				Else
					intUserID = 0
					strUserName = ""
				End If
			End If

			If strEventType = "Decommission Device" Then
				strSQL = "INSERT INTO Events (Type,LGTag,Notes,EventDate,EventTime,Category,Site,Model,EnteredBy,Warranty,Resolved,ResolvedDate,ResolvedTime) VALUES ('"
				strSQL = strSQL & Replace(strEventType,"'","''") & "','"
				strSQL = strSQL & intTag & "','"
				strSQL = strSQL & Replace(strNotes,"'","''") & "',#"
				strSQL = strSQL & datDate & "#,#"
				strSQL = strSQL & datTime & "#,'"
				strSQL = strSQL & Replace(strCategory,"'","''") & "','"
				strSQL = strSQL & Replace(objDevice(1),"'","''") & "','"
				strSQL = strSQL & Replace(objDevice(2),"'","''") & "','"
				strSQL = strSQL & strUser & "',"
				strSQL = strSQL & bolWarranty & ","
				strSQL = strSQL & "True,#"
				strSQL = strSQL & datDate & "#,#"
				strSQL = strSQL & datTime & "#)"
				Application("Connection").Execute(strSQL)
				strAddEventMessage = "<div Class=""Information"">Device Decommissioned</div>"

				strSQL = "UPDATE Devices SET Active=False,DateDisabled=#" & Date() & "# WHERE LGTag='" & intTag & "'"
				Application("Connection").Execute(strSQL)

				strSQL = "SELECT ID FROM Events WHERE LGTag='" & intTag & "' AND "
				strSQL = strSQL & "Type='" & Replace(strEventType,"'","''") & "' AND "
				strSQL = strSQL & "Category='" & Replace(strCategory,"'","''") & "' AND "
				strSQL = strSQL & "EventDate=#" & datDate & "# AND "
				strSQL = strSQL & "EventTime=#" & datTime & "#"
				Set objEventLookup = Application("Connection").Execute(strSQL)

				UpdateLog "DeviceDecommissioned",intTag,strUserName,"",strEventType & " - " & strCategory,objEventLookup(0)

			Else

				strSQL = "INSERT INTO Events (Type,LGTag,Notes,EventDate,EventTime,Category,UserID,Site,Model,EnteredBy,Warranty) VALUES ('"
				strSQL = strSQL & Replace(strEventType,"'","''") & "','"
				strSQL = strSQL & intTag & "','"
				strSQL = strSQL & Replace(strNotes,"'","''") & "',#"
				strSQL = strSQL & datDate & "#,#"
				strSQL = strSQL & datTime & "#,'"
				strSQL = strSQL & Replace(strCategory,"'","''") & "',"
				strSQL = strSQL & Int(intUserID) & ",'"
				strSQL = strSQL & Replace(objDevice(1),"'","''") & "','"
				strSQL = strSQL & Replace(objDevice(2),"'","''") & "','"
				strSQL = strSQL & strUser & "',"
				strSQL = strSQL & bolWarranty & ")"
				Application("Connection").Execute(strSQL)

				strAddEventMessage = "<div Class=""Information"">Event Added</div>"

				strSQL = "SELECT ID FROM Events WHERE LGTag='" & intTag & "' AND "
				strSQL = strSQL & "Type='" & Replace(strEventType,"'","''") & "' AND "
				strSQL = strSQL & "Category='" & Replace(strCategory,"'","''") & "' AND "
				strSQL = strSQL & "EventDate=#" & datDate & "# AND "
				strSQL = strSQL & "EventTime=#" & datTime & "#"
				Set objEventLookup = Application("Connection").Execute(strSQL)

				If intUserID > 0 Then
					UpdateLog "EventAdded",intTag,GetUserName(intUserID),"",strEventType & " - " & strCategory,objEventLookup(0)
					strUserName = GetUserName(intUserID)

					If strEventType = "Insurance Claim" Then
						BillUser intUserID,"Insurance Copay",strUserName
					End If
				Else
					UpdateLog "EventAdded",intTag,"","",strEventType & " - " & strCategory,objEventLookup(0)
				End If

				If bolResolved Then
					strSQL = "UPDATE Events SET Resolved=True,ResolvedDate=#" & datDate & "#,ResolvedTime=#" & datTime & "#," & _
					"CompletedBy='" & strUser & "' WHERE ID=" & objEventLookup(0)
					Application("Connection").Execute(strSQL)

					If intUserID > 0 Then
						UpdateLog "EventClosed",intTag,GetUserName(intUserID),"",strEventType & " - " & strCategory,objEventLookup(0)
					Else
						UpdateLog "EventClosed",intTag,"","",strEventType & " - " & strCategory,objEventLookup(0)
					End If

				End If
				
				Application("Connection").Execute("UPDATE Devices SET HasEvent=True WHERE LGTag='" & intTag & "'")

			End If

		Else

			strAddEventMessage = "<div Class=""Error"">Category Missing</div>"

		End If
	Else

		strAddEventMessage = "<div Class=""Error"">Event Type Missing</div>"

	End If

End Sub%>

<% Sub UpdateEvent

	Dim strNotes, datDate, datTime, bolResolved, strSQL, strCategory, bolWarranty, strUserName, strOldEventType
	Dim objUserID, intUserID, objOldValues, strOldNotes, strOldCategory, bolOldWarranty, strEventType

	'Get the userID
	strSQL = "SELECT AssignedTo FROM Assignments WHERE Active=True AND LGTag='" & intTag & "'"
	Set objUserID = Application("Connection").Execute(strSQL)
	If Not objUserID.EOF Then
		intUserID = objUserID(0)
	Else
		intUserID = 0
	End If

	'Get the variables from the form
	intEventID = Request.Form("EventID")
	strNotes = Request.Form("Notes")
	strCategory = Request.Form("Category")
	bolResolved = Request.Form("Resolved")
	strEventType = Request.Form("EventType")

	If Request.Form("Warranty") = "True" Then
   	bolWarranty = True
   Else
   	bolWarranty = False
   End If

	datDate = Date()
	datTime = Time()

	'Get the current values from the database
	strSQL = "SELECT Notes,Category,Type,Warranty FROM Events WHERE ID=" & intEventID
	Set objOldValues = Application("Connection").Execute(strSQL)

	'Record the old values before they change
	strOldNotes = objOldValues(0)
	strOldCategory = objOldValues(1)
	strOldEventType = objOldValues(2)
	bolOldWarranty = objOldValues(3)

	strSQL = "UPDATE Events" & vbCRLF
	strSQL = strSQL & "SET Notes='" & Replace(strNotes,"'","''") & "',"
	strSQL = strSQL & "Category='" & Replace(strCategory,"'","''") & "',"
	strSQL = strSQL & "Type='" & Replace(strEventType,"'","''") & "',"
	strSQL = strSQL & "Warranty=" & bolWarranty

	If bolResolved Then
		strSQL = strSQL & ",Resolved=True,ResolvedDate=#" & datDate & "#,ResolvedTime=#" & datTime & "#," & _
			"CompletedBy='" & strUser & "'"
	
		Application("Connection").Execute("UPDATE Devices SET HasEvent=False WHERE LGTag='" & intTag & "'")
		
	End If

	strSQL = strSQL & vbCRLF & "WHERE ID=" & intEventID
	Application("Connection").Execute(strSQL)

	'Log the updated values
	If intUserID > 0 Then
		strUserName = GetUserName(intUserID)
	Else
		strUserName = ""
	End If

	If strNotes <> strOldNotes Then
		UpdateLog "EventUpdatedNotes",intTag,strUserName,strOldNotes,strNotes,intEventID
	End If
	If strCategory <> strOldCategory Then
		UpdateLog "EventUpdatedCategory",intTag,strUserName,strOldCategory,strCategory,intEventID
	End If
	If strEventType <> strOldEventType Then
		UpdateLog "EventUpdatedType",intTag,strUserName,strOldEventType,strEventType,intEventID
	End If
	If bolWarranty <> bolOldWarranty Then
		UpdateLog "EventUpdatedWarranty",intTag,strUserName,bolOldWarranty,bolWarranty,intEventID
	End If
	If bolResolved Then
		UpdateLog "EventClosed",intTag,strUserName,"","",intEventID
	End If

End Sub %>

<%Sub UpdateDevice

	Dim strSite, strRoom, bolInsured, strTags, strSQL, arrTags, intIndex, strTag, objTagCheck
	Dim bolTagFound, strAppleID, strMACAddress, strNotes, objOldValues, objUserID, intUserID
	Dim strOldSite, strOldRoom, strOldAppleID, bolOldInsured, strOldMACAddress, strOldNotes
	Dim strOldTag, strUserName, strUpdatedAssetTag, strUpdatedMake, strUpdatedModel, strUpdatedSerial
	Dim strUpdatedBOCESTag, strUpdatedPurchased, strOldMake, strOldModel, strOldSerial, strOldBOCESTag
	Dim strOldPurchased, objDeviceCheck, strOldAssetTag, bolAssetTagChanged

	'Get the userID
	strSQL = "SELECT AssignedTo FROM Assignments WHERE Active=True AND LGTag='" & intTag & "'"
	Set objUserID = Application("Connection").Execute(strSQL)
	If Not objUserID.EOF Then
		intUserID = objUserID(0)
	Else
		intUserID = ""
	End If

	bolAssetTagChanged = False

	'Get the variables from the form
	strSite = Request.Form("Site")
	strRoom = Request.Form("Room")
	'strAppleID = Request.Form("AppleID")
	bolInsured = Request.Form("Insured")
	strTags = Request.Form("Tags")
	'strMACAddress = Request.Form("MACAddress")
	strNotes = Request.Form("Notes")

	'Get the variables from the form when in edit mode
	strUpdatedAssetTag = Request.Form("AssetTag")
	strUpdatedMake = Request.Form("Make")
	strUpdatedModel = Request.Form("Model")
	strUpdatedSerial = Request.Form("serialNumber")
	strUpdatedBOCESTag = Request.Form("BOCESTag")
	strUpdatedPurchased = Request.Form("Purchased")
	If Not bolInsured Then
		bolInsured = False
	End If

	'Make sure the MACAddress is in the right format
' 	If strMACAddress <> "" Then
' 		If Not ValidMACAddress(strMACAddress) Then
' 			strLocationMessage = "<div Class=""Error"">Invalid MAC Address</div>"
' 			Exit Sub
' 		End If
' 	End If

	'Make sure the AppleID is in the right format
' 	If strAppleID <> "" Then
' 		If Not ValidEMailAddress(strAppleID) Then
' 			strLocationMessage = "<div Class=""Error"">Invalid Apple ID</div>"
' 			Exit Sub
' 		End If
' 	End If

	'Make sure the purchased date is in the right format
	If strUpdatedPurchased <> "" Then
		If Not IsDate(strUpdatedPurchased) Then
			strLocationMessage = "<div Class=""Error"">Invalid Date Purchased</div>"
			Exit Sub
		End If
	End If

	'Get the current values from the database
	strSQL = "SELECT Site,Room,AppleID,MACAddress,Notes,HasInsurance,Manufacturer,Model,SerialNumber,BOCESTag,DatePurchased FROM Devices WHERE LGTag='" & intTag & "'"
	Set objOldValues = Application("Connection").Execute(strSQL)

	'Record the old values before they change
	strOldSite = objOldValues(0)
	strOldRoom = objOldValues(1)
	strOldAppleID = objOldValues(2)
	strOldMACAddress = objOldValues(3)
	strOldNotes = objOldValues(4)
	bolOldInsured = objOldValues(5)
	strOldMake = objOldValues(6)
	strOldModel = objOldValues(7)
	strOldSerial = objOldValues(8)
	strOldBOCESTag = objOldValues(9)
	strOldPurchased = objOldValues(10)
	strOldAssetTag = intTag

	'Fix the old values if they were null
	If IsNull(strOldSite) Then
		strOldSite = ""
	End If
	If IsNull(strOldRoom) Then
		strOldRoom = ""
	End If
	If IsNull(strOldAppleID) Then
		strOldAppleID = ""
	End If
	If IsNull(strOldMACAddress) Then
		strOldMACAddress = ""
	End If
	If IsNull(strOldNotes) Then
		strOldNotes = ""
	End If
	If IsNull(strOldMake) Then
		strOldMake = ""
	End If
	If IsNull(strOldModel) Then
		strOldModel = ""
	End If
	If IsNull(strOldSerial) Then
		strOldSerial = ""
	End If
	If IsNull(strOldBOCESTag) Then
		strOldBOCESTag = ""
	End If
	If IsNull(strOldPurchased) Then
		strOldPurchased = ""
	End If

	'Make sure the asset tag doesn't already exist
	If CStr(strUpdatedAssetTag) <> CStr(strOldAssetTag) Then
		bolAssetTagChanged = True
		strSQL = "SELECT ID FROM Devices WHERE LGTag='" & strUpdatedAssetTag & "'"
		Set objDeviceCheck = Application("Connection").Execute(strSQL)
		If Not objDeviceCheck.EOF Then
			strLocationMessage = "<div Class=""Error"">Duplicate Asset Tag</div>"
			Exit Sub
		End If
	End If

	'Make sure the BOCES tag doesn't already exist
	If strUpdatedBOCESTag <> strOldBOCESTag Then
		strSQL = "SELECT ID FROM Devices WHERE BOCESTag='" & strUpdatedBOCESTag & "'"
		Set objDeviceCheck = Application("Connection").Execute(strSQL)
		If Not objDeviceCheck.EOF Then
			strLocationMessage = "<div Class=""Error"">Duplicate BOCES Tag</div>"
			Exit Sub
		End If
	End If

	'Update the values in the database
	strSQL = "UPDATE Devices Set "
	strSQL = strSQL & "Site='" & Replace(strSite,"'","''") & "',"
	strSQL = strSQL & "Room='" & Replace(strRoom,"'","''") & "',"
	'strSQL = strSQL & "AppleID='" & Replace(strAppleID,"'","''") & "',"
	'strSQL = strSQL & "MACAddress='" & Replace(strMACAddress,"'","''") & "',"
	strSQL = strSQL & "Notes='" & Replace(strNotes,"'","''") & "',"
	strSQL = strSQL & "Manufacturer='" & Replace(strUpdatedMake,"'","''") & "',"
	strSQL = strSQL & "Model='" & Replace(strUpdatedModel,"'","''") & "',"
	strSQL = strSQL & "SerialNumber='" & Replace(strUpdatedSerial,"'","''") & "',"
	strSQL = strSQL & "BOCESTag='" & Replace(strUpdatedBOCESTag,"'","''") & "',"
	strSQL = strSQL & "DatePurchased=#" & Replace(strUpdatedPurchased,"'","''") & "#,"
	strSQL = strSQL & "HasInsurance=" & bolInsured & vbCRLF
	strSQL = strSQL & "WHERE LGTag='" & intTag & "'"
	Application("Connection").Execute(strSQL)

	If bolAssetTagChanged Then
		strSQL = "UPDATE DEVICES SET LGTag='" & strUpdatedAssetTag & "' WHERE LGTag='" & strOldAssetTag & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Assignments SET LGTag='" & strUpdatedAssetTag & "' WHERE LGTag='" & strOldAssetTag & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Events SET LGTag='" & strUpdatedAssetTag & "' WHERE LGTag='" & strOldAssetTag & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Log SET LGTag='" & strUpdatedAssetTag & "' WHERE LGTag='" & strOldAssetTag & "'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Tags SET LGTag='" & strUpdatedAssetTag & "' WHERE LGTag='" & strOldAssetTag & "'"
		Application("Connection").Execute(strSQL)
	End If

	'Get the username
	strUserName = GetUserName(intUserID)

	'Log the updated values
	If bolAssetTagChanged Then
		UpdateLog "DeviceUpdatedAssetTag",strUpdatedAssetTag,strUserName,strOldAssetTag,strUpdatedAssetTag,""
		intTag = strUpdatedAssetTag
	End If
	If strSite <> strOldSite Then
		UpdateLog "DeviceUpdatedSite",intTag,strUserName,strOldSite,strSite,""
	End If
	If strRoom <> strOldRoom Then
		UpdateLog "DeviceUpdatedRoom",intTag,strUserName,strOldRoom,strRoom,""
	End If
' 	If strAppleID <> strOldAppleID Then
' 		UpdateLog "DeviceUpdatedAppleID",intTag,strUserName,strOldAppleID,strAppleID,""
' 	End If
' 	If strMACAddress <> strOldMACAddress Then
' 		UpdateLog "DeviceUpdatedMACAddress",intTag,strUserName,strOldMACAddress,strMACAddress,""
' 	End If
	If strNotes <> strOldNotes Then
		UpdateLog "DeviceUpdatedNotes",intTag,strUserName,strOldNotes,strNotes,""
	End If
	If strUpdatedMake <> strOldMake Then
		UpdateLog "DeviceUpdatedMake",intTag,strUserName,strOldMake,strUpdatedMake,""
	End If
	If strUpdatedModel <> strOldModel Then
		UpdateLog "DeviceUpdatedModel",intTag,strUserName,strOldModel,strUpdatedModel,""
	End If
	If strUpdatedSerial <> strOldSerial Then
		UpdateLog "DeviceUpdatedSerial",intTag,strUserName,strOldSerial,strUpdatedSerial,""
	End If
	If strUpdatedBOCESTag <> strOldBOCESTag Then
		UpdateLog "DeviceUpdatedBOCESTag",intTag,strUserName,strOldBOCESTag,strUpdatedBOCESTag,""
	End If
	If cDate(strUpdatedPurchased) <> cDate(strOldPurchased) Then
		UpdateLog "DeviceUpdatedPurchased",intTag,strUserName,strOldPurchased,strUpdatedPurchased,""
	End If
	If CStr(bolInsured) <> CStr(bolOldInsured) Then
		UpdateLog "DeviceUpdatedInsurance",intTag,strUserName,bolOldInsured,bolInsured,""
	End If

	arrTags = Split(strTags,",")
	For intIndex = 0 to UBound(arrTags)
		strTag = Trim(arrTags(intIndex))

		If strTag <> "" Then

			strSQL = "SELECT Tag FROM Tags WHERE Tag='" & Replace(strTag,"'","''") & "' AND LGTag='" & intTag & "'"
			Set objTagCheck = Application("Connection").Execute(strSQL)

			If objTagCheck.EOF Then
				strSQL = "INSERT INTO Tags (LGTag, Tag) VALUES ('" & intTag & "','" & Replace(strTag,"'","''") & "')"
				Application("Connection").Execute(strSQL)

				UpdateLog "DeviceUpdatedTagAdded",intTag,GetUserName(intUserID),"",strTag,""

			End If

		End If

		Set objTagCheck = Nothing

	Next

	strSQL = "SELECT ID, Tag FROM Tags WHERE LGTag='" & intTag & "'"
	Set objTagCheck = Application("Connection").Execute(strSQL)

	If Not objTagCheck.EOF Then

		Do Until objTagCheck.EOF
			bolTagFound = False
			For intIndex = 0 to UBound(arrTags)
				strTag = Trim(arrTags(intIndex))

				If strTag <> "" Then
					If strTag = objTagCheck(1) Then
						bolTagFound = True
					End If
				End If
			Next

			If Not bolTagFound Then
				strOldTag = objTagCheck(1)
				strSQL = "DELETE FROM Tags WHERE ID=" & objTagCheck(0)
				Application("Connection").Execute(strSQL)

				UpdateLog "DeviceUpdatedTagDeleted",intTag,GetUserName(intUserID),strOldTag,"",""

			End If
			objTagCheck.MoveNext
		Loop
	End If

	strLocationMessage = "<div Class=""Information"">Updated</div>"

	If bolAssetTagChanged Then
		Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?Tag=" & strUpdatedAssetTag
	End If

End Sub%>

<%Sub RestoreDevice

	Dim strSQL

	'Enable the device
	strSQL = "UPDATE Devices SET Active=True WHERE LGTag='" & intTag & "'"
	Application("Connection").Execute(strSQL)

	'Update the log
	UpdateLog "DeviceRestored",intTag,"","Disabled","Enabled",""

End Sub%>

<%Sub DisableDevice

	Dim strSQL

	'Disable the device
	strSQL = "UPDATE Devices SET Active=False, DateDisabled=Date() WHERE LGTag='" & intTag & "'"
	Application("Connection").Execute(strSQL)

	'Update the log
	UpdateLog "DeviceDisabled",intTag,"","Enabled","Disabled",""

End Sub%>

<%Sub AssignDevice

   Dim intStudent, bolInsurance, strSQL, objDeviceCheck, objAssignmentCheck, objAssignedTo
   Dim objDeviceCount, intDeviceCount

   'Grade the data from the form
   intStudent = Request.Form("StudentID")
   bolInsurance = False

   'Make sure they submitted something
   If intStudent = "" Or intTag = "" Then
      strNewAssignmentMessage = "<div Class=""Error"">Missing Data</div>"
   Else

      'Check and see if the tag is in the database
      strSQL = "SELECT ID FROM Devices WHERE LGTag='" & intTag & "'"
      Set objDeviceCheck = Application("Connection").Execute(strSQL)

      If Not objDeviceCheck.EOF Then

         'Check and see if the device is already assigned
         strSQL = "SELECT AssignedTo FROM Assignments WHERE LGTag='" & intTag & "' And Active=True"
         Set objAssignmentCheck = Application("Connection").Execute(strSQL)


         If objAssignmentCheck.EOF Then

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

            'Find out who the device is already assigned to
            strSQL = "SELECT FirstName, LastName" & vbCRLF
            strSQL = strSQL & "FROM People" &vbCRLF
            strSQL = strSQL & "WHERE ID=" & objAssignmentCheck(0)
            Set objAssignedTo = Application("Connection").Execute(strSQL)

            strNewAssignmentMessage = "<div Class=""Error"">Device already assigned to " & objAssignedTo(0) & " " & objAssignedTo(1) & "</div>"

         End If

      Else
         strNewAssignmentMessage = "<div Class=""Error"">Device not found</div>"
      End If

   End If

End Sub%>

<%Sub ReturnDevice

   Dim strSerial, strNotes, strSQL, objDeviceCheck, objAssignmentCheck, bolDamaged
   Dim objTagCheck

   'Get the serial number and the notes
   strSerial = Request.Form("Serial")
   strNotes = Request.Form("Notes")
   bolAdapterReturned = Request.Form("Adapter")
   bolCaseReturned = Request.Form("Case")
   bolDamaged = Request.Form("Damaged")

   If Not bolAdapterReturned Then
   	bolAdapterReturned = False
   End If

   If Not bolCaseReturned Then
   	bolCaseReturned = False
   End If

   If Not bolDamaged Then
   	bolDamaged = False
   End If

   'Make sure they submitted something
   If strSerial = "" And intTag = "" Then
      strNewAssignmentMessage = "<div Class=""Error"">Missing Data</div>"
   Else

      'If they entered something in both fields we'll only use the tag
      If intTag <> "" Then

         'Check and see if the tag is in the database
         strSQL = "SELECT ID FROM Devices WHERE LGTag='" & intTag & "'"
         Set objDeviceCheck = Application("Connection").Execute(strSQL)

         If Not objDeviceCheck.EOF Then

            'Check and see if the device is assigned
            strSQL = "SELECT ID FROM Assignments WHERE LGTag='" & intTag & "' And Active=True"
            Set objAssignmentCheck = Application("Connection").Execute(strSQL)

            If Not objAssignmentCheck.EOF Then

               'Return the device
               UpdateDB objAssignmentCheck(0), strNotes

               'Add the "Returned Damaged" Tag if it's damaged.
               If bolDamaged Then
               	strSQL = "SELECT ID FROM Tags WHERE Tag='Returned Damaged' AND LGTag='" & intTag & "'"
               	Set objTagCheck = Application("Connection").Execute(strSQL)

               	If objTagCheck.EOF Then
               		strSQL = "INSERT INTO Tags (LGTag, Tag) VALUES ('" & intTag & "','Returned Damaged')"
							Application("Connection").Execute(strSQL)
               	End If

               End If

               'Set the message to return
               'strMessage = "<div Class=""Bold "">" & intTag & " is no longer assigned to " & strAssignedTo & "</div>"
               strOldAssignmentMessage = "<div Class=""Information"">Device Returned</div>"

            Else

               'strMessage = "<div Class=""Error"">" & intTag & " is not currently assigned to anyone.</div>"

            End If

         Else
            strNewAssignmentMessage = "<div Class=""Error"">Device not found</div>"
         End If

      Else

         'Check and see if the serial is in the database
         strSQL = "SELECT ID FROM Devices WHERE SerialNumber='" & strSerial & "'"
         Set objDeviceCheck = Application("Connection").Execute(strSQL)

         If Not objDeviceCheck.EOF Then

            strSQL = "SELECT Assignments.ID" & vbCRLF
            strSQL = strSQL & "FROM Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag" & vbCRLF
            strSQL = strSQL & "WHERE SerialNumber='" & strSerial & "' AND Assignments.Active=True"
            Set objAssignmentCheck = Application("Connection").Execute(strSQL)

            If Not objAssignmentCheck.EOF Then

               'Return the device
               UpdateDB objAssignmentCheck(0), strNotes

               'Set the message to return
               'strMessage = "<div Class=""Bold "">" & strSerial & " is no longer assigned to " & strAssignedTo
               strOldAssignmentMessage = "<div Class=""Information"">Device Returned</div>"

            Else

               'strMessage = "<div Class=""Error "">""A device with serial " & strSerial & " is not assigned to anyone."

            End If

         Else

            strNewAssignmentMessage = "<div Class=""Error "">""Device not found</div>"

         End If

      End If

   End If

End Sub%>

<%Sub UpdateDB(intID, strNotes)

   Dim strSQL, objAssignmentInfo, objAssignedTo, objDeviceCount, intDeviceCount, strUserName

   'Update the assignment
   strSQL = "UPDATE Assignments SET Active=False,DateReturned=#" & Date & "#,TimeReturned=#"
   strSQL = strSQL & Time & "#,ReturnedBy='" & strUser & "',Notes='" & Replace(strNotes,"'","''") & "'" & vbCRLF
   strSQL = strSQL & "WHERE ID=" & intID
   Application("Connection").Execute(strSQL)

   'Get the ID for the person and device
   strSQL = "SELECT AssignedTo, LGTag FROM Assignments WHERE ID=" & intID
   Set objAssignmentInfo = Application("Connection").Execute(strSQL)

  	'Get the current number of devices the user has
	strSQL = "SELECT DeviceCount FROM People WHERE ID=" & objAssignmentInfo(0)
	Set objDeviceCount = Application("Connection").Execute(strSQL)
	If IsNull(objDeviceCount(0)) Then
		intDeviceCount = 0
	Else
		intDeviceCount = objDeviceCount(0) - 1
	End If

   'Update the person's device status
   strSQL = "UPDATE People" & vbCRLF
   If intDeviceCount = 0 Then
   	strSQL = strSQL & "SET HasDevice=False, DeviceCount=" & intDeviceCount & vbCRLF
   Else
   	strSQL = strSQL & "SET DeviceCount=" & intDeviceCount & vbCRLF
   End If
   strSQL = strSQL & "WHERE ID=" & objAssignmentInfo(0)
   Application("Connection").Execute(strSQL)

   'If they forgot their case or bag then turn on the warning flag
   If Not bolAdapterReturned Or Not bolCaseReturned Then
   	strSQL = "UPDATE People SET Warning=True Where ID=" & objAssignmentInfo(0)
   	Application("Connection").Execute(strSQL)
   End If

   'Update the device to show it's not assigned
   strSQL = "UPDATE Devices SET "
   strSQL = strSQL & "Assigned=False,"
   strSQL = strSQL & "FirstName='',LastName='',UserName=''" & vbCRLF
   strSQL = strSQL & "WHERE LGTag='" & objAssignmentInfo(1) & "'"
   Application("Connection").Execute(strSQL)

   'Get the person's name from the database
   strSQL = "SELECT FirstName, LastName, UserName" & vbCRLF
   strSQL = strSQL & "FROM People" &vbCRLF
   strSQL = strSQL & "WHERE ID=" & objAssignmentInfo(0)
   'Set objAssignedTo = Application("Connection").Execute(strSQL)

   strUserName = GetUserName(objAssignmentInfo(0))

   If Not bolAdapterReturned Then
   	If InStr(objDevice(2),"MacBook") Then
   		BillUser objAssignmentInfo(0),objDevice(2) & " Charger",strUserName
   	End If
   	If InStr(objDevice(2),"iPad") Then
   		BillUser objAssignmentInfo(0),"iPad Charger",strUserName
   		BillUser objAssignmentInfo(0),"iPad Cable",strUserName
   	End If
   	UpdateLog "DeviceReturnedAdapterMissing",intTag,strUserName,GetDisplayName(strUserName),"",""
   End If

   If Not bolCaseReturned Then
   	If InStr(objDevice(2),"MacBook") Then
   		BillUser objAssignmentInfo(0),"Laptop Case",strUserName
   	End If
   	UpdateLog "DeviceReturnedCaseMissing",intTag,strUserName,GetDisplayName(strUserName),"",""
   End If

   If Request.Form("Damaged") Then
   	UpdateLog "DeviceReturnedDamaged",intTag,strUserName,GetDisplayName(strUserName),strNotes,""
   Else
		UpdateLog "DeviceReturned",intTag,strUserName,GetDisplayName(strUserName),strNotes,""
	End If

   'strAssignedTo = objAssignedTo(0) & " " & objAssignedTo(1)

End Sub%>

<%Sub BillUser(intUserID,strItem,strUserName)

	Dim strSQL, objPrice,intLoanID, objLoanedItem, bolReturnable

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

		EMailBusinessOffice intUserID,"BillUser",strItem,objPrice(0)

	End If

End Sub%>

<%Sub EMailBusinessOffice(intUserID,strType,strItem,intPrice)

	Const cdoSendUsingPickup = 1

   Dim strSMTPPickupFolder, strFrom, objMessage, objConf, strMessage, strBCC, objUser
   Dim strSQL, strUserName, objUserID, strSubject, strUser, objNetwork

	strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"
	strFrom = Application("EMailNotifications")

	strSQL = "SELECT ID,FirstName,LastName,UserName FROM People WHERE ID=" & intUserID & " AND NOT Deleted"
	Set objUser = Application("Connection").Execute(strSQL)

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

	strMessage =  "<p>This is an automated email from the inventory system.</p>"

	Select Case strType
		Case "BillUser"
			strSubject = "User Owes $" & intPrice & " for " & strItem

			strMessage = strMessage & "<p><a href=""http://" & Request.ServerVariables("SERVER_NAME") & Replace(Request.ServerVariables("URL"),"device","user") & "?UserName=" & objUser(3) & """>" & _
				objUser(1) & " " & objUser(2) & "</a> has been added to the system as owing $" & intPrice & " for " & strItem & ".  " & _
				"Please send them an invoice.</p></br />"

		Case "ItemReturned"
			strSubject = "Equipment has been Returned"

			strMessage = strMessage & "<p><a href=""http://" & Request.ServerVariables("SERVER_NAME") & Replace(Request.ServerVariables("URL"),"device","user") & "?UserName=" & objUser(3) & """>" & _
				objUser(1) & " " & objUser(2) & "</a> has has returned " & strItem & " so now they don't owe $" & intPrice & ".  " & _
				"Please update your records.</p></br />"

	End Select

	strMessage = strMessage & "<p>Thank you <br />"
	strMessage = strMessage & "Please do not respond to this message...</p>"

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

	Set objMessage = Nothing
	Set objConf = Nothing

End Sub%>

<%Sub EMailTeachers

	Const cdoSendUsingPickup = 1

   Dim strSMTPPickupFolder, strFrom, objMessage, objConf, strMessage, strBCC
   Dim objUserInfo, strSQL, intUserID, strUserName, objUserID

   strSQL = "SELECT FirstName,LastName" & vbCRLF
   strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
   strSQL = strSQL & "WHERE Assignments.Active=True and LGTag='" & intTag & "'"
   Set objUserInfo = Application("Connection").Execute(strSQL)

   If NOT objUserInfo.EOF Then

		strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"
		strFrom = "helpdesk@lkgeorge.org"
		strBCC = Application("EMailNotifications") & ";" & strUser & "@lkgeorge.org"

		'Create the objects required to send the mail.
		Set objMessage = CreateObject("CDO.Message")
		Set objConf = objMessage.Configuration
		With objConf.Fields
			.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
			.item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
			.Update
		End With

		strMessage =  "<p>This is an automated email from the Help Desk.</p>"

		strMessage = strMessage & "<p>" & objUserInfo(0) & " " & objUserInfo(1) & " has been assigned device " & intTag & ".</p>"

		strMessage = strMessage & "<p>Thank you <br />"
		strMessage = strMessage & "Please do not respond to this message...</p>"

		With objMessage
			.To = strFrom
			.From = strFrom
			.Subject = "Student Computer Assignment"
			.HTMLBody = strMessage
			If strBCC <> "" Then
				.BCC = strBCC
			End If
		  .Send
		End With

		Set objMessage = Nothing
		Set objConf = Nothing

		strNewAssignmentMessage = "<div class=""Information"">Message Sent</div>"

		'Get the userID
		strSQL = "SELECT AssignedTo FROM Assignments WHERE Active=True AND LGTag='" & intTag & "'"
		Set objUserID = Application("Connection").Execute(strSQL)
		If Not objUserID.EOF Then
			intUserID = objUserID(0)
		Else
			intUserID = 0
		End If

		strUserName = GetUserName(intUserID)

		UpdateLog "EmailedAssignmentToTeachers",intTag,strUserName,"","",""

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
		GetUserName = ""
	End If

End Function%>

<%Function GetDisplayName(Username)

	Dim strSQL, objUserInfo

	If Username <> "" Then

		strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & Replace(UserName,"'","''") & "'"
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
	If InStr(strType,"Notes") > 0 Or InStr(strType,"DeviceReturned") > 0 Then
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
