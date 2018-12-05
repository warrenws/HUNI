<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 2/16/16
'Last Updated 1/14/18

'This page upgrades inventory database

'Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser, objReports, strReport, strSubmitTo, strColumns
Dim strUpgradeMessage

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
		Case "Upgrade Database"
         UpgradeDatabase
         UpdateLog "DatabaseUpgraded","","","","0.060",""
   End Select
   
   'Get the URL used to submit forms
   If Request.ServerVariables("QUERY_STRING") = "" Then
      strSubmitTo = "dbupgrade.asp"
   Else   
      strSubmitTo = "dbupgrade.asp?" & Request.ServerVariables("QUERY_STRING")
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
   
   	<script>
   		$(function() {
   		
   		<%	If Not IsMobile And Not IsiPad Then%>
					$( document ).tooltip({track: true});
			<% End If %>
			})
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
			<form method="POST" action="<%=strSubmitTo%>">
				<div Class="HeaderCard">
					<%=strUpgradeMessage%>
					<input Class="Button" type="submit" value="Upgrade Database" name="Submit" />
				</div>
			</form>
		</div> 
		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
   </body>

   </html>

<%End Sub%>


<%Sub UpgradeDatabase 

	Dim strSQL
	
	'Grab the catalog from the database.
	Set objCatalog = CreateObject("ADOX.Catalog")
   objCatalog.ActiveConnection = Application("Connection")
   
   'Check in see if the new tables are present
	
	bolLogTableFound = False
	bolPendingTasksFound = False
	bolHistoryFound = False
	bolOwedFound = False
	bolPurchasableFound = False
	bolGroupMappings = False
	bolCheckInHistoryFound = False
	bolLibreNMSFound = False
	bolInternetTypes = False
   For Each Table in objCatalog.Tables
   	Select Case LCase(Table.Name)
   		Case "log"
   			bolLogTableFound = True
   		Case "pendingtasks"
   			bolPendingTasksFound = True
   		Case "counthistory"
   			bolCountHistoryFound = True
   		Case "owed"
      		bolOwedFound = True
			Case "purchasable"
				bolPurchasableFound = True
			Case "groupmappings"
				bolGroupMappings = True
			Case "checkinhistory"
				bolCheckInHistoryFound = True
			Case "librenms"
				bolLibreNMSFound = True
			Case "parents"
				bolParentsFound = True
			Case "internettypes"
				bolInternetTypes = True
   	End Select
   Next
   
   '***********************************************************************************
   
	'Check the People table for required fields.
	bolHomeRoomEmailFound = False
	bolAUPFound = False
	bolPWordFound = False
	bolNotesFound = False
	bolDeleted = False
	bolDateAdded = False
	bolDateDisabled = False
	bolDateDeleted = False
	bolPWordLastSet = False
	bolSexFound = False
	bolBirthdayFound = False
	bolPendingFound = False
	bolPhoneNumberFound = False
	bolRoomNumberFound = False
	bolDescriptionFound = False
	bolPWordNeverExpires = False
	bolLastExternalCheckIn = False
	bolLastInternalCheckIn = False
	bolInternetAccessFound = False
	
	Set objPeopleTable = objCatalog.Tables("People")
	For Each Column in objPeopleTable.Columns
      Select Case LCase(Column.Name)
         Case "homeroomemail"
            bolHomeRoomEmailFound = True
         Case "aup"
            bolAUPFound = True
         Case "pword"
            bolPWordFound = True
         Case "notes"
            bolNotesFound = True
         Case "deleted"
         	bolDeleted = True
         Case "dateadded"
         	bolDateAdded = True
         Case "datedisabled"
         	bolDateDisabled = True
         Case "datedeleted"
         	bolDateDeleted = True
         Case "pwordlastset"
         	bolPWordLastSet = True
         Case "sex"
         	bolSexFound = True
         Case "birthday"
         	bolBirthdayFound = True
         Case "pending"
         	bolPendingFound = True
         Case "phonenumber"
         	bolPhoneNumberFound = True
         Case "roomnumber"
         	bolRoomNumberFound = True
         Case "description"
         	bolDescriptionFound = True
         Case "pwordneverexpires"
         	bolPWordNeverExpires = True
         Case "lastexternalcheckin"
         	bolLastExternalCheckIn = True
         Case "lastinternalcheckin"
         	bolLastInternalCheckIn = True
         Case "internetaccess"
         	bolInternetAccessFound = True
      End Select
   Next
	
	'Add the needed columns to the database
   If NOT bolHomeRoomEmailFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add HomeRoomEmail TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolAUPFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add AUP BIT"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolPWordFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add PWord TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolNotesFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add Notes LONGTEXT WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
	If NOT bolDeleted Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add Deleted BIT"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDateAdded Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add DateAdded DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDateDisabled Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add DateDisabled DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDateDeleted Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add DateDeleted DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolPWordLastSet Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add PWordLastSet DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolSexFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add Sex TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolBirthdayFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add Birthday DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolPendingFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add Pending BIT"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolPhoneNumberFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add PhoneNumber TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolRoomNumberFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add RoomNumber TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDescriptionFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add Description TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolPWordNeverExpires Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add PWordNeverExpires BIT"
      Application("Connection").Execute(strSQL)
   End If
   If Not bolLastExternalCheckIn Then
   	strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add LastExternalCheckIn DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If Not bolLastInternalCheckIn Then
   	strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add LastInternalCheckIn DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolInternetAccessFound Then
      strSQL = "ALTER TABLE People" & vbCRLF
      strSQL = strSQL & "Add InternetAccess TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   
   
   '***********************************************************************************
   
   'Check the Devices table for required fields.
	bolMACAddressFound = False
	bolAppleIDFound = False
	bolNotesFound = False
	bolDeleted = False
	bolDateAdded = False
	bolDateDisabled = False
	bolDateDeleted = False
	bolDeviceTypeFound = False
	bolInternalIPFound = False
	bolExternalIPFound = False
	bolLastUser = False
	bolOSVersion = False
	bolLastCheckInDate = False
	bolLastCheckInTime = False
	bolHasEvent = False
	Set objDevicesTable = objCatalog.Tables("Devices")
	For Each Column in objDevicesTable.Columns
      Select Case LCase(Column.Name)
         Case "macaddress"
            bolMACAddressFound = True
         Case "appleid"
            bolAppleIDFound = True
         Case "notes"
            bolNotesFound = True
         Case "deleted"
         	bolDeleted = True
         Case "dateadded"
         	bolDateAdded = True
         Case "datedisabled"
         	bolDateDisabled = True
         Case "datedeleted"
         	bolDateDeleted = True
         Case "devicetype"
         	bolDeviceType = True
         Case "internalip"
         	bolInternalIPFound = True
         Case "externalip"
         	bolExternalIPFound = True
         Case "lastuser"
         	bolLastUser = True
         Case "osversion"
         	bolOSVersion = True
         Case "lastcheckindate"
         	bolLastCheckInDate = True
         Case "lastcheckintime"
         	bolLastCheckInTime = True
	      Case "hasevent"
	      	bolHasEvent = True
	      End Select
   Next
	
	'Add the needed columns to the database
	If NOT bolMACAddressFound Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add MACAddress TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolAppleIDFound Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add AppleID TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolNotesFound Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add Notes LONGTEXT WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDeleted Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add Deleted BIT"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDateAdded Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add DateAdded DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDateDisabled Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add DateDisabled DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDateDeleted Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add DateDeleted DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolDeviceType Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add DeviceType TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolInternalIPFound Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add InternalIP TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolExternalIPFound Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add ExternalIP TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolLastUser Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add LastUser TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolOSVersion Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add OSVersion TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolLastCheckInDate Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add LastCheckInDate DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolLastCheckInTime Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add LastCheckInTime DATETIME"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolHasEvent Then
      strSQL = "ALTER TABLE Devices" & vbCRLF
      strSQL = strSQL & "Add HasEvent BIT"
      Application("Connection").Execute(strSQL)
      
      strSQL = "SELECT LGTag FROM Events WHERE Resolved=False"
      Set objDevicesWithOpenEvents = Application("Connection").Execute(strSQL)
      
      If Not objDevicesWithOpenEvents.EOF Then
      	Do Until objDevicesWithOpenEvents.EOF
      		strSQL = "UPDATE Devices SET HasEvent=True WHERE LGTag='" & objDevicesWithOpenEvents(0) & "'"
      		Application("Connection").Execute(strSQL)
      		objDevicesWithOpenEvents.MoveNext
      	Loop
      End If
      
   End If
	
	strSQL = "UPDATE EventTypes SET Active=False WHERE EventType='Notes'"
	Application("Connection").Execute(strSQL)
	
   strUpgradeMessage = "Database Upgraded"

	'***********************************************************************************

	'Check the Events table for required fields.
	bolEnteredByFound = False
	bolSiteFound = False
	bolModelFound = False
	bolCompletedBy = False
	Set objDevicesTable = objCatalog.Tables("Events")
	For Each Column in objDevicesTable.Columns
      Select Case LCase(Column.Name)
         Case "enteredby"
            bolEnteredByFound = True
         Case "site"
            bolSiteFound = True
         Case "model"
            bolModelFound = True
         Case "completedby"
         	bolCompletedBy = True
      End Select
   Next
   
   'Add the needed columns to the database
	If NOT bolEnteredByFound Then
      strSQL = "ALTER TABLE Events" & vbCRLF
      strSQL = strSQL & "Add EnteredBy TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolSiteFound Then
      strSQL = "ALTER TABLE Events" & vbCRLF
      strSQL = strSQL & "Add Site TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolModelFound Then
      strSQL = "ALTER TABLE Events" & vbCRLF
      strSQL = strSQL & "Add Model TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   If NOT bolCompletedBy Then
      strSQL = "ALTER TABLE Events" & vbCRLF
      strSQL = strSQL & "Add CompletedBy TEXT(255) WITH COMPRESSION"
      Application("Connection").Execute(strSQL)
   End If
   
   '***********************************************************************************
   
   'Fix the Roles table
   bolRoleIDFound = False
	Set objDevicesTable = objCatalog.Tables("Roles")
	For Each Column in objDevicesTable.Columns
      Select Case LCase(Column.Name)
         Case "roleid"
            bolRoleIDFound = True
      End Select
   Next
   
   'Check and see of the roles are still split into sites.  If so rebuild the roles
   'table and fix all the existing accounts
   If bolRoleIDFound Then
   	strSQL = "SELECT ID FROM Roles WHERE RoleID=110"
   	Set objRoleTest = Application("Connection").Execute(strSQL)
   	
   	If Not objRoleTest.EOF Then
   		bolRoleIDFound = False
   		
   		strSQL = "UPDATE People SET ClassOf=20 WHERE ClassOf=50"
   		'Application("Connection").Execute(strSQL)
   		
   		strSQL = "UPDATE People SET ClassOf=30 WHERE ClassOf=60"
   		'Application("Connection").Execute(strSQL)
   		
   		strSQL = "UPDATE People SET ClassOf=40 WHERE ClassOf=70"
   		'Application("Connection").Execute(strSQL)
   		
   		strSQL = "UPDATE People SET ClassOf=50 WHERE ClassOf=80"
   		'Application("Connection").Execute(strSQL)
   		
   		strSQL = "UPDATE People SET ClassOf=50 WHERE ClassOf=90"
   		'Application("Connection").Execute(strSQL)
   		
   		strSQL = "UPDATE People SET ClassOf=60 WHERE ClassOf=100"
   		'Application("Connection").Execute(strSQL)
   		
   		strSQL = "UPDATE People SET ClassOf=60 WHERE ClassOf=110"
   		'Application("Connection").Execute(strSQL)
   		
   	End If
   	Set objRoleTest = Nothing
   End If 
   
   'Rebuild the Roles table
   If Not bolRoleIDFound Then
   	strSQL = "DROP TABLE Roles"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "CREATE TABLE Roles ("
   	strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
   	strSQL = strSQL & "Role TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "RoleID INTEGER,"
   	strSQL = strSQL & "Active BIT)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Tech Staff',10,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Teachers',20,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Staff',30,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('TAs',40,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Long Term Subs',50,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Student Teachers',60,True)"
   	Application("Connection").Execute(strSQL)
   	   	
   End If
   
   'Add the latest roles if needed.
   strSQL = "SELECT ID FROM Roles WHERE RoleID=120"
   Set objRoleCheck = Application("Connection").Execute(strSQL)
   
   If objRoleCheck.EOF Then
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Custodians',70,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Kitchen',80,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Board of Education',90,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Bus Drivers',100,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Coaches',110,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Interns',120,True)"
   	Application("Connection").Execute(strSQL)
   	
   	strSQL = "INSERT INTO Roles (Role,RoleID,Active) VALUES ('Outside Vendors',130,True)"
   	Application("Connection").Execute(strSQL)
   End If

	'***********************************************************************************

	'Build the Log table if it's missing
	If Not bolLogTableFound Then
	
		strSQL = "CREATE TABLE Log ("
   	strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
   	strSQL = strSQL & "LGTag TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "UserName TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "EventNumber INTEGER,"
   	strSQL = strSQL & "Type TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "OldValue TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "NewValue TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "UpdatedBy TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "LogDate DATETIME,"
   	strSQL = strSQL & "LogTime DATETIME,"
   	strSQL = strSQL & "Active BIT,"
   	strSQL = strSQL & "Deleted BIT,"
   	strSQL = strSQL & "OldNotes LONGTEXT WITH COMPRESSION,"
   	strSQL = strSQL & "NewNotes LONGTEXT WITH COMPRESSION)"
   	Application("Connection").Execute(strSQL)
	End If

	'Build the PendingTasks table if it's missing
	If Not bolPendingTasksFound Then
	
		strSQL = "CREATE TABLE PendingTasks ("
   	strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
   	strSQL = strSQL & "Task TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "UserName TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "NewValue TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "UpdatedBy TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "TaskDate DATETIME,"
   	strSQL = strSQL & "TaskTime DATETIME,"
   	strSQL = strSQL & "Active BIT)"
   	Application("Connection").Execute(strSQL)
	End If
	
	'Build the History table if it's missing
	If Not bolCountHistoryFound Then
		strSQL = "CREATE TABLE CountHistory ("
   	strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
   	strSQL = strSQL & "Role TEXT(255) WITH COMPRESSION,"
   	strSQL = strSQL & "StudentCount INTEGER,"
   	strSQL = strSQL & "RecordedDate DATETIME)"
   	Application("Connection").Execute(strSQL)
	End If
	
	'Build the Owed table if it's missing
	If Not bolOwedFound Then
		strSQL = "CREATE TABLE Owed ("
		strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
		strSQL = strSQL & "Item TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "Price DECIMAL(5,2),"
		strSQL = strSQL & "OwedBy INTEGER,"
		strSQL = strSQL & "RecordedDate DATETIME,"
		strSQL = strSQL & "PaidDate DATETIME,"
		strSQL = strSQL & "Returnable BIT,"
		strSQL = strSQL & "Active BIT)"
		Application("Connection").Execute(strSQL)
		
		'Change the log so it shows money owed or paid for past entries
		strSQL = "UPDATE Log SET Type='MoneyOwed',NewValue='Insurance Copay - $100' WHERE Type='LoanedOutItem' AND NewValue='Owes $100 For Insurance'"
		Set objInsurance = Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Log SET Type='MoneyPaid',NewValue='Paid Insurance Copay - $100' WHERE Type='LoanedOutItemReturned' AND NewValue='Owes $100 For Insurance'"
		Set objInsurance = Application("Connection").Execute(strSQL)
		
		'Get the list of past insurance claims and convert them to the new format
		strSQL = "SELECT ID,AssignedTo,LoanDate,ReturnDate,Returned FROM Loaned WHERE Item='Owes $100 For Insurance'"
		Set objInsurance = Application("Connection").Execute(strSQL)
		
		If Not objInsurance.EOF Then
			
			Do Until objInsurance.EOF
			
				If objInsurance(4) Then
					strSQL = "INSERT INTO Owed (Item,Price,OwedBy,RecordedDate,PaidDate,Active) VALUES ('Insurance Copay',100.00,"
					strSQL = strSQL & objInsurance(1) & ",#" & objInsurance(2) & "#,#" & objInsurance(3) & "#,False)"
					Application("Connection").Execute(strSQL)
				Else
					strSQL = "INSERT INTO Owed (Item,Price,OwedBy,RecordedDate,Active) VALUES ('Insurance Copay',100.00,"
					strSQL = strSQL & objInsurance(1) & ",#" & objInsurance(2) & "#,True)"
					Application("Connection").Execute(strSQL)
					strSQL = "UPDATE People SET Warning=True WHERE ID=" & objInsurance(1)
					Application("Connection").Execute(strSQL)
				End If
				
				strSQL = "DELETE FROM Loaned WHERE ID=" & objInsurance(0)
				Application("Connection").Execute(strSQL)

				objInsurance.MoveNext
			Loop
		End If
		
		'Convert missing items to the new system
		strSQL = "SELECT LGTag,AssignedTo,DateReturned FROM Assignments WHERE Active=False AND AdapterReturned=False"
		Set objMissing = Application("Connection").Execute(strSQL)
		
		If Not objMissing.EOF Then
			Do Until objMissing.EOF
				
				strSQL = "SELECT Model FROM Devices WHERE LGTag='" & objMissing(0) & "'"
				Set objComputer = Application("Connection").Execute(strSQL)
				
				strSQL = "INSERT INTO Owed (Item,Price,OwedBy,RecordedDate,Returnable,Active) VALUES ('" & objComputer(0) & " Charger',79.00,"
				strSQL = strSQL & objMissing(1) & ",#" & objMissing(2) & "#,True,True)"
				Application("Connection").Execute(strSQL)
				
				objMissing.MoveNext
			Loop 
		End If
		
		strSQL = "SELECT LGTag,AssignedTo,DateReturned FROM Assignments WHERE Active=False AND CaseReturned=False"
		Set objMissing = Application("Connection").Execute(strSQL)
		
		If Not objMissing.EOF Then
			Do Until objMissing.EOF
				
				strSQL = "INSERT INTO Owed (Item,Price,OwedBy,RecordedDate,Returnable,Active) VALUES ('Laptop Case',20.00,"
				strSQL = strSQL & objMissing(1) & ",#" & objMissing(2) & "#,True,True)"
				Application("Connection").Execute(strSQL)
				
				objMissing.MoveNext
			Loop 
		End If
		
		Set objMissing = Nothing
		Set objInsurance = Nothing
		
		'Remove the missing items fields
		strSQL = "ALTER TABLE Assignments DROP AdapterReturned"
		Application("Connection").Execute(strSQL)
		strSQL = "ALTER TABLE Assignments DROP CaseReturned"
		Application("Connection").Execute(strSQL)
		
		
		'Unify the laptop cases under one name
		strSQL = "UPDATE Items SET Active=False WHERE Item='Belkin Laptop Case'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Items SET Active=False WHERE Item='Case Logic Laptop Case'"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO Items (Item,Active) VALUES ('Laptop Case',True)"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Items SET Item='Laptop Case' WHERE Item='Belkin Laptop Case'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Items SET Item='Laptop Case' WHERE Item='Case Logic Laptop Case'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Log SET NewValue='Laptop Case' WHERE NewValue='Belkin Laptop Case'"
		Application("Connection").Execute(strSQL)
		strSQL = "UPDATE Log SET NewValue='Laptop Case' WHERE NewValue='Case Logic Laptop Case'"
		Application("Connection").Execute(strSQL)
		
		'Remove the old way to work with insurance 
		strSQL = "UPDATE Items SET Active = False WHERE Item='Owes $100 For Insurance'"
		Application("Connection").Execute(strSQL)
		
	End If
	
	'Build the Purchasable table if it's missing
	If Not bolPurchasableFound Then
		strSQL = "CREATE TABLE Purchasable ("
		strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
		strSQL = strSQL & "Item TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "Price DECIMAL(5,2),"
		strSQL = strSQL & "Active BIT)"
		Application("Connection").Execute(strSQL)
		
		strSQL = "INSERT INTO Purchasable (Item,Price,Active) VALUES ('Insurance Copay',100.00,True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO Purchasable (Item,Price,Active) VALUES ('MacBook Air Charger',79.00,True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO Purchasable (Item,Price,Active) VALUES ('MacBook Pro Charger',79.00,True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO Purchasable (Item,Price,Active) VALUES ('iPad Charger',19.00,True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO Purchasable (Item,Price,Active) VALUES ('iPad Cable',19.00,True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO Purchasable (Item,Price,Active) VALUES ('Laptop Case',20.00,True)"
		Application("Connection").Execute(strSQL)
	
	End If
	
	If Not bolGroupMappings Then
	
		'Build the GroupMappings table if it's missing
		strDomainAdmins = "CN=Domain Admins,CN=Users,DC=lkgeorge,DC=org"
	
		strSQL = "CREATE TABLE GroupMappings ("
		strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
		strSQL = strSQL & "RoleID INTEGER,"
		strSQL = strSQL & "Site TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "DN TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "Active BIT)"
		Application("Connection").Execute(strSQL)
		
		strSQL = "INSERT INTO GroupMappings (RoleID,Site,DN,Active) VALUES (10,'High School','" &  strDomainAdmins & "',True)"
		'Application("Connection").Execute(strSQL)
		
	End If
	
	If Not bolCheckInHistoryFound Then
	
		strSQL = "CREATE TABLE CheckInHistory ("
		strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
		strSQL = strSQL & "Role TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "CheckInCount INTEGER,"
		strSQL = strSQL & "TotalStudentCount INTEGER,"
		strSQL = strSQL & "RecordedDate DATETIME,"
		strSQL = strSQL & "InternalCheckIn BIT)"
		Application("Connection").Execute(strSQL)
	
	End If
	
	If Not bolLibreNMSFound Then
	
		strSQL = "CREATE TABLE LibreNMS ("
		strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
		strSQL = strSQL & "PortID INTEGER,"
		strSQL = strSQL & "Site TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "Room TEXT(255) WITH COMPRESSION)"
		Application("Connection").Execute(strSQL)
	
	End If
	
	If Not bolInternetTypes Then
	
		strSQL = "CREATE TABLE InternetTypes ("
		strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
		strSQL = strSQL & "InternetType TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "Active BIT)"
		Application("Connection").Execute(strSQL)
		
		strSQL = "INSERT INTO InternetTypes (InternetType,Active) VALUES ('Cable Modem',True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO InternetTypes (InternetType,Active) VALUES ('DSL',True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO InternetTypes (InternetType,Active) VALUES ('Satellite',True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO InternetTypes (InternetType,Active) VALUES ('Dial Up',True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO InternetTypes (InternetType,Active) VALUES ('Cellular Hot Spot',True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO InternetTypes (InternetType,Active) VALUES ('No Internet',True)"
		Application("Connection").Execute(strSQL)
	End If
	
	If Not bolParentsFound Then
		
		strSQL = "CREATE TABLE Parents ("
		strSQL = strSQL & "ID AUTOINCREMENT PRIMARY KEY,"
		strSQL = strSQL & "StudentID INTEGER,"
		strSQL = strSQL & "ParentID INTEGER,"
		strSQL = strSQL & "FirstName TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "LastName TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "Relationship TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "EMail TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "StreetNumber TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "UnitNumber TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "StreetName TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "Address2 TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "City TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "State TEXT(255) WITH COMPRESSION,"
		strSQL = strSQL & "ZIP TEXT(255) WITH COMPRESSION)"
		Application("Connection").Execute(strSQL)
		
	End If
	
	'Add Lost Device to the inventory
	strSQL = "SELECT ID FROM EventTypes WHERE EventType='Lost Device'"
	Set objEventTypesCheck = Application("Connection").Execute(strSQL)
	
	If objEventTypesCheck.EOF Then
		strSQL = "INSERT INTO EventTypes (EventType,Active) VALUES ('Lost Device', True)"
		Application("Connection").Execute(strSQL)
		strSQL = "INSERT INTO Categories (Category,Active) VALUES ('Lost Device', True)"
		Application("Connection").Execute(strSQL)
	End If
	
	'Remove all the Decommission Device events from the events table
	strSQL = "DELETE FROM Events WHERE Events.Type='Decommission Device'"
	Application("Connection").Execute(strSQL)
	
End Sub%>

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