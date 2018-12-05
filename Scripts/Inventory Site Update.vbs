'Created by Matthew Hull on 3/14/16
'Last Updated 6/14/16

'This script will perform the needed tasks to update the inventory database

Option Explicit

'On Error Resume Next

Dim objDBConnection, objADCommand, objArguments, strDomain, strStudentShare, strStudentOU, strURL
Dim intHSBegins, strHomeDrive, strScript, strGroupRoot, objPerms, strFromEMail, datFirstDayOfSchool
Dim intPasswordLife, strFromEMailAdmin, bolEnableEmail, datStartPasswordNag, datEndPasswordNag
Dim strEMailOverride, bolSummerMode, strDestinyExport, datTime, strHour, strMinute, objKeepEnabledInAD
Dim objKeepEnabledInDB, strOUBase, strPaperCut

Set objKeepEnabledInAD = CreateObject("Scripting.Dictionary")
Set objKeepEnabledInDB = CreateObject("Scripting.Dictionary")

'Set some variables
'strHomeDrive = "H:"
'strStudentShare = "\\lgfps01\StudentHome$\"
strScript = "logon.bat"
strStudentOU = "OU=Students,OU=Lake George CSD,DC=lkgeorge,DC=org"
strGroupRoot = "CN=Students,OU=Security Groups,OU=Lake George CSD,DC=lkgeorge,DC=org"
strPaperCut = "CN=PaperCut,OU=Security Groups,OU=Lake George CSD,DC=lkgeorge,DC=org"
strPaperCut = ""
strOUBase = "OU=%ROLE%,OU=Users,OU=%SITE%,OU=Lake George CSD,DC=lkgeorge,DC=org"
strDomain = "lkgeorge.org"
intHSBegins = 7
strFromEMail = "fullenn@lkgeorge.org"
strFromEMailAdmin = "inventory@lkgeorge.org"
datFirstDayOfSchool = "9/07/18"
strURL = "https://helpdesk.lkgeorge.org/inventory/"
strDestinyExport = "\\10.15.79.28\c$\Follett\FSC-Patron\StudentData.csv"
datStartPasswordNag = "6/10/18"
datEndPasswordNag = "6/22/18" 
intPasswordLife = 90 'Days
bolSummerMode = False
bolEnableEmail = True 'Use this to turn off all email messages.
strEMailOverride = "" '"hullm@lkgeorge.org" 'Use this to redirect all mail to one address for testing.
objKeepEnabledInAD.Add "17powellc", "9/15/18"

CONST FIRSTNAME = 0
CONST LASTNAME = 1
CONST CLASSOF = 2
CONST SITE = 3
CONST HOMEROOM = 4
CONST HOMEROOMEMAIL = 5
CONST STUDENTID = 6
CONST DATECREATED = 7
CONST SEX = 8
CONST BIRTHDAY = 9
CONST USERNAME = 10
Const PASSWORD = 11

'Set the groups that will have access to the student's home folder
Set objPerms = CreateObject("Scripting.Dictionary")
objPerms.Add "Domain Admins", "f"
objPerms.Add "Faculty", "c"
objPerms.Add "Teachers", "c"

'Created needed objects
Set objDBConnection = ConnectToDatabase
Set objADCommand = ConnectToActiveDirectory
Set objArguments = WScript.Arguments

'Get the curent time information
datTime = Time
strHour = DatePart("h",datTime)
strMinute = DatePart("n",datTime)

If strHour = 23 And strMinute = 59 Then
	UpdateCheckInHistory
End If

'Early daily run
If strHour = 5 And strMinute = 30 Then
   RunPendingTasks
   ScanImportFileForChanges True
   UpdatePasswordExpirationDate False
   ValidateADAccounts strStudentOU
   FixDestinyExport
   UpdateStudentCountHistory
   VerifyADandDBMatch

'Normal daily run
ElseIf strHour = 8 And strMinute = 0 Then
   RunPendingTasks
   UpdatePasswordExpirationDate True

'Hourly run
ElseIf strMinute = 0 Then
   RunPendingTasks
   UpdatePasswordExpirationDate False

'Every 5 minutes
ElseIf strMinute Mod 5 = 0 Then
   RunPendingTasks
End If

'Uncomment if you need to run out of cycle
RunAll

Set objDBConnection = Nothing
Set objADCommand = Nothing
Set objArguments = Nothing

Sub RunAll
	RunPendingTasks
  	'ScanImportFileForChanges True
  	'UpdatePasswordExpirationDate False
  	'ValidateADAccounts strStudentOU
  	'FixDestinyExport
  	'UpdateStudentCountHistory
  	'VerifyADandDBMatch
  'MsgBox "Done"
End Sub

Sub ScanImportFileForChanges(bolResetPasswords)

	'This will scan the import file for new users.  If found, a new account and home folder will be 
	'created for the user, and they will be added to the inventory site.  The new account will be
	'disabled waiting for the official password before it's activated.  If the user in the import 
	'file already exists their information in AD and the database will be update.  Also any users
	'who have left the district will be disabled in the database.  All of the users will be disabled
	'in the database and enabled once they're found in the import file.

	Dim objFSO, strCurrentFolder, strCSV, txtSourceCSV, strSQL, objUserLookup, arrUserData
	Dim strMessage, objShell, strUser
	
	'Exit if we're in summer mode, we don't need to scan the import file until home rooms are set
	If bolSummerMode Then
		Exit Sub
	End If
		
	CleanExportFile	
		
	'Get the CSV path
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCurrentFolder = objFSO.GetAbsolutePathName(".")
	strCurrentFolder = strCurrentFolder & "\CSV\"
	strCSV = strCurrentFolder & "Import.csv"
	
	'Disable all the student accounts in the database, only the active ones will be
	'enabled later.
	strSQL = "UPDATE People SET Active=False WHERE ClassOf > 2000"
	objDBConnection.Execute(strSQL)
	
	'Open the source CSV
	Set txtSourceCSV = objFSO.OpenTextFile(strCSV)
	
	'Discard the header row
	txtSourceCSV.ReadLine
	
	'Loop through each line 
	While txtSourceCSV.AtEndOfLine = False
				
		'Get the data from the row and add it to an array
		arrUserData = GetUserDataFromImportedData(txtSourceCSV.ReadLine)
		
		'If they aren't in the database then add them
		If Not ExistsInDatabase(arrUserData(STUDENTID),arrUserData(CLASSOF)) Then
		   
         'Find the first available UserName
         arrUserData(USERNAME) = CreateUsername(arrUserData(FIRSTNAME),arrUserData(LASTNAME),arrUserData(CLASSOF))
         
			'Create the account in Active Directory
			CreateStudentADAccount arrUserData

			'Create a home folder for them
      	'CreateHomeFolder arrUserData(USERNAME)
      	
      	'Add the user to the inventory database
			AddUserToDatabase arrUserData
      	
      	'Update the log
			UpdateLog "NewStudentDetected","",arrUserData(USERNAME),"",arrUserData(LASTNAME) & ", " & arrUserData(FIRSTNAME),""

			'Send email about the new student found.
			SendEmail "Natalie","fullenn@lkgeorge.org", "NewStudentFound", arrUserData
			SendEmail "Rene","palmerr@lkgeorge.org", "NewStudentFound", arrUserData
			SendEmail "Matt","hullm@lkgeorge.org", "NewStudentFound", arrUserData
			SendEmail "Janine","wayj@lkgeorge.org", "NewStudentFound", arrUserData

		Else
			
			'Get the user's username and password
			strSQL = "SELECT UserName,PWord,AUP FROM People WHERE StudentID=" & arrUserData(STUDENTID)
			Set objUserLookup = objDBConnection.Execute(strSQL)
			arrUserData(USERNAME) = objUserLookup(0)
			
			'The user already exists, but we will update a few values in the database, 
			'this is where the account is enabled in the database
			ModifyUserInDatabase arrUserData

			'If the account is found in AD then reset some settings
			If ExistsInActiveDirectory(arrUserData(USERNAME)) Then
			
				'Modify the account in AD
				ModifyADAccount arrUserData, objUserLookup(1), bolResetPasswords

				'Create the home folder if needed
				'CreateHomeFolder arrUserData(USERNAME)
				
				'Enable or disable the account based on the AUP
				If Not bolSummerMode Then
					CheckAUPStatus arrUserData, objUserLookup(2)
				End If
				
			End If
		
		End If
	
	Wend
	
	'If there's an override enabled then turn back on the account in the database
	For Each strUser in objKeepEnabledInDB
		If CDate(objKeepEnabledInDB.Item(strUser)) >= Date() Then
			arrUserData = GetUserDataFromDatabase(strUser)
			ModifyUserInDatabase arrUserData
		End If
	Next
	
	'Close objects
	Set objFSO = Nothing
	Set txtSourceCSV = Nothing

End Sub

Sub RunPendingTasks

	'This will run the pending tasks in the inventory database

	Dim strSQL, objPendingTasks, arrUserData, intID, strTask, strUserName, strNewValue
	Dim objUserLookUp, objRoleLookup, strSite, strRole
	
	'Get the list of pending tasks
	strSQL = "SELECT ID,Task,UserName,NewValue FROM PendingTasks WHERE Active=True"
	Set objPendingTasks = objDBConnection.Execute(strSQL)
	
	If Not objPendingTasks.EOF Then

		Do Until objPendingTasks.EOF
		
			'Set the variables using information from the database
			intID = objPendingTasks(0)
			strTask = objPendingTasks(1)
			strUserName = objPendingTasks(2)
			strNewValue = objPendingTasks(3)
			arrUserData = GetUserDataFromDatabase(strUserName)
	
			Select Case strTask
	
				Case "AUPEnable"
				
					'The AUP has been enabled so enable the AD account if needed					
					If Not IsActiveInActiveDirectory(strUserName) Then
						EnableInActiveDirectory strUserName
						UpdateLog "AccountEnabledAUP","",strUserName,"","",""
						SendEmail "Natalie", "fullenn@lkgeorge.org", "EnabledAUPAdmin", arrUserData
						SendEmail "Matt", "hullm@lkgeorge.org", "EnabledAUPAdmin", arrUserData
						SendEmail "Janine", "wayj@lkgeorge.org", "EnabledAUPAdmin", arrUserData
						SendEMailToTeachers "EnabledAUP", arrUserData
					End If
					
				Case "ActivateStudent"
					
					'Activate the account in the database
					strSQL = "UPDATE People SET Notes='',Active=True,Pending=False WHERE UserName='" & strUserName & "'"
					objDBConnection.Execute(strSQL)
					
					'Enable the account and set the password
					EnableInActiveDirectory strUserName
					ModifyADAccount arrUserData,strNewValue,True
					UpdateUserPassword strUserName, strNewValue
					
					'Update the log and send notifications
					UpdateLog "NewStudentReady","",strUserName,"","",""
					SendEmail "Natalie", "fullenn@lkgeorge.org", "NewStudentReadyAdmin", arrUserData
					SendEmail "Matt", "hullm@lkgeorge.org", "NewStudentReadyAdmin", arrUserData
					SendEmail "Rene","palmerr@lkgeorge.org", "NewStudentReady", arrUserData
					SendEmail "Janine","wayj@lkgeorge.org", "NewStudentReady", arrUserData
					SendEMailToTeachers "NewStudentReady", arrUserData
					
				Case "UpdateDescription"
					UpdateUserDescription strUserName, strNewValue
					
				Case "UpdatePhone"
					UpdateUserPhone strUserName, strNewValue
					
				Case "UpdateFirstName"
					UpdateUserFirstName strUserName, strNewValue
				
				Case "UpdateLastName"
					UpdateUserLastName strUserName, strNewValue
				
				Case "UpdateUserName"
					UpdateUserUserName strUserName, strNewValue
					
				Case "UpdatePassword"
					UpdateUserPassword strUserName, strNewValue
					
				Case "UpdateRoom"
					UpdateUserRoom strUserName, strNewValue
					
				Case "PasswordNeverExpires"
					TogglePasswordStatus strUserName, False 
				
				Case "PasswordExpires"
					TogglePasswordStatus strUserName, True
					
				Case "MoveUser"
					strSQL = "SELECT CLassOf,Site FROM People WHERE UserName='" & strUserName & "'"
					Set objUserLookUp = objDBConnection.Execute(strSQL)
					strSite = objUserLookUp(1)
					
					strSQL = "SELECT Role FROM Roles WHERE RoleID=" & objUserLookUp(0)
   				Set objRoleLookup = objDBConnection.Execute(strSQL)
   				strRole = objRoleLookup(0)
   				
   				MoveUser strUserName, strRole, strSite
   				
   			Case "CreateAccount"
   				CreateAdultADAccount arrUserData, strNewValue
   				
   			Case "EnableUser"
   				EnableInActiveDirectory strUserName
   			
   			Case "DisableUser"
   				DisableInActiveDirectory strUserName, ""
					
			End Select
			
			'Deactivate the task
			strSQL = "UPDATE PendingTasks SET Active=False WHERE ID=" & intID
			objDBConnection.Execute(strSQL)
	
			objPendingTasks.MoveNext
		
		Loop
	
		'Close objects
		Set objPendingTasks = Nothing
	
	End If

End Sub

Sub CheckAUPStatus(arrUserData, bolAUP)

	'If the user hasn't turned in their AUP then see if two weeks have passed since the start of school.
	'All students have two weeks at the start of school to turn in their AUP.  If a student starts after
	'the first two weeks of school then they have 2 weeks from when they were entered into SchoolTool to
	'turn in their AUP.  This will disabled the account if needed.

	'Have they turned in their AUP
	If Not bolAUP Then
		
		'Check and see if 2 weeks have passed from the start of school
		If DateDiff("d",datFirstDayOfSchool,Date) > 14 Then
			
			'Check and see if 2 weeks have passed from when they were created in SchoolTool
			If DateDiff("d",arrUserData(DATECREATED),Date) > 14 Then
				If IsActiveInActiveDirectory(arrUserData(USERNAME)) Then
					DisableInActiveDirectory arrUserData(USERNAME), "AUP"
					UpdateLog "AccountDisabledAUP","",arrUserData(USERNAME),"","",""
					SendEmail "Natalie", "fullenn@lkgeorge.org", "DisabledAUPAdmin", arrUserData
					SendEmail "Matt", "hullm@lkgeorge.org", "DisabledAUPAdmin", arrUserData
					SendEmail "Janine", "wayj@lkgeorge.org", "DisabledAUPAdmin", arrUserData
					SendEMailToTeachers "DisabledAUP", arrUserData
				End If
			End If
		End If
	End If

End Sub

Function GetUserDataFromImportedData(strImportedData)

	'This function will return an array that contains all the information about a user from the
	'import file.

	Dim arrRow, strLastName, strFirstName, intClassOf, intStudentID, strSex, datBirthday, datDateCreated
	Dim strHomeRoom, strHomeRoomEmail, arrLastName, strSite
	
   'On Error Resume Next
   
	arrRow = Split(strImportedData,",")
		
	'Get the variables from the row
	strLastName= Trim(arrRow(0))
	strFirstName = Trim(arrRow(1))
	intClassOf = Right(Trim(arrRow(2)),4)
	intStudentID = Trim(arrRow(3))
	strSex = Trim(arrRow(4))
	datBirthday = Trim(arrRow(5))
	datDateCreated = Trim(arrRow(6))

	'Fix the homeroom variables if there are two teachers
	If InStr(arrRow(7),"/") <> 0 Then
		strHomeRoomEmail = Replace(Trim(arrRow(7))," / ",";")
		strHomeRoom = Replace(Trim(arrRow(8)),"""","")
	Else
		'On Error Resume Next
      strHomeRoomEmail = Trim(arrRow(7))
		strHomeRoom = Replace(Trim(arrRow(8)) & ", " & Trim(arrRow(9)),"""","")
      
      'If Err Then
      '   msgbox arrRow(8)
      '   err.Clear
      '   wscript.exit
      'End If
	End If
   
   'If Err Then
   '   MsgBox strFirstName & " " & strLastName
   '   Wscript.Quote
   'End If

	'Fix the last name
	If InStr(strLastName," ") <> 0 Then
		arrLastName = Split(strLastName," ")
	
		'Fix the suffix
		Select Case LCase(arrLastName(1))
			Case "jr", "jr."
				strLastName = arrLastName(0) & " Jr"

			Case "ii", "2", "2nd"
				strLastName = arrLastName(0) & " II"
		
			Case "iii", "3", "3rd"
				strLastName = arrLastName(0) & " III"
		
			Case "iv", "4", "4th"
				strLastName = arrLastName(0) & " IV"
			
		End Select
	End If
	
	'Set the site
	If GetGrade(intClassOf) = "K" Then
		strSite = "Elementary"
	Else
		If GetGrade(intClassOf) >= intHSBegins Then
			strSite = "High School"
		Else
			strSite = "Elementary"
		End If
	End If
	
	'Add the user's data to an array
	GetUserDataFromImportedData = Array(strFirstName,strLastName,intClassOf,strSite,strHomeRoom, _
	strHomeRoomEmail,intStudentID,datDateCreated,strSex,datBirthday,"","")

End Function

Function GetUserDataFromDatabase(strUserName)
	
	'This function will return an array that contains all the needed information about a user
	
	Dim strSQL, objUserLookup, strFirstName, strLastName, intClassOf, strSite, strHomeRoom
	Dim strHomeRoomEMail, intStudentID, datDateCreated, strSex, datBirthday, strPassword
	
	'Get the user's information from the database
	strSQL = "SELECT FirstName,LastName,ClassOf,Site,HomeRoom,HomeRoomEmail,StudentID,DateAdded," & _
		"Sex,Birthday,PWord FROM People WHERE UserName='" & strUserName & "'"
	Set objUserLookup = objDBConnection.Execute(strSQL)
	
	'Write the information to an array
	If Not objUserLookUp.EOF Then
	
		'Set the variables 
		strFirstName = objUserLookup(0)
		strLastName = objUserLookup(1)
		intClassOf = objUserLookup(2)
		strSite = objUserLookup(3)
		strHomeRoom = objUserLookup(4)
		strHomeRoomEMail = objUserLookup(5)
		intStudentID = objUserLookup(6)
		datDateCreated = objUserLookup(7)
		strSex = objUserLookup(8)
		datBirthday = objUserLookup(9)
		strPassword = objUserLookup(10)
		
		'Create the array
		GetUserDataFromDatabase = Array(strFirstName,strLastName,intClassOf,strSite,strHomeRoom, _
			strHomeRoomEmail,intStudentID,datDateCreated,strSex,datBirthday,strUserName,strPassword)
	
	Else
		
		'Return an empty array if the user isn't found
		GetUserDataFromDatabase = Array("","","","","","","","","","","")
	End If
	
End Function

Sub CreateStudentADAccount(arrUserData)
	
	'This will create a user account in Active Directory
	
	Dim objOU, objGroup, objUser, objPaperCut
	
	'Get the needed objects 
	Set objOU = GetObject("LDAP://OU=" & arrUserData(CLASSOF) & "," & strStudentOU)
	Set objGroup = GetObject("LDAP://" & strGroupRoot)

	'Create the user account in Active Directory 
	Set objUser = objOU.Create("User", "cn=" & arrUserData(FIRSTNAME) & " " & arrUserData(LASTNAME))
	objUser.Put "SAMAccountName", arrUserData(USERNAME)
	objUser.Put "SN", arrUserData(LASTNAME)
	objUser.Put "GivenName", arrUserData(FIRSTNAME)
	objUser.Put "DisplayName", arrUserData(FIRSTNAME) & " " & arrUserData(LASTNAME)
	objUser.Put "UserPrincipalName", arrUserData(USERNAME) & "@" & strDomain
	objUser.Put "UserAccountControl", 544 '544 = normal account, no password required
	objUser.Put "ScriptPath", strScript
	'objUser.Put "HomeDirectory", strStudentShare & arrUserData(USERNAME)
	'objUser.Put "HomeDrive", strHomeDrive
	objUser.Put "Description", "Class of " & arrUserData(CLASSOF) & " - Pending Account"
	objUser.SetInfo
	objUser.Setpassword("P@ssw0rd")
	objUser.Put "UserAccountControl", 66050 'Account Disabled
	objUser.SetInfo
	objGroup.Add(objUser.ADSPath)
	
	'Add the new user to the PaperCut group if it's set
	If strPaperCut <> "" Then
		Set objPaperCut = GetObject("LDAP://" & strPaperCut)
		objPaperCut.Add(objUser.ADSPath)
		Set objPaperCut = Nothing
	End If
	
   'Close objects
	Set objOU = Nothing
	Set objGroup = Nothing
	Set objUser = Nothing
   
End Sub

Sub CreateAdultADAccount(arrUserData,strPassword)
	
	'This will create a user account in Active Directory
	
	Dim objOU, objGroup, objUser, strSQL, objGroups, objUserInfo, strDescription, bolRequireChange

	'Get the Description from the database
	strSQL = "SELECT Description, PWordNeverExpires FROM People WHERE UserName='" & arrUserData(USERNAME) & "'"
	Set objUserInfo = objDBConnection.Execute(strSQL)
	
	If Not objUserInfo.EOF Then
		strDescription = objUserInfo(0)
		If objUserInfo(1) Then
			bolRequireChange = True
		Else
			bolRequireChange = False
		End If
	Else
		strDescription = ""
		bolRequireChange = ""
	End If
	
	'Reset the PWordNeverExpires database setting
	If bolRequireChange Then
		strSQL = "UPDATE People SET PWordNeverExpires=False WHERE UserName='" & arrUserData(USERNAME) & "'"
		objDBConnection.Execute(strSQL)
	End If

	'Get the needed objects 
	Set objOU = GetOU(GetRole(arrUserData(CLASSOF)),arrUserData(SITE))

	'Create the user account in Active Directory 
	Set objUser = objOU.Create("User", "cn=" & arrUserData(FIRSTNAME) & " " & arrUserData(LASTNAME))
	objUser.Put "SAMAccountName", arrUserData(USERNAME)
	objUser.Put "SN", arrUserData(LASTNAME)
	objUser.Put "GivenName", arrUserData(FIRSTNAME)
	objUser.Put "DisplayName", arrUserData(FIRSTNAME) & " " & arrUserData(LASTNAME)
	objUser.Put "UserPrincipalName", arrUserData(USERNAME) & "@" & strDomain
	objUser.Put "UserAccountControl", 544 '544 = normal account, no password required
	objUser.Put "ScriptPath", strScript
	If strDescription <> "" Then
		objUser.Put "Description", strDescription
	End If
	'objUser.Put "HomeDirectory", strStudentShare & arrUserData(USERNAME)
	'objUser.Put "HomeDrive", strHomeDrive
	'objUser.Put "Description", "Class of " & arrUserData(CLASSOF) & " - Pending Account"
	objUser.SetInfo
	objUser.Setpassword(strPassword)
	objUser.Put "UserAccountControl", 512 'Account Enabled, Password Expires
	objUser.SetInfo
	
	If bolRequireChange Then
		objUser.Put "userAccountControl", 512
		objUser.Put "PwdLastSet", 0
		objUser.SetInfo	
	End If
	
	'Add the new user to the correct groups
	strSQL = "SELECT DN FROM GroupMappings WHERE RoleID=" & arrUserData(CLASSOF) & " AND Site='" & arrUserData(SITE) & "'"
	Set objGroups = objDBConnection.Execute(strSQL)
	If Not objGroups.EOF Then
		Do Until objGroups.EOF
			Set objGroup = GetObject("LDAP://" & objGroups(0))
			objGroup.Add(objUser.ADSPath)
			Set objGroup = Nothing
			objGroups.MoveNext
		Loop
	End If
	
	'Email people to let them know the account was created
	SendEmail "Matt","hullm@lkgeorge.org", "NewAccountCreated", arrUserData
	SendEmail "Dane","davisd@lkgeorge.org", "NewAccountCreated", arrUserData
	SendEmail "Janine","wayj@lkgeorge.org", "NewAccountCreated", arrUserData
	
	'Update the log
	UpdateLog "NewADAccountCreated","",arrUserData(USERNAME),"","",""
	
   'Close objects
	Set objOU = Nothing
	Set objGroups = Nothing
	Set objUser = Nothing
   
End Sub

Sub ModifyADAccount(arrUserData,strPassword,bolResetPassword)

   On Error Resume Next

	'This will update the users settings in AD

	Dim objRootDSE, objUserLookup, objUser, objGroup

	'Create a RootDSE object for the domain
	 Set objRootDSE = GetObject("LDAP://RootDSE")
	
	'Get the user object from Active Directory	
	objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
	">;(&(objectClass=user)(samAccountName=" & arrUserData(USERNAME) & "));distinguishedName"
	Set objUserLookup = objADCommand.Execute
	Set objUser = GetObject("LDAP://" & objUserLookup(0))
			
	'Update the description as long as the account isn't waiting to be activated
	If InStr(objUser.Description,"Pending") = 0 Then
		objUser.Put "Description", "Class of " & arrUserData(CLASSOF)
	End If

	'Add the student to the student's group if needed
	Set objGroup = GetObject("LDAP://" & strGroupRoot)
	If objGroup.IsMember("LDAP://" & objUserLookup(0)) = False Then
		objGroup.Add(objUser.ADSPath)
	End If

	'Fix some of the other properties of the account
	objUser.Put "ScriptPath", strScript
   objUser.PutEx 1,"HomeDirectory",vbNullString
   objUser.PutEx 1,"HomeDrive",vbNullString
	'objUser.Put "HomeDirectory", strStudentShare & arrUserData(USERNAME)
	'objUser.Put "HomeDrive", strHomeDrive

	'Reset the password if requested
	If bolResetPassword Then
		If DateDiff("h",GetLastPasswordSet(arrUserData(USERNAME)),Date) < 23 Then
			If DateDiff("h",arrUserData(DATECREATED),Date) < 24 Then
				objUser.SetPassword(strPassword)
			End If
		End If
	End If
   
   If Err Then
      'MsgBox objUser.sAMAccountName & " " & strPassword & "."
      Err.Clear
   End If

	'Save changes
	objUser.SetInfo

	'Close object
	Set objRootDSE = Nothing
	Set objUserLookup = Nothing
	Set objUser = Nothing
	Set objGroup = Nothing

End Sub

Sub AddUserToDatabase(arrUserData)

	'This will add the new user to the database with a password of NewAccount.  The
	'user will be flagged as deleted in the database.

	Dim strSQL

	strSQL = "INSERT INTO People (FirstName,LastName,Username,Role,ClassOf,Site,HomeRoom,HomeRoomEmail,Sex,Birthday,PWord,StudentID,Active,Pending,AUP,Deleted,DateAdded)" & vbCRLF
	strSQL = strSQL & "VALUES ('"
	strSQL = strSQL & Replace(arrUserData(FIRSTNAME),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(LASTNAME),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(USERNAME),"'","''") & "','"
	strSQL = strSQL & "Student','"
	strSQL = strSQL & arrUserData(CLASSOF) & "','"
	strSQL = strSQL & Replace(arrUserData(SITE),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(HOMEROOM),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(HOMEROOMEMAIL),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(SEX),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(BIRTHDAY),"'","''") & "','"
	strSQL = strSQL & "NewAccount',"
	strSQL = strSQL & arrUserData(STUDENTID) & ",False,True,"
	strSQL = strSQL & "False,True,#" & arrUserData(DATECREATED) & "#)"
	objDBConnection.Execute(strSQL)
	
End Sub

Sub ModifyUserInDatabase(arrUserData)

	'This will update the settings in the inventory database for the user

	Dim strSQL, objPendingCheck, bolActive, bolDeleted
	
	'Find out if the account is waiting to be created, if so then
	strSQL = "SELECT Pending FROM People WHERE UserName='" & arrUserData(USERNAME) & "'"
	Set objPendingCheck = objDBConnection.Execute(strSQL)
		 
	'If the account is still pending then don't set it active and keep it deleted
	If objPendingCheck(0) Then
		bolActive = False
		bolDeleted = True
	Else
		bolActive = True
		bolDeleted = False
	End If
	
	'Update the settings in the database
	strSQL = "UPDATE People SET "
	strSQL = strSQL & "HomeRoom='" & Replace(arrUserData(HOMEROOM),"'","''") & "'," 
	strSQL = strSQL & "HomeRoomEmail='" & Replace(arrUserData(HOMEROOMEMAIL),"'","''") & "'," 
	strSQL = strSQL & "Site='" & Replace(arrUserData(SITE),"'","''") & "'," 
	strSQL = strSQL & "Deleted= " & bolDeleted & "," 
	strSQL = strSQL & "Active=" & bolActive & "," 
	strSQL = strSQL & "Sex='" & arrUserData(SEX) & "',"
	strSQL = strSQL & "Birthday=#" & arrUserData(BIRTHDAY) & "#,"
	strSQL = strSQL & "DateAdded=#" & arrUserData(DATECREATED) & "# " 
	strSQL = strSQL & "WHERE Role='Student' AND StudentID=" & arrUserData(STUDENTID)
	objDBConnection.Execute(strSQL)

End Sub

Sub CreateHomeFolder(strFolderName)

	'This will create the folder and set the permissions

	Dim objFSO, objShell, strCMD, Key

	'Create the needed file system object 
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'Create the home folder and set permissions
	If Not objFSO.FolderExists(strStudentShare & strFolderName) Then
	
		objFSO.CreateFolder(strStudentShare & strFolderName) 
	
		'Set the permissions on the new folder
		Set objShell = CreateObject("Wscript.Shell")
		objPerms.Add strFolderName, "c"
		strCMD = "cmd /c echo y| cacls " & """" & strStudentShare & strFolderName & """ /c /t /g "
		For Each Key in objPerms
			strCMD = strCMD & """" & strDomain & "\" & Key & """" & ":" & objPerms.Item(Key) & " "      
		Next
		objShell.Run strCMD,0,true 
		objPerms.Remove(strFolderName)
		Set objShell = Nothing
	
	End If

	'Close objects
	Set objFSO = Nothing

End Sub

Sub UpdateUserDescription(strUserName,strDescription)

	'This will update the description on a user's account
	
	Const ADS_PROPERTY_CLEAR = 1
	
	Dim objRootDSE, objUserLookUp, objUser, arrUserData, strSQL, objOldDescription, strOldDescription, bolError
   
   bolError = False
   
   'Make sure the user account exists before trying to update it.
   If ExistsInActiveDirectory(strUserName) Then
   	
   	If Len(strDescription) <=250 Then
   	
			'Create a RootDSE object for the domain
			Set objRootDSE = GetObject("LDAP://RootDSE")
	
			'Get the user object from Active Directory	
			objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
			">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
			Set objUserLookup = objADCommand.Execute
	
			'Build the user object
			Set objUser = GetObject("LDAP://" & objUserLookup(0))
		
			'Add or remove the attribute
			If strDescription = "" Then
				objUser.PutEx ADS_PROPERTY_CLEAR, "Description", null
			Else
				objUser.Put "Description", strDescription
			End If
		
			objUser.SetInfo
		
		Else
   		bolError = True 'Description too long
   	End If
   Else  
   	bolError = True 'User not in Active Directory
	End If
	
	If bolError Then
	
		'Something wen't wrong, we can't find the account in AD, so let people know.
   	arrUserData = GetUserDataFromDatabase(strUserName)
		SendEmail "Matt", "hullm@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Dane", "davisd@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Janine", "wayj@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		
		'Get the old description so we can set it back.
		strSQL = "SELECT OldValue FROM Log WHERE Type='UserUpdatedDescription' AND UserName='" & strUserName & "' ORDER BY ID DESC"
		Set objOldDescription = objDBConnection.Execute(strSQL)
   	
   	'Change the description back to the old one in the database if we were able to find it.
		If Not objOldDescription.EOF Then
			strOldDescription = objOldDescription(0)
			UpdateLog "UserUpdatedDescription","",strUserName,strDescription,strOldDescription,""
			strSQL = "UPDATE People SET Description='" & strOldDescription & "' WHERE UserName='" & strUserName & "'"
			objDBConnection.Execute(strSQL)
		End If
	End If
   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing

End Sub

Sub UpdateUserPhone(strUserName,strPhone)

	'This will update the phone number on a user's account
	
	Const ADS_PROPERTY_CLEAR = 1
	
	Dim objRootDSE, objUserLookUp, objUser, arrUserData, strSQL, objOldPhoneNumber, strOldPhoneNumber, bolError
	
	bolError = False
	
	'Make sure the user account exists before trying to update it.
   If ExistsInActiveDirectory(strUserName) Then
	
		'Make sure the provided input isn't crazy long.
		If Len(strPhone) <= 25 Then
		
			'Create a RootDSE object for the domain
			Set objRootDSE = GetObject("LDAP://RootDSE")
	
			'Get the user object from Active Directory	
			objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
			">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
			Set objUserLookup = objADCommand.Execute
	
			'Build the user object
			Set objUser = GetObject("LDAP://" & objUserLookup(0))
	
			'Add or remove the attribute
			If strPhone = "" Then
				objUser.PutEx ADS_PROPERTY_CLEAR, "telephoneNumber", null
				objUser.PutEx ADS_PROPERTY_CLEAR, "ipPhone", null
			Else
				objUser.Put "telephoneNumber", strPhone
				objUser.Put "ipPhone", strPhone
			End If
		
			objUser.SetInfo

		Else
			bolError = True 'Phone number too long
   	End If
   Else
   	bolError = True 'User not in Active Directory
   End If
   
   If bolError Then
   	
   	'Something went wrong, let people know
   	arrUserData = GetUserDataFromDatabase(strUserName)
		SendEmail "Matt", "hullm@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Dane", "davisd@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Janine", "wayj@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		
		'Get the old phone number so we can set it back.
		strSQL = "SELECT OldValue FROM Log WHERE Type='UserUpdatedPhone' AND UserName='" & strUserName & "' ORDER BY ID DESC"
		Set objOldPhoneNumber = objDBConnection.Execute(strSQL)
   	
   	'Change the phone number back to the old one in the database if we were able to find it.
		If Not objOldPhoneNumber.EOF Then
			strOldPhoneNumber = objOldPhoneNumber(0)
			UpdateLog "UserUpdatedPhone","",strUserName,strPhone,strOldPhoneNumber,""
			strSQL = "UPDATE People SET PhoneNumber='" & strOldPhoneNumber & "' WHERE UserName='" & strUserName & "'"
			objDBConnection.Execute(strSQL)
		End If
   
   End If
   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing

End Sub

Sub UpdateUserRoom(strUserName,strRoom)

	'This will update the room on a user's account
	
	Const ADS_PROPERTY_CLEAR = 1
	
	Dim objRootDSE, objUserLookUp, objUser, arrUserData, strSQL, objOldRoomNumber, strOldRoomNumber, bolError
	
	bolError = False
	
	'Make sure the user account exists before trying to update it.
   If ExistsInActiveDirectory(strUserName) Then
	
		'Make sure the provided input isn't crazy long.
		If Len(strRoom) <= 50 Then
		
			'Create a RootDSE object for the domain
			Set objRootDSE = GetObject("LDAP://RootDSE")
	
			'Get the user object from Active Directory	
			objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
			">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
			Set objUserLookup = objADCommand.Execute
	
			'Build the user object
			Set objUser = GetObject("LDAP://" & objUserLookup(0))
	
			'Add or remove the attribute
			If strRoom = "" Then
				objUser.PutEx ADS_PROPERTY_CLEAR, "physicalDeliveryOfficeName", null
			Else
				objUser.Put "physicalDeliveryOfficeName", strRoom
			End If
		
			objUser.SetInfo

		Else
			bolError = True 'Room too long
   	End If
   Else
   	bolError = True 'User not in Active Directory
   End If
   
   If bolError Then
   	
   	'Something went wrong, let people know
   	arrUserData = GetUserDataFromDatabase(strUserName)
		SendEmail "Matt", "hullm@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Dane", "davisd@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Janine", "wayj@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		
		'Get the old room so we can set it back.
		strSQL = "SELECT OldValue FROM Log WHERE Type='UserUpdatedRoom' AND UserName='" & strUserName & "' ORDER BY ID DESC"
		Set objOldRoomNumber = objDBConnection.Execute(strSQL)
   	
   	'Change the phone number back to the old one in the database if we were able to find it.
		If Not objOldRoomNumber.EOF Then
			strOldRoomNumber = objOldRoomNumber(0)
			UpdateLog "UserUpdatedRoom","",strUserName,strRoom,strOldRoomNumber,""
			strSQL = "UPDATE People SET RoomNumber='" & strOldRoomNumber & "' WHERE UserName='" & strUserName & "'"
			objDBConnection.Execute(strSQL)
		End If
   
   End If
   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing

End Sub

Sub UpdateUserPassword(strUserName,strPassword)

	'This will update the password on a user's account
	
	Dim objRootDSE, objUserLookUp, objUser, arrUserData, strSQL, objOldPassword, strOldPassword, bolError
	
	bolError = False
	
	'Make sure the user account exists before trying to update it.
   If ExistsInActiveDirectory(strUserName) Then
	
		If PasswordValid(strPassword) Then
			
			'Create a RootDSE object for the domain
			 Set objRootDSE = GetObject("LDAP://RootDSE")

			'Get the user object from Active Directory	
			objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
			">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
			Set objUserLookup = objADCommand.Execute

			'Build the user object
			Set objUser = GetObject("LDAP://" & objUserLookup(0))

			objUser.SetPassword(strPassword)
			objUser.SetInfo

		Else
			bolError = True 'Password is not a valid password
		End If	
	Else
		bolError = True 'User not in Active Directory
	End If
	
	If bolError Then
	
		'Something went wrong, let people know
		arrUserData = GetUserDataFromDatabase(strUserName)
		SendEmail "Matt", "hullm@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Dane", "davisd@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Janine", "wayj@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		
		'Get the old password so we can set it back
		strSQL = "SELECT OldValue FROM Log WHERE Type='UserUpdatedPassword' AND UserName='" & strUserName & "' ORDER BY ID DESC"
		Set objOldPassword = objDBConnection.Execute(strSQL)
		
		'Change the password back to the old password in the database if we were able to find it
		If Not objOldPassword.EOF Then
			strOldPassword = objOldPassword(0)
			UpdateLog "UserUpdatedPassword","",strUserName,strPassword,strOldPassword,""
			strSQL = "UPDATE People SET PWord='" & strOldPassword & "' WHERE UserName='" & strUserName & "'"
			objDBConnection.Execute(strSQL)
		End If
	
	End If
   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing

End Sub

Sub UpdateUserFirstName(strUserName,strFirstName)

	'This will update the first name on a user's account
	
	Dim objRootDSE, objUserLookUp, objUser, strLastName, objOU, strOU, arrUserData, strSQL, objOldFirstName, strOldFirstName, bolError
	
	bolError = False
	
	'Make sure the user account exists before trying to update it
   If ExistsInActiveDirectory(strUserName) Then
	
		'Make sure they didn't subit a blank value
		If strFirstName <> "" Then

			'Make sure the provided input isn't crazy long.
			If Len(strFirstName) <= 50 Then 
			
				'Create a RootDSE object for the domain
				 Set objRootDSE = GetObject("LDAP://RootDSE")
	
				'Get the user object from Active Directory	
				objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
				">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
				Set objUserLookup = objADCommand.Execute
	
				'Build the user object
				Set objUser = GetObject("LDAP://" & objUserLookup(0))
	
				'Set the new name on the user object
				strLastName = objUser.sn
				objUser.Put "givenName", strFirstName
				objUser.Put "displayName", strFirstName & " " & strLastName
				objUser.SetInfo
	
				'Move the user object to the new name
				strOU = Right(Replace(objUserLookup(0),"\,",""),Len(Replace(objUserLookup(0),"\,",""))-InStr(Replace(objUserLookup(0),"\,",""),","))
				Set objOU = GetObject("LDAP://" & strOU)
				objOU.MoveHere "LDAP://" & objUserLookup(0),"CN=" & strFirstName & " " & strLastName
		
			Else 
				bolError = True 'First name is too long
			End If
		Else
			bolError = True 'First name is blank
		End If
	Else
		bolError = True 'User not in Active Directory
	End If
	
	'Fix things if there was an error
	If bolError Then
	
		'Something went wrong, let people know
		arrUserData = GetUserDataFromDatabase(strUserName)
		SendEmail "Matt", "hullm@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Dane", "davisd@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Janine", "wayj@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
	
		'Get the old first name so we can set it back
		strSQL = "SELECT OldValue FROM Log WHERE Type='UserUpdatedFirstName' AND UserName='" & strUserName & "' ORDER BY ID DESC"
		Set objOldFirstName = objDBConnection.Execute(strSQL)
		
		'Change the first name back to the old first name in the database if we were able to find it
		If Not objOldFirstName.EOF Then
			strOldFirstName = objOldFirstName(0)
			UpdateLog "UserUpdatedFirstName","",strUserName,strFirstName,strOldFirstName,""
			strSQL = "UPDATE People SET FirstName='" & strOldFirstName & "' WHERE UserName='" & strUserName & "'"
			objDBConnection.Execute(strSQL)
		End If
	
	End If
   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing
   Set objOU = Nothing

End Sub

Sub UpdateUserLastName(strUserName,strLastName)

	'This will update the last name on a user's account
	
	Dim objRootDSE, objUserLookUp, objUser, strFirstName, objOU, strOU, arrUserData, strSQL, objOldLastName, strOldLastName, bolError
	
	bolError = False
	
	'Make sure the user account exists before trying to update it
   If ExistsInActiveDirectory(strUserName) Then
   
   	'Make sure they didn't subit a blank value
		If strLastName <> "" Then
	
			'Make sure the provided input isn't crazy long.
			If Len(strLastName) <= 50 Then 
	
				'Create a RootDSE object for the domain
				 Set objRootDSE = GetObject("LDAP://RootDSE")
	
				'Get the user object from Active Directory	
				objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
				">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
				Set objUserLookup = objADCommand.Execute
	
				'Build the user object
				Set objUser = GetObject("LDAP://" & objUserLookup(0))
	
				strFirstName = objUser.givenName
				objUser.Put "sn", strLastName
				objUser.Put "displayName", strFirstName & " " & strLastName
				objUser.SetInfo
	
				strOU = Right(Replace(objUserLookup(0),"\,",""),Len(Replace(objUserLookup(0),"\,",""))-InStr(Replace(objUserLookup(0),"\,",""),","))
				Set objOU = GetObject("LDAP://" & strOU)
				objOU.MoveHere "LDAP://" & objUserLookup(0),"CN=" & strFirstName & " " & strLastName
				
			Else
				bolError = True 'Last name is too long
			End If
		Else
   		bolError = True 'Last name is blank
   	End If
   Else
   	bolError = True 'User not in Active Directory 
   End If
   
   'Fix things if there was an error
	If bolError Then
	
		'Something went wrong, let people know
		arrUserData = GetUserDataFromDatabase(strUserName)
		SendEmail "Matt", "hullm@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Dane", "davisd@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Janine", "wayj@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
	
		'Get the old last name so we can set it back
		strSQL = "SELECT OldValue FROM Log WHERE Type='UserUpdatedLastName' AND UserName='" & strUserName & "' ORDER BY ID DESC"
		Set objOldLastName = objDBConnection.Execute(strSQL)
		
		'Change the first name back to the old first name in the database if we were able to find it
		If Not objOldLastName.EOF Then
			strOldLastName = objOldLastName(0)
			UpdateLog "UserUpdatedLastName","",strUserName,strFirstName,strOldLastName,""
			strSQL = "UPDATE People SET LastName='" & strOldLastName & "' WHERE UserName='" & strUserName & "'"
			objDBConnection.Execute(strSQL)
		End If
	
	End If
   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing
   Set objOU = Nothing

End Sub

Sub MoveUser(strUserName,strRole,strSite)

	'This will move a user to a new Organizational Unit in Active Directory
	
	Const ADS_PROPERTY_DELETE = 4
	
	Dim objOU, objRootDSE, objUserLookup, objUser, objGroupList, objGroup, arrUserData, strSQL, objGroups, strGroup
	
	'Make sure the user account exists before trying to update it
   If ExistsInActiveDirectory(strUserName) Then
	
		'Make sure the site and role aren't blank
		If strRole <> "" And strSite <> "" Then
		
			'Create a RootDSE object for the domain
			 Set objRootDSE = GetObject("LDAP://RootDSE")

			'Get the user DN from Active Directory	
			objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
			">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
			Set objUserLookup = objADCommand.Execute

			'Move the object
			Set objOU = GetOU(strRole,strSite)
			objOU.MoveHere "LDAP://" & objUserLookup(0), Left(objUserLookup(0),InStr(objUserLookup(0),",")-1)
			
			'Get the new user DN from Active Directory	
			objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
			">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
			Set objUserLookup = objADCommand.Execute
			
			'Build the user object
			Set objUser = GetObject("LDAP://" & objUserLookup(0))
			
			'If the user is only a member of Domain Users, or not a member of a group this section will crash
			'We have to check for, and handle the crash.
			On Error Resume Next
			
			'Get the list of groups of which the user is a member.
			objGroupList = objUser.GetEx("memberOf")
			
			'Remove the user from each group.
			If Not Err Then
				For Each strGroup In objGroupList
					Set objGroup = GetObject("LDAP://" & strGroup)  
					objGroup.PutEx ADS_PROPERTY_DELETE,"member", Array(objUser.Get("distinguishedName")) 
					objGroup.SetInfo 
					Set objGroup = Nothing
				Next
			Else
				Err.Clear
			End If
			
			'Add the user to the correct groups
			arrUserData = GetUserDataFromDatabase(strUserName)
			strSQL = "SELECT DN FROM GroupMappings WHERE RoleID=" & arrUserData(CLASSOF) & " AND Site='" & arrUserData(SITE) & "'"
			Set objGroups = objDBConnection.Execute(strSQL)
			If Not objGroups.EOF Then
				Do Until objGroups.EOF
					Set objGroup = GetObject("LDAP://" & objGroups(0))
					objGroup.Add(objUser.ADSPath)
					Set objGroup = Nothing
					objGroups.MoveNext
				Loop
			End If
			
			'Close open objects
			Set objOU = Nothing
			Set objUserLookup = Nothing
			Set objRootDSE = Nothing
			
		End If
	End If

End Sub

Sub UpdateUserUserName(strUserName,strNewUserName)

	'This will update the last name on a user's account
	
	Dim objRootDSE, objUserLookUp, objUser, objUserConflictCheck, strSQL, arrUserData, bolError
	
	bolError = False
	
	'Make sure the new username isn't already in use
	If Not ExistsInActiveDirectory(strNewUserName) Then
	
		'Make sure the user account exists before trying to update it
  		If ExistsInActiveDirectory(strUserName) Then
		
			'Make sure they didn't subit a blank value
			If strNewUserName <> "" Then
	
				'Make sure the provided input isn't crazy long.
				If Len(strNewUserName) <= 50 Then 
			
					'Create a RootDSE object for the domain
					Set objRootDSE = GetObject("LDAP://RootDSE")
	
					'Get the user object from Active Directory	
					objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
					">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
					Set objUserLookup = objADCommand.Execute
	
					'Build the user object
					Set objUser = GetObject("LDAP://" & objUserLookup(0))
	
					objUser.Put "sAMAccountName", strNewUserName
					objUser.Put "userPrincipalName", strNewUserName & "@" & strDomain
					objUser.Put "mail", strNewUserName & "@" & strDomain
					objUser.SetInfo
				
				Else
					bolError = True 'User name is too long
				End If
			Else
				bolError = True 'User name is blank
			End If
		Else
			bolError = True 'User not in Active Directory
		End If 
	Else 
		bolError = True 'New username already exists in Active Directory
	End If
	
	If bolError Then
	
		'Change the username back in the database
		UpdateLog "UserUpdatedUserName","",strUserName,strNewUserName,strUserName,""
		strSQL = "UPDATE People SET UserName='" & strUserName & "' WHERE UserName='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Log SET UserName='" & strUserName & "' WHERE UserName='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Log SET UpdatedBy='" & strUserName & "' WHERE UpdatedBy='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Devices SET UserName='" & strUserName & "' WHERE UserName='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Devices SET LastUser='" & strUserName & "' WHERE LastUser='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Assignments SET ReturnedBy='" & strUserName & "' WHERE ReturnedBy='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Assignments SET IssuedBy='" & strUserName & "' WHERE IssuedBy='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Events SET EnteredBy='" & strUserName & "' WHERE EnteredBy='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Events SET CompletedBy='" & strUserName & "' WHERE CompletedBy='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE Sessions SET Username='" & strUserName & "' WHERE Username='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		strSQL = "UPDATE PendingTasks SET Username='" & strUserName & "' WHERE Username='" & strNewUserName & "'"
		objDBConnection.Execute(strSQL)
		arrUserData = GetUserDataFromDatabase(strUserName)
		SendEmail "Matt", "hullm@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Dane", "davisd@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
		SendEmail "Janine", "wayj@lkgeorge.org", "UserNotUpdatedInAD", arrUserData
	
	End If
	   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing

End Sub

Sub TogglePasswordStatus(strUserName, bolPasswordExpire)

	Dim objRootDSE, objUserLookup, objUser
	
	'Make sure the user account exists before trying to update it
  	If ExistsInActiveDirectory(strUserName) Then
  	
  		'Create a RootDSE object for the domain
		Set objRootDSE = GetObject("LDAP://RootDSE")

		'Get the user object from Active Directory	
		objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
		">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
		Set objUserLookup = objADCommand.Execute

		'Build the user object
		Set objUser = GetObject("LDAP://" & objUserLookup(0))
		
		'Either disable or enable the password expire option
		If bolPasswordExpire Then
			objUser.Put "userAccountControl", 512 'Normal account, password expires
		Else
			objUser.Put "userAccountControl", 66048 'Normal account, password doesn't expire
		End If
		objUser.SetInfo
  	
  	End If

End Sub

Sub EnableInActiveDirectory(strUserName)
	
	'This will disable a user in Active Directory and set the description to the reason why
	'the account was disabled.
	
	Dim objRootDSE, objUserLookUp, objUser, arrUserData, strSQL, objGroups, objGroup, objPaperCut
	
	'Get the user's information
	arrUserData = GetUserDataFromDatabase(strUserName)
	
	'Create a RootDSE object for the domain
	 Set objRootDSE = GetObject("LDAP://RootDSE")

	'Get the user object from Active Directory	
	objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
	">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
	Set objUserLookup = objADCommand.Execute
	
	'Build the user object
   Set objUser = GetObject("LDAP://" & objUserLookup(0))
   
   'Enable the user
   objUser.Put "UserAccountControl", 66048 'Account Enabled
   
   'Update the description if the user is a student
   If arrUserData(CLASSOF) > 1000 Then
   	objUser.Put "Description", "Class of 20" & Left(objUser.SAMAccountName,2)
   	
   	'Add the student to the PaperCut group if it's set
		If strPaperCut <> "" Then
			Set objPaperCut = GetObject("LDAP://" & strPaperCut)
			objPaperCut.Add(objUser.ADSPath)
			Set objPaperCut = Nothing
		End If
		
		'Add the student to the students group
		Set objGroup = GetObject("LDAP://" & strGroupRoot)
   	If objGroup.IsMember("LDAP://" & objUser.distinguishedName) = False Then
         objGroup.Add(objUser.ADSPath)
      End If
      Set objGroup = Nothing
      'If Err Then
      '   MsgBox objUser.sAMAccountName
      'End If

   Else
   
		'Add the user to the correct groups
		strSQL = "SELECT DN FROM GroupMappings WHERE RoleID=" & arrUserData(CLASSOF) & " AND Site='" & arrUserData(SITE) & "'"
		Set objGroups = objDBConnection.Execute(strSQL)
		If Not objGroups.EOF Then
			Do Until objGroups.EOF
				Set objGroup = GetObject("LDAP://" & objGroups(0))
				'objGroup.Add(objUser.ADSPath)  FIX THIS!!!!!!!
				Set objGroup = Nothing
				objGroups.MoveNext
			Loop
		End If
   
   End If
   
   'Save the changes
   objUser.SetInfo
   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing
	
End Sub 

Sub DisableInActiveDirectory(strUserName, strReason)
	
	'This will disable a user in Active Directory and set the description to the reason why
	'the account was disabled.
	
	Dim objRootDSE, objUserLookUp, objUser, strSQL, arrUserData, objGroups, objGroup, objPaperCut
	
	'Create a RootDSE object for the domain
	Set objRootDSE = GetObject("LDAP://RootDSE")

	'Get the user object from Active Directory	
	objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
	">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
	Set objUserLookup = objADCommand.Execute

	'Build the user object
	Set objUser = GetObject("LDAP://" & objUserLookup(0))

	objUser.Put "UserAccountControl", 66050 'Account Disabled

	'Set the description based on the reason
	Select Case strReason
		Case "AUP"
			objUser.Put "Description", "Class of 20" & Left(objUser.SAMAccountName,2) & " - Missing AUP"
		Case "SchoolTool"
			objUser.Put "Description", "Class of 20" & Left(objUser.SAMAccountName,2) & " - Not Active In SchoolTool"
	End Select

	objUser.SetInfo
	
	'Get the user's information
	arrUserData = GetUserDataFromDatabase(strUserName)
	
	'Take care of group memberships
	If arrUserData(CLASSOF) > 1000 Then

		'Remove the student from the PaperCut group if it's set
		If strPaperCut <> "" Then
			Set objPaperCut = GetObject("LDAP://" & strPaperCut)
			objPaperCut.Remove(objUser.ADSPath)
			Set objPaperCut = Nothing
		End If
		
		'Remove the student from the students group
		'Set objGroup = GetObject("LDAP://" & strGroupRoot)
   	'objGroup.Remove(objUser.ADSPath)
	
	Else
	
		'Remove the user to the groups
		strSQL = "SELECT DN FROM GroupMappings WHERE RoleID=" & arrUserData(CLASSOF) & " AND Site='" & arrUserData(SITE) & "'"
		Set objGroups = objDBConnection.Execute(strSQL)
		If Not objGroups.EOF Then
			Do Until objGroups.EOF
				Set objGroup = GetObject("LDAP://" & objGroups(0))
				'objGroup.Remove(objUser.ADSPath)
				Set objGroup = Nothing
				objGroups.MoveNext
			Loop
		End If
	
	End If
	
	'Turn off the password from expiring in the database so we don't get an AD/DB conflit
	strSQL = "UPDATE People SET PWordNeverExpires=True WHERE UserName='" &  strUserName & "'"
   objDBConnection.Execute(strSQL)
   
   'Close objects
   Set objRootDSE = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing
	
End Sub 

Function IsActiveInActiveDirectory(strUserName)

	'This function will look up a user in Active Directory and return true if the account is
	'enabled, and false if it's disabled.

	Dim objRootDSE, objUserLookup

	'Create a RootDSE object for the domain
   Set objRootDSE = GetObject("LDAP://RootDSE")
   
   'Get the user object from Active Diretory
	objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
	">;(&(objectClass=user)(samAccountName=" & strUserName & "));UserAccountControl"
	Set objUserLookup = objADCommand.Execute

	If Not objUserLookup.EOF Then
	
		Select Case objUserLookup(0)
			Case 66050, 514
				IsActiveInActiveDirectory = False
			Case 66048, 512
				IsActiveInActiveDirectory = True
		End Select 
	
	Else 
		IsActiveInActiveDirectory = False
	End If
	
	'Close objects
	Set objRootDSE = Nothing
	Set objUserLookup = Nothing

End Function

Function ExistsInActiveDirectory(strUserName)

	'This will check and see if a user exists in Active Directory.  It will return a 
	'true or false.

	Dim objRootDSE, objUserLookup

	'Create a RootDSE object for the domain
	 Set objRootDSE = GetObject("LDAP://RootDSE")
	
	'Get the user object from Active Directory	
	objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
	">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
	Set objUserLookup = objADCommand.Execute
	
	'Find out if the account exists in Active Directory
	If objUserLookup.EOF Then
		ExistsInActiveDirectory = False
	Else
		ExistsInActiveDirectory = True
	End If
	
	'Close object
	Set objRootDSE = Nothing
	
End Function

Function ExistsInDatabase(intStudentID,intClassOf)

	'This will check and see if a user exists in the database, it will also correct a students 
	'class off setting if their grade has changed.  The function will return a true or false.

	Dim strSQL, objStudent

   'Check and see if they are in the students table
   strSQL = "SELECT ID,ClassOf" & vbCRLF
   strSQL = strSQL & "FROM People" & vbCRLF
   strSQL = strSQL & "WHERE ClassOf>2000 AND StudentID=" & intStudentID
   Set objStudent = objDBConnection.Execute(strSQL)
   
   'If they aren't in the people table then they aren't in the database.
   If objStudent.EOF Then
      ExistsInDatabase = False
   Else

      'If the course has changed in the export file then change the user account in AD and the database.
      If CInt(intClassOf) <> CInt(objStudent(1)) Then 
         RollBackStudent objStudent(0),intClassOf
      End If
      
      ExistsInDatabase = True
   End If

End Function

Sub ValidateADAccounts(strOU)

	Dim objOU, objUser, UserData, objUserActive, strStatus, strSQL, strUser, bolKeepEnabled, arrUserData

	'Create the OU object
	Set objOU = GetObject("LDAP://" & strOU)
	
	'Loop through each object in the OU
	For Each objUser in objOU
	
		'Reset the status indicator
		strStatus = ""
	
		'Find out if the object is a user or OU
		Select Case objUser.Class
		
			'If it's an OU recursively call the sub  
			Case "organizationalUnit"
				
				'The first case is the list of OU's you want to ignore
				Select Case objUser.Name
					Case "OU=Misc"
					Case Else
						ValidateADAccounts objUser.DistinguishedName
				End Select
			
			'If it's a user then see if they need to be disabled
			Case "user"
				
				'Find out if the user is active in the database
				strSQL = "SELECT Active,AUP FROM People WHERE UserName='" & objUser.SamAccountName & "'"
				Set objUserActive = objDBConnection.Execute(strSQL)
				
				If Not objUserActive.EOF Then
					
					'If the user is active in the database then make sure their AD account is active
					If objUserActive(0) Then
						If objUser.UserAccountControl <> 66048 Then
						
							'Make sure they have in their AUP before enabling the account
							If objUserActive(1) Then
								strStatus = "Enabled"
								EnableInActiveDirectory objUser.SamAccountName
							
							'Enable the account if it's before the start of school, or before the first 2 weeks
							ElseIf DateDiff("d",datFirstDayOfSchool,Date) <= 14 Then
								strStatus = "Enabled"
								EnableInActiveDirectory objUser.SamAccountName
								
							Else
							
								'Get the information needed to send the email messages
								'UserData = GetUserDataFromDatabase(objUser.SamAccountName)
							
								'SendEmail "Natalie", "fullenn@lkgeorge.org", "ReturningStudentMissingAUP", arrUserData
								'SendEmail "Matt", "hullm@lkgeorge.org", "ReturningStudentMissingAUP", arrUserData
								'SendEmail "Janine", "wayj@lkgeorge.org", "ReturningStudentMissingAUP", arrUserData
								
							End If
						End If
					
					'Disabled the account in AD if they aren't active in the database
					Else
						If objUser.UserAccountControl <> 66050 Then
							strStatus = "Disabled"
							bolKeepEnabled = False
							For Each strUser in objKeepEnabledInAD
								If CDate(objKeepEnabledInAD.Item(strUser)) >= Date() Then
									If strUser = objUser.SamAccountName Then
										bolKeepEnabled = True
									End If
								End If
							Next
							
							If Not bolKeepEnabled Then
								DisableInActiveDirectory objUser.SamAccountName, "SchoolTool"
							End If
						End If
				
					End If
					
					'If the status indicator was changed then do what needs to be done
					If strStatus <> "" Then
						
						'Get the information needed to send the email messages
						arrUserData = GetUserDataFromDatabase(objUser.SamAccountName)

						
						Select Case strStatus
							
							Case "Enabled"
								UpdateLog "AccountEnabledSchoolTool","",objUser.SamAccountName,"","",""
								SendEmail "Natalie", "fullenn@lkgeorge.org", "EnabledSchoolToolAdmin", arrUserData
								SendEmail "Matt", "hullm@lkgeorge.org", "EnabledSchoolToolAdmin", arrUserData
								SendEmail "Rene","palmerr@lkgeorge.org", "EnabledSchoolToolAdmin", arrUserData
								SendEmail "Janine","wayj@lkgeorge.org", "EnabledSchoolToolAdmin", arrUserData
								SendEMailToTeachers "EnabledSchoolTool", arrUserData
							
							Case "Disabled"
								If Not bolKeepEnabled Then
									UpdateLog "AccountDisabledSchoolTool","",objUser.SamAccountName,"","",""
									SendEmail "Natalie", "fullenn@lkgeorge.org", "DisabledSchoolToolAdmin", arrUserData
									SendEmail "Matt", "hullm@lkgeorge.org", "DisabledSchoolToolAdmin", arrUserData
									SendEmail "Rene","palmerr@lkgeorge.org", "DisabledSchoolToolAdmin", arrUserData
									SendEmail "Janine","wayj@lkgeorge.org", "DisabledSchoolToolAdmin", arrUserData
									SendEMailToTeachers "DisabledSchoolTool", arrUserData
								End If
						End Select 
						
					End If
				
				'Disable the account if they aren't found in the database	
				Else
					If objUser.UserAccountControl <> 66050 Then
						UpdateLog "AccountDisabledSchoolTool","",objUser.SamAccountName,"","",""
						DisableInActiveDirectory objUser.SamAccountName, "SchoolTool"
					End If
				End If
					
		End Select
				
	Next
	
	'Close objects
	Set objOU = Nothing
	Set objUserActive = Nothing
	Set objUser = Nothing

End Sub

Sub VerifyADandDBMatch

	'This will sync the phone numbers stored in the database to Active Directory

	On Error Resume Next

	Dim objADCommand, objDBConnection, strSQL, objDBUsers, objADUser, objRootDSE, objUserLookup, bolError, bolADActive
	Dim strUserName, strFirstName, strLastName, strDescription, strRoom, strPhone, bolActive, strSite, intClassOf
	Dim bolPWordNeverExpires, arrUserData, strUserAccountControl, strDBOU
	
	'Create a connection to the database
	Set objDBConnection = ConnectToDatabase
    
	'Get the list of users in the database with phone numbers
	strSQL = "SELECT UserName,FirstName,LastName,Description,RoomNumber,PhoneNumber,Active,Site,ClassOf,PWordNeverExpires FROM People WHERE Role='Teacher'"
	Set objDBUsers = objDBConnection.Execute(strSQL)
	
	If Not objDBUsers.EOF Then
		
		'Create the domain objects
		Set objADCommand = ConnectToActiveDirectory
		Set objRootDSE = GetObject("LDAP://RootDSE")
		
		'Loop through each user in the database with a phone number
		Do Until objDBUsers.EOF
		
			bolError = False
		
			'Get the data from the object and assign it to variables
			strUserName = objDBUsers(0)
			strFirstName = objDBUsers(1)
			strLastName = objDBUsers(2)
			strDescription = objDBUsers(3)
			strRoom = objDBUsers(4)
			strPhone = objDBUsers(5)
			bolActive = objDBUsers(6)
			strSite = objDBUsers(7)
			intClassOf = objDBUsers(8)
			bolPWordNeverExpires = objDBUsers(9)

			'Get the user's distinguished name from Active Directory	
			objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
			">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
			Set objUserLookup = objADCommand.Execute
		
			'Build the user object
			Set objADUser = GetObject("LDAP://" & objUserLookup(0))

			'Check for differences
			If LCase(objADUser.Get("samAccountName")) <> LCase(strUserName) Then
				bolError = True
			End If
			If objADUser.Get("givenName") <> strFirstName Then
				If Err Then
					If strFirstName <> "" Then
						bolError = True
						Err.Clear
					End If
				Else
					bolError = True
				End If
			End If
			If objADUser.Get("sn") <> strLastName Then
				If Err Then
					If strLastName <> "" Then
						bolError = True
						Err.Clear
					End If
				Else
					bolError = True
				End If
			End If
			If objADUser.Get("description") <> strDescription Then
				If Err Then
					If strDescription <> "" Then
						bolError = True
						Err.Clear
					End If
				Else
					bolError = True
				End If
			End If
			If objADUser.Get("physicalDeliveryOfficeName") <> strRoom Then
				If Err Then
					If strRoom <> "" Then
						bolError = True
						Err.Clear
					End If
				Else
					bolError = True
				End If
			End If
			If objADUser.Get("telephoneNumber") <> strPhone Then
				If Err Then
					If strPhone <> "" Then
						bolError = True
						Err.Clear
					End If
				Else
					bolError = True
				End If
			End If
			If objADUser.Get("ipPhone") <> strPhone Then
				If Err Then
					If strPhone <> "" Then
						bolError = True
						Err.Clear
					End If
				Else
					bolError = True
				End If
			End If
			
			'Build what the database thinks the OU should be for the user
			strDBOU = Replace(strOUBase,"%SITE%",strSite)
			strDBOU = Replace(strDBOU,"%ROLE%",GetRole(intClassOf))
			strDBOU = Replace(strDBOU,"Elementary","Elementary School")
			strDBOU = "CN=" & strFirstName & " " & strLastName & "," & strDBOU
			
			'Make sure the user is in the correct OU
			If objUserLookup(0) <> strDBOU Then
				bolError = True
			End If
			
			'Make sure the active status is the same in both
			strUserAccountControl = objADUser.Get("userAccountControl")
			Select Case strUserAccountControl
			
				Case 66048, 512
					bolADActive = True
					
				Case Else
					bolADActive = False
					
			End Select
			If bolADActive <> bolActive Then
				bolError = True
			End If
			
			'See if their password setting is right
			If bolActive Then
				If bolPWordNeverExpires Then
					If strUserAccountControl <> 66048 Then
						bolError = True
					End If
				Else
					If strUserAccountControl <> 512 Then
						bolError = True
					End If
				End If
			Else
				If bolPWordNeverExpires Then
					If strUserAccountControl <> 66050 Then
						bolError = True
					End If
				Else
					If strUserAccountControl <> 514 Then
						bolError = True
					End If
				End If
			End If
		
			'Something went wrong, let people know
			If bolError Then 
				arrUserData = GetUserDataFromDatabase(strUserName)
				SendEmail "Matt", "hullm@lkgeorge.org", "UserNotInSync", arrUserData
				SendEmail "Dane", "davisd@lkgeorge.org", "UserNotInSync", arrUserData
				SendEmail "Janine", "wayj@lkgeorge.org", "UserNotInSync", arrUserData
			End If
		
			'Save the values to Active Directory
			objADUser.SetInfo
		
			'Move to the next user
			objDBUsers.MoveNext
			
		Loop
		
		'Close open objects
		Set objRootDSE = Nothing
		Set objADCommand = Nothing
		
	End If

	'Close open objects
	Set objDBConnection = Nothing

End Sub


Sub RollBackStudent(intStudentID,intCourse)

	'This will update a student's grade if their graduating year changes.  It will move their account
	'in Active Directory, rename their home folder, and update the database.  It will move a student
	'either back, or forward a grade depending on the change.

   Dim strSQL, objFixStudent, strNewUserName, strNewHome, strNewEMail, strNewUserPrincipalName, strOU
   Dim strNewSAMAccountName, strNewDescription, objUserLookup, objUser, strOldUserName, objRootDSE
   Dim objOU, objFSO, arrUserData

   'Get the student from the database
   strSQL = "SELECT UserName" & vbCRLF
   strSQL = strSQL & "FROM People" & vbCRLF
   strSQL = strSQL & "WHERE ID=" & intStudentID
   Set objFixStudent = objDBConnection.Execute(strSQL)
   
   'Set the new variables
   strNewUserName = Right(intCourse,2) & Right(objFixStudent(0),Len(objFixStudent(0)) - 2) 
   strNewHome = strStudentShare & strNewUserName
   strNewEMail = strNewUserName & "@" & strDomain
   strNewUserPrincipalName = strNewEMail
   strNewSAMAccountName = strNewUserName
   strNewDescription = "Class of " & intCourse
   strOldUserName = objFixStudent(0)

	'Create a RootDSE object for the domain
   Set objRootDSE = GetObject("LDAP://RootDSE")
   
   'Get the user object from Active Diretory
   objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
   ">;(&(objectClass=user)(samAccountName=" & strOldUserName & "));distinguishedName"
   Set objUserLookup = objADCommand.Execute
  
   'Build the user object
   Set objUser = GetObject("LDAP://" & objUserLookup(0))
   
   'Set the users properties
   'objUser.Put "homeDirectory", strNewHome
   objUser.Put "mail", strNewEMail
   objUser.Put "sAMAccountName", strNewUserName
   objUser.Put "userPrincipalName", strNewUserPrincipalName
   objUser.Put "description", strNewDescription
   objUser.SetInfo
   
   'Move the user to the correct OU
   strOU = "LDAP://OU=" & intCourse & "," & strStudentOU
   Set objOU = GetObject(strOU)
   objOU.MoveHere "LDAP://" & objUserLookup(0),vbNullString
   
   'Rename the home folder
   'Set objFSO = CreateObject("Scripting.FileSystemObject")
   'objFSO.MoveFolder strStudentShare & strOldUserName, strStudentShare & strNewUserName
   
   'Update the inventory database.
   strSQL = "UPDATE People SET Username='" & strNewUserName & "',ClassOf='" & intCourse & "' WHERE Username='" & strOldUserName & "'"
   objDBConnection.Execute(strSQL)
   strSQL = "UPDATE Log SET Username='" & strNewUserName & "' WHERE Username='" & strOldUserName & "'"
   objDBConnection.Execute(strSQL)
   
   UpdateLog "StudentGradeChange","",strNewUserName,strOldUserName,strNewUserName,""
   
   'Send the email
   arrUserData = GetUserDataFromDatabase(strNewUserName)
   SendEmail "Natalie", "fullenn@lkgeorge.org", "GradeChangeAdmin", arrUserData
   SendEmail "Matt", "hullm@lkgeorge.org", "GradeChangeAdmin", arrUserData
   SendEmail "Dane", "davisd@lkgeorge.org", "GradeChangeAdmin", arrUserData
   SendEmail "Janine", "wayj@lkgeorge.org", "GradeChangeAdmin", arrUserData
   
   'Close objects
   Set objFixStudent = Nothing
   Set objUserLookup = Nothing
   Set objUser = Nothing
   Set objOU = Nothing

End Sub

Sub FixDestinyExport

	Dim objFSO, strCurrentFolder, strDestinyData, txtSourceCSV, txtDestCSV, strImportedData, strOutPut
	Dim arrData, intStudentID, intCounter, strItem

	'Get the CSV path
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCurrentFolder = objFSO.GetAbsolutePathName(".")
	strCurrentFolder = strCurrentFolder & "\CSV\"
	strDestinyData = strCurrentFolder & "DestinyData.csv"

	'Open the source and destination files
	Set txtSourceCSV = objFSO.OpenTextFile(strDestinyData)
	Set txtDestCSV = objFSO.CreateTextFile(strDestinyExport)
	 
	'Loop through each line 
	While txtSourceCSV.AtEndOfLine = False

		'Grab the active line
		strImportedData = txtSourceCSV.ReadLine
	
		'Split the line into an array using the comma
		arrData = Split(strImportedData,",")
	
		'Grab the student ID from the data
		intStudentID = Replace(arrData(0),"""","")
	
		'Recreate the data replacing the proposed username with the correct one on the way.
		strOutput = ""
		intCounter = 0
		
		For Each strItem in arrData
	
			'If we are on element 18 then we need to get the correct username from the database
			If intCounter = 18 Then
				strOutPut = strOutput & """" & GetUserName(intStudentID) & ""","
			Else
				strOutput = strOutput & strItem & ","
			End If
		
			'Increase the counter by 1
			intCounter = intCounter + 1
		
		Next
	
		'Remove the end comma
		strOutput = Left(strOutPut,Len(strOutPut) - 1)
	
		'Write the output to a file
		txtDestCSV.Writeline strOutput
	
	Wend

	'Close objects
	Set objFSO = Nothing
	Set txtSourceCSV = Nothing
	Set txtDestCSV = Nothing

End Sub

Function GetUserName(intStudentID)

	'This will connect to the database and return the username for the student

	Dim strSQL, objStudent

   'Check and see if they are in the people table
   strSQL = "SELECT UserName" & vbCRLF
   strSQL = strSQL & "FROM People" & vbCRLF
   strSQL = strSQL & "WHERE Role='Student' AND StudentID=" & intStudentID
   Set objStudent = objDBConnection.Execute(strSQL)  
   
   'Return the status
   If objStudent.EOF Then
      GetUserName = ""
   Else
      GetUserName = objStudent(0)
   End If

End Function

Sub SendEmail(strFirstName, strEMail, strType, arrUserData)

	'This will send out an email

	Dim strSMTPPickupFolder, objMessage, objConf, strMessage, strSubject, bolHTMLMEssage
	Dim bolSendAsAdmin

   Const cdoSendUsingPickup = 1

   strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"
   
   'Get the message body
   strMessage =  GetEMailMessage(strFirstName, strType, arrUserData)
   
   'Set the subject
   Select Case strType
   	Case "NewStudentFound"	
   		strSubject = "New Student Found"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = False
   	Case "NewStudentReady", "NewStudentReadyAdmin"
   		strSubject = "New Student Account Created"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = False
   	Case "EnabledSchoolTool", "EnabledSchoolToolAdmin"
   		strSubject = "Student Account Enabled - Student Returned"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = False
   	Case "DisabledSchoolTool", "DisabledSchoolToolAdmin"
   		strSubject = "Student Account Disabled - Student Left District"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = False
   	Case "EnabledAUP", "EnabledAUPAdmin"
   		strSubject = "Student Account Enabled - AUP Turned In"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = False
   	Case "DisabledAUP", "DisabledAUPAdmin"
   		strSubject = "Student Account Disabled - Missing AUP"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = False
   	Case "GradeChange", "GradeChangeAdmin"
   		strSubject = "Student Grade Changed"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = False
   	Case "PasswordExpired"
   		strSubject = "Your Password Has Expired"
   		bolHTMLMEssage = True
   		bolSendAsAdmin = True
   	Case "PasswordExpiring"
   		strSubject = "Password Will Expire in " & arrUserData(PASSWORD) & " Days"
   		bolHTMLMEssage = True
   		bolSendAsAdmin = True
   	Case "PasswordExpiresToday"
   		strSubject = "Password Expires Today"
   		bolHTMLMEssage = True
   		bolSendAsAdmin = True
   	Case "PasswordExpiringAdmin"
   		strSubject = "Passwords About to Expire"
   		bolHTMLMEssage = True
   		bolSendAsAdmin = True
   	Case "EndOfYearPasswordReset"
   		strSubject = "End Of Year Password Change"
   		bolHTMLMEssage = True
   		bolSendAsAdmin = True
   	Case "UserNotUpdatedInAD"
   		strSubject = "User Update Failed in Active Directory"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = True
   	Case "UserNotInSync"	
   		strSubject = "Inventory Site and Active Directory Out Of Sync"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = True
   	Case "NewAccountCreated"
   		strSubject = "New User Account Created in Active Directory"
   		bolHTMLMEssage = False
   		bolSendAsAdmin = True
   End Select

   'Create the objects required to send the mail.
   Set objMessage = CreateObject("CDO.Message")
   Set objConf = objMessage.Configuration
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
      .Update
   End With
   
   'Set who the message is from.
   If bolSendAsAdmin Then
   	objMessage.From = strFromEMailAdmin
   Else
   	objMessage.From = strFromEMail
   	objMessage.BCC = strFromEMail
   End If
   
   'Send the message
   If strEMailOverride = "" Then
   	objMessage.To = strEMail
   Else
   	objMessage.To = strEMailOverride
   	objMessage.BCC = ""
   End If
   objMessage.Subject = strSubject
   If bolHTMLMEssage Then
		objMessage.HTMLBody = strMessage
	Else
		objMessage.TextBody = strMessage
	End If
	If bolEnableEmail Then
   	objMessage.Send
   End If
   
   'Close objects
   Set objMessage = Nothing
   Set objConf = Nothing
   
End Sub

Function GetEMailMessage(strFirstName,strType,arrUserData)

	'This will return the requested stored email message

	Dim objFSO, strEMailPath, strMessage, txtEMailMessage, arrUsers, strUser, arrData
	
	'Get the email message
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strEMailPath = objFSO.GetAbsolutePathName(".")
	strEMailPath = strEMailPath & "\EMail\"
	Set txtEMailMessage = objFSO.OpenTextFile(strEMailPath & strType & ".txt")
	
	'Read in the stored email message
	strMessage = txtEMailMessage.ReadAll
	txtEMailMessage.Close

	'Change the variables in the message to their values
	strMessage = Replace(strMessage,"#RECIPIENT#",strFirstName)
	strMessage = Replace(strMessage,"#FIRSTNAME#", arrUserData(FIRSTNAME))
	strMessage = Replace(strMessage,"#LASTNAME#", arrUserData(LASTNAME))
	strMessage = Replace(strMessage,"#USERNAME#", arrUserData(USERNAME))
	
	strMessage = Replace(strMessage,"#STUDENTURL#", strURL & "user.asp?UserName=" & arrUserData(USERNAME))
	strMessage = Replace(strMessage,"#STUDENTURLADMIN#", strURL & "admin/user.asp?UserName=" & arrUserData(USERNAME))
	strMessage = Replace(strMessage,"#ADDURL#",strURL & "admin/add.asp")
	 
	Select Case strType
		
		Case "PasswordExpiringAdmin"
	
			arrUsers = arrUserData(PASSWORD)
			For Each strUser in arrUsers 
				arrData = Split(strUser,":")
				If arrData(1) <= 1 Then
					strMessage = strMessage & "<tr><td bgcolor=""Red"">" & arrData(0) & "</td>" & vbCRLF
					strMessage = strMessage & "<td align=""center"" bgcolor=""Red"">" & arrData(1) & "</td></tr>" & vbCRLF
				ElseIf arrData(1) <= 3 Then
					strMessage = strMessage & "<tr><td bgcolor=""Yellow"">" & arrData(0) & "</td>" & vbCRLF
					strMessage = strMessage & "<td align=""center"" bgcolor=""Yellow"">" & arrData(1) & "</td></tr>" & vbCRLF
				Else
					strMessage = strMessage & "<tr><td>" & arrData(0) & "</td>" & vbCRLF
					strMessage = strMessage & "<td align=""center"">" & arrData(1) & "</td></tr>" & vbCRLF
				End If
			Next
			strMessage = strMessage & "</table>"
	
		Case "PasswordExpiring", "PasswordExpiresToday"
	
			'We're using the Password field to store the array needed to send the email to admins, putting this
			'in the else section prevents an error when doing the replace.
			If Not IsNull(arrUserData(PASSWORD)) Then
				strMessage = Replace(strMessage,"#PASSWORD#", arrUserData(PASSWORD))
				strMessage = Replace(strMessage,"#DAYSREMAINING#",arrUserData(PASSWORD))
				If IsDate(arrUserData(PASSWORD)) Then
					strMessage = Replace(strMessage,"#PASSWORDRESETTIME#",FormatDateTime(arrUserData(PASSWORD),3))
				End If
			End If
			
		Case Else
		
			If Not IsNull(arrUserData(PASSWORD)) Then
				strMessage = Replace(strMessage,"#PASSWORD#", arrUserData(PASSWORD))
			End If
	
	End Select
	
	GetEMailMessage = strMessage
	
	'Close objects
	Set objFSO = Nothing
	Set txtEMailMessage = Nothing
	
End Function

Sub SendEMailToTeachers(strMessageType,arrUserData)

	'This sub will send the requested message to the appropriate teachers.  Some classes have one teacher
	'some have two.  This will split email list if needed and send to each teacher.

	Dim arrEMails, arrTeacherNames, intCurrentGrade
	
	'If there are two teachers then there will be a ; in the email string, if found send to both teachers
	If InStr(arrUserData(HOMEROOMEMAIL),";") <> 0 Then
		arrEMails = Split(arrUserData(HOMEROOMEMAIL),";")
		arrTeacherNames = Split(arrUserData(HOMEROOM),"/")
		SendEmail arrTeacherNames(0), arrEMails(0), strMessageType, arrUserData
		SendEmail Trim(arrTeacherNames(1)), arrEMails(1), strMessageType, arrUserData
	Else
		arrTeacherNames = Split(arrUserData(HOMEROOM),",")
		SendEmail Trim(arrTeacherNames(1)), arrUserData(HOMEROOMEMAIL), strMessageType, arrUserData
	End If
	
	'Get their grade, change K to a 0
	If IsNumeric(GetGrade(arrUserData(CLASSOF))) Then
		intCurrentGrade = GetGrade(arrUserData(CLASSOF))
	Else
		intCurrentGrade = 0
	End If
	
	'EMail the building specific people
	If intCurrentGrade >= intHSBegins Then
		SendEmail "Sarah", "olsons@lkgeorge.org", strMessageType, arrUserData
		SendEmail "Dane", "davisd@lkgeorge.org", strMessageType, arrUserData
	Else
		SendEmail "Bridget", "crossmanb@lkgeorge.org", strMessageType, arrUserData
		'SendEmail "Matt", "hullm@lkgeorge.org", strMessageType, arrUserData
	End If

End Sub

Sub UpdatePasswordExpirationDate(bolSendEmail)

	'This will update the PWordLastSet value in the People table using info from Active Directory

   Dim strSQL, objUsers, datPwdLastSet, datExpires, intDaysUntilAccountExpires, strUserList
   Dim arrUserData, arrUserList

   'Grab the list of active users from the database
   strSQL = "SELECT UserName,PWordLastSet,ClassOf FROM People"
   Set objUsers = objDBConnection.Execute(strSQL)

   If Not objUsers.EOF Then
      Do Until objUsers.EOF
      
      	'Get the date the password was last set and calculate how many days remain until they need
      	'to change it.
      	datPwdLastSet = GetLastPasswordSet(objUsers(0))
			datExpires = DateAdd("d",intPasswordLife,datPwdLastSet)
			intDaysUntilAccountExpires = DateDiff("d",Date,datExpires)
      	
      	'Send the user an email if their account is about to expire, and build this list of users for the 
      	'email that will be sent to the admins.
      	If bolSendEmail Then
				arrUserData = GetUserDataFromDatabase(objUsers(0))
				If arrUserData(CLASSOF) < 2000 Then 'Only send email to adults.
					If PasswordExpires(arrUserData(USERNAME)) Then 'Only send if their password is set to expire.
						If intDaysUntilAccountExpires >= -10 And intDaysUntilAccountExpires <= 10 Then
							arrUserData = GetUserDataFromDatabase(objUsers(0))
							arrUserData(PASSWORD) = intDaysUntilAccountExpires 'Use the password field to store the days remaining
							If intDaysUntilAccountExpires > 0 Then
								SendEmail arrUserData(FIRSTNAME), arrUserData(USERNAME) & "@" & strDomain, "PasswordExpiring", arrUserData
							ElseIf intDaysUntilAccountExpires < 0 Then
								SendEmail arrUserData(FIRSTNAME), arrUserData(USERNAME) & "@" & strDomain, "PasswordExpired", arrUserData
							Else
								arrUserData(PASSWORD) = datExpires
								SendEmail arrUserData(FIRSTNAME), arrUserData(USERNAME) & "@" & strDomain, "PasswordExpiresToday", arrUserData
							End If
							strUserList = strUserList & arrUserData(LASTNAME) & ", " & arrUserData(FIRSTNAME) & ":" & intDaysUntilAccountExpires & ";"
						End If
					End If
				End If
      	
      		'Send the nag email if it's the end of the year.  This will prompt users to change their password even
      		'if they've recently changed their password.  This is so their password won't expire over the summer.
      		If objUsers(2) < 2000 Then
					If Date() >= CDate(datStartPasswordNag) And Date() <= CDate(datEndPasswordNag) Then
   					If DateDiff("d",datStartPasswordNag,datPwdLastSet) < 0 Then
							arrUserData = GetUserDataFromDatabase(objUsers(0))
							SendEmail arrUserData(FIRSTNAME), arrUserData(USERNAME) & "@" & strDomain, "EndOfYearPasswordReset", arrUserData
						End If
					End If
				End If
				   
      	End If
      	
      	'Send the date the password was last set to the database
         strSQL = "UPDATE People SET PWordLastSet=#" & datPwdLastSet & "# WHERE UserName='" & Replace(objUsers(0),"'","''") & "'"
         objDBConnection.Execute(strSQL)			
         
         objUsers.MoveNext
      Loop
      
      'Send the message to the admins
      If bolSendEmail Then
      	If strUserList <> "" Then
				strUserList = Left(strUserList,Len(strUserList) - 1)
				arrUserList = Split(strUserList,";")
				arrUserList = SortArray(arrUserList)
				arrUserData = Array("","","","","","","","","","","",arrUserList)
				SendEmail "Matt", "hullm@lkgeorge.org", "PasswordExpiringAdmin", arrUserData
				SendEmail "Dane", "davisd@lkgeorge.org", "PasswordExpiringAdmin", arrUserData
				SendEmail "Janine", "wayj@lkgeorge.org", "PasswordExpiringAdmin", arrUserData
			End If
      End If
      
   End If

End Sub

Sub UpdateStudentCountHistory

   'This will update the student count history to the latest numbers if needed.

   Dim strSQL, objStudentCount, intTotalCount, objCurrentCount
   
   'Get the number of students per grade from the database
   strSQL = "SELECT ClassOf, Count(ID) AS CountOfID" & vbCRLF
   strSQL = strSQL & "FROM People" & vbCRLF
   strSQL = strSQL & "GROUP BY ClassOf, Active" & vbCRLF
   strSQL = strSQL & "HAVING ClassOf>2000 AND Active=True"
   Set objStudentCount = objDBConnection.Execute(strSQL)
   
   'Initialize the counter that will be used to count the total number of students
   intTotalCount = 0
   
   'Loop through each grade level and write the count to the database
   If Not objStudentCount.EOF Then
      Do Until objStudentCount.EOF
         intTotalCount = intTotalCount + objStudentCount(1)
         
         'Get the last value from the database
         strSQL = "SELECT StudentCount " & vbCRLF
         strSQL = strSQL & "FROM CountHistory" & vbCRLF
         strSQL = strSQL & "WHERE Role='" & objStudentCount(0) & "'"
         strSQL = strSQL & "ORDER BY ID DESC"
         Set objCurrentCount = objDBConnection.Execute(strSQL)
         
         'If this the first time this is run record the starting data
         If objCurrentCount.EOF Then
            strSQL = "INSERT INTO CountHistory (Role,StudentCount,RecordedDate)" & vbCRLF
            strSQL = strSQL & "VALUES ("
            strSQL = strSQL & "'" & objStudentCount(0) & "',"
            strSQL = strSQL & objStudentCount(1) & ","
            strSQL = strSQL & "#" & Date & "#)"
            objDBConnection.Execute(strSQL)
         Else
            
            'If any value has changed add it to the database
            If objCurrentCount(0) <> objStudentCount(1) Then
               strSQL = "INSERT INTO CountHistory (Role,StudentCount,RecordedDate)" & vbCRLF
               strSQL = strSQL & "VALUES ("
               strSQL = strSQL & "'" & objStudentCount(0) & "',"
               strSQL = strSQL & objStudentCount(1) & ","
               strSQL = strSQL & "#" & Date & "#)"
               objDBConnection.Execute(strSQL)
            End If
         End If
         
         objStudentCount.MoveNext
      Loop
   End If
   
   strSQL = "SELECT StudentCount FROM CountHistory WHERE Role='TotalCount' ORDER BY ID DESC"
   Set objCurrentCount = objDBConnection.Execute(strSQL)
   
   'If this the first time this is run record the starting data
   If objCurrentCount.EOF Then
      strSQL = "INSERT INTO CountHistory (Role,StudentCount,RecordedDate)" & vbCRLF
      strSQL = strSQL & "VALUES ("
      strSQL = strSQL & "'TotalCount',"
      strSQL = strSQL & intTotalCount & ","
      strSQL = strSQL & "#" & Date & "#)"
      objDBConnection.Execute(strSQL)
   Else
   
      'If any value has changed add it to the database
      If objCurrentCount(0) <> intTotalCount Then
         strSQL = "INSERT INTO CountHistory (Role,StudentCount,RecordedDate)" & vbCRLF
         strSQL = strSQL & "VALUES ("
         strSQL = strSQL & "'TotalCount',"
         strSQL = strSQL & intTotalCount & ","
         strSQL = strSQL & "#" & Date & "#)"
         objDBConnection.Execute(strSQL)
      End If
   End If 

End Sub

Sub UpdateCheckInHistory

   'This will update the check in history count to the latest numbers if needed.

   Dim strSQL, objStudentCount, objInsideCheckIns, intTotalCount, objOutsideCheckIns, strHSStudents, intInsideCount, intOutsideCount
   
   strHSStudents = GetGraduationYear(5)
   
   'Get the total count of students per grade
   strSQL = "SELECT ClassOf, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (HomeRoom Is Not Null And HomeRoom  <>'') And ClassOf <= "  & strHSStudents & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentCount = objDBConnection.Execute(strSQL)
	
	'Get the total count of student per grade who have used a school device from the internal network
	strSQL = "SELECT ClassOf, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (HomeRoom Is Not Null And HomeRoom  <>'') And ClassOf <= " & strHSStudents & " AND LastInternalCheckIn=Date()" & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objInsideCheckIns = objDBConnection.Execute(strSQL)
	
	'Get the total count of student per grade who have used a school device from outside of the district
	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (HomeRoom Is Not Null And HomeRoom  <>'') And ClassOf <= " & strHSStudents & " AND LastExternalCheckIn=Date()" & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objOutsideCheckIns = objDBConnection.Execute(strSQL)
   
   'Initialize the counter that will be used to count the total number of students
   intTotalCount = 0
   intInsideCount = 0
   intOutsideCount = 0
   
   'Loop through each grade level and write the count to the database
   If Not objStudentCount.EOF Then
   
   	'Delete any existing entries on the same day
   	strSQL = "DELETE FROM CheckInHistory WHERE RecordedDate=#" & Date & "#"
   	objDBConnection.Execute(strSQL)
   	
      Do Until objStudentCount.EOF
         intTotalCount = intTotalCount + objStudentCount(1)
         
         'Loop through each classes internal checkin count and add it to the inventory
         If Not objInsideCheckIns.EOF Then
         	Do Until objInsideCheckIns.EOF
         		

         		'Make sure you find the right class
         		If objInsideCheckIns(0) = objStudentCount(0) Then
         			
         			'Add the data to the database
         			strSQL = "INSERT INTO CheckInHistory (Role,CheckInCount,TotalStudentCount,RecordedDate,InternalCheckIn) Values (" & vbCRLF
         			strSQL = strSQL & "'" & objInsideCheckIns(0) & "',"
         			strSQL = strSQL & objInsideCheckIns(1) & ","
         			strSQL = strSQL & objStudentCount(1) & ","
         			strSQL = strSQL & "#" & Date & "#,"
         			strSQL = strSQL & "True)"
         			objDBConnection.Execute(strSQL)
         			
         			intInsideCount = intInsideCount + objInsideCheckIns(1)
         			
         		End If
         		
         		objInsideCheckIns.MoveNext
         	Loop
         	objInsideCheckIns.MoveFirst
         End If
         
         'Loop through each classes external checkin count and add it to the inventory
         If Not objOutsideCheckIns.EOF Then
         	Do Until objOutsideCheckIns.EOF
         		
         		'Make sure you find the right class
         		If objOutsideCheckIns(0) = objStudentCount(0) Then
         			
         			'Add the data to the database
         			strSQL = "INSERT INTO CheckInHistory (Role,CheckInCount,TotalStudentCount,RecordedDate,InternalCheckIn) Values (" & vbCRLF
         			strSQL = strSQL & "'" & objOutsideCheckIns(0) & "',"
         			strSQL = strSQL & objOutsideCheckIns(1) & ","
         			strSQL = strSQL & objStudentCount(1) & ","
         			strSQL = strSQL & "#" & Date & "#,"
         			strSQL = strSQL & "False)"
         			objDBConnection.Execute(strSQL)
         			
         			intOutsideCount = intOutsideCount + objOutsideCheckIns(1)
         			
         		End If
         		
         		objOutsideCheckIns.MoveNext
         	Loop
         	objOutsideCheckIns.MoveFirst
         End If
         
         objStudentCount.MoveNext
      Loop
      
      'Add the internal total count to the database
		strSQL = "INSERT INTO CheckInHistory (Role,CheckInCount,TotalStudentCount,RecordedDate,InternalCheckIn) Values (" & vbCRLF
		strSQL = strSQL & "'Total',"
		strSQL = strSQL & intInsideCount & ","
		strSQL = strSQL & intTotalCount & ","
		strSQL = strSQL & "#" & Date & "#,"
		strSQL = strSQL & "True)"
		objDBConnection.Execute(strSQL)
		
		'Add the external total count to the database
		strSQL = "INSERT INTO CheckInHistory (Role,CheckInCount,TotalStudentCount,RecordedDate,InternalCheckIn) Values (" & vbCRLF
		strSQL = strSQL & "'Total',"
		strSQL = strSQL & intOutsideCount & ","
		strSQL = strSQL & intTotalCount & ","
		strSQL = strSQL & "#" & Date & "#,"
		strSQL = strSQL & "False)"
		objDBConnection.Execute(strSQL)
      
   End If

End Sub

Function GetLastPasswordSet(strUserName)

   'This function returns a date of when the user's password was last set

   Dim objRootDSE, objUserLookup, objUser, objPwdLastSet
   
   'Create a RootDSE object for the domain
   Set objRootDSE = GetObject("LDAP://RootDSE")
   
   'Get the user object from Active Diretory
   objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
   ">;(&(objectClass=user)(SamAccountName=" & strUserName & "));DistinguishedName"
   Set objUserLookup = objADCommand.Execute
	
	If Not objUserLookup.EOF Then
		Set objUser = GetObject("LDAP://" & objUserLookup(0))
	
		'Get the date the password expires
		Set objPwdLastSet = objUser.pwdLastSet
		GetLastPasswordSet = ConvertToDate(objPwdLastSet)
	Else
	
		'The user wasn't found in AD return the default value
		GetLastPasswordSet = #1/1/1601 12:00:00AM#
	
	End If
   
   Set objRootDSE = Nothing

End Function

Function ConvertToDate(objDate)

   'This fuction takes the pwdLastSet property from AD and returns the date

   Dim intDate, intHigh, intLow, intTimeZoneAdjust
   
   If IsDST(Date) Then
   	intTimeZoneAdjust = 240
   Else
   	intTimeZoneAdjust = 300
   End If

   'Grab the high and low parts of the 64bit number
   intHigh = objDate.HighPart
   intLow = objDate.LowPart

   'Correct for difference if the low value is 0
   If (intLow < 0) Then
      intHigh = intHigh + 1
   End If
    
   'Add the number of 100 nanosecond intervals to 1/1/1601 then convert that to a date
   intDate = #1/1/1601# + (((intHigh * (2 ^ 32)) + intLow) / 600000000 - intTimeZoneAdjust) / 1440
   ConvertToDate = CDate(intDate)

End Function

Function IsDST(datDate)

   'DST starts on the second sunday of March and ends on the first Sunday in November.  
   'This function will determine the start and end dates for the year passed to it and see
   'if the data falls within the range.  If so it will return a true, of not a false.

   Dim intIndex, datStartOfDST, datEndOfDST
   
   'Get the start date
   datStartOfDST = "3/1/" & Year(datDate)
   For intIndex = 0 to 6 
      If Weekday(DateAdd("d",intIndex,datStartOfDST)) = 1 Then
         datStartOfDST = DateAdd("d",intIndex + 7,datStartOfDST)
         Exit For
      End If
   Next
   
   'Get the end date
   datEndOfDST = "11/1/" & Year(datDate)
   For intIndex = 0 to 6 
      If Weekday(DateAdd("d",intIndex,datEndOfDST)) = 1 Then
         datEndOfDST = DateAdd("d",intIndex,datEndOfDST)
         Exit For
      End If
   Next
   
   'If the date falls in the range then it's in DST
   If CDate(datDate) >= CDate(datStartOfDST) AND CDate(datDate) < CDate(datEndOfDST) Then
      IsDST = True
   Else
      IsDST = False
   End If

End Function

Sub CleanExportFile

	'This will remove unwanted character from an input file

	Dim objFSO, strCurrentFolder, strSourceCSV, strDestinationCSV, txtSourceCSV, txtDestinationCSV
	Dim strImportedData, intIndex, strCharacter

	'Exit if we're in summer mode, we don't need to scan the import file until home rooms are set
	If bolSummerMode Then
		Exit Sub
	End If

	'Get the paths to the CSV's
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCurrentFolder = objFSO.GetAbsolutePathName(".")
	strCurrentFolder = strCurrentFolder & "\CSV\"
	strSourceCSV = strCurrentFolder & "Export.csv"
	strDestinationCSV = strCurrentFolder & "Import.csv" 

	'Open the source CSV
	Set txtSourceCSV = objFSO.OpenTextFile(strSourceCSV)
	
	'Create the destination CSV
	Set txtDestinationCSV = objFSO.CreateTextFile(strDestinationCSV)
	
	'Replace the unwanted characters 
	While txtSourceCSV.AtEndOfLine = False
		strImportedData = txtSourceCSV.ReadLine
		strImportedData = Replace(strImportedData,"´","")
		strImportedData = Replace(strImportedData,"  "," ")
		strImportedData = Replace(strImportedData,"!","")
		strImportedData = Replace(strImportedData,"""","")
		strImportedData = Replace(strImportedData,"#","")
		strImportedData = Replace(strImportedData,"$","")
		strImportedData = Replace(strImportedData,"%","")
		strImportedData = Replace(strImportedData,"&","")
		strImportedData = Replace(strImportedData,"(","")
		strImportedData = Replace(strImportedData,")","")
		strImportedData = Replace(strImportedData,"*","")
		strImportedData = Replace(strImportedData,"+","")
		strImportedData = Replace(strImportedData,"`","")
		strImportedData = Replace(strImportedData," 12:00:00 AM","")
		
		'Write the fixed data to the new file
		txtDestinationCSV.Write(strImportedData & vbCRLF)
	
	Wend
	
	'Close the CSV files
	txtSourceCSV.Close
	txtDestinationCSV.Close

	'Close objects
	Set objFSO = Nothing
	Set txtSourceCSV = Nothing
	Set txtDestinationCSV = Nothing

End Sub 

Function PasswordValid(strPassword)

   'This is a very basic validation function.  It verifies the length of the password is 8 characters or more.
   'It's more of a place holder incase you want to get more complex in the future

   If Len(strPassword) >= 8 Then
      If Len(strPassword) <= 120 Then
      	PasswordValid = True
      Else
      	PasswordValid = False
      End If
   Else
      PasswordValid = False
   End If
   
End Function

Function PasswordExpires(strUserName)

	'This will retrun true if the password is set to expire, and false if not.

	Dim objADCommand, objRootDSE, objUserLookup, objUser, strUACBinary
	
	'Create the domain objects
	Set objADCommand = ConnectToActiveDirectory
	Set objRootDSE = GetObject("LDAP://RootDSE")
		
   'Get the user's distinguished name from Active Directory	
	objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
	">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
	Set objUserLookup = objADCommand.Execute
		
	If Not objUserLookup.EOF Then
      'Build the user object
      Set objUser = GetObject("LDAP://" & objUserLookup(0))
      
      'See if the change password bit is turned on.
      strUACBinary = ConvertToBinary(objUser.userAccountControl,32)
      If Mid(strUACBinary,16,1) = "1" Then
         PasswordExpires = False
      Else
         PasswordExpires = True
      End If
   Else
      PasswordExpires = "User Not Found"
   End If
   
	'Close open objects
	Set objRootDSE = Nothing
	Set objADCommand = Nothing

End Function

Function ConvertToBinary(intNumber, intBits)

   Dim intExponent, intIndex

   'Set the number of bits you want to process
   intExponent = 2 ^ (intBits - 1)
   
   'Keep looping until you get down to the last bit.
   For intIndex = intBits To 1 Step -1
      
      'If the number is bigger then the exponet then turn on
      'the 1 and subtract the exponent from the number, otherwise
      'set the value to 0.
      If intNumber >= intExponent Then
         ConvertToBinary = ConvertToBinary & "1"
         intNumber = intNumber - intExponent
      Else
         ConvertToBinary = ConvertToBinary & "0"
      End If
      
      'Move to the next exponent
      intExponent = intExponent / 2
   
   Next

End Function

Function CreateUsername(strFirst,strLast,intClassOf)
   
   'This function will generate a student's username and make sure it doesn't already
   'exist in Active Directory before returning it.
   
   Dim strFixedLast, strInitial, strUserName, objLast, strYear, objUserLookup, intIndex
   Dim objRootDSE
   
   'Create a RootDSE object for the domain
   Set objRootDSE = GetObject("LDAP://RootDSE")
   
   'Remove unwanted characters from the first name
   strFirst = Replace(strFirst,"'","")
   strFirst = Replace(strFirst," ","")
   strFirst = Replace(strFirst,"-","")
   
   'Remove unwanted characters from the last name
   objLast = Split(strLast," ")
   strFixedLast = objLast(0)
   strFixedLast = Replace(strFixedLast,"'","")
   strFixedLast = Replace(strFixedLast," ","")
   strFixedLast = Replace(strFixedLast,"-","")
   strYear = Right(intClassOf,2)

   'Find an available username
   For intIndex = 1 to Len(strFirst)
      
		strInitial = Left(strFirst,intIndex)

		strUserName = strYear & strFixedLast & strInitial

		'Get the user object from Active Diretory
		objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
		">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
		Set objUserLookup = objADCommand.Execute
		
		If objUserLookup.EOF Then
			CreateUsername = LCase(strUserName)
			Exit For
		End If
   
   Next
   
   Set objRootDSE = Nothing
   
End Function

Function GetGrade(intGraduatingYear)
   
   'This function will return a students current grade based on their graduating
   'year using the format K-12.  Students are moved up to the next grade level
   'on July 1st.
   
   Dim datToday, intMonth, intCurrentYear, intShortGraduatingYear
   
   'Get the current month and year
   datToday = Date
   intMonth = DatePart("m",datToday)
   intCurrentYear = Right(DatePart("yyyy",datToday),2)
   intShortGraduatingYear = Right(intGraduatingYear,2)

	'If it's after July then the graduating year is the next year
   If intMonth >= 7 And intMonth <= 12 Then
      intCurrentYear = intCurrentYear + 1
   End If
   
   'Subtract the number of years remaining in school from 12 to find the current grade
   GetGrade = 12 - (intShortGraduatingYear - intCurrentYear)
   
   'Fix kindergarten
   If GetGrade = 0 Then
      GetGrade = "K"
   End If

End Function

Function GetGraduationYear(intGrade)

	Dim datToday, intMonth, intCurrentYear

   datToday = Date
   intMonth = DatePart("m",datToday)
   intCurrentYear = DatePart("yyyy",datToday)
   
   If intMonth >= 7 And intMonth <= 12 Then
      intCurrentYear = intCurrentYear + 1
   End If
   
   GetGraduationYear = intCurrentyear + (12 - intGrade)
   
End Function

Function GetOU(strRole,strSite)

	'This function will retrun an OU object
	
	strOUBase = Replace(strOUBase,"%ROLE%",strRole)
	strOUBase = Replace(strOUBase,"%SITE%",strSite)
	strOUBase = Replace(strOUBase,"Elementary", "Elementary School")
	Set GetOU = GetObject("LDAP://" & strOUBase)
		
End Function

Function GetRole(intYear)

   Dim datToday, intMonth, intCurrentYear, intGrade, strSQL, objRole
   
   'If they're an adult then get their role from the database
   If intYear <= 1000 Then
   	strSQL = "SELECT Role FROM Roles WHERE RoleID=" & intYear
   	Set objRole = objDBConnection.Execute(strSQL)
   	
   	If Not objRole.EOF Then
   		GetRole = objRole(0)
   		Exit Function
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

End Function

Function ConnectToDatabase

   'This function returns a connection object used to run SQL commands against the database

   Dim objFSO, strCurrentFolder, strDatabase, strConnection

   'Get the database path
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   strCurrentFolder = objFSO.GetAbsolutePathName(".")
   strCurrentFolder = objFSO.GetParentFolderName(strCurrentFolder)
   strCurrentFolder = strCurrentFolder & "\Database\"
   strDatabase = strCurrentFolder & "Inventory.mdb"

   'Create the connection to the database
   Set ConnectToDatabase = CreateObject("ADODB.Connection")
   strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strDatabase & ";"
   ConnectToDatabase.Open strConnection
   
   Set objFSO = Nothing

End Function

Function ConnectToActiveDirectory

   'This function returns a command object used to run commands against Active Directory

   Dim objADConnection

   'Establish a connection to Active Directory using ActiveX Data Object
   Set objADConnection = CreateObject("ADODB.Connection")
   objADConnection.Open "Provider=ADSDSOObject;"

   'Create the command object and attach it to the connection object
   Set ConnectToActiveDirectory = CreateObject("ADODB.Command")
   ConnectToActiveDirectory.ActiveConnection = objADConnection
 
   Set objADConnection = Nothing

End Function

Function SortArray(arrArray)

   Dim i, j, strTemp
   
   'i is used to keep track of what has been done so far.  There is no need to sort
   'later parts of the array with each loop because the have already been sorted.  
   For i = UBound(arrArray) - 1 To 0 Step - 1
      
      'Loop through each item in the array, each loop will push the values later
      'in the alphabet farther to the right.  We only have to go to i since with
      'each pass we will be pushing the last value all the way to the right.  there
      'is no need to sort the values all the way to the right as the loop progresses.
      For j = 0 to i
         
         'If the current value is later in the alphabet push it to the right
         If arrArray(j) > arrArray(j + 1) Then
            strTemp = arrArray(j + 1)
            arrArray(j + 1) = arrArray(j)
            arrArray(j) = strTemp
         End If
      Next
   Next
   
   'Return the array
   SortArray = arrArray

End Function

Sub UpdateLog(EntryType,DeviceTag,UserName,OldValue,NewValue,EventNumber)

	'This will update the log with what change has been made

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
	
	'Zero out the event number if nothing's there
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
	
	'Write the log entry to the database
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
	objDBConnection.Execute(strSQL)
	
End Sub