'Created by Matthew Hull on 10/9/17
'Last Updated 10/9/17

'This script will send the phone number and room number in the inventory site to Active Directory

Option Explicit

'On Error Resume Next

VerifyADandDBMatch

MsgBox "Done"

Sub VerifyADandDBMatch

	'This will sync the phone numbers stored in the database to Active Directory

	On Error Resume Next

	Dim objADCommand, objDBConnection, strSQL, objDBUsers, objADUser, objRootDSE, objUserLookup, bolError, bolADActive
	Dim strUserName, strFirstName, strLastName, strDescription, strRoom, strPhone, bolActive, strSite, intClassOf
	
	'Create a connection to the database
   Set objDBConnection = ConnectToDatabase

	'Get the list of users in the database with phone numbers
	strSQL = "SELECT UserName,FirstName,LastName,Description,RoomNumber,PhoneNumber,Active,Site,ClassOf FROM People WHERE Role='Teacher'"
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
		
			'Get the user's distinguished name from Active Directory	
			objADCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
			">;(&(objectClass=user)(samAccountName=" & strUserName & "));distinguishedName"
			Set objUserLookup = objADCommand.Execute
		
			'Build the user object
			Set objADUser = GetObject("LDAP://" & objUserLookup(0))

			'Check for differences
			If LCase(objADUser.Get("samAccountName")) <> LCase(strUserName) Then
				MsgBox "Database = " & strUserName & vbCRLF & "Active Directory = " & objADUser.Get("samAccountName"),,strUserName
				bolError = True
			End If
			If objADUser.Get("givenName") <> strFirstName Then
				If Err Then
					If strFirstName <> "" Then
						MsgBox "FirstName",,strUserName
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
						MsgBox "LastName",,strUserName
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
						MsgBox "Description",,strUserName
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
                  MsgBox "Room",,strUserName
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
                  MsgBox "Phone",,strUserName
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
                  MsgBox "IP Phone",,strUserName
						bolError = True
						Err.Clear
					End If
				Else
					bolError = True
				End If
			End If
			Select Case objADUser.Get("userAccountControl")
			
				Case 66048, 512
					bolADActive = True
					
				Case Else
					bolADActive = False
					
			End Select
			If bolADActive <> bolActive Then
				bolError = True
			End If
		
			If bolError Then 
				MsgBox strUserName
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