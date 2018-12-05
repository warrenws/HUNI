<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 3/22/17
'Last Updated 1/14/18

'

'Option Explicit

'On Error Resume Next

CONST LGTAG = 0
CONST EXTERNALIP = 1
CONST INTERNALIP = 2
CONST COMPUTERNAME = 3
CONST LASTUSER = 4
CONST OSVERSION = 5
CONST INTERNAL_IP_ROOT = "10.15."
CONST EXTERNAL_IP_ROOT = "163.153.220."

Dim strSerialNumber, strIPAddress, strComputerName, strLastUser

strSerialNumber = Request.QueryString("Serial")
strInternalIP = Request.QueryString("InternalIP")
strExternalIP = Request.QueryString("ExternalIP")
strComputerName = Request.QueryString("ComputerName")
strLastUser = Request.QueryString("LastUser")
strOSVersion = Request.QueryString("OS")
strTask = Request.QueryString("Task")
datDate = Date()
datTime = Time()

If InStr(strInternalIP,",") Then
	arrIPAddresses = Split(strInternalIP,",")
	strInternalIP = arrIPAddresses(0)
End If 

strSQL = "SELECT LGTag,ExternalIP,InternalIP,ComputerName,LastUser,OSVersion FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
Set objDevice = Application("Connection").Execute(strSQL)

Select Case strTask

	Case "Imaged"
		UpdateLog "ComputerImaged",objDevice(0),"","","",""
	Case "AdminCheckIn"
		AdminCheckIn
	Case Else
		UpdateInventory

End Select

Sub AdminCheckIn
		
		strSQL = "SELECT ID FROM Tags WHERE Tag='Screenshot' AND LGTag='" & objDevice(0) & "'"
		Set objScreenshot = Application("Connection").Execute(strSQL)
		
		If Not objScreenshot.EOF Then
			Response.Write "Screenshot"
		End If
		
End Sub

Sub UpdateInventory

	strSQL = ""
	If Not objDevice.EOF Then

		If IsNull(objDevice(LASTUSER)) Or objDevice(LASTUSER) <> strLastUser Then
		
			strSQL = "SELECT ID FROM People WHERE UserName='" & strLastUser & "'"
			Set objUser = Application("Connection").Execute(strSQL)
		
			If Not objUser.EOF Then
				strSQLLastUser = "LastUser='" & strLastUser & "',"
				If IsNull(objDevice(LASTUSER)) Then
					UpdateLog "LastUserChange",objDevice(LGTAG),strLastUser,"",strLastUser,""
				Else
					UpdateLog "LastUserChange",objDevice(LGTAG),strLastUser,objDevice(LASTUSER),strLastUser,""
				End If
			Else
				strLastUser = ""
			End If
		End If

		If IsNull(objDevice(EXTERNALIP)) Or objDevice(EXTERNALIP) <> strExternalIP Then
			If Left(strExternalIP,Len(EXTERNAL_IP_ROOT)) <> EXTERNAL_IP_ROOT Then
				strSQLExternalIP = "ExternalIP='" & strExternalIP & "',"
				If IsNull(objDevice(EXTERNALIP)) Then
					UpdateLog "ExternalIPChange",objDevice(LGTAG),strLastUser,"",strExternalIP,""
				Else
					UpdateLog "ExternalIPChange",objDevice(LGTAG),strLastUser,objDevice(EXTERNALIP),strExternalIP,""
				End If
			End If
		End If

		If ISNull(objDevice(INTERNALIP)) Or objDevice(INTERNALIP) <> strInternalIP Then
			If Left(strInternalIP,Len(INTERNAL_IP_ROOT)) = INTERNAL_IP_ROOT Then
				strSQLInternalIP = "InternalIP='" & strInternalIP & "',"
				If IsNull(objDevice(INTERNALIP)) Then
					UpdateLog "InternalIPChange",objDevice(LGTAG),strLastUser,"",strInternalIP,""
				Else
					UpdateLog "InternalIPChange",objDevice(LGTAG),strLastUser,objDevice(INTERNALIP),strInternalIP,""
				End If
			End If
		End If
	
		If IsNull(objDevice(COMPUTERNAME)) Or objDevice(COMPUTERNAME) <> strComputerName Then
			strSQLComputerName = "ComputerName='" & strComputerName & "',"
			If IsNull(objDevice(COMPUTERNAME)) Then
				UpdateLog "ComputerNameChange",objDevice(LGTAG),strLastUser,"",strComputerName,""
			Else
				UpdateLog "ComputerNameChange",objDevice(LGTAG),strLastUser,objDevice(COMPUTERNAME),strComputerName,""
			End If
		End If
	
		If IsNull(objDevice(OSVERSION)) Or objDevice(OSVERSION) <> strOSVersion Then
			strSQLOSVersion = "OSVersion='" & strOSVersion & "',"
			If IsNull(objDevice(OSVERSION)) Then
				UpdateLog "OSVersionChange",objDevice(LGTAG),strLastUser,"",strOSVersion,""
			Else
				UpdateLog "OSVersionChange",objDevice(LGTAG),strLastUser,objDevice(OSVERSION),strOSVersion,""
			End If
		End If
	
		If strSQLExternalIP <> "" Or strSQLInternalIP <> "" Or strSQLComputerName <> "" Or strSQLLastUser <> "" Or strSQLOSVersion <> "" Then
			strSQL = "UPDATE Devices SET "
			strSQL = strSQL & strSQLExternalIP & strSQLInternalIP & strSQLComputerName & strSQLLastUser & strSQLOSVersion
			strSQL = Left(strSQL,Len(strSQL) - 1) 'Drop Last Comma
			strSQL = strSQL & " WHERE SerialNumber='" & strSerialNumber & "'"
			Application("Connection").Execute(strSQL)
		End If
		
		If strLastUser <> "" Then
			If strExternalIP <> "" Then
				If Left(strExternalIP,Len(EXTERNAL_IP_ROOT)) <> EXTERNAL_IP_ROOT Then
					strSQL = "UPDATE People SET LastExternalCheckIn=#" & datDate & "# WHERE UserName='" & strLastUser & "'"
					Application("Connection").Execute(strSQL)
				End If
			End If
		End If
		
		If strLastUser <> "" Then
			If strInternalIP <> "" Then
				If Left(strInternalIP,Len(INTERNAL_IP_ROOT)) = INTERNAL_IP_ROOT Then
					strSQL = "UPDATE People SET LastInternalCheckIn=#" & datDate & "# WHERE UserName='" & strLastUser & "'"
					Application("Connection").Execute(strSQL)
				End If
			End If
		End If
	
		strSQL = "UPDATE Devices SET LastCheckInDate=#" & datDate & "#,LastCheckInTime=#" & datTime & "# WHERE SerialNumber='" & strSerialNumber & "'"
		Application("Connection").Execute(strSQL)

		strSQL = "SELECT ID FROM Events WHERE Resolved=False AND Type='Lost Device' AND LGTag='" & objDevice(0) & "'"
		Set objLostDeviceCheck = Application("Connection").Execute(strSQL)
	
		bolProblemFound = False
	
		If Not objLostDeviceCheck.EOF Then
			Response.Write "Lost"
			bolProblemFound = True
			SendLostEMail
		End If
	
		strSQL = "SELECT ID FROM Tags WHERE (Tag='Missing' OR Tag='Lost') AND LGTag='" & objDevice(0) & "'"
		Set objLostDeviceCheck = Application("Connection").Execute(strSQL)
	
		If Not objLostDeviceCheck.EOF Then
			Response.Write "Lost"
			bolProblemFound = True
			SendLostEMail
		End If
	
		strSQL = "SELECT ID FROM Tags WHERE Tag='Disable' AND LGTag='" & objDevice(0) & "'"
		Set objLostDeviceCheck = Application("Connection").Execute(strSQL)
	
		If Not objLostDeviceCheck.EOF Then
			Response.Write "Disable"
			bolProblemFound = True
			SendLostEMail
		End If
		
		strSQL = "SELECT ID FROM Tags WHERE Tag='Screenshot' AND LGTag='" & objDevice(0) & "'"
		Set objScreenshot = Application("Connection").Execute(strSQL)
		
		If Not objScreenshot.EOF Then
			Response.Write "Screenshot"
			bolProblemFound = True
		End If
		
		strSQL = "SELECT ID FROM Tags WHERE Tag='Notify' AND LGTag='" & objDevice(0) & "'"
		Set objNotify = Application("Connection").Execute(strSQL)
		
		If Not objNotify.EOF Then
			SendNotificationEMail
		End If
		
		
		'Written by Dane.
		strSQL = "SELECT ID FROM Tags WHERE Tag='See Me' and LGTag='" & objDevice(0) & "'"
		Set objSeeMe = Application("Connection").Execute(strSQL)
		
		'Written by Dane.
		If Not objSeeMe.EOF Then
			Response.Write "SeeMe"
			bolProblemFound = True
		End If
				
		If Not bolProblemFound Then
			Response.Write "Good"
		End If

	End If 
	
End Sub

Sub SendLostEMail

	'This will send out an email

	Dim strSMTPPickupFolder, objMessage, objConf, strMessage, strSubject, bolHTMLMEssage
	Dim bolSendAsAdmin, strURL

   Const cdoSendUsingPickup = 1

   strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"
   
   strURL = "http://" & Request.ServerVariables("server_name")
   strURL = strURL & Left(Request.ServerVariables("path_info"),Len(Request.ServerVariables("path_info")) - 16)
   strURL = strURL & "device.asp?Tag=" & objDevice(0)
   
   'Get the message body
   strMessage =  "The device with the tag of " & objDevice(0) & " has checked into the inventory system.  "
   If strLastUser <> "" Then
   	strMessage = strMessage & "The last user was " &  strLastUser & ".  "
   End If

   strMessage = strMessage & vbCRLF & vbCRLF & strURL
   
   strSubject = "Missing Device Checked In - Asset Tag " & objDevice(0)

   'Create the objects required to send the mail.
   Set objMessage = CreateObject("CDO.Message")
   Set objConf = objMessage.Configuration
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
      .Update
   End With
   
   objMessage.TextBody = strMessage
   objMessage.From = Application("EMailNotifications")
   objMessage.To = Application("LostDeviceNotify")
   objMessage.Subject = strSubject
   objMessage.Send
   
   'Close objects
   Set objMessage = Nothing
   Set objConf = Nothing
   
End Sub

Sub SendNotificationEMail

	'This will send out an email

	Dim strSMTPPickupFolder, objMessage, objConf, strMessage, strSubject, bolHTMLMEssage
	Dim bolSendAsAdmin, strURL

   Const cdoSendUsingPickup = 1

   strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"
   
   strURL = "http://" & Request.ServerVariables("server_name")
   strURL = strURL & Left(Request.ServerVariables("path_info"),Len(Request.ServerVariables("path_info")) - 16)
   strURL = strURL & "device.asp?Tag=" & objDevice(0)
   
   'Get the message body
   strMessage =  "The device with the tag of " & objDevice(0) & " has checked into the inventory system.  "
   If strLastUser <> "" Then
   	strMessage = strMessage & "The last user was " &  strLastUser & ".  "
   End If

   strMessage = strMessage & vbCRLF & vbCRLF & strURL
   
   strSubject = "Inventory Check In - Asset Tag " & objDevice(0)

   'Create the objects required to send the mail.
   Set objMessage = CreateObject("CDO.Message")
   Set objConf = objMessage.Configuration
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
      .Update
   End With
   
   objMessage.TextBody = strMessage
   objMessage.From = Application("EMailNotifications")
   objMessage.To = Application("LostDeviceNotify")
   objMessage.Subject = strSubject
   objMessage.Send
   
   'Close objects
   Set objMessage = Nothing
   Set objConf = Nothing
   
End Sub

Sub UpdateLog(EntryType,DeviceTag,UserName,OldValue,NewValue,EventNumber)

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
	
End Sub %>