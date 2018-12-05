'Created by Matthew Hull 9/15/17

'On Error Resume Next

'Get the inventory database path
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentFolder = objFSO.GetAbsolutePathName(".")
strInsuranceFile = strCurrentFolder & "\Insurance.csv"
strCurrentFolder = strCurrentFolder & "\..\..\Database"
strInventoryDatabase = strCurrentFolder & "\Inventory.mdb"

'Create the connection to the inventory database
Set objConnection = CreateObject("ADODB.Connection")
strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strInventoryDatabase & ";"
objConnection.Open strConnection

'Open the file
Set txtInsuranceFile = objFSO.OpenTextFile(strInsuranceFile)

'Loop through each line in the file
While txtInsuranceFile.AtEndOfLine = False

	'Read in the next line of the text file
	strSerial = txtInsuranceFile.ReadLine
	
   'Update the student ID in the database
   strSQL = "UPDATE Devices SET "
   strSQL = strSQL & "HasInsurance=True "
   strSQL = strSQL & "WHERE SerialNumber='" & Replace(strSerial,"'","''") & "'"
   objConnection.Execute(strSQL)

WEnd

MsgBox "Done"

'Close the text file
txtInsuranceFile.Close

'Close the objects
Set objFSO = Nothing
Set txtInsuranceFile = Nothing