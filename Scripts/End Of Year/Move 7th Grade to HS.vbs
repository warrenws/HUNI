'Created by Matthew Hull 7/8/16

'On Error Resume Next

'Get the inventory database path
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentFolder = objFSO.GetAbsolutePathName(".")
strCurrentFolder = strCurrentFolder & "\..\..\Database"
strInventoryDatabase = strCurrentFolder & "\Inventory.mdb"

'Create the connection to the inventory database
Set objConnection = CreateObject("ADODB.Connection")
strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strInventoryDatabase & ";"
objConnection.Open strConnection

strCurrentYear = Year(Date)

'Move the computers to the HS
strSQL = "UPDATE People INNER JOIN (Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag) ON People.ID = Assignments.AssignedTo "
strSQL = strSQL & "SET Devices.Site = 'High School' "
strSQL = strSQL & "WHERE (((People.ClassOf)=" & strCurrentYear + 6 & ") AND ((Assignments.DateReturned)>=#6/15/" & strCurrentYear & "# And (Assignments.DateReturned)<=#" & Date & "#));"
objConnection.Execute(strSQL)

'Move the students to the HS
strSQL = "UPDATE People Set Site='High School' WHERE Active=True AND ClassOf=" & strCurrentYear + 6
objConnection.Execute(strSQL)


MsgBox "Done"
