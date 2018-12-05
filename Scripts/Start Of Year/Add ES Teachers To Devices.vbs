'Created by Matthew Hull 8/7/15

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

strSQL = "UPDATE People INNER JOIN (Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag) ON People.ID = Assignments.AssignedTo "
strSQL = strSQL & "SET Devices.Room="
strSQLWHERE = " WHERE Assignments.Active=True And People.HomeRoom='%Teacher%'"

'*******************************************************************************************************
'1st Grade Devices

objConnection.Execute(strSQL & "'Abrantes'" & Replace(strSQLWhere,"%Teacher%","Abrantes, Sarah"))
objConnection.Execute(strSQL & "'KellyK'" & Replace(strSQLWhere,"%Teacher%","Kelly, Krista"))
objConnection.Execute(strSQL & "'Bennett'" & Replace(strSQLWhere,"%Teacher%","Bennett, Kimberly"))

'*******************************************************************************************************
'2nd Grade Devices

objConnection.Execute(strSQL &"'Zehr'"  & Replace(strSQLWhere,"%Teacher%","Zehr, Anna"))
objConnection.Execute(strSQL &"'Dudla'" & Replace(strSQLWhere,"%Teacher%","Dudla, Kellie"))

'*******************************************************************************************************
'3rd Grade Devices

objConnection.Execute(strSQL & "'Allen'" & Replace(strSQLWhere,"%Teacher%","Allen, Jeffrey"))
objConnection.Execute(strSQL & "'Poetzsch'"& Replace(strSQLWhere,"%Teacher%","Poetzsch, Alexandra"))
objConnection.Execute(strSQL & "'Gershen'" & Replace(strSQLWhere,"%Teacher%","Gershen, Ashley" ))

'*******************************************************************************************************
'4th Grade Devices

objConnection.Execute(strSQL & "'Thomsen'" & Replace(strSQLWhere,"%Teacher%","Thomsen, Brian"))
objConnection.Execute(strSQL & "'Holderman / Aspland'" & Replace(strSQLWhere,"%Teacher%","Holderman / Aspland"))
objConnection.Execute(strSQL & "'Lindsay'" & Replace(strSQLWhere,"%Teacher%","Lindsay, Lisa"))

'*******************************************************************************************************
'5th Grade Devices

objConnection.Execute(strSQL & "'Gereau'" & Replace(strSQLWhere,"%Teacher%","Gereau, Talia"))
objConnection.Execute(strSQL & "'Hoover / Catarelli'" & Replace(strSQLWhere,"%Teacher%","Hoover / Catarelli"))
objConnection.Execute(strSQL & "'Buckley'" & Replace(strSQLWhere,"%Teacher%","Buckley, Kelly"))

'*******************************************************************************************************
'6th Grade Devices

objConnection.Execute(strSQL & "'Butler'" & Replace(strSQLWhere,"%Teacher%","Butler, Matthew"))
objConnection.Execute(strSQL & "'Crotty'" & Replace(strSQLWhere,"%Teacher%","Crotty, Jeffrey"))
objConnection.Execute(strSQL & "'Lewis'" & Replace(strSQLWhere,"%Teacher%","Lewis, Jonathan"))

'*******************************************************************************************************

MsgBox "Done"