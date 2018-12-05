'Created by Matthew Hull 7/7/16

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

strSQL = "UPDATE Devices "
strSQL = strSQL & "SET Room = '' "
strSQLWhere = "WHERE Room='%Teacher%' AND Model="

'*******************************************************************************************************
'Kindergarden Devices
strModel = "iPad Air 1"

objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Hendry / Lavigne") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","O''Connell") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Bennett") & "'" & strModel & "'")


'*******************************************************************************************************
'1st Grade Devices
strModel = "iPad Pro 9.7"

objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Abrantes / Jaeger") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Abrantes") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Bennett") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","KellyK") & "'" & strModel & "'")

'*******************************************************************************************************
'2nd Grade Devices
strModel = "iPad Pro 9.7"

objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Buckley") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Zehr / Aspland") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Dudla") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Zehr") & "'" & strModel & "'")

'*******************************************************************************************************
'3rd Grade Devices
strModel = "iPad Air 2"

objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Allen") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Hendry") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Gershen / Goncerz") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Poetzsch") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Gershen") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Holderman / Aspland") & "'" & strModel & "'")

'*******************************************************************************************************
'4th Grade Devices
strModel = "iPad Air 1"

objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","KellyP") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Holderman / Brennan") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Lindsay") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Holderman / Aspland") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Thomsen") & "'" & strModel & "'")

'*******************************************************************************************************
'5th Grade Devices
strModel = "MacBook Air"

objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Catarelli") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Hoover") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Thomsen / Spring") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Buckley") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Gereau") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Hoover / Catarelli") & "'" & strModel & "'")

'*******************************************************************************************************
'6th Grade Devices
strModel = "MacBook Air"

objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Butler") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Crotty") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Lewis / Compositor") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Lewis") & "'" & strModel & "'")

'*******************************************************************************************************

'Left Over Devices
strModel = "MacBook Pro"

objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Butler") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Crotty") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Lewis / Compositor") & "'" & strModel & "'")
objConnection.Execute(strSQL & Replace(strSQLWhere,"%Teacher%","Lewis") & "'" & strModel & "'")


MsgBox "Done"