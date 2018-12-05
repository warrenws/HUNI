'Created by Matthew Hull 9/1/15

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

strSQL = "SELECT OldValue,NewValue FROM Log WHERE Type='StudentGradeChange'"
Set objChangedAccounts = objConnection.Execute(strSQL)

If Not objChangedAccounts.EOF Then
   Do Until objChangedAccounts.EOF
      
      strSQL = "UPDATE Log SET UserName='" & Replace(objChangedAccounts(1),"'","''") & "' WHERE UserName='" & Replace(objChangedAccounts(0),"'","''") & "'"
      objConnection.Execute(strSQL)
      
      objChangedAccounts.MoveNext
   Loop
End If

MsgBox "Done"