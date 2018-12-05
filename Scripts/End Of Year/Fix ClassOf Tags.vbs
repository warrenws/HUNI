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

'Clear the existing tags
For intIndex = 2000 to 2100
   strSQL = "SELECT ID,LGTag FROM Tags WHERE Tag='" & intIndex & "'"
   Set objDeviceList = objConnection.Execute(strSQL)
   
   If Not objDeviceList.EOF Then
   
      Do Until objDeviceList.EOF
         
         strSQL = "SELECT ID FROM Tags WHERE Tag='Spare' AND LGTag='" & objDeviceList(1) & "'"
         Set objDeviceCheck = objConnection.Execute(strSQL)
         
         If objDeviceCheck.EOF Then
            strSQL = "DELETE FROM Tags WHERE ID=" & objDeviceList(0)
            objConnection.Execute(strSQL)
         End If
         
         objDeviceList.MoveNext
      Loop
   
   End If
   
Next

'Get all the active assignments from the database
strSQL = "SELECT Assignments.LGTag, ClassOf" & vbCRLF
strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
strSQL = strSQL & "WHERE ClassOf>1000 AND HasDevice=True AND Assignments.Active=True"
Set objAssignments = objConnection.Execute(strSQL)

'Loop through each active assignment
If Not objAssignments.EOF Then
   Do Until objAssignments.EOF
      
      strSQL = "SELECT ID FROM Tags WHERE Tag='Spare' AND LGTag='" & objAssignments(0) & "'"
      Set objDeviceCheck = objConnection.Execute(strSQL)
      
      If objDeviceCheck.EOF Then
    
         'Update the tag
         strSQL = "INSERT INTO Tags (LGTag,Tag) VALUES ("
         strSQL = strSQL &  "'" & objAssignments(0) & "','" & objAssignments(1) & "')"
         objConnection.Execute(strSQL)
      End If
      
      objAssignments.MoveNext
   Loop
End If

MsgBox "Done"