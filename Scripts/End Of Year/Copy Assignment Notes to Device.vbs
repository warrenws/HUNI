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

'Get all the notes for the devices that were returned today.
strSQL = "SELECT LGTag, Notes" & vbCRLF
strSQL = strSQL & "FROM Assignments" & vbCRLF
strSQL = strSQL & "WHERE Notes<>'' AND DateReturned=#" & Date() & "#"
Set objNotes = objConnection.Execute(strSQL)

'Loop through each returned device
If Not objNotes.EOF Then
   Do Until objNotes.EOF
      
      strSQL = "SELECT Notes FROM Devices WHERE LGTag='" & objNotes(0) & "'"
      Set objCurrentNotes = objConnection.Execute(strSQL)
      
      If objCurrentNotes(0) <> "" Then
         strNotes = objCurrentNotes(0) & vbCRLF & vbCRLF & objNotes(1)
      Else
         strNotes = objNotes(1)
      End If
     
      'Set the assigned field to true on the device
      strSQL = "UPDATE Devices SET Notes='" & Replace(strNotes,"'","''") & "' WHERE LGTag='" & objNotes(0) & "'"
      objConnection.Execute(strSQL)
      
      objNotes.MoveNext
   Loop
End If
  
MsgBox "Done"