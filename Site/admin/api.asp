<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/11/18
'Last Updated 6/11/18

'This is the API for the inventory site

Option Explicit

'On Error Resume Next

Dim strLookupType, strUserName, strSQL, objLookUp

strLookupType = Request.QueryString("Type")

Select Case strLookupType

	Case "StudentID"
	
		strUserName = Request.QueryString("UserName")
	
		strSQL = "SELECT StudentID FROM PEOPLE WHERE UserName='" & strUserName & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)
		
		If Not objLookup.EOF Then
			Response.Write(objLookUp(0))
		End If

End Select
%>