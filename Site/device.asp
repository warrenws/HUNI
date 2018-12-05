<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/16/14
'Last Updated 1/14/18

'This page shows the details for a single device in the inventory website

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim objDevice, strActiveChecked, objAssignment, objOldAssignments, objSites, objEventTypes, objEvents
Dim intTag, strSubmitTo, strMessage, strAssignedTo, objClasses, intClassOf, intStudent
Dim bolAdapterReturned, bolCaseReturned, strInsuredChecked, strCardType

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

   Dim strSQL

   'Get the variables from the URL
   If Application("UseLeadingZeros") Then
		intTag = Request.QueryString("Tag")
	Else
		If IsNumeric(Request.QueryString("Tag")) Then
			intTag = Int(Request.QueryString("Tag"))
		Else
			intTag = Request.QueryString("Tag")
		End If
	End If
   
   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Add Event"
         AddSubmittedEvent
      Case "Assign"
      	AssignDevice
      Case "Return"
         ReturnDevice
      Case "Update Event"
      	UpdateEvent
      Case "Update Device"
      	UpdateDevice
   End Select
   
	'Setup the assign a device form
   intClassOf = Request.Form("ClassOf")
   intStudent = Request.Form("StudentID")
   If intClassOf = "" Then
      intClassOf = 0
   End If
   If intStudent = "" Then
      intStudent = 0
   End If      

   'Get the information about the device
   strSQL = "SELECT Manufacturer,Site,Model,Room,SerialNumber,Cart,BOCESTag,HasInsurance,DatePurchased,Active,AppleID,MACAddress,Notes" & vbCRLF
   strSQL = strSQL & "FROM Devices" & vbCRLF
   strSQL = strSQL & "WHERE LGTag='" & intTag & "'"
   Set objDevice = Application("Connection").Execute(strSQL)
   
   'If the device isn't found send them back to the index page.
   If objDevice.EOF Then
      Response.Redirect("index.asp?Error=DeviceNotFound")
   End If
   
   'Set the status of the insured checkbox
   If objDevice(7) Then
      strInsuredChecked = "checked=""checked"""
   Else
      strInsuredChecked = ""
   End If
   
   'Set the status of the active checkbox
   If objDevice(9) Then
      strActiveChecked = "checked=""checked"""
   Else
      strActiveChecked = ""
   End If
  
  	'Get the current assignment 
  	strSQL = "SELECT FirstName,LastName,ClassOf,Assignments.Notes,HasInsurance,StudentID,UserName,Role,HomeRoom,People.Active,Warning,Loaned,People.ID" & _
   	",Username,PWord,AUP,DateIssued,IssuedBy" & vbCRLF
   strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
   strSQL = strSQL & "WHERE LGTag='" & intTag & "' AND Assignments.Active=True"
   Set objAssignment = Application("Connection").Execute(strSQL)
   
   'Get the old assignments
   strSQL = "SELECT FirstName,LastName,ClassOf,DateIssued,DateReturned,Assignments.Notes,StudentID,UserName,Role,HomeRoom" & vbCRLF
   strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
   strSQL = strSQL & "WHERE LGTag='" & intTag & "' AND Assignments.Active=False"
   Set objOldAssignments = Application("Connection").Execute(strSQL)
   
   'Get the list of classes for the assign a device drop down menu
   strSQL = "SELECT DISTINCT ClassOf FROM People ORDER BY ClassOf DESC"
   Set objClasses = Application("Connection").Execute(strSQL)
   
   'Get the list of events for this device
   strSQL = "SELECT ID,Type,Notes,EventDate,EventTime,Resolved,ResolvedDate,ResolvedTime FROM Events WHERE LGTag='" & intTag & "'"
   Set objEvents = Application("Connection").Execute(strSQL)
   
   'Get the list of sites for the site drop down menu
   strSQL = "SELECT Site FROM Sites WHERE Active=True ORDER BY Site"
   Set objSites = Application("Connection").Execute(strSQL)
   
   'Get the list of event types for the event types drop down menu
   strSQL = "SELECT EventType FROM EventTypes WHERE Active=True ORDER BY EventType"
   Set objEventTypes = Application("Connection").Execute(strSQL)
   
   'Get the URL used to submit forms
   If Request.ServerVariables("QUERY_STRING") = "" Then
      strSubmitTo = "device.asp"
   Else   
      strSubmitTo = "device.asp?" & Request.ServerVariables("QUERY_STRING")
   End If
   
   'Set up the variables needed for the site then load it
   SetupSite
   DisplaySite
   
End Sub%>

<%Sub DisplaySite 
	
	Dim intCounter, datToday, strSQL, objNames %> 

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title><%=Application("SiteName")%></title>
      <link rel="stylesheet" type="text/css" href="style.css" /> 
      <link rel="apple-touch-icon" href="images/inventory.png" /> 
      <link rel="shortcut icon" href="images/inventory.ico" />
      <meta name="viewport" content="width=device-width" />
      <meta name="theme-color" content="#333333">
      <link rel="stylesheet" href="assets/css/jquery-ui.css">
		<script src="assets/js/jquery.js"></script>
		<script src="assets/js/jquery-ui.js"></script>
		<script>
   		$(function() {
   		<% If Not IsMobile Then %>
					$( document ).tooltip({track: true});
			<% End If %>
			})
   	</script>
   </head>
   
   <body class="<%=strSiteVersion%>">
   
      <div class="Header"><%=Application("SiteName")%></div>
      <div>
         <ul class="NavBar" align="center">
            <li><a href="index.asp"><img src="images/home.png" title="Home" height="32" width="32"/></a></li>
            <li><a href="log.asp"><img src="images/log.png" title="System Log" height="32" width="32"/></a></li>
            <li><a href="login.asp?action=logout"><img src="images/logout.png" title="Log Out" height="32" width="32"/></a></li>
         </ul>
      </div>   

    <% 
    DeviceInformation
    'AddEvent     
    ActiveAssignments   
    'OldAssignments
    'Events
    'SearchForDevice
    %>
    <%=strMessage%>
		<div class="Version">Version <%=Application("Version")%></div>
   </body>
   </html>

<%End Sub%>

<%Sub DeviceInformation

	Dim objFSO

   Set objFSO = CreateObject("Scripting.FileSystemObject")%>

<% If objDevice(9) Then
		strCardType = "NormalCard"
	Else
		strCardType = "DisabledCard"
	End If %>
	<div class="Card <%=strCardType%>">
	<div class="CardTitle">Asset Tag <%=intTag%></div>
	<div Class="ImageSectionInCard">
	<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDevice(2)," ","") & ".png") Then %>      
			<img class="PhotoCard" src="images/devices/<%=Replace(objDevice(2)," ","")%>.png" width="96" />
	<% Else %>
			<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
	<% End If %>
	</div>
	<div Class="RightOfImageInCard">
      <div>
         <div Class="PhotoCardColumn1">Make: </div>
         <div Class="PhotoCardColumn2"><%=objDevice(0)%></div>
      </div>
      <div>
         <div Class="PhotoCardColumn1">Model: </div>
         <div Class="PhotoCardColumn2"><%=objDevice(2)%></div>
      </div>
      <div>
         <div Class="PhotoCardColumn1">Serial: </div>
         <div Class="PhotoCardColumn2"><%=objDevice(4)%></div>
      </div>
      <% If objDevice(6) <> "" Then %>
			<div>
				<div Class="CardMerged">BOCES Tag: <%=objDevice(6)%></div>
			</div>
		<% End If %>
		<div>
         <div Class="CardMerged">Site: <%=objDevice(1)%></div>
      </div>
   <% If objDevice(3) <> "" Then %>
			<div>
				<div Class="CardMerged">Room: <%=objDevice(3)%></div>
			</div>
	<% End If %>
		</div>
   </div> 
      
<%End Sub%>

<%Sub ActiveAssignments%>

<% If Not objAssignment.EOF Then 

      Dim objFSO

      Set objFSO = CreateObject("Scripting.FileSystemObject")
      
      Do Until objAssignment.EOF %>
		
			<div class="Card NormalCard">   
				<div class="CardTitle">
				<% If objAssignment(7) = "Student" Then %>
					<% If Application("ShowPasswords") Then %>
						<% If objAssignment(15) Then %>
								<image src="images/yes.png" width="15" height="15" title="AUP Signed" />
						<% Else %>
								<image src="images/no.png" width="15" height="15" title="AUP Not Signed" />
						<% End If %>
					<% End If %>
				<% End If %>
					Active Assignment
				</div>
				<div Class="ImageSectionInCard">
			<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objAssignment(7) & "s\" & objAssignment(5) & ".jpg") Then %>   
					<a href="user.asp?UserName=<%=objAssignment(6)%>">   
						<img class="PhotoCard" src="/photos/<%=objAssignment(7)%>s/<%=objAssignment(5)%>.jpg" width="96" />
					</a>
			<% Else %>
					<a href="user.asp?UserName=<%=objAssignment(6)%>">
						<img class="PhotoCard" src="/photos/<%=objAssignment(7)%>s/missing.png" width="96" />
					</a>
			<% End If %>
				</div>
				<div Class="RightOfImageInCard">
					<div Class="PhotoCardColumn1">Name: </div>
					<div Class="PhotoCardColumn2">
						<a href="user.asp?UserName=<%=objAssignment(6)%>"><%=objAssignment(0) & " " & objAssignment(1)%></a>
					</div>
				<% If objAssignment(7) = "Student" Then %>	
						<div>
							<div Class="PhotoCardColumn1">Grade: </div>
							<div Class="PhotoCardColumn2Long">
								<a href="users.asp?Role=<%=objAssignment(2)%>"><%=GetRole(objAssignment(2))%></a>
							</div>
						</div>
				<% End If %>
				<% If objAssignment(7) = "Student" Then 
						If Application("ShowPasswords") Then %>
							<div>
								<div Class="CardMerged">Username: <%=objAssignment(13)%></div>
							</div>
							<div>
								<div Class="CardMerged">Password: <%=objAssignment(14)%></div>
							</div>
					<% End If %>
						<div Class="CardMerged">
							<div>Guide:
								<a href="users.asp?GuideRoom=<%=objAssignment(8)%>"><%=objAssignment(8)%></a>
							</div>
						</div>
				<% End If %>
				</div>
			</div>
   
      <% objAssignment.MoveNext
      Loop 
   
   End If %>

<%End Sub%>

<%Sub OldAssignments%>

   <% If Not objOldAssignments.EOF Then 

      Dim objFSO

      Set objFSO = CreateObject("Scripting.FileSystemObject")   
   
      Do Until objOldAssignments.EOF %>

			<div class="Card OldAssignmentCard">
				<div>
			<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objOldAssignments(8) & "s\" & objOldAssignments(6) & ".jpg") Then %>   
					<a href="user.asp?UserName=<%=objOldAssignments(7)%>">  
						<img class="PhotoCard" src="/photos/<%=objOldAssignments(8)%>s/<%=objOldAssignments(6)%>.jpg" width="96" />
					</a>
			<% Else %>
					<a href="user.asp?UserName=<%=objOldAssignments(7)%>">
						<img class="PhotoCard" src="/photos/<%=objOldAssignments(8)%>s/missing.png" width="96" />
					</a>
			<% End If %>
				</div>
				<div class="CardTitle">Old Assignment</div>
				<div>
					<div Class="PhotoCardColumn1">Name: </div>
					<div Class="PhotoCardColumn2">
						<a href="user.asp?UserName=<%=objOldAssignments(7)%>"><%=objOldAssignments(0) & " " & objOldAssignments(1)%></a>
					</div>
				</div>
				<div>
					<div Class="PhotoCardColumn1">Date: </div>
					<div Class="PhotoCardColumn2"><%=ShortenDate(objOldAssignments(3)) & " - " & ShortenDate(objOldAssignments(4))%></div>
				</div>
		<% If objOldAssignments(8) = "Student" Then %>		
				<div>
					<div Class="Bold">Guide Room: </div>
					<div>
						<a href="users.asp?GuideRoom=<%=objOldAssignments(9)%>"><%=objOldAssignments(9)%></a>
					</div>
				</div>
		<% End If %>
		<% If objOldAssignments(5) <> "" Then %> 
			<% If objOldAssignments(8) = "Student" Then %>	
					<div>&nbsp;</div>
			<% Else %>
					<div>&nbsp;</div>
					<div>&nbsp;</div>
			<% End If %>
				<div Class="Bold">Notes: </div>
				<div><%=objOldAssignments(5)%></div>
		<% End If %>
			</div>
         
      <% objOldAssignments.MoveNext
      Loop %>     
   <% End If %>
   
<%End Sub%>

<%Sub Events%>

<% If Not objEvents.EOF Then %>

   <% Do Until objEvents.EOF 
   
   		If objEvents(5) Then %> 

				<div class="Card NormalCard">
					<div class="CardTitle">Resolved Event</div>
					<div>
						<div Class="CardColumn1">Event Type: </div>
						<div Class="CardColumn2"><%=objEvents(1)%></div>
					</div>
					<div>
						<div Class="CardColumn1">Date: </div>
				<% If ShortenDate(objEvents(3)) = ShortenDate(objEvents(6)) Then %>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3))%></div>
				<% Else %>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3)) & " - " & ShortenDate(objEvents(6))%></div>
				<% End If %>
					</div>
				<% If objEvents(1) <> "" Then %> 
						<div Class="Bold">Notes: </div>
						<div><%=objEvents(2)%></div>
				<% End If %>
				</div>
				
		<% Else %>
		
				<div class="Card NormalCard">
					<form method="POST" action="<%=strSubmitTo%>">
					<input type="hidden" name="EventID" value="<%=objEvents(0)%>" />
					<div class="CardTitle">Open Event</div>
					<div>
						<div Class="CardColumn1">Event Type: </div>
						<div Class="CardColumn2"><%=objEvents(1)%></div>
					</div>
					<div>
						<div Class="CardColumn1">Submitted: </div>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3)) & " " & ShortenTime(objEvents(4))%></div>
					</div>
					<div>
						<div Class="CardColumn1">Complete: </div>
						<div Class="CardColumn2">
							<input Class="TwoColumnCard" type="checkbox" value="TRUE" name="Resolved" />
						</div>
					</div>
					<div Class="Bold">Notes: </div>
					<div>
						<textarea Class="TwoColumnCard" rows="5" name="Notes" cols="90" style="width: 99%;"><%=objEvents(2)%></textarea>
					</div>
					<div>&nbsp;</div>
					<div Class="Button"><input type="submit" value="Update Event" name="Submit" /></div>
					</form>
				</div>

         
      <% End If
      	objEvents.MoveNext
      Loop 
   End If %>

<%End Sub%>

<%Sub SearchForDevice%>
	<div class="Card NormalCard">
		<form method="POST" action="index.asp">
		<div class="CardTitle">Search for a Device</div>
		<div>
			<div Class="CardColumn1">Asset tag: </div>
			<div Class="CardColumn2">
				<input class="TwoColumnCard InputWidthSmall" type="text" name="Tag" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">BOCES tag: </div>
			<div Class="CardColumn2">
				<input class="TwoColumnCard InputWidthSmall" type="text" name="BOCESTag" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Serial #: </div>
			<div Class="CardColumn2">
				<input class="TwoColumnCard InputWidthLarge" type="text" name="Serial" />
			</div>
		</div>
		<div>
			<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		</div>
		</form>
	</div>
<%End Sub%>

<%Sub AddEvent%>

   <div Class="Card">
      <form method="POST" action="<%=strSubmitTo%>">
      <div Class="CardTitle">Add Event</div>
      <div Class="Bold">Event Type: 
         <select Class="SingleColumnCard" name="EventType">
            <option value=""></option>
      <% Do Until objEventTypes.EOF %>
            <option value="<%=objEventTypes(0)%>"><%=objEventTypes(0)%></option>
      <%    objEventTypes.MoveNext
         Loop %>
         </select>
      </div>
      <div>&nbsp;</div>
      <div Class="Bold">Notes:</div>
      <div><textarea Class="TwoColumnCard" rows="5" name="Notes" cols="90" style="width: 99%;"></textarea></div>
      <div>&nbsp;</div>
      <div Class="Button"><input type="submit" value="Add Event" name="Submit" /></div>
      </form>
   </div>  

<%End Sub%>


<%Sub AddSubmittedEvent 

   Dim strEventType, strNotes, strSQL
   
   strEventType = Replace(Request.Form("EventType"),"'","''")
   strNotes = Replace(Request.Form("Notes"),"'","''")
   
   strSQL = "INSERT INTO Events (Type,LGTag,Notes,EventDate,EventTime) VALUES ('"
   strSQL = strSQL & strEventType & "','"
   strSQL = strSQL & intTag & "','"
   strSQL = strSQL & strNotes & "',#"
   strSQL = strSQL & Date() & "#,#"
   strSQL = strSQL & Time() & "#)"
   Application("Connection").Execute(strSQL)
  

End Sub%>

<% Sub UpdateEvent 
	
	Dim intEventID, strNotes, datDate, datTime, bolResolved, strSQL
	
	intEventID = Request.Form("EventID")
	strNotes = Request.Form("Notes")
	bolResolved = Request.Form("Resolved")
	datDate = Date()
	datTime = Time()
	
	strSQL = "UPDATE Events" & vbCRLF 
	strSQL = strSQL & "SET Notes='" & Replace(strNotes,"'","''") & "'" & vbCRLF
	
	If bolResolved Then
		strSQL = strSQL & ",Resolved=True,ResolvedDate=#" & datDate & "#,ResolvedTime=#" & datTime & "#" & vbCRLF
	End If
	
	strSQL = strSQL & "WHERE ID=" & intEventID
	Application("Connection").Execute(strSQL)

End Sub %>

<%Sub UpdateDevice

	Dim strSite, strRoom, bolInsured, bolActive, strSQL
	
	strSite = Request.Form("Site")
	strRoom = Request.Form("Room")
	bolInsured = Request.Form("Insured")
	bolActive = Request.Form("Active")
	
	If Not bolInsured Then
		bolInsured = False
	End If
	
	If Not bolActive Then
		bolActive = False
	End If
	
	strSQL = "UPDATE Devices Set "
	strSQL = strSQL & "Site='" & strSite & "',"
	strSQL = strSQL & "Room='" & strRoom & "',"
	strSQL = strSQL & "HasInsurance=" & bolInsured & ","
	strSQL = strSQL & "Active=" & bolActive & vbCRLF
	strSQL = strSQL & "WHERE LGTag='" & intTag & "'"
	Application("Connection").Execute(strSQL)

End Sub%>

<%Sub AssignDevice
   
   Dim intStudent, bolInsurance, strSQL, objDeviceCheck, objAssignmentCheck, objAssignedTo
   
   'Grade the data from the form
   intStudent = Request.Form("StudentID")
   bolInsurance = False
   
   'Make sure they submitted something
   If intStudent = "" Or intTag = "" Then
      strMessage = "<div Class=""Error"">Missing Data</div>"
   Else
  
      'Check and see if the tag is in the database
      strSQL = "SELECT ID FROM Devices WHERE LGTag='" & intTag & "'"
      Set objDeviceCheck = Application("Connection").Execute(strSQL)
      
      If Not objDeviceCheck.EOF Then
      
         'Check and see if the device is already assigned
         strSQL = "SELECT AssignedTo FROM Assignments WHERE LGTag='" & intTag & "' And Active=True"
         Set objAssignmentCheck = Application("Connection").Execute(strSQL)
         
         If objAssignmentCheck.EOF Then
         
            'Make sure the insurance variable is ready
            If Not bolInsurance Then
               bolInsurance = False
            End If
            
            'Create the assignment in the database
            strSQL = "INSERT INTO Assignments (LGTag, DateIssued, TimeIssued, Active, AssignedTo, HasInsurance, IssuedBy)" & vbCRLF
            strSQL = strSQL & "VALUES ('"
            strSQL = strSQL & intTag & "',#"
            strSQL = strSQL & Date & "#,#"
            strSQL = strSQL & Time & "#,"
            strSQL = strSQL & True & ","
            strSQL = strSQL & intStudent & ","
            strSQL = strSQL & bolInsurance & ",'"
            strSQL = strSQL & strUser & "')"
            Application("Connection").Execute(strSQL)
            
            'Update the student to show they have a device
            strSQL = "UPDATE People" & vbCRLF
            strSQL = strSQL & "SET HasDevice = True" & vbCRLF
            strSQL = strSQL & "WHERE ID = " & intStudent
            Application("Connection").Execute(strSQL)
            
            'Set the message to return
            strMessage = "<div class=""Bold"">" & intTag & " Assigned</div>"
         
         Else
         
            'Find out who the device is already assigned to
            strSQL = "SELECT FirstName, LastName" & vbCRLF
            strSQL = strSQL & "FROM People" &vbCRLF
            strSQL = strSQL & "WHERE ID=" & objAssignmentCheck(0)
            Set objAssignedTo = Application("Connection").Execute(strSQL)
            
            strMessage = "<div Class=""Error"">Device already assigned to " & objAssignedTo(0) & " " & objAssignedTo(1) & "</div>"
            
         End If
      
      Else
         strMessage = "<div Class=""Error"">Device not found</div>"
      End If
   
   End If
   
End Sub%>

<%Sub ReturnDevice   

   Dim intTag, strSerial, strNotes, strSQL, objDeviceCheck, objAssignmentCheck
   
   'Get the tag
   If IsNumeric(Request.QueryString("id")) Then
      intTag = Int(Request.QueryString("id"))
   Else
      intTag = Request.QueryString("id")
   End If
 
   'Get the serial number and the notes
   strSerial = Request.Form("Serial")
   strNotes = Request.Form("Notes")
   bolAdapterReturned = Request.Form("Adapter")
   bolCaseReturned = Request.Form("Case")
   
   If Not bolAdapterReturned Then
   	bolAdapterReturned = False
   End If
   
   If Not bolCaseReturned Then
   	bolCaseReturned = False
   End If
 
   'Make sure they submitted something
   If strSerial = "" And intTag = "" Then
      strMessage = "<div Class=""Error"">Missing Data</div>"
   Else
  
      'If they entered something in both fields we'll only use the tag
      If intTag <> "" Then
  
         'Check and see if the tag is in the database
         strSQL = "SELECT ID FROM Devices WHERE LGTag='" & intTag & "'"
         Set objDeviceCheck = Application("Connection").Execute(strSQL)
         
         If Not objDeviceCheck.EOF Then
         
            'Check and see if the device is assigned
            strSQL = "SELECT ID FROM Assignments WHERE LGTag='" & intTag & "' And Active=True"
            Set objAssignmentCheck = Application("Connection").Execute(strSQL)
            
            If Not objAssignmentCheck.EOF Then
            
               'Return the device
               UpdateDB objAssignmentCheck(0), strNotes 
               
               'Set the message to return
               strMessage = "<div Class=""Bold "">" & intTag & " is no longer assigned to " & strAssignedTo & "</div>"
            
            Else
               
               strMessage = "<div Class=""Error"">" & intTag & " is not currently assigned to anyone.</div>"
               
            End If
      
         Else
            strMessage = "<div Class=""Error"">Device not found</div>"
         End If
         
      Else
         
         'Check and see if the serial is in the database
         strSQL = "SELECT ID FROM Devices WHERE SerialNumber='" & strSerial & "'"
         Set objDeviceCheck = Application("Connection").Execute(strSQL)
         
         If Not objDeviceCheck.EOF Then
         
            strSQL = "SELECT Assignments.ID" & vbCRLF
            strSQL = strSQL & "FROM Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag" & vbCRLF
            strSQL = strSQL & "WHERE SerialNumber='" & strSerial & "' AND Assignments.Active=True"
            Set objAssignmentCheck = Application("Connection").Execute(strSQL)      

            If Not objAssignmentCheck.EOF Then
            
               'Return the device
               UpdateDB objAssignmentCheck(0), strNotes
               
               'Set the message to return
               strMessage = "<div Class=""Bold "">" & strSerial & " is no longer assigned to " & strAssignedTo
            
            Else
               
               strMessage = "<div Class=""Error "">""A device with serial " & strSerial & " is not assigned to anyone."
               
            End If
            
         Else
            
            Submit "<div Class=""Error "">""Device not found</div>"
         
         End If
         
      End If
   
   End If
   
End Sub%>

<%Sub UpdateDB(intID, strNotes)

   Dim strSQL, objStudentID, objAssignedTo

   'Update the assignment
   strSQL = "UPDATE Assignments SET Active=False,DateReturned=#" & Date & "#,TimeReturned=#" 
   strSQL = strSQL & Time & "#,ReturnedBy='" & strUser & "',Notes='" & Replace(strNotes,"'","''") & "'," 
   strSQL = strSQL & "AdapterReturned=" & bolAdapterReturned & ",CaseReturned=" & bolCaseReturned & vbCRLF
   strSQL = strSQL & "WHERE ID=" & intID
   Application("Connection").Execute(strSQL)
  
   'Get the students ID
   strSQL = "SELECT AssignedTo FROM Assignments WHERE ID=" & intID
   Set objStudentID = Application("Connection").Execute(strSQL)
  
   'Update the student to show they don't have a device
   strSQL = "UPDATE People" & vbCRLF
   strSQL = strSQL & "SET HasDevice = False" & vbCRLF
   strSQL = strSQL & "WHERE ID = " & objStudentID(0)
   Application("Connection").Execute(strSQL)
   
   'Get the students name from the database
   strSQL = "SELECT FirstName, LastName" & vbCRLF
   strSQL = strSQL & "FROM People" &vbCRLF
   strSQL = strSQL & "WHERE ID=" & objStudentID(0)
   Set objAssignedTo = Application("Connection").Execute(strSQL)
   
   strAssignedTo = objAssignedTo(0) & " " & objAssignedTo(1)

End Sub%>

<%Function GetAge(strDate)

	If Month(Date) >= Month(strDate) Then 
		GetAge = Year(Date) - Year(strDate) + 1
	Else
		GetAge = Year(Date) - Year(strDate)
	End If

End Function %>

<%Function ShortenDate(strDate)
	
	If Not IsNull(strDate) Then
		If strDate <> "" Then
			ShortenDate = Left(strDate,Len(strDate) - 4) & Right(strDate,2)
		End If
	End If

End Function %>

<%Function ShortenTime(strTime)
	
	If Not IsNull(strTime) Then
		If strTime <> "" Then
			ShortenTime = Left(strTime,Len(strTime) - 6) & " " & Right(strTime,2)
		End If
	End If

End Function %>

<%Function GetRole(intYear)

   Dim datToday, intMonth, intCurrentYear, intGrade
   
   'Set the role to the correct staff group
   Select Case intYear
      Case 10
         GetRole = "Tech Staff"   
      Case 20
         GetRole = "Teacher"
      Case 30
         GetRole = "Staff"
      Case 40
         GetRole = "TA"
      Case 50
         GetRole = "Sub"
      Case 60
         GetRole = "Student Teacher"
   End Select 
      
   'Convert the graduating year to a grade
   datToday = Date
   intMonth = DatePart("m",datToday)
   intCurrentYear = Right(DatePart("yyyy",datToday),2)
   intYear = Right(intYear,2)

   If intMonth >= 7 And intMonth <= 12 Then
      intCurrentYear = intCurrentYear + 1
   End If

   intGrade = 12 - (intYear - intCurrentYear)

	If GetRole = "" Then
	
		'Change the grade number into text
		Select Case intGrade
			Case 0
				GetRole = "Kindergarten"
			Case 1
				GetRole = "1st"
			Case 2
				GetRole = "2nd"
			Case 3
				GetRole = "3rd"
			Case 4
				GetRole = "4th"   
			Case 5
				GetRole = "5th"
			Case 6
				GetRole = "6th"
			Case 7
				GetRole = "7th"
			Case 8
				GetRole = "8th"
			Case 9
				GetRole = "9th"
			Case 10
				GetRole = "10th"
			Case 11
				GetRole = "11th"
			Case 12
				GetRole = "12th"
			Case Else
				GetRole = "Graduated"
		End Select

	End If

End Function%>

<%
' Anything below here should exist on all pages 
%>

<%Sub DenyAccess 

   'If we're not using basic authentication then send them to the login screen
   If bolShowLogout Then
      Response.Redirect("login.asp?action=logout")
   Else
   
   SetupSite
   
   %>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title><%=Application("SiteName")%></title>
      <link rel="stylesheet" type="text/css" href="style.css" /> 
      <link rel="apple-touch-icon" href="images/inventory.png" /> 
      <link rel="shortcut icon" href="images/inventory.ico" />
      <meta name="viewport" content="width=device-width" />
   </head>
   <body>
      <center><b>Access Denied</b></center>
   </body>
   </html>
   
<% End If

End Sub%>

<%Function AccessGranted

   Dim objNetwork, strUserAgent, strSQL, strRole, objNameCheckSet

   'Redirect the user the SSL version if required
   If Application("ForceSSL") Then
      If Request.ServerVariables("SERVER_PORT")=80 Then
         If Request.ServerVariables("QUERY_STRING") = "" Then
            Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
         Else
            Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
         End If
      End If
   End If

   'Get the users logon name
   Set objNetwork = CreateObject("WSCRIPT.Network")   
   strUser = objNetwork.UserName
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

   'Check and see if anonymous access is enabled
   If LCase(Left(strUser,4)) = "iusr" Then
      strUser = GetUser
      bolShowLogout = True
   Else
      bolShowLogout = False
   End If

   'Build the SQL string, this will check the userlevel of the user.
   strSQL = "Select Role" & vbCRLF
   strSQL = strSQL & "From Sessions" & vbCRLF
   strSQL = strSQL & "WHERE UserName='" & strUser & "' And SessionID='" & Request.Cookies("SessionID") & "'"
   Set objNameCheckSet = Application("Connection").Execute(strSQL)
   strRole = objNameCheckSet(0)

   If strRole = "Admin" or strRole = "User" Then
      AccessGranted = True
   Else
      AccessGranted = False
   End If

End Function%>

<%Function GetUser

   Const USERNAME = 1

   Dim strUserAgent, strSessionID, objSessionLookup, strSQL
   
   'Get some needed data
   strSessionID = Request.Cookies("SessionID")
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
   
   'Send them to the logon screen if they don't have a Session ID
   If strSessionID = "" Then
      SendToLogonScreen

   'Get the username from the database
   Else
   
      strSQL = "SELECT ID,UserName,SessionID,IPAddress,UserAgent,ExpirationDate FROM Sessions "
      strSQL = strSQL & "WHERE UserAgent='" & Left(Replace(strUserAgent,"'","''"),250) & "' And SessionID='" & Replace(strSessionID,"'","''") & "'"
      strSQL = strSQL & " And ExpirationDate > Date()"
      Set objSessionLookup = Application("Connection").Execute(strSQL)
      
      'If a session isn't found for then kick them out
      If objSessionLookup.EOF Then
         SendToLogonScreen
      Else
         GetUser = objSessionLookup(USERNAME)
      End If
   End If  
   
End Function%>

<%Function IsMobile

   Dim strUserAgent

   'Get the User Agent from the client so we know what browser they are using
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

   'Check the user agent for signs they are on a mobile browser
   If InStr(strUserAgent,"iPhone") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"iPad") Then
      IsMobile = False
   ElseIf InStr(strUserAgent,"Android") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"Windows Phone") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"BlackBerry") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"Nintendo") Then
      IsMobile = True 
   ElseIf InStr(strUserAgent,"PlayStation Vita") Then
      IsMobile = True
   Else
      IsMobile = False
   End If 
   
   If InStr(strUserAgent,"Nexus 9") Then  
      IsMobile = False
   End If
End Function%>

<%Sub SendToLogonScreen

   Dim strReturnLink, strSourcePage
      
   'Build the return link before sending them away.
   strReturnLink =  "?" & Request.ServerVariables("QUERY_STRING")
   strSourcePage = Request.ServerVariables("SCRIPT_NAME")
   strSourcePage = Right(strSourcePage,Len(strSourcePage) - InStrRev(strSourcePage,"/"))
   If strReturnLink = "?" Then
      strReturnLink = "?SourcePage=" & strSourcePage
   Else
      strReturnLink = strReturnLink & "&SourcePage=" & strSourcePage
   End If
   
   Response.Redirect("login.asp" & strReturnLink)
   
End Sub %>

<%Sub SetupSite
   
   If IsMobile Then
      strSiteVersion = "Mobile"
   Else
      strSiteVersion = "Full"
   End If
   
End Sub%>