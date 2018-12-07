<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 8/22/15
'Last Updated 1/14/18

'This page shows a list of events as a result of a search.

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim objEvents, strEventType, strCategory, strView, strWarranty, strComplete, objEventTypes
Dim intEventCount, strNotes, strSubmitTo, objCategories, intEventID
Dim datStartDate, datEndDate, strSearchMessage, strBackLink, strColumns, intEventNumber
Dim strModel, strSite, intDeviceYear, objLastNames

'See if the user has the rights to visit this page
If AccessGranted Then
	ProcessSubmissions 
Else
	DenyAccess
End If %>

<%Sub ProcessSubmissions

	Dim strSQL, strSQLWhere
	
	'Get the variables from the URL
	intEventNumber = Request.QueryString("EventNumber")
	strEventType = Request.QueryString("EventType")
	strCategory = Request.QueryString("Category")
	strModel = Request.QueryString("EventModel")
	strSite = Request.QueryString("EventSite")
	intDeviceYear = Request.QueryString("Year")
	strNotes = Request.QueryString("EventNotes")
	strView = Request.QueryString("View")
	datStartDate = Request.QueryString("StartDate")
	datEndDate = Request.QueryString("EndDate")
	strBackLink = BackLink
	
	If datStartDate <> "" Then
		If Not IsDate(datStartDate) Then
			datStartDate = GetStartOfFiscalYear(Date)
		End If
	End If

	If datEndDate <> "" Then
		If Not IsDate(datEndDate) Then
			datEndDate = Date
		End If
	End If
	
	Select Case Request.QueryString("Warranty")
		Case "Yes"
			strWarranty = "Yes"
		Case "No"
			strWarranty = "No"
	End Select
	
	Select Case Request.QueryString("Complete")
		Case "Yes"
			strComplete = "Yes"
		Case "No"
			strComplete = "No"
		Case "All"
			strComplete = "All"
	End Select

	'Check and see if anything was submitted to the site
	Select Case Request.Form("Submit")
		Case "Update Event"
			UpdateEvent
		Case "Save"
			SaveSearch
			
	End Select
	
	'Get the list of devices
	strSQLWhere = "WHERE "
	If intEventNumber <> "" Then
		If IsNumeric(intEventNumber) Then
			strSQLWhere = strSQLWhere & "Events.ID=" & intEventNumber & " AND "
		End If
	End If
	If strEventType <> "" Then
		strSQLWhere = strSQLWhere & "Events.Type='" & Replace(strEventType,"'","''") & "' AND "
	End If
	If strCategory <> "" Then
		strSQLWhere = strSQLWhere & "Events.Category='" & Replace(strCategory,"'","''") & "' AND "
	End If
	If strNotes <> "" Then
		strSQLWhere = strSQLWhere & "Events.Notes Like '%" & Replace(strNotes,"'","''") & "%' AND "
	End If
	If strModel <> "" Then
		strSQLWhere = strSQLWhere & "Events.Model Like '%" & Replace(strModel,"'","''") & "%' AND "
	End If
	If strSite <> "" Then
		strSQLWhere = strSQLWhere & "Events.Site='" & Replace(strSite,"'","''") & "' AND "
	End If
	If intDeviceYear <> "" Then
		strSQLWhere = strSQLWhere & _
		"Devices.DatePurchased>=#" & DateAdd("yyyy",intDeviceYear * -1,Date) & "# AND " & _
		"Devices.DatePurchased<=#" & DateAdd("yyyy",(intDeviceYear -1) * -1,Date) & "# AND "
	End If
	
	Select Case strWarranty
		Case "Yes"
			strSQLWhere = strSQLWhere & "Events.Warranty=True AND "
		Case "No"
			strSQLWhere = strSQLWhere & "Events.Warranty=False AND "
	End Select
	
	Select Case strComplete
		Case "Yes"
			strSQLWhere = strSQLWhere & "Events.Resolved=True AND "
		Case "No"
			strSQLWhere = strSQLWhere & "Events.Resolved=False AND "
		Case "All"
		Case Else
			If intEventNumber = "" Then
				strSQLWhere = strSQLWhere & "Events.Resolved=False AND "
			End If
	End Select
	
	If datStartDate <> "" Then
		strSQLWhere = strSQLWhere & "Events.EventDate>=#" & datStartDate & "# AND "
	End If

	If datEndDate <> "" Then
		strSQLWhere = strSQLWhere & "Events.EventDate<=#" & datEndDate & "# AND "
	End If
	
	If strSQLWhere <> "WHERE " Then
		strSQLWhere = Left(strSQLWhere,Len(strSQLWhere) - 5)
	Else
		strSQLWhere = ""
	End If
		
	strSQL = "SELECT Events.ID,Events.Type,Events.Notes,Events.EventDate,Events.EventTime,Events.Resolved,Events.ResolvedDate,Events.ResolvedTime,Events.Category," & _
		"Events.Warranty,Events.LGTag,Events.UserID,Events.Site,Events.Model,Events.EnteredBy,Events.CompletedBy,Devices.SerialNumber" & vbCRLF
	strSQL = strSQL & "FROM Devices INNER JOIN Events ON Devices.LGTag = Events.LGTag " & strSQLWhere
	Set objEvents = Application("Connection").Execute(strSQL)

	'If no user is found send them back to the index page.
	If objEvents.EOF Then
		Response.Redirect("index.asp?Error=NoEventsFound")
	Else
		intEventCount = 0
		Do Until objEvents.EOF 
			intEventCount = intEventCount + 1
			objEvents.MoveNext
		Loop
		objEvents.MoveFirst
	End If

	'Get the URL used to submit forms
	If Request.ServerVariables("QUERY_STRING") = "" Then
		strSubmitTo = "events.asp"
	Else   
		strSubmitTo = "events.asp?" & Request.ServerVariables("QUERY_STRING")
	End If

	'Get the list of event types for the event types drop down menu
	strSQL = "SELECT EventType FROM EventTypes WHERE Active=True ORDER BY EventType"
	Set objEventTypes = Application("Connection").Execute(strSQL)
	
	'Get the list of categories from the category drop down menu
	strSQL = "SELECT Category FROM Categories WHERE Active=True ORDER BY Category"
	Set objCategories = Application("Connection").Execute(strSQL)

	'Get the list of lastnames for the auto complete
	strSQL = "SELECT DISTINCT LastName FROM People WHERE Active=True"
	Set objLastNames = Application("Connection").Execute(strSQL)

	'Set up the variables needed for the site then load it
	SetupSite
	If LCase(strView) = "table" Then
		strSiteVersion = "Full"
	End If
	DisplaySite
	
End Sub%>

<%Sub DisplaySite %>

	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	<html>
	<head>
		<title><%=Application("SiteName")%></title> 
		<link rel="stylesheet" type="text/css" href="../style.css" /> 
		<link rel="apple-touch-icon" href="../images/inventory.png" /> 
		<link rel="shortcut icon" href="../images/inventory.ico" />
		<meta name="viewport" content="width=device-width,user-scalable=0" />
		<meta name="theme-color" content="#333333">
		<link rel="stylesheet" href="../assets/css/jquery-ui.css">
		<script src="../assets/js/jquery.js"></script>
		<script src="../assets/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="../assets/css/jquery.dataTables.min.css">
		<link rel="stylesheet" href="../assets/css/buttons.dataTables.min.css">
		<script src="../assets/js/jquery.dataTables.min.js"></script>
		<script src="../assets/js/dataTables.buttons.min.js"></script>
		<script src="../assets/js/buttons.colVis.min.js"></script>
		<script src="../assets/js/buttons.html5.min.js"></script>
		<script src="../assets/js/jszip.min.js"></script>
		<script type="text/javascript">
			$(document).ready( function () {
			
			<%	If Not IsMobile And Not IsiPad Then%>
					$( document ).tooltip({track: true});
				End If %>	
			
				var table = $('#ListView').DataTable( {
					paging: false,
					"info": false,
					"autoWidth": false,
					dom: 'Bfrtip',
					// stateSave: true,
					buttons: [
						{
							extend: 'colvis',
							text: 'Show/Hide Columns'
						}
				<%	If Not IsMobile Then %>		
						,
						{
							extend: 'csvHtml5',
							text: 'Download CSV'
						}
				<%	End If %>
					]
				
				});
				
				<%	If Not objLastNames.EOF Then %>
						var possibleLastNames = [
					<%	Do Until objLastNames.EOF %>	
							"<%=objLastNames(0)%>",
						<%	objLastNames.MoveNext
						Loop %>
					];
						$( "#LastNames" ).autocomplete({
							source: possibleLastNames
						});
				<%	End If %>
				
		<%	If IsMobile Then %>	
				table.columns([3,4,5,6,7,8,9,10,11,12,13]).visible(false);	
		<%	Else %>		
				table.columns([4,5,6,8,10,11,12]).visible(false);
		<%	End If %>

			} );
		</script>
	</head>

	<body class="<%=strSiteVersion%>">
	
		<div class="Header"><%=Application("SiteName")%> (<%=intEventCount%>)</div>
		<div>
			<ul class="NavBar" align="center">
				<li><a href="index.asp"><img src="../images/home.png" title="Home" height="32" width="32"/></a></li>
				<li><a href="search.asp"><img src="../images/search.png" title="Search" height="32" width="32"/></a></li>
				<li><a href="stats.asp"><img src="../images/stats.png" title="Stats" height="32" width="32"/></a></li>
				<li><a href="log.asp"><img src="../images/log.png" title="System Log" height="32" width="32"/></a></li>
				<li><a href="add.asp"><img src="../images/add.png" title="Add Person or Device" height="32" width="32"/></a></li>
				<li><a href="login.asp?action=logout"><img src="../images/logout.png" title="Log Out" height="32" width="32"/></a></li>
			</ul>
		</div>
		<%JumpToDevice%>
<%	If Not objEvents.EOF Then
	
		Select Case LCase(strView)
			Case "table"
				ShowEventTable
			Case "card"
				ShowEventCards
			Case Else
				
				If IsMobile Then
				
					If LCase(Application("DefaultViewMobile")) = "table" Then
						If intEventCount < Application("CardThreshold") Then
							ShowEventCards
						Else
							ShowEventTable
						End If
					Else
						ShowEventCards
					End If
				
				Else
			
					If LCase(Application("DefaultView")) = "table" Then
						If intEventCount < Application("CardThreshold") Then
							ShowEventCards
						Else
							ShowEventTable
						End If
					Else
						ShowEventCards
					End If
					
				End If
		End Select
		SaveAsSearch
	End If %>
		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
	</body>

	</html>

<%End Sub%>

<%Sub ShowEventCards 

	Dim strSelected, strWarrantyChecked%>
	
	<div class="ViewButton">
		<a href="<%=SwitchView("Table")%>"><img src="../images/table.png" title="Table View" height="32" width="32"/></a>
	</div>
	<div class="center"><%=FilterBar%></div>
	<div Class="<%=strColumns%>">
<%	If Not objEvents.EOF Then 

		Do Until objEvents.EOF 
		
			If objEvents(9) Then
				strWarrantyChecked = "checked=""checked"""
			Else
				strWarrantyChecked = ""
			End If
	
			If objEvents(5) Then %> 

				<div class="Card NormalCard">
					<div class="CardTitle">Event <%=objEvents(0)%></div>
					<div>
						<div Class="CardColumn1">Asset Tag: </div>
						<div Class="CardColumn2">
							<a href="device.asp?Tag=<%=objEvents(10)%><%=strBackLink%>"><%=objEvents(10)%></a>
						</div>
					</div>
					<div>
						<div Class="CardColumn1">Event Type: </div>
						<div Class="CardColumn2"><%=objEvents(1)%></div>
					</div>
					<div>
						<div Class="CardColumn1">Date: </div>
				<%	If ShortenDate(objEvents(3)) = ShortenDate(objEvents(6)) Then %>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3))%></div>
				<%	Else %>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3)) & " - " & ShortenDate(objEvents(6))%></div>
				<%	End If %>
					</div>
					<div>
						<div Class="CardColumn1">Category: </div>
						<div Class="CardColumn2"><%=objEvents(8)%></div>
					</div>
					<div>
						<div Class="CardColumn1">Warranty: </div>
						<div Class="CardColumn2"><input type="checkbox" name="Warranty" value="True" <%=strWarrantyChecked%> /></div>
					</div>
				<%	If objEvents(2) <> "" And Not IsNull(objEvents(2)) Then %> 
						<div>Notes: </div>
						<div><%=objEvents(2)%></div>
				<%	End If %>
				</div>
				
		<%	Else %>
		
				<div class="Card NormalCard">
					<form method="POST" action="<%=strSubmitTo%>">
					<input type="hidden" name="EventID" value="<%=objEvents(0)%>" />
					<div class="CardTitle">Event <%=objEvents(0)%></div>
					<div>
						<div Class="CardColumn1">Asset Tag: </div>
						<div Class="CardColumn2">
							<a href="device.asp?Tag=<%=objEvents(10)%><%=strBackLink%>"><%=objEvents(10)%></a>
						</div>
					</div>
					<div>	
						<div Class="CardColumn1">Event Type: </div>
						<div Class="CardColumn2">
							<select Class="Card" name="EventType">
								<option value=""></option>
						<%	If Not objEventTypes.EOF Then
								Do Until objEventTypes.EOF 
									If objEvents(1) = objEventTypes(0) Then
										strSelected = "selected=""selected"""
									Else
										strSelected = ""
									End If %>
									<option value="<%=objEventTypes(0)%>" <%=strSelected%>><%=objEventTypes(0)%></option>
								<%	objEventTypes.MoveNext
								Loop 
							End If 
							objEventTypes.MoveFirst %>
								<option value="Decommission Device">Decommission Device</option>
							</select>
						</div>
					</div>
					<div>
						<div Class="CardColumn1">Submitted: </div>
						<div Class="CardColumn2"><%=ShortenDate(objEvents(3)) & " " & ShortenTime(objEvents(4))%></div>
					</div>
					<div Class="CardColumn1">Category:</div>
					<div Class="CardColumn2">
						<select Class="Card" name="Category">
							<option value=""></option>
					<%	If Not objCategories.EOF Then
							Do Until objCategories.EOF 
								If objCategories(0) = objEvents(8) Then
									strSelected = "selected=""selected"""
								Else
									strSelected = "" 
								End If %>
								<option value="<%=objCategories(0)%>" <%=strSelected%>><%=objCategories(0)%></option>
							<%	objCategories.MoveNext
							Loop
							objCategories.MoveFirst 
						End If %>
						</select>
					</div>
					<div Class="CardColumn1">Warranty:</div>
					<div Class="CardColumn2">
						<input type="checkbox" name="Warranty" value="True" <%=strWarrantyChecked%> />
					</div>
					<div>
						<div Class="CardColumn1">Complete: </div>
						<div Class="CardColumn2">
							<input Class="Card" type="checkbox" value="TRUE" name="Resolved" />
						</div>
					</div>
					<div>Notes: </div>
					<div>
						<textarea Class="Card" rows="5" name="Notes" cols="90" style="width: 99%;"><%=objEvents(2)%></textarea>
					</div>
					<div>&nbsp;</div>
					<div Class="Button"><input type="submit" value="Update Event" name="Submit" /></div>
			<%	If CInt(intEventID) = CInt(objEvents(0)) Then %>
					<div>
						<div class="Information">Updated</div>
					</div>
			<%	End If %>
					</form>
				</div>
			
		<%	End If
			objEvents.MoveNext
		Loop %>
		</div>
<%	End If 

End Sub%>

<%Sub ShowEventTable

	Dim strWarrantyInfo, strCompleteInfo, strSQL, objName%>

	<div class="ViewButton">
		<a href="<%=SwitchView("Card")%>"><img src="../images/card.png" title="Card View" height="32" width="32"/></a>
	</div>
	<div class="center"><%=Replace(FilterBar,"?","?View=Table&")%></div>

	<div>
		<table align="center" Class="ListView" id="ListView">
			<thead>
			<th>Event</th>
			<th>Asset Tag</th>
			<th>Type</th>
			<th>Category</th>
			<th>Model</th>
			<th>Serial Number</th>
			<th>Site</th>
			<th>Start Date</th>
			<th>End Date</th>
			<th>User</th>
			<th>Warranty</th>
			<th>Entered By</th>
			<th>Completed By</th>
			<th>Complete</th>
			<th>Event Notes</th>
			</thead>
			<tbody>
<%	Do Until objEvents.EOF

		If objEvents(9) Then
			strWarrantyInfo = "Yes"
		Else
			strWarrantyInfo = "No"
		End If 
		
		If objEvents(5) Then
			strCompleteInfo = "Yes"
		Else
			strCompleteInfo = "No"
		End If%>
		
			<tr>
				<td id="center"><%=objEvents(0)%></th>
				<td id="center"><a href="device.asp?Tag=<%=objEvents(10)%><%=strBackLink%>"><%=objEvents(10)%></a></td>
				<td><a href="events.asp?EventType=<%=objEvents(1)%>&View=Table"><%=objEvents(1)%></a></td>
				<td><a href="events.asp?Category=<%=objEvents(8)%>&View=Table"><%=objEvents(8)%></a></td>
				<td><a href="events.asp?EventModel=<%=objEvents(13)%>&View=Table"><%=objEvents(13)%></a></td>
				<td id="center"><a href="device.asp?Tag=<%=objEvents(10)%><%=strBackLink%>"><%=objEvents(16)%></a></td>
				<td><a href="events.asp?EventSite=<%=objEvents(12)%>&View=Table"><%=objEvents(12)%></a></td>
				<td><%=ShortenDate(objEvents(3))%></td>
				<td><%=ShortenDate(objEvents(6))%></td>
			<%	If objEvents(11) <> "" Then 
			
					strSQL = "SELECT FirstName,LastName,UserName FROM People WHERE ID=" & objEvents(11) 
					Set objName = Application("Connection").Execute(strSQL)
					
					If Not objName.EOF Then %>
						<td>
							<a href="user.asp?UserName=<%=objName(2)%><%=strBackLink%>"><%=objName(1)%>, <%=objName(0)%></a>
						</td>
				<%	Else %>		
						<td></td>
				<%	End If %>
					
			<%	Else %>
					<td></td>
			<%	End If %>
				<td id="center"><a href="events.asp?Warranty=<%=strWarrantyInfo%>&View=Table"><%=strWarrantyInfo%></a></td>
				
			<%	If objEvents(14) <> "" Then 
			
					strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objEvents(14) & "'"
					Set objName = Application("Connection").Execute(strSQL)
					
					If Not objName.EOF Then %>
						<td><%=objName(1)%>, <%=objName(0)%></td>
				<%	Else %>		
						<td></td>
				<%	End If %>
					
			<%	Else %>
					<td></td>
			<%	End If %>
			
			<%	If objEvents(15) <> "" Then 
			
					strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & objEvents(15) & "'"
					Set objName = Application("Connection").Execute(strSQL)
					
					If Not objName.EOF Then %>
						<td><%=objName(1)%>, <%=objName(0)%></td>
				<%	Else %>		
						<td></td>
				<%	End If %>
					
			<%	Else %>
					<td></td>
			<%	End If %>
				
				<td id="center"><a href="events.asp?Complete=<%=strCompleteInfo%>&View=Table"><%=strCompleteInfo%></a></td>
				<td><%=Replace(objEvents(2),vbCRLF,"<br />")%></td>
			</tr>
	<%	objEvents.MoveNext
	Loop %>
			</tbody>
		</table>	
	</div>
<%End Sub %>

<%	Sub UpdateEvent 
	
	Dim strNotes, datDate, datTime, bolResolved, strSQL, strCategory, bolWarranty, strUserName, strOldEventType, objDeviceLookup
	Dim objUserID, intUserID, objOldValues, strOldNotes, strOldCategory, bolOldWarranty, strEventType, intTag
	
	'Get the variables from the form
	intEventID = Request.Form("EventID")
	strNotes = Request.Form("Notes")
	strCategory = Request.Form("Category")
	bolResolved = Request.Form("Resolved")
	strEventType = Request.Form("EventType")
	
	If Request.Form("Warranty") = "True" Then
		bolWarranty = True
	Else
		bolWarranty = False
	End If
	
	datDate = Date()
	datTime = Time()
	
	'Get the device's tag
	strSQL = "SELECT LGTag FROM Events WHERE ID=" & intEventID
	Set objDeviceLookup = Application("Connection").Execute(strSQL)
	intTag = objDeviceLookup(0)
	
	'Get the userID
	strSQL = "SELECT AssignedTo FROM Assignments WHERE Active=True AND LGTag='" & intTag & "'"
	Set objUserID = Application("Connection").Execute(strSQL)
	If Not objUserID.EOF Then
		intUserID = objUserID(0)
	Else
		intUserID = 0
	End If
	
	'Get the current values from the database
	strSQL = "SELECT Notes,Category,Type,Warranty FROM Events WHERE ID=" & intEventID
	Set objOldValues = Application("Connection").Execute(strSQL)
	
	'Record the old values before they change
	strOldNotes = objOldValues(0)
	strOldCategory = objOldValues(1)
	strOldEventType = objOldValues(2)
	bolOldWarranty = objOldValues(3)
	
	strSQL = "UPDATE Events" & vbCRLF 
	strSQL = strSQL & "SET Notes='" & Replace(strNotes,"'","''") & "'," 
	strSQL = strSQL & "Category='" & Replace(strCategory,"'","''") & "',"
	strSQL = strSQL & "Type='" & Replace(strEventType,"'","''") & "',"
	strSQL = strSQL & "Warranty=" & bolWarranty
	
	If bolResolved Then
		strSQL = strSQL & ",Resolved=True,ResolvedDate=#" & datDate & "#,ResolvedTime=#" & datTime & "#," & _
			"CompletedBy='" & strUser & "'"
		Application("Connection").Execute("UPDATE Devices SET HasEvent=False WHERE LGTag='" & intTag & "'")
	End If
	
	strSQL = strSQL & vbCRLF & "WHERE ID=" & intEventID
	Application("Connection").Execute(strSQL)
	
	'Log the updated values
	If intUserID > 0 Then 
		strUserName = GetUserName(intUserID)
	Else
		strUserName = ""
	End If
	
	If strNotes <> strOldNotes Then
		UpdateLog "EventUpdatedNotes",intTag,strUserName,strOldNotes,strNotes,intEventID
	End If
	If strCategory <> strOldCategory Then
		UpdateLog "EventUpdatedCategory",intTag,strUserName,strOldCategory,strCategory,intEventID
	End If
	If strEventType <> strOldEventType Then
		UpdateLog "EventUpdatedType",intTag,strUserName,strOldEventType,strEventType,intEventID
	End If
	If bolWarranty <> bolOldWarranty Then
		UpdateLog "EventUpdatedWarranty",intTag,strUserName,bolOldWarranty,bolWarranty,intEventID
	End If
	If bolResolved Then
		UpdateLog "EventClosed",intTag,strUserName,"","",intEventID
	End If
	
End Sub %>

<%Sub JumpToDevice%>

	<div Class="HeaderCard">
		<form method="POST" action="search.asp">
		Asset tag: <input class="Card InputWidthSmall" type="text" name="SmartBox" id="LastNames" />
		<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		<input type="hidden" value="" name="BOCESTag" />
		</form>
	</div>

<%End Sub%>

<%Function FilterBar

	If intEventNumber <> "" Then
		FilterBar = FilterBar & "Number = <a href=""events.asp?EventNumber=" & intEventNumber & """>" & intEventNumber & "</a> | "
	End If
	
	If strEventType <> "" Then
		FilterBar = FilterBar & "Type = <a href=""events.asp?EventType=" & strEventType & """>" & strEventType & "</a> | "
	End If
	
	If strCategory <> "" Then
		FilterBar = FilterBar & "Category = <a href=""events.asp?Category=" & strCategory & """>" & strCategory & "</a> | "
	End If
	
	If strModel <> "" Then
		FilterBar = FilterBar & "Model = <a href=""events.asp?EventModel=" & strModel & """>" & strModel & "</a> | "
	End If
	
	If strSite <> "" Then
		FilterBar = FilterBar & "Site = <a href=""events.asp?EventSite=" & strSite & """>" & strSite & "</a> | "
	End If
	
	If intDeviceYear <> "" Then
		FilterBar = FilterBar & "Year = <a href=""devices.asp?Year=" & intDeviceYear & """>" & intDeviceYear & "</a> | "
	End If
	
	If strWarranty <> "" Then
		FilterBar = FilterBar & "Warranty = <a href=""events.asp?Warranty=" & strWarranty & """>" & strWarranty & "</a> | "
	End If
	
	If strComplete <> "" Then
		FilterBar = FilterBar & "Complete = <a href=""events.asp?Complete=" & strComplete & """>" & strComplete & "</a> | "
	End If
	
	If strNotes <> "" Then
		FilterBar = FilterBar & "Notes = <a href=""events.asp?Notes=" & strNotes & """>" & strNotes & "</a> | "
	End If
	
	If datStartDate <> "" Then
		FilterBar = FilterBar & "Start Date = <a href=""events.asp?StartDate=" & datStartDate & """>" & datStartDate & "</a> | "
	End If
	
	If datEndDate <> "" Then
		FilterBar = FilterBar & "End Date = <a href=""events.asp?EndDate=" & datEndDate & """>" & datEndDate & "</a> | "
	End If
	
	If FilterBar <> "" Then
		FilterBar = Left(FilterBar,Len(FilterBar) - 3)
	End If

End Function %>

<%Function BackLink

	Dim strPage
	
	If Request.QueryString("Back") = "" Then
		BackLink = "&Back=" & Server.UrlEncode(Request.ServerVariables("QUERY_STRING"))
		strPage = Right(Request.ServerVariables("SCRIPT_NAME"),Len(Request.ServerVariables("SCRIPT_NAME")) - InStrRev(Request.ServerVariables("SCRIPT_NAME"),"/"))
		BackLink = BackLink & "&Page=" & strPage
	Else
		BackLink = "&Back=" & Server.UrlEncode(Request.QueryString("Back")) & "&Page=" & Request.QueryString("Page")
	End If
	 
End Function%>

<%Sub SaveAsSearch 

	Dim strPage

	strPage = Right(Request.ServerVariables("SCRIPT_NAME"),Len(Request.ServerVariables("SCRIPT_NAME")) - InStrRev(Request.ServerVariables("SCRIPT_NAME"),"/"))
	strPage = strPage & "?" & Request.ServerVariables("QUERY_STRING")%>
	
	<br />
	
	<div class="Card NormalCard">
		<form method="POST" action="<%=strPage%>">
			<div class="CardTitle">Save Search</div>
			<div Class="CardColumn1">Search Name:</div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="SearchName""/>
			</div>
			<div>
				<div class="Button"><input type="submit" value="Save" name="Submit" /></div>
			</div>
		<%	If strSearchMessage <> "" Then %>
			<div>
				<%=strSearchMessage%>
			</div>
	<%	End If %> 
		</form>
		</div>

<%End Sub%>

<%Sub SaveSearch

	Dim strPage, strSearchName, strQueryString, strSQL
	
	strPage = Right(Request.ServerVariables("SCRIPT_NAME"),Len(Request.ServerVariables("SCRIPT_NAME")) - InStrRev(Request.ServerVariables("SCRIPT_NAME"),"/"))
	strQueryString = Replace(Replace(Request.ServerVariables("QUERY_STRING"),"'","''")," ","%20")
	strSearchName = Request.Form("SearchName")
	
	strSQL = "INSERT INTO SavedSearches (SearchName,Page,QueryString,UserName,Active) VALUES (" 
	strSQL = strSQL & "'" & Replace(strSearchName,"'","''") & "',"
	strSQL = strSQL & "'" & strPage & "',"
	strSQL = strSQL & "'" & strQueryString & "',"
	strSQL = strSQL & "'" & strUser & "',"
	strSQL = strSQL & "True)"
	Application("Connection").Execute(strSQL)
	
	UpdateLog "SearchSaved","","","",strSearchName,""
	
	strSearchMessage = "<div Class=""Information"">Saved</div>"

End Sub%>

<%Function SwitchView(strView)

	Dim strURL
	
	If LCase(Request.QueryString("View")) = "" Then
		If Request.ServerVariables("QUERY_STRING") = "" Then
			SwitchView = "devices.asp?View=" & strView
		Else
			SwitchView = "devices.asp?" & Request.ServerVariables("QUERY_STRING") & "&View=" & strView
		End If	
	Else
		Select Case LCase(strView)
			Case "card"  
				SwitchView = "devices.asp?" & Replace(Request.ServerVariables("QUERY_STRING"),"Table","Card")
			Case "table"
				SwitchView = "devices.asp?" & Replace(Request.ServerVariables("QUERY_STRING"),"Card","Table")
		End Select
	End If
		
End Function%>

<%Function GetStartOfFiscalYear(datToday)

	If IsDate(datToday) Then
		If Month(datToday) >= 7 Then
			GetStartOfFiscalYear = "7/1/" & Year(datToday)
		Else
			GetStartOfFiscalYear = "7/1/" & Year(datToday) - 1
		End If
	Else
		GetStartOfFiscalYear = GetStartOfFiscalYear(Date)
	End If
	
End Function%>

<%Function SwitchView(strView)

	Dim strURL
	
	If LCase(Request.QueryString("View")) = "" Then
		If Request.ServerVariables("QUERY_STRING") = "" Then
			SwitchView = "events.asp?View=" & strView
		Else
			SwitchView = "events.asp?" & Request.ServerVariables("QUERY_STRING") & "&View=" & strView
		End If
	Else
		Select Case LCase(strView)
			Case "card"  
				SwitchView = "events.asp?" & Replace(Request.ServerVariables("QUERY_STRING"),"Table","Card")
			Case "table"
				SwitchView = "events.asp?" & Replace(Request.ServerVariables("QUERY_STRING"),"Card","Table")
		End Select
	End If
		
End Function%>

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

<%Function GetUserName(UserID)

	Dim strSQL, objUserInfo
	
	If UserID <> "" Then
		
		strSQL = "SELECT UserName FROM People WHERE ID=" & UserID
		Set objUserInfo = Application("Connection").Execute(strSQL)
		
		If Not objUserInfo.EOF Then
			GetUserName = objUserInfo(0)
		Else
			GetUserName = UserID
		End If
	Else
		GetUserName = UserID
	End If

End Function%>

<%Function GetDisplayName(Username)

	Dim strSQL, objUserInfo

	If Username <> "" Then
		
		strSQL = "SELECT FirstName,LastName FROM People WHERE UserName='" & Username & "'"
		Set objUserInfo = Application("Connection").Execute(strSQL)
		
		If Not objUserInfo.EOF Then
			GetDisplayName = objUserInfo(1) & ", " & objUserInfo(0)
		Else
			GetDisplayName = Username
		End If
	Else
		GetDisplayName = Username
	End If

End Function%>

<%Sub UpdateLog(EntryType,DeviceTag,UserName,OldValue,NewValue,EventNumber)

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
	
End Sub%>

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
		<link rel="stylesheet" type="text/css" href="../style.css" /> 
		<link rel="apple-touch-icon" href="../images/inventory.png" /> 
		<link rel="shortcut icon" href="../images/inventory.ico" />
		<meta name="viewport" content="width=device-width" />
	</head>
	<body>
		<center><b>Access Denied</b></center>
	</body>
	</html>
	
<%	End If

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

	If strRole = "Admin" Then
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
End Function %>

<%Function IsiPad

	Dim strUserAgent

	'Get the User Agent from the client so we know what browser they are using
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Check the user agent for signs they are on a mobile browser
	If InStr(strUserAgent,"iPad") Then
		IsiPad = True
	Else
		IsiPad = False
	End If
	
End Function %>

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
	
	If Application("MultiColumn") Then
		strColumns = "MultiColumn"
	End If
	
End Sub%>