<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/16/14
'Last Updated 1/14/18

'This page shows statistics on the information in the inventory database

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser, objReports, strReport, strSubmitTo, strColumns


'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

	Dim strSQL, strURL
	
	strSQL = "SELECT ID,Report,ShortName FROM Reports ORDER BY Report"
	Set objReports = Application("Connection").Execute(strSQL)

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "View Report"
      	Response.Redirect("stats.asp?Report=" & Replace(Request.Form("Report")," ","%20"))
      Case "Return"
         ReturnMissingItem
      Case "Update"
      	If Request.Form("StartDate") <> "" Then
      		strURL = strURL & "&StartDate=" & Request.Form("StartDate")
      	End If
      	If Request.Form("EndDate") <> "" Then
      		strURL = strURL & "&EndDate=" & Request.Form("EndDate")
      	End If
      	If strURL <> "" Then
      		strURL = "?" & Right(strURL,Len(strURL) - 1)
      		Response.Redirect("stats.asp" & strURL)
      	End If
         
   End Select
   
   'Get the URL used to submit forms
   If Request.ServerVariables("QUERY_STRING") = "" Then
      strSubmitTo = "stats.asp"
   Else   
      strSubmitTo = "stats.asp?" & Request.ServerVariables("QUERY_STRING")
   End If
   
   'Set up the variables needed for the site then load it
   SetupSite
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
		<script src="//www.google.com/jsapi"></script>
		<script>
			$(function() {
			
			<%	If Not IsMobile And Not IsiPad Then%>
					$( document ).tooltip({track: true});
			<% End If %>
				
			 $( "#from" ).datepicker({
				changeMonth: true,
				changeYear: true,
				showOtherMonths: true,
				selectOtherMonths: true,
				onClose: function( selectedDate ) {
				  $( "#to" ).datepicker( "option", "minDate", selectedDate );
				}
			 });
			 $( "#to" ).datepicker({
				changeMonth: true,
				changeYear: true,
				showOtherMonths: true,
				selectOtherMonths: true,
				onClose: function( selectedDate ) {
				  $( "#from" ).datepicker( "option", "maxDate", selectedDate );
				}
			 });
		  });
		  
		  	google.load("visualization", "1", {packages:["corechart"]});
		   <%
		   DeviceYearJavaScript
		   MacBookYearJavaScript
			iPadsYearJavaScript
		  	%>
		  	
  		</script>
   </head>
   <body class="<%=strSiteVersion%>">
   
      <div class="Header"><%=Application("SiteName")%></div>
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
		<div Class="<%=strColumns%>"> 
<% Select Case Request.QueryString("Report")
		Case "MissingAdaptersAndCases"
			MissingAdaptersAndCases
		Case Else
			'ChooseReport
			DatabaseStats
			DeviceSiteStats
			EventTypeStats
			'DeviceAgeStats
			DeviceYearCard
			MacBookYearCard
			iPadYearCard
			GradeLevelStats
			GraduationYearToGradeCard
			PersonTypeStats
			DeviceTypeStats
			EventCategoryStats
			'EventStats
			'EventCategoryStats
			'DevicesPerRoomHS
			'MissingAdaptersAndCases
	End Select %>
		</div> 
		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
   </body>

   </html>

<%End Sub%>

<%Sub ChooseReport %>

	<div class="TwoColumnCard">
		<form method="POST" action="reports.asp">
		<div class="TwoColumnCardTitle">Choose a Report</div>
		<div Class="TwoColumnCardMerged">Report:  
			<select Class="TwoColumnCard" name="Report">
				<option value=""></option>
			<% Do Until objReports.EOF %>
						<option value="<%=objReports(2)%>"><%=objReports(1)%></option>
			<%    objReports.MoveNext
				Loop %>
			</select>
		</div>
		<div>
			<div class="Button"><input type="submit" value="View Report" name="Submit" /></div>
		</div>
	</div>
	
<%End Sub%>

<%Sub DeviceYearCard%>
	<div class="Card NormalCard"> 
		<div class="CardTitle">Devices By Year</div>
		<div id="deviceYear"></div>
	</div>
<%End Sub%>

<%Sub MacBookYearCard%>
	<div class="Card NormalCard"> 
		<div class="CardTitle">MacBooks By Year</div>
		<div id="macBookYear"></div>
	</div>
<%End Sub%>

<%Sub iPadYearCard%>
	<div class="Card NormalCard"> 
		<div class="CardTitle">iPads By Year</div>
		<div id="ipadYear"></div>
	</div>
<%End Sub%>

<%Sub DatabaseStats 

	Dim strSQL, objDeviceCount, intDeviceCount, objPeopleCount, intPeopleCount
	Dim objAssignmentCount, intAssignmentCount

   'Find out how many active devices are in the database.
   strSQL = "SELECT Count(ID) AS CountofID FROM Devices GROUP BY Active HAVING Active=True"
   Set objDeviceCount = Application("Connection").Execute(strSQL) 
   If Not objDeviceCount.EOF Then
   	intDeviceCount = objDeviceCount(0)
   Else 
   	intDeviceCount = 0
   End If
   
   'Find out how many active people are in the database.
   strSQL = "SELECT Count(ID) AS CountofID FROM People GROUP BY Active HAVING Active=True"
   Set objPeopleCount = Application("Connection").Execute(strSQL) 
   If Not objPeopleCount.EOF Then
   	intPeopleCount = objPeopleCount(0)
   Else
   	intPeopleCount = 0
   End If 
   
   'Find out how many active assignments are in the database.
   strSQL = "SELECT Count(ID) AS CountofID FROM Assignments GROUP BY Active HAVING Active=True"
   Set objAssignmentCount = Application("Connection").Execute(strSQL) 
   If Not objAssignmentCount.EOF Then
   	intAssignmentCount = objAssignmentCount(0)
   Else
   	intAssignmentCount = 0
   End If %>

	<div class="Card NormalCard">
		<div class="CardTitle">Database Statistics</div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Type</th>
					<th>Count</th>
				</thead>
				<tbody>
					<tr>
						<td>Total Devices</td>
						<td id="center">
							<a href="devices.asp?DeviceStatus=Enabled&View=Table"><%=intDeviceCount%></a>
						</td>
					</tr>
					<tr>
						<td>Total People</td>
						<td id="center">
							<a href="users.asp?UserStatus=Enabled&View=Table"><%=intPeopleCount%></a>
						</td>
					</tr>
					<tr>
						<td>Assigned Devices</td>
						<td id="center">
							<a href="devices.asp?Assigned=Yes&View=Table"><%=intAssignmentCount%></a>
						</td>
					</tr>
				</tbody>
			</table>
		</div>
	</div>

<%End Sub%>

<%Sub GradeLevelStats

	Dim strSQL, objGradeCounts, objAssignmentCounts, bolGradeFound, intDeviceCount, intDifference
	Dim intTotalUsers, intTotalDevices, intTotalDifference
	
	strSQL = "SELECT ClassOf, Count(ID) AS CountofID FROM People WHERE Active=True GROUP BY ClassOf HAVING ClassOf>2000 ORDER BY ClassOf DESC"
	Set objGradeCounts = Application("Connection").Execute(strSQL) 
	
	strSQL = "SELECT ClassOf, Count(ID) AS CountofID FROM People WHERE HasDevice=True GROUP BY ClassOf HAVING ClassOf>2000 ORDER BY ClassOf DESC"
	Set objAssignmentCounts = Application("Connection").Execute(strSQL)
	
	intTotalUsers = 0
	intTotalDevices = 0
	intTotalDifference = 0 %>

	<div class="Card NormalCard">
		<div class="CardTitle">Grade Level Statistics</div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Grade</th>
					<th>Users</th>
					<th>Devices</th>
					<th>Difference</th>
				</thead>
				<tbody>
			<% If Not objGradeCounts.EOF Then 
					Do Until objGradeCounts.EOF %>
						<tr>
							<td><%=GetRole(objGradeCounts(0))%></td>
							<td id="center">
							<% If objGradeCounts(1) = 0 Then %>
								0
							<% Else 
								intTotalUsers = intTotalUsers + objGradeCounts(1) %>
								<a href="users.asp?Role=<%=objGradeCounts(0)%>&View=Table"><%=objGradeCounts(1)%></a>
							<% End If %>
							</td>
						<% If Not objAssignmentCounts.EOF Then
								bolGradeFound = False
								Do Until objAssignmentCounts.EOF
									
									If Int(objAssignmentCounts(0)) = Int(objGradeCounts(0)) Then
										intDeviceCount = objAssignmentCounts(1)
										bolGradeFound = True
									End If
									objAssignmentCounts.MoveNext	
								Loop
								objAssignmentCounts.MoveFirst
								If Not bolGradeFound Then
									intDeviceCount = 0
								End If
							End If %>
							<td id="center">
							<% If intDeviceCount = 0 Then %>
								0
							<% Else 
								intTotalDevices = intTotalDevices + intDeviceCount %>
								<a href="users.asp?Role=<%=objGradeCounts(0)%>&WithDevice=Yes&View=Table"><%=intDeviceCount%>
							<% End If %>
							</td>
							
							<% intDifference = objGradeCounts(1) - intDeviceCount 
								If intDifference = 0 Then %>
								<td id="center">0</td>
							<%	ElseIf intDifference > 0 Then 
									
								intTotalDifference = intTotalDifference + intDifference %>
								<td id="center"><a href="users.asp?Role=<%=objGradeCounts(0)%>&WithDevice=No&View=Table"><%=intDifference%></a></td>
							<% Else 
								intTotalDifference = intTotalDifference + intDifference %>
								<td id="center" Class="Disabled"><a href="users.asp?Role=<%=objGradeCounts(0)%>&WithDevice=Yes&UserStatus=Disabled&View=Table"><%=intDifference%></a></td>
							<% End If %>
						</tr>
					<%	objGradeCounts.MoveNext
					Loop
				End If %>
					<tr>
						<td>Total</td>
						<td id="center">
							<a href="users.asp?Role=Student&Status=Enabled&View=Table"><%=intTotalUsers%></a>
						</td>
						<td id="center">
							<a href="users.asp?Role=Student&WithDevice=Yes&View=Table"><%=intTotalDevices%></a>
						</td>
						<td id="center">
							<a href="users.asp?Role=Student&WithDevice=No&View=Table"><%=intTotalDifference%></a>
						</td>
					</tr>
				</tbody>
			</table>
		</div>
	</div>

<%End Sub %>

<%Sub DeviceAgeStats

	Dim strSQL, objOldestDevice, datOldestDevice, intYears, intIndex, objDeviceCount, intDeviceCount

	'Get the oldest device from the inventory
   strSQL = "SELECT DatePurchased FROM Devices WHERE DatePurchased Is Not Null And Active=True ORDER BY DatePurchased"
   Set objOldestDevice = Application("Connection").Execute(strSQL)
   If Not objOldestDevice.EOF Then
      datOldestDevice = objOldestDevice(0)
      intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice) 
   End If
%>

	<div class="Card NormalCard">
		<div class="CardTitle">Device Year Statistics</div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Year</th>
					<th>Devices</th>
					<th>MacBooks</th>
					<th>iPads</th>
				</thead>
				<tbody>
			 <% For intIndex = 1 to intYears + 1  %>
					<tr>
						<td>Year <%=intIndex%></td>
					<% intDeviceCount = 0
						strSQL = "SELECT Active, Count(ID) AS CountofID FROM Devices "
						strSQL = strSQL & "WHERE Active=True AND ("
						strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
						strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "	
						strSQL = strSQL & "GROUP BY Active"
						Set objDeviceCount = Application("Connection").Execute(strSQL)
						If Not objDeviceCount.EOF Then
							intDeviceCount = objDeviceCount(1)
						End If
					%>
						<td id="center">
						<% If intDeviceCount = 0 Then %>
							0
						<% Else %>
							<a href="devices.asp?Year=<%=intIndex%>&View=Table"><%=intDeviceCount%></a>
						<% End If %>
						</td>
					<% intDeviceCount = 0
						strSQL = "SELECT Active, Count(ID) AS CountofID FROM Devices "
						strSQL = strSQL & "WHERE Active=True AND Model LIKE '%MacBook%' AND ("
						strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
						strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "	
						strSQL = strSQL & "GROUP BY Active"
						Set objDeviceCount = Application("Connection").Execute(strSQL)
						If Not objDeviceCount.EOF Then
							intDeviceCount = objDeviceCount(1)
						End If
					%>
						<td id="center">
						<% If intDeviceCount = 0 Then %>
							0
						<% Else %>
							<a href="devices.asp?Model=MacBook&Year=<%=intIndex%>&View=Table"><%=intDeviceCount%></a>
						<% End If %>
						</td>
					<% intDeviceCount = 0
						strSQL = "SELECT Active, Count(ID) AS CountofID FROM Devices "
						strSQL = strSQL & "WHERE Active=True AND Model LIKE '%iPad%' AND ("
						strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
						strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "	
						strSQL = strSQL & "GROUP BY Active"
						Set objDeviceCount = Application("Connection").Execute(strSQL)
						If Not objDeviceCount.EOF Then
							intDeviceCount = objDeviceCount(1)
						End If
					%>
						<td id="center">
						<% If intDeviceCount = 0 Then %>
							0
						<% Else %>
							<a href="devices.asp?Model=iPad&Year=<%=intIndex%>&View=Table"><%=intDeviceCount%></a>
						<% End If %>
						</td>
					</tr>
			<% Next %>
				</tbody>
			</table>
		</div>
	</div>

<%End Sub%>

<%Sub DeviceSiteStats

	Dim strSQL, objSiteCount, objDeviceCount, intDeviceCount
	
	strSQL = "SELECT Site, Count(ID) AS CountofID FROM Devices WHERE Active=True GROUP BY Site" 
	Set objSiteCount = Application("Connection").Execute(strSQL)%>
	
	<div class="Card NormalCard">
		<div class="CardTitle">Device Site Statistics</div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Site</th>
					<th>Devices</th>
					<th>MacBooks</th>
					<th>iPads</th>
				</thead>
				<tbody>
			<% If Not objSiteCount.EOF Then 
					Do Until objSiteCount.EOF %>
						<tr>
							<td><%=objSiteCount(0)%></td>
							<td id="center">
								<a href="devices.asp?DeviceSite=<%=objSiteCount(0)%>&View=Table"><%=objSiteCount(1)%></a>
							</td>
							<% intDeviceCount = 0
							strSQL = "SELECT Active, Count(ID) AS CountofID FROM Devices "
							strSQL = strSQL & "WHERE Active=True AND Model LIKE '%MacBook%' AND Site='" & objSiteCount(0) & "'"
							strSQL = strSQL & "GROUP BY Active"
							Set objDeviceCount = Application("Connection").Execute(strSQL)
							If Not objDeviceCount.EOF Then
								intDeviceCount = objDeviceCount(1)
							End If
						%>
							<td id="center">
							<% If intDeviceCount = 0 Then %>
								0
							<% Else %>
								<a href="devices.asp?Model=MacBook&DeviceSite=<%=objSiteCount(0)%>&View=Table"><%=intDeviceCount%></a>
							<% End If %>
							</td>
							<% intDeviceCount = 0
							strSQL = "SELECT Active, Count(ID) AS CountofID FROM Devices "
							strSQL = strSQL & "WHERE Active=True AND Model LIKE '%iPad%' AND Site='" & objSiteCount(0) & "'"
							strSQL = strSQL & "GROUP BY Active"
							Set objDeviceCount = Application("Connection").Execute(strSQL)
							If Not objDeviceCount.EOF Then
								intDeviceCount = objDeviceCount(1)
							End If
						%>
							<td id="center">
							<% If intDeviceCount = 0 Then %>
								0
							<% Else %>
								<a href="devices.asp?Model=iPad&DeviceSite=<%=objSiteCount(0)%>&View=Table"><%=intDeviceCount%></a>
							<% End If %>
							</td>
						</tr>
					<%	objSiteCount.MoveNext
					Loop 
				End If %>
						
				</tbody>
			</table>
		</div>
	</div>

<%End Sub%>

<%Sub DeviceTypeStats
	
	Dim strSQL, objDeviceTypeCount
	strSQL = "SELECT DeviceType, COUNT(DeviceType) FROM Devices WHERE Active = True GROUP BY DeviceType"
	Set objDeviceTypeCount = Application("Connection").Execute(strSQL)%>
	
	<div class="Card NormalCard">
		<div class="CardTitle">Device Type Statistics</div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Type</th>
					<th>Count</th>
				</thead>
				<tbody>
			<% If Not objDeviceTypeCount.EOF Then
					Do Until objDeviceTypeCount.EOF %>
						<tr>
							<td><a href="devices.asp?DeviceType=<%=objDeviceTypeCount(0)%>"><%=objDeviceTypeCount(0)%></a></td>
							<td id="center"><%=objDeviceTypeCount(1)%></td>
						</tr>
				<% objDeviceTypeCount.MoveNext
				Loop
			End If%>
			
				</tbody>
			</table>
		</div>
	</div>

<%End Sub%>


<%Sub PersonTypeStats
	
	Dim strSQL, objPersonTypeCount, objPersonHasDevice, intTotalDevices, intTotalUsers, intTotalDifference
	
	strSQL = "SELECT Roles.Role, Count(*) AS Total, RoleID, SUM (IIF(People.HasDevice = True, 1, 0)) AS WithDevice, (Total-WithDevice) AS Difference "
	strSQL = strSQL & "FROM Roles INNER JOIN People ON Roles.RoleID = People.ClassOf "
	strSQL = strSQL & "WHERE People.Active = True GROUP BY Roles.Role, Roles.RoleID"
	Set objPersonTypeCount = Application("Connection").Execute(strSQL)
	
	intTotalUsers = 0
	intTotalDevices = 0
	intTotalDifference = 0 %>
	
		
	<div class="Card NormalCard">
		<div class="CardTitle">Person Type Statistics</div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Role</th>
					<th>Users</th>
					<th>Devices</th>
					<th>Difference</th>					
				</thead>
				<tbody>
			<% If Not objPersonTypeCount.EOF Then
					Do Until objPersonTypeCount.EOF %>
						<tr>
							<td><%=objPersonTypeCount(0)%></td>
							<td id="center"><a href="users.asp?Role=<%=objPersonTypeCount(2)%>"><%=objPersonTypeCount(1)%></a></td>
							<% intTotalDevices = intTotalDevices + objPersonTypeCount(3) %>
							<% If objPersonTypeCount(3) = 0 Then %>
								<td id="center"><%=objPersonTypeCount(3)%></td>
								<% intTotalUsers = intTotalUsers + objPersonTypeCount(1) %>
							<% Else %>
								<td id="center"><a href="users.asp?Role=<%=objPersonTypeCount(2)%>&WithDevice=Yes"><%=objPersonTypeCount(3)%></td>
								<% intTotalUsers = intTotalUsers + objPersonTypeCount(1) %>
							<% End If %>
							<% If objPersonTypeCount(4) = 0 Then %>
								<td id="center"><%=objPersonTypeCount(4)%></td>
							<% Else  %>
								<td id="center"><a href="users.asp?Role=<%=objPersonTypeCount(2)%>&WithDevice=No"><%=objPersonTypeCount(4)%></td>
							<% End If %>
						</tr>
				<% objPersonTypeCount.MoveNext
				Loop
			End If%>
						<tr>
							<td>Total</td>
							<td><%=intTotalUsers%></td>
							<td><%=intTotalDevices%></td>
							<td><%=intTotalUsers-intTotalDevices%></td>
						</tr>			
				</tbody>
			</table>
		</div>
	</div>

<%End Sub%>



<%Sub EventTypeStats 

Dim strSQL, datStartDate, datEndDate, objEvents, strURL
	
	datStartDate = Request.QueryString("StartDate")
	datEndDate = Request.QueryString("EndDate")
	
	If datStartDate = "" Then
		datStartDate = GetStartOfFiscalYear(Date)
	End If
	
	If datEndDate = "" Then
		datEndDate = Date
	End If

	strURL = strURL & "&StartDate=" & datStartDate
	strURL = strURL & "&EndDate=" & datEndDate
	strURL = strURL & "&Complete=All"
   strURL = "?" & Right(strURL, Len(strURL) - 1)
	
	strSQL = "SELECT Type, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
	strSQL = strSQL & "GROUP BY Type"
	Set objEvents = Application("Connection").Execute(strSQL)%>

	<div class="Card NormalCard">
		<div class="CardTitle">Events by Type <a href="eventstats.asp?Lookup=Types"><image src="../images/dig.png" width="15" height="15" title="Dig Deeper"/></a></div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Event Type</th>
					<th>Count</th>
				</thead>
				<tbody>
			<% If Not objEvents.EOF Then
			 		Do Until objEvents.EOF %>
			 			<tr>
			 				<td>
			 					<%=objEvents(0)%>
			 				</td>
			 				<td id="center">
			 					<a href="events.asp<%=strURL%>&EventType=<%=objEvents(0)%>&View=Table"><%=objEvents(1)%></a>
			 				</td>
			 			</tr>
			 		<%	objEvents.MoveNext
			 		Loop
			 	End If %>
				</tbody>
			</table>
		</div>
	</div>

<%End Sub%>

<%Sub EventCategoryStats 
	
	Dim strSQL, datStartDate, datEndDate, objEvents, strURL
	
	datStartDate = Request.QueryString("StartDate")
	datEndDate = Request.QueryString("EndDate")
	
	If datStartDate = "" Then
		datStartDate = GetStartOfFiscalYear(Date)
	End If
	
	If datEndDate = "" Then
		datEndDate = Date
	End If

	strURL = strURL & "&StartDate=" & datStartDate
	strURL = strURL & "&EndDate=" & datEndDate
	strURL = strURL & "&Complete=All"
   strURL = "?" & Right(strURL, Len(strURL) - 1)
	
	strSQL = "SELECT Category, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
	strSQL = strSQL & "GROUP BY Category"
	Set objEvents = Application("Connection").Execute(strSQL)%>
	
	<div class="Card NormalCard">
		<div class="CardTitle">Events by Category <a href="eventstats.asp?Lookup=Categories"><image src="../images/dig.png" width="15" height="15" title="Dig Deeper"/></a></div>	
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Category</th>
					<th>Count</th>
				</thead>
				<tbody>
			<% If Not objEvents.EOF Then
			 		Do Until objEvents.EOF %>
			 			<tr>
			 				<td>
			 					<%=objEvents(0)%>
			 				</td>
			 				<td id="center">
			 					<a href="events.asp<%=strURL%>&Category=<%=objEvents(0)%>&View=Table"><%=objEvents(1)%></a>
			 				</td>
			 			</tr>
			 		<%	objEvents.MoveNext
			 		Loop
			 	End If %>
				</tbody>
			</table>
			<form method="POST" action="stats.asp">
			<div>&nbsp;</div>
			&nbsp;&nbsp;&nbsp;&nbsp;Range:
			<input class="SingleColumnCard InputWidthSmall" type="text" name="StartDate" value="<%=datStartDate%>" id="from"> - 
			<input class="SingleColumnCard InputWidthSmall" type="text" name="EndDate" value="<%=datEndDate%>" id="to">
			<div class="Button"><input type="submit" value="Update" name="Submit" /></div>
			</form>
		</div>
		
	</div>

<%End Sub %>

<%Sub EventStats

	Dim strSQL, datStartDate, datEndDate, objEvents, strURL
	
	datStartDate = Request.QueryString("StartDate")
	datEndDate = Request.QueryString("EndDate")
	
	If datStartDate = "" Then
		datStartDate = GetStartOfFiscalYear(Date)
	End If
	
	If datEndDate = "" Then
		datEndDate = Date
	End If

	strURL = strURL & "&StartDate=" & datStartDate
	strURL = strURL & "&EndDate=" & datEndDate
	strURL = strURL & "&Complete=All"
   strURL = "?" & Right(strURL, Len(strURL) - 1)
	
	strSQL = "SELECT Type, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
	strSQL = strSQL & "GROUP BY Type"
	Set objEvents = Application("Connection").Execute(strSQL)%>

	<div class="Card NormalCard">
		<div class="CardTitle">Events</div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Event Type</th>
					<th>Count</th>
				</thead>
				<tbody>
			<% If Not objEvents.EOF Then
			 		Do Until objEvents.EOF %>
			 			<tr>
			 				<td>
			 					<%=objEvents(0)%>
			 				</td>
			 				<td id="center">
			 					<a href="events.asp<%=strURL%>&EventType=<%=objEvents(0)%>&View=Table"><%=objEvents(1)%></a>
			 				</td>
			 			</tr>
			 		<%	objEvents.MoveNext
			 		Loop
			 	End If %>
				</tbody>
			</table>
		</div>
		
<%	strSQL = "SELECT Category, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
	strSQL = strSQL & "GROUP BY Category"
	Set objEvents = Application("Connection").Execute(strSQL)%>

		
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Category</th>
					<th>Count</th>
				</thead>
				<tbody>
			<% If Not objEvents.EOF Then
			 		Do Until objEvents.EOF %>
			 			<tr>
			 				<td>
			 					<%=objEvents(0)%>
			 				</td>
			 				<td id="center">
			 					<a href="events.asp<%=strURL%>&Category=<%=objEvents(0)%>&View=Table"><%=objEvents(1)%></a>
			 				</td>
			 			</tr>
			 		<%	objEvents.MoveNext
			 		Loop
			 	End If %>
				</tbody>
			</table>
			<form method="POST" action="stats.asp">
			<div>&nbsp;</div>
			&nbsp;&nbsp;&nbsp;&nbsp;Range:
			<input class="SingleColumnCard InputWidthSmall" type="text" name="StartDate" value="<%=datStartDate%>" id="from"> - 
			<input class="SingleColumnCard InputWidthSmall" type="text" name="EndDate" value="<%=datEndDate%>" id="to">
			<div class="Button"><input type="submit" value="Update" name="Submit" /></div>
			</form>
		</div>
		
	</div>


<%End Sub %>

<%Sub GraduationYearToGradeCard

	Dim intIndex, strMessage
	
	strMessage = "<table align=""center"" Class=""ListView"">" & vbCRLF

	For intIndex = 0 to 12
		strMessage = strMessage & "<tr>"
		If intIndex = 0 Then
			strMessage = strMessage & "<td id=""center"">K</td><td id=""center"">" & GetGraduationYear(intIndex) & "</td></tr>"
		Else
			strMessage = strMessage & "<td id=""center"">" & intIndex & "</td><td id=""center"">" & GetGraduationYear(intIndex) & "</td></tr>"
		End If
	Next
	strMessage = strMessage & vbCRLF & "</table>" %>
	
	<div class="Card NormalCard"> 
		<div class="CardTitle">Grade to Graduation Year</div>
		<div><%=strMessage%></div>
	</div>
<%End Sub%>

<%Sub DevicesPerRoomHS

	Dim strSQL, objRooms

   strSQL = "SELECT Room, Count(ID) AS CountofID FROM Devices GROUP BY Room, Site HAVING Site='High School'"
   Set objRooms = Application("Connection").Execute(strSQL) %>

	<div class="Card NormalCard">
		<div class="CardTitle">High School - Device Per Room</div>
		<div>
			<table align="center" Class="ListView">
				<thead>
					<th>Room</th>
					<th>Devices</th>
				</thead>
				<tbody>
			<% If Not objRooms.EOF Then
			 		Do Until objRooms.EOF %>
			 			<tr>
			 				<td>
			 					<%=objRooms(0)%>
			 				</td>
			 				<td id="center">
			 					<%=objRooms(1)%>
			 				</td>
			 			</tr>
			 		<%	objRooms.MoveNext
			 		Loop
			 	End If %>
				</tbody>
			</table>
		</div>
	</div>

<%End Sub%>

<%Sub MissingAdaptersAndCases
	
	Dim strSQL, objReportData, objFSO, intItemCount
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	strSQL = "SELECT LGTag,UserName,FirstName,LastName,AdapterReturned,CaseReturned,StudentID,Role,Assignments.ID,People.Active,Warning" & vbCRLF
	strSQL = strSQL & "FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
	strSQL = strSQL & "WHERE (Assignments.Active=False) AND (Assignments.AdapterReturned=False OR Assignments.CaseReturned=False)"
	Set objReportData = Application("Connection").Execute(strSQL) 
	
	intItemCount = 0
	Do Until objReportData.EOF
		intItemCount = intItemCount + 1
		objReportData.MoveNext
	Loop
	objReportData.MoveFirst
	%>
	
	<div align="center">Missing Adapters and Cases (<%=intItemCount%> People Found)</div>
	
	<% Do Until objReportData.EOF %>
		
		<% If objReportData(10) Then%>
				<div class="Card WarningCard">
		<% ElseIf objReportData(9) Then %>
				<div class="Card PhotoCard">
		<% Else %>
				<div class="Card DisabledCard">
		<% End If %>	

				<div>
					<a href="user.asp?UserName=<%=objReportData(1)%>">
			<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objReportData(7) & "s\" & objReportData(6) & ".jpg") Then %>      
					<img class="PhotoCard" src="/photos/<%=objReportData(7)%>s/<%=objReportData(6)%>.jpg" width="96" />
			<% Else %>
					<img class="PhotoCard" src="/photos/<%=objReportData(7)%>s/missing.jpg" width="96" />
			<% End If %>
					</a>
				</div>
				<div class="PhotoCardTitle"><%=objReportData(2) & " " & objReportData(3)%></div>
			
			<% If Not objReportData(4) Then %>
					<form method="POST" action="<%=strSubmitTo%>">   
					<input type="hidden" name="AssignmentID" value="<%=objReportData(8)%>" />
					<div>
						<a href="device.asp?Tag=<%=objReportData(0)%>"><%=objReportData(0)%></a>
						 - Adapter: <input type="checkbox" name="Adapter" value="True" />
						<div class="Button"><input type="submit" value="Return" name="Submit" /></div>
					</div>
					</form>
				<% If Not objReportData(5) Then %>
						<div>&nbsp;</div>
				<% End If %>
			<% End If %>
		
			<% If Not objReportData(5) Then %>
					<form method="POST" action="<%=strSubmitTo%>">   
					<input type="hidden" name="AssignmentID" value="<%=objReportData(8)%>" />
					<div>
						<a href="device.asp?Tag=<%=objReportData(0)%>"><%=objReportData(0)%></a>
						 - Case: <input type="checkbox" name="Case" value="True" />
						<div class="Button"><input type="submit" value="Return" name="Submit" /></div>
					</div>
					</form>
			<% End If %>

			</div>
		<% objReportData.MoveNext 
		Loop %>
	
<% End Sub %>

<%Sub DeviceYearJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intYears, intIndex, objDeviceCount, intDeviceCount, strDeviceData

	'Get the oldest device from the inventory
   strSQL = "SELECT DatePurchased FROM Devices WHERE DatePurchased Is Not Null AND Active=True ORDER BY DatePurchased"
   Set objOldestDevice = Application("Connection").Execute(strSQL)
   If Not objOldestDevice.EOF Then
      datOldestDevice = objOldestDevice(0)
      intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice) 
   End If

	strDeviceData = "['Year','Count',{ role: 'annotation' } ],"

	For intIndex = 1 to intYears + 1

		intDeviceCount = 0
		
		strSQL = "SELECT Active, Count(ID) AS CountofID FROM Devices "
		strSQL = strSQL & "WHERE Active=True AND ("
		strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
		strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "	
		strSQL = strSQL & "GROUP BY Active"
		Set objDeviceCount = Application("Connection").Execute(strSQL)

		If Not objDeviceCount.EOF Then
			intDeviceCount = objDeviceCount(1)
		End If
		
		strDeviceData = strDeviceData & "['Year " & intIndex & "', " & intDeviceCount & ",'" & intIndex & "'],"
		
	Next%>
	
	google.setOnLoadCallback(drawDeviceYear);
	
	function drawDeviceYear() {
		
		var data = google.visualization.arrayToDataTable([
			<%=strDeviceData%>
		]);
		
		var options = {
			titlePosition: 'none',
			chartArea:{width:'90%', height:'85%'},
			animation: {startup: 'true', duration: 1000, easing: 'out'},
			is3D: 'true',
			pieSliceText: 'value',
			hAxis: {title: '', minValue: 0},
			vAxis: {title: ''}
		};
	
		var chart = new google.visualization.PieChart(document.getElementById('deviceYear'));
		chart.draw(data, options);
	
		google.visualization.events.addListener(chart, 'select', selectHandler);
		
		function selectHandler(e) {
			window.open('devices.asp?Year=' + data.getValue(chart.getSelection()[0].row, 2) + '&View=Table','_self');
		}
	
	}
      
<%End Sub %>

<%Sub MacBookYearJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intYears, intIndex, objDeviceCount, intDeviceCount, strMacBookData

	'Get the oldest device from the inventory
   strSQL = "SELECT DatePurchased FROM Devices WHERE DatePurchased Is Not Null AND Active=True AND Model LIKE '%MacBook%' ORDER BY DatePurchased"
   Set objOldestDevice = Application("Connection").Execute(strSQL)
   If Not objOldestDevice.EOF Then
      datOldestDevice = objOldestDevice(0)
      intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice) 
   End If

	strMacBookData = "['Year','Count',{ role: 'annotation' } ],"

	For intIndex = 1 to intYears + 1

		intDeviceCount = 0
		
		strSQL = "SELECT Active, Count(ID) AS CountofID FROM Devices "
		strSQL = strSQL & "WHERE Active=True AND Model LIKE '%MacBook%' AND ("
		strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
		strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "	
		strSQL = strSQL & "GROUP BY Active"
		Set objDeviceCount = Application("Connection").Execute(strSQL)

		If Not objDeviceCount.EOF Then
			intDeviceCount = objDeviceCount(1)
		End If
		
		strMacBookData = strMacBookData & "['Year " & intIndex & "', " & intDeviceCount & ",'" & intIndex & "'],"
		
	Next%>
	
	google.setOnLoadCallback(drawMacBookYear);
	
	function drawMacBookYear() {
		
		var data = google.visualization.arrayToDataTable([
			<%=strMacBookData%>
		]);
		
		var options = {
			titlePosition: 'none',
			chartArea:{width:'90%', height:'85%'},
			animation: {startup: 'true', duration: 1000, easing: 'out'},
			is3D: 'true',
			pieSliceText: 'value',
			hAxis: {title: '', minValue: 0},
			vAxis: {title: ''}
		};
	
		var chart = new google.visualization.PieChart(document.getElementById('macBookYear'));
		chart.draw(data, options);
	
		google.visualization.events.addListener(chart, 'select', selectHandler);
		
		function selectHandler(e) {
			window.open('devices.asp?Year=' + data.getValue(chart.getSelection()[0].row, 2) + '&Model=MacBook&View=Table','_self');
		}
	
	}
      
<%End Sub %>

<%Sub iPadsYearJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intYears, intIndex, objDeviceCount, intDeviceCount, striPadData

	'Get the oldest device from the inventory
   strSQL = "SELECT DatePurchased FROM Devices WHERE DatePurchased Is Not Null AND Active=True AND Model LIKE '%iPad%' ORDER BY DatePurchased"
   Set objOldestDevice = Application("Connection").Execute(strSQL)
   If Not objOldestDevice.EOF Then
      datOldestDevice = objOldestDevice(0)
      intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice) 
   End If

	striPadData = "['Year','Count',{ role: 'annotation' } ],"

	For intIndex = 1 to intYears + 1

		intDeviceCount = 0
		
		strSQL = "SELECT Active, Count(ID) AS CountofID FROM Devices "
		strSQL = strSQL & "WHERE Active=True AND Model LIKE '%iPad%' AND ("
		strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
		strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "	
		strSQL = strSQL & "GROUP BY Active"
		Set objDeviceCount = Application("Connection").Execute(strSQL)

		If Not objDeviceCount.EOF Then
			intDeviceCount = objDeviceCount(1)
		End If
		
		striPadData = striPadData & "['Year " & intIndex & "', " & intDeviceCount & ",'" & intIndex & "'],"
		
	Next%>
	
	google.setOnLoadCallback(drawiPadYear);
	
	function drawiPadYear() {
		
		var data = google.visualization.arrayToDataTable([
			<%=striPadData%>
		]);
		
		var options = {
			titlePosition: 'none',
			chartArea:{width:'90%', height:'85%'},
			animation: {startup: 'true', duration: 1000, easing: 'out'},
			is3D: 'true',
			pieSliceText: 'value',
			hAxis: {title: '', minValue: 0},
			vAxis: {title: ''}
		};
	
		var chart = new google.visualization.PieChart(document.getElementById('ipadYear'));
		chart.draw(data, options);
	
		google.visualization.events.addListener(chart, 'select', selectHandler);
		
		function selectHandler(e) {
			window.open('devices.asp?Year=' + data.getValue(chart.getSelection()[0].row, 2) + '&Model=iPad&View=Table','_self');
		}
	
	}
      
<%End Sub %>

<%Sub ReturnMissingItem

	Dim intAssignmentID, bolAdapterReturned, bolCaseReturned, strSQL, strUserName
	Dim objUser, objItemCheck

	intAssignmentID = Request.Form("AssignmentID")
	bolAdapterReturned = Request.Form("Adapter")
	bolCaseReturned = Request.Form("Case")
	
	strSQL = "SELECT UserName FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
	strSQL = strSQL & "WHERE Assignments.ID=" & intAssignmentID
	Set objUser = Application("Connection").Execute(strSQL)
	strUserName = objUser(0)
	
	
	If bolAdapterReturned Then
		strSQL = "UPDATE Assignments SET AdapterReturned=True WHERE ID=" & intAssignmentID
		Application("Connection").Execute(strSQL)
	End If
	
	If bolCaseReturned Then
		strSQL = "UPDATE Assignments SET CaseReturned=True WHERE ID=" & intAssignmentID
		Application("Connection").Execute(strSQL)
	End If
	
	strSQL = "SELECT Assignments.ID FROM People INNER JOIN Assignments ON People.ID = Assignments.AssignedTo" & vbCRLF
	strSQL = strSQL & "WHERE (UserName='" & strUserName & "' AND Assignments.Active=False) AND "
	strSQL = strSQL & "(AdapterReturned=False OR CaseReturned=False)"
	Set objItemCheck = Application("Connection").Execute(strSQL)
	
	If objItemCheck.EOF Then
		strSQL = "UPDATE People SET Warning=False WHERE UserName='" & strUserName & "'"
		Application("Connection").Execute(strSQL)
	End If
	
End Sub%>

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

<%Function GetGraduationYear(intGrade)

	Dim datToday, intMonth, intCurrentYear

   datToday = Date
   intMonth = DatePart("m",datToday)
   intCurrentYear = DatePart("yyyy",datToday)
   
   If intMonth >= 7 And intMonth <= 12 Then
      intCurrentYear = intCurrentYear + 1
   End If
   
   GetGraduationYear = intCurrentyear + (12 - intGrade)
   
End Function %>

<%Function GetRole(intYear)

   Dim datToday, intMonth, intCurrentYear, intGrade, strSQL, objRole
   
   'If they're an adult then get their role from the database
   If intYear <= 1000 Then
   	strSQL = "SELECT Role FROM Roles WHERE RoleID=" & intYear
   	Set objRole = Application("Connection").Execute(strSQL)
   	
   	If Not objRole.EOF Then
   		GetRole = objRole(0)
   	End If
   End If 
      
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
				GetRole = "1st Grade"
			Case 2
				GetRole = "2nd Grade"
			Case 3
				GetRole = "3rd Grade"
			Case 4
				GetRole = "4th Grade"   
			Case 5
				GetRole = "5th Grade"
			Case 6
				GetRole = "6th Grade"
			Case 7
				GetRole = "7th Grade"
			Case 8
				GetRole = "8th Grade"
			Case 9
				GetRole = "9th Grade"
			Case 10
				GetRole = "10th Grade"
			Case 11
				GetRole = "11th Grade"
			Case 12
				GetRole = "12th Grade"
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
      <link rel="stylesheet" type="text/css" href="../style.css" /> 
      <link rel="apple-touch-icon" href="../images/inventory.png" /> 
      <link rel="shortcut icon" href="../images/inventory.ico" />
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