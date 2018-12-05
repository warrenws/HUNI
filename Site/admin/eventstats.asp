<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/3/16
'Last Updated 1/14/18

'This page shows stats about the events

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim intTag, strFirstName, strLastName, strRole, strDeviceMessage, strUserMessage, strEventMessage
Dim strColumns, objEventData, intDisabledUsersCount, intEventNumber, intYears, datStartDate, datEndDate
Dim strSubmitTo, strCardList

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

   Dim strSQL, objOldestDevice, datOldestDevice, strURL

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
		Case "Update"
			strURL = BuildDateURL
			If Request.QueryString("Model") <> "" Then
				strURL = strURL & "&Model=" & Request.QueryString("Model")
			End If
			If Request.QueryString("LookUp") <> "" Then
				strURL = strURL & "&LookUp=" & Request.QueryString("LookUp")
			End If
			Response.Redirect("eventstats.asp" & strURL)
   End Select
   
   If Request.QueryString("StartDate") = "" Then
   	datStartDate = GetStartOfFiscalYear(Date)
   Else
   	datStartDate = Request.QueryString("StartDate")
   End If
   
   If Request.QueryString("EndDate") = "" Then
		datEndDate = Date
	Else
		datEndDate = Request.QueryString("EndDate")
	End If
   
   'Get the list of models with events from the database
   If Request.QueryString("Model") = "" Then
   	strSQL = "SELECT DISTINCT Model FROM Events WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "# ORDER BY Model" 
   Else   	
   	
		'Get the oldest device
		intYears = 0
		strSQL = "SELECT DatePurchased" & vbCRLF
		strSQL = strSQL & "FROM Devices INNER JOIN Events ON Devices.LGTag = Events.LGTag" & vbCRLF
		strSQL = strSQL & "WHERE DatePurchased Is Not Null AND Devices.Model='" & Replace(Request.QueryString("Model"),"'","''") & "' AND EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "# AND Category <> 'End Of Life'" & vbCRLF
		strSQL = strSQL & "ORDER BY DatePurchased"
		Set objOldestDevice = Application("Connection").Execute(strSQL)
		If Not objOldestDevice.EOF Then
			datOldestDevice = objOldestDevice(0)
			intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice) 
		End If
   	
   End If
   Set objEventData = Application("Connection").Execute(strSQL)
   
   'Display the error message if they were sent back to this page
   Select Case Request.QueryString("Error")
      Case "DeviceNotFound"
         strDeviceMessage = "<div Class=""Error"">Device not found</div>"
      Case "NoDevicesFound"
         strDeviceMessage = "<div Class=""Error"">No devices found</div>"
      Case "UserNotFound"
         strUserMessage = "<div Class=""Error"">User not found</div>"
      Case "NoUsersFound"
         strUserMessage = "<div Class=""Error"">No users found</div>"
      Case "NoEventsFound"
      	strEventMessage = "<div Class=""Error"">No events found</div>"
   End Select
   
   'Get the URL used to submit forms
   If Request.ServerVariables("QUERY_STRING") = "" Then
      strSubmitTo = "eventstats.asp"
   Else   
      strSubmitTo = "eventstats.asp?" & Request.ServerVariables("QUERY_STRING")
   End If
   
   'Set up the variables needed for the site then load it
   SetupSite
   DisplaySite
   
End Sub%>

<%Sub DisplaySite

	Dim strSQL, arrCards, strCard %>

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
		<script src="../assets/js/jquery.flip.min.js"></script>
		<script src="//www.google.com/jsapi"></script>
		<script>
		
		  	google.load("visualization", "1", {packages:["corechart"]});
			
			$(document).ready( function () {
				$(".EventStats").flip({
					axis: 'y',
					trigger: 'manual'
				});
							
			<%
			If Request.QueryString("Model") = "" Then
				EventCategoriesJavaScript
				EventTypesJavaScript
				CategoriesListJavaScript
				TypesListJavaScript
			Else
				ModelCategoriesJavaScript
				ModelTypesJavaScript
				CategoriesByYearJavaScript
				TypesByYearJavaScript
			End If%>
			
				var viewCategoriesButton = document.getElementById("ViewCategories");
				var viewTypesButton = document.getElementById("ViewTypes");
				
				viewCategoriesButton.onclick = function() {
					$(".EventStats").flip('toggle');
					return false;
				}
				
				viewTypesButton.onclick = function() {
					$(".EventStats").flip('toggle'); 	 
					return false;
				}
		  	
		  		$(".FlipCard").on('click', function(e) {
					$(e.target).closest(".EventStats").flip('toggle');
					return false;
				});
		  	
		  	});
			
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
		
		<%
		Select Case Request.QueryString("LookUp")
			Case "Type", "Types"
				EventTypesPage
			Case "Category", "Categories"
				EventCategoriesPage
			Case Else
				EventTypesPage
		End Select
		%>

      </div>
      <div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
   </body>
   </html>

<%End Sub%>

<%Sub EventTypesPage

	DatePickerCard
	If Request.QueryString("Model") = "" Then %>
	
		<div class="flip EventStats" > 
			<div class="front"> 
				<%EventTypesCard%>
			</div>
			<div class="back">
				<%EventCategoriesCard%>
			</div>
		</div>
	
	<%	
		EventListCards

	Else %>
	
		<div class="flip EventStats" > 
			<div class="front"> 
				<%ModelTypesCard%>
			</div>
			<div class="back">
				<%ModelCategoriesCard%>
			</div>
		</div>
	<%	
		EventsListByYearCards

	End If

End Sub%>

<%Sub EventCategoriesPage

	DatePickerCard
	If Request.QueryString("Model") = "" Then %>
		
		<div class="flip EventStats" > 
			<div class="front"> 
				<%EventCategoriesCard%>
			</div>
			<div class="back">
				<%EventTypesCard%>
			</div>
		</div>
		
	<%	
		EventListCards

	Else %>
		
		<div class="flip EventStats" > 
			<div class="front"> 
				<%ModelCategoriesCard%>
			</div>
			<div class="back">
				<%ModelTypesCard%>
			</div>
		</div>
	<%	
		EventsListByYearCards

	End If

End Sub %>

<%Sub DatePickerCard%>
	<div Class="HeaderCard">
		<form method="POST" action="<%=strSubmitTo%>">
		&nbsp;&nbsp;&nbsp;&nbsp;<input class="SingleColumnCard InputWidthDate" type="text" name="StartDate" value="<%=datStartDate%>" id="from"> - 
		<input class="SingleColumnCard InputWidthDate" type="text" name="EndDate" value="<%=datEndDate%>" id="to">
		<div class="Button"><input type="submit" value="Update" name="Submit" /></div>
		</form>
	</div>
<%End Sub%>

<%Sub EventCategoriesCard%>
	<div class="Card NormalCard"> 
		<div class="CardTitle">Event Categories</div>
		<div id="eventCategories" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="index.asp"><image src="../images/back.png" width="20" height="20" title="Return to the Main Page"/></a>
				<a href="" id="ViewTypes"><image src="../images/swap.png" width="20" height="20" title="View Types"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub EventTypesCard%>
	<div class="Card NormalCard"> 
		<div class="CardTitle">Event Types</div>
		<div id="eventTypes" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="index.asp"><image src="../images/back.png" width="20" height="20" title="Return to the Main Page"/></a>
				<a href="" id="ViewCategories"><image src="../images/swap.png" width="20" height="20" title="View Categories"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub ModelCategoriesCard%>
	<div class="Card NormalCard"> 
		<div class="CardTitle"><%=Request.QueryString("Model")%> Events</div>
		<div id="modelCategories" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="eventstats.asp<%=FixReturnLink("Categories")%>"><image src="../images/back.png" width="20" height="20" title="Return to Events"/></a>
				<a href="" id="ViewTypes"><image src="../images/swap.png" width="20" height="20" title="View Types"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub ModelTypesCard%>
	<div class="Card NormalCard"> 
		<div class="CardTitle"><%=Request.QueryString("Model")%> Events</div>
		<div id="modelTypes" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="eventstats.asp<%=FixReturnLink("Types")%>"><image src="../images/back.png" width="20" height="20" title="Return to Events"/></a>
				<a href="" id="ViewCategories"><image src="../images/swap.png" width="20" height="20" title="View Categories"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Function FixReturnLink(strLinkType)

	FixReturnLink = BuildDateURL
	If InStr(FixReturnLink,"LookUp=") Then
		Select Case strLinkType
			Case "Types"
				FixReturnLink = Replace(FixReturnLink,"LookUp=Categories","Lookup=Types")
			Case "Categories"
				FixReturnLink = Replace(FixReturnLink,"LookUp=Types","Lookup=Categories")
		End Select
	Else
		Select Case strLinkType
			Case "Types"
				FixReturnLink = FixReturnLink & "&LookUp=Types"
			Case "Categories"
				FixReturnLink = FixReturnLink & "&LookUp=Categories"
		End Select
	End If
	
	If Right(FixReturnLink,1) <> "?" Then
		FixReturnLink = "?" & Right(FixReturnLink,Len(FixReturnLink) - 1)
	End If

End Function %>

<%Sub EventListCards

	Dim strSQL, strCardTitle, objOldestDevice, datOldestDevice, intIndex, objEvents, intYearsWithEvents, strCardName, strDigLink

	If Not objEventData.EOF Then
	
		Do Until objEventData.EOF 
		
			'Get the oldest device
			intYears = 0
			intYearsWithEvents = 0
			datOldestDevice = ""
			strSQL = "SELECT DatePurchased" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Events ON Devices.LGTag = Events.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE DatePurchased Is Not Null AND Devices.Model='" & Replace(objEventData(0),"'","''") & "' AND EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
			strSQL = strSQL & "ORDER BY DatePurchased"
			Set objOldestDevice = Application("Connection").Execute(strSQL)
			If Not objOldestDevice.EOF Then
				datOldestDevice = objOldestDevice(0)
				intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice) 
			End If 
		
			For intIndex = 1 to intYears + 1
		
				strSQL = "SELECT Type, Count(Events.ID) AS CountOfID" & vbCRLF
				strSQL = strSQL & "FROM Devices INNER JOIN Events ON Devices.LGTag = Events.LGTag" & vbCRLF
				strSQL = strSQL & "WHERE Events.Model='" & Replace(objEventData(0),"'","''") & "' AND "
				strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
				strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "# AND "
				strSQL = strSQL & "Events.EventDate>=#" & datStartDate & "# AND "
				strSQL = strSQL & "Events.EventDate<=#" & datEndDate & "#" & vbCRLF
				strSQL = strSQL & "GROUP BY Type" & vbCRLF
				Set objEvents = Application("Connection").Execute(strSQL)
			
				If Not objEvents.EOF Then
					intYearsWithEvents = intYearsWithEvents + 1
				End If
		
			Next

			If intYearsWithEvents > 1 Then 
				If Request.ServerVariables("QUERY_STRING") = "" Then
					strDigLink = "<a href=""eventstats.asp?Model=" & objEventData(0) & """><image src=""../images/dig.png"" width=""20"" height=""20"" title=""Dig Deeper""/></a>"
				Else
					strDigLink = "<a href=""eventstats.asp?Model=" & objEventData(0) & "&" & Request.ServerVariables("QUERY_STRING") & """><image src=""../images/dig.png"" width=""20"" height=""20"" title=""Dig Deeper""/></a>"
				End If
			Else 
				strDigLink = ""
			End If

			strCardTitle = objEventData(0) & " Events"
			strCardName = objEventData(0)
			strCardName = Replace(strCardName," ","")
			strCardName = Replace(strCardName,".","")
			strCardName = Replace(strCardName,"'","")
			strCardName = Replace(strCardName,"-","")	

			Select Case Request.QueryString("LookUp")
				Case "Type", "Types"
					ListTypeCards strCardName,strCardTitle,strDigLink
				Case "Category", "Categories"
					ListCategoryCards strCardName,strCardTitle,strDigLink
				Case Else
					ListTypeCards strCardName,strCardTitle,strDigLink
			End Select
		
			objEventData.MoveNext
		Loop
		objEventData.MoveFirst
   End If

End Sub%>

<%Sub ListTypeCards(strCardName,strCardTitle,strDigLink) %>

	<div class="flip EventStats" >
		<div class="front"> 
			<div class="Card NormalCard"> 
				<div class="CardTitle"><%=strCardTitle%></div>
				<div id="Types<%=strCardName%>" Class="Chart"></div>
				<div class="ChartBottomBar">
					<div class="ChartBackLink">
						<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Categories"/></a>
						<%=Replace(strDigLink,"LookUp=Categories","LookUp=Types")%>
					</div>
				</div>
			</div>
		</div>
		<div class="back">
			<div class="Card NormalCard"> 
				<div class="CardTitle"><%=strCardTitle%></div>
				<div id="Categories<%=strCardName%>" Class="Chart"></div>
				<div class="ChartBottomBar">
					<div class="ChartBackLink">
						<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Types"/></a>
						<%=Replace(strDigLink,"LookUp=Types","LookUp=Categories")%>
					</div>
				</div>
			</div>
		</div>
	</div>

<%End Sub%>

<%Sub ListCategoryCards(strCardName,strCardTitle,strDigLink) %>

	<div class="flip EventStats" >
		<div class="front"> 
			<div class="Card NormalCard"> 
				<div class="CardTitle"><%=strCardTitle%></div>
				<div id="Categories<%=strCardName%>" Class="Chart"></div>
				<div class="ChartBottomBar">
					<div class="ChartBackLink">
						<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Categories"/></a>
						<%=Replace(strDigLink,"LookUp=Types","LookUp=Categories")%>
					</div>
				</div>
			</div>
		</div>
		<div class="back">
			<div class="Card NormalCard"> 
				<div class="CardTitle"><%=strCardTitle%></div>
				<div id="Types<%=strCardName%>" Class="Chart"></div>
				<div class="ChartBottomBar">
					<div class="ChartBackLink">
						<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Types"/></a>
						<%=Replace(strDigLink,"LookUp=Categories","LookUp=Types")%>
					</div>
				</div>
			</div>
		</div>
	</div>

<%End Sub%>

<%Sub EventsListByYearCards

	Dim strSQL, intIndex, objEvents

	If intYears > 0 Then
		For intIndex = 1 to intYears + 1
		
			strSQL = "SELECT Type, Count(Events.ID) AS CountOfID" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Events ON Devices.LGTag = Events.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Events.Model='" & Replace(Request.QueryString("Model"),"'","''") & "' AND "
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "# AND "
			strSQL = strSQL & "Events.EventDate>=#" & datStartDate & "# AND "
			strSQL = strSQL & "Events.EventDate<=#" & datEndDate & "#" & vbCRLF
			strSQL = strSQL & "GROUP BY Type" & vbCRLF
			Set objEvents = Application("Connection").Execute(strSQL)
			
			If Not objEvents.EOF Then %>
			
			<div class="flip EventStats" >
				<div class="front"> 
					<div class="Card NormalCard"> 
						<div class="CardTitle"><%=Request.QueryString("Model")%> Year <%=intIndex%></div>
						<div id="TypesYear<%=intIndex%>" Class="Chart"></div>
						<div class="ChartBottomBar">
							<div class="ChartBackLink">
								<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Categories"/></a>
							</div>
						</div>
					</div>
				</div>
				<div class="back">
					<div class="Card NormalCard"> 
						<div class="CardTitle"><%=Request.QueryString("Model")%> Year <%=intIndex%></div>
						<div id="CategoriesYear<%=intIndex%>" Class="Chart"></div>
						<div class="ChartBottomBar">
							<div class="ChartBackLink">
								<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Types"/></a>
							</div>
						</div>
					</div>
				</div>
			</div>
			
		<% End If
		 Next
   End If

End Sub%>

<%Sub EventCategoriesJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intIndex, objDeviceCount, intDeviceCount
	Dim objEvents, strEventData

	strSQL = "SELECT Category, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
	strSQL = strSQL & "GROUP BY Category" & vbCRLF
	Set objEvents = Application("Connection").Execute(strSQL)
	
	strEventData = "['Category','Count',{ role: 'annotation' }, { role: 'annotation' } ],"
	
	If Not objEvents.EOF Then
		Do Until objEvents.EOF
			strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "','" & datEndDate & "'],"
			objEvents.MoveNext
		Loop
	End If%>
	
	google.setOnLoadCallback(drawEventCategories);
	
	function drawEventCategories() {
		
		var data = google.visualization.arrayToDataTable([
			<%=strEventData%>
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
	
		var chart = new google.visualization.PieChart(document.getElementById('eventCategories'));
		chart.draw(data, options);
	
		google.visualization.events.addListener(chart, 'select', selectHandler);
		
		function selectHandler(e) {
			var startDate = data.getValue(chart.getSelection()[0].row, 2)
			var endDate = data.getValue(chart.getSelection()[0].row, 3)
			window.open('events.asp?Category=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + startDate + '&EndDate=' + endDate + '&Complete=All&View=Table','_self');
		}
	
	}
      
<%End Sub %>

<%Sub EventTypesJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intIndex, objDeviceCount, intDeviceCount
	Dim objEvents, strEventData

	strSQL = "SELECT Type, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
	strSQL = strSQL & "GROUP BY Type" & vbCRLF
	Set objEvents = Application("Connection").Execute(strSQL)
	
	strEventData = "['Category','Count',{ role: 'annotation' }, { role: 'annotation' } ],"
	
	If Not objEvents.EOF Then
		Do Until objEvents.EOF
			strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "','" & datEndDate & "'],"
			objEvents.MoveNext
		Loop
	End If%>
	
	google.setOnLoadCallback(draweventTypes);
	
	function draweventTypes() {
		
		var data = google.visualization.arrayToDataTable([
			<%=strEventData%>
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
	
		var chart = new google.visualization.PieChart(document.getElementById('eventTypes'));
		chart.draw(data, options);
	
		google.visualization.events.addListener(chart, 'select', selectHandler);
		
		function selectHandler(e) {
			var startDate = data.getValue(chart.getSelection()[0].row, 2)
			var endDate = data.getValue(chart.getSelection()[0].row, 3)
			window.open('events.asp?EventType=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + startDate + '&EndDate=' + endDate + '&Complete=All&View=Table','_self');
		}
	
	}
      
<%End Sub %>

<%Sub ModelCategoriesJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intIndex, objDeviceCount, intDeviceCount
	Dim objEvents, strEventData

	strSQL = "SELECT Category, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "# AND Model='" & Replace(Request.QueryString("Model"),"'","''") & "'" & vbCRLF
	strSQL = strSQL & "GROUP BY Category" & vbCRLF
	Set objEvents = Application("Connection").Execute(strSQL)
	
	strEventData = "['Category','Count',{ role: 'annotation' }, { role: 'annotation' } ],"
	
	If Not objEvents.EOF Then
		Do Until objEvents.EOF
			strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "','" & datEndDate & "'],"
			objEvents.MoveNext
		Loop
	End If%>
	
	google.setOnLoadCallback(drawModelCategories);
	
	function drawModelCategories() {
		
		var data = google.visualization.arrayToDataTable([
			<%=strEventData%>
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
	
		var chart = new google.visualization.PieChart(document.getElementById('modelCategories'));
		chart.draw(data, options);
	
		google.visualization.events.addListener(chart, 'select', selectHandler);
		
		function selectHandler(e) {
			var startDate = data.getValue(chart.getSelection()[0].row, 2)
			var endDate = data.getValue(chart.getSelection()[0].row, 3)
			window.open('events.asp?Category=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + startDate + '&EndDate=' + endDate + '&EventModel=<%=Request.QueryString("Model")%>&Complete=All&View=Table','_self');
		}
	
	}
      
<%End Sub %>

<%Sub ModelTypesJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intIndex, objDeviceCount, intDeviceCount
	Dim objEvents, strEventData

	strSQL = "SELECT Type, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "# AND Model='" & Replace(Request.QueryString("Model"),"'","''") & "'" & vbCRLF
	strSQL = strSQL & "GROUP BY Type" & vbCRLF
	Set objEvents = Application("Connection").Execute(strSQL)
	
	strEventData = "['Category','Count',{ role: 'annotation' }, { role: 'annotation' } ],"
	
	If Not objEvents.EOF Then
		Do Until objEvents.EOF
			strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "','" & datEndDate & "'],"
			objEvents.MoveNext
		Loop
	End If%>
	
	google.setOnLoadCallback(drawModelTypes);
	
	function drawModelTypes() {
		
		var data = google.visualization.arrayToDataTable([
			<%=strEventData%>
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
	
		var chart = new google.visualization.PieChart(document.getElementById('modelTypes'));
		chart.draw(data, options);
	
		google.visualization.events.addListener(chart, 'select', selectHandler);
		
		function selectHandler(e) {
			var startDate = data.getValue(chart.getSelection()[0].row, 2)
			var endDate = data.getValue(chart.getSelection()[0].row, 3)
			window.open('events.asp?EventType=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + startDate + '&EndDate=' + endDate + '&EventModel=<%=Request.QueryString("Model")%>&Complete=All&View=Table','_self');
		}
	
	}
      
<%End Sub %>

<%Sub CategoriesListJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intIndex, objDeviceCount, intDeviceCount
	Dim objEvents, strEventData, strFunctionName, strCardName
	
	If Not objEventData.EOF Then
		Do Until objEventData.EOF 

			strSQL = "SELECT Category, Count(ID) AS CountOfID" & vbCRLF
			strSQL = strSQL & "FROM Events" & vbCRLF
			strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "# AND Model='" & objEventData(0) & "'" & vbCRLF
			strSQL = strSQL & "GROUP BY Category" & vbCRLF
			Set objEvents = Application("Connection").Execute(strSQL)
	
			strEventData = "['Category','Count',{ role: 'annotation' }, { role: 'annotation' } ],"
	
			strFunctionName = objEventData(0)
			strFunctionName = Replace(strFunctionName," ","")
			strFunctionName = Replace(strFunctionName,".","")
			strFunctionName = Replace(strFunctionName,"'","")
			strFunctionName = Replace(strFunctionName,"-","")
			strCardName = strFunctionName
			strFunctionName = "Categories" & strFunctionName
	
			If Not objEvents.EOF Then
				Do Until objEvents.EOF
					strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "','" & datEndDate & "'],"
					objEvents.MoveNext
				Loop
			End If%>
			
			google.setOnLoadCallback(draw<%=strFunctionName%>);
	
			function draw<%=strFunctionName%>() {
		
				var data = google.visualization.arrayToDataTable([
					<%=strEventData%>
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
	
				var chart = new google.visualization.PieChart(document.getElementById('Categories<%=strCardName%>'));
				chart.draw(data, options);
	
				google.visualization.events.addListener(chart, 'select', selectHandler);
		
				function selectHandler(e) {
					var startDate = data.getValue(chart.getSelection()[0].row, 2)
					var endDate = data.getValue(chart.getSelection()[0].row, 3)
					window.open('events.asp?Category=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + startDate + '&EndDate=' + endDate + '&Complete=All&View=Table&EventModel=<%=objEventData(0)%>','_self');
				}
	
			}
		<% objEventData.MoveNext
		Loop
		objEventData.MoveFirst
	End If	
      
End Sub %>

<%Sub TypesListJavaScript 

	Dim strSQL, objOldestDevice, datOldestDevice, intIndex, objDeviceCount, intDeviceCount
	Dim objEvents, strEventData, strFunctionName, strCardName
	
	If Not objEventData.EOF Then
		Do Until objEventData.EOF 

			strSQL = "SELECT Type, Count(ID) AS CountOfID" & vbCRLF
			strSQL = strSQL & "FROM Events" & vbCRLF
			strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "# AND Model='" & objEventData(0) & "'" & vbCRLF
			strSQL = strSQL & "GROUP BY Type" & vbCRLF
			Set objEvents = Application("Connection").Execute(strSQL)
	
			strEventData = "['Category','Count',{ role: 'annotation' }, { role: 'annotation' } ],"
	
			strFunctionName = objEventData(0)
			strFunctionName = Replace(strFunctionName," ","")
			strFunctionName = Replace(strFunctionName,".","")
			strFunctionName = Replace(strFunctionName,"'","")
			strFunctionName = Replace(strFunctionName,"-","")
			strCardName = strFunctionName
			strFunctionName = "Types" & strFunctionName
	
			If Not objEvents.EOF Then
				Do Until objEvents.EOF
					strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "','" & datEndDate & "'],"
					objEvents.MoveNext
				Loop
			End If%>
	
			google.setOnLoadCallback(draw<%=strFunctionName%>);
	
			function draw<%=strFunctionName%>() {
		
				var data = google.visualization.arrayToDataTable([
					<%=strEventData%>
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
	
				var chart = new google.visualization.PieChart(document.getElementById('Types<%=strCardName%>'));
				chart.draw(data, options);
	
				google.visualization.events.addListener(chart, 'select', selectHandler);
		
				function selectHandler(e) {
					var startDate = data.getValue(chart.getSelection()[0].row, 2)
					var endDate = data.getValue(chart.getSelection()[0].row, 3)
					window.open('events.asp?EventType=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + startDate + '&EndDate=' + endDate + '&Complete=All&View=Table&EventModel=<%=objEventData(0)%>','_self');
				}
	
			}
						
			$(document).ready( function () {
				$("#<%=strCardName%>Card").flip({
					axis: 'y',
					trigger: 'manual'
			 	})
			 });
		
		
		<% 
		
			If strCardList <>  "" Then
				strCardList = strCardList & "," & strCardName & "Card"
			Else
				strCardList = strCardName & "Card"
			End If
		
			objEventData.MoveNext
		Loop
		objEventData.MoveFirst
	End If	
      
End Sub %>

<%Sub CategoriesByYearJavaScript

	Dim strSQL, objOldestDevice, datOldestDevice, intIndex, objDeviceCount, intDeviceCount
	Dim objEvents, strEventData
	
	If intYears > 0 Then
		For intIndex = 1 to intYears + 1

			strSQL = "SELECT Category, Count(Events.ID) AS CountOfID" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Events ON Devices.LGTag = Events.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Events.Model='" & Replace(Request.QueryString("Model"),"'","''") & "' AND "
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "# AND "
			strSQL = strSQL & "Events.EventDate>=#" & datStartDate & "# AND "
			strSQL = strSQL & "Events.EventDate<=#" & datEndDate & "#" & vbCRLF
			strSQL = strSQL & "GROUP BY Category" & vbCRLF
			Set objEvents = Application("Connection").Execute(strSQL)
			
			strEventData = "['Category','Count',{ role: 'annotation' }, { role: 'annotation' } ],"
	
			If Not objEvents.EOF Then
				Do Until objEvents.EOF
					strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "','" & datEndDate & "'],"
					objEvents.MoveNext
				Loop  %>
	
				google.setOnLoadCallback(drawCategoriesForYear<%=intIndex%>);
	
				function drawCategoriesForYear<%=intIndex%>() {
		
					var data = google.visualization.arrayToDataTable([
						<%=strEventData%>
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
	
					var chart = new google.visualization.PieChart(document.getElementById('CategoriesYear<%=intIndex%>'));
					chart.draw(data, options);
	
					google.visualization.events.addListener(chart, 'select', selectHandler);
		
					function selectHandler(e) {
						var startDate = data.getValue(chart.getSelection()[0].row, 2)
						var endDate = data.getValue(chart.getSelection()[0].row, 3)
						window.open('events.asp?Category=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + startDate + '&EndDate=' + endDate + '&Complete=All&View=Table&EventModel=<%=Request.QueryString("Model")%>&Year=<%=intIndex%>','_self');
					}
	
				}
		<% End If
		Next
	End If

End Sub%>

<%Sub TypesByYearJavaScript

	Dim strSQL, objOldestDevice, datOldestDevice, intIndex, objDeviceCount, intDeviceCount
	Dim objEvents, strEventData
	
	If intYears > 0 Then
		For intIndex = 1 to intYears + 1

			strSQL = "SELECT Type, Count(Events.ID) AS CountOfID" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Events ON Devices.LGTag = Events.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Events.Model='" & Replace(Request.QueryString("Model"),"'","''") & "' AND "
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "# AND "
			strSQL = strSQL & "Events.EventDate>=#" & datStartDate & "# AND "
			strSQL = strSQL & "Events.EventDate<=#" & datEndDate & "#" & vbCRLF
			strSQL = strSQL & "GROUP BY Type" & vbCRLF
			Set objEvents = Application("Connection").Execute(strSQL)
			
			strEventData = "['Category','Count',{ role: 'annotation' }, { role: 'annotation' } ],"
	
			If Not objEvents.EOF Then
				Do Until objEvents.EOF
					strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "','" & datEndDate & "'],"
					objEvents.MoveNext
				Loop  %>
	
				google.setOnLoadCallback(drawTypesForYear<%=intIndex%>);
	
				function drawTypesForYear<%=intIndex%>() {
		
					var data = google.visualization.arrayToDataTable([
						<%=strEventData%>
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
	
					var chart = new google.visualization.PieChart(document.getElementById('TypesYear<%=intIndex%>'));
					chart.draw(data, options);
	
					google.visualization.events.addListener(chart, 'select', selectHandler);
		
					function selectHandler(e) {
						var startDate = data.getValue(chart.getSelection()[0].row, 2)
						var endDate = data.getValue(chart.getSelection()[0].row, 3)
						window.open('events.asp?EventType=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + startDate + '&EndDate=' + endDate + '&Complete=All&View=Table&EventModel=<%=Request.QueryString("Model")%>&Year=<%=intIndex%>','_self');
					}
	
				}
		<% End If
		Next
	End If

End Sub%>

<%Function BuildDateURL

	If Request.Form("StartDate") <> "" Then
		BuildDateURL = BuildDateURL & "&StartDate=" & Request.Form("StartDate")
	End If
	If Request.Form("EndDate") <> "" Then
		BuildDateURL = BuildDateURL & "&EndDate=" & Request.Form("EndDate")
	End If
	
	If BuildDateURL = "" Then
		If Request.QueryString("StartDate") <> "" Then
			BuildDateURL = BuildDateURL & "&StartDate=" & Request.QueryString("StartDate")
		End If
		If Request.QueryString("EndDate") <> "" Then
			BuildDateURL = BuildDateURL & "&EndDate=" & Request.QueryString("EndDate")
		End If
	End If
	
	If BuildDateURL <> "" Then
		BuildDateURL = "?" & Right(BuildDateURL,Len(BuildDateURL) - 1)
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
End Function%>

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