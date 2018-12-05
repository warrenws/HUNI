<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/16/14
'Last Updated 1/14/18

'This is the main admin page for the inventory website

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim intTag, strFirstName, strLastName, strRole, strDeviceMessage, strUserMessage, strEventMessage
Dim strColumns, objDisabledUsers, intDisabledUsersCount, intEventNumber, objFirstNames, objLastNames
Dim strStudentsPerGradeInfo, strOpenEventsInfo, objShell, strStudentsHomeInternetInfo

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions

   Dim strSQL

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Search"
         LookupUser
			LookupDevice
			LookupEvent
   End Select

   'Get number of devices assigned to disabled students
   strSQL = "SELECT ID, UserName FROM People WHERE People.Active=False AND People.HasDevice=True"
   Set objDisabledUsers = Application("Connection").Execute(strSQL)
   If Not objDisabledUsers.EOF Then
   	Do Until objDisabledUsers.EOF
   		intDisabledUsersCount = intDisabledUsersCount + 1
   		objDisabledUsers.MoveNext
   	Loop
   End If

   'Get the list of firstnames for the auto complete
   strSQL = "SELECT DISTINCT FirstName FROM People WHERE Active=True"
   Set objFirstNames = Application("Connection").Execute(strSQL)

   'Get the list of lastnames for the auto complete
   strSQL = "SELECT DISTINCT LastName FROM People WHERE Active=True"
   Set objLastNames = Application("Connection").Execute(strSQL)

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

'   Get the chart if needed
'    If Application("LibreNMSServer") <> "" Then
'    	Set objShell = Server.CreateObject("WScript.Shell")
'    	objShell.Run "powershell ../scripts/CreateChart.ps1 -server """ &  Application("LibreNMSServer") & """ -token """ & Application("LibreNMSToken") & """" & _
'    		" -portID " & Application("BandwidthPort") & " -filename Bandwidth.svg",0,true
'    	Set objShell = Nothing
'    End If

   'Set up the variables needed for the site then load it
   SetupSite
   DisplaySite

End Sub%>

<%Sub DisplaySite

	Dim strSQL, arrIndexCards, strCard  %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title><%=Application("SiteName")%></title>
      <link rel="stylesheet" href="../assets/css/flipclock.css">
      <link rel="stylesheet" type="text/css" href="../style.css" />
      <link rel="apple-touch-icon" href="../images/inventory.png" />
      <link rel="shortcut icon" href="../images/inventory.ico" />
      <meta name="viewport" content="width=device-width,user-scalable=0" />
      <meta name="theme-color" content="#333333">
      <link rel="stylesheet" href="../assets/css/jquery-ui.css">
		<script src="../assets/js/jquery.js"></script>
		<script src="../assets/js/jquery-ui.js"></script>
		<script src="../assets/js/flipclock.min.js"></script>
		<script src="../assets/js/jquery.flip.min.js"></script>
		<script src="//www.google.com/jsapi"></script>
		<script>

		  	$(document).ready( function () {

				$(".Flipable").flip({
					axis: 'y',
					trigger: 'manual'
				});

				$(".FlipCard").on('click', function(e) {
					$(e.target).closest(".Flipable").flip('toggle');
					return false;
				});

		  	});

		  	google.load("visualization", "1", {packages:["corechart"]});

			<%
			arrIndexCards = Split(Application("IndexCards"),",")

			For Each strCard in arrIndexCards
				Select Case strCard

					Case "StudentsPerGradeWithMissing"
						StudentsPerGradeWithMissingJavaScript

					Case "StudentsPerGrade"
						StudentsPerGradeJavaScript

					Case "OpenEvents"
						OpenEventsJavaScript

					Case "SpareMacBooks"
						SpareMacBooksByGradeJavaScript

					Case "SpareiPadsByGrade"
						SpareiPadsByGradeJavaScript

					Case "SpareiPadsByType"
						SpareiPadsByTypeJavaScript

					Case "SpareMacBooksFlipiPads"
						SpareiPadsByGradeJavaScript
						'//SpareiPadsByTypeJavaScript
						SpareMacBooksByGradeJavaScript

					Case "SpareiPadsFlipMacBooks"
						SpareiPadsByGradeJavaScript
						'SpareiPadsByTypeJavaScript
						SpareMacBooksByGradeJavaScript

					Case "AccessHistory"
						StudentsPerGradeWithOutsideAccess
						StudentsPerGradeWithInsideAccess

					Case "EventStats"
						EventCategoriesJavaScript
						EventTypesJavaScript

				End Select
			Next

			%>

			$(document).ready(function () {

			<%	If Not IsMobile And Not IsiPad Then%>
					$( document ).tooltip({track: true});
			<% End If %>

		<% If Application("CountdownTimerTitle") <> "" Then %>
				var date = new Date("<%=Application("CountdownTimerDate")%>");
				var now = new Date();
				var diff = date.getTime()/1000 - now.getTime()/1000;
				if (diff > 0) {
					var clock = $('.your-clock').FlipClock(diff, {
						 clockFace: 'DailyCounter',
					 	countdown: true
					});
				}
		<% End If %>

		<% If Not objFirstNames.EOF Then %>
				var possibleFirstNames = [
			<% Do Until objFirstNames.EOF %>
					"<%=objFirstNames(0)%>",
				<%	objFirstNames.MoveNext
				Loop %>
			];
				$( "#FirstNames" ).autocomplete({
					source: possibleFirstNames
				});
		<% End If %>

		<% If Not objLastNames.EOF Then %>
				var possibleLastNames = [
			<% Do Until objLastNames.EOF %>
					"<%=objLastNames(0)%>",
				<%	objLastNames.MoveNext
				Loop %>
			];
				$( "#LastNames" ).autocomplete({
					source: possibleLastNames
				});
		<% End If %>
			})

  		</script>

   </head>

   <body class="<%=strSiteVersion%>" >

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
		arrIndexCards = Split(Application("IndexCards"),",")

		For Each strCard in arrIndexCards
			Select Case strCard
				Case "Warning"
					WarningCard

				Case "Search"
					SearchCard

				Case "StudentsPerGrade", "StudentsPerGradeWithMissing"
					StudentsPerGradeCard

				Case "OpenEvents"
					OpenEventsCard

				Case "SpareMacBooks"
					SpareMacBooksCard

				Case "SpareiPadsByGrade", "SpareiPadsByType"
					SpareiPadsCard

				Case "SpareMacBooksFlipiPads"
					SpareMacBooksFlipiPads

				Case "SpareiPadsFlipMacBooks"
					SpareiPadsFlipMacBooks

				Case "AccessHistory"
					AccessHistory

				Case "EventStats"
					EventStats

        Case "Bandwidth"
          If Application("LibreNMSServer") <> "" Then
            BandwidthCard
          End If

        Case "BandwidthWide"
          If Application("LibreNMSServer") <> "" Then
            If IsMobile Then
              BandwidthCard
            Else
              WideBandwidthCard
            End If
          End If

			End Select
		Next

		If Application("CountdownTimerTitle") <> "" Then
			If CDate(Date & " " & Time) < CDate(Application("CountdownTimerDate")) Then
				CountdownTimer
			End If
		End If
		%>
      </div>

      <div class="Version">Version <%=Application("Version")%></div>
      <div class="CopyRight"><%=Application("Copyright")%></div>
   </body>
   </html>

<%End Sub%>

<%Sub SpareMacBooksFlipiPads%>

	<div class="flip Flipable" >
		<div class="front">
			<%SpareMacBooksCardFlip%>
		</div>
		<div class="back">
			<%SpareiPadsCardFlip%>
		</div>
	</div>

<%End Sub%>

<%Sub BandwidthCard%>

  <div class="Card NormalCard">
    <div class="CardTitle">24 Hour Internet Usage</div>
    <a href="<%=Application("LibreNMSServer")%>/graphs/id=<%=Application("BandwidthPort")%>/type=port_bits" target="_blank">
      <img src="<%=Application("LibreNMSServer")%>/graph.php?id=<%=Application("BandwidthPort")%>&type=port_bits&width=300&height=150&from=end-24h">
    </a>
  </div>

<%End Sub%>

<%Sub WideBandwidthCard%>
  <br />
  <div class="Col3Card NormalCard">
    <div class="Col3CardTitle">24 Hour Internet Usage</div>
    <a href="<%=Application("LibreNMSServer")%>/graphs/id=<%=Application("BandwidthPort")%>/type=port_bits" target="_blank">
      <img src="<%=Application("LibreNMSServer")%>/graph.php?id=<%=Application("BandwidthPort")%>&type=port_bits&width=650&height=150&from=end-24h">
    </a>
  </div>
  <br />
<%End Sub%>

<%Sub SpareiPadsFlipMacBooks%>

	<div class="flip Flipable" >
		<div class="front">
			<%SpareiPadsCardFlip%>
		</div>
		<div class="back">
			<%SpareMacBooksCardFlip%>
		</div>
	</div>

<%End Sub%>

<%Sub AccessHistory %>
	<div class="flip Flipable" >
		<div class="front">
			<%StudentsWithOutsideAccessCard%>
		</div>
		<div class="back">
			<%StudentsWithInsideAccessCard%>
		</div>
	</div>
<%End Sub%>

<%Sub EventStats %>
	<div class="flip Flipable" >
		<div class="front">
			<%EventTypesCard%>
		</div>
		<div class="back">
			<%EventCategoriesCard%>
		</div>
	</div>
<%End Sub%>

<%Sub WarningCard

	 If intDisabledUsersCount >= 1 Then %>
			<div Class="SiteErrorCard">
				<div Class="Center">
					<a href="users.asp?UserStatus=Disabled&WithDevice=Yes">
					Devices are Assigned to Disabled Users
					</a>
				</div>
			</div>

	<% End If

End Sub%>

<%Sub CountdownTimer %>
	<div class="Card NormalCard">
		<div class="CardTitle"><%=Application("CountdownTimerTitle")%></div>
		<div class="your-clock"></div>
	</div>
<%End Sub%>

<%Sub SearchCard%>

	<div Class="Card NormalCard">
		<form method="POST" action="index.asp">
		<div class="CardTitle">Quick Search</div>
		<div>
			<div Class="CardColumn1">Asset tag: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthSmall" type="text" name="Tag" value="<%=intTag%>" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Event number: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthSmall" type="text" name="EventNumber" value="<%=intEventNumber%>" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">First name: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="FirstName" value="<%=strFirstName%>" id="FirstNames" />
			</div>
		</div>
		<div>
			<div Class="CardColumn1">Last name: </div>
			<div Class="CardColumn2">
				<input class="Card InputWidthLarge" type="text" name="LastName" value="<%=strLastName%>" id="LastNames" />
			</div>
		</div>

		<div class="Button"><input type="submit" value="Search" name="Submit" /></div>

	<% If strUserMessage <> "" Or strDeviceMessage <> "" Then %>
		<div>
			<div Class="Error">No Results</div>
		</div>
	<% End If %>

		<input type="hidden" value="" name="BOCESTag" />
		</form>
	</div>

<%End Sub%>

<%Sub StudentsPerGradeCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">
			Students Per Grade
			<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strStudentsPerGradeInfo%>"  />&nbsp;</div>
		</div>
		<div id="studentsPerGrade"></div>
	</div>
<%End Sub%>

<%Sub StudentsWithOutsideAccessCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">
			Outside Access Over the Past Week
			<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strStudentsHomeInternetInfo%>"  />&nbsp;</div>
		</div>
		<div id="studentsWithAccess" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Inside Access"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub StudentsWithInsideAccessCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">
			Inside Access Over the Past Week
		</div>
		<div id="studentsWithInsideAccess" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Outside Access"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub OpenEventsCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">
			Open Events
			<div class="Button">
			     <a href="users.asp?Owes=True&UserStatus=All&Display=Owed"><image src="../images/info.png" width="20" height="20" title="<%=strOpenEventsInfo%>"  /></a>&nbsp;</div>
		  </div>
		<div id="openEvents"></div>
	</div>
<%End Sub%>

<%Sub SpareMacBooksCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">Spare MacBooks</div>
		<div id="spareMacBooks" Class="Chart"></div>
	</div>
<%End Sub%>

<%Sub SpareMacBooksCardFlip%>
	<div class="Card NormalCard">
		<div class="CardTitle">Spare MacBooks</div>
		<div id="spareMacBooks" Class="Chart"></div>
		<div class="ChartBottomBar" >
			<div class="ChartBackLink">
				<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Spare iPads"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub SpareiPadsCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">Spare iPads</div>
		<div id="spareiPads" Class="Chart"></div>

	</div>
<%End Sub%>

<%Sub SpareiPadsCardFlip%>
	<div class="Card NormalCard">
		<div class="CardTitle">Spare iPads</div>
		<div id="spareiPads" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Spare MacBooks"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub EventCategoriesCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">Event Categories </div>
		<div id="eventCategories" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Types"/></a>
				<a href="eventstats.asp?LookUp=Categories"><image src="../images/dig.png" width="20" height="20" title="Dig Deeper"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub EventTypesCard%>
	<div class="Card NormalCard">
		<div class="CardTitle">Event Types </div>
		<div id="eventTypes" Class="Chart"></div>
		<div class="ChartBottomBar">
			<div class="ChartBackLink">
				<a href="" class="FlipCard"><image src="../images/swap.png" width="20" height="20" title="View Categories"/></a>
				<a href="eventstats.asp?LookUp=Types"><image src="../images/dig.png" width="20" height="20" title="Dig Deeper"/></a>
			</div>
		</div>
	</div>
<%End Sub%>

<%Sub StudentsPerGradeJavaScript

	Dim strSQL, objStudentsPerGrade, strStudentsPerGradeData, intHighestValue, intTotalStudentCount
	Dim intHSCount, intESCount, intMSCount, intHSMSCount, intNewComputerCount, intGrade, intSixToElevenCount
	Dim intSchuylervilleES, intSchuylervilleMS, intSchuylervilleHS, intSchuylervilleCount


	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND ClassOf > 2000" & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentsPerGrade = Application("Connection").Execute(strSQL)

	intHighestValue = 0
	intHSCount = 0
	intESCount = 0
	intMSCount = 0
	intHSMSCount = 0
	intSixToElevenCount = 0
	intTotalStudentCount = 0
	strStudentsPerGradeData = "['Grade','Students',{ role: 'annotation' },{ role: 'tooltip' }],"
	If Not objStudentsPerGrade.EOF Then
		Do Until objStudentsPerGrade.EOF
			strStudentsPerGradeData = strStudentsPerGradeData & "['" & GetGrade(Replace(objStudentsPerGrade(0),"'","\'")) & "', " & objStudentsPerGrade(1) & _
				",'" & objStudentsPerGrade(1) & "','" & objStudentsPerGrade(0) & "'],"
			intTotalStudentCount = intTotalStudentCount + objStudentsPerGrade(1)

			intGrade = GetGrade(objStudentsPerGrade(0))

			If intGrade = "K" Then
				intGrade = 0
			End If

			If intGrade >= 7 Then
				intHSMSCount = intHSMSCount + objStudentsPerGrade(1)
			Else
				intESCount = intESCount + objStudentsPerGrade(1)
			End If

			If intGrade = 7 Or intGrade = 8 Then
				intMSCount = intMSCount + objStudentsPerGrade(1)
			End If

			If intGrade = 4 Or intGrade = 8 Then
				intNewComputerCount = intNewComputerCount + objStudentsPerGrade(1)
			End If

			If intGrade >= 9 Then
				intHSCount = intHSCount + objStudentsPerGrade(1)
			End If

			If intGrade >=6 And intGrade <=11 Then
				intSixToElevenCount = intSixToElevenCount + objStudentsPerGrade(1)
			End If

			If objStudentsPerGrade(1) > intHighestValue Then
            intHighestValue = objStudentsPerGrade(1)
         End If

         If intGrade <= 5 Then
				intSchuylervilleES = intSchuylervilleES + objStudentsPerGrade(1)
			End If

			If intGrade >=6 And intGrade <=8 Then
				intSchuylervilleMS = intSchuylervilleMS + objStudentsPerGrade(1)
			End If

			 If intGrade >= 9 Then
				intSchuylervilleHS = intSchuylervilleHS + objStudentsPerGrade(1)
			End If

			If intGrade = 2 Or intGrade = 5 Then
				intSchuylervilleCount = intSchuylervilleCount + objStudentsPerGrade(1)
			End If

			objStudentsPerGrade.MoveNext
		Loop
		objStudentsPerGrade.MoveFirst
	End If
	strStudentsPerGradeData = Left(strStudentsPerGradeData,Len(strStudentsPerGradeData) - 1)

	Select Case Application("SiteName")

		Case "Lake George Inventory", "Lake George Inventory - Dev"

			strStudentsPerGradeInfo = "Elementary: " & intESCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Jr.-Sr. High: " & intHSMSCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "-----------------------" & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Elementary: " & intESCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Junior High: " & intMSCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Senior High: " & intHSCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "-----------------------" & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "4th and 8th: " & intNewComputerCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "5th - 11th: " & intSixToElevenCount

		Case "Schuylerville Inventory"
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Elementary: " & intSchuylervilleES & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Middle: " & intSchuylervilleMS & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "High School: " & intSchuylervilleHS & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "-----------------------" & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "2nd and 5th: " & intSchuylervilleCount

		Case Else

			strStudentsPerGradeInfo = ""

	End Select
	%>

	google.setOnLoadCallback(drawStudentPerGrade);

	function drawStudentPerGrade() {

		var data = google.visualization.arrayToDataTable([
			<%=strStudentsPerGradeData%>
		]);

		var options = {
			title: 'Total = <%=intTotalStudentCount%>',
			bar: {groupWidth: "90%"},
			chartArea:{width:'90%', height:'85%'},
			legend:{position: 'none'},
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {title: '', minValue: 0},
			vAxis: {viewWindow: {max : <%=intHighestValue%>},minValue: 0}
		};

		var chart = new google.visualization.ColumnChart(document.getElementById('studentsPerGrade'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			window.open('users.asp?Role=' + getGraduationYear(data.getValue(chart.getSelection()[0].row, 0)) + '&View=Table','_self');
		}

		function getGraduationYear(x) {

			if (x == 'K') {
				x = 0;
			}

			var d = new Date();
			var currentYear = d.getFullYear()
			var currentMonth = d.getMonth()
			currentMonth = currentMonth + 1;

			if (currentMonth>=7 && currentMonth<=12) {
				currentYear = currentYear + 1;
			}
			return currentYear + (12 - x)
		}
	}

<%End Sub %>

<%Sub StudentsPerGradeWithMissingJavaScript

	Dim strSQL, objStudentsPerGrade, strStudentsPerGradeData, intHighestValue, intTotalStudentCount
	Dim intHSCount, intESCount, intMSCount, intHSMSCount, intNewComputerCount, intGrade, bolShowTotalCount
	Dim objStudentsWithoutDevices, objStudentsWithDevices, intStudentsWithDevices, intStudentsWithoutDevices
	Dim intTotalWithDevices, intTotalWithoutDevices, intAssignedPercent, intSixToElevenCount
	Dim strToolTip, intSchuylervilleES, intSchuylervilleMS, intSchuylervilleHS, intSchuylervilleCount

	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND ClassOf > 2000" & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentsPerGrade = Application("Connection").Execute(strSQL)

	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND ClassOf > 2000 AND HasDevice=False" & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentsWithoutDevices = Application("Connection").Execute(strSQL)

	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND ClassOf > 2000 AND HasDevice=True" & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentsWithDevices = Application("Connection").Execute(strSQL)

	intHighestValue = 0
	intHSCount = 0
	intESCount = 0
	intMSCount = 0
	intHSMSCount = 0
	intTotalStudentCount = 0
	intTotalWithDevices = 0
	intTotalWithoutDevices = 0
	intSchuylervilleES = 0
	intSchuylervilleMS = 0
	intSchuylervilleHS = 0
	intSchuylervilleCount = 0
	strStudentsPerGradeData = "['Grade','Students',{ role: 'annotation' },{ role: 'tooltip' },'Without',{ role: 'annotation' },{ role: 'tooltip' }],"
	If Not objStudentsPerGrade.EOF Then
		Do Until objStudentsPerGrade.EOF

			intStudentsWithDevices = 0
			intStudentsWithoutDevices = 0
			bolShowTotalCount = False

			If Not objStudentsWithoutDevices.EOF Then
				Do Until objStudentsWithoutDevices.EOF
					If objStudentsWithoutDevices(0) = objStudentsPerGrade(0) Then
						intStudentsWithoutDevices = objStudentsWithoutDevices(1)
						intTotalWithoutDevices = intTotalWithoutDevices + intStudentsWithoutDevices
					End If
					objStudentsWithoutDevices.MoveNext
				Loop
			End If
			objStudentsWithoutDevices.MoveFirst

			If Not objStudentsWithDevices.EOF Then
				Do Until objStudentsWithDevices.EOF
					If objStudentsWithDevices(0) = objStudentsPerGrade(0) Then
						intStudentsWithDevices = objStudentsWithDevices(1)
						intTotalWithDevices = intTotalWithDevices + intStudentsWithDevices
					End If
					objStudentsWithDevices.MoveNext
				Loop
			End If
			objStudentsWithDevices.MoveFirst

			If intStudentsWithDevices = 0 Then
				intStudentsWithDevices = ""
				bolShowTotalCount = True
			End If
			If intStudentsWithoutDevices = 0 Then
				intStudentsWithoutDevices = ""
				bolShowTotalCount = True
			End If

			If bolShowTotalCount Then
				strToolTip = objStudentsPerGrade(0)
			Else
				strToolTip = objStudentsPerGrade(0) & " - " & objStudentsPerGrade(1)
			End If

			strStudentsPerGradeData = strStudentsPerGradeData & "['" & GetGrade(Replace(objStudentsPerGrade(0),"'","\'")) & _
			"', " & intStudentsWithDevices & _
			",'" & intStudentsWithDevices & _
			"','" & strToolTip & _
			"'," & intStudentsWithoutDevices & _
			",'" & intStudentsWithoutDevices & _
			"','" & strToolTip & "'],"
			intTotalStudentCount = intTotalStudentCount + objStudentsPerGrade(1)

			intGrade = GetGrade(objStudentsPerGrade(0))

			If intGrade = "K" Then
				intGrade = 0
			End If

			If intGrade >= 7 Then
				intHSMSCount = intHSMSCount + objStudentsPerGrade(1)
			Else
				intESCount = intESCount + objStudentsPerGrade(1)
			End If

			If intGrade = 7 Or intGrade = 8 Then
				intMSCount = intMSCount + objStudentsPerGrade(1)
			End If

			If intGrade = 4 Or intGrade = 8 Then
				intNewComputerCount = intNewComputerCount + objStudentsPerGrade(1)
			End If

			If intGrade >= 9 Then
				intHSCount = intHSCount + objStudentsPerGrade(1)
			End If

			If intGrade >=6 And intGrade <=11 Then
				intSixToElevenCount = intSixToElevenCount + objStudentsPerGrade(1)
			End If

			If objStudentsPerGrade(1) > intHighestValue Then
            intHighestValue = objStudentsPerGrade(1)
         End If

         If intGrade <= 5 Then
				intSchuylervilleES = intSchuylervilleES + objStudentsPerGrade(1)
			End If

			If intGrade >=6 And intGrade <=8 Then
				intSchuylervilleMS = intSchuylervilleMS + objStudentsPerGrade(1)
			End If

			 If intGrade >= 9 Then
				intSchuylervilleHS = intSchuylervilleHS + objStudentsPerGrade(1)
			End If

			If intGrade = 2 Or intGrade = 5 Then
				intSchuylervilleCount = intSchuylervilleCount + objStudentsPerGrade(1)
			End If

			objStudentsPerGrade.MoveNext
		Loop
		objStudentsPerGrade.MoveFirst
	End If

	If intTotalStudentCount <> 0 Then
		intAssignedPercent = Round((intTotalWithDevices/intTotalStudentCount)*100,2)
	Else
		intAssignedPercent = 0
	End If

	strStudentsPerGradeData = Left(strStudentsPerGradeData,Len(strStudentsPerGradeData) - 1)

	Select Case Application("SiteName")

		Case "Lake George Inventory", "Lake George Inventory - Dev"

			strStudentsPerGradeInfo = "Elementary: " & intESCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Jr.-Sr. High: " & intHSMSCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "-----------------------" & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Elementary: " & intESCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Junior High: " & intMSCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Senior High: " & intHSCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "-----------------------" & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "4th and 8th: " & intNewComputerCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "5th - 11th: " & intSixToElevenCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "-----------------------" & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Assigned: " & intAssignedPercent & "%"

		Case "Schuylerville Inventory"

			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Elementary: " & intSchuylervilleES & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Middle: " & intSchuylervilleMS & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "High School: " & intSchuylervilleHS & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "-----------------------" & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "2nd and 5th: " & intSchuylervilleCount & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "-----------------------" & " &#013 "
			strStudentsPerGradeInfo = strStudentsPerGradeInfo & "Assigned: " & intAssignedPercent & "%"

		Case Else

			strStudentsPerGradeInfo = ""

	End Select
	%>

	google.setOnLoadCallback(drawStudentPerGrade);

	function drawStudentPerGrade() {

		var data = google.visualization.arrayToDataTable([
			<%=strStudentsPerGradeData%>
		]);

		var options = {
			title: 'Total = <%=intTotalStudentCount%>',
			bar: {groupWidth: "90%"},
			chartArea:{width:'90%', height:'85%'},
			legend:{position: 'none'},
			isStacked: true,
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {title: '', minValue: 0},
			vAxis: {viewWindow: {max : <%=intHighestValue%>}, minValue: 0}
		};

		var chart = new google.visualization.ColumnChart(document.getElementById('studentsPerGrade'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			window.open('users.asp?Role=' + getGraduationYear(data.getValue(chart.getSelection()[0].row, 0)) + '&View=Table','_self');
		}

		function getGraduationYear(x) {

			if (x == 'K') {
				x = 0;
			}

			var d = new Date();
			var currentYear = d.getFullYear()
			var currentMonth = d.getMonth()
			currentMonth = currentMonth + 1;

			if (currentMonth>=7 && currentMonth<=12) {
				currentYear = currentYear + 1;
			}
			return currentYear + (12 - x)
		}
	}

<%End Sub %>

<%Sub StudentsPerGradeWithOutsideAccess

	Dim strSQL, objStudentsPerGrade, strStudentsPerGradeData, intHighestValue, intTotalStudentCount
	Dim intHSCount, intESCount, intMSCount, intHSMSCount, intNewComputerCount, intGrade, bolShowTotalCount
	Dim objStudentsWithoutDevices, objStudentsWithAccess, intStudentsWithAccess, intStudentsWithoutAccess
	Dim intTotalWithAccess, intTotalWithoutAccess, intTotalPercentWithAccess, strHSStudents
	Dim strToolTip, objInternetTypes, intUnknownInternetCount

	strHSStudents = GetGraduationYear(7)

	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (HomeRoom Is Not Null And HomeRoom  <>'') And ClassOf <= "  & strHSStudents & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentsPerGrade = Application("Connection").Execute(strSQL)

	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (HomeRoom Is Not Null And HomeRoom  <>'') And ClassOf <= " & strHSStudents & " AND LastExternalCheckIn>=Date()-7" & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentsWithAccess = Application("Connection").Execute(strSQL)

 	intHighestValue = 0
 	intHSCount = 0
 	intESCount = 0
 	intMSCount = 0
 	intHSMSCount = 0
 	intTotalStudentCount = 0
 	intTotalWithAccess = 0
 	intTotalWithoutAccess = 0
	strStudentsPerGradeData = "['Grade','Students',{ role: 'annotation' },{ role: 'tooltip' },'With Access',{ role: 'annotation' },{ role: 'tooltip' }],"
	If Not objStudentsPerGrade.EOF Then
		Do Until objStudentsPerGrade.EOF

			intStudentsWithAccess = 0
 			intStudentsWithoutAccess = 0
 			bolShowTotalCount = False


			If Not objStudentsWithAccess.EOF Then
				Do Until objStudentsWithAccess.EOF
					If objStudentsWithAccess(0) = objStudentsPerGrade(0) Then
						intStudentsWithAccess = objStudentsWithAccess(1)
						intTotalWithAccess = intTotalWithAccess + intStudentsWithAccess
					End If
					objStudentsWithAccess.MoveNext
				Loop
				objStudentsWithAccess.MoveFirst
			End If


			intStudentsWithoutAccess = objStudentsPerGrade(1) - intStudentsWithAccess

			If intStudentsWithAccess = 0 Then
				intStudentsWithAccess = ""
				bolShowTotalCount = True
			End If
			If intStudentsWithoutAccess = 0 Then
				intStudentsWithoutAccess = ""
				bolShowTotalCount = True
			End If

			If bolShowTotalCount Then
				strToolTip = objStudentsPerGrade(1)
			Else
				strToolTip = objStudentsPerGrade(1)' & " - " & objStudentsPerGrade(1)
			End If

			strStudentsPerGradeData = strStudentsPerGradeData & "['" & GetGrade(Replace(objStudentsPerGrade(0),"'","\'")) & _
			"', " & intStudentsWithAccess & _
			",'" & intStudentsWithAccess & _
			"','Have Accessed'" & _
			"," & intStudentsWithoutAccess & _
			",'" & intStudentsWithoutAccess & _
			"','Have Not Accessed'],"
			intTotalStudentCount = intTotalStudentCount + objStudentsPerGrade(1)

			If objStudentsPerGrade(1) > intHighestValue Then
            intHighestValue = objStudentsPerGrade(1)
         End If


			objStudentsPerGrade.MoveNext
		Loop
		objStudentsPerGrade.MoveFirst
	End If

	If intTotalStudentCount <> 0 Then
		intTotalPercentWithAccess = Round((intTotalWithAccess/intTotalStudentCount)*100,2)
	Else
		intTotalPercentWithAccess = 0
	End If 
	
	'Get the list of Internet types    
   strSQL = "SELECT InternetAccess, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (InternetAccess Is Not Null AND Not InternetAccess='')" & vbCRLF
	strSQL = strSQL & "GROUP BY InternetAccess" & vbCRLF
	strSQL = strSQL & "ORDER BY Count(ID) DESC;"
   Set objInternetTypes = Application("Connection").Execute(strSQL)
   
   intUnknownInternetCount = intTotalStudentCount
   Do Until objInternetTypes.EOF
   	strStudentsHomeInternetInfo = strStudentsHomeInternetInfo & objInternetTypes(0) & ": " & objInternetTypes(1) & " - " & Round((objInternetTypes(1)/intTotalStudentCount)*100,2) & "% &#013 "
   	intUnknownInternetCount = intUnknownInternetCount - objInternetTypes(1)
   	objInternetTypes.MoveNext
   Loop
	strStudentsHomeInternetInfo = strStudentsHomeInternetInfo & "Unknown: " & intUnknownInternetCount & " - " & Round((intUnknownInternetCount/intTotalStudentCount)*100,2) & "%"
	%>

	google.setOnLoadCallback(drawStudentWithAccess);

	function drawStudentWithAccess() {

		var data = google.visualization.arrayToDataTable([
			<%=strStudentsPerGradeData%>
		]);

		var options = {
			title: 'Total = <%=intTotalPercentWithAccess%>%',
			bar: {groupWidth: "90%"},
			chartArea:{left:40, width:'90%', height:'85%'},
			legend:{position: 'none'},
			isStacked: true,
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {title: '', minValue: 0},
			vAxis: {viewWindow: {max : <%=intHighestValue%>}, minValue: 0}
		};

		var chart = new google.visualization.ColumnChart(document.getElementById('studentsWithAccess'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			window.open('users.asp?Role=' + getGraduationYear(data.getValue(chart.getSelection()[0].row, 0)) + '&View=Table&Display=CheckIn&Internal=False','_self');
		}

		function getGraduationYear(x) {

			if (x == 'K') {
				x = 0;
			}

			var d = new Date();
			var currentYear = d.getFullYear()
			var currentMonth = d.getMonth()
			currentMonth = currentMonth + 1;

			if (currentMonth>=7 && currentMonth<=12) {
				currentYear = currentYear + 1;
			}
			return currentYear + (12 - x)
		}
	}

<%End Sub %>

<%Sub StudentsPerGradeWithInsideAccess

	Dim strSQL, objStudentsPerGrade, strStudentsPerGradeData, intHighestValue, intTotalStudentCount
	Dim intHSCount, intESCount, intMSCount, intHSMSCount, intNewComputerCount, intGrade, bolShowTotalCount
	Dim objStudentsWithoutDevices, objStudentsWithAccess, intStudentsWithAccess, intStudentsWithoutAccess
	Dim intTotalWithAccess, intTotalWithoutAccess, intTotalPercentWithAccess, strHSStudents
	Dim strToolTip

	strHSStudents = GetGraduationYear(5)

	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (HomeRoom Is Not Null And HomeRoom  <>'') And ClassOf <= "  & strHSStudents & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentsPerGrade = Application("Connection").Execute(strSQL)

	strSQL = "SELECT ClassOf, Count(People.ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE Active=True AND (HomeRoom Is Not Null And HomeRoom  <>'') And ClassOf <= " & strHSStudents & " AND LastInternalCheckIn>=Date()-7" & vbCRLF
	strSQL = strSQL & "GROUP BY ClassOf" & vbCRLF
	strSQL = strSQL & "ORDER BY ClassOf DESC"
	Set objStudentsWithAccess = Application("Connection").Execute(strSQL)

 	intHighestValue = 0
 	intHSCount = 0
 	intESCount = 0
 	intMSCount = 0
 	intHSMSCount = 0
 	intTotalStudentCount = 0
 	intTotalWithAccess = 0
 	intTotalWithoutAccess = 0
	strStudentsPerGradeData = "['Grade','Students',{ role: 'annotation' },{ role: 'tooltip' },'With Access',{ role: 'annotation' },{ role: 'tooltip' }],"
	If Not objStudentsPerGrade.EOF Then
		Do Until objStudentsPerGrade.EOF

			intStudentsWithAccess = 0
 			intStudentsWithoutAccess = 0
 			bolShowTotalCount = False


			If Not objStudentsWithAccess.EOF Then
				Do Until objStudentsWithAccess.EOF
					If objStudentsWithAccess(0) = objStudentsPerGrade(0) Then
						intStudentsWithAccess = objStudentsWithAccess(1)
						intTotalWithAccess = intTotalWithAccess + intStudentsWithAccess
					End If
					objStudentsWithAccess.MoveNext
				Loop
				objStudentsWithAccess.MoveFirst
			End If

			intStudentsWithoutAccess = objStudentsPerGrade(1) - intStudentsWithAccess

			If intStudentsWithAccess = 0 Then
				intStudentsWithAccess = ""
				bolShowTotalCount = True
			End If
			If intStudentsWithoutAccess = 0 Then
				intStudentsWithoutAccess = ""
				bolShowTotalCount = True
			End If

			If bolShowTotalCount Then
				strToolTip = objStudentsPerGrade(1)
			Else
				strToolTip = objStudentsPerGrade(1)' & " - " & objStudentsPerGrade(1)
			End If

			strStudentsPerGradeData = strStudentsPerGradeData & "['" & GetGrade(Replace(objStudentsPerGrade(0),"'","\'")) & _
			"', " & intStudentsWithAccess & _
			",'" & intStudentsWithAccess & _
			"','Have Accessed'" & _
			"," & intStudentsWithoutAccess & _
			",'" & intStudentsWithoutAccess & _
			"','Have Not Accessed'],"
			intTotalStudentCount = intTotalStudentCount + objStudentsPerGrade(1)

			If objStudentsPerGrade(1) > intHighestValue Then
            intHighestValue = objStudentsPerGrade(1)
         End If


			objStudentsPerGrade.MoveNext
		Loop
		objStudentsPerGrade.MoveFirst
	End If

	If intTotalStudentCount <> 0 Then
		intTotalPercentWithAccess = Round((intTotalWithAccess/intTotalStudentCount)*100,2)
	Else
		intTotalPercentWithAccess = 0
	End If %>

	google.setOnLoadCallback(drawStudentWithInsideAccess);

	function drawStudentWithInsideAccess() {

		var data = google.visualization.arrayToDataTable([
			<%=strStudentsPerGradeData%>
		]);

		var options = {
			title: 'Total = <%=intTotalPercentWithAccess%>%',
			bar: {groupWidth: "90%"},
			chartArea:{left:40, width:'90%', height:'85%'},
			legend:{position: 'none'},
			isStacked: true,
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {title: '', minValue: 0},
			vAxis: {viewWindow: {max : <%=intHighestValue%>}, minValue: 0}
		};

		var chart = new google.visualization.ColumnChart(document.getElementById('studentsWithInsideAccess'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			window.open('users.asp?Role=' + getGraduationYear(data.getValue(chart.getSelection()[0].row, 0)) + '&View=Table&Display=CheckIn&Internal=True','_self');
		}

		function getGraduationYear(x) {

			if (x == 'K') {
				x = 0;
			}

			var d = new Date();
			var currentYear = d.getFullYear()
			var currentMonth = d.getMonth()
			currentMonth = currentMonth + 1;

			if (currentMonth>=7 && currentMonth<=12) {
				currentYear = currentYear + 1;
			}
			return currentYear + (12 - x)
		}
	}

<%End Sub %>

<%Sub OpenEventsJavaScript

	Dim strSQL, objOpenEventCounts, strOpenEventsData, intHighestValue, objOwedMoney

	strSQL = "SELECT Type, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE Resolved=False" & vbCRLF
	strSQL = strSQL & "GROUP BY Type"
	Set objOpenEventCounts = Application("Connection").Execute(strSQL)

	strSQL = "SELECT Sum(Price) AS SumOfPrice FROM Owed GROUP BY Active HAVING Active=True"
	Set objOwedMoney = Application("Connection").Execute(strSQL)

	If Not objOwedMoney.EOF Then
		strOpenEventsInfo = "$" & objOwedMoney(0) & " Currently Owed"
	Else
		strOpenEventsInfo = "No Money Owed"
	End If

	strOpenEventsData = "['Type','Events',{ role: 'annotation' } ],"
	If Not objOpenEventCounts.EOF Then
		Do Until objOpenEventCounts.EOF

			If objOpenEventCounts(1) > intHighestValue Then
            intHighestValue = objOpenEventCounts(1)
         End If

			strOpenEventsData = strOpenEventsData & "['" & Replace(objOpenEventCounts(0),"'","\'") & "', " & objOpenEventCounts(1) & ",'" & objOpenEventCounts(1) & "'],"
			objOpenEventCounts.MoveNext


		Loop
		objOpenEventCounts.MoveFirst
	End If
	strOpenEventsData = Left(strOpenEventsData,Len(strOpenEventsData) - 1)%>

	google.setOnLoadCallback(drawOpenEvents);

	function drawOpenEvents() {

		var data = google.visualization.arrayToDataTable([
			<%=strOpenEventsData%>
		]);

		var options = {
			titlePosition: 'none',
			chartArea:{left:85, width:'90%', height:'85%'},
			legend: 'none',
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {viewWindow: {max : <%=intHighestValue*1.1%>},minValue: 0},
			vAxis: {title: ''}
		};

		var chart = new google.visualization.BarChart(document.getElementById('openEvents'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			window.open('events.asp?EventType=' + data.getValue(chart.getSelection()[0].row, 0) + '&Complete=No&View=Table','_self');
		}

	}

<%End Sub %>

<%Sub SpareMacBooksJavaScript

	Dim strSQL, objSpareDevices, strSpareDevicesData, intDeviceYear, strDeviceName, strDeviceSite, intYears, objOldestDevice, datOldestDevice
	Dim intIndex, objSpares, intHighestValue, strModel, objLoanedSpares, intAvailableSpares, intLoanedSpares, objAvailableSpares
	Dim strAvailableSparesLabel, strLoanedSpareLabel, objReplacements, intReplacements, intReplacementLabel

	'Get the oldest device from the inventory
   strSQL = "SELECT DatePurchased FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
   strSQL = strSQL & "WHERE DatePurchased Is Not Null AND Active=True AND Tag='Spare' AND (Model Like '%MacBook%' Or Model Like '%iPad%') AND Assigned=False ORDER BY DatePurchased"

   Set objOldestDevice = Application("Connection").Execute(strSQL)
   If Not objOldestDevice.EOF Then
      datOldestDevice = objOldestDevice(0)
      intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice)
   End If

	strSpareDevicesData = "['Device','Available',{ role: 'annotation' },'Loaned Out',{ role: 'annotation' },'Issued as Replacements',{ role: 'annotation' }],"
	intHighestValue = 0

	For intIndex = 1 to intYears

		'Get the number of spares
		strSQL = "SELECT Site, Model, Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
		strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND (Model Like '%MacBook%') AND ("
		strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
		strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "
		strSQL = strSQL & "GROUP BY Site, Model"
		Set objSpares = Application("Connection").Execute(strSQL)

		Do Until objSpares.EOF

			'Get the number of available spares
			strSQL = "SELECT Site, Model, Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
			strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND (Model Like '%MacBook%') AND Site='" & objSpares(0) & "' AND Assigned=False AND ("
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "
			strSQL = strSQL & "GROUP BY Site, Model"
			Set objAvailableSpares = Application("Connection").Execute(strSQL)

			If objAvailableSpares.EOF Then
				intAvailableSpares = 0
			Else
				intAvailableSpares = objAvailableSpares(2)
			End If

			'Get the number of loaned spares
			strSQL = "SELECT Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
			strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND Model='" & objSpares(1) & "' AND Assigned=True AND ("
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) AND "
			strSQL = strSQL & "Site='" & objSpares(0) & "'"
			Set objLoanedSpares = Application("Connection").Execute(strSQL)

			'Get the number of devices issued as replacement
			strSQL = "SELECT Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
			strSQL = strSQL & "WHERE Tag='Replacement' AND Active=True AND Model='" & objSpares(1) & "' AND Assigned=True AND ("
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) AND "
			strSQL = strSQL & "Site='" & objSpares(0) & "'"
			Set objReplacements = Application("Connection").Execute(strSQL)

			If Not objLoanedSpares.EOF Then
				intLoanedSpares = objLoanedSpares(0)
			Else
				intLoanedSpares = 0
			End If

			If Not objReplacements.EOF Then
				intReplacements = objReplacements(0)
			Else
				intReplacements = 0
			End If

			If intAvailableSpares <= 1 Then
				strAvailableSparesLabel = ""
			Else
				strAvailableSparesLabel = intAvailableSpares
			End If

			If intReplacements = 0 Then
				intReplacementLabel = ""
			ElseIf intReplacements = 1 Then
				intReplacementLabel = ""
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			Else
				intReplacementLabel = intReplacements
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			End If

			If intLoanedSpares <= 1 Then
				strLoanedSpareLabel = ""
			Else
				strLoanedSpareLabel = intLoanedSpares
			End If

			If (intAvailableSpares + intLoanedSpares + intReplacements) > intHighestValue Then
            intHighestValue = (intAvailableSpares + intLoanedSpares + intReplacements)
         End If

			Select Case objSpares(1)
				Case "MacBook Air"
					strModel = "Air"
				Case "MacBook Pro"
					strModel = "Pro"
				Case Else
					strModel = objSpares(1)
			End Select

			Select Case objSpares(0)
				Case "Elementary"
					strDeviceSite = "ES"
				Case "High School"
					strDeviceSite = "HS"
				Case Else
					strDeviceSite  =objSpares(0)
			End Select

			strSpareDevicesData = strSpareDevicesData & "['" & strModel & " - " & strDeviceSite & " - Yr " & _
			intIndex & "'," & intAvailableSpares & ",'" & strAvailableSparesLabel & "'," & intLoanedSpares & _
			",'" & strLoanedSpareLabel & "'," & intReplacements & ",'" & intReplacementLabel & "'],"
			objSpares.MoveNext
		Loop

	Next
	strSpareDevicesData = Left(strSpareDevicesData,Len(strSpareDevicesData) - 1)%>

	google.setOnLoadCallback(drawSpareMacBooks);

	function drawSpareMacBooks() {

		var data = google.visualization.arrayToDataTable([
			<%=strSpareDevicesData%>
		]);

		var options = {
			titlePosition: 'none',
			chartArea:{left:75, width:'90%', height:'85%'},
			legend: 'none',
			isStacked: true,
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {viewWindow: {max : <%=intHighestValue*1.1%>},minValue: 0},
			vAxis: {title: ''}

		};

		var chart = new google.visualization.BarChart(document.getElementById('spareMacBooks'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			var dataFromChart = data.getValue(chart.getSelection()[0].row, 0).split('-')

			var deviceModel = dataFromChart[0].trim();
			var deviceSite = dataFromChart[1].trim();
			var deviceYear = dataFromChart[2].trim();
			deviceYear = deviceYear.substr(deviceYear.length - 1);

			//window.alert(data.getValue(chart.getSelection()[0].row, 3));

			switch(deviceSite) {
				case 'ES':
					deviceSite = 'Elementary';
					break;
				case 'HS':
					deviceSite = 'High School';
					break;
			}

			switch(deviceModel) {
				case 'Air':
					deviceModel = 'MacBook Air';
					break;
				case 'Pro':
					deviceModel = 'MacBook Pro';
					break;
			}

			window.open('devices.asp?Model=' + deviceModel + '&DeviceSite=' + deviceSite + '&Year=' + deviceYear + '&Tags=Spare&View=Table','_self');
		}

	}

<%End Sub %>

<%Sub SpareMacBooksByGradeJavaScript

	Dim strSQL, objSpareDevices, strSpareDevicesData, intDeviceYear, strDeviceName, strDeviceSite, intYears, objOldestDevice, datOldestDevice
	Dim intIndex, objSpares, intHighestValue, strModel, objLoanedSpares, intAvailableSpares, intLoanedSpares, objAvailableSpares
	Dim strAvailableSparesLabel, strLoanedSpareLabel, objReplacements, intReplacements, intReplacementLabel, strYear

	'Get the current seniors graduating year
	strYear = GetGraduationYear(12)

	strSpareDevicesData = "['Device','Available',{ role: 'annotation' },'Loaned Out',{ role: 'annotation' },'Issued as Replacements',{ role: 'annotation' }],"
	intHighestValue = 0

	For intIndex = strYear to strYear + 7

		'Get the number of spares
		strSQL = "SELECT Count(TagCount) AS CountofCount" & vbCRLF
		strSQL = strSQL & "FROM" & vbCRLF
		strSQL = strSQL & "(SELECT Count(Tags.Tag) AS TagCount,Devices.LGTag" & vbCRLF
		strSQL = strSQL & "FROM(" & vbCRLF
		strSQL = strSQL & "SELECT Devices.ID, Devices.LGTag, Tags.Tag" & vbCRLF
		strSQL = strSQL & "FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
		strSQL = strSQL & "WHERE Devices.Active=True AND Devices.Deleted=False AND Model Like '%MacBook%')" & vbCRLF
		strSQL = strSQL & "WHERE (Tags.Tag='" & intIndex & "' OR Tags.Tag='Spare')" & vbCRLF
		strSQL = strSQL & "GROUP BY Devices.LGTag)" & vbCRLF
		strSQL = strSQL & "WHERE TagCount=2"
		Set objSpares = Application("Connection").Execute(strSQL)

		If Not IsEmpty(objSpares) Then

			'Get the number of available spares
			strSQL = "SELECT Count(TagCount) AS CountofCount" & vbCRLF
			strSQL = strSQL & "FROM" & vbCRLF
			strSQL = strSQL & "(SELECT Count(Tags.Tag) AS TagCount,Devices.LGTag" & vbCRLF
			strSQL = strSQL & "FROM(" & vbCRLF
			strSQL = strSQL & "SELECT Devices.ID, Devices.LGTag, Tags.Tag" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Devices.Active=True AND Devices.Deleted=False AND Assigned=False AND Model Like '%MacBook%')" & vbCRLF
			strSQL = strSQL & "WHERE (Tags.Tag='" & intIndex & "' OR Tags.Tag='Spare')" & vbCRLF
			strSQL = strSQL & "GROUP BY Devices.LGTag)" & vbCRLF
			strSQL = strSQL & "WHERE TagCount=2"
			Set objAvailableSpares = Application("Connection").Execute(strSQL)

			'Get the number of loaned spares
			strSQL = "SELECT Count(TagCount) AS CountofCount" & vbCRLF
			strSQL = strSQL & "FROM" & vbCRLF
			strSQL = strSQL & "(SELECT Count(Tags.Tag) AS TagCount,Devices.LGTag" & vbCRLF
			strSQL = strSQL & "FROM(" & vbCRLF
			strSQL = strSQL & "SELECT Devices.ID, Devices.LGTag, Tags.Tag" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Devices.Active=True AND Devices.Deleted=False AND Assigned=True AND Model Like '%MacBook%')" & vbCRLF
			strSQL = strSQL & "WHERE (Tags.Tag='" & intIndex & "' OR Tags.Tag='Spare')" & vbCRLF
			strSQL = strSQL & "GROUP BY Devices.LGTag)" & vbCRLF
			strSQL = strSQL & "WHERE TagCount=2"
			Set objLoanedSpares = Application("Connection").Execute(strSQL)

			'Get the number of devices issued as replacement
			strSQL = "SELECT Count(TagCount) AS CountofCount" & vbCRLF
			strSQL = strSQL & "FROM" & vbCRLF
			strSQL = strSQL & "(SELECT Count(Tags.Tag) AS TagCount,Devices.LGTag" & vbCRLF
			strSQL = strSQL & "FROM(" & vbCRLF
			strSQL = strSQL & "SELECT Devices.ID, Devices.LGTag, Tags.Tag" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Devices.Active=True AND Devices.Deleted=False AND Assigned=True AND Model Like '%MacBook%')" & vbCRLF
			strSQL = strSQL & "WHERE (Tags.Tag='" & intIndex & "' OR Tags.Tag='Spare' OR Tags.Tag='Replacement')" & vbCRLF
			strSQL = strSQL & "GROUP BY Devices.LGTag)" & vbCRLF
			strSQL = strSQL & "WHERE TagCount=3"
			Set objReplacements = Application("Connection").Execute(strSQL)

			If objAvailableSpares.EOF Then
				intAvailableSpares = 0
			Else
				intAvailableSpares = objAvailableSpares(0)
			End If

			If Not objLoanedSpares.EOF Then
				intLoanedSpares = objLoanedSpares(0)
			Else
				intLoanedSpares = 0
			End If

			If Not objReplacements.EOF Then
				intReplacements = objReplacements(0)
			Else
				intReplacements = 0
			End If

			If intAvailableSpares <= 1 Then
				strAvailableSparesLabel = ""
			Else
				strAvailableSparesLabel = intAvailableSpares
			End If

			If intReplacements = 0 Then
				intReplacementLabel = ""
			ElseIf intReplacements = 1 Then
				intReplacementLabel = ""
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			Else
				intReplacementLabel = intReplacements
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			End If

			If intLoanedSpares <= 1 Then
				strLoanedSpareLabel = ""
			Else
				strLoanedSpareLabel = intLoanedSpares
			End If

			If (intAvailableSpares + intLoanedSpares + intReplacements) > intHighestValue Then
            intHighestValue = (intAvailableSpares + intLoanedSpares + intReplacements)
         End If

			strSpareDevicesData = strSpareDevicesData & "['" & intIndex & "'," & intAvailableSpares & _
			",'" & strAvailableSparesLabel & "'," & intLoanedSpares & _
			",'" & strLoanedSpareLabel & "'," & intReplacements & ",'" & intReplacementLabel & "'],"

		End If

	Next
	strSpareDevicesData = Left(strSpareDevicesData,Len(strSpareDevicesData) - 1)%>

	google.setOnLoadCallback(drawSpareMacBooks);

	function drawSpareMacBooks() {

		var data = google.visualization.arrayToDataTable([
			<%=strSpareDevicesData%>
		]);

		var options = {
			titlePosition: 'none',
			chartArea:{left:40, width:'90%', height:'85%'},
			legend: 'none',
			isStacked: true,
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {viewWindow: {max : <%=intHighestValue*1.1%>},minValue: 0},
			vAxis: {title: ''}

		};

		var chart = new google.visualization.BarChart(document.getElementById('spareMacBooks'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			var deviceYear = data.getValue(chart.getSelection()[0].row, 0)
			window.open('devices.asp?Model=MacBook&Tags=Spare,' + deviceYear + '&View=Table','_self');
		}

	}

<%End Sub %>

<%Sub SpareiPadJavaScript

	Dim strSQL, objSpareDevices, strSpareDevicesData, intDeviceYear, strDeviceName, strDeviceSite, intYears, intDeviceCount, objOldestDevice, datOldestDevice
	Dim intIndex, objSpares, intHighestValue, strModel, objLoanedSpares, intAvailableSpares, intLoanedSpares
	Dim strAvailableSparesLabel, strLoanedSpareLabel, objReplacements, intReplacements, intReplacementLabel, objAvailableSpares

	'Get the oldest device from the inventory
   strSQL = "SELECT DatePurchased FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
   strSQL = strSQL & "WHERE DatePurchased Is Not Null AND Active=True AND Tag='Spare' AND (Model Like '%MacBook%' Or Model Like '%iPad%') AND Assigned=False ORDER BY DatePurchased"

   Set objOldestDevice = Application("Connection").Execute(strSQL)
   If Not objOldestDevice.EOF Then
      datOldestDevice = objOldestDevice(0)
      intYears = DatePart("yyyy",Date) - DatePart("yyyy",datOldestDevice)
   End If

	strSpareDevicesData = "['Device','Available',{ role: 'annotation' },'Loaned Out',{ role: 'annotation' },'Issued as Replacements',{ role: 'annotation' }],"
	intHighestValue = 0

	For intIndex = 1 to intYears

		intAvailableSpares = 0
		intLoanedSpares = 0

		'Get the number of available spares
		strSQL = "SELECT Site, Model, Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
		strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND (Model Like '%iPad%') AND ("
		strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
		strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "
		strSQL = strSQL & "GROUP BY Site, Model"
		Set objSpares = Application("Connection").Execute(strSQL)

		Do Until objSpares.EOF

			'Get the number of available spares
			strSQL = "SELECT Site, Model, Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
			strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND (Model Like '%iPad%') AND Assigned=False AND Site='" & objSpares(0) & "' AND Assigned=False AND ("
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) "
			strSQL = strSQL & "GROUP BY Site, Model"
			Set objAvailableSpares = Application("Connection").Execute(strSQL)

			If objAvailableSpares.EOF Then
				intAvailableSpares = 0
			Else
				intAvailableSpares = objAvailableSpares(2)
			End If

			'Get the number of loaned spares
			strSQL = "SELECT Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
			strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND Model='" & objSpares(1) & "' AND Assigned=True AND ("
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) AND "
			strSQL = strSQL & "Site='" & objSpares(0) & "'"
			Set objLoanedSpares = Application("Connection").Execute(strSQL)

			'Get the number of devices issued as replacement
			strSQL = "SELECT Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
			strSQL = strSQL & "WHERE Tag='Replacement' AND Active=True AND Model='" & objSpares(1) & "' AND Assigned=True AND ("
			strSQL = strSQL & "DatePurchased>=#" & DateAdd("yyyy",intIndex * -1,Date) & "# AND "
			strSQL = strSQL & "DatePurchased<=#" & DateAdd("yyyy",(intIndex -1) * -1,Date) & "#) AND "
			strSQL = strSQL & "Site='" & objSpares(0) & "'"
			Set objReplacements = Application("Connection").Execute(strSQL)

			If Not objLoanedSpares.EOF Then
				intLoanedSpares = objLoanedSpares(0)
			Else
				intLoanedSpares = 0
			End If

			If Not objReplacements.EOF Then
				intReplacements = objReplacements(0)
			Else
				intReplacements = 0
			End If

			If intAvailableSpares <= 1 Then
				strAvailableSparesLabel = ""
			Else
				strAvailableSparesLabel = intAvailableSpares
			End If

			If intReplacements = 0 Then
				intReplacementLabel = ""
			ElseIf intReplacements = 1 Then
				intReplacementLabel = ""
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			Else
				intReplacementLabel = intReplacements
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			End If

			If intLoanedSpares <= 1 Then
				strLoanedSpareLabel = ""
			Else
				strLoanedSpareLabel = intLoanedSpares
			End If

			If intAvailableSpares + intLoanedSpares > intHighestValue Then
            intHighestValue = intAvailableSpares + intLoanedSpares
         End If

			Select Case objSpares(1)
				Case "iPad Air", "iPad Air 1"
					strModel = "Air 1"
				Case "iPad Air 2"
					strModel = "Air 2"
				Case "iPad Pro 12.9"
					strModel = "Pro 13"
				Case "iPad Pro 9.7"
					strModel = "Pro"
				Case Else
					strModel = objSpares(1)
			End Select

			Select Case objSpares(0)
				Case "Elementary"
					strDeviceSite = "ES"
				Case "High School"
					strDeviceSite = "HS"
				Case Else
					strDeviceSite  =objSpares(0)
			End Select

			strSpareDevicesData = strSpareDevicesData & "['" & strModel & " - " & strDeviceSite & " - Yr " & _
			intIndex & "'," & intAvailableSpares & ",'" & strAvailableSparesLabel & "'," & intLoanedSpares & _
			",'" & strLoanedSpareLabel & "'," & intReplacements & ",'" & intReplacementLabel & "'],"
			objSpares.MoveNext
		Loop

	Next
	strSpareDevicesData = Left(strSpareDevicesData,Len(strSpareDevicesData) - 1)%>

	google.setOnLoadCallback(drawSpareiPads);

	function drawSpareiPads() {

		var data = google.visualization.arrayToDataTable([
			<%=strSpareDevicesData%>
		]);

		var options = {
			titlePosition: 'none',
			chartArea:{left:85, width:'90%', height:'85%'},
			legend: 'none',
			isStacked: true,
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {viewWindow: {max : <%=intHighestValue*1.2%>},minValue: 0},
			vAxis: {title: ''}
		};

		var chart = new google.visualization.BarChart(document.getElementById('spareiPads'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			var dataFromChart = data.getValue(chart.getSelection()[0].row, 0).split('-')

			var deviceModel = dataFromChart[0].trim();
			var deviceSite = dataFromChart[1].trim();
			var deviceYear = dataFromChart[2].trim();
			deviceYear = deviceYear.substr(deviceYear.length - 1);

			switch(deviceSite) {
				case 'ES':
					deviceSite = 'Elementary';
					break;
				case 'HS':
					deviceSite = 'High School';
					break;
			}

			switch(deviceModel) {
				case 'Air 1':
					deviceModel = 'iPad Air 1';
					break;
				case 'Air 2':
					deviceModel = 'iPad Air 2';
					break;
				case 'Pro':
					deviceModel = 'iPad Pro';
					break;
			}

			window.open('devices.asp?Model=' + deviceModel + '&DeviceSite=' + deviceSite + '&Year=' + deviceYear + '&Tags=Spare&View=Table','_self');
		}

	}

<%End Sub %>

<%Sub SpareiPadsByGradeJavaScript

	Dim strSQL, objSpareDevices, strSpareDevicesData, intDeviceYear, strDeviceName, strDeviceSite, intYears, objOldestDevice, datOldestDevice
	Dim intIndex, objSpares, intHighestValue, strModel, objLoanedSpares, intAvailableSpares, intLoanedSpares, objAvailableSpares
	Dim strAvailableSparesLabel, strLoanedSpareLabel, objReplacements, intReplacements, intReplacementLabel, strYear

	'Get the current seniors graduating year
	strYear = GetGraduationYear(4)

	strSpareDevicesData = "['Device','Available',{ role: 'annotation' },'Loaned Out',{ role: 'annotation' },'Issued as Replacements',{ role: 'annotation' }],"
	intHighestValue = 0

	For intIndex = strYear to strYear + 4

		'Get the number of spares
		strSQL = "SELECT Count(TagCount) AS CountofCount" & vbCRLF
		strSQL = strSQL & "FROM" & vbCRLF
		strSQL = strSQL & "(SELECT Count(Tags.Tag) AS TagCount,Devices.LGTag" & vbCRLF
		strSQL = strSQL & "FROM(" & vbCRLF
		strSQL = strSQL & "SELECT Devices.ID, Devices.LGTag, Tags.Tag" & vbCRLF
		strSQL = strSQL & "FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
		strSQL = strSQL & "WHERE Devices.Active=True AND Devices.Deleted=False AND Model Like '%iPad%')" & vbCRLF
		strSQL = strSQL & "WHERE (Tags.Tag='" & intIndex & "' OR Tags.Tag='Spare')" & vbCRLF
		strSQL = strSQL & "GROUP BY Devices.LGTag)" & vbCRLF
		strSQL = strSQL & "WHERE TagCount=2"
		Set objSpares = Application("Connection").Execute(strSQL)

		If Not IsEmpty(objSpares) Then

			'Get the number of available spares
			strSQL = "SELECT Count(TagCount) AS CountofCount" & vbCRLF
			strSQL = strSQL & "FROM" & vbCRLF
			strSQL = strSQL & "(SELECT Count(Tags.Tag) AS TagCount,Devices.LGTag" & vbCRLF
			strSQL = strSQL & "FROM(" & vbCRLF
			strSQL = strSQL & "SELECT Devices.ID, Devices.LGTag, Tags.Tag" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Devices.Active=True AND Devices.Deleted=False AND Assigned=False AND Model Like '%iPad%')" & vbCRLF
			strSQL = strSQL & "WHERE (Tags.Tag='" & intIndex & "' OR Tags.Tag='Spare')" & vbCRLF
			strSQL = strSQL & "GROUP BY Devices.LGTag)" & vbCRLF
			strSQL = strSQL & "WHERE TagCount=2"
			Set objAvailableSpares = Application("Connection").Execute(strSQL)

			'Get the number of loaned spares
			strSQL = "SELECT Count(TagCount) AS CountofCount" & vbCRLF
			strSQL = strSQL & "FROM" & vbCRLF
			strSQL = strSQL & "(SELECT Count(Tags.Tag) AS TagCount,Devices.LGTag" & vbCRLF
			strSQL = strSQL & "FROM(" & vbCRLF
			strSQL = strSQL & "SELECT Devices.ID, Devices.LGTag, Tags.Tag" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Devices.Active=True AND Devices.Deleted=False AND Assigned=True AND Model Like '%iPad%')" & vbCRLF
			strSQL = strSQL & "WHERE (Tags.Tag='" & intIndex & "' OR Tags.Tag='Spare')" & vbCRLF
			strSQL = strSQL & "GROUP BY Devices.LGTag)" & vbCRLF
			strSQL = strSQL & "WHERE TagCount=2"
			Set objLoanedSpares = Application("Connection").Execute(strSQL)

			'Get the number of devices issued as replacement
			strSQL = "SELECT Count(TagCount) AS CountofCount" & vbCRLF
			strSQL = strSQL & "FROM" & vbCRLF
			strSQL = strSQL & "(SELECT Count(Tags.Tag) AS TagCount,Devices.LGTag" & vbCRLF
			strSQL = strSQL & "FROM(" & vbCRLF
			strSQL = strSQL & "SELECT Devices.ID, Devices.LGTag, Tags.Tag" & vbCRLF
			strSQL = strSQL & "FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag" & vbCRLF
			strSQL = strSQL & "WHERE Devices.Active=True AND Devices.Deleted=False AND Assigned=True AND Model Like '%iPad%')" & vbCRLF
			strSQL = strSQL & "WHERE (Tags.Tag='" & intIndex & "' OR Tags.Tag='Spare' OR Tags.Tag='Replacement')" & vbCRLF
			strSQL = strSQL & "GROUP BY Devices.LGTag)" & vbCRLF
			strSQL = strSQL & "WHERE TagCount=3"
			Set objReplacements = Application("Connection").Execute(strSQL)

			If objAvailableSpares.EOF Then
				intAvailableSpares = 0
			Else
				intAvailableSpares = objAvailableSpares(0)
			End If

			If Not objLoanedSpares.EOF Then
				intLoanedSpares = objLoanedSpares(0)
			Else
				intLoanedSpares = 0
			End If

			If Not objReplacements.EOF Then
				intReplacements = objReplacements(0)
			Else
				intReplacements = 0
			End If

			If intAvailableSpares <= 1 Then
				strAvailableSparesLabel = ""
			Else
				strAvailableSparesLabel = intAvailableSpares
			End If

			If intReplacements = 0 Then
				intReplacementLabel = ""
			ElseIf intReplacements = 1 Then
				intReplacementLabel = ""
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			Else
				intReplacementLabel = intReplacements
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			End If

			If intLoanedSpares <= 1 Then
				strLoanedSpareLabel = ""
			Else
				strLoanedSpareLabel = intLoanedSpares
			End If

			If intAvailableSpares + intLoanedSpares > intHighestValue Then
            intHighestValue = intAvailableSpares + intLoanedSpares
         End If

			strSpareDevicesData = strSpareDevicesData & "['" & intIndex & "'," & intAvailableSpares & _
			",'" & strAvailableSparesLabel & "'," & intLoanedSpares & _
			",'" & strLoanedSpareLabel & "'," & intReplacements & ",'" & intReplacementLabel & "'],"

		End If

	Next
	strSpareDevicesData = Left(strSpareDevicesData,Len(strSpareDevicesData) - 1)%>

	google.setOnLoadCallback(drawSpareiPads);

	function drawSpareiPads() {

		var data = google.visualization.arrayToDataTable([
			<%=strSpareDevicesData%>
		]);

		var options = {
			titlePosition: 'none',
			chartArea:{left:40, width:'90%', height:'85%'},
			legend: 'none',
			isStacked: true,
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {viewWindow: {max : <%=intHighestValue*1.1%>},minValue: 0},
			vAxis: {title: ''}

		};

		var chart = new google.visualization.BarChart(document.getElementById('spareiPads'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			var deviceYear = data.getValue(chart.getSelection()[0].row, 0)
			window.open('devices.asp?Model=iPad&Tags=Spare,' + deviceYear + '&View=Table','_self');
		}

	}

<%End Sub %>

<%Sub SpareiPadsByTypeJavaScript

	Dim strSQL, objSpareDevices, strSpareDevicesData, intDeviceYear, strDeviceName, strDeviceSite, intYears, intDeviceCount, objOldestDevice, datOldestDevice
	Dim intIndex, objSpares, intHighestValue, strModel, objLoanedSpares, intAvailableSpares, intLoanedSpares
	Dim strAvailableSparesLabel, strLoanedSpareLabel, objReplacements, intReplacements, intReplacementLabel, objAvailableSpares

	strSpareDevicesData = "['Device','Available',{ role: 'annotation' },'Loaned Out',{ role: 'annotation' },'Issued as Replacements',{ role: 'annotation' }],"
	intHighestValue = 0
	intAvailableSpares = 0
	intLoanedSpares = 0

	'Get the number of available spares
	strSQL = "SELECT Model, Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
	strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND (Model Like '%iPad%') "
	strSQL = strSQL & "GROUP BY Model"
	Set objSpares = Application("Connection").Execute(strSQL)

	'Get the number of available spares
	strSQL = "SELECT Model, Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
	strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND (Model Like '%iPad%') AND Assigned=False "
	strSQL = strSQL & "GROUP BY Model"
	Set objAvailableSpares = Application("Connection").Execute(strSQL)

	'Get the number of loaned spares
	strSQL = "SELECT Model, Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
	strSQL = strSQL & "WHERE Tag='Spare' AND Active=True AND Assigned=True "
	strSQL = strSQL & "GROUP BY Model"
	Set objLoanedSpares = Application("Connection").Execute(strSQL)

	'Get the number of devices issued as replacement
	strSQL = "SELECT Model, Count(Devices.ID) AS CountofID FROM Devices INNER JOIN Tags ON Devices.LGTag = Tags.LGTag "
	strSQL = strSQL & "WHERE Tag='Replacement' AND Active=True AND Assigned=True "
	strSQL = strSQL & "GROUP BY Model"
	Set objReplacements = Application("Connection").Execute(strSQL)

	If Not objSpares.EOF Then
		Do Until objSpares.EOF

			strModel = objSpares(0)
			intAvailableSpares = 0
			intLoanedSpares = 0
			intReplacements = 0

			If objAvailableSpares.EOF Then
				intAvailableSpares = 0
			Else
				Do Until objAvailableSpares.EOF
					If strModel = objAvailableSpares(0) Then
						intAvailableSpares = objAvailableSpares(1)
					End If
					objAvailableSpares.MoveNext
				Loop
				objAvailableSpares.MoveFirst
			End If

			If objLoanedSpares.EOF Then
				intLoanedSpares = 0
			Else
				Do Until objLoanedSpares.EOF
					If strModel = objLoanedSpares(0) Then
						intLoanedSpares = objLoanedSpares(1)
					End If
					objLoanedSpares.MoveNext
				Loop
				objLoanedSpares.MoveFirst
			End If

			If objReplacements.EOF Then
				intReplacements = 0
			Else
				Do Until objReplacements.EOF
					If strModel = objReplacements(0) Then
						intReplacements = objReplacements(1)
					End If
					objReplacements.MoveNext
				Loop
				objReplacements.MoveFirst
			End If

			If intAvailableSpares <= 1 Then
				strAvailableSparesLabel = ""
			Else
				strAvailableSparesLabel = intAvailableSpares
			End If

			If intReplacements = 0 Then
				intReplacementLabel = ""
			ElseIf intReplacements = 1 Then
				intReplacementLabel = ""
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			Else
				intReplacementLabel = intReplacements
				intLoanedSpares = intLoanedSpares - intReplacements
				strLoanedSpareLabel = intLoanedSpares
			End If

			If intLoanedSpares <= 1 Then
				strLoanedSpareLabel = ""
			Else
				strLoanedSpareLabel = intLoanedSpares
			End If

			If intAvailableSpares + intLoanedSpares > intHighestValue Then
				intHighestValue = intAvailableSpares + intLoanedSpares
			End If

			If intLoanedSpares = "" Then
				intLoanedSpares = 0
			End If

			If intReplacements = "" Then
				intReplacements = 0
			End If

			strSpareDevicesData = strSpareDevicesData & "['" & strModel & _
			intIndex & "'," & intAvailableSpares & ",'" & strAvailableSparesLabel & "'," & intLoanedSpares & _
			",'" & strLoanedSpareLabel & "'," & intReplacements & ",'" & intReplacementLabel & "'],"
			objSpares.MoveNext


		Loop
	End If

	strSpareDevicesData = Left(strSpareDevicesData,Len(strSpareDevicesData) - 1)%>

	google.setOnLoadCallback(drawSpareiPads);

	function drawSpareiPads() {

		var data = google.visualization.arrayToDataTable([
			<%=strSpareDevicesData%>
		]);

		var options = {
			titlePosition: 'none',
			chartArea:{left:85, width:'90%', height:'85%'},
			legend: 'none',
			isStacked: true,
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			hAxis: {viewWindow: {max : <%=intHighestValue*1.2%>},minValue: 0},
			vAxis: {title: ''}
		};

		var chart = new google.visualization.BarChart(document.getElementById('spareiPads'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			var dataFromChart = data.getValue(chart.getSelection()[0].row, 0).split('-')

			var deviceModel = dataFromChart[0].trim();

			window.open('devices.asp?Model=' + deviceModel + '&Tags=Spare&View=Table','_self');
		}

	}

<%End Sub %>

<%Sub EventCategoriesJavaScript

	Dim strSQL, objOldestDevice, datOldestDevice, intYears, intIndex, objSpares, intDeviceCount
	Dim datStartDate, datEndDate, objEvents, strEventData

	datStartDate = GetStartOfFiscalYear(Date)
	datEndDate = Date

	strSQL = "SELECT Category, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
	strSQL = strSQL & "GROUP BY Category" & vbCRLF 'ORDER BY Count(ID) DESC
	Set objEvents = Application("Connection").Execute(strSQL)

	strEventData = "['Category','Count',{ role: 'annotation' } ],"


	If Not objEvents.EOF Then
		Do Until objEvents.EOF
			strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "'],"
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
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			is3D: 'true',
			pieSliceText: 'value',
			hAxis: {title: '', minValue: 0},
			vAxis: {title: ''}
		};

		var chart = new google.visualization.PieChart(document.getElementById('eventCategories'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			window.open('events.asp?Category=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + data.getValue(chart.getSelection()[0].row, 2) + '&Complete=All&View=Table','_self');
		}

	}

<%End Sub %>

<%Sub EventTypesJavaScript

	Dim strSQL, objOldestDevice, datOldestDevice, intYears, intIndex, objSpares, intDeviceCount
	Dim datStartDate, datEndDate, objEvents, strEventData

	datStartDate = GetStartOfFiscalYear(Date)
	datEndDate = Date

	strSQL = "SELECT Type, Count(ID) AS CountOfID" & vbCRLF
	strSQL = strSQL & "FROM Events" & vbCRLF
	strSQL = strSQL & "WHERE EventDate>=#" & datStartDate & "# AND EventDate<=#" & datEndDate & "#" & vbCRLF
	strSQL = strSQL & "GROUP BY Type" & vbCRLF 'ORDER BY Count(ID) DESC
	Set objEvents = Application("Connection").Execute(strSQL)

	strEventData = "['Category','Count',{ role: 'annotation' } ],"


	If Not objEvents.EOF Then
		Do Until objEvents.EOF
			strEventData = strEventData & "['" & objEvents(0) & "', " & objEvents(1) & ",'" & datStartDate & "'],"
			objEvents.MoveNext
		Loop
	End If%>

	google.setOnLoadCallback(drawEventTypes);

	function drawEventTypes() {

		var data = google.visualization.arrayToDataTable([
			<%=strEventData%>
		]);

		var options = {
			titlePosition: 'none',
			chartArea:{width:'90%', height:'85%'},
			//animation: {startup: 'true', duration: 1000, easing: 'out'},
			is3D: 'true',
			pieSliceText: 'value',
			hAxis: {title: '', minValue: 0},
			vAxis: {title: ''}
		};

		var chart = new google.visualization.PieChart(document.getElementById('eventTypes'));
		chart.draw(data, options);

		google.visualization.events.addListener(chart, 'select', selectHandler);

		function selectHandler(e) {
			window.open('events.asp?EventType=' + data.getValue(chart.getSelection()[0].row, 0) + '&StartDate=' + data.getValue(chart.getSelection()[0].row, 2) + '&Complete=All&View=Table','_self');
		}

	}

<%End Sub %>

<%Sub LookupDevice

   Dim strSQLWhere, strSQL, objDeviceLookup, intDeviceCount

   'Get the variables from the URL or form and fix it
   intTag = Request.Form("Tag")
   If intTag = "" then
      intTag = Request.QueryString("Tag")
   End If
   If Not Application("UseLeadingZeros") Then
		If IsNumeric(intTag) Then
			intTag = Int(intTag)
		Else
			If Left(intTag,4) = "TECH" Then
		   	intTag = Replace(intTag,"TECH","")
				If IsNumeric(intTag) Then
		      	intTag = Int(intTag)
			  	End If
			End If
		End If
	End If

	If intTag = 0 Then
		intTag = ""
	End If

   If intTag = "" Then
      strDeviceMessage = "<div Class=""Error"">Device not found</div>"

   Else

      If intTag <> "" Then
         strSQLWhere = strSQLWhere & "WHERE Devices.LGTag='" & Replace(intTag,"'","''") & "' AND "
      End If

      strSQLWhere = strSQLWhere & "Devices.Deleted=False"

      strSQL = "SELECT ID, LGTag FROM Devices " & strSQLWhere
      Set objDeviceLookup = Application("Connection").Execute(strSQL)

      If Not objDeviceLookup.EOF Then
      	Response.Redirect("device.asp?Tag=" & intTag)
      Else
      	strDeviceMessage = "<div Class=""Error"">Device not found</div>"
      End If

   End If

End Sub%>

<%Sub LookupUser

   Dim  strSQL, objUserLookup, intUserCount, strURL, strSQLWhere

   strFirstName = Request.Form("FirstName")
   strLastName = Request.Form("LastName")
   intUserCount = 0

   If strFirstName = "" And strLastName = "" Then
      strUserMessage = "<div Class=""Error"">User not found</div>"

   Else

      strSQLWhere = "WHERE "
      strSQLWhere = strSQLWhere & "People.Active=True AND "
      If strFirstName <> "" Then
         strSQLWhere = strSQLWhere & "People.FirstName Like '%" & Replace(strFirstName,"'","''") & "%' AND "
         strURL = strURL & "&FirstName=" & strFirstName
      End If
      If strLastName <> "" Then
         strSQLWhere = strSQLWhere & "People.LastName Like '%" & Replace(strLastName,"'","''") & "%' AND "
         strURL = strURL & "&LastName=" & strLastName
      End If

      strSQLWhere = strSQLWhere & "People.Deleted=False"

   	strSQL = "SELECT ID, UserName FROM People " & strSQLWhere
      Set objUserLookup = Application("Connection").Execute(strSQL)

      If Not objUserLookup.EOF Then
         Do Until objUserLookup.EOF
            intUserCount = intUserCount + 1
            objUserLookup.MoveNext
         Loop
         objUserLookup.MoveFirst
      End If

      If strURL <> "" Then
         strURL = Right(strURL, Len(strURL) - 1)
      End If

      Select Case intUserCount
         Case 0
            strUserMessage = "<div Class=""Error"">User not found</div>"
         Case 1
            Response.Redirect("user.asp?UserName=" & objUserLookup(1))
         Case Else
            Response.Redirect("users.asp?" & strURL)
      End Select

   End If

End Sub%>

<%Sub LookupEvent

	Dim strSQL, objEventLookup, intEventCount

	intEventNumber = Request.Form("EventNumber")

 	If intEventNumber = "" Then
 		strEventMessage = "<div Class=""Error"">No events found</div>"

 	Else

		strSQL = "SELECT ID, LGTag FROM Events WHERE ID=" & intEventNumber
		Set objEventLookup = Application("Connection").Execute(strSQL)

		If Not objEventLookup.EOF Then
			Response.Redirect("events.asp?EventNumber=" & intEventNumber)
		Else
			strEventMessage = "<div Class=""Error"">No events found</div>"
		End If

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

<%Function GetGrade(intYear)

	Dim datToday, intMonth, intCurrentYear

   datToday = Date
   intMonth = DatePart("m",datToday)
   intCurrentYear = Right(DatePart("yyyy",datToday),2)
   intYear = Right(intYear,2)

   If intMonth >= 7 And intMonth <= 12 Then
      intCurrentYear = intCurrentYear + 1
   End If

   GetGrade = 12 - (intYear - intCurrentYear)

   If GetGrade = 0 Then
      GetGrade = "K"
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
