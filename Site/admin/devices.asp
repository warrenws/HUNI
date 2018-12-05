<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/22/15
'Last Updated 1/14/18

'This page shows a list of devices as a result of a search.

Option Explicit

'On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser
Dim objDeviceList, strSite, intDeviceCount, strView, intDeviceYear, strModel, strTags, strRoom
Dim strAssigned, intTagCount, strStatus, strSearchMessage, strBackLink, strCardType, strColumns
Dim strNotes, strMake, strDeviceType, objLastNames, objUserList

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

	Dim strSQL, arrTags, intIndex, strSQLWhere, strCurrentTag, strCurrentPage

	'Get the variables from the URL
	strSite = Request.QueryString("DeviceSite")
	intDeviceYear = Request.QueryString("Year")
	strView = Request.QueryString("View")
	strDeviceType = Request.QueryString("DeviceType")
	strMake = Request.QueryString("Make")
	strModel = Request.QueryString("Model")
	strTags = Request.QueryString("Tags")
	strRoom = Request.QueryString("Room")
	strAssigned = Request.QueryString("Assigned")
	strNotes = Request.QueryString("DeviceNotes")
	strStatus = Request.QueryString("DeviceStatus")
	strBackLink = BackLink
	strCurrentPage = LCase(Right(Request.ServerVariables("URL"),Len(Request.ServerVariables("URL")) - InStrRev(Request.ServerVariables("URL"),"/")))
	
	'If nothing was submitted send them back to the index page
   If strSite = "" And intDeviceYear = "" And strMake = "" And strModel = "" And strRoom = "" And strTags = "" And strAssigned = "" And strStatus = "" And strNotes = "" And strDeviceType = "" Then
   	If Request.QueryString("Source") <> "" Then
   		Response.Redirect("search.asp?Error=NoDevicesFound")
   	ElseIf strCurrentPage = "devices.asp" Then
   		Response.Redirect("index.asp?Error=NoDevicesFound")
   	Else
   		Response.Redirect(Request.QueryString("Source") & "?Error=NoDevicesFound")
   	End If
   End If

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Save"
      	SaveSearch
         
   End Select
   
   'Get the list of devices
	strSQLWhere = "WHERE "
	
	If strDeviceType <> "" Then
   	strSQLWhere = strSQLWhere & "Devices.DeviceType='" & Replace(strDeviceType,"'","''") & "' AND "
   End If
	
	If strMake <> "" Then
   	strSQLWhere = strSQLWhere & "Devices.Manufacturer Like '%" & Replace(strMake,"'","''") & "%' AND "
   End If
	
	If strModel <> "" Then
   	strSQLWhere = strSQLWhere & "Devices.Model Like '%" & Replace(strModel,"'","''") & "%' AND "
   End If
   
   If strRoom <> "" Then
   	strSQLWhere = strSQLWhere & "Devices.Room='" & Replace(strRoom,"'","''") & "' AND "
		
		strSQL = "SELECT ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,Active,Warning,Loaned,PWord,AUP,Site,Notes,PWordLastSet,PhoneNumber,RoomNumber,Description " & _
   		"FROM People WHERE Active=True AND RoomNumber='" & Replace(strRoom,"'","''") & "'"

		If strSite <> "" Then
			strSQL = strSQL & " AND Site='" & Replace(strSite,"'","''") & "'"
		End If    	
   	
   	Set objUserList = Application("Connection").Execute(strSQL)
   
   Else
   
   	strSQL = "SELECT ID,FirstName,LastName,UserName,StudentID,ClassOf,HomeRoom,Role,DeviceCount,Active,Warning,Loaned,PWord,AUP,Site,Notes,PWordLastSet,PhoneNumber,RoomNumber,Description " & _
   		"FROM People WHERE RoomNumber='NO ROOM NUMBER GIVEN'"
   	Set objUserList = Application("Connection").Execute(strSQL)
   
   End If
	
	If strSite <> "" Then
      strSQLWhere = strSQLWhere & "Devices.Site = '" & strSite & "' AND "
   End If
   
   If strNotes <> "" Then
   	strSQLWhere = strSQLWhere & "Devices.Notes Like '%" & Replace(strNotes,"'","''") & "%' AND "
   End If
   
   If intDeviceYear <> "" Then
		strSQLWhere = strSQLWhere & _
		"Devices.DatePurchased>=#" & DateAdd("yyyy",intDeviceYear * -1,Date) & "# AND " & _
		"Devices.DatePurchased<=#" & DateAdd("yyyy",(intDeviceYear -1) * -1,Date) & "# AND "
   End If
   
   Select Case strAssigned
		Case "Yes"
			strSQLWhere = strSQLWhere & "Assigned=True AND "
		Case "No"
			strSQLWhere = strSQLWhere & "Assigned=False AND "
   End Select
   
   Select Case strStatus
		Case "Enabled"
			strSQLWhere = strSQLWhere & "Devices.Active=True AND "
		Case "Disabled"
			strSQLWhere = strSQLWhere & "Devices.Active=False AND "
		Case "All"
		Case Else
			strSQLWhere = strSQLWhere & "Devices.Active=True AND "
   End Select
   
   strSQLWhere = strSQLWhere & "Devices.Deleted=False AND "
   
	If strSQLWhere <> "WHERE " Then
		strSQLWhere = Left(strSQLWhere,Len(strSQLWhere) - 5)
	Else
		strSQLWhere = ""
	End If
	
	intTagCount = 0
	If strTags = "" Then
		strSQL = "SELECT ID,LGTag,Manufacturer,Model,SerialNumber,Room,DatePurchased,Site,Active,FirstName,LastName,UserName,BOCESTag," & _
			"HasInsurance,MACAddress,AppleID,Notes,DeviceType,InternalIP,ExternalIP,LastUser,OSVersion,ComputerName,LastCheckInDate,LastCheckInTime,HasEvent  FROM Devices "
			strSQL = strSQL & strSQLWhere & " ORDER BY Devices.LGTag"
   Else
		strSQL = "SELECT Devices.ID,Devices.LGTag,Devices.Manufacturer,Devices.Model,Devices.SerialNumber,Devices.Room," & _
			"Devices.DatePurchased,Devices.Site,Devices.Active,Tags.Tag,FirstName,LastName,UserName,BOCESTag,HasInsurance," & _
			"MACAddress,AppleID,Devices.Notes,DeviceType,InternalIP,ExternalIP,LastUser,OSVersion,ComputerName,LastCheckInDate,LastCheckInTime,HasEvent FROM Tags INNER JOIN " & _
			"Devices ON Tags.LGTag = Devices.LGTag "
	
		arrTags = Split(strTags,",")
		intTagCount = UBound(arrTags) + 1
		
		strSQL = "SELECT COUNT(Tags.Tag) AS TagCount,Devices.LGTag,Devices.Manufacturer,Devices.Model,Devices.SerialNumber," & _
			"Devices.Room,Devices.DatePurchased,Devices.Site,Devices.Active,FirstName,LastName,UserName,BOCESTag,HasInsurance," & _
			"MACAddress,AppleID,Devices.Notes,DeviceType,InternalIP,ExternalIP,LastUser,OSVersion,ComputerName,LastCheckInDate,LastCheckInTime,HasEvent " & _
			"FROM(" & strSQL & strSQLWhere & ") WHERE "

		For intIndex = 0 to UBound(arrTags)
			strSQL = strSQL & "Tags.Tag='" & Trim(Replace(arrTags(intIndex),"'","''")) & "' OR "
		Next     
		strSQL = Left(strSQL,Len(strSQL) - 4)
		strSQL = strSQL & " GROUP BY Devices.LGTag,Devices.Manufacturer,Devices.Model,Devices.SerialNumber,Devices.Room," & _
			"Devices.DatePurchased,Devices.Site,Devices.Active,FirstName,LastName,UserName,BOCESTag,HasInsurance," & _
			"MACAddress,AppleID,Devices.Notes,DeviceType,InternalIP,ExternalIP,LastUser,OSVersion,ComputerName,LastCheckInDate,LastCheckInTime,HasEvent ORDER BY Devices.LGTag"
			  
	End If

	Set objDeviceList = Application("Connection").Execute(strSQL)

   'If no user is found send them back to the index page.
   If objDeviceList.EOF Then
   	If Request.QueryString("Source") = "" Then
   		Response.Redirect("search.asp?Error=NoDevicesFound")
   	Else
   		Response.Redirect(Request.QueryString("Source") & "?Error=NoDevicesFound")
   	End If
	Else
		Do Until objDeviceList.EOF
			If intTagCount > 0 Then
				If objDeviceList(0) = intTagCount Then
					intDeviceCount = intDeviceCount + 1
				End If
			Else   
				intDeviceCount = intDeviceCount + 1
			End If
			objDeviceList.MoveNext
		Loop
		objDeviceList.MoveFirst
	End If

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
			<% End If %>	
				
    			var table = $('#ListView').DataTable( {
    				paging: false,
    				"info": false,
    				"autoWidth": false,
    				"order": [[ 1, "asc" ]],
    				dom: 'Bfrtip',
    				// stateSave: true,
    				buttons: [
						{
							extend: 'colvis',
							text: 'Show/Hide Columns'
						}
				<% If Not IsMobile Then %>		
						,
						{
							extend: 'csvHtml5',
							text: 'Download CSV',
							title: 'Inventory - Devices'
						}
				<% End If %>
        			]
        		
    			});
    			
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
    			
    	<% If IsMobile Then %>	
    			table.columns([0,2,3,4,5,8,9,10,11,12,13,14,15,16,17]).visible(false);	
    	<% Else %>		
				table.columns([0,2,3,4,8,10,12,13,14,16,17]).visible(false);
		<% End If %>
				$('#body').show();

    		} );
    	</script>
   </head>

   <body class="<%=strSiteVersion%>" id="body" style="display:none;">
   
      <div class="Header"><%=Application("SiteName")%> (<%=intDeviceCount%>)</div>
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

<% If Not objDeviceList.EOF Then
	
		Select Case LCase(strView)
			Case "table"
				ShowDeviceTable
			Case "card"
				ShowDeviceCards
			Case Else
			
				If IsMobile Then
					If LCase(Application("DefaultViewMobile")) = "table" Then
						If intDeviceCount < Application("CardThreshold") Then
							ShowDeviceCards
						Else
							ShowDeviceTable
						End If
					Else
						ShowDeviceCards
					End If
				Else

					If LCase(Application("DefaultView")) = "table" Then
						If intDeviceCount < Application("CardThreshold") Then
							ShowDeviceCards
						Else
							ShowDeviceTable
						End If
					Else
						ShowDeviceCards
					End If
					
				End If
		End Select
		ShowChart
      SaveAsSearch
   End If %>
	<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
   </body>

   </html>

<%End Sub%>

<%Sub ShowDeviceCards

	If IsMobile Then %>
		<div class="ViewButtonMobile">
<% Else %>
		<div class="ViewButton">
<% End If %>
		<a href="<%=SwitchView("Table")%>"><img src="../images/table.png" title="Table View" height="32" width="32"/></a>
	</div>
	
	<div class="center"><%=FilterBar%></div>

<%	If Not objUserList.EOF Then
		ShowUserCards
   Else %>
		<div Class="<%=strColumns%>">
<%	End If %>	
	
<%	Do Until objDeviceList.EOF 
		If intTagCount > 0 Then
			If objDeviceList(0) = intTagCount Then
				DrawCard
			End If
		Else
			DrawCard
		End If
		objDeviceList.MoveNext
   Loop  %>
	</div>
<%End Sub%>

<%Sub DrawCard

	Dim objFSO, objUserList, strSQL, objTagList, strDeviceInfo
	
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	
	strDeviceInfo = "Name: " & objDeviceList(22) & " &#013 "
	strDeviceInfo = strDeviceInfo & "Last User: " & objDeviceList(20) & " &#013 "
	strDeviceInfo = strDeviceInfo & "OS Version: " & objDeviceList(21) & " &#013 "
	strDeviceInfo = strDeviceInfo & "Last Checkin: " & objDeviceList(23) & " - " & objDeviceList(24)
	
	%>

<% If objDeviceList(25) Then
		strCardType = "WarningCard"
	ElseIf objDeviceList(8) Then
		strCardType = "NormalCard"
	Else
		strCardType = "DisabledCard"
	End If %>
	<div class="Card <%=strCardType%>">
		<div class="CardTitle">

		<% If objDeviceList(22) <> "" Then %>
			<% If Application("MunkiReportServer") = "" Then %>
					<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strDeviceInfo%>"  />&nbsp;</div>
			<% Else %>
					<a href="<%=Application("MunkiReportServer")%>/index.php?/clients/detail/<%=objDeviceList(4)%>" target="_blank">
						<div class="Button"><image src="../images/info.png" width="20" height="20" title="<%=strDeviceInfo%>"  />&nbsp;</div>
					</a>
			<% End If %>
		<% End If %>
		
		<% If objDeviceList(13) Then %>
				<image src="../images/yes.png" width="15" height="15" title="Insured" />
		<% End If %>
			<a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>">Asset Tag <%=objDeviceList(1)%></a>
		</div>
		<div Class="ImageSectionInCard">		
		<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(3)," ","") & ".png") Then %>      
				<a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>"> 
					<img class="PhotoCard" src="../images/devices/<%=Replace(objDeviceList(3)," ","")%>.png" width="96" />
				</a>
		<% Else %>
				<a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>"> 
					<img class="PhotoCard" src="../images/devices/missing.png" width="96" />
				</a>
		<% End If %>
		
	</div>		
		<div Class="RightOfImageInCard">
			<div>
				<div Class="PhotoCardColumn1">Make: </div>
				<div Class="PhotoCardColumn2"><%=objDeviceList(2)%></div>
			</div>
			<div>
				<div Class="PhotoCardColumn1">Model: </div>
				<div Class="PhotoCardColumn2"><%=objDeviceList(3)%></div>
			</div>
			<div>
				<div Class="PhotoCardColumn1">Serial: </div>
				<div Class="PhotoCardColumn2">
			<% Select Case objDeviceList(2)
					Case "Apple" %>
						<a href="https://checkcoverage.apple.com/us/en/?sn=<%=objDeviceList(4)%>" target="_blank"><%=objDeviceList(4)%></a>
				<% Case "Dell" %>
						<a href="http://www.dell.com/support/home/us/en/19/product-support/servicetag/<%=objDeviceList(4)%>" target="_blank"><%=objDeviceList(4)%></a>
				<% Case Else %>
						<%=objDeviceList(4)%>
			 <% End Select %> 	
				</div>
			</div>
		<% If objDeviceList(6) <> "" Then %>
				<div>
					<div Class="CardMerged">
						Purchased: <%=ShortenDate(objDeviceList(6))%> - <a href="devices.asp?Year=<%=GetAge(objDeviceList(6))%>">Year <%=GetAge(objDeviceList(6))%></a>
					</div>
				</div>
		<% End If %>
		<% If objDeviceList(5) <> "" Then %>
			<div>
				<div Class="PhotoCardColumn1">Room: </div>
				<div Class="PhotoCardColumn2">
					<a href="devices.asp?Room=<%=objDeviceList(5)%>&DeviceSite=<%=objDeviceList(7)%>&View=Card"><%=objDeviceList(5)%></a>
				</div>
			</div>
		<% End If %>
		<% strSQL = "SELECT Tag FROM Tags WHERE LGTag='" & objDeviceList(1) & "' ORDER BY Tag"
			Set objTagList = Application("Connection").Execute(strSQL)
		
			If Not objTagList.EOF Then %>
				<div>
					<div Class="PhotoCardColumn1">Tags: </div>
					<div Class="PhotoCardColumn2">
			<% Do Until objTagList.EOF %>	
					<a href="devices.asp?Tags=<%=objTagList(0)%>"><%=objTagList(0)%></a>
			<% 	objTagList.MoveNext
				Loop %>
					</div>
				</div>
		<% End If %>

		<% If objDeviceList(11) <> "" Then %>
				<div>
					<div Class="CardMerged">Assigned To: 
						<a href="user.asp?UserName=<%=objDeviceList(11)%><%=strBackLink%>"><%=objDeviceList(9)%>&nbsp;<%=objDeviceList(10)%></a>
					</div>
				</div>
		<%	End If %>
	
		<% If objDeviceList(16) <> "" Then %>
				<div>Device Notes: <%=Replace(objDeviceList(16),vbCRLF,"<br />")%> </div>
		<% End If %>
		
		<%	If objDeviceList(18) <> "" Then %>
			<br />
		
			<% DrawIcon "Remote",0,""
			End If %>
		</div>
	</div>
	 
<%End Sub %>

<%Sub ShowDeviceTable

		If IsMobile Then %>
		<div class="ViewButtonMobile">
<% Else %>
		<div class="ViewButton">
<% End If %>
		<a href="<%=SwitchView("Card")%>"><img src="../images/card.png" title="Card View" height="32" width="32"/></a>
	</div>
	
	<div class="center"><%=Replace(FilterBar,"?","?View=Table&")%></div>

	<div>
		<table align="center" Class="ListView" id="ListView">
			<thead>
			<th>Photo</th>
			<th>Asset Tag</th>
			<th>BOCES Tag</th>
			<th>Serial Number</th>
			<th>Device Type</th>
			<th>Make</th>
			<th>Model</th>
			<th>Assigned To</th>
			<th>Site</th>
			<th>Room</th>
			<th>Purchased</th>
			<th>Year</th>	
			<th>Insured</th>
			<th>MAC Address</th>
			<th>Apple ID</th>
			<th>Tags</th>
			<th>Device Notes</th>
			<th>Last Check In</th>
			</thead>
			<tbody>
	<% Do Until objDeviceList.EOF 
		If intTagCount > 0 Then
			If objDeviceList(0) = intTagCount Then
				DrawRow
			End If
		Else
			DrawRow
		End If
		objDeviceList.MoveNext
   Loop  %> 
			</tbody>
  	 	</table>	
	</div>
<%End Sub %>

<%Sub DrawRow 

	Dim objTagList, strSQL, objUserList, strRowClass, objFSO

	If objDeviceList(25) Then
		strRowClass = " Class=""Warning"""
	ElseIf objDeviceList(8) Then
		strRowClass = ""
	Else 
		strRowClass = " Class=""Disabled"""
	End If 
	
	Set objFSO = CreateObject("Scripting.FileSystemObject") %>
	
	<tr<%=strRowClass%>>
	
	<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(3)," ","") & ".png") Then %>      
			<td <%=strRowClass%> id="center" width="1px">
				<a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>"> 
			<% If InStr(LCase(objDeviceList(3)),"ipad") Then %>
					<img src="../images/devices/<%=Replace(objDeviceList(3)," ","")%>.png" width="52" />
			<% Else %>
					<img src="../images/devices/<%=Replace(objDeviceList(3)," ","")%>.png" width="72" />
			<% End If %>
				</a>
			</td>
	<% Else %>
			<td <%=strRowClass%> id="center" width="1px">
				<a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>"> 
					<img class="PhotoCard" src="../images/devices/missing.png" width="72" />
				</a>
			</td>
	<% End If %>
	
		<td <%=strRowClass%> id="center"><a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>"><%=objDeviceList(1)%></a></td>
		<td <%=strRowClass%>><%=objDeviceList(12)%></td>
		<% Select Case objDeviceList(2)
				Case "Apple" %>
					<td <%=strRowClass%>><a href="https://checkcoverage.apple.com/us/en/?sn=<%=objDeviceList(4)%>" target="_blank"><%=objDeviceList(4)%></a></td>
			<% Case "Dell" %>
					<td <%=strRowClass%>><a href="http://www.dell.com/support/home/us/en/19/product-support/servicetag/<%=objDeviceList(4)%>" target="_blank"><%=objDeviceList(4)%></a></td>
			<% Case Else %>
					<td <%=strRowClass%>><%=objDeviceList(4)%></td>
		 <% End Select %> 
		<td <%=strRowClass%>><%=objDeviceList(17)%></td>
		<td <%=strRowClass%>><%=objDeviceList(2)%></td>
		<td <%=strRowClass%>><%=objDeviceList(3)%></td>
		
<% If objDeviceList(11) <> "" Then %>
		<td <%=strRowClass%>><a href="user.asp?UserName=<%=objDeviceList(11)%><%=strBackLink%>"><%=objDeviceList(10)%>, <%=objDeviceList(9)%></a></td>
<% Else %>
		<td <%=strRowClass%>></td>
<% End If %>	
		<td <%=strRowClass%> id="center"><%=objDeviceList(7)%></td>
		<td <%=strRowClass%>><a href="devices.asp?Room=<%=objDeviceList(5)%>&View=Table"><%=objDeviceList(5)%></a></td>
		<td <%=strRowClass%> id="center"><%=objDeviceList(6)%></td>
<% If objDeviceList(6) <> "" Then %>	
		<td id="center"><a href="devices.asp?Year=<%=GetAge(objDeviceList(6))%>&View=Table"><%=GetAge(objDeviceList(6))%></a></td>
<% Else %>
		<td <%=strRowClass%>>N/A</td>
<% End If %>
<% If objDeviceList(13) Then %>
		<td <%=strRowClass%> id="center">Yes</td>
<% Else %>
		<td <%=strRowClass%> id="center">No</td>
<% End If %>
		<td <%=strRowClass%>><%=objDeviceList(14)%></td>
		<td <%=strRowClass%>><%=objDeviceList(15)%></td>
		<td <%=strRowClass%>>
	<% strSQL = "SELECT Tag FROM Tags WHERE LGTag='" & objDeviceList(1) & "' ORDER BY Tag"
		Set objTagList = Application("Connection").Execute(strSQL)
	
		If Not objTagList.EOF Then %>
		<% Do Until objTagList.EOF %>	
				<a href="devices.asp?Tags=<%=objTagList(0)%>&View=Table"><%=objTagList(0)%></a>
		<% 	objTagList.MoveNext
			Loop %>
	<%	End If %>	
		</td>
	<% If NOT IsNull(objDeviceList(16)) Then %>
			<td <%=strRowClass%>><%=Replace(objDeviceList(16),vbCRLF,"<br />")%></td>
	<% Else %>
			<td <%=strRowClass%>><%=objDeviceList(16)%></td>
	<% End If%>	
		<td <%=strRowClass%> id="center"><%=objDeviceList(23)%></td>
	</tr>

<%End Sub %>

<%Sub JumpToDevice%>

	<div Class="HeaderCard">
		<form method="POST" action="search.asp">
		Asset tag: <input class="Card InputWidthSmall" type="text" name="SmartBox" id="LastNames" />
		<div class="Button"><input type="submit" value="Lookup Device" name="Submit" /></div>
		<input type="hidden" value="" name="BOCESTag" />
		</form>
	</div>

<%End Sub%>

<%Sub ShowUserCards 

	Dim arrPWordLastSet %>

	<div Class="<%=strColumns%>">
<%	If Not objUserList.EOF Then
		
		Dim objFSO, objDeviceList, strSQL, strDeviceList, intDaysRemaining
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		Do Until objUserList.EOF
		
			If objUserList(10) Then 
				strCardType = "WarningCard"
		 	ElseIf objUserList(11) Then 
				strCardType = "LoanedCard"
			ElseIf objUserList(9) Then 
				strCardType = "NormalCard"
		 	Else 
				strCardType = "DisabledCard"
			End If %>
			
			<div class="Card <%=strCardType%>">
				<div class="CardTitle">
				<% If objUserList(7) = "Student" Then %>
					<% If Application("ShowPasswords") Then %>
						<% If objUserList(13) Then %>
								<image src="../images/yes.png" width="15" height="15" title="AUP Signed" />
						<% Else %>
								<image src="../images/no.png" width="15" height="15" title="AUP Not Signed" />
						<% End If %>
					<% End If %>
				<% End If %>
					<a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>"><%=objUserList(1) & " " & objUserList(2)%></a>
				</div>
				<div Class="ImageSectionInCard">
				<% If objFSO.FileExists(Application("PhotoLocation") & "\" & objUserList(7) & "s\" & objUserList(4) & ".jpg") Then %>   
						<a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>">   
							<img class="PhotoCard" src="/photos/<%=objUserList(7)%>s/<%=objUserList(4)%>.jpg" title="<%=objUserList(4)%>" width="96" />
						</a>
				<% Else %>
						<a href="user.asp?UserName=<%=objUserList(3)%><%=strBackLink%>">
							<img class="PhotoCard" src="/photos/<%=objUserList(7)%>s/missing.png" title="<%=objUserList(4)%>" width="96" />
						</a>
				<% End If %>
			
				</div>
				<div Class="RightOfImageInCard">
				<div>
					<div Class="PhotoCardColumn1">Role: </div>
					<div Class="PhotoCardColumn2Long">
						<a href="users.asp?Role=<%=objUserList(5)%>"><%=GetRole(objUserList(5))%></a>
					</div>
				</div>
			<% If objUserList(7) = "Student" Then %>	
					<div>
						<div Class="PhotoCardColumn1">Guide: </div>
						<div Class="PhotoCardColumn2Long">
							<a href="users.asp?GuideRoom=<%=objUserList(6)%>"><%=objUserList(6)%></a>
						</div>
					</div>
				<%	If Application("ShowPasswords") Then %>
				   <div>
						<div Class="CardMerged">Username: <%=objUserList(3)%></div>
					</div>
				   <div>
						<div Class="CardMerged">Password: <%=objUserList(12)%></div>
					</div>
				<% End If
				Else 
					
					If Not IsNull(objUserList(16)) Then
						arrPWordLastSet = Split(objUserList(16)," ")
						If CDate(arrPWordLastSet(0)) > #1/1/80# Then 
					
							intDaysRemaining = DateDiff("d",Date(),DateAdd("d",Application("PasswordsExpire"),arrPWordLastSet(0))) %>
					
							<div Class="CardMerged">Password Changed: <%=ShortenDate(arrPWordLastSet(0))%></div>
						
						<% If intDaysRemaining > 10 Then %>
								<div Class="CardMerged">Days Remaining: <%=intDaysRemaining%></div>
						<% ElseIf intDaysRemaining >= -500 Then %>
								<div Class="CardMerged Error">Days Remaining: <%=intDaysRemaining%></div>
						<% End If 
						
						Else %>
						
							<div Class="CardMerged">Password Changed: ---</div>
							<div Class="CardMerged Error">Days Remaining: Expired</div>
							
					<%	End If %>
					
				<%	End If
				End If %>
				
			<% If objUserList(17) <> "" Then %>
				<div Class="CardMerged">Phone: <%=objUserList(17)%></div>
			<% End If %>
			
			<% If objUserList(8) > 0 Then
					
					strSQL = "SELECT Assignments.LGTag,Devices.Model" & vbCRLF
					strSQL = strSQL & "FROM Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag" & vbCRLF
					strSQL = strSQL & "WHERE Assignments.AssignedTo=" & objUserList(0) & " AND Assignments.Active=True"
					Set objDeviceList = Application("Connection").Execute(strSQL)
					
					If Not objDeviceList.EOF Then
						strDeviceList = "" %>
						<div Class="CardMerged">Assigned: 
					
					<%	Do Until objDeviceList.EOF
							strDeviceList = strDeviceList & "<a href=""device.asp?Tag=" & objDeviceList(0) & strBackLink & """>" & objDeviceList(1) & "</a>, "
							objDeviceList.MoveNext
						Loop 
	
						strDeviceList = Left(strDeviceList,Len(strDeviceList) - 2)%>
						<%=strDeviceList%>
						</div>
				<%	End If
			   End If %>
			   	</div>
			<% If objUserList(15) <> "" Then %>
					<div>User Notes: <%=Replace(objUserList(15),vbCRLF,"<br />")%> </div>
			<% End If %>
			</div>
			
   
      <% objUserList.MoveNext
      Loop %>

<% End If 
   
End Sub%>

<%Sub ShowChart

	If Application("LibreNMSServer") <> "" Then

		Dim strSQL, objChartList
	
		strSQL = "SELECT PortID FROM LibreNMS WHERE Room='" & Replace(strRoom,"'","''") & "'"
		If strSite <> "" Then
			strSQL = strSQL & " AND Site='" & Replace(strSite,"'","''") & "'"
		End If
		
		Set objChartList = Application("Connection").Execute(strSQL)
		
		If Not objChartList.EOF Then
			Do Until objChartList.EOF
				If Not IsMobile Then %>
					<div class="Col3Card NormalCard">
						<div class="Col3CardTitle">24 Hour AP Usage</div>
						<img src="<%=Application("LibreNMSServer")%>/graph.php?id=<%=objChartList(0)%>&type=port_bits&width=650&height=150&from=end-24h">
					</div>
			<% Else %>
				<div class="Card NormalCard">
					<div class="CardTitle">24 Hour AP Usage</div>
					<img src="<%=Application("LibreNMSServer")%>/graph.php?id=<%=objChartList(0)%>&type=port_bits&width=300&height=150&from=end-24h">
				</div>
			<% End If 
				objChartList.MoveNext
			Loop
		End If
	
	End If

End Sub%>

<%Function FilterBar
   
   If strSite <> "" Then
   	FilterBar = FilterBar & "Site = <a href=""devices.asp?DeviceSite=" & strSite & """>" & strSite & "</a> | "
   End If
   
   If intDeviceYear <> "" Then
   	FilterBar = FilterBar & "Year = <a href=""devices.asp?Year=" & intDeviceYear & """>" & intDeviceYear & "</a> | "
   End If
   
   If strDeviceType <> "" Then
   	FilterBar = FilterBar & "Device Type = <a href=""devices.asp?DeviceType=" & strDeviceType & """>" & strDeviceType & "</a> | "
   End If
   
   If strMake <> "" Then
   	FilterBar = FilterBar & "Make = <a href=""devices.asp?Make=" & strMake & """>" & strMake & "</a> | "
   End If
   
   If strModel <> "" Then
   	FilterBar = FilterBar & "Model = <a href=""devices.asp?Model=" & strModel & """>" & strModel & "</a> | "
   End If
   
   If strTags <> "" Then
   	FilterBar = FilterBar & "Tags = <a href=""devices.asp?Tags=" & strTags & """>" & strTags & "</a> | "
   End If
   
   If strRoom <> "" Then
   	FilterBar = FilterBar & "Room = <a href=""devices.asp?Room=" & strRoom & """>" & strRoom & "</a> | "
   End If
   
   If strAssigned <> "" Then
   	FilterBar = FilterBar & "Assigned = <a href=""devices.asp?Assigned=" & strAssigned & """>" & strAssigned & "</a> | "
   End If
   
   If strStatus <> "" Then
   	FilterBar = FilterBar & "Status = <a href=""devices.asp?DeviceStatus=" & strStatus & """>" & strStatus & "</a> | "
   End If
   
   If strNotes <> "" Then
   	FilterBar = FilterBar & "Notes = <a href=""devices.asp?DeviceNotes=" & strNotes & """>" & strNotes & "</a> | "
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

<%Sub DrawIcon(strIconType,intPosition,strURLData)

	Dim strPosition

	Select Case intPosition
		Case 0
			strPosition = "Left"
		Case 1
			strPosition = "One"
		Case 2
			strPosition = "Two"
		Case 3
			strPosition = "Three"
		Case 4
			strPosition = "Four"
	End Select

	Select Case strIconType
		Case "HelpDesk" %>
			<a href="<%=Application("HelpDeskURL")%>/index.asp?UserName=<%=strURLData%>" class="Button<%=strPosition%>">
				<image src="../images/helpdesk.png" width="20" height="20" title="Enter Help Desk Ticket" />
			</a>
	<%	Case "Back" %>
			<a href="<%=Request.QueryString("Page")%>?<%=Request.QueryString("Back")%>" class="Button<%=strPosition%>">
				<image src="../images/back.png" width="20" height="20" title="Return to Search Results"/>
			</a>
	<% Case "Log" %>
			<a href="" class="Button<%=strPosition%>" id="logToggle">
				<image src="../images/log.png" height="20" width="20" title="View Log">
			</a>
			
	<% Case "Assignments" %>	
			<a href="" class="Button<%=strPosition%>" id="assignmentsToggle">
				<image src="../images/assignment.png" height="20" width="20" title="View Old Assignments">
			</a>
	
	<% Case "Events" %>
			<a href="" class="Button<%=strPosition%>" id="eventsToggle">
				<image src="../images/event.png" height="20" width="20" title="View Events">
			</a>
			
	<% Case "ViewAll" %>
			<a href="" class="Button<%=strPosition%>" id="viewAllToggle">
				<image src="../images/viewall.png" height="20" width="20" title="View All">
			</a>
	
	<% Case "Remote" %>
			<div class="ButtonLeftDevices">
			<a href="vnc://<%=objDeviceList(18)%>:5900" class="ButtonLeft" >
				<image src="../images/remote.png" height="20" width="20" title="Remote Control">
			</a>
			<a href="ssh://admin@<%=objDeviceList(18)%>" class="ButtonLeft" >
				<image src="../images/ssh.png" height="20" width="20" title="SSH ">
			</a>
			</div>
			
<%	End Select

End Sub %>

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
      <% If strSearchMessage <> "" Then %>
   		<div>
   			<%=strSearchMessage%>
   		</div>
   <% End If %> 
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
				GetRole = "Kindergarten Student"
			Case 1
				GetRole = "1st Grade Student"
			Case 2
				GetRole = "2nd Grade Student"
			Case 3
				GetRole = "3rd Grade Student"
			Case 4
				GetRole = "4th Grade Student"   
			Case 5
				GetRole = "5th Grade Student"
			Case 6
				GetRole = "6th Grade Student"
			Case 7
				GetRole = "7th Grade Student"
			Case 8
				GetRole = "8th Grade Student"
			Case 9
				GetRole = "9th Grade Student"
			Case 10
				GetRole = "10th Grade Student"
			Case 11
				GetRole = "11th Grade Student"
			Case 12
				GetRole = "12th Grade Student"
			Case Else
				GetRole = "Graduated"
		End Select

	End If

End Function%>

<%Function GetAge(strDate)

   Dim strMonth, strDay, strYear, intIndex, datStartofYear, datEndofYear
   
   strMonth = Month(Date)
   strDay = Day(Date)
   strYear = Year(Date)
   
   For intIndex = 0 to 100 
   
      datStartofYear = strMonth & "/" & strDay & "/" & strYear - intIndex - 1
      datEndofYear = strMonth & "/" & strDay & "/" & strYear - intIndex
      If CDate(strDate) >= CDate(datStartofYear) And CDate(strDate) <= CDate(datEndofYear)  Then
         GetAge = intIndex + 1
         Exit For
      End If
      
   Next
	
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