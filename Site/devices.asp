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

'See if the user has the rights to visit this page
If AccessGranted Then
   ProcessSubmissions 
Else
   DenyAccess
End If %>

<%Sub ProcessSubmissions  

	Dim strSQL, arrTags, intIndex, strSQLWhere, strCurrentTag
	
	'Get the variables from the URL
	strSite = Request.QueryString("DeviceSite")
	intDeviceYear = Request.QueryString("Year")
	strView = Request.QueryString("View")
	strModel = Request.QueryString("Model")
	strTags = Request.QueryString("Tags")
	strRoom = Request.QueryString("Room")
	strAssigned = Request.QueryString("Assigned")
	strStatus = Request.QueryString("DeviceStatus")
	strBackLink = BackLink
	
	'If nothing was submitted send them back to the index page
   If strSite = "" And intDeviceYear = "" And strModel = "" And strRoom = "" And strTags = "" And strAssigned = "" And strStatus = "" Then
   	Response.Redirect("index.asp?Error=NoDevicesFound")
   End If

   'Check and see if anything was submitted to the site
   Select Case Request.Form("Submit")
      Case "Save"
      	SaveSearch
         
   End Select
   
   'Get the list of devices
	strSQLWhere = "WHERE "
	
	If strModel <> "" Then
   	strSQLWhere = strSQLWhere & "Devices.Model Like '%" & Replace(strModel,"'","''") & "%' AND "
   End If
   
   If strRoom <> "" Then
   	strSQLWhere = strSQLWhere & "Devices.Room='" & Replace(strRoom,"'","''") & "' AND "
   End If
	
	If strSite <> "" Then
      strSQLWhere = strSQLWhere & "Devices.Site = '" & strSite & "' AND "
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
   
	If strSQLWhere <> "WHERE " Then
		strSQLWhere = Left(strSQLWhere,Len(strSQLWhere) - 5)
	Else
		strSQLWhere = ""
	End If
	
	intTagCount = 0
	If strTags = "" Then
		strSQL = "SELECT ID,LGTag,Manufacturer,Model,SerialNumber,Room,DatePurchased,Site,Active,FirstName,LastName,UserName FROM Devices "
			strSQL = strSQL & strSQLWhere & " ORDER BY Devices.LGTag"
   Else
		strSQL = "SELECT Devices.ID,Devices.LGTag,Devices.Manufacturer,Devices.Model,Devices.SerialNumber,Devices.Room," & _
			"Devices.DatePurchased,Devices.Site,Devices.Active,Tags.Tag,FirstName,LastName,UserName FROM Tags INNER JOIN " & _
			"Devices ON Tags.LGTag = Devices.LGTag "
	
		arrTags = Split(strTags,",")
		intTagCount = UBound(arrTags) + 1
		
		strSQL = "SELECT COUNT(Tags.Tag) AS TagCount,Devices.LGTag,Devices.Manufacturer,Devices.Model,Devices.SerialNumber," & _
			"Devices.Room,Devices.DatePurchased,Devices.Site,Devices.Active,FirstName,LastName,UserName " & _
			"FROM(" & strSQL & strSQLWhere & ") WHERE "

		For intIndex = 0 to UBound(arrTags)
			strSQL = strSQL & "Tags.Tag='" & Trim(Replace(arrTags(intIndex),"'","''")) & "' OR "
		Next     
		strSQL = Left(strSQL,Len(strSQL) - 4)
		strSQL = strSQL & " GROUP BY Devices.LGTag,Devices.Manufacturer,Devices.Model,Devices.SerialNumber,Devices.Room," & _
			"Devices.DatePurchased,Devices.Site,Devices.Active,FirstName,LastName,UserName ORDER BY Devices.LGTag"
			  
	End If

	Set objDeviceList = Application("Connection").Execute(strSQL)

   'If no user is found send them back to the index page.
   If objDeviceList.EOF Then
   	Response.Redirect("index.asp?Error=NoDevicesFound")
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
      <link rel="stylesheet" type="text/css" href="style.css" /> 
      <link rel="apple-touch-icon" href="images/inventory.png" /> 
      <link rel="shortcut icon" href="images/inventory.ico" />
      <meta name="viewport" content="width=device-width" />
      <meta name="theme-color" content="#333333">
		<link rel="stylesheet" href="assets/css/jquery-ui.css">
		<script src="assets/js/jquery.js"></script>
		<script src="assets/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="assets/css/jquery.dataTables.min.css">
		<link rel="stylesheet" href="assets/css/buttons.dataTables.min.css">
		<script src="assets/js/jquery.dataTables.min.js"></script>
		<script src="assets/js/dataTables.buttons.min.js"></script>
		<script src="assets/js/buttons.colVis.min.js"></script>
		<script src="assets/js/buttons.html5.min.js"></script>
		<script src="assets/js/jszip.min.js"></script>
		<script type="text/javascript">
			$(document).ready( function () {
    			
    		<% If Not IsMobile Then %>
					$( document ).tooltip({track: true});
			<% End If %>
				
    			$('#ListView').DataTable( {
    				paging: false,
    				"info": false
    			});
    		} );
    	</script>
   </head>

   <body class="<%=strSiteVersion%>">
   
      <div class="Header"><%=Application("SiteName")%> (<%=intDeviceCount%>)</div>
      <div>
         <ul class="NavBar" align="center">
            <li><a href="index.asp"><img src="images/home.png" title="Home" height="32" width="32"/></a></li>
            <li><a href="log.asp"><img src="images/log.png" title="System Log" height="32" width="32"/></a></li>
            <li><a href="login.asp?action=logout"><img src="images/logout.png" title="Log Out" height="32" width="32"/></a></li>
         </ul>
      </div>  

<% If Not objDeviceList.EOF Then
	
		Select Case LCase(strView)
			Case "table"
				ShowDeviceTable
			Case "card"
				ShowDeviceCards
			Case Else
				If LCase(Application("DefaultView")) = "table" Then
					ShowDeviceTable
				Else
					ShowDeviceCards
				End If
		End Select
      SaveAsSearch
   End If %>
	<div class="Version">Version <%=Application("Version")%></div>
   </body>

   </html>

<%End Sub%>

<%Sub ShowDeviceCards %>  
	<div class="ViewButton">
		<a href="<%=SwitchView("Table")%>"><img src="images/table.png" title="Table View" height="32" width="32"/></a>
	</div>
	<div class="center"><%=FilterBar%></div>
	<div Class="<%=strColumns%>">
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

	Dim objFSO, objUserList, strSQL, objTagList
	
	Set objFSO = CreateObject("Scripting.FileSystemObject") %>

<% If objDeviceList(8) Then
		strCardType = "NormalCard"
	Else
		strCardType = "DisabledCard"
	End If %>
	<div class="Card <%=strCardType%>">
		<div>	
		<div class="CardTitle">
			<a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>">Asset Tag <%=objDeviceList(1)%></a>
		</div>		
		<% If objFSO.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\devices\" & Replace(objDeviceList(3)," ","") & ".png") Then %>      
				<a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>"> 
				<% If InStr(LCase(objDeviceList(3)),"ipad") Then %>
						<img class="PhotoCard" src="images/devices/<%=Replace(objDeviceList(3)," ","")%>.png" width="70" />
				<% Else %>
						<img class="PhotoCard" src="images/devices/<%=Replace(objDeviceList(3)," ","")%>.png" width="96" />
				<% End If %>
				</a>
		<% Else %>
				<a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>"> 
					<img class="PhotoCard" src="images/devices/missing.png" width="96" />
				</a>
		<% End If %>
		</div>
		
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
					<a href="https://selfsolve.apple.com/agreementWarrantyDynamic.do?caller=sp&sn=<%=objDeviceList(4)%>" target="_blank"><%=objDeviceList(4)%></a>
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
			<div Class="CardColumn1">Room: </div>
			<div Class="CardColumn2">
				<a href="devices.asp?Room=<%=objDeviceList(5)%>"><%=objDeviceList(5)%></a>
			</div>
		</div>
	<% End If %>
	<% strSQL = "SELECT Tag FROM Tags WHERE LGTag='" & objDeviceList(1) & "' ORDER BY Tag"
		Set objTagList = Application("Connection").Execute(strSQL)
		
		If Not objTagList.EOF Then %>
			<div>
				<div Class="CardColumn1">Tags: </div>
				<div Class="CardColumn2">
		<% Do Until objTagList.EOF %>	
				<a href="devices.asp?Tags=<%=objTagList(0)%>"><%=objTagList(0)%></a>
		<% 	objTagList.MoveNext
			Loop %>
				</div>
			</div>
	<% End If %>
	<% If objDeviceList(11) <> "" Then %>
			<div>
				<div Class="CardColumn1">Assigned To: </div>
				<div Class="CardColumn2">
					<a href="user.asp?UserName=<%=objDeviceList(11)%><%=strBackLink%>"><%=objDeviceList(9)%>&nbsp;<%=objDeviceList(10)%></a>
				</div>
			</div>
	<%	End If %>
	</div>
	 
<%End Sub %>

<%Sub ShowDeviceTable%>

	<div class="ViewButton">
		<a href="<%=SwitchView("Card")%>"><img src="images/card.png" title="Card View" height="32" width="32"/></a>
	</div>
	<div class="center"><%=Replace(FilterBar,"?","?View=Table&")%></div>

	<div>
		<table align="center" Class="ListView" id="ListView">
			<thead>
			<th>LG Tag</th>
			<th>Make</th>
			<th>Model</th>
			<th>Assigned To</th>
			<th>Room</th>
			<th>Year</th>
			<th>Tags</th>
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

	Dim objTagList, strSQL, objUserList, strRowClass

	If objDeviceList(8) Then
		strRowClass = ""
	Else 
		strRowClass = " Class=""Disabled"""
	End If %>
	
	<tr<%=strRowClass%>>
		<td id="center"><a href="device.asp?Tag=<%=objDeviceList(1)%><%=strBackLink%>"><%=objDeviceList(1)%></a></td>
		<td><%=objDeviceList(2)%></td>
		<td><%=objDeviceList(3)%></td>
		
<% If objDeviceList(11) <> "" Then %>
		<td><a href="user.asp?UserName=<%=objDeviceList(11)%><%=strBackLink%>"><%=objDeviceList(10)%>,&nbsp;<%=objDeviceList(9)%></a></td>
<% Else %>
		<td></td>
<% End If %>	
	
		<td><a href="devices.asp?Room=<%=objDeviceList(5)%>&View=Table"><%=objDeviceList(5)%></a></td>
<% If objDeviceList(6) <> "" Then %>	
		<td id="center"><a href="devices.asp?Year=<%=GetAge(objDeviceList(6))%>&View=Table"><%=GetAge(objDeviceList(6))%></a></td>
<% Else %>
		<td>N/A</td>
<% End If %>

		<td>
	<% strSQL = "SELECT Tag FROM Tags WHERE LGTag='" & objDeviceList(1) & "' ORDER BY Tag"
		Set objTagList = Application("Connection").Execute(strSQL)
	
		If Not objTagList.EOF Then %>
		<% Do Until objTagList.EOF %>	
				<a href="devices.asp?Tags=<%=objTagList(0)%>&View=Table"><%=objTagList(0)%></a>
		<% 	objTagList.MoveNext
			Loop %>
	<%	End If %>	
		</td>
	</tr>

<%End Sub %>

<%Function FilterBar
   
   If strSite <> "" Then
   	FilterBar = FilterBar & "Site = <a href=""devices.asp?DeviceSite=" & strSite & """>" & strSite & "</a> | "
   End If
   
   If intDeviceYear <> "" Then
   	FilterBar = FilterBar & "Year = <a href=""devices.asp?Year=" & intDeviceYear & """>" & intDeviceYear & "</a> | "
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