<%
	Response.CacheControl = "no-cache"
	Response.AddHeader("Pragma", "no-cache")
	Response.Expires = -1%>
<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of HR.Intranet.Server.NavLinksViewModel)" %>
<%@ Import Namespace="DMI.NET.Classes" %>
<%@ Import Namespace="DMI.NET.Code" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server.Interfaces" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%-- For other devs: Do not remove below line. --%>
<%="" %>
<%-- For other devs: Do not remove above line. --%>

<link id="SSIthemeLink" href="" rel="stylesheet" type="text/css" />
<link href="<%:Url.LatestContent("~/Content/jquery.mCustomScrollbar.min.css")%>" rel="stylesheet" />
<link href="<%= Url.LatestContent("~/Content/jquery.gridster.css")%>" rel="stylesheet" type="text/css" />
<script src="<%:Url.LatestContent("~/Scripts/jquery/jquery.mCustomScrollbar.min.js")%>"></script>
<script src="<%:Url.LatestContent("~/Scripts/FormScripts/linksMain.js")%>"></script>
<script src="<%:Url.LatestContent("~/Scripts/jquery/jquery.tablesorter.min.js")%>"></script>

<%Session("recordID") = 0
	Session("singleRecordID") = 0
	
	Dim fWFDisplayPendingSteps As Boolean
	Dim _PendingWorkflowStepsHTMLTable As New StringBuilder	'Used to construct the (temporary) HTML table that will be transformed into a jQuey grid table
	Dim _StepCount As Integer = 0

	Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)
	Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)
	
	'Get the pendings workflow steps from the database	
	If Licence.IsModuleLicenced(SoftwareModule.Workflow) Then

		Try
			
			fWFDisplayPendingSteps = NullSafeInteger(Session("SSILinkViewID")) = NullSafeInteger(Session("SingleRecordViewID"))
			
			Dim _rstDefSelRecords = objDataAccess.GetDataTable("spASRIntCheckPendingWorkflowSteps", CommandType.StoredProcedure)
		
			With _PendingWorkflowStepsHTMLTable
				.Append("<table id=""PendingStepsTable_Dash"">")
				.Append("<tr>")
				.Append("<th id=""DescriptionHeader"">Description</th>")
				.Append("<th id=""URLHeader"">URL</th>")
				.Append("<th id=""NameHeader"">URL</th>")
				.Append("</tr>")
			End With
			'Loop over the records
			For Each objRow As DataRow In _rstDefSelRecords.Rows
			
				_StepCount += 1
				With _PendingWorkflowStepsHTMLTable
					.Append("<tr>")
					.Append("<td>" & objRow("description").ToString() & "</td>")
					.Append("<td>" & objRow("url").ToString() & "</td>")
					.Append("<td>" & objRow("name").ToString() & "</td>")
					.Append("</tr>")
				End With
			Next
						
			_PendingWorkflowStepsHTMLTable.Append("</table>")
				
		
		Catch ex As Exception
			Throw
		
		End Try
	End If
		
%>

<div id="" class="DashContent" style="display: block;">
	<div class="tileContent">
		<%Dim fFirstSeparator = True
			Const iMaxRows As Integer = 4
			Dim iRowNum As Integer = 1
			Dim iHideablePopupIconID As Integer = 1
			Dim iHideableDrillDownIconID As Integer = 1
			Dim iColNum As Integer = 1
			Dim iSeparatorNum As Integer = 0
			Dim sOnClick As String = ""
			Dim sText As String = ""
			Dim sURL As String = ""
			Dim classIcon As String = ""
			Dim sNewWindow As String = ""
			Dim sAppFilePath As String = ""
			Dim sAppParameters As String = ""%>

		<div class="pendingworkflowlinks">
			<ul class="pendingworkflowsframe cols2">
				<li class="pendingworkflowlink-displaytype">
					<div class="wrapupcontainer">
						<div class="wrapuptext">
							<p class="pendingworkflowlinkseparator">Pending Workflows (max. four)</p>
						</div>
					</div>
					<div class="gridster pendingworkflowlinkcontent" id="gridster_PendingWorkflow">
						<ul id="pendingworkflowstepstiles">
						</ul>
					</div>
				</li>
			</ul>
		</div>

		<%fFirstSeparator = True%>
		<div class="hypertextlinks">
			<%
				Dim tileCount = 1
				For Each navlink In Model.NavigationLinks.FindAll(Function(n) n.LinkType = LinkType.HyperLink)
					Dim sTileColourClass = "Colour" & CStr(CInt(Math.Ceiling(Rnd() * 7)))
							
					If navlink.Element_Type = 1 Or navlink.LinkOrder = 0 or fFirstSeparator Then		' separator
						iRowNum = 1
						iColNum = 1
						If fFirstSeparator Then
							fFirstSeparator = False
						Else
			%>
								</ul>
		</div>
		</li> </ul>
							<%
							End If
						
							iSeparatorNum += 1
							
							If navlink.Text.Length > 0 And navlink.Element_Type = 1 Then
								sText = Html.Encode(navlink.Text)
								sText = sText.Replace("--", "")
								sText = sText.Replace("'", """")
							Else
								sText = "Hypertext Links"
							End If
							%>

		<ul class="hypertextlinkseparatorframe" id="hypertextlinkseparatorframe_<%=iSeparatorNum %>">
			<li class="hypertextlink-displaytype">
				<div class="wrapupcontainer hypertextlinktextseparator">
					<div class="wrapuptext hypertextlinktextseparator">
						<p class="hypertextlinkseparator hypertextlinkseparator-font hypertextlinkseparator-colour hypertextlinkseparator-size hypertextlinkseparator-bold hypertextlinkseparator-italics"><%=sText%></p>
					</div>
				</div>
				<div class="gridster hypertextlinkcontent" id="gridster_Hypertextlink_<%=tileCount%>">
					<ul>
						<%
						End If
						
						If navlink.Element_Type <> 1 Then	' not a separator; i.e. add a tile.
							If iRowNum > iMaxRows Then
								iColNum += 1
								iRowNum = 1
						%>
						<script type="text/javascript">
							$("#hypertextlinkseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
							$("#hypertextlinkseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
						</script>
						<%
						End If

						classIcon = ""
						sNewWindow = ""
								
						Select Case navlink.Element_Type
							Case ElementType.ButtonLink
										
								sURL = NullSafeString(navlink.URL).Replace("'", "\'")
								sURL = sURL.Replace("&", "&amp;")
								sURL = sURL.Replace("""", "&quot;")
								sURL = sURL.Replace(">", "&gt;")
								sURL = sURL.Replace("<", "&lt;")
									
								sAppFilePath = navlink.AppFilePath.Replace("\", "\\")
								sAppParameters = navlink.AppParameters.Replace("\", "\\")
								
								classIcon = "icon-external-link"
								If navlink.AppFilePath.Length > 0 Then
									sOnClick = "goApp('" & sAppFilePath & "', '" & sAppParameters & "')"
								ElseIf navlink.URL.Length > 0 Then
									If navlink.NewWindow = True Then
										sNewWindow = "1"
									Else
										sNewWindow = "0"
									End If
			
									sOnClick = "goURL('" & sURL & "', " & sNewWindow & ", true, '" & server.HtmlEncode(navlink.Text.Replace("'", "\'")) & "')"

								Else
									Dim sUtilityType = Convert.ToString(navlink.UtilityType)
									Dim sUtilityID = Convert.ToString(navlink.UtilityID)
									Dim sUtilityDef = sUtilityType & "_" & sUtilityID
									Dim sUtilityBaseTable = CStr(navlink.BaseTableID)
									sOnClick = "goUtility(" & sUtilityType & ", " & sUtilityID & ", '" & navlink.Text.Replace("'", "") & "', " & sUtilityBaseTable & ")"
										
								End If
									
										
							Case ElementType.OrgChart
								classIcon = "icon-sitemap"
								sOnClick = "loadPartialView('OrgChart', 'home', 'workframe')"
									
						End Select
						%>
						<li class="hypertextlinktext hypertextlinktext-highlightcolour <%=sTileColourClass%> flipTile" data-col="<%=iColNum %>" data-row="<%=iRowNum %>"
							data-sizex="1" data-sizey="1" onclick="<%=sOnclick%>">
							<%
								sText = navlink.Text
								If navlink.Text.Length > 30 Then
									sText = navlink.Text.Substring(0, 30) + "..."
								End If
							%>
							<a class="hypertextlinktext-font hypertextlinktext-colour hypertextlinktext-size hypertextlinktext-bold hypertextlinktext-italics" href="#" title="<%: navlink.Text%>"><%: sText%></a>
							<p class="hypertextlinktileIcon"><i class="<%=classIcon %>"></i></p>
						</li>
						<%
							iRowNum += 1
						End If
						tileCount += 1
					Next
								
					Dim objNavigation = CType(Session("NavigationLinks"), clsNavigationLinks)
								
					' Get the navigation hypertext links.
							
					Dim sDestination As String
							
					For Each objNavLink In objNavigation.GetNavigationLinks(False, LinkType.HyperLink)
							
						Dim sLinkText As New StringBuilder
						If objNavLink.Text1.Trim().Length > 0 Then sLinkText.Append(Html.Encode(objNavLink.Text1) & " ")
						sLinkText.Append(Html.Encode(objNavLink.Text2.Trim()))
						sText = sLinkText.ToString()
						
						If sLinkText.Length > 30 Then
							sText = sLinkText.ToString().Substring(0, 30) + "..."
						End If
		
						If objNavLink.LinkToFind = 0 Then
							sDestination = "linksMain?" & CStr(objNavLink.TableID) & "!" & CStr(objNavLink.ViewID)
			
							If objNavLink.SingleRecord = 1 Then
								sDestination = sDestination & "_0"
							Else
								sDestination = sDestination & "_" & CStr(Session("TopLevelRecID"))
							End If
						Else
							sDestination = "recordEditMain?multifind_0_" & CStr(objNavLink.TableID) & "!" & CStr(objNavLink.ViewID)
						End If
						If fFirstSeparator Then		' add a separator
							iRowNum = 1
							iColNum = 1
							If fFirstSeparator Then
								fFirstSeparator = False
							Else
						%>
					</ul>
				</div>
			</li>
		</ul>

		<%
		End If
		iSeparatorNum += 1
			
		%>

		<ul class="hypertextlinkseparatorframe" id="hypertextlinkseparatorframe_<%=iSeparatorNum %>">
			<li class="hypertextlink-displaytype">
				<div class="wrapupcontainer">
					<div class="wrapuptext">
						<p class="hypertextlinkseparator">Fixed Links</p>
					</div>
				</div>
				<div class="gridster hypertextlinkcontent" id="gridster_Hypertextlink_<%=tileCount%>">

					<ul>
						<%
						End If
						If iRowNum > iMaxRows Then
							iColNum += 1
							iRowNum = 1
						%>
						<script type="text/javascript">
							$("#hypertextlinkseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
							$("#hypertextlinkseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
						</script>
						<%
						End If
						%>
						<li class="hypertextlinktext Colour4" data-col="<%=iColNum %>" data-row="<%=iRowNum %>"
							data-sizex="1" data-sizey="1" onclick="goURL('<%=sDestination%>', 0, false)">
							<a class="hypertextlinktext-font hypertextlinktext-colour hypertextlinktext-size hypertextlinktext-bold hypertextlinktext-italics" href="#" title="<%=sText%>"><%=sText%></a>
							<p class="hypertextlinktileIcon"><i class="icon-external-link-sign"></i></p>
						</li>
						<%
							iRowNum += 1
							tileCount += 1
						Next

						If Not fFirstSeparator Then		' close off the hypertext group%>
					</ul>
				</div>
			</li>
		</ul>

		<%
		End If
		%>
	</div>

	<%fFirstSeparator = True%>
	<div class="linkspagebutton">		
		<div class="linkspagebuttonbox">
		<div class="ButtonLinkColumn">
			<%sOnClick = ""
				Dim sLinkKey As String = ""
				sAppFilePath = ""
				sAppParameters = ""
				sNewWindow = "0"
				Dim sTileBackColourStyle As String = ""
				Dim sTileForeColourStyle As String = ""

				For Each navlink In Model.NavigationLinks.FindAll(Function(n) n.LinkType = LinkType.Button)
					Dim sTileColourClass As String = ""																			

					If navlink.AppFilePath.Length > 0 Then
						sAppFilePath = NullSafeString(navlink.AppFilePath).Replace("\", "\\")
						sAppParameters = NullSafeString(navlink.AppParameters).Replace("\", "\\")
						' TODO: apps???
						sOnClick = "//goApp('" & sAppFilePath & "', '" & sAppParameters & "')"
						' sCheckKeyPressed = "CheckKeyPressed('APP', '" & sAppFilePath & "', 0, '" & sAppParameters & "')"
			
					ElseIf NullSafeString(navlink.URL).Length > 0 Then
						sURL = NullSafeString(navlink.URL)
						sURL = sURL.Replace("&", "&amp;")
						sURL = sURL.Replace("""", "&quot;")
						sURL = sURL.Replace(">", "&gt;")
						sURL = sURL.Replace("<", "&lt;")

						If navlink.NewWindow = True Then
							sNewWindow = "1"
						Else
							sNewWindow = "0"
						End If
			
						sOnClick = "goURL('" & sURL & "', " & sNewWindow & ", true, '" & server.HtmlEncode(navlink.Text.Replace("'", "\'")) & "')"

					Else
						If navlink.UtilityID > 0 Then
							Dim sUtilityType = CStr(navlink.UtilityType)
							Dim sUtilityID = CStr(navlink.UtilityID)
							Dim sUtilityBaseTable = CStr(navlink.BaseTableID)
												
							sOnClick = "goUtility(" & sUtilityType & ", " & sUtilityID & ", '" & navlink.Text & "', " & sUtilityBaseTable & ")"
						Else
							sLinkKey = "recedit" & "_" & Session("TopLevelRecID").ToString() & "_" & navlink.ID
												
							sOnClick = "goScreen('" & sLinkKey & "')"
						End If
					End If

					If navlink.Element_Type = 1 Or navlink.LinkOrder = 0 or fFirstSeparator Then			' separator
						iRowNum = 1
						iColNum = 1
						Dim sSeparatorColor = ""
						
						If navlink.SeparatorColour <> "" And navlink.SeparatorColour <> "#FFFFFF" Then sSeparatorColor = "border-bottom: 2px solid " & navlink.SeparatorColour & ";"
						
						If Session("CurrentLayout").ToString() = Layout.tiles.ToString() Then
							If navlink.SeparatorColour <> "" And navlink.SeparatorColour <> "#FFFFFF" Then
								sSeparatorColor = "border-bottom: 2px solid " & navlink.SeparatorColour & ";"
								sTileBackColourStyle = "background-color: " & navlink.SeparatorColour & ";"
								sTileForeColourStyle = "color: white;"
							Else
								sTileBackColourStyle = ""
								sTileForeColourStyle = ""
							End If
							
						End If
						
						
							If fFirstSeparator Then
								fFirstSeparator = False
							Else
			%>
														</ul>
		</div>
		</li> </ul>
												<%									
												End If
												If navlink.SeparatorOrientation = 1 Then	' Vertical break/new column %>
	</div>
	<div class="ButtonLinkColumn">
		<%
		End If
		
		iSeparatorNum += 1
		
		If navlink.Text.Length > 30 Then
			navlink.Text = navlink.Text.Substring(0, 30)
		End If
		
		If navlink.Text.Length > 0 And navlink.Element_Type = 1 Then
			sText = Html.Encode(navlink.Text)
			sText = sText.Replace("--", "")
			sText = sText.Replace("'", """")
		Else
			sText = "Button Links"
		End If

		
		%>
		<ul class="linkspagebuttonseparatorframe" id="linkspagebuttonseparatorframe_<%=iSeparatorNum %>">
			<li class="linkspagebutton-displaytype">
				<div class="wrapupcontainer linkspagebuttonseparator-bordercolour" style="<%=sSeparatorColor%>">
					<div class="wrapuptext">
						<p class="linkspagebuttonseparator linkspagebuttonseparator-font linkspagebuttonseparator-colour linkspagebuttonseparator-size linkspagebuttonseparator-bold linkspagebuttonseparator-italics"><%=sText%></p>
					</div>
				</div>
				<div class="gridster buttonlinkcontent" id="gridster_buttonlink_<%=tileCount%>">
					<ul>
						<%											
						End If  ' end separator...
						If navlink.Element_Type <> 1 Then	' not a separator...
							If iRowNum > iMaxRows Then	 ' start a new column if required (affects tiles only)
								iColNum += 1
								iRowNum = 1
						%>
						<script type="text/javascript">
							$("#linkspagebuttonseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
							$("#linkspagebuttonseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
						</script>
						<%
						End If

						
						' if style not set...
						If sTileBackColourStyle = "" Then
							sTileColourClass = "Colour" & CStr(CInt(Math.Ceiling(Rnd() * 7)))
						Else
							sTileColourClass = ""
						End If
						
						Select Case navlink.Element_Type

							Case ElementType.ButtonLink
								Dim sIconClass As String = "icon-file"
									
								If navlink.UtilityType = -1 Then	' screen view
									sIconClass = "icon-table"
								ElseIf navlink.UtilityType = 25 Then
									sIconClass = "icon-magic"
								End If
						
								If navlink.Text.Length > 30 Then
									navlink.Text = navlink.Text.Substring(0, 30) + "..."
								End If
								%>
						<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1" style="<%:sTileBackColourStyle%><%:sTileForeColourStyle%>" class="linkspagebuttontext <%=sTileColourClass%>" onclick="<%=sOnclick%>">
							<a style="<%:sTileForeColourStyle%>" class="linkspagebutton-displaytype linkspagebuttontext-alignment linkspagebutton-colourtheme" href="#"><span class="linkspageprompttext-font linkspageprompttext-colour linkspageprompttext-size linkspageprompttext-bold linkspageprompttext-italics"><%: navlink.Prompt.Replace("...", "") & " "%></span>
								<span class="linkspagebuttontext-font linkspagebuttontext-colour linkspagebuttontext-size linkspagebuttontext-bold linkspagebuttontext-italics"><%: navlink.Text %></span></a>
							<p class="linkspagebuttontileIcon"><i class="<%=sIconClass%>"></i></p>
						</li>
						<%
							iRowNum += 1
									
						Case ElementType.Chart
									
							Dim iChart_TableID As Long = navlink.Chart_TableID
							Dim iChart_ColumnID As Long = navlink.Chart_ColumnID
							Dim iChart_FilterID As Long = navlink.Chart_FilterID
							Dim iChart_AggregateType As Long = navlink.Chart_AggregateType
							Dim iChart_ElementType As ElementType = navlink.Element_Type
							'Dim fChart_ShowLegend = navlink.Chart_ShowLegend
							Dim iChart_Type = navlink.Chart_Type
							'Dim fChart_ShowGrid = navlink.Chart_ShowGrid
							'Dim fChart_StackSeries = navlink.Chart_StackSeries
							'Dim fChart_ShowValues = navlink.Chart_ShowValues
							'Dim sChart_ColumnName = Replace(navlink.Chart_ColumnName, "_", " ")
							'Dim sChart_ColumnName_2 = Replace(navlink.Chart_ColumnName_2, "_", " ")
		
							Dim iChart_TableID_2 As Long = navlink.Chart_TableID_2
							Dim iChart_ColumnID_2 As Long = navlink.Chart_ColumnID_2
							Dim iChart_TableID_3 As Long = navlink.Chart_TableID_3
							Dim iChart_ColumnID_3 As Long = navlink.Chart_ColumnID_3
		
							'Dim iChartInitialDisplayMode = CleanNumeric(navlink.InitialDisplayMode)
		
							Dim iChart_SortOrderID As Long = navlink.Chart_SortOrderID
							Dim iChart_SortDirection As Integer = navlink.Chart_SortDirection
							Dim iChart_ColourID As Long = navlink.Chart_ColourID
		
							'Dim fChart_ShowPercentages = navlink.Chart_ShowPercentages
		
							Dim fMultiAxis As Boolean
									
							If iChart_TableID_2 > 0 Or iChart_TableID_3 > 0 Then
								fMultiAxis = True
							Else
								fMultiAxis = False
							End If
									
							' Drilldown?
							If navlink.UtilityID > 0 Then
								' sOnclick = "goUtilityDash('" & navlink.UtilityType & "_" & navlink.UtilityID.ToString() & "_" & navlink.BaseTable
								sOnClick = "goUtility(" & navlink.UtilityType & ", " & navlink.UtilityID & ", '" & navlink.Text & "', " & navlink.BaseTableID & ")"
							Else
								sOnClick = ""
							End If
									
						%>

						<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1" style="<%:sTileBackColourStyle%>" class="linkspagebuttontext ui-state-disabled <%=sTileColourClass%> displayonly">
							<a style="<%:sTileForeColourStyle%>" href="#"><%If Session("CurrentLayout").ToString() <> Layout.tiles.ToString() Then Response.Write(navlink.Text)%>
								<%
									If navlink.UtilityID > 0 And navlink.DrillDownHidden = False Then
										iHideableDrillDownIconID += 1
								%>
								<img id="drillDownIcon<%=iHideableDrillDownIconID%>" src="<%:Url.Content("~/Content/images/Utilities.gif")%>" style="<%=IIf(Session("CurrentLayout").ToString() = Layout.tiles.ToString(), "background: wheat;", "")%> float: right; cursor: pointer; width: 16px; height: 16px; vertical-align: bottom; margin-right: 5px" alt="Drilldown..." title="Drill down to data..."
									onclick="<%=sOnClick %>" />
								<%
								End If
								%>
								<img id="popupIcon<%=iHideablePopupIconID%>" src="<%:Url.Content("~/Content/images/Chart_Popout.png")%>" style="<%=IIf(Session("CurrentLayout").ToString() = Layout.tiles.ToString(), "background: wheat;", "")%> float: right; cursor: pointer; width: 16px; height: 16px; vertical-align: bottom;" alt="Popout chart..." title="View this chart in a new window"
									onclick="popoutchart('<%=fMultiAxis%>', '<%=navlink.Chart_ShowLegend%>', '<%=navlink.Chart_ShowGrid%>', '<%=navlink.Chart_ShowValues%>', '<%=navlink.Chart_StackSeries%>', '<%=navlink.Chart_ShowPercentages%>', '<%=iChart_Type%>', '<%=iChart_TableID%>', '<%=iChart_ColumnID%>', '<%=iChart_FilterID%>', '<%=iChart_AggregateType%>', '<%=CInt(iChart_ElementType)%>', '<%=iChart_TableID_2%>', '<%=iChart_ColumnID_2%>', '<%=iChart_TableID_3%>', '<%=iChart_ColumnID_3%>', '<%=iChart_SortOrderID%>', '<%=iChart_SortDirection%>', '<%=iChart_ColourID%>','<%=navlink.Text%>', '<%=Session("ui-admin-theme").ToString() %>')" />
							</a>
							<%If Not (navlink.InitialDisplayMode = 0 Or Session("CurrentLayout").ToString() = Layout.tiles.ToString()) Then%>
							<p class="linkspagebuttontileIcon">
								<i class="icon-bar-chart"></i>
							</p>
							<%End If%>
							<%
								iHideablePopupIconID += 1
								
								'Set some values that depend on the current layout
								Dim Height As Integer = 296
								Dim Width As Integer = 412
								Dim ShowLegend As Boolean = navlink.Chart_ShowLegend
								Dim ShowLabels As Boolean = True
								If Session("CurrentLayout").ToString() = Layout.tiles.ToString() Then
									Height = 120 'Tile height
									Width = 120	'Tile width
									ShowLegend = False
									ShowLabels = False
								End If

								If navlink.InitialDisplayMode = 0 Then%>
							<div class="widgetplaceholder chart">
								<%If fMultiAxis Then%>
								<div>
									<img onerror="$('#popupIcon<%=iHideablePopupIconID - 1%>').hide(); $(this).parent().parent().css('height', '20px'); $(this).parent().parent().html('<%If Session("CurrentLayout").ToString() = Layout.tiles.ToString() Then Response.Write("<p title=\'" & navlink.Text & "\' class=\'linkspagebuttontileIcon\'><i class=\'icon-bar-chart\'></i></p><p style=\'font-size: smaller; text-align: center\'>" & navlink.Text & "<br/><br/>(No records)</p>") Else Response.Write("No matching records")%>');" 
											 src="<%:Url.Action("GetMultiAxisChart", "Home", New With {.Height = Height, .Width = Width, .ShowLegend = ShowLegend, .DottedGrid = navlink.Chart_ShowGrid, .ShowValues = navlink.Chart_ShowValues, .Stack = navlink.Chart_StackSeries, .ShowPercent = navlink.Chart_ShowPercentages, .ChartType = iChart_Type, .TableID = iChart_TableID, .ColumnID = iChart_ColumnID, .FilterID = iChart_FilterID, .AggregateType = iChart_AggregateType, .ElementType = CInt(iChart_ElementType), .TableID_2 = iChart_TableID_2, .ColumnID_2 = iChart_ColumnID_2, .TableID_3 = iChart_TableID_3, .ColumnID_3 = iChart_ColumnID_3, .SortOrderID = iChart_SortOrderID, .SortDirection = iChart_SortDirection, .ColourID = iChart_ColourID, .ShowLabels = ShowLabels})%>"
											 alt="Chart"
											 title="<%:navlink.Text%>"
											 style="cursor: pointer"
										   onclick="popoutchart('<%=fMultiAxis%>', '<%=navlink.Chart_ShowLegend%>', '<%=navlink.Chart_ShowGrid%>', '<%=navlink.Chart_ShowValues%>', '<%=navlink.Chart_StackSeries%>', '<%=navlink.Chart_ShowPercentages%>', '<%=iChart_Type%>', '<%=iChart_TableID%>', '<%=iChart_ColumnID%>', '<%=iChart_FilterID%>', '<%=iChart_AggregateType%>', '<%=CInt(iChart_ElementType)%>', '<%=iChart_TableID_2%>', '<%=iChart_ColumnID_2%>', '<%=iChart_TableID_3%>', '<%=iChart_ColumnID_3%>', '<%=iChart_SortOrderID%>', '<%=iChart_SortDirection%>', '<%=iChart_ColourID%>','<%=navlink.Text%>', '<%=Session("ui-admin-theme").ToString()%>', '<%=navlink.Chart_ColumnName%>', '<%=navlink.Chart_ColumnName_2%>')"
										 />
								</div>
								<%Else%>
								<div>
									<img onerror="$('#popupIcon<%=iHideablePopupIconID - 1%>').hide(); $(this).parent().parent().css('height', '20px'); $(this).parent().parent().html('<%If Session("CurrentLayout").ToString() = Layout.tiles.ToString() Then Response.Write("<p title=\'" & navlink.Text & "\' class=\'linkspagebuttontileIcon\'><i class=\'icon-bar-chart\'></i></p><p style=\'font-size: smaller; text-align: center\'>" & navlink.Text & "<br/><br/>(No records)</p>") Else Response.Write("No matching records")%>');"
											 src="<%:Url.Action("GetChart", "Home", New With {.Height = Height, .Width = Width, .ShowLegend = ShowLegend, .DottedGrid = navlink.Chart_ShowGrid, .ShowValues = navlink.Chart_ShowValues, .Stack = navlink.Chart_StackSeries, .ShowPercent = navlink.Chart_ShowPercentages, .ChartType = iChart_Type, .TableID = iChart_TableID, .ColumnID = iChart_ColumnID, .FilterID = iChart_FilterID, .AggregateType = iChart_AggregateType, .ElementType = CInt(iChart_ElementType), .SortOrderID = iChart_SortOrderID, .SortDirection = iChart_SortDirection, .ColourID = iChart_ColourID, .ShowLabels = ShowLabels})%>" 
											 alt="Chart" 
											 title="<%:navlink.Text%>"
											 style="cursor: pointer"
											 onclick="popoutchart('<%=fMultiAxis%>', '<%=navlink.Chart_ShowLegend%>', '<%=navlink.Chart_ShowGrid%>', '<%=navlink.Chart_ShowValues%>', '<%=navlink.Chart_StackSeries%>', '<%=navlink.Chart_ShowPercentages%>', '<%=iChart_Type%>', '<%=iChart_TableID%>', '<%=iChart_ColumnID%>', '<%=iChart_FilterID%>', '<%=iChart_AggregateType%>', '<%=CInt(iChart_ElementType)%>', '<%=iChart_TableID_2%>', '<%=iChart_ColumnID_2%>', '<%=iChart_TableID_3%>', '<%=iChart_ColumnID_3%>', '<%=iChart_SortOrderID%>', '<%=iChart_SortDirection%>', '<%=iChart_ColourID%>','<%=navlink.Text%>', '<%=Session("ui-admin-theme").ToString()%>', '<%=navlink.Chart_ColumnName%>', '<%=navlink.Chart_ColumnName_2%>')"
										 />
								</div>
								<%End If%>
																			
								<a href="#"></a>
							</div>
							<%
							Else
								Dim objChart As IChart
								Dim sErrorDescription As String = ""
								' Dim fFormatting_Use1000Separator As Boolean = (navlink.Formatting_Use1000Separator = 1)
																								
								If fMultiAxis = True Then
									objChart = New HR.Intranet.Server.clsMultiAxisChart
								Else
									objChart = New HR.Intranet.Server.clsChart
								End If

								objChart.SessionInfo = CType(Session("SessionContext"), SessionInfo)

								Dim mrstChartData As DataTable
								Err.Clear()
			
								If fMultiAxis = True Then
									mrstChartData = objChart.GetChartData(iChart_TableID, iChart_ColumnID, iChart_FilterID, iChart_AggregateType, iChart_ElementType, iChart_TableID_2, iChart_ColumnID_2, iChart_TableID_3, iChart_ColumnID_3, iChart_SortOrderID, iChart_SortDirection, iChart_ColourID)
								Else
									mrstChartData = objChart.GetChartData(iChart_TableID, iChart_ColumnID, iChart_FilterID, iChart_AggregateType, iChart_ElementType, 0, 0, 0, 0, iChart_SortOrderID, iChart_SortDirection, iChart_ColourID)
								End If

								If (Err.Number <> 0) Then
									sErrorDescription = "The Chart field values could not be retrieved." & vbCrLf & FormatError(Err.Description)
								End If
			
								If Not mrstChartData Is Nothing Then
									If mrstChartData.Rows.Count > 500 Then mrstChartData = Nothing ' limit to 500 rows as get row buffer limit exceeded error.
								End If

								Dim Chart_AggregateType As ChartAggregateType = navlink.Chart_AggregateType

								If Session("CurrentLayout").ToString() = Layout.tiles.ToString() Then 'Put and icon in tile mode
									If mrstChartData.Rows.Count > 0 AndAlso (TryCast(mrstChartData.Rows(0)(0), String) <> "No Access" And TryCast(mrstChartData.Rows(0)(0), String) <> "No Data") Then
										Dim popupChartCall As String = "popoutchart('" & fMultiAxis & "', '" & navlink.Chart_ShowLegend & "', '" & navlink.Chart_ShowGrid & "', '" & navlink.Chart_ShowValues & "', '" & navlink.Chart_StackSeries & "', '" & navlink.Chart_ShowPercentages & "', '" & iChart_Type & "', '" & iChart_TableID & "', '" & iChart_ColumnID & "', '" & iChart_FilterID & "', '" & iChart_AggregateType & "', '" & CInt(iChart_ElementType) & "', '" & iChart_TableID_2 & "', '" & iChart_ColumnID_2 & "', '" & iChart_TableID_3 & "', '" & iChart_ColumnID_3 & "', '" & iChart_SortOrderID & "', '" & iChart_SortDirection & "', '" & iChart_ColourID & "',encodeURI('" & navlink.Text & "').replace(/&/g,'%26'),'" & Session("ui-admin-theme").ToString() & "')"
										Dim popuChartCallOnclick = "onclick=" & Chr(34) & popupChartCall & Chr(34)
										Response.Write("<p " & popuChartCallOnclick & " title='" & navlink.Text & "' class='linkspagebuttontileIcon'><i class='icon-bar-chart'></i></p><p " & popupChartCall & " style='text-align: center'>" & navlink.Text & "</p>")
									Else
										Response.Write("<p title='" & navlink.Text & "' class='linkspagebuttontileIcon'><i class='icon-bar-chart'></i></p><p style='font-size: smaller; text-align: center'>(No records)</p>")
									End If
								End If
								%>
							<div class="widgetplaceholder datagrid" id="WidgetPlaceHolder<%=iRowNum%>">
								<table id="DataTable<%=iRowNum%>" cellspacing="0" cellpadding="5" rules="all" frame="box" style="width: 100%; vertical-align: top; border: 3px solid lightgray">
									<%If mrstChartData.Rows.Count > 0 AndAlso (TryCast(mrstChartData.Rows(0)(0), String) <> "No Access" And TryCast(mrstChartData.Rows(0)(0), String) <> "No Data") Then%>
									<thead>
									<tr>
										<th style="font-weight: normal; text-align: left; cursor: default">
											<%=Left(NullSafeString(navlink.Chart_ColumnName), 50)%>
										</th>
										<%If fMultiAxis Then%>
										<th style="font-weight: normal; text-align: left; cursor: default">
											<%=Trim(Left(NullSafeString(navlink.Chart_ColumnName_2), 50))%>
										</th>
										<th style="font-weight: normal; text-align: right; cursor: default">
											<%Else%>
										<th style="font-weight: normal; text-align: right; cursor: default">
											<%End If%>
											<%Response.Write(Chart_AggregateType.ToString())%>
										</th>
									</tr>
									</thead>
									<tbody>
									<%
										If mrstChartData.Rows.Count > 0 Then
											For Each objRow As DataRow In mrstChartData.Rows
									%>
									<tr>
										<td class="bordered" style="width: 150px; text-align: left; white-space: nowrap">
											<%If fMultiAxis Then%>
											<%=Trim(Left(NullSafeString(objRow(1)), 50))%>
											<%Else%>
											<%=Trim(Left(NullSafeString(objRow(0)), 50))%>
											<%End If%>
										</td>
										<%If fMultiAxis Then%>
										<td class="bordered" style="text-align: left; white-space: nowrap">
											<div style="width: 150px; white-space: nowrap">
												<%=Trim(Left(NullSafeString(objRow(3)), 50))%>
											</div>
										</td>
										<%End If%>
										<td class="bordered" style="text-align: right; vertical-align: top; padding-bottom: 0; white-space: nowrap; overflow: hidden">
											<%If fMultiAxis Then%>
											<%If navlink.UseFormatting = True And (TryCast(objRow(4), String) <> "No Access" And TryCast(objRow(4), String) <> "No Data") Then%>
											<%=FormatNumber(CDbl(Trim(Left(NullSafeString(objRow(4)), 50))), navlink.Formatting_DecimalPlaces, , , TriState.UseDefault)%>
											<%Else%>
											<%=Trim(Left(NullSafeString(objRow(4)), 50))%>
											<%End If
											Else
												If navlink.UseFormatting = True And (TryCast(objRow(1), String) <> "No Access" And TryCast(objRow(1), String) <> "No Data") Then%>
											<%=FormatNumber(CDbl(Trim(Left(NullSafeString(objRow(1)), 50))), navlink.Formatting_DecimalPlaces, , , TriState.UseDefault)%>
											<%Else%>
											<%=Trim(Left(NullSafeString(objRow(1)), 50))%>
											<%
											End If
										End If%>
										</td>
									</tr>
									<%    
											
									Next
									%>
										</tbody>
									<%
								End If
							Else
									If Session("CurrentLayout").ToString() = Layout.tiles.ToString() Then	'Put and icon in tile mode
										Response.Write("<p title='" & navlink.Text & "' class='linkspagebuttontileIcon'><i class='icon-bar-chart'></i></p><p style='font-size: smaller; text-align: center'>(No records)</p>")
									Else
									%>
									<tr>
										<td class="bordered" style="text-align: center;" rowspan="3">No matching records found</td>
									</tr>
									<%End If%>
									<script type="text/javascript">
										// No data on this chart, adjust UI accordingly
										<%If Session("CurrentLayout").ToString() <> Layout.tiles.ToString() Then%>
										$("#WidgetPlaceHolder<%=iRowNum%>").css('height', "40px"); //Reduce the size of the parent div ('widgetplaceholder')
										$("#WidgetPlaceHolder<%=iRowNum%>").children(0).css('border', 'none'); //Remove the border of the table
										<%End If%>
										$("#drillDownIcon<%=iHideableDrillDownIconID%>").hide(); //Hide the drilldown icon
										$("#popupIcon<%=iHideablePopupIconID - 1%>").hide(); //Hide the popup icon
									</script>
									<%End If%>
								</table>
								<script type="text/javascript">
									//Attach table sorter to the table
									$("#DataTable<%=iRowNum%>").tablesorter();
								</script>
							</div>
							<%End If%>
						</li>
						<%iRowNum += 1%>

						<%Case ElementType.PendingWorkflows And Licence.IsModuleLicenced(SoftwareModule.Workflow)
								sText = "No pending workflow steps"
								sText = String.Format("{0} Pending workflow step{1}", _StepCount, IIf(_StepCount <> 1, "s", ""))
								%>
						<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="2" data-sizey="1" style="<%:sTileBackColourStyle%><%:sTileForeColourStyle%>" class="linkspagebuttontext <%=sTileColourClass%> displayonly pwfslink" onclick="relocateURL('WorkflowPendingSteps', 0)">
							<div class="pwfTile <%=sTileColourClass%>">
								<p class="linkspagebuttontileIcon">
									<i class="icon-inbox"></i>
									<div class="workflowCount"></div>
								</p>
								<div class="widgetplaceholder generaltheme">
									<div><i class="icon-inbox"></i></div>
									<a style="<%:sTileForeColourStyle%>" class="linkspageprompttext-font linkspageprompttext-colour linkspageprompttext-size linkspageprompttext-bold linkspageprompttext-italics" href="#"><%=sText%></a>
								</div>
							</div>
							<div class="pwfList <%=sTileColourClass%>" style="display: none; <%:sTileBackColourStyle%><%:sTileForeColourStyle%>">
								<p><span>Pending steps:</span></p>
								<table></table>
							</div>
						</li>
						<%
							iRowNum += 1
							fWFDisplayPendingSteps = False

						Case ElementType.DatabaseValue
									
							' DBValue Formatting options...
							Dim fUseFormatting = navlink.UseFormatting
									
							Dim iFormatting_DecimalPlaces = navlink.Formatting_DecimalPlaces
							Dim fFormatting_Use1000Separator = navlink.Formatting_Use1000Separator
							Dim sFormatting_Prefix = Html.Encode(navlink.Formatting_Prefix)
							Dim sFormatting_Suffix = Html.Encode(navlink.Formatting_Suffix)
		
							' DBValue Conditional Formatting options...
							Dim fUseConditionalFormatting = navlink.UseConditionalFormatting

							Dim sCFOperator(2) As String
							Dim sCFValue(2) As String
							Dim sCFStyle(2) As String
							Dim sCFColour(2) As String
									
							sCFOperator(0) = navlink.ConditionalFormatting_Operator_1
							sCFOperator(1) = navlink.ConditionalFormatting_Operator_2
							sCFOperator(2) = navlink.ConditionalFormatting_Operator_3
		
							sCFValue(0) = navlink.ConditionalFormatting_Value_1
							sCFValue(1) = navlink.ConditionalFormatting_Value_2
							sCFValue(2) = navlink.ConditionalFormatting_Value_3
		
							sCFStyle(0) = navlink.ConditionalFormatting_Style_1
							sCFStyle(1) = navlink.ConditionalFormatting_Style_2
							sCFStyle(2) = navlink.ConditionalFormatting_Style_3
		
							sCFColour(0) = navlink.ConditionalFormatting_Colour_1
							sCFColour(1) = navlink.ConditionalFormatting_Colour_2
							sCFColour(2) = navlink.ConditionalFormatting_Colour_3

							' Set the conditional formatting defaults
							Dim sCFForeColor = "" + Session("Config-linkspagebuttontext-colour")
							Dim sCFFontBold = "" + Session("Config-linkspagebuttontext-bold")
							Dim sCFFontItalic = "" + Session("Config-linkspagebuttontext-italic")
							Dim sCFVisible = True
		
							Dim fFormattingApplies = True
									
							Dim sErrorDescription = ""
							Dim sPrompt = navlink.Text
							sText = ""
									
							' Create the reference to the DLL (Report Class)
							Dim objChart = New HR.Intranet.Server.clsChart
							objChart.SessionInfo = objSession
			
							Err.Clear()
							Dim mrstDbValueData = objChart.GetChartData(navlink.Chart_TableID, navlink.Chart_ColumnID, navlink.Chart_FilterID, _
																													navlink.Chart_AggregateType, navlink.Element_Type, 0, 0, 0, 0, navlink.Chart_SortOrderID, _
																													navlink.Chart_SortDirection, navlink.Chart_ColourID)

							If Err.Number <> 0 Then
								sErrorDescription = "The Database Values could not be retrieved." & vbCrLf & FormatError(Err.Description)
							End If
									
							If Len(sErrorDescription) = 0 Then

								For Each objRow As DataRow In mrstDbValueData.Rows
									sText = objRow(0).ToString()
								Next
								Dim fDoFormatting As Boolean
								If fUseConditionalFormatting = True Then
									For jnCount = 0 To 2
										fDoFormatting = False
										If sCFValue(jnCount) <> vbNullString And sText <> "No Data" And sText <> "No Access" Then 'objChart.GetChartData returns either "No Data" or "No Access" if there is a permission-denied problem
											Select Case sCFOperator(jnCount)
												Case "is equal to"
													If CType(sText, Int32) = CType(sCFValue(jnCount), Int32) Then fDoFormatting = True
												Case "is not equal to"
													If CType(sText, Int32) <> CType(sCFValue(jnCount), Int32) Then fDoFormatting = True
												Case "is less than or equal to"
													If CType(sText, Int32) <= CType(sCFValue(jnCount), Int32) Then fDoFormatting = True
												Case "is greater than or equal to"
													If CType(sText, Int32) >= CType(sCFValue(jnCount), Int32) Then fDoFormatting = True
												Case "is less than"
													If CType(sText, Int32) < CType(sCFValue(jnCount), Int32) Then fDoFormatting = True
												Case "is greater than"
													If CType(sText, Int32) > CType(sCFValue(jnCount), Int32) Then fDoFormatting = True
											End Select
														
											If fDoFormatting Then
												sCFForeColor = sCFColour(jnCount)
												Select Case sCFStyle(jnCount)
													Case "Bold"
														sCFFontBold = "font-weight:bold"
													Case "Italic"
														sCFFontItalic = "font-style:italic"
													Case "Bold & Italic"
														sCFFontItalic = "font-weight:bold;font-style:italic"
													Case "Hidden"
														sCFVisible = False
													Case "Normal"
														fFormattingApplies = True
													Case Else
														fFormattingApplies = False
												End Select
												Exit For
											End If
										End If
									Next
								Else
									fFormattingApplies = False
								End If


							Else	 ' no results - return zero
								sText = "No Data"
							End If
								
							If sText <> "No Data" And sCFVisible = True Then
									
								If fFormattingApplies Then
						%>
						<li id="li_<%: navlink.id %>" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
							data-sizey="1" style="<%:sTileBackColourStyle%><%:sTileForeColourStyle%>" class="linkspagebuttontext ui-state-disabled <%=sTileColourClass%> displayonly linkspagebuttontext-font linkspagebuttontext-colour linkspagebuttontext-size linkspagebuttontext-bold linkspagebuttontext-italics">
							<div class="DBValueScroller" id="marqueeDBV<%: navlink.id %>">
								<p class="DBValue" style="color: <%=sCFForeColor%>; <%=sCFFontBold%>; <%=sCFFontItalic%>" id="DBV<%: navlink.id %>">
									<%If fUseFormatting = True Then%>
									<span class="DBVSpan"><%=sFormatting_Prefix%><%=FormatNumber(cdbl(sText), iFormatting_DecimalPlaces,,,fFormatting_Use1000Separator)%><%=sFormatting_Suffix%></span>
									<%Else%>
									<span class="DBVSpan"><%: sText %></span>
									<%End If%>
								</p>
							</div>
							<a style="<%:sTileForeColourStyle%>" href="#">
								<p class="DBValueCaption" style="color: <%=sCFForeColor%>; <%=sCFFontBold%>; <%=sCFFontItalic%>">
									<%sText = navlink.Text
										If sText.Length > 30 Then
											sText = sText.Substring(0, 30) + "..."
										End If
									sText = sText.Replace("--", "")
									sText = sText.Replace("'", """")%>
									<%: sText %>
								</p>
							</a>
						</li>

						<%Else%>
						<li id="li_<%: navlink.id %>" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
							data-sizey="1" style="<%:sTileBackColourStyle%><%:sTileForeColourStyle%>" class="linkspagebuttontext ui-state-disabled <%=sTileColourClass%> displayonly linkspagebuttontext-font linkspagebuttontext-colour linkspagebuttontext-size linkspagebuttontext-bold linkspagebuttontext-italics">
							<div class="DBValueScroller" id="marqueeDBV<%: navlink.id %>">
								<p class="DBValue" id="DBV<%: navlink.id %>">
									<%If fUseFormatting = True Then%>
									<span class="DBVSpan"><%=sFormatting_Prefix%><%=FormatNumber(cdbl(sText), iFormatting_DecimalPlaces,,,fFormatting_Use1000Separator)%><%=sFormatting_Suffix%></span>
									<%Else%>
									<span class="DBVSpan"><%: sText %></span>
									<%End If%>
								</p>
							</div>
							<a style="<%:sTileForeColourStyle%>" href="#">
								<p class="DBValueCaption">
									<%sText = navlink.Text
										If sText.Length > 30 Then
											sText = sText.Substring(0, 30) + "..."
										End If
										sText = sText.Replace("--", "")
										sText = sText.Replace("'", """")%>
									<%: sText%>
								</p>
							</a>
						</li>
						<%End If
						End If%>

						<script type="text/javascript">							//loadjscssfile('$.getScript("../scripts/widgetscripts/wdg_oHRDBV.js", function () { initialiseWidget(<%: navlink.id %>, "DBV<%: navlink.id %>", "DBV<%: navlink.Text %>", ""); });', 'ajax');</script>
						<%iRowNum += 1

						Case ElementType.TodaysEvents%>
						<li data-col="<%=iColNum %>" style="<%:sTileBackColourStyle%><%:sTileForeColourStyle%>" data-row="<%=iRowNum %>" data-sizex="2" data-sizey="1" class="linkspagebuttontext <%=sTileColourClass%> displayonly TELink">
							<div class="TETile <%=sTileColourClass%>">
								<p class="linkspagebuttontileIcon">
									<i class="icon-calendar"></i>
									<div class="TECount"></div>
								</p>
								<p>
									<a style="<%:sTileForeColourStyle%>" href="#"><%=FormatDateTime(Now, vbLongDate)%></a>
								</p>
								<div class="widgetplaceholder generaltheme">
									<div><i class="icon-calendar"></i></div>
									<a style="<%:sTileForeColourStyle%>" href="#">Today's Events</a>
								</div>
							</div>
							<div style="<%:sTileBackColourStyle%><%:sTileForeColourStyle%>" class="TEList <%=sTileColourClass%>">
								<p><span>Today's Events:</span></p>
								<table style="width: 100%;">
									<%											
										' ----------------------- DIARY LINKS -----------------------------
										Dim sErrorDescription As String = ""
										Dim iRecNum As Integer
										Dim iNumberOfEvents As Integer = 0
												
										' Create the reference to the DLL
										Dim objDiaryEvents As clsDiary = New clsDiary
										objDiaryEvents.SessionInfo = CType(Session("SessionContext"), SessionInfo)
										objDiaryEvents.CheckAccessToSystemEvents()
			
										Err.Clear()
										Dim mrstEventData = objDiaryEvents.GetDiaryData(False, Now.Date, Now.Date)
																									
										If (Err.Number() <> 0) Then
											sErrorDescription = "The Event Data could not be retrieved." & vbCrLf & FormatError(Err.Description)
										End If
										iRecNum = 0
												
										If sErrorDescription.Length = 0 Then
											If mrstEventData.Rows.Count Then
									%>
									<tr>
										<td colspan="2" style="font-weight: bold; font-size: small; border-bottom: 1px solid gray">Diary Links</td>
									</tr>
									<%    
										For Each objRow As DataRow In mrstEventData.Rows
													
									%>
									<tr>
										<td colspan="2" style="font-weight: normal; font-size: small"><%=objRow(3).ToString%></td>
									</tr>
									<%                
										iRecNum = iRecNum + 1
									Next
								End If

							End If
											
							iNumberOfEvents += iRecNum
											
							' ----------------------- OUTLOOK LINKS -----------------------------
							' Create the reference to the DLL
							Dim objOutlookEvents As HR.Intranet.Server.clsOutlookLinks = New HR.Intranet.Server.clsOutlookLinks
							objOutlookEvents.SessionInfo = CType(Session("SessionContext"), SessionInfo)
			
							Err.Clear()
							mrstEventData = objOutlookEvents.GetOutlookLinks()

							If (Err.Number <> 0) Then
								sErrorDescription = "The Outlook Links Data could not be retrieved." & vbCrLf & FormatError(Err.Description)
							End If
							iRecNum = 0
											
							If Len(sErrorDescription) = 0 Then
								If mrstEventData.Rows.Count > 0 Then
									%>
									<tr>
										<td colspan="2" style="font-weight: bold; font-size: small; border-bottom: 1px solid gray">Outlook Calendar Links</td>
									</tr>
									<%
										For Each objRow As DataRow In mrstEventData.Rows
									%>
									<tr>
										<td colspan="2" style="font-weight: normal; font-size: small"><%=Trim(objRow(2).ToString())%></td>
									</tr>
									<%
										iRecNum += 1
									Next
								End If


							End If
									
							iNumberOfEvents += iRecNum
											

							' ----------------------- TODAY'S ABSENCES -----------------------------
							' Create the reference to the DLL
							Dim objTodaysEvents As clsTodaysAbsence = New clsTodaysAbsence
							objTodaysEvents.SessionInfo = CType(Session("SessionContext"), SessionInfo)
				
							Err.Clear()
							mrstEventData = objTodaysEvents.GetTodaysAbsences(CleanNumeric(Session("TopLevelRecID")))
							iRecNum = 0
											
							If Len(sErrorDescription) = 0 And Not mrstEventData Is Nothing Then
								If mrstEventData.Rows.Count > 0 Then
									%>
									<tr>
										<td colspan="2" style="font-weight: bold; font-size: small; border-bottom: 1px solid gray">Today's Absences</td>
									</tr>
									<%             
												
										For Each objRow As DataRow In mrstEventData.Rows
													
									%>
									<tr>
										<td colspan="2" style="font-weight: normal; font-size: small"><%=Trim(objRow(0).ToString)%></td>
									</tr>
									<%                
										iRecNum = iRecNum + 1
									Next
								End If
								iNumberOfEvents += iRecNum
																																
							End If
									%>
								</table>
							</div>

							<div class="linkspagebuttontileIcon"><span>
								<p><%=iNumberOfEvents%></p>
								<p style="font-size: small;">Event<%If iNumberOfEvents <> 1 Then%>s<%End If%></p>
							</span></div>
						</li>
						<%
							iRowNum += 1
							
						Case ElementType.OrgChart
							sOnClick = "loadPartialView('OrgChart', 'home', 'workframe')"%>
						<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1" style="<%:sTileBackColourStyle%><%:sTileForeColourStyle%>" class="linkspagebuttontext <%=sTileColourClass%>" onclick="<%=sOnclick%>">
							<a style="<%:sTileForeColourStyle%>" href="#"><%: navlink.Text %></a>
							<p class="linkspagebuttontileIcon"><i class="icon-sitemap"></i></p>
						</li>
						<%iRowNum += 1%>
						
						<%Case Else%>
						<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1" style="<%:sTileBackColourStyle%><%:sTileForeColourStyle%>"
							class="linkspagebuttontext <%=sTileColourClass%> displayonly"><a style="<%:sTileForeColourStyle%>" href="#">
								<%: navlink.Text %></a></li>
						<%iRowNum += 1

					End Select

				End If

				tileCount += 1
			Next
								
			' Get the navigation hypertext links.
			For Each objNavLink In objNavigation.GetNavigationLinks(False, LinkType.Button)
	
				Dim sLinkText As New StringBuilder
				If objNavLink.Text1.Trim().Length > 0 Then sLinkText.Append(Html.Encode(objNavLink.Text1) & " ")
				sLinkText.Append(Html.Encode(objNavLink.Text2.Trim()))
				sText = sLinkText.ToString()
						
				If sLinkText.Length > 30 Then
					sText = sLinkText.ToString().Substring(0, 30) + "..."
				End If

		
				If objNavLink.LinkToFind = 0 Then
					sDestination = "linksMain?" & CStr(objNavLink.TableID) & "!" & CStr(objNavLink.ViewID)
			
					If objNavLink.SingleRecord = 1 Then
						sDestination = sDestination & "_0"
					Else
						sDestination = sDestination & "_" & CStr(Session("TopLevelRecID"))
					End If
				Else
					sDestination = "recordEditMain?multifind_0_" & CStr(objNavLink.TableID) & "!" & CStr(objNavLink.ViewID)
				End If
				If fFirstSeparator Then		' add a separator
					iRowNum = 1
					iColNum = 1
					If fFirstSeparator Then
						fFirstSeparator = False
					Else
						%>
					</ul>
				</div>
			</li>
		</ul>

		<%
		End If
		iSeparatorNum += 1
			
		%>

		<ul class="linkspagebuttonseparatorframe" id="linkspagebuttonseparatorframe_<%=iSeparatorNum %>">
			<li class="linkspagebutton-displaytype">
				<div class="wrapupcontainer linkspagebuttonseparator-bordercolour" style="">
					<div class="wrapuptext">
						<p class="linkspagebuttonseparator linkspagebuttonseparator-font linkspagebuttonseparator-colour linkspagebuttonseparator-size linkspagebuttonseparator-bold linkspagebuttonseparator-italics">Fixed Links</p>
					</div>
				</div>
				<div class="gridster buttonlinkcontent" id="gridster_buttonlink_<%=tileCount%>">
					<ul>
						<%
						End If
						If iRowNum > iMaxRows Then
							iColNum += 1
							iRowNum = 1
						%>
						<script type="text/javascript">
							$("#linkspagebuttonseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
							$("#linkspagebuttonseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
						</script>
						<%
						End If

						%>
						<li class="linkspagebuttontext Colour4" data-col="<%=iColNum %>" data-row="<%=iRowNum %>"
							data-sizex="1" data-sizey="1" onclick="goURL('<%=sDestination%>', 0, false)">
							<a class="linkspagebutton-displaytype linkspagebuttontext-alignment linkspagebutton-colourtheme" href="#"><%=sText%></a>
							<p class="linkspagebuttontileIcon"><i class="icon-external-link-sign"></i></p>
						</li>
						<%
							iRowNum += 1
							tileCount += 1
						Next						
						
						
			If Not fFirstSeparator Then%>
					</ul>
				</div>
			</li>
		</ul>
		<%
		End If
		%>
	</div>	
	</div>    <%-- close linkspagebuttonbox div--%>
</div>

<%If Model.NumberOfLinks > 0 Then%>
<div class="dropdownlinks">
	<ul class="dropdownlinkseparatorframe" id="dropdownlinkseparatorframe_<%=iSeparatorNum %>">
		<li class="dropdownlink-displaytype">
			<%If (Model.NavigationLinks.FindAll(Function(n) n.LinkType = LinkType.DropDown).Count + objNavigation.GetNavigationLinks(False, LinkType.DropDown).Count) > 0 Then%>
			<p class="dropdownlinkseparator">Dropdown links:</p>
			<div class="gridster dropdownlinkcontent" id="gridster_DropdownLinks">
				<ul class="DropDownListMenu">
					<%iRowNum = 1
						iColNum = 1

						For Each navlink In Model.NavigationLinks.FindAll(Function(n) n.LinkType = LinkType.DropDown)
						
							Dim sTileColourClass = "Colour" & CStr(CInt(Math.Ceiling(Rnd() * 7)))
							Dim sValue As String, sUtilityType As String, sUtilityID As String, sUtilityDef As String, sUtilityName As String
						
							If Len(navlink.AppFilePath) > 0 Then
								sAppFilePath = Replace(navlink.AppFilePath, "\", "\\")
								sAppParameters = Replace(navlink.AppParameters, "\", "\\")
			
								sValue = "5_" & sAppFilePath & "_" & sAppParameters
								sOnClick = "goDropLink('" + sValue + "')"

							ElseIf navlink.Element_Type = ElementType.OrgChart Then
								sValue = "6_OrgChart"
								sOnClick = "loadPartialView('OrgChart', 'home', 'workframe')"
							
							
							ElseIf Len(navlink.URL) > 0 Then
								sURL = Html.Encode(navlink.URL)
								sURL = Replace(sURL, "'", "\'")

								If navlink.NewWindow = True Then
									sNewWindow = "1"
								Else
									sNewWindow = "0"
								End If
		 
								sValue = "0_" & sNewWindow & "_" & sURL & "_" & Server.HtmlEncode(navlink.Text)
								sOnClick = "goDropLink('" + sValue + "')"
							
							Else
								If navlink.UtilityID > 0 Then
									sUtilityType = CStr(navlink.UtilityType)
									sUtilityID = CStr(navlink.UtilityID)
									sUtilityName = navlink.Text.Replace("'", "")
									sUtilityDef = sUtilityType & "_" & sUtilityID & "_" & sUtilityName
				
									sValue = "2_" & sUtilityDef
				
								Else
									sLinkKey = "recedit" & _
										"_" & Session("TopLevelRecID").ToString() & _
										"_" & navlink.ID
					
									sValue = "1_" & sLinkKey
				
								End If
							
								sOnClick = "goDropLink('" + sValue + "')"
							
							End If

							If iRowNum > iMaxRows Then	 ' start a new column if required (affects tiles only)
								iColNum += 1
								iRowNum = 1%>
					<script type="text/javascript">
						$("#dropdownlinkseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
						$("#dropdownlinkseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
					</script>
					<%End If%>
					<li class="dropdownlinktext <%=sTileColourClass%>" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
						data-sizey="1" onclick="<%=sOnclick%>">
						<p class="dropdownlinktileIcon">
							<i class="icon-external-link"></i>
						</p>
						<p>
							<a href="#" data-ddlvalue="<%=sValue%>">
								<%If navlink.Text.Length > 30 Then
										navlink.Text = navlink.Text.Substring(0, 30) + "..."
									End If
									%>
								<%: navlink.Text %>
							</a>
						</p>
					</li>
					<%iRowNum += 1

					Next
			
					For Each objNavLink In objNavigation.GetNavigationLinks(False, LinkType.DropDown)
	
						Dim sLinkText As New StringBuilder
						If objNavLink.Text1.Trim().Length > 0 Then sLinkText.Append(Html.Encode(objNavLink.Text1) & " ")
						sLinkText.Append(Html.Encode(objNavLink.Text2.Trim()))
						
						sText = sLinkText.ToString()
						
						If sLinkText.Length > 30 Then
							sText = sLinkText.ToString().Substring(0, 30) + "..."
						End If
						
						Dim sValue As String = ""

		
						If objNavLink.LinkToFind = 0 Then
							sValue = "7_" & CStr(objNavLink.TableID) & "!" & CStr(objNavLink.ViewID)
							'sDestination = "linksMain?" & CStr(objNavLink.TableID) & "!" & CStr(objNavLink.ViewID)
			
							If objNavLink.SingleRecord = 1 Then
								'sDestination = sDestination & "_0"
								sValue &= "_0"
							Else
								'sDestination = sDestination & "_" & CStr(Session("TopLevelRecID"))
								sValue &= "_" & CStr(Session("TopLevelRecID"))
							End If
						Else
							'sDestination = "recordEditMain?multifind_0_" & CStr(objNavLink.TableID) & "!" & CStr(objNavLink.ViewID)
							sValue = "4_" & CStr(objNavLink.TableID) & "!" & CStr(objNavLink.ViewID)
						End If
						
						sOnClick = "goDropLink('" + sValue + "')"

						If iRowNum > iMaxRows Then
							iColNum += 1
							iRowNum = 1
						%>
						<script type="text/javascript">
							$("#dropdownlinkseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
							$("#dropdownlinkseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
						</script>
						<%End If%>
					<li class="dropdownlinktext Colour4" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
						data-sizey="1" onclick="<%=sOnClick%>">
						<p class="dropdownlinktileIcon">
							<i class="icon-external-link"></i>
						</p>
						<p>
							<a href="#" data-ddlvalue="<%=sValue%>">
								<%=sText %></a>
						</p>
					</li>					
						<%
							iRowNum += 1
							tileCount += 1
						Next						

%>
				</ul>
				<a class="DropLinkGoText" style="text-decoration: none; margin-left: 10px;" href="#" onclick="goDropLink()">Go...</a>
			</div>
		  <%End If%>
		</li>

	</ul>
</div>

<%End If%>




<%If Session("CurrentLayout").ToString() = Layout.tiles.ToString() Then%>
<%If Model.DocumentDisplayLinkCount > 0 Then%>
<div class="docdisplaylinks">
	<ul class="docdisplaylinkseparatorframe" id="docdisplaylinkseparatorframe_<%=iSeparatorNum %>">
		<li class="docdisplaylink-displaytype">
			<p class="docdisplaylinkseparator">Document Display links:</p>
			<div class="gridster docdisplaylinkcontent" id="gridster_docdisplayLinks">
				<ul class="DocDisplayListMenu">
					<%iRowNum = 1
						iColNum = 1

						For Each navlink In Model.NavigationLinks.FindAll(Function(n) n.LinkType = LinkType.DocumentDisplay)
						
							Dim sTileColourClass = "Colour" & CStr(CInt(Math.Ceiling(Rnd() * 7)))
							Dim sValue As String
							
							sURL = Html.Encode(navlink.DocumentFilePath)
							sURL = Replace(sURL, "'", "\'")

							sNewWindow = "1"
		 
							sValue = "0_" & sNewWindow & "_" & sURL
							sOnClick = "goDropLink('" + sValue + "')"
							

							If iRowNum > iMaxRows Then	 ' start a new column if required (affects tiles only)
								iColNum += 1
								iRowNum = 1%>
					<script type="text/javascript">
						$("#docdisplaylinksseparatorframe<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
						$("#docdisplaylinksseparatorframe<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
					</script>
					<%End If%>
					<li class="docdisplaylinktext <%=sTileColourClass%>" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
						data-sizey="1" onclick="<%=sOnclick%>">
						<p class="docdisplaylinktileIcon">
							<i class="icon-external-link"></i>
						</p>
						<p>
							<a href="#" data-ddlvalue="<%=sValue%>">
								<%If navlink.Text.Length > 30 Then
										navlink.Text = navlink.Text.Substring(0, 30) + "..."
									End If
									%>
								<%: navlink.Text %>
							</a>
						</p>
					</li>
					<%iRowNum += 1

					Next				

%>
				</ul>				
			</div>
		</li>

	</ul>
</div>
		</div>
	</div>
<%End If%>

<%Else%>

		</div>
	</div>

<div id="documentDisplay" class="ui-widget-content">
	<div id="divResize">
		<img id="splitToggle" src="" alt="Show Document Display"
			onclick="setDocumentDisplayVisible();" />
	</div>
	<div id="documentDisplayContent" rowspan="4" width="340px" valign="top" nowrap="nowrap">
		<%Html.RenderPartial("~/Views/Home/documentDisplay.ascx")%>
	</div>
</div>

<%End If%>

<div id="pwfs"><%Response.Write(_PendingWorkflowStepsHTMLTable.ToString())%></div>

<div id="frmMenuInfo" >
	<%
		Response.Write("<INPUT type=""hidden"" id=txtDefaultStartPage name=txtDefaultStartPage value=""" & Replace(Session("DefaultStartPage"), """", "&quot;") & """>")
	%>

	<input type="hidden" id="txtIsDMIUser" name="txtIsDMIUser" value=<%= objSession.LoginInfo.IsDMIUser%>>
	<input type="hidden" id="txtIsSSIUser" name="txtIsSSIUser" value='<%= objSession.LoginInfo.IsSSIUser%>'>

	<input type="hidden" id="txtPersonnel_EmpTableID" name="txtPersonnel_EmpTableID" value='<%:SettingsConfig.Personnel_EmpTableID%>'>

	<input type="hidden" id="txtTB_EmpTableID" name="txtTB_EmpTableID" value='<%=session("TB_EmpTableID")%>'>
	<input type="hidden" id="txtTB_CourseTableID" name="txtTB_CourseTableID" value='<%=session("TB_CourseTableID")%>'>
	<input type="hidden" id="txtTB_CourseCancelDateColumnID" name="txtTB_CourseCancelDateColumnID" value='<%=session("TB_CourseCancelDateColumnID")%>'>
	<input type="hidden" id="txtWaitListOverRideColumnID" name="txtWaitListOverRideColumnID" value='<%=session("TB_WaitListOverRideColumnID")%>'>
	<input type="hidden" id="txtTB_TBTableID" name="txtTB_TBTableID" value='<%=session("TB_TBTableID")%>'>
	<input type="hidden" id="txtTB_TBTableSelect" name="txtTB_TBTableSelect" value='<%=session("TB_TBTableSelect")%>'>
	<input type="hidden" id="txtTB_TBTableInsert" name="txtTB_TBTableInsert" value='<%=session("TB_TBTableInsert")%>'>
	<input type="hidden" id="txtTB_TBTableUpdate" name="txtTB_TBTableUpdate" value='<%=session("TB_TBTableUpdate")%>'>
	<input type="hidden" id="txtTB_TBStatusColumnID" name="txtTB_TBStatusColumnID" value='<%=session("TB_TBStatusColumnID")%>'>
	<input type="hidden" id="txtTB_TBStatusColumnUpdate" name="txtTB_TBStatusColumnUpdate" value='<%=session("TB_TBStatusColumnUpdate")%>'>
	<input type="hidden" id="txtTB_TBCancelDateColumnID" name="txtTB_TBCancelDateColumnID" value='<%=session("TB_TBCancelDateColumnID")%>'>
	<input type="hidden" id="txtTB_TBCancelDateColumnUpdate" name="txtTB_TBCancelDateColumnUpdate" value='<%=session("TB_TBCancelDateColumnUpdate")%>'>
	<input type="hidden" id="txtTB_TBStatusPExists" name="txtTB_TBStatusPExists" value='<%=session("TB_TBStatusPExists")%>'>
	<input type="hidden" id="txtTB_WaitListTableID" name="txtTB_WaitListTableID" value='<%=session("TB_WaitListTableID")%>'>
	<input type="hidden" id="txtTB_WaitListTableInsert" name="txtTB_WaitListTableInsert" value='<%=session("TB_WaitListTableInsert")%>'>
	<input type="hidden" id="txtTB_WaitListTableDelete" name="txtTB_WaitListTableDelete" value='<%=session("TB_WaitListTableDelete")%>'>
	<input type="hidden" id="txtTB_WaitListCourseTitleColumnID" name="txtTB_WaitListCourseTitleColumnID" value='<%=session("TB_WaitListCourseTitleColumnID")%>'>
	<input type="hidden" id="txtTB_WaitListCourseTitleColumnUpdate" name="txtTB_WaitListCourseTitleColumnUpdate" value='<%=session("TB_WaitListCourseTitleColumnUpdate")%>'>
	<input type="hidden" id="txtTB_WaitListCourseTitleColumnSelect" name="txtTB_WaitListCourseTitleColumnSelect" value='<%=session("TB_WaitListCourseTitleColumnSelect")%>'>
	<input type="hidden" id="txtPrimaryStartMode" name="txtPrimaryStartMode" value='<%=session("PrimaryStartMode")%>'>
	<input type="hidden" id="txtHistoryStartMode" name="txtHistoryStartMode" value='<%=session("HistoryStartMode")%>'>
	<input type="hidden" id="txtLookupStartMode" name="txtLookupStartMode" value='<%=session("LookupStartMode")%>'>
	<input type="hidden" id="txtQuickAccessStartMode" name="txtQuickAccessStartMode" value='<%=session("QuickAccessStartMode")%>'>
	<input type="hidden" id="txtDesktopColour" name="txtDesktopColour" value='<%=session("DesktopColour")%>'>

	<input type="hidden" id="txtWFEnabled" name="txtWFEnabled" value='<%=session("WF_Enabled")%>'>
	<input type="hidden" id="txtWFOutOfOfficeEnabled" name="txtWFOutOfOfficeEnabled" value='<%=session("WF_OutOfOfficeConfigured")%>'>

	<input type="hidden" id="txtDoneDatabaseMenu" name="txtDoneDatabaseMenu" value="0">
	<input type="hidden" id="txtDoneQuickEntryMenu" name="txtDoneQuickEntryMenu" value="0">
	<input type="hidden" id="txtDoneTableScreensMenu" name="txtDoneTableScreensMenu" value="0">

</div>

<div id="workflowDisplay" class="absolutefull" style="display: none; background-color: transparent; text-align: center;">
	<div class="pageTitleDiv" style="text-align: left;">		
		<a href='<%=Url.Action("MainSSI", "Home")%>' title='Back'>
			<i class='pageTitleIcon icon-circle-arrow-left'></i>
		</a>
		<span class="pageTitle">Workflow</span>
	</div>

	<iframe id="externalContentFrame" style="width: 90%; margin: 0 auto;"></iframe>
</div>

<form action="WorkAreaRefresh" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
		<%=Html.AntiForgeryToken()%>
</form>

<script type="text/javascript">	
	function setupTiles() {

		$('.gridster').each(function () {
			var id = $(this).attr('id');
			griditup(id, true);
		});

		//add mousewheel scrollability to the main content window
		if ('<%=session("isMobileDevice")%>' != "True") {
		$(".DashContent").mousewheel(function (event, delta) {
			//this.scrollLeft -= (delta * 30);
			//event.preventDefault();
		});
	} else {
		$('.DashContent').css('overflow-x', 'auto');
	}
}

	//Display Pending Workflow Steps if appropriate
	if (('<%=fWFDisplayPendingSteps%>' == 'True') && (Number('<%=_StepCount%>') > 0) && ('<%=Session("LoggingIn")%>' == '')) {		
		relocateURL('WorkflowPendingSteps', 0);
	}

	$(".sp-container.sp-hidden").css("display", "none"); //The color picker plugin sometimes leaves visible bits; remove them

	$(document).ready(function () {

		$("#fixedlinksframe").show();

		$("#workframe").attr("data-framesource", "linksmain");
		$('#workframe').css('height', '100%');
		//$('#SSILinksFrame').css('height', '100%');
		
		showDefaultRibbon();

		menu_setVisibleMenuItem('mnutoolFixedWorkflowOutOfOffice', "<%:ViewData("showOutOfOffice")%>");

		$("#toolbarHome").show();
		$("#toolbarHome").click();

		refreshPendingWorkflowTiles();

		if (window.currentLayout == "tiles") {
			setupTiles();

			//Reduce the dbvalue text size to fit its tile if too big.
			$('.DBValue').each(function () {
				var originalFontSize = 26;
				var sectionWidth = $(this).width();

				var spanWidth = $(this).find('span').width();
				var newFontSize = (sectionWidth / spanWidth) * originalFontSize;
				if (newFontSize < originalFontSize) {
					$(this).find('span').css({ "font-size": newFontSize, "line-height": newFontSize / 1.2 + "px" });
				}
			});
			
			//Limit number of characters in a tile to 30 when in Tiles mode
			//$('.DBValueCaption').each(function () {
			//	alert($(this).value);
			//	if( $(this).length > 30) {
			//		var output = $(this).substring(0, 30);
			//		var ellipse = "...";
			//		$(this)["text html"] = concat(output, ellipse);
			//	}
			//});

		} else {
			// for wireframe layout, convert the dropdownlinks to a <select> element
			$(function () {
				$('ul.DropDownListMenu').each(function () {
					var $select = $('<select class="DropdownlistSelect"/>');

					$(this).find('a').each(function () {
						var $option = $('<option />');
						$option.attr('value', $(this).attr('data-DDLValue')).html($(this).html());
						$select.append($option);
					});

					$(this).replaceWith($select);


				});
			});

			//Show document display (not tiles)
			//get cookie...
			var showDocBar = window.getCookie('displayDocBar');
			if (showDocBar.length == 0) showDocBar = 'true';

			if (showDocBar == 'true') {
				setDocumentDisplayVisible('true');
			} else {
				setDocumentDisplayVisible('false');
			}

		}

		if ((window.currentLayout == "wireframe") || (window.currentLayout == "winkit")) {
			//set up the classes 

			$(".hypertextlinks").addClass("ui-accordion ui-widget ui-helper-reset");
			$(".ButtonLinkColumn").addClass("ui-accordion ui-widget ui-helper-reset");
			$(".wrapupcontainer").addClass("ui-accordion-header ui-helper-reset ui-state-default ui-accordion-icons ui-accordion-header-default ui-state-default ui-corner-top");
			//$(".hypertextlinkcontent").addClass("ui-accordion-content ui-helper-reset ui-widget-content ui-corner-bottom ui-accordion-content-active");
			//menu style
			//$(".hypertextlinkcontent>ul").addClass("ui-menu ui-widget ui-widget-content ui-corner-all");
			$('.hypertextlinkcontent>ul').menu();
			$('.hypertextlinkcontent>ul').removeClass('ui-corner-all').addClass('ui-corner-bottom');
			$('.buttonlinkcontent>ul').menu();
			$('.buttonlinkcontent>ul').removeClass('ui-corner-all').addClass('ui-corner-bottom');

			$('.DashContent').addClass("ui-widget ui-widget-content");
			//$('.ViewDescription').addClass('ui-widget ui-widget-content');

		}

		// This replaces the big fat grey scrollbar with the nice thin dark one. (HRPRO-2952)
		if (('<%=session("isMobileDevice")%>' !== "True") && (window.currentLayout == 'tiles')) {
			setTimeout('$(".DashContent").mCustomScrollbar({ horizontalScroll: true, theme:"dark-thin", autoHideScrollbar: true, scrollInertia: 150 });', 750);
		} else {
			$('.DashContent').attr('overflow', 'auto');
		}


		//resize columns that have wide tiles
		$("li[data-sizex='2']").each(function () {

			var ulelement = $(this).closest('.linkspagebuttonseparatorframe');

			if ($(ulelement).hasClass('cols2')) {
				$(ulelement).removeClass('cols2');
				$(ulelement).addClass('cols3');
			} else if ($(ulelement).hasClass('cols3')) {
				$(ulelement).removeClass('cols3');
				$(ulelement).addClass('cols4');
			} else if ($(ulelement).hasClass('cols4')) {
				$(ulelement).removeClass('cols4');
				$(ulelement).addClass('cols5');
			} else {
				//no cols class, so add one.
				$(ulelement).addClass('cols2');
			}

		});


		//display view details
		$('.ViewDescription p').html('<%=Html.Encode(Session("ViewDescription").ToString())%>');

		if (window.currentLayout == "tiles") {
			$('header').css('background-image', 'none').css('background-color', 'none');			
		}

		$('header').show();

	});

</script>

<%Session("LoggingIn") = False%>


<div id="chartPopout"></div>
<script type="text/javascript">	
	function popoutchart(multiAxis, showLegend, showGrid, showValues, stackSeries, showPercentages, chartType, chartTableID, chartColumnID, chartFilterID,
				aggregateType, elementType, chartTableID2, chartColumnID2, chartTableID3, chartColumnID3, sortOrderID, sortDirection, colourID, chartText, themeName,
				chartColumnName, chartColumnName2) {
		
		var chartModel = {
			multiAxis: multiAxis,
			showLegend: showLegend,
			showGrid: showGrid,
			showValues: showValues,
			stackSeries: stackSeries,
			showPercentages: showPercentages,
			chartType: chartType,
			chartTableID: chartTableID,
			chartColumnID: chartColumnID,
			chartFilterID: chartFilterID,
			aggregateType: aggregateType,
			elementType: elementType,
			chartTableID2: chartTableID2,
			chartColumnID2: chartColumnID2,
			chartTableID3: chartTableID3,
			chartColumnID3: chartColumnID3,
			sortOrderID: sortOrderID,
			sortDirection: sortDirection,
			colourID: colourID,
			chartText: chartText,
			themeName: themeName,
			chartColumnName: chartColumnName,
			chartColumnName2: chartColumnName2
		};
		
		$('#chartPopout').dialog({
			autoOpen: true,
			width: 700,
			height: 560,
			resizable: true,
			title: chartText,
			modal: true,
			open: function (event, ui) {
				//Load the showChart action which will return 
				// the partial view _showChart
				$(this).load("<%:Url.Action("ShowChart", "Home")%>", chartModel);
			},
			resize: function () {
				var doit;				
					clearTimeout(doit);
					doit = setTimeout(loadChart, 100);				
			},
			buttons: {
					"Close": function () {
						$(this).dialog("close");
					}
			},
			minHeight: 300,
			minWidth: 300
		});		
	}
	//set up window variable for ismobilebrowser.
	window.isMobileBrowser = '<%=Session("isMobileDevice").ToString().ToLower()%>';
	
</script>

