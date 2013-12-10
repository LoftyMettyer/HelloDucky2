<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader("Pragma", "no-cache")%>
<% Response.Expires = -1 %>
<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of DMI.NET.NavLinksViewModel)" %>
<%@Import namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>

<%-- For other devs: Do not remove below line. --%>
<%="" %>
<%-- For other devs: Do not remove above line. --%>

<link id="SSIthemeLink" href="" rel="stylesheet" type="text/css" />
<link href="<%:Url.Content("~/Content/jquery.mCustomScrollbar.min.css")%>" rel="stylesheet" />
<link href="<%= Url.LatestContent("~/Content/jquery.gridster.css")%>" rel="stylesheet" type="text/css" />
<script src="<%: Url.LatestContent("~/Scripts/jquery/jquery.gridster.js")%>" type="text/javascript"></script>
<script src="<%: Url.LatestContent("~/Scripts/jquery/jquery.mousewheel.js")%>" type="text/javascript"></script>
<script src="<%:Url.Content("~/Scripts/jquery/jquery.mCustomScrollbar.min.js")%>"></script>

<%Session("recordID") = 0%>
<%Dim fWFDisplayPendingSteps As Boolean = True%>

	<script type="text/javascript">

		dragged = 0;
		function loadjscssfile(filename, filetype) {

			var fileref;

			if (filetype == "ajax") {
				fileref = document.createElement("script");
				fileref.setAttribute("type", "text/javascript");
				fileref.innerHTML = filename;
			}
			else if (filetype == "js") { //if filename is a external JavaScript file
				fileref = document.createElement('script');
				fileref.id = "myScript1";
				fileref.setAttribute("type", "text/javascript");
				fileref.setAttribute("src", filename);
			}
			else if (filetype == "css") { //if filename is an external CSS file
				fileref = document.createElement("link");
				fileref.setAttribute("rel", "stylesheet");
				fileref.setAttribute("type", "text/css");
				fileref.setAttribute("href", filename);
			}
			if (typeof fileref != "undefined") {
				document.getElementsByTagName("head")[0].appendChild(fileref);
			}
		}

		//function showPWFS() {
		//	$('#pwfs').dialog('open');
		//}

		function popoutchart(MultiAxis, Chart_ShowLegend, Chart_ShowGrid, Chart_ShowValues, Chart_StackSeries, Chart_ShowPercentages, iChart_Type, iChart_TableID,iChart_ColumnID, iChart_FilterID, iChart_AggregateType, iChart_ElementType,iChart_TableID_2,iChart_ColumnID_2,iChart_TableID_3,  iChart_ColumnID_3, iChart_SortOrderID,  iChart_SortDirection,iChart_ColourID) {
			
			var windowHeight = 500;
			var windowWidth = 500;
			
			var w = window.open('Chart', '_blank', 'width=' + windowWidth + ', height=' + windowHeight + ',location=no,resizable=yes,toolbar=no,titlebar=no,menubar=no');
			w.document.open();
			w.document.write('<div style="width: 100%; height: 100%;"><img id="chartImage" style="" src="" alt="Chart" />');
			w.document.write('</div>');
			w.document.write('<div style="position: fixed; bottom: 0;">');
			w.document.write('<table align="center" border="solid 1px" bgcolor="#cccccc">');
			w.document.write("<tr style='font-family:Verdana;font-size:x-small'>");
			w.document.write("<td>");
			w.document.write('<input type="button" style="" value="Redraw chart" onclick="loadChart();"/>');
			w.document.write('</td>');
			w.document.write("<td>");
			w.document.write("Chart Type: <select id='selChartType'>");
			w.document.write("<option value=\"0\"" + ((iChart_Type == 0) ? " selected " : "") + ">3D Bar</option>");
			w.document.write("<option value=\"1\"" + ((iChart_Type == 1) ? " selected " : "") + ">2D Bar</option>");
			w.document.write("<option value=\"2\"" + ((iChart_Type == 2) ? " selected " : "") + ">3D Line</option>");
			w.document.write("<option value=\"3\"" + ((iChart_Type == 3) ? " selected " : "") + ">2D Line</option>");
			w.document.write("<option value=\"4\"" + ((iChart_Type == 4) ? " selected " : "") + ">3D Area</option>");
			w.document.write("<option value=\"5\"" + ((iChart_Type == 5) ? " selected " : "") + ">2D Area</option>");
			w.document.write("<option value=\"6\"" + ((iChart_Type == 6) ? " selected " : "") + ">3D Step</option>");
			w.document.write("<option value=\"7\"" + ((iChart_Type == 7) ? " selected " : "") + ">2D Step</option>");
			w.document.write("<option value=\"14\"" + ((iChart_Type == 14) ? " selected " : "") + ">2D Pie</option>");
			w.document.write("<option value=\"16\"" + ((iChart_Type == 16) ? " selected " : "") + ">2D XY</option>");
			w.document.write("</select></td>");
			w.document.write("<td>Show Legend:<input id='chkshowLegend' type='checkbox' ");
			w.document.write((Chart_ShowLegend == "True") ? "Checked " : "");
			w.document.write("/></td>");
			w.document.write("<td>Stack Series:<input id='chkstackSeries' type='checkbox' ");
			w.document.write((Chart_StackSeries == 'True') ? "Checked " : "");
			w.document.write("/></td> ");
			w.document.write("<td>Show Gridlines:<input id='chkShowGrid' type='checkbox' ");
			w.document.write((Chart_ShowGrid == 'True') ? "Checked " : "");
			w.document.write(" /></td> ");
			w.document.write("<td>Show Values As:<select id='lstValueType' >");
			w.document.write("  <option value=\"Values\"" + (Chart_ShowPercentages == 'False' ? " selected " : "") + ">Values</option>");
			w.document.write("  <option value=\"Percentages\"" + (Chart_ShowPercentages == 'True' ? " selected " : "") + ">Percentages</option>");
			w.document.write("</select></td>");
			w.document.write("<td><input value='Print' id='btnPrint' type='button' onClick='window.print()'/></td>");
			w.document.write("</tr>");
			w.document.write("</table>");
			w.document.write('</div>');
			w.document.write('<scri');
			w.document.write('pt type="text/javascript">');
			w.document.write('function loadChart() {');
			w.document.write('var windowHeight = window.innerHeight - 80;'); //reduce height for toolbar.
			w.document.write('var chartType = document.getElementById("selChartType").value;');
			w.document.write('var chartShowLegend = (document.getElementById("chkshowLegend").checked==true);');
			w.document.write('var chartStackSeries = (document.getElementById("chkstackSeries").checked==true);');
			w.document.write('var chartShowGridlines = (document.getElementById("chkShowGrid").checked==true);');			
			w.document.write('var chartShowPercentages = (document.getElementById("lstValueType").value == "Percentages");');
			if (MultiAxis == 'True') {
				w.document.write('var psURL = "GetMultiAxisChart?');
			} else {
				w.document.write('var psURL = "GetChart?');
			}
			w.document.write('height=" + windowHeight + "');
			w.document.write('&width=" + window.innerWidth + "');
			w.document.write('&ShowLegend=" + chartShowLegend + "');
			w.document.write('&DottedGrid=" + chartShowGridlines + "');
			w.document.write('&ShowValues=true');
			w.document.write('&Stack=" + chartStackSeries + "');
			w.document.write('&ShowPercent=" + chartShowPercentages + "');
			w.document.write('&ChartType=" + chartType + "');
			w.document.write('&TableID=' + iChart_TableID);
			w.document.write('&ColumnID=' + iChart_ColumnID);
			w.document.write('&FilterID=' + iChart_FilterID);
			w.document.write('&AggregateType=' + iChart_AggregateType);
			w.document.write('&ElementType=' + iChart_ElementType);
			if (MultiAxis == 'True') {
				w.document.write('&TableID_2=' + iChart_TableID_2);
				w.document.write('&ColumnID_2=' + iChart_ColumnID_2);
				w.document.write('&TableID_3=' + iChart_TableID_3);
				w.document.write('&ColumnID_3=' + iChart_ColumnID_3);
			}
			w.document.write('&SortOrderID=' + iChart_SortOrderID);
			w.document.write('&SortDirection=' + iChart_SortDirection);
			w.document.write('&ColourID=' + iChart_ColourID + '";');
			w.document.write('document.getElementById("chartImage").src = psURL;');
			w.document.write('}');
			w.document.write('loadChart();');
			w.document.write('</scri');
			w.document.write('pt>');
			w.document.close();

			

		}


		function refreshPendingWorkflowTiles() {
			//Add pending worklow tiles if in tiles mode
			if ((window.currentLayout == "tiles") && ($("#PendingStepsTable_Dash td").length > 0)) {				
				$('.pendingworkflowlinks').show();
				var rowNumber = 1;
				$("#PendingStepsTable_Dash tr td:nth-child(1)").each(function () {
					if (rowNumber > 4) return false;

					var desc = $(this).html();
					var name = $(this).next().next().html();
					if (desc.substring(0, name.length) === name) {
						//description starts with 'name' string, so remove it.
						desc = desc.substr(name.length + 2); // 2chars for the dash.
					}

					var url = 'launchWorkflow(\'' + $(this).next().html() + '\', \'' + name + '\')';
					
					var lihtml = '<li class="pendingworkflowtext White" data-col="1" data-row="' + rowNumber + '" ';
					lihtml += 'data-sizex="2" data-sizey="1" onclick="' + url + '">';
					lihtml += '<a href="#">';
					lihtml += '<span class="pendingworkflowname">' + name + '</span>';
					lihtml += '<br />';
					lihtml += '<span class="pendingworkflowdesc">' + desc + '</span>';
					lihtml += '</a>';
					lihtml += '<p class="pendingworkflowtileIcon"><i class="icon-adjust"></i></p>';
					lihtml += '</li>';

					$('#pendingworkflowstepstiles').append(lihtml);
					rowNumber += 1;
				});
			}
		}


		function setDocumentDisplayVisible(newSetting) {
			
			//are we toggling?
			if (newSetting == undefined) newSetting = (($("#documentDisplay").width() < 10)?'true':'false');
			
			if (newSetting == 'true') {
				//show the bar
				$("#documentDisplay").animate({ width: '340px' }, 350);
				$('#splitToggle').attr('src', '../Content/images/splitterRight.bmp');

				window.setCookie('displayDocBar', 'true', 365);

			} else {
				//hide the bar
				$("#documentDisplay").animate({ width: '6px' }, 350);
				$('#splitToggle').attr('src', '../Content/images/splitterLeft.bmp');
				
				window.setCookie('displayDocBar', 'false', 365);
			}

		}


		$(document).ready(function () {

			$("#fixedlinksframe").show();
			
			//Hide DMI button for non-IE browsers			
			if(('True' !== '<%=Session("MSBrowser")%>') && ('TRUE' == '<%=Session("DMIRequiresIE")%>')) $('#mnutoolFixedOpenHR').hide();

			showDefaultRibbon();
			$("#toolbarHome").show();
			$("#toolbarHome").click();


			$("#workframe").attr("data-framesource", "linksmain");
			$('#workframe').css('height', '100%');
			//$('#SSILinksFrame').css('height', '100%');

			refreshPendingWorkflowTiles();

			if (window.currentLayout == "tiles") {
				setupTiles();
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

			if (window.currentLayout == "wireframe") {
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


			//Load Poll.asp, then reload every 30 seconds to keep
			//session alive, and check for server messages.
			loadPartialView("poll", "home"); // first time
			// re-call the function each 30 seconds
			window.setInterval("loadPartialView('poll', 'home')", 30000);

			// This replaces the big fat grey scrollbar with the nice thin dark one. (HRPRO-2952)
			if ('<%=session("isMobileDevice")%>' != "True") {				
				setTimeout('$(".DashContent").mCustomScrollbar({ horizontalScroll: true, theme:"dark-thin" });', 500);
			}
			

			//resize columns that have wide tiles
			$("li[data-sizex='2']").each(function () {

				var ulelement = $(this).closest('.linkspagebuttonseparatorframe');

				if ($(ulelement).hasClass('cols2')) {
					$(ulelement).removeClass('cols2');
					$(ulelement).addClass('cols3');
				}
				else if ($(ulelement).hasClass('cols3')) {
					$(ulelement).removeClass('cols3');
					$(ulelement).addClass('cols4');
				}
				else if ($(ulelement).hasClass('cols4')) {
					$(ulelement).removeClass('cols4');
					$(ulelement).addClass('cols5');
				} else {
					//no cols class, so add one.
					$(ulelement).addClass('cols2');
				}

			});
			

			//display view details
			$('.ViewDescription p').text('<%=Session("ViewDescription")%>');

		});

		function setupTiles() {
			//apply the gridster functionality.
			//griditup(true);

			$('.gridster').each(function() {
				var id = $(this).attr('id');
				griditup(id, true);
			});
			
			//add mousewheel scrollability to the main content window
			if ('<%=session("isMobileDevice")%>' != "True") {
				$(".DashContent").mousewheel(function(event, delta) {
					this.scrollLeft -= (delta * 30);
					event.preventDefault();
				});
			} else {
				$('.DashContent').css('overflow-x', 'auto');
			}
		}

		function griditup(objectID, mode) {
			if (mode == true) {
				$("#" + objectID + " > ul").gridster({
					widget_margins: [5, 5],
					widget_base_dimensions: [120, 120],
					min_rows: 4,
					min_cols: 1,
					avoid_overlapped_widgets: true,
					draggable: {
						start: function (event, ui) {
							//dragged = 1;
						}
					}
				});
				
				var gridster = $("#" + objectID + " > ul").gridster().data('gridster');
				gridster.disable();

			}
		}

		function changeLayout(newLayoutName) {

			setCookie('Intranet_Layout', newLayoutName, 365);
			if (newLayoutName == "winkit") {
				setCookie('Intranet_Theme', "white", 365);
			} else {
				setCookie('Intranet_Theme', "blue", 365);
			}

			$(".DashContent").fadeOut("slow");

			$(".DashContent").promise().done(function () {
				window.location = "Main?SSIMode=True";
			});
		}

		function changeTheme(newThemeName) {
			
			$("link[id=SSIthemeLink]").attr({ href: "../Content/themes/" + newThemeName + "/jquery-ui.min.css" });
			setCookie('Intranet_Wireframe_Theme', newThemeName, 365);
		}

		function applyImportedTheme(newValue) {
			if (newValue == false) {
				$("link[id=WireframethemeLink]").attr({ href: "" });
			} else {
				$("link[id=WireframethemeLink]").attr({ href: "../Content/DashboardStyles/themes/upgraded.css" });
			}

			setCookie('Apply_Wireframe_Theme', newValue, 365);

		}


		//NPG20082901 Fault 12873
		function isEMail(psURL) {			
			var pblnIsEMail, psSearchString;
			psSearchString = 'mailto';
			pblnIsEMail = psURL.search(psSearchString);			
			return (pblnIsEMail);
		}

		function refreshSession() {
			// Submit the refresh.asp to keep the session alive
			try {
				var frmRefresh = document.getElementById('frmRefresh');
				OpenHR.submitForm(frmRefresh);
			}
			catch (e) { }
		}

		function relocateURL(psURL, pfNewWindow) {

			if (!dragged) {
				// Submit the refresh.asp to keep the session alive

				refreshSession();

				//NPG20081102 Fault 12873
				if (isEMail(psURL) == 0) {
					window.location.href = psURL;
					return false;
				}
				if (pfNewWindow == 1) {
					window.open(psURL);
				} else {
					try {
						var aParameters = psURL.split('?');
						loadPartialView(psURL, 'home', 'workframe', aParameters[1]);
					}
					catch (e) {
						alert('error in link');
					}
				}
			}
		}


		function goURL(psURL, pfNewWindow, pfExternal) {

			try {
				pfNewWindow = (pfExternal==true?1:0);
				//if (txtHypertextLinksEnabled.value != 0) {
				relocateURL(psURL, pfNewWindow);
				//}
			}
			catch (e) {
			}
		}


		function goScreen(psScreenInfo) {

			//check to see if we're completing a drag event
			if (!dragged) {
				var sDestination;
				menu_disableMenu();				
				loadPartialView("recordEditMain", "home", "workframe", psScreenInfo);
			}
			//reset drag value
			dragged = 0;
			// Submit the refresh.asp to keep the session alive
			//refreshSession();
			//psScreenInfo = escape(psScreenInfo);

			//sDestination = "recordEditMain.asp?";
			//sDestination = sDestination.concat(psScreenInfo);
			//window.frames("linksworkframe").location.replace(sDestination);
		}

		function goUtility(sUtilityType, sUtilityID, sUtilityName, sUtilityBaseTable) {

			if (!dragged) {
				//menu_disableMenu();


				if (sUtilityType == "25") {
					// Workflow
					var frmWorkflow = document.getElementById('frmUtilityPrompt');
					frmWorkflow.utiltype.value = sUtilityType;
					frmWorkflow.utilid.value = sUtilityID;
					frmWorkflow.utilname.value = sUtilityName;
					frmWorkflow.action.value = "run";

					var sUtilId = new String(sUtilityID);
					frmWorkflow.target = sUtilId;
					frmWorkflow.action = "util_run_workflow";				
					
					//submit but leave hidden - no point showing the message.
					OpenHR.submitForm(frmWorkflow, 'workframe', false);
					$('#SSILinksFrame').hide();
					$('#optionframe').show();
					
				} else {
					//Not a workflow!
					$('#SSILinksFrame').fadeOut();
					$('#SSILinksFrame').promise().done(function () {
						var frmPrompt = OpenHR.getForm("utilities", "frmUtilityPrompt");
						frmPrompt.utiltype.value = sUtilityType;
						frmPrompt.utilid.value = sUtilityID;
						frmPrompt.utilname.value = sUtilityName;
						OpenHR.submitForm(frmPrompt, "workframe", false);
						$('#workframe').fadeIn();
					});
				}
			}
		}

		function goDropLink(sLinkInfo) {
			
			if (sLinkInfo == undefined) {
				sLinkInfo = $('.DropdownlistSelect').val();				
			}


			var sLinkType = sLinkInfo.substr(0, 1);
			sLinkInfo = sLinkInfo.substr(2);
			var sNewWindow;
			var sAppFilePath;
			var sAppParameters;

			if (sLinkType == "0") {
				// URL link
				sNewWindow = sLinkInfo.substr(0, 1);
				sLinkInfo = sLinkInfo.substr(2);

				goURL(sLinkInfo, sNewWindow);
			}
			else {
				if (sLinkType == "2") {
					// Utility link
					goUtility(sLinkInfo);
				}
				// Org Chart
				else if (sLinkType == "6") {
					loadPartialView('OrgChart', 'home', 'workframe')

				}
				else if (sLinkType == "5") {
					// Application link
					sAppFilePath = sLinkInfo.substr(0, sLinkInfo.indexOf('_', 0));
					sAppParameters = sLinkInfo.substr(sLinkInfo.indexOf('_', 0) + 1);
					goApp(sAppFilePath, sAppParameters);
				}
				else {
					if (sLinkType == "4") {
						// Mulitple record find page
						sLinkInfo = "recordEditMain.asp?multifind_0_" + sLinkInfo;
						goURL(sLinkInfo, 0);
					}
					else {
						// HR Pro screen link
						goScreen(sLinkInfo);
					}
				}
			}
		}

		function launchWorkflow(url, name) {

			$('.pageTitle').text(name);
			$('#externalContentFrame').attr('src', url);
			$('.DashContent').fadeOut();
			$('#workflowDisplay').fadeIn();

			//var newWindow = window.open(url);
			//if (window.focus) {
			//	newWindow.focus();
			//}
		}

	</script>


<%
	Dim _PendingWorkflowStepsHTMLTable As New StringBuilder	'Used to construct the (temporary) HTML table that will be transformed into a jQuey grid table
	Dim _StepCount As Integer = 0
	Dim _WorkflowGood As Boolean = True
		
	'Get the pendings workflow steps from the database
	Dim _cmdDefSelRecords = New Command
	_cmdDefSelRecords.CommandText = "spASRSysMobileCheckPendingWorkflowSteps"
	_cmdDefSelRecords.CommandType = CommandTypeEnum.adCmdStoredProc
	_cmdDefSelRecords.ActiveConnection = Session("databaseConnection")

	Dim prmKeyParameter = _cmdDefSelRecords.CreateParameter("screenID", 200, 1, 8000)	
	_cmdDefSelRecords.Parameters.Append(prmKeyParameter)
	prmKeyParameter.Value = Session("username")

	Err.Clear()
	Dim _rstDefSelRecords = _cmdDefSelRecords.Execute
		
	If (Err.Number <> 0) Then		
	' Workflow not licensed or configured. Go to default page.
	_WorkflowGood = False
	Else
	With _PendingWorkflowStepsHTMLTable
			.Append("<table id=""PendingStepsTable_Dash"">")
		.Append("<tr>")
		.Append("<th id=""DescriptionHeader"">Description</th>")
		.Append("<th id=""URLHeader"">URL</th>")
		.Append("<th id=""NameHeader"">URL</th>")
		.Append("</tr>")
	End With
	'Loop over the records
	Do Until _rstDefSelRecords.eof
		_StepCount += 1
		With _PendingWorkflowStepsHTMLTable
			.Append("<tr>")
			.Append("<td>" & _rstDefSelRecords.Fields("description").Value & "</td>")
			.Append("<td>" & _rstDefSelRecords.Fields("url").Value & "</td>")
			.Append("<td>" & _rstDefSelRecords.Fields("name").Value & "</td>")
			.Append("</tr>")
		End With
		_rstDefSelRecords.movenext()
	Loop
						
	_PendingWorkflowStepsHTMLTable.Append("</table>")
						
	_rstDefSelRecords.close()
	_rstDefSelRecords = Nothing
	End If
				
	' Release the ADO command object.
	_cmdDefSelRecords = Nothing

%>

	<div id="" class="DashContent" style="display: block;">
		<div class="tileContent">
		<%Dim fFirstSeparator = True
			Const iMaxRows As Integer = 4
			Dim iRowNum = 1
			Dim iColNum = 1
			Dim iSeparatorNum = 0
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
					<div class="wrapupcontainer"><div class="wrapuptext"><p class="pendingworkflowlinkseparator">To-do list (Pending workflows)</p></div></div>					
					<div class="gridster pendingworkflowlinkcontent" id="gridster_PendingWorkflow" >
						<ul id="pendingworkflowstepstiles">
						</ul>
					</div>					
				</li>
			</ul>	
			</div>

			<%fFirstSeparator = True%>
			<div class="hypertextlinks">
				<%Dim tileCount = 1
					For Each navlink In Model.NavigationLinks
						Dim sTileColourClass = "Colour" & CStr(CInt(Math.Ceiling(Rnd() * 7)))
						If navlink.LinkType = NavigationLinkType.HyperLink Then
							If (navlink.Element_Type = 1 Or navlink.LinkOrder = 0) And navlink.UtilityType = -1 Then		' separator
								iRowNum = 1
								iColNum = 1
								If fFirstSeparator Then
									fFirstSeparator = False
									Else%>
											</ul>
											</div>
												</li> </ul>
			<%End If
				iSeparatorNum += 1
				
				If navlink.Text.Length > 0 Then
					sText = Html.Encode(navlink.Text)
					sText = sText.Replace("--", "")
					sText = sText.Replace("'", """")
				Else
					sText = ""
				End If%>
			
			<ul class="hypertextlinkseparatorframe" id="hypertextlinkseparatorframe_<%=iSeparatorNum %>">
				<li class="hypertextlink-displaytype">
					<div class="wrapupcontainer hypertextlinktextseparator">
						<div class="wrapuptext hypertextlinktextseparator">
							<p class="hypertextlinkseparator hypertextlinkseparator-font hypertextlinkseparator-colour hypertextlinkseparator-size hypertextlinkseparator-bold hypertextlinkseparator-italics"><%=sText%></p>
						</div>
					</div>					
					<div class="gridster hypertextlinkcontent" id="gridster_Hypertextlink_<%=tileCount%>">
						<ul>
							<%Else%>
							<%If iRowNum > iMaxRows Then%>
							<% iColNum += 1%>
							<%iRowNum = 1%>
							<script type="text/javascript">
								$("#hypertextlinkseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
								$("#hypertextlinkseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
							</script>
							<%End If%>
							<%
								classIcon = ""
								sNewWindow = ""
								
								Select Case navlink.Element_Type%>
							<%Case ElementType.ButtonLink
									sURL = NullSafeString(navlink.URL).Replace("'", "\'")
									sURL = sURL.Replace("&", "&amp;")
									sURL = sURL.Replace("""", "&quot;")
									sURL = sURL.Replace(">", "&gt;")
									sURL = sURL.Replace("<", "&lt;")
									
									sAppFilePath = navlink.AppFilePath.Replace("\", "\\")
									sAppParameters = navlink.AppParameters.Replace("\", "\\")
								
									classIcon = "icon-external-link"
									If navlink.AppFilePath.Length > 0 Then
										sOnclick = "goApp('" & sAppFilePath & "', '" & sAppParameters & "')"
										' sCheckKeyPressed = "CheckKeyPressed('APP', '" & sDestination & "',0,'')"
									ElseIf navlink.URL.Length > 0 Then
										If navlink.NewWindow = True Then
											sNewWindow = "1"
										Else
											sNewWindow = "0"
										End If
			
										sOnclick = "goURL('" & sURL & "', " & sNewWindow & ", true)"
										' sCheckKeyPressed = "CheckKeyPressed('HYPERLINK', '" & sURL & "', " & sNewWindow & ",'')"
									Else
										Dim sUtilityType = Convert.ToString(navlink.UtilityType)
										Dim sUtilityID = Convert.ToString(navlink.UtilityID)
										Dim sUtilityDef = sUtilityType & "_" & sUtilityID
										Dim sUtilityBaseTable = CStr(navlink.BaseTable)

										sOnclick = "goUtility(" & sUtilityType & ", " & sUtilityID & ", '" & navlink.Text & "', " & sUtilityBaseTable & ")"

									End If
									
								Case ElementType.OrgChart
									sOnClick = "loadPartialView('OrgChart', 'home', 'workframe')"
									
									
							End Select%>
							<li class="hypertextlinktext hypertextlinktext-highlightcolour <%=sTileColourClass%> flipTile" data-col="<%=iColNum %>" data-row="<%=iRowNum %>"
								data-sizex="1" data-sizey="1" onclick="<%=sOnclick%>">
								<a class="hypertextlinktext-font hypertextlinktext-colour hypertextlinktext-size hypertextlinktext-bold hypertextlinktext-italics" href="#" title="<%: navlink.Text%>"><%: navlink.Text %></a>
								<p class="hypertextlinktileIcon"><i class="<%=classIcon %>"></i></p>
							</li>
							<%iRowNum += 1%>
							<%End If%>
							<%End If%>
							<%tileCount += 1%>
							<%Next
								
								Dim objNavigation = New HR.Intranet.Server.clsNavigationLinks
								objNavigation.Connection = Session("databaseConnection")
								
								' Get the navigation hypertext links.
								Dim iFindPage As Int16 = 0
								'If sWorkPage = "FIND" Then
								'	iFindPage = 1
								'Else
								'	iFindPage = 0
								'End If
								Dim objNavigationHyperlinkInfo = objNavigation.GetNavigationLinks(0, CBool(iFindPage))
								
								Dim sDestination As String
								
								For iCount = 1 To objNavigationHyperlinkInfo.Count
									sText = Html.Encode(objNavigationHyperlinkInfo(iCount).text1)
		
									If objNavigationHyperlinkInfo(iCount).linkToFind = 0 Then
										sDestination = "linksMain?" & CStr(objNavigationHyperlinkInfo(iCount).tableID) & "!" & CStr(objNavigationHyperlinkInfo(iCount).viewID)
			
										If objNavigationHyperlinkInfo(iCount).singleRecord = 1 Then
											sDestination = sDestination & "_0"
										Else
											sDestination = sDestination & "_" & CStr(Session("TopLevelRecID"))
										End If
									Else
										sDestination = "recordEditMain?multifind_0_" & CStr(objNavigationHyperlinkInfo(iCount).tableID) & "!" & CStr(objNavigationHyperlinkInfo(iCount).viewID)
									End If
							%>
							<%			If fFirstSeparator Then		' add a separator
									iRowNum = 1
									iColNum = 1
									If fFirstSeparator Then
										fFirstSeparator = False
									Else%>
						</ul>
					</div>
				</li>
			</ul>
			<%End If
				iSeparatorNum += 1
				
				'If sText.Length > 0 Then
				'	sText = Html.Encode(sText)
				'	sText = sText.Replace("--", "")
				'	sText = sText.Replace("'", """")
				'Else
				'sText = ""
				'End If
				
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
							<%end if%>
							<%If iRowNum > iMaxRows Then%>
							<%	iColNum += 1%>
							<%iRowNum = 1%>
							<script type="text/javascript">
								$("#hypertextlinkseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
								$("#hypertextlinkseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
							</script>
							<%End If%>
							<li class="hypertextlinktext Colour4" data-col="<%=iColNum %>" data-row="<%=iRowNum %>"
								data-sizex="1" data-sizey="1" onclick="goURL('<%=sDestination%>', 0, false)">
								<a class="hypertextlinktext hypertextlinktext-font hypertextlinktext-colour hypertextlinktext-size hypertextlinktext-bold hypertextlinktext-italics" href="#"><%=sText%></a>
								<p class="hypertextlinktileIcon"><i class="icon-external-link-sign"></i></p>
							</li>
							<%iRowNum += 1%>
							
							<%tileCount += 1%>
							<%Next%>


							<%If Not fFirstSeparator Then		' close off the hypertext group%>
						</ul>
					</div>
				</li>
			</ul>
			<%End If%>
		</div>

		<%fFirstSeparator = True%>
		<div class="linkspagebutton">
			<div class="ButtonLinkColumn">
								<%sOnclick = ""
									Dim sLinkKey As String = ""
									sAppFilePath = ""
									sAppParameters = ""
									sNewWindow = "0"									

									For Each navlink In Model.NavigationLinks
				
										Dim sTileColourClass = "Colour" & CStr(CInt(Math.Ceiling(Rnd() * 7)))

										If navlink.LinkType = NavigationLinkType.Button Then

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
			
												sOnClick = "goURL('" & sURL & "', " & sNewWindow & ", true)"

											Else
												If navlink.UtilityID > 0 Then
													Dim sUtilityType = CStr(navlink.UtilityType)
													Dim sUtilityID = CStr(navlink.UtilityID)
													Dim sUtilityBaseTable = CStr(navlink.BaseTable)
												
													sOnClick = "goUtility(" & sUtilityType & ", " & sUtilityID & ", '" & navlink.Text & "', " & sUtilityBaseTable & ")"
												Else
													sLinkKey = "recedit" & "_" & Session("TopLevelRecID").ToString() & "_" & navlink.ID
												
													sOnClick = "goScreen('" & sLinkKey & "')"
												End If
											End If

											If navlink.Element_Type = 1 Then		' separator
												iRowNum = 1
												iColNum = 1
												Dim sSeparatorColor = ""
												If navlink.SeparatorColour <> "" And navlink.SeparatorColour <> "#FFFFFF" Then sSeparatorColor = "background-color: " & navlink.SeparatorColour & "!important;"
												If fFirstSeparator Then
													fFirstSeparator = False
																 Else%>
				</ul>
			</div>
			</li> </ul>
			<%End If
				If navlink.SeparatorOrientation = 1 Then	' Vertical break/new column %>
		</div>
		<div class="ButtonLinkColumn">
			<%End If%>
			<%iSeparatorNum += 1%>
			<ul class="linkspagebuttonseparatorframe" id="linkspagebuttonseparatorframe_<%=iSeparatorNum %>">
				<li class="linkspagebutton-displaytype">					
					<div class="wrapupcontainer linkspagebuttonseparator-bordercolour" style="<%=sSeparatorColor%>">
						<div class="wrapuptext">							
							<p class="linkspagebuttonseparator linkspagebuttonseparator-font linkspagebuttonseparator-colour linkspagebuttonseparator-size linkspagebuttonseparator-bold linkspagebuttonseparator-italics"><%: navlink.Text %></p>
						</div>
					</div>
					<div class="gridster buttonlinkcontent" id="gridster_buttonlink_<%=tileCount%>">
						<ul>
							<%Else
									If iRowNum > iMaxRows Then	 ' start a new column if required (affects tiles only)
										iColNum += 1
										iRowNum = 1%>
							<script type="text/javascript">
								$("#linkspagebuttonseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
								$("#linkspagebuttonseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
							</script>
							<%End If
														
								Select Case navlink.Element_Type

									Case ElementType.ButtonLink
										Dim sIconClass As String = "icon-file"
									
										If navlink.UtilityType = -1 Then	' screen view
											sIconClass = "icon-table"
										ElseIf navlink.UtilityType = 25 Then
											sIconClass = "icon-magic"
												 End If%>
							
									<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext <%=sTileColourClass%>" onclick="<%=sOnclick%>">
										<a class="linkspagebutton-displaytype linkspagebuttontext-alignment linkspagebutton-colourtheme" href="#"><span class="linkspageprompttext-font linkspageprompttext-colour linkspageprompttext-size linkspageprompttext-bold linkspageprompttext-italics"><%: navlink.Prompt.Replace("...", "") & " "%></span>
											<span class="linkspagebuttontext-font linkspagebuttontext-colour linkspagebuttontext-size linkspagebuttontext-bold linkspagebuttontext-italics"><%: navlink.Text %></span><img src="<%: Url.Content("~/Content/images/extlink2.png") %>" alt=""/></a>
										<p class="linkspagebuttontileIcon"><i class="<%=sIconClass%>" ></i></p>
									</li>																																																																								
								<%iRowNum += 1

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
										sOnClick = "goUtility(" & navlink.UtilityType & ", " & navlink.UtilityID & ", '" & navlink.Text & "', " & navlink.BaseTable & ")"
									Else
										sOnClick = ""
									End If
									
									%>
							
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext <%=sTileColourClass%> displayonly">
									<a href="#"><%: navlink.Text %>
										<%If navlink.UtilityID > 0 And navlink.DrillDownHidden = False Then%>
										<img src="<%:Url.Content("~/Content/images/Utilities.gif")%>" style="float: right; cursor: pointer; width: 16; height: 16; vertical-align: bottom;" alt="Drilldown..." title="Drill down to data..."
											onclick="<%=sOnClick %>" />
										<%End If%>
										<img src="<%:Url.Content("~/Content/images/Chart_Popout.png")%>" style="float: right; cursor: pointer;width:16px;height:16px;vertical-align:bottom;" alt="Popout chart..." title="View this chart in a new window" 
											onclick="popoutchart('<%=fMultiAxis%>', '<%=navlink.Chart_ShowLegend%>', '<%=navlink.Chart_ShowGrid%>', '<%=navlink.Chart_ShowValues%>', '<%=navlink.Chart_StackSeries%>', '<%=navlink.Chart_ShowPercentages%>', '<%=iChart_Type%>', '<%=iChart_TableID%>', '<%=iChart_ColumnID%>', '<%=iChart_FilterID%>', '<%=iChart_AggregateType%>', '<%=iChart_ElementType%>', '<%=iChart_TableID_2%>', '<%=iChart_ColumnID_2%>', '<%=iChart_TableID_3%>', '<%=iChart_ColumnID_3%>', '<%=iChart_SortOrderID%>', '<%=iChart_SortDirection%>', '<%=iChart_ColourID%>')"/>
									</a>
									<p class="linkspagebuttontileIcon">
										<i class="icon-bar-chart"></i>
									</p>
									<%If navlink.InitialDisplayMode = 0 Then%>
									<div class="widgetplaceholder chart">
										<%If fMultiAxis Then%>
										<div><img onerror="$(this).parent().parent().hide();" src="<%:Url.Action("GetMultiAxisChart", "Home", New With {.Height = 296, .Width = 412, .ShowLegend = navlink.Chart_ShowLegend, .DottedGrid = navlink.Chart_ShowGrid, .ShowValues = navlink.Chart_ShowValues, .Stack = navlink.Chart_StackSeries, .ShowPercent = navlink.Chart_ShowPercentages, .ChartType = iChart_Type, .TableID = iChart_TableID, .ColumnID = iChart_ColumnID, .FilterID = iChart_FilterID, .AggregateType = iChart_AggregateType, .ElementType = iChart_ElementType, .TableID_2 = iChart_TableID_2, .ColumnID_2 = iChart_ColumnID_2, .TableID_3 = iChart_TableID_3, .ColumnID_3 = iChart_ColumnID_3, .SortOrderID = iChart_SortOrderID, .SortDirection = iChart_SortDirection, .ColourID = iChart_ColourID})%>" alt="Chart" /></div>
										<%Else%>
										<div><img onerror="$(this).parent().parent().hide();" src="<%:Url.Action("GetChart", "Home", New With {.Height = 296, .Width = 412, .ShowLegend = navlink.Chart_ShowLegend, .DottedGrid = navlink.Chart_ShowGrid, .ShowValues = navlink.Chart_ShowValues, .Stack = navlink.Chart_StackSeries, .ShowPercent = navlink.Chart_ShowPercentages, .ChartType = iChart_Type, .TableID = iChart_TableID, .ColumnID = iChart_ColumnID, .FilterID = iChart_FilterID, .AggregateType = iChart_AggregateType, .ElementType = iChart_ElementType, .SortOrderID = iChart_SortOrderID, .SortDirection = iChart_SortDirection, .ColourID = iChart_ColourID})%>" alt="Chart" /></div>
										<%End If%>
										<a href="#"></a>
									</div>
									<%Else%>
									<div class="widgetplaceholder datagrid">
										<table cellspacing="0" cellpadding="5" rules="all" frame="box" style="width:100%;vertical-align:top;border:3px solid lightgray">
											<tr>
												<th onclick="fsort();" style="font-size: 10pt; font-weight: normal; text-align: left">
													<%=Left(NullSafeString(navlink.Chart_ColumnName), 50)%>
												</th>
												<%If fMultiAxis Then%>
												<th onclick="fsort();" style="font-size: 10pt; font-weight: normal; text-align: left">
													<%=Trim(Left(NullSafeString(navlink.Chart_ColumnName_2), 50))%>
												</th>
												<th onclick="fsort();" style="font-size: 10pt; font-weight: normal; text-align: right">
													<%Else%>
												<th onclick="fsort();" style="font-size: 10pt; font-weight: normal; text-align: right">
													<%End If%>
													<%If navlink.Chart_AggregateType = 0 Then%>
														Count
													<%ElseIf navlink.Chart_AggregateType = 1 Then%>
														Total
													<%ElseIf navlink.Chart_AggregateType = 2 Then%>
														Average
													<%ElseIf navlink.Chart_AggregateType = 3 Then%>
														Minimm
													<%ElseIf navlink.Chart_AggregateType = 4 Then%>
														Maximum
													<%End If%>
												</th>
											</tr>
											<%
												Dim objChart As Object
												Dim sErrorDescription As String = ""
												' Dim fFormatting_Use1000Separator As Boolean = (navlink.Formatting_Use1000Separator = 1)
																								
												If fMultiAxis = True Then
													objChart = New HR.Intranet.Server.clsMultiAxisChart
												Else
													objChart = New HR.Intranet.Server.clsChart
												End If

												' Pass required info to the DLL
												objChart.Username = CType(Session("username"), String)
												objChart.Connection = CType(Session("databaseConnection"), Connection)
				
												Dim mrstChartData As Recordset
												Err.Clear()
			
												If fMultiAxis = True Then
													mrstChartData = objChart.GetChartData(iChart_TableID, iChart_ColumnID, iChart_FilterID, iChart_AggregateType, iChart_ElementType, iChart_TableID_2, iChart_ColumnID_2, iChart_TableID_3, iChart_ColumnID_3, iChart_SortOrderID, iChart_SortDirection, iChart_ColourID)
												Else
													mrstChartData = objChart.GetChartData(iChart_TableID, iChart_ColumnID, iChart_FilterID, iChart_AggregateType, iChart_ElementType, iChart_SortOrderID, iChart_SortDirection, iChart_ColourID)
												End If

												If (Err.Number <> 0) Then
													sErrorDescription = "The Chart field values could not be retrieved." & vbCrLf & FormatError(Err.Description)
												End If
			
												If Not mrstChartData Is Nothing Then
													If mrstChartData.RecordCount > 500 Then mrstChartData = Nothing ' limit to 500 rows as get row buffer limit exceeded error.
												End If
										
												If Not (mrstChartData.EOF And mrstChartData.BOF) Then
													mrstChartData.MoveFirst()
								
													Do While Not mrstChartData.EOF%>
											<tr>
												<td class="bordered" style="width: 150px; text-align: left; white-space: nowrap">
													<%If fMultiAxis Then%>
													<%=Trim(Left(NullSafeString(mrstChartData.Fields(1).Value), 50))%>
													<%Else%>
													<%=Trim(Left(NullSafeString(mrstChartData.Fields(0).Value), 50))%>
													<%End If%>
												</td>
												<%If fMultiAxis Then%>
												<td class="bordered" style="text-align: left; white-space: nowrap">
													<div style="width: 150px; white-space: nowrap">
														<%=Trim(Left(NullSafeString(mrstChartData.Fields(3).Value), 50))%>
													</div>
												</td>
												<%End If%>
												<td class="bordered" style="text-align: right; vertical-align: top; padding-bottom: 0; white-space: nowrap; overflow: hidden">
													<%If fMultiAxis Then%>
													<%If navlink.UseFormatting = True And (TryCast(mrstChartData.Fields(4).Value, String) <> "No Access" And TryCast(mrstChartData.Fields(4).Value, String) <> "No Data") Then%>
													<%=FormatNumber(CDbl(Trim(Left(NullSafeString(mrstChartData.Fields(4).Value), 50))), navlink.Formatting_DecimalPlaces, , , TriState.UseDefault)%>
													<%Else%>
													<%=Trim(Left(NullSafeString(mrstChartData.Fields(4).Value), 50))%>
													<%End If%>
													<%Else%>
													<%If navlink.UseFormatting = True And (TryCast(mrstChartData.Fields(1).Value, String) <> "No Access" And TryCast(mrstChartData.Fields(1).Value, String) <> "No Data") Then%>
													<%=FormatNumber(CDbl(Trim(Left(NullSafeString(mrstChartData.Fields(1).Value), 50))), navlink.Formatting_DecimalPlaces, , , TriState.UseDefault)%>
													<%Else%>
													<%=Trim(Left(NullSafeString(mrstChartData.Fields(1).Value), 50))%>
													<%End If%>
													<%End If%>
												</td>
											</tr>
											<%    
												mrstChartData.MoveNext()
											Loop
										Else
											%>
											<tr>
												<td class="bordered" style="text-align: left" nowrap="nowrap">No Data
												</td>
												<td class="bordered" style="text-align: right" nowrap="nowrap"></td>
											</tr>

											<%End If%>
										
											</table>
									</div>
									<%End If%>
								</li>
								<%iRowNum += 1%>

							<%Case ElementType.PendingWorkflows%>
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="2" data-sizey="1"	class="linkspagebuttontext <%=sTileColourClass%> displayonly pwfslink" onclick="relocateURL('WorkflowPendingSteps', 0)">
									<div class="pwfTile <%=sTileColourClass%>">
									<p class="linkspagebuttontileIcon">
										<i class="icon-inbox"></i>
										<div class="workflowCount"></div>
									</p>
									<div class="widgetplaceholder generaltheme">
										<div><i class="icon-inbox"></i></div>
										<a class="linkspageprompttext-font linkspageprompttext-colour linkspageprompttext-size linkspageprompttext-bold linkspageprompttext-italics" href="#">Pending Workflows</a>
									</div>
									</div>
									<div class="pwfList <%=sTileColourClass%>" style="display: none;">
										<p><span>Pending steps:</span></p>
										<table></table>											
									</div>
								</li>
								<%iRowNum += 1%>
								<%fWFDisplayPendingSteps = False%>


							<%Case ElementType.DatabaseValue
									
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

									' Pass required info to the DLL
									objChart.Username = CType(Session("username"), String)
									objChart.Connection = CType(Session("databaseConnection"), Connection)
				
									Err.Clear()
									Dim mrstDbValueData = objChart.GetChartData(navlink.Chart_TableID, navlink.Chart_ColumnID, navlink.Chart_FilterID, _
																															navlink.Chart_AggregateType, navlink.Element_Type, navlink.Chart_SortOrderID, _
																															navlink.Chart_SortDirection, navlink.Chart_ColourID)

									If Err.Number <> 0 Then
										sErrorDescription = "The Database Values could not be retrieved." & vbCrLf & FormatError(Err.Description)
									End If
									
									If Len(sErrorDescription) = 0 Then
										If Not (mrstDbValueData.EOF And mrstDbValueData.BOF) Then
											Do While Not mrstDbValueData.EOF
												sText = CType(mrstDbValueData.Fields(0).Value, String)
												mrstDbValueData.MoveNext()
											Loop
											Dim fDoFormatting As Boolean
											
											If fUseConditionalFormatting = True Then
												For jnCount = 0 To 2
													fDoFormatting = False
													If sCFValue(jnCount) <> vbNullString Then
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
										mrstDbValueData.Close()
									End If

									
									If sText <> "No Data" And sCFVisible = True Then
									
										If fFormattingApplies Then
							%>
							<li id="li_<%: navlink.id %>" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
								data-sizey="1" class="linkspagebuttontext <%=sTileColourClass%> displayonly linkspagebuttontext-font linkspagebuttontext-colour linkspagebuttontext-size linkspagebuttontext-bold linkspagebuttontext-italics">
								<div class="DBValueScroller" id="marqueeDBV<%: navlink.id %>">
									<p class="DBValue" style="color: <%=sCFForeColor%>; <%=sCFFontBold%>; <%=sCFFontItalic%>" id="DBV<%: navlink.id %>">
											<%If fUseFormatting = True Then%>
										 <%=sFormatting_Prefix%><%=FormatNumber(cdbl(sText), iFormatting_DecimalPlaces,,,fFormatting_Use1000Separator)%><%=sFormatting_Suffix%>
										<%Else%>
										<%: sText %>
										<%end if %>
									</p>
								</div>
								<a href="#">
									<p class="DBValueCaption" style="color: <%=sCFForeColor%>; <%=sCFFontBold%>; <%=sCFFontItalic%>">										
										<%: navlink.Text %>
									</p>
								</a>
							</li>

							<%Else%>
							<li id="li_<%: navlink.id %>" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
								data-sizey="1" class="linkspagebuttontext <%=sTileColourClass%> displayonly linkspagebuttontext-font linkspagebuttontext-colour linkspagebuttontext-size linkspagebuttontext-bold linkspagebuttontext-italics">
								<div class="DBValueScroller" id="marqueeDBV<%: navlink.id %>">
									<p class="DBValue" id="DBV<%: navlink.id %>">
											<%If fUseFormatting = True Then%>
										 <%=sFormatting_Prefix%><%=FormatNumber(cdbl(sText), iFormatting_DecimalPlaces,,,fFormatting_Use1000Separator)%><%=sFormatting_Suffix%>
										<%Else%>
										<%: sText %>
										<%end if %>
									</p>
								</div>
								<a href="#">
									<p class="DBValueCaption">
										<%: navlink.Text %>
									</p>
								</a>
							</li>
							<%End If
								End If%>

								<script type="text/javascript">									//loadjscssfile('$.getScript("../scripts/widgetscripts/wdg_oHRDBV.js", function () { initialiseWidget(<%: navlink.id %>, "DBV<%: navlink.id %>", "DBV<%: navlink.Text %>", ""); });', 'ajax');</script>
								<%iRowNum += 1%>

							<%Case ElementType.TodaysEvents%>							
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="2" data-sizey="1"	class="linkspagebuttontext <%=sTileColourClass%> displayonly TELink">
									<div class="TETile <%=sTileColourClass%>">
									<p class="linkspagebuttontileIcon">
										<i class="icon-calendar"></i>
										<div class="TECount"></div>
									</p>
									<p>
										<a href="#"><%=FormatDateTime(Now, vbLongDate)%></a>
									</p>
									<div class="widgetplaceholder generaltheme">
										<div><i class="icon-calendar"></i></div>
										<a href="#">Today's Events</a>
									</div>
									</div>
									<div class="TEList <%=sTileColourClass%>">
										<p><span>Today's Events:</span></p>
										<table style="width: 100%;">
											<%											
												' ----------------------- DIARY LINKS -----------------------------
												Dim sErrorDescription As String = ""
												Dim iRecNum As Integer
												Dim iNumberOfEvents As Integer = 0
												
												' Create the reference to the DLL
												Dim objDiaryEvents As HR.Intranet.Server.clsDiary = New HR.Intranet.Server.clsDiary

												' Pass required info to the DLL
												objDiaryEvents.Username = CType(Session("username"), String)
												objDiaryEvents.Connection = CType(Session("databaseConnection"), Connection)
				
												Err.Clear()
												Dim mrstEventData As Recordset = objDiaryEvents.GetDiaryData(False, Now.Date, Now.Date)
													
												
												If (Err.Number() <> 0) Then
													sErrorDescription = "The Event Data could not be retrieved." & vbCrLf & FormatError(Err.Description)
												End If
												iRecNum = 0
												
												If sErrorDescription.Length = 0 Then
													If Not (mrstEventData.EOF And mrstEventData.BOF) Then
											%>
											<tr>
												<td colspan="2" style="font-weight: bold; font-size: xx-small; border-bottom: 1px solid gray">Diary Links</td>
											</tr>
											<%       
												Do While Not mrstEventData.EOF
											%>
											<tr>
												<td colspan="2" style="font-weight: normal; font-size: xx-small"><%=mrstEventData.fields(3).value %></td>
											</tr>
											<%                
												mrstEventData.MoveNext()
												iRecNum = iRecNum + 1
											Loop
										End If

										mrstEventData.close()
									End If
											
									iNumberOfEvents += iRecNum
											
									' ----------------------- OUTLOOK LINKS -----------------------------
									' Create the reference to the DLL
									Dim objOutlookEvents As HR.Intranet.Server.clsOutlookLinks = New HR.Intranet.Server.clsOutlookLinks									

									' Pass required info to the DLL
									objOutlookEvents.Username = Session("username")
									objOutlookEvents.Connection = Session("databaseConnection")
				
									Err.Clear()
									mrstEventData = objOutlookEvents.GetOutlookLinks(False, Now.Date, Now.Date)

									If (Err.Number <> 0) Then
										sErrorDescription = "The Outlook Links Data could not be retrieved." & vbCrLf & FormatError(Err.Description)
									End If
									iRecNum = 0
											
									If Len(sErrorDescription) = 0 Then
										If Not (mrstEventData.EOF And mrstEventData.BOF) Then
											%>
											<tr>
												<td colspan="2" style="font-weight: bold; font-size: xx-small; border-bottom: 1px solid gray">Outlook Calendar Links</td>
											</tr>
											<%
												
												Do While Not mrstEventData.EOF
											%>
											<tr>
												<td colspan="2" style="font-weight: normal; font-size: xx-small"><%=trim(mrstEventData.fields(2).value)%></td>
											</tr>
											<%
												mrstEventData.movenext()
												iRecNum = iRecNum + 1
											Loop
										End If

										mrstEventData.close()

									End If
									
									iNumberOfEvents += iRecNum
											

									' ----------------------- TODAY'S ABSENCES -----------------------------
									' Create the reference to the DLL
									Dim objTodaysEvents As HR.Intranet.Server.clsTodaysAbsence = New HR.Intranet.Server.clsTodaysAbsence
						

									' Pass required info to the DLL
									objTodaysEvents.Username = Session("username")
									objTodaysEvents.Connection = Session("databaseConnection")
				
									Err.Clear()
									mrstEventData = objTodaysEvents.GetTodaysAbsences(CleanNumeric(Session("TopLevelRecID")))

									If (Err.Number <> 0) Then
										sErrorDescription = "Todays Absence Data could not be retrieved." & vbCrLf & FormatError(Err.Description)
									End If
									iRecNum = 0
											
									If Len(sErrorDescription) = 0 Then
										If Not (mrstEventData.EOF And mrstEventData.bof) Then
											%>
											<tr>
												<td colspan="2" style="font-weight: bold; font-size: xx-small; border-bottom: 1px solid gray">Today's Absences</td>
											</tr>
											<%             
												
												Do While Not mrstEventData.EOF
											%>
											<tr>
												<td colspan="2" style="font-weight: normal; font-size: xx-small"><%=trim(mrstEventData.fields(0).value) %></td>
											</tr>
											<%                
												mrstEventData.movenext()
												iRecNum = iRecNum + 1
											Loop
										End If
										iNumberOfEvents += iRecNum
											
										mrstEventData.close()
										
											
									End If%>

										</table>											
									</div>

									<div class="linkspagebuttontileIcon"><span><p><%=iNumberOfEvents%></p><p style="font-size: small;">Events</p></span></div>
								</li>
								<%iRowNum += 1
							
								Case ElementType.OrgChart
									sOnclick = "loadPartialView('OrgChart', 'home', 'workframe')"%>			
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext <%=sTileColourClass%>" onclick="<%=sOnclick%>">
										<a href="#"><%: navlink.Text %><img src="<%: Url.Content("~/Content/images/extlink2.png") %>" alt=""/></a>
										<p class="linkspagebuttontileIcon"><i class="icon-sitemap" ></i></p>
									</li>				
							

							<%Case Else%>
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"
									class="linkspagebuttontext <%=sTileColourClass%> displayonly"><a href="#">
										<%: navlink.Text %></a></li>
								<%iRowNum += 1

							End Select

							End If
							End If
							tileCount += 1
						Next
						If Not fFirstSeparator Then%>
							</ul>
				</div>
						</li>
					</ul>
				<%End If%>
			</div>
		</div>

		<%If Model.NumberOfLinks > 0 Then%>
	<div class="dropdownlinks">
		<ul class="dropdownlinkseparatorframe" id="dropdownlinkseparatorframe_<%=iSeparatorNum %>">
		<li class="dropdownlink-displaytype">
			<p class="dropdownlinkseparator">Dropdown links:</p>
			<div class="gridster" id="gridster_DropdownLinks">
			<ul class="DropDownListMenu">
				<%iRowNum = 1
					iColNum = 1

					For Each navlink In Model.NavigationLinks.FindAll(Function(n) n.LinkType = NavigationLinkType.DropDown)
						
						Dim sTileColourClass = "Colour" & CStr(CInt(Math.Ceiling(Rnd() * 7)))
						Dim sValue As String, sUtilityType As String, sUtilityID As String, sUtilityBaseTable As String, sUtilityDef As String
						
						If Len(navlink.AppFilePath) > 0 Then
							sAppFilePath = Replace(navlink.AppFilePath, "\", "\\")
							sAppParameters = Replace(navlink.AppParameters, "\", "\\")
			
							sValue = "5_" & sAppFilePath & "_" & sAppParameters
							sOnclick = "goDropLink('" + sValue + "')"

						ElseIf navlink.Element_Type = ElementType.OrgChart Then
							sValue = "6_OrgChart"
							sOnclick = "loadPartialView('OrgChart', 'home', 'workframe')"
							
							
						ElseIf Len(navlink.URL) > 0 Then
							sURL = Html.Encode(navlink.URL)
							sURL = Replace(sURL, "'", "\'")

							If navlink.NewWindow = True Then
								sNewWindow = "1"
							Else
								sNewWindow = "0"
							End If
		 
							sValue = "0_" & sNewWindow & "_" & sURL
							sOnclick = "goDropLink('" + sValue + "')"
							
						Else
							If navlink.UtilityID > 0 Then
								sUtilityType = CStr(navlink.UtilityType)
								sUtilityID = CStr(navlink.UtilityID)
								sUtilityBaseTable = CStr(navlink.BaseTable)
								sUtilityDef = sUtilityType & "_" & sUtilityID & "_" & sUtilityBaseTable
				
								sValue = "2_" & sUtilityDef
				
							Else
								sLinkKey = "recedit" & _
									"_" & Session("TopLevelRecID").ToString() & _
									"_" & navlink.ID
					
								sValue = "1_" & sLinkKey
				
							End If
							
							sOnclick = "goDropLink('" + sValue + "')"
							
						End If

						If iRowNum > iMaxRows Then	 ' start a new column if required (affects tiles only)
							iColNum += 1
							iRowNum = 1%>
				<script type="text/javascript">
					$("#dropdownlinksseparatorframe<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
					$("#dropdownlinksseparatorframe<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
				</script>
				<%End If%>
				<li class="dropdownlinktext <%=sTileColourClass%>" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
					data-sizey="1">
					<p class="dropdownlinktileIcon">
						<i class="icon-external-link"></i>
					</p>
					<p>
						<a href="#" data-DDLValue="<%=sValue%>" onclick="<%=sOnclick%>">
							<%: navlink.Text %></a>
					</p>
				</li>
				<%iRowNum += 1

				Next%>
			</ul><a class="DropLinkGoText" style="text-decoration: none;margin-left: 10px;" href="#" onclick="goDropLink()">Go...</a>
			</div>
			</li>
			
			</ul>
	</div>

		<%End If%>

		</div>
	</div>


<div id="documentDisplay">
	<div id="divResize">
		<img id="splitToggle" src="" alt="Show Document Display"
			onclick="setDocumentDisplayVisible();" />
	</div>
		<div id="documentDisplayContent" rowspan="4" width="340px" valign="top" nowrap="nowrap">
		<%Html.RenderPartial("~/Views/Home/documentDisplay.ascx")%>     
		</div>
</div>

<div id="pollframeset">
		<div id="poll" data-framesource="poll.asp" style="display: none"></div>
		<div id="pollmessageframe" data-framesource="pollmessage.asp" style="display: none"><%Html.RenderPartial("~/views/home/pollmessage.ascx")%></div>
	</div>    
	
<div id="pwfs"><%Response.Write(_PendingWorkflowStepsHTMLTable.ToString())%></div>

<FORM action="" method="POST" id="frmMenuInfo" name="frmMenuInfo">
<%
	Response.Write("<INPUT type=""hidden"" id=txtDefaultStartPage name=txtDefaultStartPage value=""" & Replace(Session("DefaultStartPage"), """", "&quot;") & """>")
	Response.Write("<INPUT type=""hidden"" id=txtDatabase name=txtDatabase value=""" & Replace(Session("Database"), """", "&quot;") & """>")
%>
	<INPUT type="hidden" id=txtIEVersion name=txtIEVersion value=<%=session("IEVersion")%>>
	<INPUT type="hidden" id=txtUserType name=txtUserType value=<%=session("userType")%>>

	<INPUT type="hidden" id=txtPersonnel_EmpTableID name=txtPersonnel_EmpTableID value=<%=session("Personnel_EmpTableID")%>>

	<INPUT type="hidden" id=txtTB_EmpTableID name=txtTB_EmpTableID value=<%=session("TB_EmpTableID")%>>
	<INPUT type="hidden" id=txtTB_CourseTableID name=txtTB_CourseTableID value=<%=session("TB_CourseTableID")%>>
	<INPUT type="hidden" id=txtTB_CourseCancelDateColumnID name=txtTB_CourseCancelDateColumnID value=<%=session("TB_CourseCancelDateColumnID")%>>
	<INPUT type="hidden" id=txtWaitListOverRideColumnID name=txtWaitListOverRideColumnID value=<%=session("TB_WaitListOverRideColumnID")%>>
	<INPUT type="hidden" id=txtTB_TBTableID name=txtTB_TBTableID value=<%=session("TB_TBTableID")%>>
	<INPUT type="hidden" id=txtTB_TBTableSelect name=txtTB_TBTableSelect value=<%=session("TB_TBTableSelect")%>>
	<INPUT type="hidden" id=txtTB_TBTableInsert name=txtTB_TBTableInsert value=<%=session("TB_TBTableInsert")%>>
	<INPUT type="hidden" id=txtTB_TBTableUpdate name=txtTB_TBTableUpdate value=<%=session("TB_TBTableUpdate")%>>
	<INPUT type="hidden" id=txtTB_TBStatusColumnID name=txtTB_TBStatusColumnID value=<%=session("TB_TBStatusColumnID")%>>
	<INPUT type="hidden" id=txtTB_TBStatusColumnUpdate name=txtTB_TBStatusColumnUpdate value=<%=session("TB_TBStatusColumnUpdate")%>>
	<INPUT type="hidden" id=txtTB_TBCancelDateColumnID name=txtTB_TBCancelDateColumnID value=<%=session("TB_TBCancelDateColumnID")%>>
	<INPUT type="hidden" id=txtTB_TBCancelDateColumnUpdate name=txtTB_TBCancelDateColumnUpdate value=<%=session("TB_TBCancelDateColumnUpdate")%>>
	<INPUT type="hidden" id=txtTB_TBStatusPExists name=txtTB_TBStatusPExists value=<%=session("TB_TBStatusPExists")%>>
	<INPUT type="hidden" id=txtTB_WaitListTableID name=txtTB_WaitListTableID value=<%=session("TB_WaitListTableID")%>>
	<INPUT type="hidden" id=txtTB_WaitListTableInsert name=txtTB_WaitListTableInsert value=<%=session("TB_WaitListTableInsert")%>>
	<INPUT type="hidden" id=txtTB_WaitListTableDelete name=txtTB_WaitListTableDelete value=<%=session("TB_WaitListTableDelete")%>>
	<INPUT type="hidden" id=txtTB_WaitListCourseTitleColumnID name=txtTB_WaitListCourseTitleColumnID value=<%=session("TB_WaitListCourseTitleColumnID")%>>
	<INPUT type="hidden" id=txtTB_WaitListCourseTitleColumnUpdate name=txtTB_WaitListCourseTitleColumnUpdate value=<%=session("TB_WaitListCourseTitleColumnUpdate")%>>
	<INPUT type="hidden" id=txtTB_WaitListCourseTitleColumnSelect name=txtTB_WaitListCourseTitleColumnSelect value=<%=session("TB_WaitListCourseTitleColumnSelect")%>>
	<INPUT type="hidden" id=txtPrimaryStartMode name=txtPrimaryStartMode value=<%=session("PrimaryStartMode")%>>
	<INPUT type="hidden" id=txtHistoryStartMode name=txtHistoryStartMode value=<%=session("HistoryStartMode")%>>
	<INPUT type="hidden" id=txtLookupStartMode name=txtLookupStartMode value=<%=session("LookupStartMode")%>>
	<INPUT type="hidden" id=txtQuickAccessStartMode name=txtQuickAccessStartMode value=<%=session("QuickAccessStartMode")%>>
	<INPUT type="hidden" id=txtDesktopColour name=txtDesktopColour value=<%=session("DesktopColour")%>>

	<INPUT type="hidden" id=txtWFEnabled name=txtWFEnabled value=<%=session("WF_Enabled")%>>
	<INPUT type="hidden" id=txtWFOutOfOfficeEnabled name=txtWFOutOfOfficeEnabled value=<%=session("WF_OutOfOfficeConfigured")%>>

	<INPUT type="hidden" id=txtDoneDatabaseMenu name=txtDoneDatabaseMenu value=0>
	<INPUT type="hidden" id=txtDoneQuickEntryMenu name=txtDoneQuickEntryMenu value=0>
	<INPUT type="hidden" id=txtDoneTableScreensMenu name=txtDoneTableScreensMenu value=0>
	<INPUT type="hidden" id=txtDoneSelfServiceStart name=txtDoneSelfServiceStart value=0>

	<INPUT type="hidden" id=txtMenuSaved name=txtMenuSaved value=0>
</FORM>	
	
<div id="utilities">
	<form name="frmUtilityPrompt" method="post" action="util_run_promptedValues" id="frmUtilityPrompt" style="visibility: hidden; display: none">
		<input type="hidden" id="utiltype" name="utiltype" value="">
		<input type="hidden" id="utilid" name="utilid" value="">
		<input type="hidden" id="utilname" name="utilname" value="">
		<input type="hidden" id="action" name="action" value="run">
	</form>
</div>

<div id="workflowDisplay" class="absolutefull" style="display: none; background-color: transparent; text-align: center;">
	<div class="pageTitleDiv" style="text-align: left;">
		<a href='<%=Url.Action("Main", "Home", New With {.SSIMode = "True"})%>' title='Home'>
			<i class='pageTitleIcon icon-circle-arrow-left'></i>
		</a>
		<span class="pageTitle">Workflow</span>
	</div>

	<iframe id="externalContentFrame" style="width: 700px; height: 400px; margin: 0 auto;"></iframe>
</div>

<script type="text/javascript">
	//Display Pending Workflow Steps if appropriate
	if (('<%=fWFDisplayPendingSteps%>' == 'True') && (Number('<%=_StepCount%>') > 0) && ('<%=Session("ViewDescription")%>' == '')) {
		relocateURL('WorkflowPendingSteps', 0);
	}

	$(".sp-container.sp-hidden").css("display", "none"); //The color picker plugin sometimes leaves visible bits; remove them
</script>
