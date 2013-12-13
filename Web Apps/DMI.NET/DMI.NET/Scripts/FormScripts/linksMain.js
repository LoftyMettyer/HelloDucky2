

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

	//$('.pageTitle').text(name);
	//$('#externalContentFrame').attr('src', url);
	//$('.DashContent').fadeOut();
	//$('#workflowDisplay').fadeIn();

	var newWindow = window.open(url);
	if (window.focus) {
		newWindow.focus();
	}
}
