

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
		window.location = "MainSSI";
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

function relocateURL(psUrl, pfNewWindow) {

	if (!dragged) {

		//NPG20081102 Fault 12873
		if (isEMail(psUrl) == 0) {
			window.location.href = psUrl;
			return false;
		}
		if (pfNewWindow == 1) {
			window.open(psUrl);
		} else {
			try {
				var aParameters = psUrl.split('?');
				loadPartialView(psUrl, 'home', 'workframe', aParameters[1]);
			}
			catch (e) {
				alert('error in link');
			}
		}
	}
}

function goURL(psUrl, pfNewWindow, pfExternal, psTitle) {
	
	if (pfExternal == true) {
		if (isEMail(psUrl) == 0) {
			window.location.href = psUrl;
			return false;
		}
		
		if (pfNewWindow == 1) {
			window.open(psUrl);
			return false;
		}
		
		//external content
		$('.DashContent').hide();
		$('#workflowDisplay').show();
		$('#externalContentFrame').attr('src', psUrl);
		$('#externalContentFrame').css("height", "90%");
		$('#workflowDisplay .pageTitle').text(psTitle);

		if (window.currentLayout == "tiles") { //Tiles need some tweaking of the HTML/CSS
			$('.DashContent').show();
			$('.hypertextlinks').hide();
			$('.linkspagebutton').hide();
			$('.dropdownlinks').hide();
			$('.pendingworkflowlinks').hide();
			$('#workflowDisplay').css("position", "fixed");
			$('#externalContentFrame').css("height", "77%");
		}

		return false;
	}

	try {
		//pfNewWindow = (pfExternal==true?1:0);
		//if (txtHypertextLinksEnabled.value != 0) {
		relocateURL(psUrl, pfNewWindow);
		//}
	}
	catch (e) {
	}
}

function goScreen(psScreenInfo) {

	//check to see if we're completing a drag event
	if (!dragged) {
	

		menu_disableMenu();				
		loadPartialView("recordEditMain", "home", "workframe", psScreenInfo);
	}
	//reset drag value
	dragged = 0;
	// Submit the refresh.asp to keep the session alive
	//psScreenInfo = escape(psScreenInfo);

	//sDestination = "recordEditMain.asp?";
	//sDestination = sDestination.concat(psScreenInfo);
	//window.frames("linksworkframe").location.replace(sDestination);
}

function goUtility(sUtilityType, sUtilityID, sUtilityName, sUtilityBaseTable) {	
	if (!dragged) {

		if ((window.isMobileBrowser == "true") && (sUtilityType == "9")) {
			//No mailmerges for mobiles - see JIRA 3969
			return false;
		}

		if (sUtilityType == "25") {

			// Workflow
			var postData = {
				utiltype: sUtilityType,
				utilid: sUtilityID,
				utilname: sUtilityName,
				__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
			}

			OpenHR.submitForm(null, "workframe", null, postData, "util_run_workflow"
				, function () {
					$('#SSILinksFrame').hide();
					$('#optionframe').show();
				});
					
		} else {
			
			//Not a workflow!
			$('#SSILinksFrame').fadeOut();
			$('#SSILinksFrame').promise().done(function() {

				var postData = {
					UtilType: sUtilityType,
					ID: sUtilityID,
					Name: sUtilityName,
					__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
				};
				OpenHR.submitForm(null, "workframe", null, postData, "util_run_promptedValues");

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
	var arrLinkInfo;

	switch (sLinkType) {
		case "0":
			// URL link
			arrLinkInfo = sLinkInfo.split('_');
			sNewWindow = arrLinkInfo[0];
			sLinkInfo = arrLinkInfo[1];
			var sTitleText = arrLinkInfo[2];

			goURL(sLinkInfo, sNewWindow, true, sTitleText);
			break;
		case "2":

			// Utility link
			arrLinkInfo = sLinkInfo.split('_');

			var sUtilityType = arrLinkInfo[0];
			var sUtilityID = arrLinkInfo[1];
			var sUtilityName = arrLinkInfo[2];

			goUtility(sUtilityType, sUtilityID, sUtilityName);
			break;
		case "4":
			// Mulitple record find page
			sLinkInfo = "recordEditMain?multifind_0_" + sLinkInfo;
			goURL(sLinkInfo, 0, false);
			break;
		case "5":
			// Application link
			sAppFilePath = sLinkInfo.substr(0, sLinkInfo.indexOf('_', 0));
			sAppParameters = sLinkInfo.substr(sLinkInfo.indexOf('_', 0) + 1);
			goApp(sAppFilePath, sAppParameters);
			break;
		case "6":
			//Org Chart
			loadPartialView('OrgChart', 'home', 'workframe');
			break;
		case 7:
			//linksMain link
			sLinkInfo = "linksMain?" + sLinkInfo;
			goURL(sLinkInfo, 0, false);
			break;
		default:
			// HR Pro screen link
			goScreen(sLinkInfo);
			break;
	}

}


function goApp(sAppFilePath, sAppParameters) {

	OpenHR.modalMessage("Application links are not available.");

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
