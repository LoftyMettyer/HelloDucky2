
var formInputPrefix = "FI_";

Sys.Application.add_init(function () {
	// For postback, set up the scripts for begin and end requests...
	try {
		var inst = Sys.WebForms.PageRequestManager.getInstance();
		if (!inst.get_isInAsyncPostBack()) {
			inst.add_beginRequest(goSubmit);
			inst.add_endRequest(showMessage);
		}
	} catch (e) { }
});

jQuery.noConflict();

var jQuerySetup = function () {

	// configure date controls
	jQuery.datepicker.setDefaults({
		changeYear: true,
		changeMonth: true,
		showOtherMonths: true,
		selectOtherMonths: true,
		showOn: window.isMobile ? 'both' : 'button',
		buttonImage: 'Images/calendar16.png',
		buttonText: '',
		buttonImageOnly: true,
		dateFormat: window.localeDateFormatjQuery,
		beforeShow: function (input) {
			if (!window.androidLayerBug) return;
			var top = jQuery(input).offset().top + jQuery(input).height();
			jQuery('[id^=FI], img').filter(function () { return jQuery(this).offset().top > top && jQuery(this).offset().top < top + 100; }).addClass('androidHide');
		},
		onClose: function () {
			if (!window.androidLayerBug) return;
			jQuery('.androidHide').removeClass('androidHide');
		}
	});

	jQuery('input.date.withPicker').datepicker();

	jQuery('img.ui-datepicker-trigger').css('z-index', '3');

	jQuery('input.date.withPicker').change(function () {
		//validate a typed date and format it
		var $this = jQuery(this);
		var value = $this.val();
		try {
			var date = jQuery.datepicker.parseDate(window.localeDateFormatjQuery, value);
			if (date != null) {
				$this.val(jQuery.datepicker.formatDate(window.localeDateFormatjQuery, date));
				jQuery.datepicker.setDefaults({ defaultDate: date });
			}
		} catch (e) {
			$this.val('');
		}
	});

	jQuery('input.date.withPicker').keyup(function (e) {
		//F2 should set todays date
		if (e.which == 113) {
			var date = new Date();
			jQuery(this).val(jQuery.datepicker.formatDate(window.localeDateFormatjQuery, date));
			jQuery.datepicker.setDefaults({ defaultDate: date });
		}
	});

	//configure numeric controls
	jQuery.metadata.setType('attr', 'data-numeric');
	jQuery('input.numeric').autoNumeric({ aSep: '', aDec: window.localeDecimal, wEmpty: 'zero' });
};

jQuery(jQuerySetup);

//fault HRPRO-2270
function resizeIframe(id, newHeight) {
	//Plus one for luck (IE9 actually)
	document.getElementById(id).height = (newHeight + 1) + "px";
}

var overlay;
var wait;

function InitialiseWindow() {

	overlay = jQuery('#divOverlay');
	wait = jQuery('#pleasewaitScreen');

	//Set the current page tab	  
	SetCurrentTab(document.getElementById("hdnDefaultPageNo").value);

	window.iCurrentMessageState = 'none';

	try {
		if ("ActiveXObject" in window) {			
			var iDefHeight, iDefWidth, iResizeByHeight, iResizeByWidth;
			iDefHeight = jQuery('#pnlInputDiv').height();
			iDefWidth = jQuery('#pnlInputDiv').width();

			window.focus();

			if (iDefHeight > 0 && iDefWidth > 0) {
				iResizeByHeight = iDefHeight - window.currentHeight;
				iResizeByWidth = iDefWidth - window.currentWidth;

				window.parent.resizeBy(iResizeByWidth, iResizeByHeight);
				window.parent.moveTo((screen.availWidth - iDefWidth) / 2, (screen.availHeight - iDefHeight) / 3);
			}
		}
		
		try {
			if (window.autoFocusControl.length > 0) {
				setTimeout(function () {
					document.getElementById(window.autoFocusControl).focus();
				}, 0);
			}
		}
		catch (e) { }

		launchForms(window.$get("frmMain").hdnSiblingForms.value, false);
	}
	catch (e) { }

	/*	
	//For GPS Location functionality. Future Development. 
	//Unremark this, and same comments in default.aspx.vb to enable.
	if (navigator.geolocation) {
	var GPSObjects = document.getElementsByClassName("GPSTextBox");
	for (var i = 0; i < GPSObjects.length; i++) {
	navigator.geolocation.getCurrentPosition(function (position) {
	var lat = position.coords.latitude;
	var lng = position.coords.longitude; 
	GPSObjects[i].value = lat + "," + lng;
	});			
	}
	}	
	*/
}

function launchForms(psForms, pfFirstFormRelocate) {

	try {
		var asForms = psForms.split("\t"), sFirstForm = "";

		for (var i = 0; i < asForms.length; i++) {
			if (i == 0) {
				sFirstForm = asForms[i];
			} else {
				// Open other forms in new browsers.
				spawnWindow(asForms[i]);
			}
		}

		if (sFirstForm.length > 0) {
			if (pfFirstFormRelocate == true) {
				// Open first form in current browser.
				window.location = sFirstForm;
			} else {
				// Open first form in new browser.
				spawnWindow(sFirstForm);
			}
		}
	} catch (e) {
	}
}

function spawnWindow(psURL) {
	var newWin;
	try {
		newWin = window.open(psURL);
		try { newWin.window.focus(); } catch (e) { }
	}
	catch (e) {
		try {
			try {
				newWin.close();
			}
			catch (e) {
				alert("For your security please close your browser");
			}
		}
		catch (e) { }

		spawnWindow(psURL);
	}
}

function goSubmit() {

	if ($get("txtPostbackMode").value == "3") {
		try {
			if ($get("txtActiveDDE").value.indexOf("dde") > 0) {
				//keep the lookup open.
				//kicks off InitializeLookup BTW.
				$find($get("txtActiveDDE").value).show();
			}
		}
		catch (e) { }
		return;
	}

	wait.show();
	overlay.show();
}

//TODO replace should be 1 line of jQuery
function getElementsBySearchValue(searchValue) {
	var retVal = new Array();
	var elems = document.getElementsByTagName("input");

	for (var i = 0; i < elems.length; i++) {
		var valueProp = "";

		try {
			var nameProp = elems[i].getAttribute('name');
			if (nameProp.substr(0, 8) == "lookupFI")
				valueProp = elems[i].getAttribute('value');
		}
		catch (e) { }

		if (!(valueProp == null)) {
			if (valueProp.indexOf(searchValue) > 0) {
				retVal.push(elems[i]);
			}
		}
	}

	return retVal;
}

function showErrorMessages(state) {

	switch (state) {
		case 'max':
			jQuery('#errorMessagePanel').show();
			jQuery('#errorMessageMax').hide();
			break;
		case 'min':
			jQuery('#errorMessagePanel').hide();
			jQuery('#errorMessageMax').show();
			break;
		default:
			jQuery('#errorMessagePanel').hide();
			jQuery('#errorMessageMax').hide();
	}
	window.iCurrentMessageState = state;
}

function hasErrors() {
	return jQuery('#hdnCount_Errors').val() > 0 ||
    	    jQuery('#hdnCount_Warnings').val() > 0;
}

function launchFollowOnForms(psForms) {
	launchForms(psForms, true);
}

function overrideWarningsAndSubmit() {

	$get("frmMain").hdnOverrideWarnings.value = "1";

	try {
		document.getElementById($get("frmMain").hdnLastButtonClicked.value).click();
	}
	catch (e) { }
}

function submitForm() {
	var mode = document.getElementById("txtPostbackMode").value;

	return (mode != 0);
}

function setPostbackMode(piValue) {
	// 0 = Default
	// 1 = Submit/SaveForLater button postback (ie. WebForm submission)
	// 2 = Grid header postback
	// 3 = FileUpload button postback
	try {
		document.getElementById("txtPostbackMode").value = piValue;
	}
	catch (e) { }

}

function SR(row, rowIndex) {

	var gridId = row.parentNode.parentNode.id;

	SetScrollTopPos(gridId, document.getElementById(gridId.replace('_Grid', '_gridcontainer')).scrollTop, rowIndex);
	try {
		setPostbackMode(3);
	}
	catch (e) {
	};
	__doPostBack(gridId, 'Select$' + rowIndex);
}

function showFileUpload(pfDisplay, psElementItemID, psAlreadyUploaded) {

	try {
		if (pfDisplay == true) {

			var sAlreadyUploaded = new String(psAlreadyUploaded);
			sAlreadyUploaded = sAlreadyUploaded.substr(0, 1);
			if (sAlreadyUploaded != "1") {
				sAlreadyUploaded = "0";
			}

			$get("ifrmFileUpload").src = "FileUpload.aspx?" + sAlreadyUploaded + psElementItemID;

			overlay.show();
			showErrorMessages(hasErrors() ? 'min' : 'none');

			document.getElementById("divFileUpload").style.display = "block";
		}
		else {
			document.getElementById("divFileUpload").style.display = "none";
			overlay.hide();
		}
	}
	catch (e) { }
}

function fileUploadDone(psElementItemID, piExitMode) {
	// 0 = Cancel
	// 1 = Clear
	// 2 = File Uploaded
	// Hide the file upload dialog, and record how the fileUpload was performed.
	try {
		if ((piExitMode == 1) || (piExitMode == 2)) {
			var sID = "file" + formInputPrefix + psElementItemID + "_17_";

			if (piExitMode == 2) {
				$get("frmMain").elements.namedItem(sID).value = "1";
			}
			else {
				$get("frmMain").elements.namedItem(sID).value = "0";
			}
		}

		showFileUpload(false, '0', 0);
	}
	catch (e) { }
}

function showMessage() {

	//Reset jQuery setup
	jQuerySetup();

	//Reset current tab position
	SetCurrentTab(iCurrentTab);
	//Reset current error message display
	showErrorMessages(window.iCurrentMessageState);

	wait.hide();
	overlay.hide();

	//Reapply resizable column functionality to tables
	//This is put here to ensure functionality is reapplied after partial/full postback.
	ResizableColumns();

	if ($get("txtActiveDDE").value.indexOf("dde") > 0) {
		try {
			$find($get("txtActiveDDE").value).show();
			$get("txtActiveDDE").value = "";
		}
		catch (e) { }
	}

	if ($get("txtPostbackMode").value == 3) {
		//ShowMessage is the sub called in lieu of Application:EndRequest, i.e. Pretty much the end of
		//the postback cycle. So we'll reset all grid scroll bars to their previous position
		SetScrollTopPos("", "-1", 0);
	}

	try {
		if ($get("frmMain").hdnErrorMessage.value.length > 0) {
			showSubmissionMessage();
			return;
		}

		if (($get("txtPostbackMode").value == 2) || ($get("txtPostbackMode").value == 3)) {
			// 0 = Default
			// 1 = Submit/SaveForLater button postback (ie. WebForm submission)
			// 2 = Grid header postback
			// 3 = FileUpload button postback

			// not doing this causes the object referenced is null error:
			setPostbackMode(0);
			return;

		}

		if (hasErrors()) {
			showErrorMessages('max');
		}
		else {
			if ($get("frmMain").hdnNoSubmissionMessage.value == 1) {
				try {
					if ($get("frmMain").hdnFollowOnForms.value.length > 0) {
						launchFollowOnForms($get("frmMain").hdnFollowOnForms.value);
					}
					else {
						overlay.show();

						if (navigator.userAgent.indexOf("MSIE") > 0) {
							//Only IE can self-close windows that it didn't open
							window.close();
						} else {
							// Non-IE browsers can't self-close windows, show close message instead
							wait.show();
							wait.text('Please close your browser.');
						}
					}
				}
				catch (e) { };
			}
			else {
				if ($get("txtPostbackMode").value == 1) {
					showSubmissionMessage();
				}
			}
		}
		setPostbackMode(0);
	}
	catch (e) { }
}

function showSubmissionMessage() {

	try {
		$get("ifrmMessages").src = "SubmissionMessage.aspx";

		overlay.show();
		showErrorMessages('none');
		$get("divSubmissionMessages").style.display = "block";
		$get("divSubmissionMessages").style.visibility = "visible";

		if (window.androidLayerBug) {
			$get("divInput").style.display = "none";	
		}		
	}
	catch (e) { }
}

function FileDownload_Click(id) {
	spawnWindow("FileDownload.aspx?" + id);
}

function FileDownload_KeyPress(id) {
	// If the user presses SPACE (keyCode = 32) launch the file download.
	if (window.event.keyCode == 32) {
		spawnWindow("FileDownload.aspx?" + id);
	}
}

//TODO replace using jQuery date functions
function GetDatePart(psLocaleDateValue, psDatePart) {
	var reDATE = /[YMD]/g;
	var sLocaleDateFormat = window.localeDateFormat;
	var sLocaleDateSep = sLocaleDateFormat.replace(reDATE, "").substr(0, 1);
	var iLoop;
	var iRequiredPart = 1;
	var sValuePart1;
	var sValuePart2;
	var sValuePart3;
	var iPartCounter = 1;
	var sTemp = "";

	for (iLoop = 0; iLoop < psLocaleDateValue.length; iLoop++) {
		if (psLocaleDateValue.substr(iLoop, 1) == sLocaleDateSep) {
			if (iPartCounter == 1) {
				sValuePart1 = sTemp;
			}
			else {
				if (iPartCounter == 2) {
					sValuePart2 = sTemp;
				}
			}

			iPartCounter++;
			sTemp = "";
		}
		else {
			sTemp = sTemp + psLocaleDateValue.substr(iLoop, 1);
		}
	}
	sValuePart3 = sTemp;


	if (psDatePart == "Y") {
		if (sLocaleDateFormat.indexOf("M") < sLocaleDateFormat.indexOf("Y")) {
			iRequiredPart++;
		}
		if (sLocaleDateFormat.indexOf("D") < sLocaleDateFormat.indexOf("Y")) {
			iRequiredPart++;
		}
	}
	else {
		if (psDatePart == "M") {
			if (sLocaleDateFormat.indexOf("Y") < sLocaleDateFormat.indexOf("M")) {
				iRequiredPart++;
			}
			if (sLocaleDateFormat.indexOf("D") < sLocaleDateFormat.indexOf("M")) {
				iRequiredPart++;
			}
		}
		else {
			if (sLocaleDateFormat.indexOf("Y") < sLocaleDateFormat.indexOf("D")) {
				iRequiredPart++;
			}
			if (sLocaleDateFormat.indexOf("M") < sLocaleDateFormat.indexOf("D")) {
				iRequiredPart++;
			}
		}
	}

	if (iRequiredPart == 1) {
		return (sValuePart1);
	}
	else {
		if (iRequiredPart == 2) {
			return (sValuePart2);
		}
		else {
			if (iRequiredPart == 3) {
				return (sValuePart3);
			}
			else {
				return ("");
			}
		}
	}
}

function ResizeComboForForm(sender, args) {

	var psWebComboId = sender._id;

	//Let's set the width of the lookup panel to the width of the screen. 
	//It used to resize the screen, but don't want this happening now.

	try {
		var oEl = document.getElementById(psWebComboId.replace("dde", ""));
		if (eval(oEl)) {
			if (oEl.offsetWidth > $get("bdyMain").clientWidth) {
				oEl.style.width = $get("bdyMain").clientWidth - oEl.offsetLeft - 5 + "px";
				document.getElementById(psWebComboId.replace("dde", "gridcontainer")).style.width = oEl.style.width;
			}

			//also set left position to 0 if required (right coord > bymain.width)
			if ((oEl.offsetLeft + oEl.offsetWidth) > $get("bdyMain").clientWidth) {
				oEl.style.left = "0px";
			}

			//Hide the navigation icons as required
			//Order to hide is: nav arrows go first, then 'page 1 of x'. Finally the search box goes.
			//N.B. if the control is paged, min width is 420px before hiding the relevant controls

			//Check to see if this is a paged control...
			var oElDDL = document.getElementById(psWebComboId.replace("dde", "tcPagerDDL"));
			if (eval(oElDDL)) {
				//This is a paged control, so different rules apply.
				if (oEl.offsetWidth < 420) {
					document.getElementById(psWebComboId.replace("dde", "tcPagerBtns")).style.visibility = "hidden";
					document.getElementById(psWebComboId.replace("dde", "tcPagerBtns")).style.display = "none";
					document.getElementById(psWebComboId.replace("dde", "tcPageXofY")).style.visibility = "hidden";
					document.getElementById(psWebComboId.replace("dde", "tcPageXofY")).style.display = "none";
				}
				else {
					document.getElementById(psWebComboId.replace("dde", "tcPagerBtns")).style.visibility = "visible";
					document.getElementById(psWebComboId.replace("dde", "tcPagerBtns")).style.display = "";
					document.getElementById(psWebComboId.replace("dde", "tcPageXofY")).style.visibility = "visible";
					document.getElementById(psWebComboId.replace("dde", "tcPageXofY")).style.display = "";
				}
			}
			else {
				//Not a paged control
				if (oEl.offsetWidth < 250) {
					document.getElementById(psWebComboId.replace("dde", "tcPagerBtns")).style.visibility = "hidden";
					document.getElementById(psWebComboId.replace("dde", "tcPagerBtns")).style.display = "none";
					document.getElementById(psWebComboId.replace("dde", "tcPageXofY")).style.visibility = "hidden";
					document.getElementById(psWebComboId.replace("dde", "tcPageXofY")).style.display = "none";
				}
				else {
					document.getElementById(psWebComboId.replace("dde", "tcPagerBtns")).style.visibility = "visible";
					document.getElementById(psWebComboId.replace("dde", "tcPagerBtns")).style.display = "";
					document.getElementById(psWebComboId.replace("dde", "tcPageXofY")).style.visibility = "visible";
					document.getElementById(psWebComboId.replace("dde", "tcPageXofY")).style.display = "";
				}
			}
		}
	}
	catch (e) { }
}

function scrollHeader(iGridID) {
	//keeps the header table aligned with the gridview in record selectors and lookups.
	var leftPos = document.getElementById(iGridID).scrollLeft;
	document.getElementById(iGridID.replace("gridcontainer", "Header")).style.left = "-" + leftPos + "px";

	var hdn1 = document.getElementById(iGridID.replace("Grid", "scrollpos"));
	hdn1.value = document.getElementById(iGridID).scrollTop;

}

function InitializeLookup(sender, args) {

	if ($get("txtActiveDDE").value.indexOf("dde") >= 0) {
		// If we're in the process of displaying a filtered lookup already, do nothing and exit the function...
		return false;
	}

	var sSelectWhere = "";
	var sValueID = "";
	var sValueType = "";
	var sControlType = "";
	var sValue = "";
	var sTemp = "";
	var sSubTemp = "";
	var dtValue;
	var iIndex;
	var reTAB = /\t/g;
	var reSINGLEQUOTE = /\'/g;
	var reDECIMAL = new RegExp("\\" + window.localeDecimal, "gi");
	var psWebComboID = "";

	psWebComboID = sender._id;

	if (psWebComboID == "") { return false; }

	var sID = "lookup" + psWebComboID.replace("dde", "");
	try {
		var ctlLookupFilter = document.getElementById(sID);
		if (ctlLookupFilter) {
			sSelectWhere = ctlLookupFilter.value;

			if (sSelectWhere.length > 0) {
				// sSelectWhere has the format:
				//  <filterValueControlID><TAB><selectWhere code with TABs where the value from filterValueControlID is to be inserted>

				iIndex = sSelectWhere.indexOf("\t");
				if (iIndex >= 0) {
					sValueType = sSelectWhere.substring(0, iIndex);
					sSelectWhere = sSelectWhere.substr(iIndex + 1);
				}

				iIndex = sSelectWhere.indexOf("\t");
				if (iIndex >= 0) {
					sValueID = sSelectWhere.substring(0, iIndex);
					sSelectWhere = sSelectWhere.substr(iIndex + 1);

					sControlType = sValueID.substr(sValueID.indexOf("_") + 1);
					sControlType = sControlType.substr(sControlType.indexOf("_") + 1);
					sControlType = sControlType.substring(0, sControlType.indexOf("_"));

					if ((sControlType == 13) || (sControlType == 14)) {
						// Dropdown (13), Lookup (14)
						if (sControlType == 13) {
							sValue = document.getElementById(sValueID).value;
						}
						else {
							sValue = document.getElementById(sValueID + "TextBox").value;
						}

						if (sValueType == 11) {
							// Date value from lookup. Convert from locale format to yyyymmdd.
							if (sValue.length > 0) {
								sTemp = GetDatePart(sValue, "Y");

								sSubTemp = "0" + GetDatePart(sValue, "M");
								sTemp = sTemp + sSubTemp.substr(sSubTemp.length - 2);

								sSubTemp = "0" + GetDatePart(sValue, "D");
								sTemp = sTemp + sSubTemp.substr(sSubTemp.length - 2);

								sValue = sTemp;
							}
							else {
								sValue = "";
							}
						}
						else {
							if ((sValueType == 2) || (sValueType == 4)) {
								// numerics/integers
								if (sValue.length > 0) {
									sValue = sValue.replace(reDECIMAL, ".");
								}
								else {
									sValue = "0";
								}
							}
						}
					}
					else {
						if (sControlType == 6) {
							// Checkbox (6)
							if (document.getElementById(sValueID).checked == true) {
								sValue = "1";
							}
							else {
								sValue = "0";
							}
						}
						else {
							if (sControlType == 5) {
								// Numeric (5)
								sValue = document.getElementById(sValueID).value;
							}
							else {
								if (sControlType == 7) {
									// Date (7)
									//TODO PG have changed date control
									var ctlLookupValueDate = igdrp_getComboById(sValueID);
									dtValue = ctlLookupValueDate.getValue();
									if (dtValue) {
										// Get year part.
										sTemp = dtValue.getFullYear();

										// Get month part. Pad to 2 digits if required.
										sSubTemp = "0" + (dtValue.getMonth() + 1);
										sTemp = sTemp + sSubTemp.substr(sSubTemp.length - 2);

										// Get day part. Pad to 2 digits if required.
										sSubTemp = "0" + dtValue.getDate();
										sValue = sTemp + sSubTemp.substr(sSubTemp.length - 2);
									}
									else {
										sValue = "";
									}
								}
								else {
									// CharInput, OptionGroup
									var ctlLookupValue = document.getElementById(sValueID);
									sValue = ctlLookupValue.value;
								}
							}
						}
					}

					sValue = sValue.toUpperCase().trim().replace(reSINGLEQUOTE, "\'\'");
					sSelectWhere = sSelectWhere.replace(reTAB, sValue);

					if (sValue == "") {
						document.getElementById(psWebComboID.replace("dde", "filterSql")).value = "";
					}
					else {
						document.getElementById(psWebComboID.replace("dde", "filterSql")).value = sSelectWhere;
					}

					//This prevents the lookup closing after the filter is applied/removed

					$get("txtActiveDDE").value = psWebComboID;

					setPostbackMode(3);

					//These lines hide the lookup dropdown until it's filled with data.
					document.getElementById(psWebComboID.replace("dde", "")).style.height = "0px";
					document.getElementById(psWebComboID.replace("dde", "")).style.width = "0px";

					//This clicks the server-side button to apply filtering...                          
					//this also kicks off the gosubmit() via postback beginrequest.                          
					document.getElementById(psWebComboID.replace("dde", "refresh")).click();

					//set pbmode back to 0 to prevent recursion.                          
					setPostbackMode(0);
				}
			}
		}
	}
	catch (e) { }

	return false;
}

function FilterMobileLookup(sourceControlID) {

	var sSelectWhere = "";
	var sValueID = "";
	var sValueType = "";
	var sControlType = "";
	var sValue = "";
	var sTemp = "";
	var sSubTemp = "";
	var dtValue;
	var iIndex;
	var reTAB = /\t/g;
	var reSINGLEQUOTE = /\'/g;
	var reDECIMAL = new RegExp("\\" + window.localeDecimal, "gi");
	var psWebComboID;

	if (sourceControlID == "") { return; }

	var lookups = getElementsBySearchValue(sourceControlID);
	var AllLookupIDs = "";

	for (var i = 0; i < lookups.length; i++) {

		try {
			psWebComboID = lookups[i].name.replace("lookup", "");
		}
		catch (e) { psWebComboID = ""; }


		if (psWebComboID.length > 0) {

			var sId = "lookup" + psWebComboID;
			AllLookupIDs = AllLookupIDs + (i == 0 ? "" : "\t") + psWebComboID + "refresh";

			try {
				var ctlLookupFilter = document.getElementById(sId);
				if (ctlLookupFilter) {
					sSelectWhere = ctlLookupFilter.value;

					if (sSelectWhere.length > 0) {
						// sSelectWhere has the format:
						//  <filterValueControlID><TAB><selectWhere code with TABs where the value from filterValueControlID is to be inserted>

						iIndex = sSelectWhere.indexOf("\t");
						if (iIndex >= 0) {
							sValueType = sSelectWhere.substring(0, iIndex);
							sSelectWhere = sSelectWhere.substr(iIndex + 1);
						}

						iIndex = sSelectWhere.indexOf("\t");
						if (iIndex >= 0) {
							sValueID = sSelectWhere.substring(0, iIndex);
							sSelectWhere = sSelectWhere.substr(iIndex + 1);

							sControlType = sValueID.substr(sValueID.indexOf("_") + 1);
							sControlType = sControlType.substr(sControlType.indexOf("_") + 1);
							sControlType = sControlType.substring(0, sControlType.indexOf("_"));

							if (sControlType == 13 || sControlType == 14) {
								// Dropdown (13), Lookup (14)
								if (sControlType == 13) {
									sValue = document.getElementById(sValueID).value;
								}
								else {
									var ctlLookupValueCombo = document.getElementById(sValueID + "TextBox");
									if (!(eval(ctlLookupValueCombo))) { ctlLookupValueCombo = document.getElementById(sValueID); }

									sValue = ctlLookupValueCombo.value;
								}

								if (sValueType == 11) {
									// Date value from lookup. Convert from locale format to yyyymmdd.
									if (sValue.length > 0) {
										sTemp = GetDatePart(sValue, "Y");

										sSubTemp = "0" + GetDatePart(sValue, "M");
										sTemp = sTemp + sSubTemp.substr(sSubTemp.length - 2);

										sSubTemp = "0" + GetDatePart(sValue, "D");
										sTemp = sTemp + sSubTemp.substr(sSubTemp.length - 2);

										sValue = sTemp;
									}
									else {
										sValue = "";
									}
								}
								else {
									if ((sValueType == 2) || (sValueType == 4)) {
										// numerics/integers
										if (sValue.length > 0) {
											sValue = sValue.replace(reDECIMAL, ".");
										}
										else {
											sValue = "0";
										}
									}
								}
							}
							else {
								if (sControlType == 6) {
									// Checkbox (6)
									if (document.getElementById(sValueID).checked == true) {
										sValue = "1";
									}
									else {
										sValue = "0";
									}
								}
								else {
									if (sControlType == 5) {
										// Numeric (5)
										sValue = document.getElementById(sValueID).value;
									}
									else {
										if (sControlType == 7) {
											// Date (7) 
											//TODO PG have changed date control
											var ctlLookupValueDate = igdrp_getComboById(sValueID);
											dtValue = ctlLookupValueDate.getValue();
											if (dtValue) {
												// Get year part.
												sTemp = dtValue.getFullYear();

												// Get month part. Pad to 2 digits if required.
												sSubTemp = "0" + (dtValue.getMonth() + 1);
												sTemp = sTemp + sSubTemp.substr(sSubTemp.length - 2);

												// Get day part. Pad to 2 digits if required.
												sSubTemp = "0" + dtValue.getDate();
												sValue = sTemp + sSubTemp.substr(sSubTemp.length - 2);
											}
											else {
												sValue = "";
											}
										}
										else {
											// CharInput, OptionGroup
											sValue = document.getElementById(sValueID).value;
										}
									}
								}
							}

							sValue = sValue.toUpperCase().trim().replace(reSINGLEQUOTE, "\'\'");
							sSelectWhere = sSelectWhere.replace(reTAB, sValue);

							if (sValue == "") {
								document.getElementById(psWebComboID + "filterSql").value = "";
							}
							else {
								document.getElementById(psWebComboID + "filterSql").value = sSelectWhere;
							}
						}
					}
				}
			}
			catch (e) { }
		}
	}
	setPostbackMode(3);
	document.getElementById("hdnMobileLookupFilter").value = AllLookupIDs;

	if (AllLookupIDs.length > 0) {
		$get("frmMain").btnDoFilter.click();
	}
}

function Right(str, n) {
	if (n <= 0)
		return "";
	else if (n > String(str).length)
		return str;
	else {
		var iLen = String(str).length;
		return String(str).substring(iLen, iLen - n);
	}
}

function isGridFiltered(iGridID) {
	//searches the specified table for hidden rows and returns true if any are found...
	var table = document.getElementById(iGridID);

	for (var r = 0; r < table.rows.length; r++) {
		if (table.rows[r].style.display == 'none') {
			return true;
		}
	}
	return false;
}

function GetGridRowHeight(iGridID) {
	var table = document.getElementById(iGridID);

	for (var r = 0; r < table.rows.length; r++) {
		if (table.rows[r].style.display == '') {
			var rows = document.getElementById(iGridID).rows;
			return (rows[r].offsetHeight);
		}
	}
	return 0;
}


function SetScrollTopPos(iGridID, iPos, iRowIndex) {
	if (iPos == -1) {
		// -1 is the 'code' to reset scrollbar to stored position
		//Loop through all hidden scroll fields and reset values.
		var controlCollection = $get("frmMain").elements;
		if (controlCollection != null) {
			for (var i = 0; i < controlCollection.length; i++) {
				if (Right(controlCollection.item(i).name, 9) == "scrollpos") {
					document.getElementById(controlCollection.item(i).name.replace("scrollpos", "gridcontainer")).scrollTop = (controlCollection.item(i).value);
				}
			}
		}
	}
	else {
		//Check if this grid is quick-filtered (NOT lookup filtered)
		//If it is, calculate the scroll position to use after postback,
		//otherwise store the current scroll position for postback...
		if (isGridFiltered(iGridID)) {
			iPos = (iRowIndex * GetGridRowHeight(iGridID)) - 1;
		}
		//store the scrollbar position
		document.getElementById(iGridID.replace("Grid", "scrollpos")).value = iPos;
	}
}

function SetCurrentTab(iNewTab) {

	if (iNewTab < 1) { iNewTab = 1; }

	jQuery('.tab').removeClass('active');
	jQuery('.tab-page').hide();

	jQuery('#' + formInputPrefix + iNewTab + '_21_Panel').addClass('active');
	jQuery('#' + formInputPrefix + iNewTab + '_21_PageTab').show();

	window.iCurrentTab = iNewTab;
	document.getElementById("hdnDefaultPageNo").value = iNewTab;
}

