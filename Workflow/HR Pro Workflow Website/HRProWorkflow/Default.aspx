<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" id="htmMain">
<head runat="server">
	<title></title>
	<meta http-equiv="refresh" content="<%=Session("TimeoutSecs")%>;URL=timeout.aspx" />

	<script language="javascript" type="text/javascript">
    // <!CDATA[
		function window_onload() {
			var iDefHeight;
			var iDefWidth;
			var iResizeByHeight;
			var iResizeByWidth;
            var sControlType;
            
			try {
				iDefHeight = frmMain.hdnFormHeight.value;
				iDefWidth = frmMain.hdnFormWidth.value;

				bdyMain.bgColor = frmMain.hdnColourThemeHex.value;

				window.focus();

				if ((iDefHeight > 0) && (iDefWidth > 0)) {
					iResizeByHeight = iDefHeight - document.documentElement.clientHeight;
					iResizeByWidth = iDefWidth - document.documentElement.clientWidth;
					window.parent.moveTo((screen.availWidth - iDefWidth) / 2, (screen.availHeight - iDefHeight) / 3);
					window.parent.resizeBy(iResizeByWidth, iResizeByHeight);
				}

				try {
					if (frmMain.hdnFirstControl.value.length > 0) {
					    sControlType = frmMain.hdnFirstControl.value.substr(frmMain.hdnFirstControl.value.indexOf("_")+1);
                        sControlType = sControlType.substr(sControlType.indexOf("_")+1);
                        sControlType = sControlType.substring(0, sControlType.indexOf("_"));

                        if (sControlType == 7)
                        {
                            // Date (7)
                            igdrp_getComboById(frmMain.hdnFirstControl.value).focus();
                        }
                        else
                        {
                            if ((sControlType == 13)
                                || (sControlType == 14))
                            {
                                igcmbo_getComboById(frmMain.hdnFirstControl.value).focus();
                            }
                            else
                            {
                                if (sControlType == 11)
                                {
                                    // Record Selector (11)
                                    var grid = igtbl_getGridById(frmMain.hdnFirstControl.value);
                                    var oRows = grid.Rows;
                                    grid.Element.focus(); 
                                    
                                    if (oRows.length > 0)
                                    {
                                        oRow = grid.getActiveRow();
	                                    if (oRow != null)
	                                    {
                                            oRow.scrollToView();
                                        }
                                    }
                                }
                                else
                                {
						            document.getElementById(frmMain.hdnFirstControl.value).setActive();
						        }
						    }
                        }
					}
				}
				catch (e) { }

				if ((iDefHeight > 0) && (iDefWidth > 0)) {
					iResizeByHeight = iDefHeight - document.documentElement.clientHeight;
					iResizeByWidth = iDefWidth - document.documentElement.clientWidth;
					window.parent.resizeBy(iResizeByWidth, iResizeByHeight);
				}


				launchForms(frmMain.hdnSiblingForms.value, false);
			}
			catch (e) {}
		}

		function resizeToFit(piWidth, piHeight) {
			var iDefHeight;
			var iDefWidth;
			var iResizeByHeight;
			var iResizeByWidth;

			try {
				iDefHeight = frmMain.hdnFormHeight.value;
				iDefWidth = frmMain.hdnFormWidth.value;

				iResizeByHeight = piHeight - htmMain.clientHeight;
				iResizeByWidth = piWidth - htmMain.clientWidth;

				if (iResizeByHeight < 0) {
					iResizeByHeight = 0;
				}
				if (iResizeByWidth < 0) {
					iResizeByWidth = 0;
				}

				window.parent.resizeBy(iResizeByWidth, iResizeByHeight);
			}
			catch (e) { }
		}

		function launchForms(psForms, pfFirstFormRelocate) {
			var asForms;
			var iLoop;
			var iCount;
			var sQueryString;
			var sFirstForm;

			try {
				iCount = 0;
				sFirstForm = "";
				asForms = psForms.split("\t");

				for (iLoop = 1; iLoop < asForms.length; iLoop++) {
					sQueryString = asForms[iLoop];

					if (sQueryString.length > 0) {
						iCount = iCount + 1;

						if (iCount == 1) {
							sFirstForm = sQueryString;
						}
						else {
							// Open other forms in new browsers.
							spawnWindow(sQueryString);
						}
					}
				}

				if (sFirstForm.length > 0) {
					if (pfFirstFormRelocate == true) {
						// Open first form in current browser.
						window.location = sFirstForm;
					}
					else {
						// Open first form in new browser.
						spawnWindow(sFirstForm);
					}
				}
			}
			catch (e) { }
		}

		function spawnWindow(psURL) {
			var newWin;

			try {
				newWin = window.open(psURL);

				if (parseInt(navigator.appVersion) >= 4) {
					try {
						newWin.window.focus();
					}
					catch (e) { }
				}
			}
			catch (e) {
				try {
					newWin.close();
				}
				catch (e) { }

				spawnWindow(psURL);
			}
		}

		function goSubmit() {
			disableChildElements("pnlInput");
			showErrorMessages(false);
		}

		function closeOtherCombos(objId) {
			var theObject = document.getElementById(objId);
			var level = 0;

            // Tell the TraverseDOM function to run the doNothing function on each control. 
            // The TraverseDOM function already has code close all WebCombos, so a 'doNothibng ios all that is required.
			TraverseDOM(theObject, level, doNothing);
		}

		function doNothing(obj) {
		    // Empty function. Required - See note for closeOtherCombos function.
		}

		function disableChildElements(objId) {
		    try
		    {
			    var theObject = document.getElementById(objId);
			    var level = 0;

			    TraverseDOM(theObject, level, disableElement);
    		}
    		catch(e) {}
		}

		function disableElement(obj) {
		    try
		    {
    			obj.disabled = true;
    		}
    		catch(e) {}
		}

		function TraverseDOM(obj, lvl, actionFunc) {
		    var sControlType;
		    var sFormInputPrefix = "forminput_";
		    var sGridSuffix = "Grid";
		    
		    try
		    {
    			for (var i = 0; i < obj.childNodes.length; i++) {
    				var childObj = obj.childNodes[i];

                    // Close any lookup/dropdown grids.
                    try
                    {
                        if (childObj.id != undefined) {
                            if (childObj.id.substr(0, "forminput_".length) == "forminput_")
                            {
                                sControlType = childObj.id.substr(childObj.id.indexOf("_")+1);
                                sControlType = sControlType.substr(sControlType.indexOf("_")+1);
                                sControlType = sControlType.substring(0, sControlType.indexOf("_"));

                                if ((sControlType == 13)
                                    || (sControlType == 14))
                                {
                                    if ((childObj.id.substr(0, sFormInputPrefix.length) == sFormInputPrefix) &&
                                        (childObj.id.substr(childObj.id.length - sGridSuffix.length) != sGridSuffix))
                                    {
                                        igcmbo_getComboById(childObj.id).setDropDown(false);
                                    }
                                }
                            }
                        }
                    }
                    catch(e){}

				    if (childObj.tagName) 
				    {
					    actionFunc(childObj);
				    }

				    TraverseDOM(childObj, lvl + 1, actionFunc);
	    		}
	    	}
	    	catch(e) {}
		}

		function showErrorMessages(pfDisplay) {
			if (((frmMain.hdnCount_Errors.value > 0)
				|| (frmMain.hdnCount_Warnings.value > 0))
				&& (pfDisplay == false)) {
				imgErrorMessages_Max.style.display = "block";
				imgErrorMessages_Max.style.visibility = "visible";
			}
			else {
				imgErrorMessages_Max.style.display = "none";
				imgErrorMessages_Max.style.visibility = "hidden";
			}

			if (pfDisplay == true) {
				//divErrorMessages_Inner.style.visibility = "visible";
				divErrorMessages_Outer.style.filter = "revealTrans(duration=0.3, transition=4)";
				divErrorMessages_Outer.filters.revealTrans.apply();
				divErrorMessages_Outer.style.display = "block";
				divErrorMessages_Outer.style.visibility = "visible";
				divErrorMessages_Outer.filters.revealTrans.play();
			}
			else {
				divErrorMessages_Outer.style.filter = "revealTrans(duration=0.3, transition=5)";
				divErrorMessages_Outer.filters.revealTrans.apply();
				divErrorMessages_Outer.style.visibility = "hidden";
				//divErrorMessages_Outer.style.display = "none";
				//divErrorMessages_Inner.style.visibility = "hidden";
				divErrorMessages_Outer.filters.revealTrans.play();

			}
		}

		function launchFollowOnForms(psForms) {
			launchForms(psForms, true);
		}

		function overrideWarningsAndSubmit() {
			if (divErrorMessages_Outer.disabled == true) {
				return;
			};

			frmMain.hdnOverrideWarnings.value = 1;

			try {
				document.getElementById(frmMain.hdnLastButtonClicked.value).click();
			}
			catch (e) {
				frmMain.btnSubmit.click();
			}
		}

		function submitForm() {
			try {
				if (txtPostbackMode.value == 0) {
					document.getElementById(txtActiveElement.value).setActive();
				}
			}
			catch (e) { };

			return (txtPostbackMode.value != 0);
		}

		function setPostbackMode(piValue) {
			// 0 = Default
			// 1 = Submit/SaveForLater button postback (ie. WebForm submission)
			// 2 = Grid header postback
			// 3 = FileUpload button postback
            try {
				txtPostbackMode.value = piValue;
			}
			catch (e) { }
		}

		function activateGridPostback() {
			setPostbackMode(2);
		}

		function activateControl() {
			try {
				txtActiveElement.value = document.activeElement.id;
			}
			catch (e) { }
		}

		function checkMaxLength(iMaxLength) {
			var sClipboardText;
			var iResultantLength;
			var iCurrentFieldLength;
			var fIsPermittedKeystroke;
			var iEnteredKeystroke;
			var fActionAllowed = true;
			var iSelectionLength;

			try {
				if (iMaxLength > 0) {
					iSelectionLength = parseInt(document.selection.createRange().text.length);
					iCurrentFieldLength = parseInt(event.srcElement.value.length);

					if (event.type == "keydown") {
						// Allow non-printing, arrow and delete keys
						iEnteredKeystroke = window.event.keyCode;
						fIsPermittedKeystroke = (((iEnteredKeystroke < 32)			// Non printing - don't count
							|| (iEnteredKeystroke >= 33 && iEnteredKeystroke <= 40)	// Page Up, Down, Home, End, Arrow - don't count
							|| (iEnteredKeystroke == 46))							// Delete - doesn't count
							&& (iEnteredKeystroke != 13))							// Enter - does count

						// Decide whether the keystroke is allowed to proceed
						if (!fIsPermittedKeystroke) {
							if ((iCurrentFieldLength - iSelectionLength) >= iMaxLength) {
								fActionAllowed = false;
							}
						}

						window.event.returnValue = fActionAllowed;
						return (fActionAllowed);
					}

					if (event.type == "paste") {
						sClipboardText = window.clipboardData.getData("Text");
						iResultantLength = iCurrentFieldLength + sClipboardText.length - iSelectionLength;

						if (iResultantLength > iMaxLength) {
							fActionAllowed = false;
						}

						window.event.returnValue = fActionAllowed;
						return (fActionAllowed);
					}
				}
			}
			catch (e) { }
		}

		function dropdownControlKeyPress(pobjControlID, pNewValue, piKeyCode) {
			try {
				activateControl();

				if (piKeyCode == 32) // SPACE - drop list
				{
					var objCombo1 = igcmbo_getComboById(pobjControlID);
					objCombo1.setDropDown(true);
				}
				if (piKeyCode == 13) // RTN - close list
				{
					var objCombo2 = igcmbo_getComboById(pobjControlID);
					objCombo2.setDropDown(false);
				}
			}
			catch (e) { }
		}

		function dateControlKeyPress(pobjControl, piKeyCode, pobjEvent) {
			try {
				activateControl();

				if (piKeyCode == 113) // F2 - set today's date
				{
					var d = new Date();
					pobjControl.setValue(d);
				}
				if (piKeyCode == 117) // F6 - show calendar
				{
					pobjControl.setDropDownVisible(true);
				}
			}
			catch (e) { }
		}

		function dateControlTextChanged(pobjControl, pNewText, pobjEvent) {
			var sDate;
			var dtCurrentDate;

			try {
				if (pNewText.length > 0) {
					dtCurrentDate = pobjControl.getValue();
					txtLastDate_Month.value = dtCurrentDate.getMonth();
					txtLastDate_Day.value = dtCurrentDate.getDate();
					txtLastDate_Year.value = dtCurrentDate.getYear();
				}
			}
			catch (e) { }
		}

		function dateControlBeforeDropDown(pobjControl, pPanel, pobjEvent) {
			try {
				var sCurrentText = pobjControl.getText();
				var sLastDate_Month = txtLastDate_Month.value;
				var sLastDate_Day = txtLastDate_Day.value;
				var sLastDate_Year = txtLastDate_Year.value;
				var dtLastDate;

				if ((sCurrentText.length == 0)
                    && (sLastDate_Month.length > 0)
                    && (sLastDate_Day.length > 0)
                    && (sLastDate_Year.length > 0)) {
					dtLastDate = new Date(sLastDate_Year, sLastDate_Month, sLastDate_Day);
					pobjControl.Calendar.setSelectedDate(dtLastDate);
				}
			}
			catch (e) { }
		}

		function showFileUpload(pfDisplay, psElementItemID, psAlreadyUploaded) {
		
			try {
				if (pfDisplay == true) {

					setPostbackMode(3);

					var sAlreadyUploaded = new String(psAlreadyUploaded);
					sAlreadyUploaded = sAlreadyUploaded.substr(0, 1);
					if (sAlreadyUploaded != "1") {
						sAlreadyUploaded = "0";
					}

					try {
						txtActiveElement.value = document.activeElement.id;
					}
					catch (e) { }

					document.all.ifrmFileUpload.src = "FileUpload.aspx?" + sAlreadyUploaded + psElementItemID;

					showErrorMessages(false);

					divInput.disabled = true;
					divErrorMessages_Outer.disabled = true;
					imgErrorMessages_Max.disabled = true;
					divErrorMessages_Outer.style.display = "none";

					divFileUpload.style.filter = "revealTrans(duration=0.5, transition=12)";
					divFileUpload.filters.revealTrans.apply();
					divFileUpload.style.visibility = "visible";
					divFileUpload.style.display = "block";
					divFileUpload.filters.revealTrans.play();
				}
				else {
					divFileUpload.style.filter = "revealTrans(duration=0.5, transition=12)";
					divFileUpload.filters.revealTrans.apply();
					divFileUpload.style.visibility = "hidden";
					divFileUpload.style.display = "none";
					divFileUpload.filters.revealTrans.play();

					setPostbackMode(3);
					
					frmMain.btnReEnableControls.click();

					divInput.disabled = false;
					divErrorMessages_Outer.disabled = false;
					imgErrorMessages_Max.disabled = false;
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
					var sID = "fileforminput_" + psElementItemID + "_17_";

					if (piExitMode == 2) {
						frmMain.elements.namedItem(sID).value = "1";
					}
					else {
						frmMain.elements.namedItem(sID).value = "0";
					}
				}

				showFileUpload(false, '0', 0);
			}
			catch (e) { }
		}

		function unblockErrorMessageDIV() {
			try {
				if ((divErrorMessages_Outer.style.visibility == "hidden") &&
					(divErrorMessages_Outer.style.display != "none")) {
					divErrorMessages_Outer.style.display = "none";
				}
			}
			catch (e) { }
		}

		function showMessage() {
			try {
				if (frmMain.hdnErrorMessage.value.length > 0) {
					showSubmissionMessage();
					return;
				}

				refreshLiterals();

				if ((txtPostbackMode.value == 2)
                    || (txtPostbackMode.value == 3)) 
                {
					// 0 = Default
					// 1 = Submit/SaveForLater button postback (ie. WebForm submission)
					// 2 = Grid header postback
					// 3 = FileUpload button postback
					
					if (txtPostbackMode.value == 3) 
					{
					    document.all.ifrmFileUpload.contentWindow.enableControls();
                    }
					setPostbackMode(0);
					return;
				}

				if ((frmMain.hdnCount_Errors.value > 0)
			        || (frmMain.hdnCount_Warnings.value > 0)) {
					showErrorMessages(true);
				}
				else {
					if (frmMain.hdnNoSubmissionMessage.value == 1) {
						try {
							if (frmMain.hdnFollowOnForms.value.length > 0) {
								launchFollowOnForms(frmMain.hdnFollowOnForms.value);
							}
							else {
								window.close();
							}
						}
						catch (e) { };
					}
					else {
						if (txtPostbackMode.value == 1) {
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
				document.all.ifrmMessages.src = "SubmissionMessage.aspx";

				divInput.disabled = true;
				frmMain.hdnCount_Errors.value = 0;
				frmMain.hdnCount_Warnings.value = 0;
				divErrorMessages_Outer.style.display = "none";
				showErrorMessages(false);

				divSubmissionMessages.style.filter = "revealTrans(duration=0.5, transition=12)";
				divSubmissionMessages.filters.revealTrans.apply();
				divSubmissionMessages.style.display = "block";
				divSubmissionMessages.style.visibility = "visible";
				divSubmissionMessages.filters.revealTrans.play();
			}
			catch (e) { }
		}

		function unblockFileUploadDIV() {
			try {
				if ((divFileUpload.style.visibility == "hidden") &&
					(divFileUpload.style.display != "none")) {
					divFileUpload.style.display = "none";
				}
			}
			catch (e) { }

			try {
				document.getElementById(txtActiveElement.value).setActive();
			}
			catch (e) { }
		}

		function FileDownload_Click(psID) {
			spawnWindow("FileDownload.aspx?" + psID);
		}

		function FileDownload_KeyPress(psID) {
			// If the user presses SPACE (keyCode = 32) launch the file download.
			if (window.event.keyCode == 32) {
				spawnWindow("FileDownload.aspx?" + psID);
			}
		}

		function WARP_SetTimeout() {
			ig_shared.getCBManager()._timeLimit = <%=SubmissionTimeout()%>;
		}
	    
		function GetDatePart(psLocaleDateValue, psDatePart) {
            var reDATE = /[YMD]/g;        
            var sLocaleDateFormat = "<%=LocaleDateFormat()%>";
            var sLocaleDateSep = sLocaleDateFormat.replace(reDATE, "").substr(0, 1);
            var iLoop;
            var iRequiredPart = 1;
            var sValuePart1;
            var sValuePart2;
            var sValuePart3;
            var iPartCounter = 1;
            var sTemp = "";

            for (iLoop=0; iLoop<psLocaleDateValue.length; iLoop++)
            {
                if (psLocaleDateValue.substr(iLoop, 1) == sLocaleDateSep)
                {
                    if (iPartCounter == 1)
                    {
                        sValuePart1 = sTemp;
                    }
                    else
                    {
                        if (iPartCounter == 2)
                        {
                            sValuePart2 = sTemp;
                        }
                    }
                    
                    iPartCounter++;
                    sTemp = "";
                }
                else
                {
                    sTemp = sTemp + psLocaleDateValue.substr(iLoop, 1);
                }
            }
            sValuePart3 = sTemp;

            
            if (psDatePart == "Y")
            {    
                if (sLocaleDateFormat.indexOf("M") < sLocaleDateFormat.indexOf("Y"))
                {
                    iRequiredPart++;
                }
                if (sLocaleDateFormat.indexOf("D") < sLocaleDateFormat.indexOf("Y"))
                {
                    iRequiredPart++;
                }
            }
            else
            {
                if (psDatePart == "M")
                {
                    if (sLocaleDateFormat.indexOf("Y") < sLocaleDateFormat.indexOf("M"))
                    {
                        iRequiredPart++;
                    }
                    if (sLocaleDateFormat.indexOf("D") < sLocaleDateFormat.indexOf("M"))
                    {
                        iRequiredPart++;
                    }
                }
                else
                {
                    if (sLocaleDateFormat.indexOf("Y") < sLocaleDateFormat.indexOf("D"))
                    {
                        iRequiredPart++;
                    }
                    if (sLocaleDateFormat.indexOf("M") < sLocaleDateFormat.indexOf("D"))
                    {
                        iRequiredPart++;
                    }
                }
            }

            if (iRequiredPart == 1)
            {
                return (sValuePart1);
            }
            else
            {
                if (iRequiredPart == 2)
                {
                    return (sValuePart2);
                }
                else
                {
                    if (iRequiredPart == 3)
                    {
                        return (sValuePart3);
                    }
                    else
                    {
                        return ("");
                    }
                }
            }
		}
	    
	    function ChangeLookup(psWebComboId) {
	        // Ensure locale number formatting is applied.
	        try
	        {
                var sLocaleDecimal = "<%=LocaleDecimal()%>";
                var reDECIMAL = /\./g;        
                var objCombo = igcmbo_getComboById(psWebComboId);
	            objCombo.setDisplayValue(objCombo.displayValue.replace(reDECIMAL, sLocaleDecimal));
	        }
	        catch(e) {}
	    }
	    
	    function ResizeFormForCombo(psWebComboId) {
			var iResizeByHeight = 0;
			var iResizeByWidth = 0;
            
			try {
	            var objCombo = igcmbo_getComboById(psWebComboId);
                var grid = objCombo.getGrid();

                var oEl = grid.Element;

                if (oEl.scrollWidth > bdyMain.clientWidth)
                {
                    if (oEl.scrollWidth > screen.availWidth)
                    {
                        iResizeByWidth = screen.availWidth - bdyMain.clientWidth;
                    }
                    else
                    {
                        iResizeByWidth = oEl.scrollWidth - bdyMain.clientWidth;
                    }
                }
                
//                if (oEl.scrollHeight > bdyMain.clientHeight)
//                {
//                    if (oEl.scrollHeight > screen.availHeight)
//                    {
//                        iResizeByHeight = screen.availHeight - bdyMain.clientHeight;
//                    }
//                    else
//                    {
//                        iResizeByHeight = oEl.scrollHeight - bdyMain.clientHeight;
//                    }
//                }
                
                if ((iResizeByWidth > 0) || (iResizeByHeight > 0))
                {
                    setTimeout('window.resizeBy(' + iResizeByWidth + ',' + iResizeByHeight + ');', 100);
                }
            }
            catch(e) {}
	    }
	    
	    function InitializeLookup(psWebComboId) {
	        var sSelectWhere = "";
	        var sValueID = "";
	        var sValueType = "";
	        var sControlType = "";
            var sValue = "";
            var sTemp = "";
            var sSubTemp = "";
            var numValue = 0;
            var dtValue;
            var fValue = true;
            var iIndex;
            var iTemp;
            var reX = /x/g;        
            var reDATE = /[YMD]/g;        
            var reTAB = /\t/g;        
            var reSINGLEQUOTE = /\'/g;        
            var sLocaleDecimal = "\\<%=LocaleDecimal()%>";
        	var reDECIMAL = new RegExp(sLocaleDecimal, "gi");
	
		    try {
                closeOtherCombos(psWebComboId);

			    var sID = "lookup" + psWebComboId;
			    
			    // For some reason the WebComboID has '_' replaced with 'x'.
			    // Replace 'x' with '_' as required to get the required ID.
                sID = sID.replace(reX, "_");   

                var ctlLookupFilter = document.getElementById(sID);
                if (ctlLookupFilter)
                { 
                    sSelectWhere = ctlLookupFilter.value; 

	                if (sSelectWhere.length > 0)
	                {
	                    // sSelectWhere has the format:
	                    //  <filterValueControlID><TAB><selectWhere code with TABs where the value from filterValueControlID is to be inserted>
                        iIndex = sSelectWhere.indexOf("\t");
                        if (iIndex >= 0)
                        {
                            sValueType = sSelectWhere.substring(0, iIndex);
                            sSelectWhere = sSelectWhere.substr(iIndex+1);
                        }
                        
                        iIndex = sSelectWhere.indexOf("\t");
                        if (iIndex >= 0)
                        {
                            sValueID = sSelectWhere.substring(0, iIndex);
                            sSelectWhere = sSelectWhere.substr(iIndex+1);

                            sControlType = sValueID.substr(sValueID.indexOf("_")+1);
                            sControlType = sControlType.substr(sControlType.indexOf("_")+1);
                            sControlType = sControlType.substring(0, sControlType.indexOf("_"));
                            
                            if ((sControlType == 13)
                                || (sControlType == 14))
                            {
                                // Dropdown (13), Lookup (14)
                                var ctlLookupValueCombo = igcmbo_getComboById(sValueID);
                        	    sValue = ctlLookupValueCombo.displayValue;
                        	    
                        	    if(sValueType == 11)
                        	    {
                        	        // Date value from lookup. Convert from locale format to yyyymmdd.
                        	        if (sValue.length > 0)
                        	        {
                        	            sTemp = GetDatePart(sValue, "Y");
                        	             
                        	            sSubTemp = "0" + GetDatePart(sValue, "M");
                        	            sTemp = sTemp + sSubTemp.substr(sSubTemp.length-2);
                        	            
                        	            sSubTemp = "0" + GetDatePart(sValue, "D");
                        	            sTemp = sTemp + sSubTemp.substr(sSubTemp.length-2);

                        	            sValue = sTemp;
                          	        }
                        	        else
                        	        {
                        	            sValue = "";
                        	        }
                        	    }
                        	    else
                        	    {
                        	        if((sValueType == 2) || (sValueType == 4))
                        	        {
                        	            // numerics/integers
                        	            if (sValue.length > 0)
                        	            {
                        	                sValue = sValue.replace(reDECIMAL, ".");
                        	            }
                        	            else
                        	            {
                        	                sValue = "0";
                        	            }
                        	        }
                        	    }
                            }
                            else
                            {
                                if (sControlType == 6)
                                {
                                    // Checkbox (6)
                                    var ctlLookupValueCheckbox = document.getElementById(sValueID);
                        	        fValue = ctlLookupValueCheckbox.checked;
                                    if (fValue == true)
                                    {
                                        sValue = "1";
                                    }
                                    else
                                    {
                                        sValue = "0";
                                    }
                                }
                                else
                                {
                                    if (sControlType == 5)
                                    {
                                        // Numeric (5)
                                        var ctlLookupValueNumeric = igedit_getById(sValueID);
                    	                numValue = ctlLookupValueNumeric.getValue();
                                        sValue = numValue.toString();
                                    }
                                    else
                                    {
                                        if (sControlType == 7)
                                        {
                                            // Date (7)
                                            var ctlLookupValueDate = igdrp_getComboById(sValueID);
                    	                    dtValue = ctlLookupValueDate.getValue();
                    	                    if (dtValue)
                    	                    {
                                	            // Get year part.
                        	                    sTemp = dtValue.getFullYear();
                        	            
                        	                    // Get month part. Pad to 2 digits if required.
                        	                    sSubTemp = "0" + (dtValue.getMonth() + 1);
                                	            sTemp = sTemp + sSubTemp.substr(sSubTemp.length-2);

                        	                    // Get day part. Pad to 2 digits if required.
                                	            sSubTemp = "0" + dtValue.getDate();
                                	            sValue = sTemp + sSubTemp.substr(sSubTemp.length-2);
                                            }
                                            else
                                            {
                                                sValue = "";
                                            }
                                        }
                                        else
                                        {
                                            // CharInput, OptionGroup
	                                        var ctlLookupValue = document.getElementById(sValueID);
	                                        sValue = ctlLookupValue.value;
	                                    }
                                    }
	                            }
	                        }

	                        sValue = sValue.toUpperCase().trim().replace(reSINGLEQUOTE, "\'\'"); 
                            sSelectWhere = sSelectWhere.replace(reTAB, sValue);   
					        var objCombo = igcmbo_getComboById(psWebComboId);
	                        objCombo.selectWhere(sSelectWhere);
                        }
	                }
                }
            }
            catch (e) {}

	        return false;
		}
    // ]]>
	</script>

	<script src="scripts\WebNumericEditValidation.js" type="text/javascript"></script>

</head>
<body id="bdyMain" onload="return window_onload()" scroll="auto" style="overflow: auto;  
	text-align: center; margin: 0px; padding: 0px;">
	<img id="imgErrorMessages_Max" src="Images/uparrows_white.gif" alt="Show messages"
		style="position: absolute; right: 1px; bottom: 1px; display: none; visibility: hidden;
		z-index: 1;" onclick="showErrorMessages(true);" />
	<form runat="server" hidefocus="true" id="frmMain" onsubmit="return submitForm();">
	<asp:ScriptManager ID="ScriptManager1" runat="server">
	</asp:ScriptManager>
	<!--
        Web Form Validation Error Messages
        -->
	<div id="divErrorMessages_Outer" onfilterchange="unblockErrorMessageDIV();" style="position: absolute;
		bottom: 0px; left: 0px; right: 0px; display: none; visibility: hidden; z-index: 1">
		<div id="divErrorMessages_Inner" style="background-color: white; text-align: left;
			position: relative; margin: 0px; padding: 5px; border: 1px solid; font-size: 8pt;
			color: black; font-family: Verdana;">
			<img id="imgErrorMessages_Min" src="Images/downarrows_white.gif" alt="Hide messages"
				style="right: 1px; position: absolute; top: 0px;" onclick="showErrorMessages(false);" />
			<igmisc:WebAsyncRefreshPanel id="pnlErrorMessages" runat="server" style="position: relative;"
				width="100%" height="100%">
				<asp:Label ID="lblErrors" runat="server" Text=""></asp:Label>
				<asp:BulletedList ID="bulletErrors" runat="server" Style="margin-top: 0px; margin-bottom: 0px;
					padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" Font-Names="Verdana"
					Font-Size="8pt" BorderStyle="None">
				</asp:BulletedList>
				<asp:Label ID="lblWarnings" runat="server" Text=""></asp:Label>
				<asp:BulletedList ID="bulletWarnings" runat="server" Style="margin-top: 0px; margin-bottom: 0px;
					padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" Font-Names="Verdana"
					Font-Size="8pt" BorderStyle="None">
				</asp:BulletedList>
				<asp:Label ID="lblWarningsPrompt_1" runat="server" Text="Click"></asp:Label>
				<span id="spnClickHere" name="spnClickHere" tabindex="1" style="color:#333366;" onclick="overrideWarningsAndSubmit();" onmouseover="try{this.style.color='#ff9608'}catch(e){}"
					onmouseout="try{this.style.color='#333366';}catch(e){}" onfocus="try{this.style.color='#ff9608';}catch(e){}"
					onblur="try{this.style.color='#333366';}catch(e){}" onkeypress="try{if(window.event.keyCode == 32){spnClickHere.click()};}catch(e){}">
					<asp:Label ID="lblWarningsPrompt_2" runat="server" Text="here" Font-Underline="true" 
						style="cursor: hand;"></asp:Label>
				</span>
				<asp:Label ID="lblWarningsPrompt_3" runat="server" Text=""></asp:Label>
			</igmisc:WebAsyncRefreshPanel>
		</div>
	</div>
	<!--
    Submission and Exceptional Errors Popup 
    -->
	<div id="divSubmissionMessages" style="position: absolute; left: 0px; top: 15%; width: 100%;
		display: none; z-index: 3; visibility: hidden; text-align: center;" nowrap="nowrap">
		<iframe id="ifrmMessages" src="" frameborder="0" scrolling="no"></iframe>
	</div>
	<!--
    File Upload Popup
    -->
	<div id="divFileUpload" style="position: absolute; left: 0px; top: 15%; width: 100%;
		filter: revealTrans(duration=0.5, transition=12); display: none; z-index: 3; visibility: hidden;
		text-align: center;" nowrap="nowrap" onfilterchange="return unblockFileUploadDIV();">
		<iframe id="ifrmFileUpload" src="" frameborder="0" scrolling="no"></iframe>
	</div>
	<!--
        Web Form Controls
        -->
	<div id="divInput" style="z-index: 0; width: 100%; background-color: <%=ColourThemeHex()%>;
		padding: 0px; margin: 0px; text-align: center">
		<igmisc:WebAsyncRefreshPanel ID="pnlInput" runat="server" Style="position: relative;
			padding-right: 0px; padding-left: 0px; padding-bottom: 0px; margin-top: 0px; margin-bottom: 0px;
			margin-right: auto; margin-left: auto; padding-top: 0px;" LinkedRefreshControlID="pnlErrorMessages">
			<igtxt:WebImageButton id="btnSubmit" runat="server" style="visibility: hidden; top: 0px;
				position: absolute; left: 0px; width: 0px; height: 0px;" text="">
			</igtxt:WebImageButton>
			<igtxt:WebImageButton id="btnReEnableControls" runat="server" style="visibility: hidden;
				top: 0px; position: absolute; left: 0px; width: 0px; height: 0px;" text="">
			</igtxt:WebImageButton>
			<asp:HiddenField ID="hdnCount_Errors" runat="server" Value="" />
			<asp:HiddenField ID="hdnCount_Warnings" runat="server" Value="" />
			<asp:HiddenField ID="hdnOverrideWarnings" runat="server" Value="0" />
			<asp:HiddenField ID="hdnLastButtonClicked" runat="server" Value="" />
			<asp:HiddenField ID="hdnNoSubmissionMessage" runat="server" Value="0" />
			<asp:HiddenField ID="hdnFollowOnForms" runat="server" Value="" />
			<asp:HiddenField ID="hdnErrorMessage" runat="server" Value="" />
			<asp:HiddenField ID="hdnSiblingForms" runat="server" Value="" />
			<asp:HiddenField ID="hdnSubmissionMessage_1" runat="server" Value="" />
			<asp:HiddenField ID="hdnSubmissionMessage_2" runat="server" Value="" />
			<asp:HiddenField ID="hdnSubmissionMessage_3" runat="server" Value="" />
		</igmisc:WebAsyncRefreshPanel>
	</div>
	<!--
    Temporary values from the server
    -->
	<asp:HiddenField ID="hdnFormHeight" runat="server" Value="0" />
	<asp:HiddenField ID="hdnFormWidth" runat="server" Value="0" />
	<asp:HiddenField ID="hdnFormBackColourHex" runat="server" Value="" />
	<asp:HiddenField ID="hdnFormBackImage" runat="server" Value="" />
	<asp:HiddenField ID="hdnFormBackRepeat" runat="server" Value="" />
	<asp:HiddenField ID="hdnFormBackPosition" runat="server" Value="" />
	<asp:HiddenField ID="hdnColourThemeHex" runat="server" Value="" />
	<asp:HiddenField ID="hdnFirstControl" runat="server" Value="" />
	</form>
	<!--
    Temporary client-side values
    -->
	<input type="hidden" id="txtPostbackMode" name="txtPostbackMode" value="0" />
	<input type="hidden" id="txtActiveElement" name="txtActiveElement" value="" />
	<input type="hidden" id="txtLastDate_Month" name="txtLastDate_Month" value="" />
	<input type="hidden" id="txtLastDate_Day" name="txtLastDate_Day" value="" />
	<input type="hidden" id="txtLastDate_Year" name="txtLastDate_Year" value="" />
</body>
</html>
