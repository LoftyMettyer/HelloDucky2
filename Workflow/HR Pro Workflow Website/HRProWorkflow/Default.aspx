<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" EnableSessionState="True" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<meta http-equiv="refresh" content="<%=Session("TimeoutSecs")%>;URL=timeout.aspx" />

<html xmlns="http://www.w3.org/1999/xhtml" id="htmMain">
<head runat="server">
	  <title></title>	  
    <script src="Scripts/resizable-table.js" type="text/javascript"></script>
    <script type="text/javascript">
//      var DDE;
//      function pageLoad() {
//        DDE = $find('forminput_38886_14_DDE');
//        if (DDE._dropDownControl) {
//          $common.removeHandlers(DDE._dropDownControl, DDE._dropDownControl$delegates);
//        }
//        DDE._dropDownControl$delegates = {
//          click: Function.createDelegate(DDE, ShowMe),
//          contextmenu: Function.createDelegate(DDE, DDE._dropDownControl_oncontextmenu)
//        }
//        $addHandlers(DDE._dropDownControl, DDE._dropDownControl$delegates);
//      }

//      function ShowMe() {
//        DDE._wasClicked = true;
//      }

    </script>
    
    
</head>

<body id="bdyMain" onload="return window_onload()" scroll="auto" style="overflow: auto;  
	text-align: center; margin: 0px; padding: 0px;">
	<img id="imgErrorMessages_Max" src="Images/uparrows_white.gif" alt="Show messages"
		style="position: absolute; right: 1px; bottom: 1px; display: none; visibility: hidden;
		z-index: 1;" onclick="showErrorMessages(true);" />
	<form runat="server" hidefocus="true" id="frmMain" onsubmit="return submitForm();">
	
		<script language="javascript" type="text/javascript">
  var app = Sys.Application
  app.add_init(ApplicationInit);
  
    // <!CDATA[
    var gridViewCtl = null;
    var curSelRow = new Array();
    var selRow = new Array();
    var curSelRowBackColour = new Array();
 
    function ApplicationInit(sender) {
        try 
        {
            var prm = Sys.WebForms.PageRequestManager.getInstance();
            if (!prm.get_isInAsyncPostBack()) 
            {
              prm.add_beginRequest(goSubmit);
              prm.add_endRequest(showMessage);
            }
        }
        catch (e) {}
    }


		function window_onload() {
			var iDefHeight;
			var iDefWidth;
			var iResizeByHeight;
			var iResizeByWidth;
      var sControlType;
      var oldgridSelectedColor;
			var ScrollTopPos;
           
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
		if(txtPostbackMode.value=="2") return;		
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
			  //refresh the errors WARP panel. 
			  __doPostBack('pnlErrorMessages', '');
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
		  pbModeValue = document.getElementById("txtPostbackMode").value
			
			try {
				if (pbModeValue == 0) {
				  tAE = document.getElementById("txtActiveElement");				  
				  if(eval(tAE)) {tae.value.setActive();}
				  
				}
			}
			catch (e) { };
			
			return (pbModeValue != 0);
		}

		function setPostbackMode(piValue) {
			// 0 = Default
			// 1 = Submit/SaveForLater button postback (ie. WebForm submission)
			// 2 = Grid header postback
			// 3 = FileUpload button postback
      try {
        pbModeValue = document.getElementById("txtPostbackMode")
	      pbModeValue.value = piValue;
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
		
		//ShowMessage is the sub called in lieu of Application:EndRequest, i.e. Pretty much the end of
		//the postback cycle. So we'll reset all grid scroll bars to their previous position
		SetScrollTopPos("", "-1");
        
			try {
				if (frmMain.hdnErrorMessage.value.length > 0) {
					showSubmissionMessage();
					return;
				}

				if(txtPostbackMode.value!="2") refreshLiterals();

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
          // not doing this causes the object referenced is null error:
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

  function scrollHeader(iGridID) {
      //keeps the header table aligned with the gridview in record
      //selectors and lookups.
      var leftPos = document.getElementById(iGridID).scrollLeft;
      document.getElementById(iGridID.replace("gridcontainer", "Header")).style.left = "-" + leftPos + "px";
      
      var hdn1 = document.getElementById(iGridID.replace("Grid","hiddenfield"));
      hdn1.value = document.getElementById(iGridID).scrollTop;
      
  }
	    
  function InitializeLookup(sender, args) {
  
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
	        var psWebComboID = "";
	        
	        psWebComboID = sender._id
	        
	        if(psWebComboID=="") {return;}
	        
	        var sID = "lookup" + psWebComboID.replace("DDE","");
	        
		      try {
			    
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
                              if (sControlType == 13) {  
                                var ctlLookupValueCombo = document.getElementById(sValueID);
                        	      sValue = ctlLookupValueCombo.value;
                        	    }
                        	    else
                        	    {
                                var ctlLookupValueCombo = document.getElementById(sValueID + "TextBox");
                        	      sValue = ctlLookupValueCombo.value;                        	    
                        	    }
                        	    
                        	    
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
                                                     
                          // ASP's own gridview control doesn't have a nice built-in filter option,
                          // so we'll loop through the rows and hide any that don't meet the criteria.
                          // Need Filter column, Operator and Value.
					                //var objCombo = igcmbo_getComboById(psWebComboId);
	                        //        objCombo.selectWhere(sSelectWhere);
        	                
        	                 //PageMethods.SetGridFilter(psWebComboID, sSelectWhere, OnCallSumComplete, OnCallSumError);
        	                 
                          hideRows(psWebComboID.replace("DDE", "Grid"), sValue, sSelectWhere);
	                
                        }
	                }
                }
            }
           catch (e) {}

	        return false;
  }


function OnCallSumComplete(result,txtresult,methodName)
{
//Show the result in txtresult
}

// Callback function on error
// Callback function on complete
// First argument is always "error" if server side code throws any exception
// Second argument is usercontext control pass at the time of call
// Third argument is methodName (server side function name) 
// In this example the methodName will be "Sum"
function OnCallSumError(error,userContext,methodName)
{
//alert(error.get_message());
}

 function hideRows(gridName, searchValue, sSelectWhere)
  {
    //Get the column number
    tblTable = document.getElementById(gridName);   
    iFilterColumn = tblTable.attributes["LookupFilterColumn"].value;    
    rows = document.getElementById(gridName).rows;
    //If only the blank row exists...
    if(rows.length==1){return;}
    iVisibleRows = 1;
      
    for (i = 1; i < rows.length; i++) {
      if ((rows[i].cells[iFilterColumn].innerText == searchValue) || (searchValue == ""))  {
        rows[i].style.display = "block";
        iVisibleRows += 1;
      }
      else {
        rows[i].style.display = "none";
      }
    }

    //Set the max height of the dropdown again.
    iRowHeight = (document.getElementById(gridName.replace("Grid", "TextBox")).clientHeight) - 6;
    iRowHeight = (iRowHeight< 22)?22:iRowHeight;

    iDropHeight = (iRowHeight * ((iVisibleRows>6)?6:iVisibleRows) + 1);
    
    grdContainer = document.getElementById(gridName.replace("Grid", "gridcontainer"));
    grdContainer.height = iDropHeight + 'px';
    grdContainer.style.height = iDropHeight + 'px';
    
    grdContainer = document.getElementById(gridName.replace("Grid", ""));
    grdContainer.height = iDropHeight + 16 + 'px';
    grdContainer.style.height = iDropHeight + 16 + 'px';

  }

function getSelectedRow(iGridID, iRowIdx) {
    getGridViewControl(iGridID);
    if (null != gridViewCtl) {
    
      iIDCol = GetIDColumnNum(iGridID);
      for (i = 0; i < gridViewCtl.rows.length; i++) {
        try {
          if (gridViewCtl.rows[i].cells[iIDCol].innerText == iRowIdx)  {
            return gridViewCtl.rows[i];      
          }
        }
        catch (e) {}
      }                
    }
    return null;
  }

  function getGridViewControl(iGridID) {
    //    if (null == gridViewCtl) {
      gridViewCtl = document.getElementById(iGridID);
    //}
  }


function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

  function SetScrollTopPos(iGridID, iPos) {
    if(iPos==-1) {
    
    //Loop through all hidden scroll fields and reset values.
    var controlCollection = frmMain.elements;
	    if (controlCollection!=null) 
	    {
		    for (i=0; i<controlCollection.length; i++)  
		    {
			    if(Right(controlCollection.item(i).name, 11)=="hiddenfield") {
			    
			      document.getElementById(controlCollection.item(i).name.replace("hiddenfield", "gridcontainer")).scrollTop = (controlCollection.item(i).value);
    			}	
		    }
	    }
				
				
      // -1 is the code to reset the scrollbar
      //alert("SETTING SCROLLBAR " + iGridID.replace("Grid", "gridcontainer") + " TO " + ScrollTopPos);
      //document.getElementById(iGridID.replace("Grid", "gridcontainer")).scrollTop = (ScrollTopPos);
    }
    else { 
      //store the scrollbar position
      hdn1 = document.getElementById(iGridID.replace("Grid","hiddenfield"));
      hdn1.value = iPos;
      ScrollTopPos = iPos;          
    }
  }

  function changeRow(iGridID, iRowIdx, strHighlightCol, iIDCol) {
    //e.g. changeRow('forminput_38880_11_Grid', '0', '#FDEB9F', '7');
    
    //change the row colour and reset any previously selected rows.
    
    var iElementID = iGridID.substring(10, iGridID.indexOf("_", 10));
   
    selRow[iElementID] = getSelectedRow(iGridID, iRowIdx);
    
    if (curSelRow[iElementID] != null) {
      if (strHighlightCol == "default") {
        strHighlightCol = curSelRow[iElementID].style.backgroundColor;
      }
      curSelRow[iElementID].style.backgroundColor = curSelRowBackColour[iElementID];
    }
    if (null != selRow[iElementID]) {
      curSelRowBackColour[iElementID] = selRow[iElementID].style.backgroundColor;  //oldgridSelectedColor;  Switch this to enable row highlight on hoverover.
      curSelRow[iElementID] = selRow[iElementID];
      curSelRow[iElementID].style.backgroundColor = strHighlightCol;
    }

    //The following doesn't work in Firefox. This will be referred to in future
    //Get the record ID from the selected row and store to hidden element
    //tblTable=document.getElementById(iGridID);
    //Cell = tblTable.rows[iRowIdx].cells[iIDCol];
    //hdn1 = document.getElementById(iGridID.replace("Grid","hiddenfield"));
    //hdn1.value = Cell.innerText;
    //Firefox compliant version of the above:
    var theCells = selRow[iElementID].getElementsByTagName("td"); //Header would be th
    var theText = theCells[iIDCol].innerHTML
    hdn1 = document.getElementById(iGridID.replace("Grid","hiddenfield"));
    hdn1.value = theText;
    
    //Might need this someday too, converts innerHTML to innerText.
    //if(typeof HTMLElement!="undefined"){
    //  HTMLElement.prototype.__defineGetter__("innerText", function () { 
    //  var r = this.ownerDocument.createRange(); 
    //  r.selectNodeContents(this); 
    //  return r.toString(); 
    //  }); 
    //}
    //alert(document.getElementById("div1").innerHTML);
    //alert(document.getElementById("div1").innerText);
  }

  
  function changeDDERow(iGridID, iRowIdx, strHighlightCol, iIDCol) {
    
    //e.g. changeDDERow('forminput_38880_11_Grid', '0', '#FDEB9F', '7');
    
    //Dropdown highlight colour is fixed at system highlight...
    strHighlightCol = '#FDEB9F'
    
    //change the row colour and reset any previously selected rows.      
    var iElementID = iGridID.substring(10, iGridID.indexOf("_", 10));
    
    selRow[iElementID] = getSelectedRow(iGridID, iRowIdx);
    
    if (curSelRow[iElementID] != null) {
      if (strHighlightCol == "default") {
        strHighlightCol = curSelRow[iElementID].style.backgroundColor;
      }
      curSelRow[iElementID].style.backgroundColor = curSelRowBackColour[iElementID];
    }
    if (null != selRow[iElementID]) {
      curSelRowBackColour[iElementID] = selRow[iElementID].style.backgroundColor;  //oldgridSelectedColor;  Switch this to enable row highlight on hoverover.
      curSelRow[iElementID] = selRow[iElementID];
      curSelRow[iElementID].style.backgroundColor = strHighlightCol;
    }
    
    //Set the textbox text to the selected grid item
    tblTable = document.getElementById(iGridID);
    iLookupColumnIndex = tblTable.attributes["LookupColumnIndex"].value;
        
    if(IsNumeric(iLookupColumnIndex)) {
      Cell = tblTable.rows[iRowIdx].cells[iLookupColumnIndex];
    }
    else{
      Cell = tblTable.rows[iRowIdx].cells[0];
    }
      txtTextBox = document.getElementById(iGridID.replace("Grid", "TextBox"));
      txtTextBox.value = Cell.innerHTML.replace("&nbsp;","");
  }
  
  
  //Sort the grid/lookup when clicking on column headers.
  //Needs a bit of work to convert dates correctly for sorting.
  //and the lastcol/lastseq needs converting to an array.
  var lastcol, lastseq;
  function fsort(ao_table, ai_sortcol, ab_header) {
    var ir, ic, is, ii, id;
    
    ir = ao_table.rows.length;
    if (ir < 1) return;

    ic = ao_table.rows[1].cells.length;
    // if we have a header row, ignore the first row
    if (ab_header == true) is = 1; else is = 0;

    // take a copy of the data to shuffle in memory
    var row_data = new Array(ir);
    ii = 0;
    for (i = is; i < ir; i++) {
      var col_data = new Array(ic);
      for (j = 0; j < ic; j++) {
        col_data[j] = ao_table.rows[i].cells[j].innerHTML;
      }
      row_data[ii++] = col_data;
    }

    // sort the data
    var bswap = false;
    var row1, row2;
    var col1, col2;

    if (ai_sortcol != lastcol)
      lastseq = 'A';
    else {
      if (lastseq == 'A') lastseq = 'D'; else lastseq = 'A';
    }

    // if we have a header row we have one less row to sort
    if (ab_header == true) id = ir - 1; else id = ir;
    for (i = 0; i < id; i++) {
      bswap = false;
      for (j = 0; j < id - 1; j++) {
        // test the current value + the next and
        // swap if required.
        row1 = row_data[j];
        row2 = row_data[j + 1];

        if (IsNumeric(row1[ai_sortcol]) == true) {
          col1 = parseFloat(row1[ai_sortcol]);
          col2 = parseFloat(row2[ai_sortcol]);
        }
        else if (isDate(row1[ai_sortcol]) == true) {
          col1 = Date.parse(row1[ai_sortcol]);
          col2 = Date.parse(row2[ai_sortcol]);
        }
        else {
          col1 = row1[ai_sortcol];
          col2 = row2[ai_sortcol];
        }


        if (lastseq == "A") {
          if (col1 > col2) {
            row_data[j + 1] = row1;
            row_data[j] = row2;
            bswap = true;
          }
        }
        else {
          if (col1 < col2) {
            row_data[j + 1] = row1;
            row_data[j] = row2;
            bswap = true;
          }
        }
      }
      if (bswap == false) break;
    }

    // load the data back into the table
    // When we hit the ID of the selected row, store it so we can highlight and scroll to it.
    
    if(eval(document.getElementById(ao_table.id.replace("Grid", "TextBox")))) 
    {
      hdn1 = document.getElementById(ao_table.id.replace("Grid", "TextBox"));
      iRecordID = hdn1.value.toUpperCase().trim()
      iRecordIDColNum = ao_table.attributes["LookupColumnIndex"].value;
    }
    else
    {
      //get the selected ID from the hidden field
      hdn1 = document.getElementById(ao_table.id.replace("Grid", "hiddenfield"));
      iRecordID = hdn1.value;
      if((iRecordID < 0) || (IsNumeric(iRecordID)==false)) {iRecordID = 0;}
      iRecordIDColNum = GetIDColumnNum(ao_table.id);
    }

    ii = is;
    iCurrentRow = -1;
    
    for (i = 0; i < id; i++) {
      row1 = row_data[i];
      for (j = 0; j < ic; j++) {
        ao_table.rows[ii].cells[j].innerHTML = row1[j];
      }
      //check for ID match.
      if (iRecordIDColNum >= 0) {
        if (ao_table.rows[ii].cells[iRecordIDColNum].innerText.toUpperCase().trim() == iRecordID) {
          iCurrentRow = ii;
        }
      }
      
      ii++;
    }
    lastcol = ai_sortcol;
   
    if (iCurrentRow >= 0)   {
      doscroll(ao_table.id, iCurrentRow);
    }
  }

  function GetIDColumnNum(GridID) {

    ao_table = document.getElementById(GridID);
    
    ir = ao_table.rows.length;

    if (ir < 1) return;

    ic = ao_table.rows[1].cells.length;

    ii = 0;
      for (j = 0; j < ic; j++) {
        col_colour = ao_table.rows[1].cells[j].style.backgroundColor;
        if (col_colour == "black") {
          return j;
        }
      }
  }


  function IsNumeric(sText) {
    var ValidChars = "0123456789.-+ ";
    var IsNumber = true;
    var Char;

    for (i = 0; i < sText.length && IsNumber == true; i++) {
      Char = sText.charAt(i);
      if (ValidChars.indexOf(Char) == -1) {
        IsNumber = false;
      }
    }
    return IsNumber;
  }

  function isDate(value) {
    var d = Date.parse(value);
    return (d > 0);
  } 

  
  function doscroll(GridID, iRowNum) {
    //scrolls the grid/lookup to the specified row and highlights it.
    tblTable = document.getElementById(GridID);

    if (iRowNum <= tblTable.rows.length) {
      iRowHeight = tblTable.rows(iRowNum).offsetHeight;
      //scroll to the row
      document.getElementById(GridID.replace("Grid", "gridcontainer")).scrollTop = (iRowHeight * iRowNum);
      //Highlight the row
      if(eval(document.getElementById(GridID.replace("Grid", "TextBox"))))
      {
        changeDDERow(GridID, iRowNum, 'default', 0)
      }
      else
      {
        changeRow(GridID, iRowNum, 'default', GetIDColumnNum(GridID))
      }
    }
  }  
  
    // ]]>
	</script>

	<script src="scripts\WebNumericEditValidation.js" type="text/javascript"></script>
	
  <ajx:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server" EnablePartialRendering="true" EnablePageMethods="true">
  </ajx:ToolkitScriptManager>
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
		<%--<igmisc:WebAsyncRefreshPanel ID="pnlInput2" runat="server" Style="position: relative;
			padding-right: 0px; padding-left: 0px; padding-bottom: 0px; margin-top: 0px; margin-bottom: 0px;
			margin-right: auto; margin-left: auto; padding-top: 0px;" LinkedRefreshControlID="pnlErrorMessages">
		</igmisc:WebAsyncRefreshPanel>--%>
      
    <asp:UpdatePanel ID="pnlInput" runat="server"  >
    <ContentTemplate>
    <div id = "pnlInputDiv" runat="server" style="position:relative;padding-right:0px;padding-left:0px;padding-bottom:0px;
                            margin-top:0px;margin-bottom:0px;;margin-right:auto;margin-left:auto;padding-top:0px;">
    </div>    
      <asp:Button id="btnSubmit" runat="server" style="visibility: hidden; top: 0px;
				position: absolute; left: 0px; width: 0px; height: 0px;" text=""/>
        <asp:Button id="btnReEnableControls" runat="server" style="visibility: hidden;
				top: 0px; position: absolute; left: 0px; width: 0px; height: 0px;" text=""/>
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
		</ContentTemplate>
    </asp:UpdatePanel>			

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

<script language="javascript" type="text/javascript">

  function disposeTree(sender, args) {
  
    //http://support.microsoft.com/?kbid=2000262

    try {
  
    var elements = args.get_panelsUpdating();
    for (var i = elements.length - 1; i >= 0; i--) {
      var element = elements[i];
      var allnodes = element.getElementsByTagName('*'),
                length = allnodes.length;
      var nodes = new Array(length)
      for (var k = 0; k < length; k++) {
        nodes[k] = allnodes[k];
      }
      for (var j = 0, l = nodes.length; j < l; j++) {
        var node = nodes[j];
        if (node.nodeType === 1) {
          if (node.dispose && typeof (node.dispose) === "function") {
            node.dispose();
          }
          else if (node.control && typeof (node.control.dispose) === "function") {
            node.control.dispose();
          }

          var behaviors = node._behaviors;
          if (behaviors) {
            behaviors = Array.apply(null, behaviors);
            for (var k = behaviors.length - 1; k >= 0; k--) {
              behaviors[k].dispose();
            }
          }
        }
      }
      element.innerHTML = "";
    }} catch (e) { }
  }

  try {
    Sys.WebForms.PageRequestManager.getInstance().add_pageLoading(disposeTree);
  }
  catch (e) { }

</script>
</html>
