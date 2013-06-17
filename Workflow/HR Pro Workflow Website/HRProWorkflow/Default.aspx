<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" EnableSessionState="True" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml" id="htmMain">

<meta http-equiv="refresh" content="<%=Session("TimeoutSecs")%>;URL=timeout.aspx" />
<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=1;"/>
<!--<meta name="viewport" content="width=700; user-scalable=1;"/>-->


<head runat="server">
  <style type="text/css">
		.highlighted { background: yellow; }
	</style>

	  <title></title>	  
    <script src="Scripts/resizable-table.js" type="text/javascript"></script>	 
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
            // For postback, set up the scripts for begin and end requests...
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

      //Set the current page tab to page 1
      iCurrentTab = 1;
          
			try {
				iDefHeight = $get("frmMain").hdnFormHeight.value;
				iDefWidth = $get("frmMain").hdnFormWidth.value;
				$get("bdyMain").bgColor = $get("frmMain").hdnColourThemeHex.value;
				
				window.focus();

				if ((iDefHeight > 0) && (iDefWidth > 0)) {
					iResizeByHeight = iDefHeight - document.documentElement.clientHeight;
			
					iResizeByWidth = iDefWidth - document.documentElement.clientWidth;
					window.parent.moveTo((screen.availWidth - iDefWidth) / 2, (screen.availHeight - iDefHeight) / 3);
					window.parent.resizeBy(iResizeByWidth, iResizeByHeight);
				}
				

				try {
					if ($get("frmMain").hdnFirstControl.value.length > 0) {
					    sControlType = $get("frmMain").hdnFirstControl.value.substr($get("frmMain").hdnFirstControl.value.indexOf("_")+1);
                        sControlType = sControlType.substr(sControlType.indexOf("_")+1);
                        sControlType = sControlType.substring(0, sControlType.indexOf("_"));

                        if (sControlType == 7)
                        {
                            // Date (7)
                            igdrp_getComboById($get("frmMain").hdnFirstControl.value).focus();
                        }
                        else
                        {
                            if ((sControlType == 13)
                                || (sControlType == 14))
                            {
                                igcmbo_getComboById($get("frmMain").hdnFirstControl.value).focus();
                            }
                            else
                            {
                                if (sControlType == 11)
                                {
                                    // Record Selector (11)
                                    var grid = igtbl_getGridById($get("frmMain").hdnFirstControl.value);
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
						            document.getElementById($get("frmMain").hdnFirstControl.value).setActive();
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


				launchForms($get("frmMain").hdnSiblingForms.value, false);
			}
			catch (e) {}
		}

		function resizeToFit(piWidth, piHeight) {
			var iDefHeight;
			var iDefWidth;
			var iResizeByHeight;
			var iResizeByWidth;

			try {
				iDefHeight = $get("frmMain").hdnFormHeight.value;
				iDefWidth = $get("frmMain").hdnFormWidth.value;

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
				
		if($get("txtPostbackMode").value=="3") {
		  try {
		    if($get("txtActiveDDE").value.indexOf("dde")>0) {
		      //keep the lookup open.
		      //kicks off InitializeLookup BTW.
		      $find($get("txtActiveDDE").value).show();
		    }
		  }
		  catch (e) {}
		  return;			
		}
		//document.all.pleasewaitScreen.style.pixelTop = (document.body.scrollTop + 50);
		$get("pleasewaitScreen").style.visibility="visible";
		
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
		
			if ((($get("frmMain").hdnCount_Errors.value > 0)			
				|| ($get("frmMain").hdnCount_Warnings.value > 0))
				&& (pfDisplay == false)) {
				$get("imgErrorMessages_Max").style.display = "block";
				$get("imgErrorMessages_Max").style.visibility = "visible";
			}
			else {
				$get("imgErrorMessages_Max").style.display = "none";
				$get("imgErrorMessages_Max").style.visibility = "hidden";
			}


			// Removed all transitions in the following block as they're IE only

			if (pfDisplay == true) {
			  //refresh the errors WARP panel. 
			  __doPostBack('pnlErrorMessages', '');
				//divErrorMessages_Inner.style.visibility = "visible";
				
				//divErrorMessages_Outer.style.filter = "revealTrans(duration=0.3, transition=4)";
				//divErrorMessages_Outer.filters.revealTrans.apply();
				$get("divErrorMessages_Outer").style.display = "block";
				$get("divErrorMessages_Outer").style.visibility = "visible";
				//divErrorMessages_Outer.filters.revealTrans.play();
			}
			else {
				//divErrorMessages_Outer.style.filter = "revealTrans(duration=0.3, transition=5)";
				//divErrorMessages_Outer.filters.revealTrans.apply();
				$get("divErrorMessages_Outer").style.visibility = "hidden";
				//divErrorMessages_Outer.style.display = "none";
				//divErrorMessages_Inner.style.visibility = "hidden";
				//divErrorMessages_Outer.filters.revealTrans.play();
			}
		}

		function launchFollowOnForms(psForms) {
			launchForms(psForms, true);
		}

		function overrideWarningsAndSubmit() {
			if (divErrorMessages_Outer.disabled == true) {
				return;
			};

			$get("frmMain").hdnOverrideWarnings.value = 1;

			try {
				document.getElementById(frmMain.hdnLastButtonClicked.value).click();
			}
			catch (e) {
				$get("frmMain").btnSubmit.click();
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
			setPostbackMode(3);
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
					//txtLastDate_Year.value = dtCurrentDate.getYear();					
					txtLastDate_Year.value = dtCurrentDate.getFullYear();					
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

					$get("ifrmFileUpload").src = "FileUpload.aspx?" + sAlreadyUploaded + psElementItemID;

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
					
					$get("frmMain").btnReEnableControls.click();

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
		
		$get("pleasewaitScreen").style.visibility="hidden";
		
		//Reapply resizable column functionality to tables
		//This is put here to ensure functionality is reapplied after partial/full postback.
		ResizableColumns();
		
    if($get("txtActiveDDE").value.indexOf("dde")>0) {
      try {  
        $find($get("txtActiveDDE").value).show();        
        $get("txtActiveDDE").value="";        
      }
      catch (e) {}      
    }		
		    
		if($get("txtPostbackMode").value==3) {
		    //ShowMessage is the sub called in lieu of Application:EndRequest, i.e. Pretty much the end of
		    //the postback cycle. So we'll reset all grid scroll bars to their previous position
		    SetScrollTopPos("", "-1", 0);		    
      }
      
      
			try {
				if ($get("frmMain").hdnErrorMessage.value.length > 0) {
					showSubmissionMessage();
					return;
				}

				if($get("txtPostbackMode").value!="2") refreshLiterals();

				if (($get("txtPostbackMode").value == 2)
                    || ($get("txtPostbackMode").value == 3)) 
                {
					// 0 = Default
					// 1 = Submit/SaveForLater button postback (ie. WebForm submission)
					// 2 = Grid header postback
					// 3 = FileUpload button postback
					
					if ($get("txtPostbackMode").value == 3) 
					{
					    $get("ifrmFileUpload").contentWindow.enableControls();
          }
          // not doing this causes the object referenced is null error:
					setPostbackMode(0);
					return;
					
				}

				if (($get("frmMain").hdnCount_Errors.value > 0)
			        || ($get("frmMain").hdnCount_Warnings.value > 0)) {
					showErrorMessages(true);
				}
				else {
					if ($get("frmMain").hdnNoSubmissionMessage.value == 1) {
						try {
							if ($get("frmMain").hdnFollowOnForms.value.length > 0) {
								launchFollowOnForms($get("frmMain").hdnFollowOnForms.value);
							}
							else {							
							  if(navigator.userAgent.indexOf("MSIE")>0) {
							    //Only IE can self-close windows that it didn't open
								  window.close();
								}
								else
								{
								  // Non-IE browsers can't self-close windows.
								  //show Please Wait box, with 'please close me' text
								  disableChildElements("pnlInput");
								  $get("pleasewaitScreen").style.visibility="visible";
								  $get("pleasewaitScreen").style.width="200px";

                  labelCtl = document.getElementById("pleasewaitText");
                  if (null != labelCtl) {
                    labelCtl.innerHTML = "Workflow completed.<BR/><BR/>Please close your browser.";                    
                  }								  
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
		  //Rem out all transitions as only IE can run 'em
			try {
				$get("ifrmMessages").src = "SubmissionMessage.aspx";

				$get("divInput").disabled = true;
				$get("frmMain").hdnCount_Errors.value = 0;
				$get("frmMain").hdnCount_Warnings.value = 0;
				$get("divErrorMessages_Outer").style.display = "none";
				showErrorMessages(false);
				//$get("divSubmissionMessages").style.filter = "revealTrans(duration=0.5, transition=12)";
				//$get("divSubmissionMessages").filters.revealTrans.apply();
				$get("divSubmissionMessages").style.display = "block";
				$get("divSubmissionMessages").style.visibility = "visible";
				//$get("divSubmissionMessages").filters.revealTrans.play();
			}
			catch (e) { }
		}

		function unblockFileUploadDIV() {
			try {
				if (($get("divFileUpload").style.visibility == "hidden") &&
					($get("divFileUpload").style.display != "none")) {
					$get("divFileUpload").style.display = "none";
				}
			}
			catch (e) { }

			try {
				document.getElementById($get("txtActiveElement").value).setActive();
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

// Old resizeformforcombo - used with Infragistics only, now replaced.	    
//	    function ResizeFormForCombo(psWebComboId) {
//			var iResizeByHeight = 0;
//			var iResizeByWidth = 0;
//            
//			try {
//	            var objCombo = igcmbo_getComboById(psWebComboId);
//                var grid = objCombo.getGrid();

//                var oEl = grid.Element;

//                if (oEl.scrollWidth > bdyMain.clientWidth)
//                {
//                    if (oEl.scrollWidth > screen.availWidth)
//                    {
//                        iResizeByWidth = screen.availWidth - bdyMain.clientWidth;
//                    }
//                    else
//                    {
//                        iResizeByWidth = oEl.scrollWidth - bdyMain.clientWidth;
//                    }
//                }
//                
////                if (oEl.scrollHeight > bdyMain.clientHeight)
////                {
////                    if (oEl.scrollHeight > screen.availHeight)
////                    {
////                        iResizeByHeight = screen.availHeight - bdyMain.clientHeight;
////                    }
////                    else
////                    {
////                        iResizeByHeight = oEl.scrollHeight - bdyMain.clientHeight;
////                    }
////                }
//                
//                if ((iResizeByWidth > 0) || (iResizeByHeight > 0))
//                {
//                    setTimeout('window.resizeBy(' + iResizeByWidth + ',' + iResizeByHeight + ');', 100);
//                }
//            }
//            catch(e) {}
//	    }

function ResizeComboForForm(sender, args) {
      psWebComboID = sender._id;
            
			var iResizeByHeight = 0;
			var iResizeByWidth = 0;

      //Let's set the width of the lookup panel to the width of the screen. 
      //It used to resize the screen, but don't want this happening now.

			try {			
                var oEl = document.getElementById(psWebComboID.replace("dde", ""));
                if(eval(oEl)) 
                {
                  if (oEl.offsetWidth > $get("bdyMain").clientWidth)
                  {
                    iNewWidth = $get("bdyMain").clientWidth - oEl.offsetLeft - 5 + "px";
                    
                    oEl.style.width = iNewWidth;
                    document.getElementById(psWebComboID.replace("dde", "gridcontainer")).style.width = oEl.style.width;
                  }   
                  
                  //also set left position to 0 if required (right coord > bymain.width)
                  //alert(oEl.offsetLeft + oEl.offsetWidth + ":" + bdyMain.clientWidth);
                  if ((oEl.offsetLeft + oEl.offsetWidth) > $get("bdyMain").clientWidth)
                  {
                    oEl.style.left = "0px";
                  }                                                 
                }
            }
      catch(e) {}

}



  function scrollHeader(iGridID) {
      //keeps the header table aligned with the gridview in record
      //selectors and lookups.
      var leftPos = document.getElementById(iGridID).scrollLeft;
      document.getElementById(iGridID.replace("gridcontainer", "Header")).style.left = "-" + leftPos + "px";
      
      var hdn1 = document.getElementById(iGridID.replace("Grid","scrollpos"));
      hdn1.value = document.getElementById(iGridID).scrollTop;
      
  }
	    
  function InitializeLookup(sender, args) {
  
  if($get("txtActiveDDE").value.indexOf("dde")>=0) {
    // If we're in the process of displaying a filtered lookup already, do nothing and exit the function...
    return;
  }

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
	        
	        var sID = "lookup" + psWebComboID.replace("dde","");
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
                                
					                //var objCombo = igcmbo_getComboById(psWebComboId);
	                        //        objCombo.selectWhere(sSelectWhere);
        	                                         
                          if(sValue=="") {
                            document.getElementById(psWebComboID.replace("dde", "filterSQL")).value = "";                          
                          }
                          else {
                            document.getElementById(psWebComboID.replace("dde", "filterSQL")).value = sSelectWhere;                          
                          }
                          
                          //This prevents the lookup closing after the filter is applied/removed
                          
                          $get("txtActiveDDE").value = psWebComboID;
                          
                          setPostbackMode(3);
                          
                          //These lines hide the lookup dropdown until it's filled with data.
                          document.getElementById(psWebComboID.replace("dde","")).style.height="0px";
                          document.getElementById(psWebComboID.replace("dde","")).style.width="0px";
                          
                          //This clicks the server-side button to apply filtering...                          
                          //this also kicks off the gosubmit() via postback beginrequest.                          
                          document.getElementById(psWebComboID.replace("dde", "refresh")).click();
                          
                          //set pbmode back to 0 to prevent recursion.                          
                          setPostbackMode(0);                                                                  
                        }
	                }
                }
            }
           catch (e) {}

	        return false;
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

  function isGridFiltered(iGridID) { 
    //searches the specified table for hidden rows and returns true if any are found...
    var table = document.getElementById(iGridID)
    
    for (var r = 0; r < table.rows.length; r++) {
        if (table.rows[r].style.display == 'none') {
          return true;
        }
    }
    return false;  
  }
  
  function GetGridRowHeight(iGridID) {
    var table = document.getElementById(iGridID)


    for (var r = 0; r < table.rows.length; r++) {
        if (table.rows[r].style.display == '') {
          return document.getElementById(iGridID.replace("Grid", "Grid_row" + r)).offsetHeight;
        }
    }
    
    return 0;    
  }
  
  
  function SetScrollTopPos(iGridID, iPos, iRowIndex) {
    if(iPos==-1) {
    // -1 is the 'code' to reset scrollbar to stored position
    //Loop through all hidden scroll fields and reset values.
    var controlCollection = $get("frmMain").elements;
      if (controlCollection!=null) 
      {
	      for (i=0; i<controlCollection.length; i++)  
	      {
		      if(Right(controlCollection.item(i).name, 9)=="scrollpos") {			    
		        document.getElementById(controlCollection.item(i).name.replace("scrollpos", "gridcontainer")).scrollTop = (controlCollection.item(i).value);
  			  }	
	      }
      }							
    }
    else { 
      //Check if this grid is quick-filtered (NOT lookup filtered)
      //If it is, calculate the scroll position to use after postback,
      //otherwise store the current scroll position for postback...
      if(isGridFiltered(iGridID)) {
        iPos = (iRowIndex * GetGridRowHeight(iGridID)) - 1;
        }
      //store the scrollbar position
      hdn1 = document.getElementById(iGridID.replace("Grid","scrollpos"));
      hdn1.value = iPos;
      ScrollTopPos = iPos;          
    }
  }
  
  function SetCurrentTab(iNewTab) {
    var currentTab = "forminput_" + iCurrentTab + "_21_PageTab"
    var newTab = "forminput_" + iNewTab + "_21_PageTab"

    try {
      $get(currentTab).style.display = "none";
      $get(newTab).style.display = "block";

      iCurrentTab = iNewTab;
    }
    catch (e) {}
  }

    // ]]>
	</script>

	<script src="scripts\WebNumericEditValidation.js" type="text/javascript"></script>
	
  <ajx:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server" EnablePartialRendering="true" EnablePageMethods="true">
  </ajx:ToolkitScriptManager>
	<!--
        Web Form Validation Error Messages
        -->        

<div id="pleasewaitScreen" style="position:absolute;z-index:5;top:30%;width:150px;height:60px;left:50%;margin-left:-75px;visibility:hidden">
			<table border="0" cellspacing="0" cellpadding="10" style="top: 0px; left: 0px; width: 100%;
        height: 100%; position: relative; text-align: center; font-size: 10pt; color: black;
        font-family: Verdana; border: black 1px solid;" bgcolor="White">
				<tr>
					<td style="width:100%;height:100%;background-color:White;text-align:center;vertical-align:middle">
								<label id="pleasewaitText">Processing...<br/><br/>Please wait.<br/></label>
					</td>
				</tr>
			</table>
		</div>
		
		       
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
						style="cursor: pointer;"></asp:Label>
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
  <div id="pnlTabsDiv" style="height:auto;float:right" runat="server"></div>
	<div id="divInput" style="top:0px; left:0px; z-index: 0; background-color: <%=ColourThemeHex()%>;
		padding: 0px; margin: 0px; text-align: center;float:right" runat="server">
      
    <asp:UpdatePanel ID="pnlInput" runat="server">
    <ContentTemplate>
    <div id = "pnlInputDiv" runat="server" style="position:relative;padding-right:0px;padding-left:0px;padding-bottom:0px;
                            margin-top:0px;margin-bottom:0px;margin-right:auto;margin-left:auto;padding-top:0px;">

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
	<input type="hidden" id="txtActiveDDE" name="txtActiveDDE" value="" />	
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
