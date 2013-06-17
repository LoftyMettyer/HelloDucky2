﻿	
	    var app = Sys.Application;
	    app.add_init(ApplicationInit);

	    var formInputPrefix = "FI_";
 
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

      //fault HRPRO-2270
	    function resizeIframe(id, iNewHeight) {
        //Plus one for luck (IE9 actually)
	      iNewHeight = iNewHeight + 1;
	      document.getElementById(id).height = (iNewHeight) + "px";
	    }

	    function window_onload() {

	        var iDefHeight, iDefWidth, iResizeByHeight, iResizeByWidth;

	        //Set the current page tab	  
	        var iPageNo = document.getElementById("hdnDefaultPageNo").value;
	        
	        if(iPageNo > 0) {
	            window.iCurrentTab = iPageNo;
	        }
	        else {
	            window.iCurrentTab = 1;
	        }
	        SetCurrentTab(iCurrentTab);

	        window.iCurrentMessageState = 'none';

	        try {
	          
	            iDefHeight = window.$get("frmMain").hdnFormHeight.value;
	            iDefWidth = window.$get("frmMain").hdnFormWidth.value;
			    
	            window.focus();
	            if (iDefHeight > 0 && iDefWidth > 0) {
	              iResizeByHeight = iDefHeight - window.currentHeight; 
	              iResizeByWidth = iDefWidth - window.currentWidth;
	                
	              window.parent.resizeBy(iResizeByWidth, iResizeByHeight);	
	              window.parent.moveTo((screen.availWidth - iDefWidth) / 2, (screen.availHeight - iDefHeight) / 3);			  
	            }

	            try {
	                if (window.$get("frmMain").hdnFirstControl.value.length > 0) {
	                    document.getElementById(window.$get("frmMain").hdnFirstControl.value).focus();
	                }
	            }
	            catch (e) { }

	            launchForms(window.$get("frmMain").hdnSiblingForms.value, false);
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
	            try { newWin.window.focus(); } catch (e) { }
	        }
	        catch (e) {
	            try {
	                try {
	                    newWin.close();
	                }
	                catch(e) {
	                    alert("For your security please close your browser");
	                }
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

	        showWait(true);
	        showOverlay(true);
	    }

	    function getElementsBySearchValue(searchValue) {
	        var retVal = new Array();
	        var elems = document.getElementsByTagName("input");

	        for(var i = 0; i < elems.length; i++) {
	            var valueProp = "";
              
	            try {
	                var nameProp = elems[i].getAttribute('name');
	                if(nameProp.substr(0, 8)=="lookupFI")
	                    var valueProp = elems[i].getAttribute('value');
	            }
	            catch(e) {}              
              
	            if(!(valueProp==null)) {
	                if(valueProp.indexOf(searchValue) > 0) {
	                    retVal.push(elems[i]);
	                }         
	            }
	        }

	        return retVal;    
	    }

	    function showErrorMessages(state) {

	        switch (state) {
	            case 'max':
	                document.getElementById("errorMessagePanel").style.display = "block";
	                document.getElementById("errorMessageMax").style.display = "none";
	                break;
	            case 'min':
	                document.getElementById("errorMessagePanel").style.display = "none";
	                document.getElementById("errorMessageMax").style.display = "block";
	                break;
	            default:
	                document.getElementById("errorMessagePanel").style.display = "none";
	                document.getElementById("errorMessageMax").style.display = "none";
	        }
	        window.iCurrentMessageState = state;
	    }

	    function hasErrors() {
	        return document.getElementById("frmMain").hdnCount_Errors.value > 0 || 
    	           document.getElementById("frmMain").hdnCount_Warnings.value > 0;
        }

	    function launchFollowOnForms(psForms) {
	        launchForms(psForms, true);
	    }

	    function overrideWarningsAndSubmit() {

	        $get("frmMain").hdnOverrideWarnings.value = 1;

	        try {
	            document.getElementById($get("frmMain").hdnLastButtonClicked.value).click();
	        }
	        catch (e) {}
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

        function dateControlAndroidFix(controlId, hide) {
            var dateControl = document.getElementById(controlId);
            var nodes = dateControl.parentNode.childNodes;
            for(var i = 0; i < nodes.length; i++) {
                var ctl = nodes[i];
                if (ctl.id && ctl.id.indexOf('FI_') == 0 && ctl.id != dateControl.id) {
                    if (ctl.offsetTop > dateControl.offsetTop && ctl.offsetTop < dateControl.offsetTop + 100) {
                        ctl.style.visibility = hide ? 'hidden' : 'visible';
                    }
                }
            }
        }

        function showOverlay(display) {
	        $get("divOverlay").style.display = display ? "block" : "none";               
	    }

        function showWait(display) {
            $get("pleasewaitScreen").style.display = display ? "block" : "none"; 
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

	                showOverlay(true);
	                showErrorMessages(hasErrors() ? 'min' : 'none');
	                
	                document.getElementById("divFileUpload").style.display = "block";
	            }
	            else {
	                document.getElementById("divFileUpload").style.display = "none";
	                showOverlay(false);
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

	    var jQuerySetup;

	    function showMessage() {

	        //Reset jQuery setup
	        jQuerySetup();

	        //Reset current tab position
	        SetCurrentTab(iCurrentTab);
	        //Reset current error message display
	        showErrorMessages(window.iCurrentMessageState);
	        
	        showWait(false);
	        showOverlay(false);

	        //Reapply resizable column functionality to tables
	        //This is put here to ensure functionality is reapplied after partial/full postback.
	        ResizableColumns();		

	        if($get("txtActiveDDE").value.indexOf("dde") > 0) {
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

	            if (($get("txtPostbackMode").value == 2) || ($get("txtPostbackMode").value == 3)) 
	            {
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
	                            showOverlay(true);
	                            
	                            if(navigator.userAgent.indexOf("MSIE") > 0) {
	                                //Only IE can self-close windows that it didn't open
	                                window.close();
	                            } else {
	                                // Non-IE browsers can't self-close windows, show close message instead
	                                showWait(true);
	                                $get("pleasewaitText").innerHTML = "Please close your browser.";						  
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

	            showOverlay(true);
	            showErrorMessages('none');
	            $get("divSubmissionMessages").style.display = "block";
	            $get("divSubmissionMessages").style.visibility = "visible";
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
	    
	    function ResizeComboForForm(sender, args) {
	        psWebComboID = sender._id;
            
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
	                if ((oEl.offsetLeft + oEl.offsetWidth) > $get("bdyMain").clientWidth)
	                {
	                    oEl.style.left = "0px";
	                }                                                 
                  
	                //Hide the navigation icons as required
	                //Order to hide is: nav arrows go first, then 'page 1 of x'. Finally the search box goes.
	                //N.B. if the control is paged, min width is 420px before hiding the relevant controls

	                //Check to see if this is a paged control...
	                var oElDDL = document.getElementById(psWebComboID.replace("dde", "tcPagerDDL"));
	                if(eval(oElDDL)) {
	                    //This is a paged control, so different rules apply.
	                    if(oEl.offsetWidth<420) {
	                        document.getElementById(psWebComboID.replace("dde", "tcPagerBtns")).style.visibility = "hidden";
	                        document.getElementById(psWebComboID.replace("dde", "tcPagerBtns")).style.display = "none";
	                        document.getElementById(psWebComboID.replace("dde", "tcPageXofY")).style.visibility = "hidden";
	                        document.getElementById(psWebComboID.replace("dde", "tcPageXofY")).style.display = "none";
	                    }
	                    else {
	                        document.getElementById(psWebComboID.replace("dde", "tcPagerBtns")).style.visibility = "visible";
	                        document.getElementById(psWebComboID.replace("dde", "tcPagerBtns")).style.display = "";
	                        document.getElementById(psWebComboID.replace("dde", "tcPageXofY")).style.visibility = "visible";
	                        document.getElementById(psWebComboID.replace("dde", "tcPageXofY")).style.display = ""; 
	                    }
	                }
	                else {
	                    //Not a paged control
	                    if(oEl.offsetWidth<250) {
	                        document.getElementById(psWebComboID.replace("dde", "tcPagerBtns")).style.visibility = "hidden";
	                        document.getElementById(psWebComboID.replace("dde", "tcPagerBtns")).style.display = "none";
	                        document.getElementById(psWebComboID.replace("dde", "tcPageXofY")).style.visibility = "hidden";
	                        document.getElementById(psWebComboID.replace("dde", "tcPageXofY")).style.display = "none";
	                    }
	                    else {
	                        document.getElementById(psWebComboID.replace("dde", "tcPagerBtns")).style.visibility = "visible";
	                        document.getElementById(psWebComboID.replace("dde", "tcPagerBtns")).style.display = "";
	                        document.getElementById(psWebComboID.replace("dde", "tcPageXofY")).style.visibility = "visible";
	                        document.getElementById(psWebComboID.replace("dde", "tcPageXofY")).style.display = "";
	                    }                    
	                }
	            }
	        }
	        catch(e) {}
	    }

	    function scrollHeader(iGridID) {
	        //keeps the header table aligned with the gridview in record selectors and lookups.
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
	        var reTAB = /\t/g;        
	        var reSINGLEQUOTE = /\'/g;        
	        var reDECIMAL = new RegExp("\\" + window.localeDecimal, "gi");
	        var psWebComboID = "";

	        psWebComboID = sender._id;
	        
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
	                    if (iIndex >= 0) {
	                        sValueType = sSelectWhere.substring(0, iIndex);
	                        sSelectWhere = sSelectWhere.substr(iIndex+1);
	                    }
                        
	                    iIndex = sSelectWhere.indexOf("\t");
	                    if (iIndex >= 0) {
	                        sValueID = sSelectWhere.substring(0, iIndex);
	                        sSelectWhere = sSelectWhere.substr(iIndex+1);

	                        sControlType = sValueID.substr(sValueID.indexOf("_")+1);
	                        sControlType = sControlType.substr(sControlType.indexOf("_")+1);
	                        sControlType = sControlType.substring(0, sControlType.indexOf("_"));
                            
	                        if ((sControlType == 13) || (sControlType == 14)) {
	                            // Dropdown (13), Lookup (14)
	                            if (sControlType == 13) {  
	                                var ctlLookupValueCombo = document.getElementById(sValueID);
	                                sValue = ctlLookupValueCombo.value;
	                            }
	                            else {
	                                var ctlLookupValueCombo = document.getElementById(sValueID + "TextBox");
	                                sValue = ctlLookupValueCombo.value;                        	    
	                            }
                        	    
	                            if(sValueType == 11) {
	                                // Date value from lookup. Convert from locale format to yyyymmdd.
	                                if (sValue.length > 0) {
	                                    sTemp = GetDatePart(sValue, "Y");
                        	             
	                                    sSubTemp = "0" + GetDatePart(sValue, "M");
	                                    sTemp = sTemp + sSubTemp.substr(sSubTemp.length-2);
                        	            
	                                    sSubTemp = "0" + GetDatePart(sValue, "D");
	                                    sTemp = sTemp + sSubTemp.substr(sSubTemp.length-2);

	                                    sValue = sTemp;
	                                }
	                                else {
	                                    sValue = "";
	                                }
	                            }
	                            else {
	                                if((sValueType == 2) || (sValueType == 4)) {
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
	                                var ctlLookupValueCheckbox = document.getElementById(sValueID);
	                                fValue = ctlLookupValueCheckbox.checked;
	                                if (fValue == true) {
	                                    sValue = "1";
	                                }
	                                else {
	                                    sValue = "0";
	                                }
	                            }
	                            else {
	                                if (sControlType == 5) {
	                                    // Numeric (5)
	                                    var ctlLookupValueNumeric = igedit_getById(sValueID);
	                                    numValue = ctlLookupValueNumeric.getValue();
	                                    sValue = numValue.toString();
	                                }
	                                else {
	                                    if (sControlType == 7) {
	                                        // Date (7)
	                                        var ctlLookupValueDate = igdrp_getComboById(sValueID);
	                                        dtValue = ctlLookupValueDate.getValue();
	                                        if (dtValue) {
	                                            // Get year part.
	                                            sTemp = dtValue.getFullYear();
                        	            
	                                            // Get month part. Pad to 2 digits if required.
	                                            sSubTemp = "0" + (dtValue.getMonth() + 1);
	                                            sTemp = sTemp + sSubTemp.substr(sSubTemp.length-2);

	                                            // Get day part. Pad to 2 digits if required.
	                                            sSubTemp = "0" + dtValue.getDate();
	                                            sValue = sTemp + sSubTemp.substr(sSubTemp.length-2);
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

	    function FilterMobileLookup(sourceControlID) {
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
	        var reTAB = /\t/g;
	        var reSINGLEQUOTE = /\'/g;
	        var reDECIMAL = new RegExp("\\" + window.localeDecimal, "gi");

	        if (sourceControlID == "") { return; }

	        var lookups = getElementsBySearchValue(sourceControlID);
	        var AllLookupIDs = "";

	        for (var i = 0; i < lookups.length; i++) {

	            try {
	                var psWebComboID = lookups[i].name.replace("lookup", "");
	            }
	            catch (e) { var psWebComboID = ""; }


	            if (psWebComboID.length > 0) {

	                var sID = "lookup" + psWebComboID;
	                AllLookupIDs = AllLookupIDs + (i == 0 ? "" : "\t") + psWebComboID + "refresh";

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
	                                        var ctlLookupValueCombo = document.getElementById(sValueID);
	                                        sValue = ctlLookupValueCombo.value;
	                                    }
	                                    else {
	                                        var ctlLookupValueCombo = document.getElementById(sValueID + "TextBox");
	                                        if (!(eval(ctlLookupValueCombo))) { var ctlLookupValueCombo = document.getElementById(sValueID); }

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
	                                        var ctlLookupValueCheckbox = document.getElementById(sValueID);
	                                        fValue = ctlLookupValueCheckbox.checked;
	                                        if (fValue == true) {
	                                            sValue = "1";
	                                        }
	                                        else {
	                                            sValue = "0";
	                                        }
	                                    }
	                                    else {
	                                        if (sControlType == 5) {
	                                            // Numeric (5)
	                                            var ctlLookupValueNumeric = igedit_getById(sValueID);
	                                            numValue = ctlLookupValueNumeric.getValue();
	                                            sValue = numValue.toString();
	                                        }
	                                        else {
	                                            if (sControlType == 7) {
	                                                // Date (7)
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
	                                    document.getElementById(psWebComboID + "filterSQL").value = "";
	                                }
	                                else {
	                                    document.getElementById(psWebComboID + "filterSQL").value = sSelectWhere;
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

	        var formInputPrefix = "FI_";
	        var currentTab = $get(formInputPrefix + iCurrentTab + "_21_PageTab");
	        var currentPanel = $get(formInputPrefix + iCurrentTab + "_21_Panel");
	        var newTab = $get(formInputPrefix + iNewTab + "_21_PageTab");
	        var newPanel = $get(formInputPrefix + iNewTab + "_21_Panel");

	        document.getElementById("hdnDefaultPageNo").value = iNewTab;
    
	        try {
	            if(currentTab!=null) currentTab.style.display = "none";
      
	            if(currentPanel!=null) currentPanel.style.borderBottom = "1px solid black";
        
	            if(newTab!=null) newTab.style.display = "block";
        
	            if(newPanel!=null) newPanel.style.borderBottom = "1px solid white";
        
	            window.iCurrentTab = iNewTab;            

	        }
	        catch (e) {}
	    }

    function disposeTree(sender, args) {

        //http://support.microsoft.com/?kbid=2000262

        try {

            var elements = args.get_panelsUpdating();
            for (var i = elements.length - 1; i >= 0; i--) {
                var element = elements[i];
                var allnodes = element.getElementsByTagName('*'),
                    length = allnodes.length;
                var nodes = new Array(length);
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
            } 
        } catch (e) { }
    }

try {
    Sys.WebForms.PageRequestManager.getInstance().add_pageLoading(disposeTree);
}
catch (e) { }
