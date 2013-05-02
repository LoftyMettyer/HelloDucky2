<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

    <script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>           

<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0">
    <param name="LPKPath" value="lpks/main.lpk">
</object>

<form id="frmpicklistSelectionUseful" name="frmpicklistSelectionUseful" style="visibility: hidden; display: none">
    <input type="hidden" id="txtIEVersion" name="txtIEVersion" value='<%=session("IEVersion")%>'>
    <input type='hidden' id="txtSelectionType" name="txtSelectionType" value='<%=session("selectionType")%>'>
    <input type='hidden' id="txtTableID" name="txtTableID" value='<%=session("selectionTableID")%>'>
    <input type="hidden" id="txtMenuSaved" name="txtMenuSaved" value="0">
</form>

<script type="text/javascript">    
    function picklistSelection_window_onload() {

        $("#picklistworkframe").attr("data-framesource", "PICKLISTSELECTION");

        fOK = true;

        var frmUseful = document.getElementById("frmpicklistSelectionUseful");

        if ((frmUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
            (frmUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

            var sErrMsg = txtpicklistSelectionErrorDescription.value;
            if (sErrMsg.length > 0) {
                fOK = false;
                OpenHR.messageBox(sErrMsg);
            }

            if (fOK == true) {
                if (selectView.length == 0) {
                    fOK = false;
                    OpenHR.messageBox("You do not have permission to read the table.");
                }
            }
	
            if (fOK == true) {
                if (selectOrder.length == 0) {
                    fOK = false;
                    OpenHR.messageBox("You do not have permission to use any of the table orders.");
                }
            }
        }

        cmdCancel.focus();
	
        if ((frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") ||
            (frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST")) {
            setGridFont(ssOleDBGridSelRecords);

            // Set focus onto one of the form controls. 
            // NB. This needs to be done before making any reference to the grid
            ssOleDBGridSelRecords.focus();		

            if (ssOleDBGridSelRecords.rows > 0) {
                // Select the top row.
                ssOleDBGridSelRecords.MoveFirst();
                ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
            }

            refreshControls();
        }
        else {
            setGridFont(ssOleDBGridSelRecords);

            setMenuFont(abMainMenu);

            abMainMenu.Attach();
            abMainMenu.DataPath = "misc\\mainmenu.htm";
            abMainMenu.RecalcLayout();
		
//            window.parent.dialogLeft = new String((screen.width - (9 * screen.width / 10)) / 2) + "px";
//            window.parent.dialogTop =  new String((screen.height - (3 * screen.height / 4)) / 2) + "px";
//            window.parent.dialogWidth = new String((9 * screen.width / 10)) + "px";
//            window.parent.dialogHeight = new String((3 * screen.height / 4)) + "px";
			
            txtTableID.value = frmUseful.txtTableID.value;
            txtViewID.value = selectView.options[selectView.selectedIndex].value;
            txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;

            loadAddRecords();
        }
    }
</script>

<script type="text/javascript">

    function refreshControls() {
        var fNoneSelected;

        fNoneSelected = (ssOleDBGridSelRecords.SelBookmarks.Count == 0);

        button_disable(cmdOK, fNoneSelected);

        if ((frmpicklistSelectionUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
            (frmpicklistSelectionUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

            if (selectOrder.length <= 1) {
                combo_disable(selectOrder, true);
                button_disable(btnGoOrder, true);
            }

            if (selectView.length <= 1) {
                combo_disable(selectView, true);
                button_disable(btnGoView, true);
            }
        }
    }

    function cancelClick() {
        $(".popup").dialog("close");
        $("#workframeset").show();
    }

    function makeSelection() {

        var frmUseful = document.getElementById("frmpicklistSelectionUseful");

        if (frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") {
            try {
                var frmParentUseful = OpenHR.getForm("workframe", "frmUseful");
                frmParentUseful.txtChanged.value = 1;
            }
            catch (e) {
            }

            // Go to the prompted values form to get any required prompts. 
            frmPrompt.filterID.value = selectedRecordID();
            OpenHR.showInReportFrame(frmPrompt);

        }
        else {
            if (frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST") {
                try {
                    picklistdef_makeSelection(frmUseful.txtSelectionType.value, selectedRecordID(), "");
                }
                catch (e) {
                }
            }
            else {
                sSelectedIDs = "";

                ssOleDBGridSelRecords.Redraw = false;
                for (iIndex = 0; iIndex < ssOleDBGridSelRecords.selbookmarks.Count() ; iIndex++) {
                    ssOleDBGridSelRecords.bookmark = ssOleDBGridSelRecords.selbookmarks(iIndex);

                    sRecordID = ssOleDBGridSelRecords.Columns("id").Value;

                    if (sSelectedIDs.length > 0) {
                        sSelectedIDs = sSelectedIDs + ",";
                    }
                    sSelectedIDs = sSelectedIDs + sRecordID;
                }
                ssOleDBGridSelRecords.Redraw = true;

                try {
                    var frmParentUseful = OpenHR.getForm("workframe", "frmUseful");
                    frmParentUseful.txtChanged.value = 1;

                    picklistdef_makeSelection(frmUseful.txtSelectionType.value, 0, sSelectedIDs);
                }
                catch (e) {
                }
            }
        }
    }

    /* Return the ID of the record selected in the find form. */
    function selectedRecordID() {
        var iRecordID;

        debugger;

        iRecordID = 0;

        if (ssOleDBGridSelRecords.SelBookmarks.Count > 0) {
            iRecordID = ssOleDBGridSelRecords.Columns(0).Value;
        }

        return (iRecordID);
    }

    function locateRecord(psSearchFor) {
        var fFound;

        fFound = false;
	
        ssOleDBGridSelRecords.Redraw = false;

        ssOleDBGridSelRecords.MoveLast();
        ssOleDBGridSelRecords.MoveFirst();

        ssOleDBGridSelRecords.SelBookmarks.removeall();
	
        for (iIndex = 1; iIndex <= ssOleDBGridSelRecords.rows; iIndex++) 
        {	
            var sGridValue = new String(ssOleDBGridSelRecords.Columns(0).value);
            sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
            if (sGridValue == psSearchFor.toUpperCase()) 
            {
                ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
                fFound = true;
                break;
            }

            if (iIndex < ssOleDBGridSelRecords.rows) 
            {
                ssOleDBGridSelRecords.MoveNext();
            }
            else 
            {
                break;
            }
        }

        if ((fFound == false) && (ssOleDBGridSelRecords.rows > 0)) 
        {
            // Select the top row.
            ssOleDBGridSelRecords.MoveFirst();
            ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
        }

        ssOleDBGridSelRecords.Redraw = true;
    }

    function goView() {

        // Get the picklistSelectionData.asp to get the find records.
        var dataForm = OpenHR.getForm("dataframe", "frmPicklistGetData");
        dataForm.txtTableID.value = frmpicklistSelectionUseful.txtTableID.value;
        dataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
        dataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
        dataForm.txtFirstRecPos.value = 1;
        dataForm.txtCurrentRecCount.value = 0;
        dataForm.txtPageAction.value = "LOAD";

        refreshData();
    }

    function goOrder() {
        // Get the picklistSelectionData.asp to get the find records.
        var dataForm = OpenHR.getForm("dataframe", "frmPicklistGetData");
        dataForm.txtTableID.value = frmpicklistSelectionUseful.txtTableID.value;
        dataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
        dataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
        dataForm.txtFirstRecPos.value = 1;
        dataForm.txtCurrentRecCount.value = 0;
        dataForm.txtPageAction.value = "LOAD";

        refreshData();
    }

    function selectedOrderID() {
        return selectOrder.options[selectOrder.selectedIndex].value;
    }

    function selectedViewID() {
        return selectView.options[selectView.selectedIndex].value;
    }

    function refreshMenu() {
        var sTemp;
	
        if (abMainMenu.Bands.Count() > 0) {
            enableMenu();

            var frmData = OpenHR.getForm("dataframe","frmPicklistData");        

            for (i=0; i< abMainMenu.tools.count(); i++) {
                abMainMenu.tools(i).visible = false;
            }				
		
            abMainMenu.Bands("mnuMainMenu").visible = false;
            abMainMenu.Bands("mnubandMainToolBar").visible = true;

            // Enable the record editing options as necessary.
            abMainMenu.tools("mnutoolFirstRecord").visible = true;
            abMainMenu.tools("mnutoolFirstRecord").enabled = (frmData.txtIsFirstPage.value != "True");
            abMainMenu.tools("mnutoolPreviousRecord").visible = true;
            abMainMenu.tools("mnutoolPreviousRecord").enabled = (frmData.txtIsFirstPage.value != "True");
            abMainMenu.tools("mnutoolNextRecord").visible = true;
            abMainMenu.tools("mnutoolNextRecord").enabled = (frmData.txtIsLastPage.value != "True");
            abMainMenu.tools("mnutoolLastRecord").visible = true;
            abMainMenu.tools("mnutoolLastRecord").enabled = (frmData.txtIsLastPage.value != "True");

            abMainMenu.tools("mnutoolLocateRecordsCaption").visible = true;
            abMainMenu.tools("mnutoolLocateRecords").visible = (frmData.txtFirstColumnType.value != "-7");
            abMainMenu.Tools("mnutoolLocateRecordsLogic").CBList.Clear();
            abMainMenu.Tools("mnutoolLocateRecordsLogic").CBList.AddItem("True");
            abMainMenu.Tools("mnutoolLocateRecordsLogic").CBList.AddItem("False");
            abMainMenu.tools("mnutoolLocateRecordsLogic").visible = (frmData.txtFirstColumnType.value == "-7");

            sCaption = "";
            if (frmData.txtRecordCount.value > 0) {
                iStartPosition = new Number(frmData.txtFirstRecPos.value);
                iEndPosition = new Number(frmData.txtRecordCount.value);
                iEndPosition = iStartPosition - 1 + iEndPosition;
                sCaption = "Records " +
                    iStartPosition + 
                    " to " +
                    iEndPosition +
                    " of " +
                    frmData.txtTotalRecordCount.value;
            }
            else {

                sCaption = "No Records";
            }

            abMainMenu.tools("mnutoolRecordPosition").visible = true;
            abMainMenu.Bands("mnubandMainToolBar").tools("mnutoolRecordPosition").caption = sCaption;
						
            try
            {
                abMainMenu.Attach();
                abMainMenu.RecalcLayout();
                abMainMenu.ResetHooks();
                abMainMenu.Refresh();
            }
            catch(e) {}
		
            // Adjust the frameset dimensions to suit the size of the menu.
            /*lngMenuHeight = abMainMenu.Bands("mnubandMainToolBar").height;
            sTemp = new String(lngMenuHeight);
            if(frmUseful.txtIEVersion.value >= 5.5) 
            {
                window.parent.document.all.item("mainframeset").rows = "*, " + sTemp;
            }
            else 
            {
                window.parent.document.all.item("mainframeset").rows = "*, 0";
            }*/
        }
    }

    function reloadPage(psAction, psLocateValue) {
        var sConvertedValue;
        var sDecimalSeparator;
        var sThousandSeparator;
        var sPoint;
        var iIndex ;
        var iTempSize;
        var iTempDecimals;
	
        sDecimalSeparator = "\\";
        sDecimalSeparator = sDecimalSeparator.concat(ASRIntranetFunctions.LocaleDecimalSeparator);
        var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

        sThousandSeparator = "\\";
        sThousandSeparator = sThousandSeparator.concat(ASRIntranetFunctions.LocaleThousandSeparator);
        var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

        sPoint = "\\.";
        var rePoint = new RegExp(sPoint, "gi");

        fValidLocateValue = true;

        var dataForm = OpenHR.getForm("dataframe","frmPicklistData");        
        var getDataForm = OpenHR.getForm("dataframe", "frmPicklistGetData");

        if (psAction == "LOCATE") {
            // Check that the entered value is valid for the first order column type.
            iDataType = dataForm.txtFirstColumnType.value;

            if ((iDataType == 2) || (iDataType == 4)) {
                // Numeric/Integer column.
                // Ensure that the value entered is numeric.
                if (psLocateValue.length == 0) {
                    psLocateValue = "0";
                }

                // Convert the value from locale to UK settings for use with the isNaN funtion.
                sConvertedValue = new String(psLocateValue);
                // Remove any thousand separators.
                sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
                psLocateValue = sConvertedValue;

                // Convert any decimal separators to '.'.
                if (ASRIntranetFunctions.LocaleDecimalSeparator != ".") {
                    // Remove decimal points.
                    sConvertedValue = sConvertedValue.replace(rePoint, "A");
                    // replace the locale decimal marker with the decimal point.
                    sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
                }

                if (isNaN(sConvertedValue) == true) {
                    fValidLocateValue = false;
                    OpenHR.messageBox("Invalid numeric value entered.");
                }
                else {
                    psLocateValue = sConvertedValue;
                    iIndex = sConvertedValue.indexOf(".");
                    if (iDataType == 4) {
                        // Ensure that integer columns are compared with integer values.
                        if (iIndex >= 0 ) {
                            fValidLocateValue = false;
                            OpenHR.messageBox("Invalid integer value entered.");
                        }
                    } 
                    else {
                        // Ensure numeric columns are compared with numeric values that do not exceed
                        // their defined size and decimals settings.
                        if (iIndex >= 0) {
                            iTempSize = iIndex;
                            iTempDecimals = sConvertedValue.length - iIndex - 1;
                        }
                        else {
                            iTempSize = sConvertedValue.length;
                            iTempDecimals = 0;
                        }
					
                        if ((sConvertedValue.substr(0,1) == "+") ||
                            (sConvertedValue.substr(0,1) == "-")) {
                            iTempSize = iTempSize - 1;
                        }

                        if(iTempSize > (dataForm.txtFirstColumnSize.value - dataForm.txtFirstColumnDecimals.value)) {
                            fValidLocateValue = false;
                            OpenHR.messageBox("The value cannot have more than " + (dataForm.txtFirstColumnSize.value - dataForm.txtFirstColumnDecimals.value) + " digit(s) to the left of the decimal separator.");
                        }
                        else {
                            if(iTempDecimals > dataForm.txtFirstColumnDecimals.value) {
                                fValidLocateValue = false;
                                OpenHR.messageBox("The value cannot have more than " + dataForm.txtFirstColumnDecimals.value + " decimal place(s).");
                            }
                        }
                    }
                }
            }
            else {
                if (iDataType == 11) {
                    // Date column.
                    // Ensure that the value entered is a date.
                    if (psLocateValue.length > 0) {
                        // Convert the date to SQL format (use this as a validation check).
                        // An empty string is returned if the date is invalid.
                        psLocateValue = convertLocaleDateToSQL(psLocateValue)
                        if (psLocateValue.length = 0) {
                            fValidLocateValue = false;
                            OpenHR.messageBox("Invalid date value entered.");
                        }
                    }
                }
            }
        }
	
        if (fValidLocateValue == true) {
            disableMenu();

            // Get the optionData.asp to get the link find records.
            getDataForm.txtTableID.value = frmpicklistSelectionUseful.txtTableID.value;
            getDataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
            getDataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
            getDataForm.txtFirstRecPos.value = dataForm.txtFirstRecPos.value;
            getDataForm.txtCurrentRecCount.value = dataForm.txtRecordCount.value;
            getDataForm.txtGotoLocateValue.value = psLocateValue;
            getDataForm.txtPageAction.value = psAction;

            refreshData();
        }

        // Clear the locate value from the menu.
        abMainMenu.Tools("mnutoolLocateRecords").Text = "";
    }

    function disableMenu() {
        for(iLoop = 0; iLoop < abMainMenu.Bands.Item("mnubandMainToolBar").Tools.Count(); iLoop ++) {
            abMainMenu.Bands.Item("mnubandMainToolBar").tools.Item(iLoop).Enabled = false;
        }

        abMainMenu.RecalcLayout();
        abMainMenu.ResetHooks();
        abMainMenu.Refresh();
    }

    function enableMenu() {
        for(iLoop = 0; iLoop < abMainMenu.Bands.Item("mnubandMainToolBar").Tools.Count(); iLoop ++) {
            abMainMenu.Bands.Item("mnubandMainToolBar").tools.Item(iLoop).Enabled = true;
        }
    }

    function convertLocaleDateToSQL(psDateString)
    { 
        /* Convert the given date string (in locale format) into 
        SQL format (mm/dd/yyyy). */
        var sDateFormat;
        var iDays;
        var iMonths;
        var iYears;
        var sDays;
        var sMonths;
        var sYears;
        var iValuePos;
        var sTempValue;
        var sValue;
        var iLoop;
		
        sDateFormat = OpenHR.LocaleDateFormat.LocaleDateFormat;

        sDays="";
        sMonths="";
        sYears="";
        iValuePos = 0;

        // Trim leading spaces.
        sTempValue = psDateString.substr(iValuePos,1);
        while (sTempValue.charAt(0) == " ") 
        {
            iValuePos = iValuePos + 1;		
            sTempValue = psDateString.substr(iValuePos,1);
        }

        for (iLoop=0; iLoop<sDateFormat.length; iLoop++)  {
            if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'D') && (sDays.length==0)){
                sDays = psDateString.substr(iValuePos,1);
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) {
                    sDays = sDays.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;		
            }

            if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'M') && (sMonths.length==0)){
                sMonths = psDateString.substr(iValuePos,1);
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) {
                    sMonths = sMonths.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;
            }

            if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'Y') && (sYears.length==0)){
                sYears = psDateString.substr(iValuePos,1);
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) {
                    sYears = sYears.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) {
                    sYears = sYears.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) {
                    sYears = sYears.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;
            }

            // Skip non-numerics
            sTempValue = psDateString.substr(iValuePos,1);
            while (isNaN(sTempValue) == true) {
                iValuePos = iValuePos + 1;		
                sTempValue = psDateString.substr(iValuePos,1);
            }
        }

        while (sDays.length < 2) {
            sTempValue = "0";
            sDays = sTempValue.concat(sDays);
        }

        while (sMonths.length < 2) {
            sTempValue = "0";
            sMonths = sTempValue.concat(sMonths);
        }

        while (sYears.length < 2) {
            sTempValue = "0";
            sYears = sTempValue.concat(sYears);
        }

        if (sYears.length == 2) {
            iValue = parseInt(sYears);
            if (iValue < 30) {
                sTempValue = "20";
            }
            else {
                sTempValue = "19";
            }
		
            sYears = sTempValue.concat(sYears);
        }

        while (sYears.length < 4) {
            sTempValue = "0";
            sYears = sTempValue.concat(sYears);
        }

        sTempValue = sMonths.concat("/");
        sTempValue = sTempValue.concat(sDays);
        sTempValue = sTempValue.concat("/");
        sTempValue = sTempValue.concat(sYears);
	
        sValue = OpenHR.ConvertSQLDateToLocale(sTempValue);

        iYears = parseInt(sYears);
	
        while (sMonths.substr(0, 1) == "0") {
            sMonths = sMonths.substr(1);
        }
        iMonths = parseInt(sMonths);
	
        while (sDays.substr(0, 1) == "0") {
            sDays = sDays.substr(1);
        }
        iDays = parseInt(sDays);

        var newDateObj = new Date(iYears, iMonths - 1, iDays);
        if ((newDateObj.getDate() != iDays) || 
            (newDateObj.getMonth() + 1 != iMonths) || 
            (newDateObj.getFullYear() != iYears)) {
            return "";
        }
        else {
            return sTempValue;
        }
    }

</script>
    
<script type="text/javascript">

    function picklistSelection_addhandlers() {
        OpenHR.addActiveXHandler("abMainMenu", "DataReady", abMainMenu_DataReady);
        OpenHR.addActiveXHandler("abMainMenu", "PreCustomizeMenu", abMainMenu_PreCustomizeMenu);
        OpenHR.addActiveXHandler("abMainMenu", "Click", abMainMenu_Click);
        OpenHR.addActiveXHandler("abMainMenu", "KeyDown", abMainMenu_KeyDown);
        OpenHR.addActiveXHandler("abMainMenu", "ComboSelChange", abMainMenu_ComboSelChange);
        OpenHR.addActiveXHandler("abMainMenu", "PreSysMenu", abMainMenu_PreSysMenu);
        OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "RowColChange", ssOleDBGridSelRecords_RowColChange);
        OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "DblClick", ssOleDBGridSelRecords_DblClick);
        OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "KeyPress", ssOleDBGridSelRecords_KeyPress);
    }

     function abMainMenu_DataReady() {

         var sKey;
         sKey = new String("tempmenufilepath_");
         sKey = sKey.concat(window.parent.window.dialogArguments.window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
         sPath = ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
         if(sPath == "") {
             sPath = "c:\\";
         }

         if(sPath == "<NONE>") {
             frmpicklistSelectionUseful.txtMenuSaved.value = 1;
             abMainMenu.RecalcLayout();
         }
         else {
             if (sPath.substr(sPath.length - 1, 1) != "\\") {
                 sPath = sPath.concat("\\");
             }
		
             sPath = sPath.concat("tempmenu.asp");
             if ((abMainMenu.Bands.Count() > 0) && (frmpicklistSelectionUseful.txtMenuSaved.value == 0)) {
                 try {
                     abMainMenu.save(sPath, "");
                 }
                 catch(e) {
                     OpenHR.messageBox("The specified temporary menu file path cannot be written to. The temporary menu file path will be cleared."); 
                     sKey = new String("tempMenuFilePath_");
                     sKey = sKey.concat(window.parent.window.dialogArguments.window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
                     ASRIntranetFunctions.SaveRegistrySetting("HR Pro", "DataPaths", sKey, "<NONE>");
                 }
                 frmpicklistSelectionUseful.txtMenuSaved.value = 1;
             }
             else {
                 if ((abMainMenu.Bands.Count() == 0) && (frmpicklistSelectionUseful.txtMenuSaved.value == 1)) {
                     abMainMenu.DataPath = sPath;
                     abMainMenu.RecalcLayout();
                     return;
                 }
             }
         }
     }

     function abMainMenu_PreCustomizeMenu(pfCancel) {
         pfCancel = true;
         OpenHR.messageBox("The menu cannot be customized. Errors will occur if you attempt to customize it. Click anywhere in your browser to remove the dummy customisation menu.");         
     }
     
     function abMainMenu_Click(pTool) {
    
         switch (pTool.name) {
             case "mnutoolFirstRecord" :
                 reloadPage("MOVEFIRST", "");
                 break;
             case "mnutoolPreviousRecord" :
                 reloadPage("MOVEPREVIOUS", "");
                 break;
             case "mnutoolNextRecord" :
                 reloadPage("MOVENEXT", "");
                 break;
             case "mnutoolLastRecord" :
                 reloadPage("MOVELAST", "");
                 break;
         }     
     }

     function abMainMenu_KeyDown(piKeyCode, piShift) {
         iIndex = abMainMenu.ActiveBand.CurrentTool;
	
         if (abMainMenu.ActiveBand.Tools(iIndex).Name == "mnutoolLocateRecords") {
             if (piKeyCode == 13) {
                 sLocateValue = abMainMenu.ActiveBand.Tools(iIndex).Text;

                 reloadPage("LOCATE", sLocateValue);
             }
         }
     }

     function abMainMenu_ComboSelChange(pTool) {
         if (pTool.Name == "mnutoolLocateRecordsLogic") {
             sLocateValue = pTool.Text;

             reloadPage("LOCATE", sLocateValue);
         }
     }

     function abMainMenu_PreSysMenu(pBand) {
         if(pBand.Name == "SysCustomize") {
             pBand.Tools.RemoveAll();
         }
     }

     function ssOleDBGridSelRecords_RowColChange() {
         refreshControls();         
     }

     function ssOleDBGridSelRecords_DblClick() {
         if (frmpicklistSelectionUseful.txtSelectionType.value != "ALL") {
             makeSelection();
         }         
     }
     
     function ssOleDBGridSelRecords_KeyPress(iKeyAscii) {

         if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {	
             var dtTicker = new Date();
             var iThisTick = new Number(dtTicker.getTime());
             if (txtLastKeyFind.value.length > 0) {
                 var iLastTick = new Number(txtTicker.value);
             }
             else {
                 var iLastTick = new Number("0");
             }
		
             if (iThisTick > (iLastTick + 1500)) {
                 var sFind = String.fromCharCode(iKeyAscii);
             }
             else {
                 var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
             }
		
             txtTicker.value = iThisTick;
             txtLastKeyFind.value = sFind;

             locateRecord(sFind);
         }         
     }

</script>


<%
    If (UCase(Session("selectionType")) <> UCase("picklist")) And _
        (UCase(Session("selectionType")) <> UCase("filter")) Then
%>

<object classid="clsid:6976CB54-C39B-4181-B1DC-1A829068E2E7"
    codebase="cabs/COAInt_Client.cab#Version=1,0,0,5"
    height="32" id="abMainMenu" name="abMainMenu" style="LEFT: 0px; TOP: 0px" width="100%" viewastext>
    <param name="_ExtentX" value="847">
    <param name="_ExtentY" value="847">
</object>

<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="95%">
    <%
    Else
    %>
    <table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
        <%
        End If
        %>
        <tr>
            <td>
                <table align="center" class="invisible" cellspacing="0" cellpadding="0" width="100%" height="100%">
                    <tr height="10">
                        <td colspan="3" align="center" height="10">
                            <h3 align="center">
                                <% 
                                        If UCase(Session("selectionType")) = UCase("picklist") Then
                                %>	
                            Select Picklist
                                    <%
                                    Else
                                        If UCase(Session("selectionType")) = UCase("filter") Then
                                    %>
			                Select Filter
                                    <%
                                    Else
                                    %>		
                            Select Records
                                    <%
                                    End If
                                End If
                                    %>
                            </h3>
                        </td>
                    </tr>
                    <tr>
                        <td width="20"></td>
                        <td>
                            <%
                                    Dim sErrorDescription As String
                                    Dim cmdSelRecords
                                    Dim prmTableID
                                    Dim prmUser
                                    Dim rstSelRecords
                                    Dim lngRowCount As Long
                                    Dim cmdViewRecords
                                    Dim prmDfltOrderID
                                    Dim rstViewRecords
                                    Dim sFailureDescription As String
                                    Dim cmdOrderRecords
                                    Dim prmViewID
                                    Dim rstOrderRecords                                   
                                    
                                    Session("optionLinkViewID") = 0
                                    Session("optionLinkOrderID") = 0
	
                                    sErrorDescription = ""

                                    If (UCase(Session("selectionType")) = UCase("picklist")) Or _
                                        (UCase(Session("selectionType")) = UCase("filter")) Then

                                        cmdSelRecords = Server.CreateObject("ADODB.Command")
                                        cmdSelRecords.CommandType = 4
                                        cmdSelRecords.ActiveConnection = Session("databaseConnection")

                                        If UCase(Session("selectionType")) = UCase("picklist") Then
                                            cmdSelRecords.CommandText = "spASRIntGetAvailablePicklists"

                                            prmTableID = cmdSelRecords.CreateParameter("tableID", 3, 1) ' 3 = integer, 1 = input
                                            cmdSelRecords.Parameters.Append(prmTableID)
                                            prmTableID.value = cleanNumeric(CLng(Session("selectionTableID")))
			
                                            prmUser = cmdSelRecords.CreateParameter("user", 200, 1, 255)
                                            cmdSelRecords.Parameters.Append(prmUser)
                                            prmUser.value = Session("username")
                                        Else
                                            cmdSelRecords.CommandText = "spASRIntGetAvailableFilters"

                                            prmTableID = cmdSelRecords.CreateParameter("tableID", 3, 1) ' 3 = integer, 1 = input
                                            cmdSelRecords.Parameters.Append(prmTableID)
                                            prmTableID.value = cleanNumeric(CLng(Session("selectionTableID")))
			
                                            prmUser = cmdSelRecords.CreateParameter("user", 200, 1, 255)
                                            cmdSelRecords.Parameters.Append(prmUser)
                                            prmUser.value = Session("username")
                                        End If

                                        Err.Clear()
                                        rstSelRecords = cmdSelRecords.Execute

                                        ' Instantiate and initialise the grid. 
                            %>
                            <object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id="ssOleDBGridSelRecords" name="ssOleDBGridSelRecords" codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px">
                                <param name="ScrollBars" value="4">
                                <param name="_Version" value="196616">
                                <param name="DataMode" value="2">
                                <param name="Cols" value="0">
                                <param name="Rows" value="0">
                                <param name="BorderStyle" value="1">
                                <param name="RecordSelectors" value="0">
                                <param name="GroupHeaders" value="0">
                                <param name="ColumnHeaders" value="0">
                                <param name="GroupHeadLines" value="0">
                                <param name="HeadLines" value="0">
                                <param name="FieldDelimiter" value="(None)">
                                <param name="FieldSeparator" value="(Tab)">
                                <param name="Col.Count" value="<%=rstSelRecords.fields.count%>">
                                <param name="stylesets.count" value="0">
                                <param name="TagVariant" value="EMPTY">
                                <param name="UseGroups" value="0">
                                <param name="HeadFont3D" value="0">
                                <param name="Font3D" value="0">
                                <param name="DividerType" value="3">
                                <param name="DividerStyle" value="1">
                                <param name="DefColWidth" value="0">
                                <param name="BeveColorScheme" value="2">
                                <param name="BevelColorFrame" value="-2147483642">
                                <param name="BevelColorHighlight" value="-2147483628">
                                <param name="BevelColorShadow" value="-2147483632">
                                <param name="BevelColorFace" value="-2147483633">
                                <param name="CheckBox3D" value="-1">
                                <param name="AllowAddNew" value="0">
                                <param name="AllowDelete" value="0">
                                <param name="AllowUpdate" value="0">
                                <param name="MultiLine" value="0">
                                <param name="ActiveCellStyleSet" value="">
                                <param name="RowSelectionStyle" value="0">
                                <param name="AllowRowSizing" value="0">
                                <param name="AllowGroupSizing" value="0">
                                <param name="AllowColumnSizing" value="0">
                                <param name="AllowGroupMoving" value="0">
                                <param name="AllowColumnMoving" value="0">
                                <param name="AllowGroupSwapping" value="0">
                                <param name="AllowColumnSwapping" value="0">
                                <param name="AllowGroupShrinking" value="0">
                                <param name="AllowColumnShrinking" value="0">
                                <param name="AllowDragDrop" value="0">
                                <param name="UseExactRowCount" value="-1">
                                <param name="SelectTypeCol" value="0">
                                <param name="SelectTypeRow" value="1">
                                <param name="SelectByCell" value="-1">
                                <param name="BalloonHelp" value="0">
                                <param name="RowNavigation" value="1">
                                <param name="CellNavigation" value="0">
                                <param name="MaxSelectedRows" value="1">
                                <param name="HeadStyleSet" value="">
                                <param name="StyleSet" value="">
                                <param name="ForeColorEven" value="0">
                                <param name="ForeColorOdd" value="0">
                                <param name="BackColorEven" value="16777215">
                                <param name="BackColorOdd" value="16777215">
                                <param name="Levels" value="1">
                                <param name="RowHeight" value="503">
                                <param name="ExtraHeight" value="0">
                                <param name="ActiveRowStyleSet" value="">
                                <param name="CaptionAlignment" value="2">
                                <param name="SplitterPos" value="0">
                                <param name="SplitterVisible" value="0">
                                <param name="Columns.Count" value="<%=rstSelRecords.fields.count%>">
                                <%
                                        For iLoop = 0 To (rstSelRecords.fields.count - 1)
                                            If rstSelRecords.fields(iLoop).name <> "name" Then
                                %>
                                <param name="Columns(<%=iLoop%>).Width" value="0">
                                <param name="Columns(<%=iLoop%>).Visible" value="0">
                                <% 
                                    Else
                                %>
                                <param name="Columns(<%=iLoop%>).Width" value="100000">
                                <param name="Columns(<%=iLoop%>).Visible" value="-1">
                                <%
                                    End If
                                %>
                                <param name="Columns(<%=iLoop%>).Columns.Count" value="1">
                                <param name="Columns(<%=iLoop%>).Caption" value="<%=replace(rstSelRecords.fields(iLoop).name, "_", "> ")%>">
                                <param name="Columns(<%=iLoop%>).Name" value="<%=rstSelRecords.fields(iLoop).name%>">
                                <param name="Columns(<%=iLoop%>).Alignment" value="0">
                                <param name="Columns(<%=iLoop%>).CaptionAlignment" value="3">
                                <param name="Columns(<%=iLoop%>).Bound" value="0">
                                <param name="Columns(<%=iLoop%>).AllowSizing" value="1">
                                <param name="Columns(<%=iLoop%>).DataField" value="Column <%=iLoop%>">
                                <param name="Columns(<%=iLoop%>).DataType" value="8">
                                <param name="Columns(<%=iLoop%>).Level" value="0">
                                <param name="Columns(<%=iLoop%>).NumberFormat" value="">
                                <param name="Columns(<%=iLoop%>).Case" value="0">
                                <param name="Columns(<%=iLoop%>).FieldLen" value="4096">
                                <param name="Columns(<%=iLoop%>).VertScrollBar" value="0">
                                <param name="Columns(<%=iLoop%>).Locked" value="0">
                                <param name="Columns(<%=iLoop%>).Style" value="0">
                                <param name="Columns(<%=iLoop%>).ButtonsAlways" value="0">
                                <param name="Columns(<%=iLoop%>).RowCount" value="0">
                                <param name="Columns(<%=iLoop%>).ColCount" value="1">
                                <param name="Columns(<%=iLoop%>).HasHeadForeColor" value="0">
                                <param name="Columns(<%=iLoop%>).HasHeadBackColor" value="0">
                                <param name="Columns(<%=iLoop%>).HasForeColor" value="0">
                                <param name="Columns(<%=iLoop%>).HasBackColor" value="0">
                                <param name="Columns(<%=iLoop%>).HeadForeColor" value="0">
                                <param name="Columns(<%=iLoop%>).HeadBackColor" value="0">
                                <param name="Columns(<%=iLoop%>).ForeColor" value="0">
                                <param name="Columns(<%=iLoop%>).BackColor" value="0">
                                <param name="Columns(<%=iLoop%>).HeadStyleSet" value="">
                                <param name="Columns(<%=iLoop%>).StyleSet" value="">
                                <param name="Columns(<%=iLoop%>).Nullable" value="1">
                                <param name="Columns(<%=iLoop%>).Mask" value="">
                                <param name="Columns(<%=iLoop%>).PromptInclude" value="0">
                                <param name="Columns(<%=iLoop%>).ClipMode" value="0">
                                <param name="Columns(<%=iLoop%>).PromptChar" value="95">
                                <%
                                    Next
                                %>
                                <param name="UseDefaults" value="-1">
                                <param name="TabNavigation" value="1">
                                <param name="_ExtentX" value="17330">
                                <param name="_ExtentY" value="1323">
                                <param name="_StockProps" value="79">
                                <param name="Caption" value="">
                                <param name="ForeColor" value="0">
                                <param name="BackColor" value="16777215">
                                <param name="Enabled" value="-1">
                                <param name="DataMember" value="">
                                <% 
                                        lngRowCount = 0
                                        Do While Not rstSelRecords.EOF
                                            For iLoop = 0 To (rstSelRecords.fields.count - 1)
                                %>
                                <param name="Row(<%=lngRowCount%>).Col(<%=iLoop%>)" value="<%=replace(rstSelRecords.Fields(iLoop).Value, "_", " ")%>">
                                <%
                                    Next
                                    lngRowCount = lngRowCount + 1
                                    rstSelRecords.MoveNext()
                                Loop
                                %>
                                <param name="Row.Count" value="<%=lngRowCount%>">
                            </object>
                            <%	
                                    rstSelRecords.close()
                                    rstSelRecords = Nothing

                                    ' Release the ADO command object.
                                    cmdSelRecords = Nothing
		
                                Else
                                    ' Select individual employee records.
                            %>
                            <table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
                                <tr height="10">
                                    <td height="10">
                                        <table width="100%" height="10" class="invisible" cellspacing="0" cellpadding="0">
                                            <tr height="10">
                                                <td width="40" height="10">View :
                                                </td>
                                                <td width="10" height="10">&nbsp;
                                                </td>
                                                <td width="175" height="10">
                                                    <select id="selectView" name="selectView" class="combo" style="HEIGHT: 10px; WIDTH: 200px">
                                                        <%
                                                                If Len(sErrorDescription) = 0 Then
                                                                    ' Get the view records.
                                                                    cmdViewRecords = Server.CreateObject("ADODB.Command")
                                                                    cmdViewRecords.CommandText = "sp_ASRIntGetLinkViews"
                                                                    cmdViewRecords.CommandType = 4 ' Stored Procedure
                                                                    cmdViewRecords.ActiveConnection = Session("databaseConnection")

                                                                    prmTableID = cmdViewRecords.CreateParameter("tableID", 3, 1)
                                                                    cmdViewRecords.Parameters.Append(prmTableID)
                                                                    prmTableID.value = cleanNumeric(Session("selectionTableID"))

                                                                    prmDfltOrderID = cmdViewRecords.CreateParameter("dfltOrderID", 3, 2) ' 11=integer, 2=output
                                                                    cmdViewRecords.Parameters.Append(prmDfltOrderID)

                                                                    Err.Clear()
                                                                    rstViewRecords = cmdViewRecords.Execute

                                                                    If (Err.Number <> 0) Then
                                                                        sErrorDescription = "The Employee views could not be retrieved." & vbCrLf & FormatError(Err.Description)
                                                                    End If

                                                                    If Len(sErrorDescription) = 0 Then
                                                                        Do While Not rstViewRecords.EOF
                                                        %>
                                                        <option value="<%=rstViewRecords.Fields(0).Value%>"
                                                            <%
                                                                If rstViewRecords.Fields(0).Value = Session("optionLinkViewID") Then
%>
                                                            selected
                                                            <%
                                                            End If

                                                            If rstViewRecords.Fields(0).Value = 0 Then
%>><%=replace(rstViewRecords.Fields(1).Value, "_", " ")%></option>
                                                        <%
                                                            Else
                                                        %>
						                                >'<%=replace(rstViewRecords.Fields(1).Value, "_", " ")%>' view</OPTION>
                                                            <%
                                                            End If

                                                            rstViewRecords.MoveNext()
                                                        Loop
			
                                                        If (rstViewRecords.EOF And rstViewRecords.BOF) Then
                                                            sFailureDescription = "You do not have permission to read the Employee table."
                                                        End If
		
                                                        ' Release the ADO recordset object.
                                                        rstViewRecords.close()
                                                        rstViewRecords = Nothing
	
                                                        ' NB. IMPORTANT ADO NOTE.
                                                        ' When calling a stored procedure which returns a recordset AND has output parameters
                                                        ' you need to close the recordset and set it to nothing before using the output parameters. 
                                                        If Session("optionLinkOrderID") <= 0 Then
                                                            Session("optionLinkOrderID") = cmdViewRecords.Parameters("dfltOrderID").Value
                                                        End If
                                                    End If

                                                    ' Release the ADO command object.
                                                    cmdViewRecords = Nothing
                                                End If
                                                            %>
                                                    </select>
                                                </td>
                                                <td width="10" height="10">
                                                    <input type="button" value="Go" id="btnGoView" name="btnGoView" class="btn"
                                                        onclick="goView()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
                                                </td>
                                                <td height="10">&nbsp;
                                                </td>
                                                <td width="40" height="10">Order :
                                                </td>
                                                <td width="10" height="10">&nbsp;
                                                </td>
                                                <td width="175" height="10">
                                                    <select id="selectOrder" name="selectOrder" class="combo" style="HEIGHT: 10px; WIDTH: 200px">
                                                        <%
                                                                If Len(sErrorDescription) = 0 Then
                                                                    ' Get the order records.
                                                                    cmdOrderRecords = Server.CreateObject("ADODB.Command")
                                                                    cmdOrderRecords.CommandText = "sp_ASRIntGetTableOrders"
                                                                    cmdOrderRecords.CommandType = 4 ' Stored Procedure
                                                                    cmdOrderRecords.ActiveConnection = Session("databaseConnection")

                                                                    prmTableID = cmdOrderRecords.CreateParameter("tableID", 3, 1)
                                                                    cmdOrderRecords.Parameters.Append(prmTableID)
                                                                    prmTableID.value = cleanNumeric(Session("selectionTableID"))

                                                                    prmViewID = cmdOrderRecords.CreateParameter("viewID", 3, 1)
                                                                    cmdOrderRecords.Parameters.Append(prmViewID)
                                                                    prmViewID.value = 0

                                                                    Err.Clear()
                                                                    rstOrderRecords = cmdOrderRecords.Execute

                                                                    If (Err.Number <> 0) Then
                                                                        sErrorDescription = "The order records could not be retrieved." & vbCrLf & FormatError(Err.Description)
                                                                    End If

                                                                    If Len(sErrorDescription) = 0 Then
                                                                        Do While Not rstOrderRecords.EOF
                                                        %>
                                                        <option value="<%=rstOrderRecords.Fields(1).Value%>"
                                                            <%
                                                                If rstOrderRecords.Fields(1).Value = Session("optionLinkOrderID") Then
%>
                                                            selected
                                                            <%
                                                            End If
%>><%=replace(rstOrderRecords.Fields(0).Value, "_", " ")%></option>
                                                        <%
                                                                rstOrderRecords.MoveNext()
                                                            Loop

                                                            ' Release the ADO recordset object.
                                                            rstOrderRecords.close()
                                                            rstOrderRecords = Nothing
                                                        End If
	
                                                        ' Release the ADO command object.
                                                        cmdOrderRecords = Nothing
                                                    End If
                                                        %>
                                                    </select>
                                                </td>
                                                <td width="10" height="10">
                                                    <input type="button" value="Go" id="btnGoOrder" name="btnGoOrder" class="btn"
                                                        onclick="goOrder()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr height="10">
                                    <td height="10">&nbsp;</td>
                                </tr>
                                <tr>
                                    <td>
                                        <object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id="ssOleDBGridSelRecords" name="ssOleDBGridSelRecords" codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px">
                                            <param name="ScrollBars" value="4">
                                            <param name="_Version" value="196617">
                                            <param name="DataMode" value="2">
                                            <param name="Cols" value="0">
                                            <param name="Rows" value="0">
                                            <param name="BorderStyle" value="1">
                                            <param name="RecordSelectors" value="0">
                                            <param name="GroupHeaders" value="0">
                                            <param name="ColumnHeaders" value="-1">
                                            <param name="GroupHeadLines" value="1">
                                            <param name="HeadLines" value="1">
                                            <param name="FieldDelimiter" value="(None)">
                                            <param name="FieldSeparator" value="(Tab)">
                                            <param name="Col.Count" value="0">
                                            <param name="stylesets.count" value="0">
                                            <param name="TagVariant" value="EMPTY">
                                            <param name="UseGroups" value="0">
                                            <param name="HeadFont3D" value="0">
                                            <param name="Font3D" value="0">
                                            <param name="DividerType" value="3">
                                            <param name="DividerStyle" value="1">
                                            <param name="DefColWidth" value="0">
                                            <param name="BeveColorScheme" value="2">
                                            <param name="BevelColorFrame" value="-2147483642">
                                            <param name="BevelColorHighlight" value="-2147483628">
                                            <param name="BevelColorShadow" value="-2147483632">
                                            <param name="BevelColorFace" value="-2147483633">
                                            <param name="CheckBox3D" value="-1">
                                            <param name="AllowAddNew" value="0">
                                            <param name="AllowDelete" value="0">
                                            <param name="AllowUpdate" value="0">
                                            <param name="MultiLine" value="0">
                                            <param name="ActiveCellStyleSet" value="">
                                            <param name="RowSelectionStyle" value="0">
                                            <param name="AllowRowSizing" value="0">
                                            <param name="AllowGroupSizing" value="0">
                                            <param name="AllowColumnSizing" value="-1">
                                            <param name="AllowGroupMoving" value="0">
                                            <param name="AllowColumnMoving" value="0">
                                            <param name="AllowGroupSwapping" value="0">
                                            <param name="AllowColumnSwapping" value="0">
                                            <param name="AllowGroupShrinking" value="0">
                                            <param name="AllowColumnShrinking" value="0">
                                            <param name="AllowDragDrop" value="0">
                                            <param name="UseExactRowCount" value="-1">
                                            <param name="SelectTypeCol" value="0">
                                            <param name="SelectTypeRow" value="3">
                                            <param name="SelectByCell" value="-1">
                                            <param name="BalloonHelp" value="0">
                                            <param name="RowNavigation" value="1">
                                            <param name="CellNavigation" value="0">
                                            <param name="MaxSelectedRows" value="0">
                                            <param name="HeadStyleSet" value="">
                                            <param name="StyleSet" value="">
                                            <param name="ForeColorEven" value="0">
                                            <param name="ForeColorOdd" value="0">
                                            <param name="BackColorEven" value="16777215">
                                            <param name="BackColorOdd" value="16777215">
                                            <param name="Levels" value="1">
                                            <param name="RowHeight" value="503">
                                            <param name="ExtraHeight" value="0">
                                            <param name="ActiveRowStyleSet" value="">
                                            <param name="CaptionAlignment" value="2">
                                            <param name="SplitterPos" value="0">
                                            <param name="SplitterVisible" value="0">
                                            <param name="Columns.Count" value="0">
                                            <param name="UseDefaults" value="-1">
                                            <param name="TabNavigation" value="1">
                                            <param name="_ExtentX" value="17330">
                                            <param name="_ExtentY" value="1323">
                                            <param name="_StockProps" value="79">
                                            <param name="Caption" value="">
                                            <param name="ForeColor" value="0">
                                            <param name="BackColor" value="16777215">
                                            <param name="Enabled" value="-1">
                                            <param name="DataMember" value="">
                                            <param name="Row.Count" value="0">
                                        </object>
                                    </td>
                                </tr>
                            </table>
                            <%	
                                End If
                            %>
                            <input type='hidden' id="txtpicklistSelectionErrorDescription" name="txtpicklistSelectionErrorDescription" value="<%=sErrorDescription%>">
                        </td>
                        <td width="20"></td>
                    </tr>
                    <tr height="10">
                        <td height="10" colspan="3">&nbsp;</td>
                    </tr>
                    <tr height="10">
                        <td width="20"></td>
                        <td height="10">
                            <table width="100%" class="invisible" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td>&nbsp;</td>
                                    <td width="10">
                                        <input id="cmdOK" type="button" value="OK" name="cmdOK" class="btn" style="WIDTH: 80px" width="80"
                                            onclick="makeSelection()"
                                            onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                            onfocus="try{button_onFocus(this);}catch(e){}"
                                            onblur="try{button_onBlur(this);}catch(e){}" />
                                    </td>
                                    <td width="10">&nbsp;</td>
                                    <td width="10">
                                        <input id="cmdCancel" type="button" value="Cancel" class="btn" name="cmdCancel" style="WIDTH: 80px" width="80"
                                            onclick="cancelClick();"
                                            onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                            onfocus="try{button_onFocus(this);}catch(e){}"
                                            onblur="try{button_onBlur(this);}catch(e){}" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td width="20"></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</table>


<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">


<form name="frmPrompt" method="post" action="promptedValues" id="frmPrompt" style="visibility: hidden; display: none">
    <input type="hidden" id="filterID" name="filterID">
</form>

