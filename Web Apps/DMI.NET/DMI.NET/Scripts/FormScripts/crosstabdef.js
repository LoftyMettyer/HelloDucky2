

    function util_def_crosstabs_window_onload() {
        var fOK;
        fOK = true;
        var frmUseful = document.getElementById("frmUseful");
        var sErrMsg = frmUseful.txtErrorDescription.value;
        if (sErrMsg.length > 0) {
            fOK = false;
            OpenHR.MessageBox(sErrMsg, 48, "Cross Tabs");
            //TODO
            //window.parent.location.replace("login");

        }

        if (fOK == true) {
            setGridFont(frmDefinition.grdAccess);
            setFont(frmDefinition.txtHorStart);
            setFont(frmDefinition.txtHorStop);
            setFont(frmDefinition.txtHorStep);
            setFont(frmDefinition.txtVerStart);
            setFont(frmDefinition.txtVerStop);
            setFont(frmDefinition.txtVerStep);
            setFont(frmDefinition.txtPgbStart);
            setFont(frmDefinition.txtPgbStop);
            setFont(frmDefinition.txtPgbStep);

            // Expand the work frame and hide the option frame.
            //window.parent.document.all.item("workframeset").cols = "*, 0";	
            $("#workframe").attr("data-framesource", "UTIL_DEF_CROSSTABS");

            populateBaseTableCombo();

            if (frmUseful.txtAction.value.toUpperCase() == "NEW") {
                frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
                setBaseTable(0);
                changeBaseTable();
                frmUseful.txtSelectedColumnsLoaded.value = 1;
                frmUseful.txtSortLoaded.value = 1;
                frmDefinition.txtDescription.value = "";
                populateColumnCombos();
            } else {
                loadDefinition();
            }

            populateAccessGrid();

            if (frmUseful.txtAction.value.toUpperCase() != "EDIT") {
                frmUseful.txtUtilID.value = 0;
            }

            if (frmUseful.txtAction.value.toUpperCase() == "EDIT") {
                // Get the columns/calcs for the current table selection.
                //var frmGetDataForm = window.parent.frames("dataframe").document.forms("frmGetData");
                var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");

                frmGetDataForm.txtReportBaseTableID.value = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;
            }

            if (frmUseful.txtAction.value.toUpperCase() == "COPY") {
                frmUseful.txtChanged.value = 1;
            }

            refreshTab3Controls();

            displayPage(1);
        }
    }

function util_def_crosstabs_addhandlers() {
    OpenHR.addActiveXHandler("grdAccess", "ComboCloseUp", grdAccess_ComboCloseUp);
    OpenHR.addActiveXHandler("grdAccess", "GotFocus", grdAccess_GotFocus);
    OpenHR.addActiveXHandler("grdAccess", "RowColChange", grdAccess_RowColChange);
    OpenHR.addActiveXHandler("grdAccess", "RowLoaded", grdAccess_RowLoaded);
}

function displayPage(piPageNumber) {
        var iLoop;
        var iCurrentChildCount;
        //window.parent.frames("refreshframe").document.forms("frmRefresh").submit();
        
        if (piPageNumber == 1) {
            div1.style.visibility="visible";
            div1.style.display="block";
            div2.style.visibility="hidden";
            div2.style.display="none";
            div3.style.visibility="hidden";
            div3.style.display="none";

            try {
                frmDefinition.txtName.focus();
            }
            catch(e) {
            }
            refreshTab1Controls();

            button_disable(frmDefinition.btnTab1, true);
            button_disable(frmDefinition.btnTab2, false);
            button_disable(frmDefinition.btnTab3, false);
        }

        if (piPageNumber == 2) {
            // Get the columns/calcs for the current table selection.
            //var frmGetDataForm = window.parent.frames("dataframe").document.forms("frmGetData");
            var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
            

            div1.style.visibility="hidden";
            div1.style.display="none";
            div2.style.display="block";
            div2.style.visibility="visible";
            div3.style.visibility="hidden";
            div3.style.display="none";
		
            refreshTab2Controls();

            if (frmUseful.txtSecondTabShown.value == 0)
            {
                var isEnabled = frmDefinition.txtHorStart.Enabled;
                frmDefinition.txtHorStart.Enabled = true;
                frmDefinition.txtHorStart.focus();
                frmDefinition.txtHorStart.Enabled = isEnabled;
                try {
                    frmDefinition.cboHor.focus();
                }
                catch (e) {}
            }
            frmUseful.txtSecondTabShown.value = 1;
			
            button_disable(frmDefinition.btnTab1, false);
            button_disable(frmDefinition.btnTab2, true);
            button_disable(frmDefinition.btnTab3, false);
        }

        if (piPageNumber == 3) {
            div1.style.visibility="hidden";
            div1.style.display="none";
            div2.style.visibility="hidden";
            div2.style.display="none";
            div3.style.visibility="visible";
            div3.style.display="block";

            refreshTab3Controls();
            try {
                frmDefinition.chkPrintFilter.focus();
            }
            catch(e) {
            }

            button_disable(frmDefinition.btnTab1, false);
            button_disable(frmDefinition.btnTab2, false);
            button_disable(frmDefinition.btnTab3, true);
        }

        // Little dodge to get around a browser bug that
        // does not refresh the display on all controls.
        try
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);
        }
        catch(e) {}

        frmUseful.txtLoading.value = 'N';
    }

function populateBaseTableCombo()
{
    var i;

    //Clear the existing data in the child table combo
    while (frmDefinition.cboBaseTable.options.length > 0) {
        frmDefinition.cboBaseTable.options.remove(0);
    }

    var dataCollection = frmTables.elements;
    if (dataCollection!=null) {
        for (i=0; i<dataCollection.length; i++)  {
            var sControlName = dataCollection.item(i).name;
            sControlTag = sControlName.substr(0, 13);
            if (sControlTag == "txtTableName_") {
                sTableID = sControlName.substr(13);
                var oOption = document.createElement("OPTION");
                frmDefinition.cboBaseTable.options.add(oOption);
                oOption.innerText = dataCollection.item(i).value;
                oOption.value = sTableID;			
            }
        }
    }

}

function cboHor_Change() {
    if (frmUseful.txtLoading.value == 'N') {
        frmUseful.txtCurrentHColID.value = frmDefinition.cboHor.options[frmDefinition.cboHor.selectedIndex].value;
        if (frmUseful.txtCurrentHColID.value == frmUseful.txtCurrentVColID.value) {
            frmUseful.txtCurrentVColID.value = 0;
        }
        if (frmUseful.txtCurrentHColID.value == frmUseful.txtCurrentPColID.value) {
            frmUseful.txtCurrentPColID.value = 0;
        }

        //loadAvailableColumns();
        loadAvailableColumns2(false,true,true,false);
    }
}

function cboVer_Change() {
    if (frmUseful.txtLoading.value == 'N') {
        frmUseful.txtCurrentVColID.value = frmDefinition.cboVer.options[frmDefinition.cboVer.selectedIndex].value;
        if (frmUseful.txtCurrentVColID.value == frmUseful.txtCurrentPColID.value) {
            frmUseful.txtCurrentPColID.value = 0;
        }
        //loadAvailableColumns();
        loadAvailableColumns2(false,false,true,false);
    }
}

function cboPgb_Change() {
    if (frmUseful.txtLoading.value == 'N') {
        frmUseful.txtCurrentPColID.value = frmDefinition.cboPgb.options[frmDefinition.cboPgb.selectedIndex].value;
        //loadAvailableColumns();
        loadAvailableColumns2(false,false,false,false);
    }
    refreshTab2Controls();
}

function cboInt_Change() {
    refreshTab2Controls();
    if (frmDefinition.cboInt.options[frmDefinition.cboInt.selectedIndex].value == 0) {
        frmDefinition.cboIntType.selectedIndex = 1;
    }
}

function populateColumnCombos() {
    //var frmGetDataForm = window.parent.frames("dataframe").document.forms("frmGetData");
    var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
        

    frmGetDataForm.txtAction.value = "LOADREPORTCOLUMNS";
    //frmGetDataForm.txtReportBaseTableID.value = 20;		//frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].Value;
    frmGetDataForm.txtReportBaseTableID.value = frmUseful.txtCurrentBaseTableID.value;
    frmGetDataForm.txtReportParent1TableID.value = 0;
    frmGetDataForm.txtReportParent2TableID.value = 0;
    frmGetDataForm.txtReportChildTableID.value = 0;
    data_refreshData();

    frmUseful.txtLoading.value = 'Y';
}

function loadAvailableColumns() {
    loadAvailableColumns2(true, true, true, true);
    frmUseful.txtLoading.value = 'N';
    //window.parent.frames("menuframe").refreshMenu();
    menu_refreshMenu();

}

function loadAvailableColumns2(horChange, VerChange, PgbChange, IntChange) {

    try {

        frmUseful.txtLoading.value = 'Y';

        var lngHor = 0;
        var lngVer = 0;
        var lngPgb = 0;
        var lngInt = 0;

        var iHNumeric = 0;
        var iHSize = 0;
        var iHDecimals = 0;
        var iVNumeric = 0;
        var iVSize = 0;
        var iVDecimals = 0;
        var iPNumeric = 0;
        var iPSize = 0;
        var iPDecimals = 0;


        if (horChange == true) {
            while (frmDefinition.cboHor.options.length > 0) {
                frmDefinition.cboHor.options.remove(0);
            }
        }

        if (VerChange == true) {
            while (frmDefinition.cboVer.options.length > 0) {
                frmDefinition.cboVer.options.remove(0);
            }
        }
	
        if (PgbChange == true) {
            while (frmDefinition.cboPgb.options.length > 0) {
                frmDefinition.cboPgb.options.remove(0);
            }

            var oOption = document.createElement("OPTION");
            frmDefinition.cboPgb.options.add(oOption);
            oOption.innerText = '<None>';
            oOption.value = 0;			
        }

        if (IntChange == true) {
            while (frmDefinition.cboInt.options.length > 0) {
                frmDefinition.cboInt.options.remove(0);
            }

            var oOption = document.createElement("OPTION");
            frmDefinition.cboInt.options.add(oOption);
            oOption.innerText = '<None>';
            oOption.value = 0;			
        }

        //var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
        var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
            
        var dataCollection = frmUtilDefForm.elements;

        //type
        //tableid
        //colid
        //colname
        //hidden
        //numeric

        if (dataCollection!=null) {
            for (i=0; i<dataCollection.length; i++)  {
                sControlName = dataCollection.item(i).name;
                sControlName = sControlName.substr(0, 10);
                if (sControlName=="txtRepCol_") {

                    sTemp = dataCollection.item(i).value;
                    iIndex = sTemp.indexOf("	");

                    if (iIndex >= 0) {
                        sType = sTemp.substr(0, iIndex);
                        sTemp = sTemp.substr(iIndex + 1);
                        iIndex = sTemp.indexOf("	");
                    }

                    if (sType == "C") {

                        if (iIndex >= 0) {
                            //sTableID = sTemp.substr(0, iIndex);
                            sTemp = sTemp.substr(iIndex + 1);
                            iIndex = sTemp.indexOf("	");
                        }

                        if (iIndex >= 0) {
                            sID = sTemp.substr(0, iIndex);
                            sTemp = sTemp.substr(iIndex + 1);
                            iIndex = sTemp.indexOf("	");
                        }

                        if (iIndex >= 0) {
                            iIndex = sTemp.indexOf(".");
                            sTemp = sTemp.substr(iIndex + 1);
                            iIndex = sTemp.indexOf("	");

                            sColName = sTemp.substr(0, iIndex);
                            sTemp = sTemp.substr(iIndex + 1);
                            iIndex = sTemp.indexOf("	");
                        }

                        if (iIndex >= 0) {
                            sSize = sTemp.substr(0, iIndex);
                            sTemp = sTemp.substr(iIndex + 1);
                            iIndex = sTemp.indexOf("	");

                            sDecimals = sTemp.substr(0, iIndex);
                            sTemp = sTemp.substr(iIndex + 1);
                            iIndex = sTemp.indexOf("	");
                        }

                        if (iIndex >= 0) {

                            sNumeric = sTemp.substr(sTemp.length-1,1);

                            if (horChange == true) {
                                var oOption = document.createElement("OPTION");
                                frmDefinition.cboHor.options.add(oOption);
                                oOption.innerText = sColName;
                                oOption.value = sID;
                            }
                            if (frmUseful.txtCurrentHColID.value == 0) {
                                frmUseful.txtCurrentHColID.value = sID;
                            }
                            if (sID == frmUseful.txtCurrentHColID.value) {
                                lngHor = frmDefinition.cboHor.options.length-1;
                                iHNumeric = parseInt(sNumeric);
                                iHSize = parseInt(sSize);
                                iHDecimals = parseInt(sDecimals);
                            }

                            if (sID != frmUseful.txtCurrentHColID.value) {
                                if (VerChange == true) {
                                    var oOption = document.createElement("OPTION");
                                    frmDefinition.cboVer.options.add(oOption);
                                    oOption.innerText = sColName;
                                    oOption.value = sID;			
                                }
                                if (frmUseful.txtCurrentVColID.value == 0) {
                                    frmUseful.txtCurrentVColID.value = sID;
                                }
                                if (sID == frmUseful.txtCurrentVColID.value) {
                                    lngVer = frmDefinition.cboVer.options.length-1;
                                    iVNumeric = parseInt(sNumeric);
                                    iVSize = parseInt(sSize);
                                    iVDecimals = parseInt(sDecimals);
                                }
                            }

                            if (PgbChange == true) {
                                if ((sID != frmUseful.txtCurrentHColID.value) && (sID != frmUseful.txtCurrentVColID.value)) {
                                    var oOption = document.createElement("OPTION");
                                    frmDefinition.cboPgb.options.add(oOption);
                                    oOption.innerText = sColName;
                                    oOption.value = sID;			
                                    if (sID == frmUseful.txtCurrentPColID.value) {
                                        lngPgb = frmDefinition.cboPgb.options.length-1;
                                    }
                                }
                            }
                            if (sID == frmUseful.txtCurrentPColID.value) {
                                iPNumeric = parseInt(sNumeric);
                                iPSize = parseInt(sSize);
                                iPDecimals = parseInt(sDecimals);
                            }

                            if (sNumeric != "0") {
                                if (IntChange == true) {
                                    var oOption = document.createElement("OPTION");
                                    frmDefinition.cboInt.options.add(oOption);
                                    oOption.innerText = sColName;
                                    oOption.value = sID;			
                                    if (sID == frmUseful.txtCurrentIColID.value) {
                                        lngInt = frmDefinition.cboInt.options.length-1;
                                    }
                                }
                            }

                        }
                    }
                }
            }
        }	

        FormatRange(1, iHNumeric, iHSize, iHDecimals);
        FormatRange(2, iVNumeric, iVSize, iVDecimals);
        FormatRange(3, iPNumeric, iPSize, iPDecimals);

        if (horChange == true) {
            frmDefinition.cboHor.selectedIndex = lngHor;
            frmUseful.txtCurrentHColID.value = frmDefinition.cboHor.options[lngHor].value;
        }

        if (VerChange == true) {
            frmDefinition.cboVer.selectedIndex = lngVer;
            frmUseful.txtCurrentVColID.value = frmDefinition.cboVer.options[lngVer].value;
        }

        if (PgbChange == true) {
            frmDefinition.cboPgb.selectedIndex = lngPgb;
            frmUseful.txtCurrentPColID.value = frmDefinition.cboPgb.options[lngPgb].value;
        }

        if (IntChange == true) {
            frmDefinition.cboInt.selectedIndex = lngInt;
            frmUseful.txtCurrentIColID.value = frmDefinition.cboInt.options[lngInt].value;
        }

    }
    catch (e) {
        alert(e.description);
    }

    frmUseful.txtLoading.value = 'N';

}

function FormatRange(lngIndex, iNumeric, iSize, iDecimals)
{

    try {
        if (iNumeric != 0) {
            sMask = "0";
            sMaxValue = "9";
            iDigitsBeforeDecimal = iSize - iDecimals;
            for (i=1; i<iDigitsBeforeDecimal; i++)  {
                sMask = "#" + sMask;
                sMaxValue = "9" + sMaxValue;
            }
            if (iDecimals > 0) {
                sMask += ".";
                sMaxValue += ".";
                for (i=0; i<iDecimals; i++)  {
                    sMask += "0";
                    sMaxValue += "9";
                }
            }
        }
        else {
            sMask = "#";
            sMaxValue = "9";
        }

        if (lngIndex == 1) {
            EnableControl(frmDefinition.txtHorStart, iNumeric, sMask, sMaxValue);
            EnableControl(frmDefinition.txtHorStop, iNumeric, sMask, sMaxValue);
            EnableControl(frmDefinition.txtHorStep, iNumeric, sMask, sMaxValue);
        }

        if (lngIndex == 2) {
            EnableControl(frmDefinition.txtVerStart, iNumeric, sMask, sMaxValue);
            EnableControl(frmDefinition.txtVerStop, iNumeric, sMask, sMaxValue);
            EnableControl(frmDefinition.txtVerStep, iNumeric, sMask, sMaxValue);
        }

        if (lngIndex == 3) {
            EnableControl(frmDefinition.txtPgbStart, iNumeric, sMask, sMaxValue);
            EnableControl(frmDefinition.txtPgbStop, iNumeric, sMask, sMaxValue);
            EnableControl(frmDefinition.txtPgbStep, iNumeric, sMask, sMaxValue);
        }
    }
    catch(e) {
        alert(e.description);
    }

}

function EnableControl(tempCtl, iNumeric, sMask, sMaxValue)
{
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

    try {
        ctlVal = tempCtl.Value;
        if (ctlVal > parseInt(sMaxValue)) {
            ctlVal = parseInt(sMaxValue);
        }
        if (ctlVal < (parseInt(sMaxValue) * -1)) {
            ctlVal = (parseInt(sMaxValue) * -1);
        }
    }
    catch (e) {
        ctlVal = 0;
    }

    try {
        tempCtl.Format = sMask;
        tempCtl.DisplayFormat = sMask;
        tempCtl.MaxValue = parseInt(sMaxValue);
    }
    catch (e) {
        alert(e.description);
    }


    try {
        if (iNumeric == 0) {
            tempCtl.Value = 0;
        }
        else {
            tempCtl.Value = ctlVal;
        }
		
        if ((iNumeric == 0) || (fViewing == true)) {
            tempCtl.Enabled = false;
            tempCtl.BackColor = 15004669;
            tempCtl.ForeColor = 11375765;
        }
        else {
            tempCtl.Enabled = true;
            tempCtl.BackColor = 15988214; 
            tempCtl.ForeColor = 6697779;
        }			
    }
    catch (e) {
        alert(e.description);
    }
}

function setBaseTable(piTableID) 
{
    var i;
	
    if (piTableID == 0) piTableID = frmUseful.txtPersonnelTableID.value;

    if (piTableID > 0) {
        for (i=0; i<frmDefinition.cboBaseTable.options.length; i++)  {
            if (frmDefinition.cboBaseTable.options(i).value == piTableID) {
                frmDefinition.cboBaseTable.selectedIndex = i;
                frmUseful.txtCurrentBaseTableID.value = piTableID;
                break;
            }		
        }
    }
    else {
        if (frmDefinition.cboBaseTable.options.length > 0) {
            frmDefinition.cboBaseTable.selectedIndex = 0;
            frmUseful.txtCurrentBaseTableID.value = frmDefinition.cboBaseTable.options(0).value;
        }
    }
}

function selectEmailGroup()
{
    var sURL;
	
    frmEmailSelection.EmailSelCurrentID.value = frmDefinition.txtEmailGroupID.value; 

    sURL = "util_emailSelection" +
        "?EmailSelCurrentID=" + frmEmailSelection.EmailSelCurrentID.value;
    openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");
}

function changeBaseTable() 
{
    var i;
    var iAnswer;
    //frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");
    var frmRefresh = OpenHR.getForm("pollframe", "frmHit");
        
    iAnswer = 7;
    if (frmUseful.txtLoading.value == 'N') {
	
        if (frmUseful.txtAction.value.toUpperCase() != "NEW") {

            iAnswer = OpenHR.MessageBox("Warning: Changing the base table will result in all table/column specific aspects of this report definition being cleared. Are you sure you wish to continue?",36,"Cross Tabs");
            if (iAnswer == 7)	{
                // cancel and change back ! (txtcurrentbasetable)
                setBaseTable(frmUseful.txtCurrentBaseTableID.value);
                return;
            }
            else	{
                frmUseful.txtSelectedColumnsLoaded.value = 1;
                frmUseful.txtSortLoaded.value = 1;
                frmUseful.txtChanged.value = 1;
            }
        }
        else {
            frmUseful.txtChanged.value = 1;
        }
    }

    clearBaseTableRecordOptions();
	
    refreshTab1Controls();
    if (frmDefinition.cboBaseTable.options.selectedIndex != -1) {
        frmUseful.txtCurrentBaseTableID.value = frmDefinition.cboBaseTable.options(frmDefinition.cboBaseTable.options.selectedIndex).value;
        frmUseful.txtTablesChanged.value = 1;
    }

    //if (iAnswer != 7) {
    frmUseful.txtCurrentHColID.value = 0;
    frmUseful.txtCurrentVColID.value = 0;
    frmUseful.txtCurrentPColID.value = 0;
    populateColumnCombos();
    //}

}

function refreshTab1Controls()
{
    var fIsForcedHidden;
    var fViewing;
    var fIsNotOwner;
    var fAllAlreadyHidden;
    var fSilent;
	
    fSilent = ((frmUseful.txtAction.value.toUpperCase() == "COPY") &&
        (frmUseful.txtLoading.value == "Y"));
		
    fIsForcedHidden = ((frmSelectionAccess.baseHidden.value == "Y") || 
        (frmSelectionAccess.childHidden.value == "Y") || 
        (frmSelectionAccess.calcsHiddenCount.value > 0));
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());
    fAllAlreadyHidden = AllHiddenAccess(frmDefinition.grdAccess);

    if (fIsForcedHidden == true) {
        if (fAllAlreadyHidden != true) {
            if (fSilent == false) {
                OpenHR.MessageBox("This definition will now be made hidden as it contains a hidden picklist/filter/calculation.", 64);
            }
            ForceAccess(frmDefinition.grdAccess, "HD");
            frmUseful.txtChanged.value = 1;
        }
        else
        {
            if (frmSelectionAccess.forcedHidden.value == "N") {
                //MH20040816 Fault 9049
                //if (fSilent == false) {
                if ((fSilent == false) && (frmUseful.txtLoading.value != "Y")) {
                    OpenHR.MessageBox("The definition access cannot be changed as it contains a hidden picklist/filter/calculation.", 64);
                }
            }
        }
        frmSelectionAccess.forcedHidden.value = "Y";
    }
    else {
        if (frmSelectionAccess.forcedHidden.value == "Y") {
            // No longer forced hidden.
            if (fSilent == false) {
                OpenHR.MessageBox("This definition no longer has to be hidden.", 64);
            }
            frmSelectionAccess.forcedHidden.value = "N";
            frmDefinition.grdAccess.MoveFirst();
            frmDefinition.grdAccess.Col = 1;
        }
    }
    frmDefinition.grdAccess.refresh();

    text_disable(frmDefinition.txtName, (fViewing == true));
    text_disable(frmDefinition.txtDescription, (fViewing == true));
    combo_disable(frmDefinition.cboBaseTable, (fViewing == true));
    button_disable(frmDefinition.cmdBasePicklist, ((frmDefinition.optRecordSelection2.checked == false)
        || (fViewing == true)));
    button_disable(frmDefinition.cmdBaseFilter, ((frmDefinition.optRecordSelection3.checked == false)
        || (fViewing == true)));

    if (frmDefinition.optRecordSelection1.checked == true) {
        checkbox_disable(frmDefinition.chkPrintFilter, true);
        frmDefinition.chkPrintFilter.checked = false;
    }
    else {
        checkbox_disable(frmDefinition.chkPrintFilter, (fViewing == true));
    }

    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
        (fViewing == true)));

    // Little dodge to get around a browser bug that
    // does not refresh the display on all controls.
    try
    {
        window.resizeBy(0,-1);
        window.resizeBy(0,1);
    }
    catch(e) {}

}

function refreshTab2Controls()
{

    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

    combo_disable(frmDefinition.cboHor, (fViewing == true));
    combo_disable(frmDefinition.cboVer, (fViewing == true));
    combo_disable(frmDefinition.cboPgb, (fViewing == true));
    combo_disable(frmDefinition.cboInt, (fViewing == true));

    if (frmDefinition.cboInt.selectedIndex != -1) {
        if (frmDefinition.cboInt.options[frmDefinition.cboInt.selectedIndex].value > 0) {
            combo_disable(frmDefinition.cboIntType, (fViewing == true));
        }
        else {
            combo_disable(frmDefinition.cboIntType, true);
        }
    }

    if (frmDefinition.cboPgb.selectedIndex != -1) {
        if ((frmDefinition.cboPgb.options[frmDefinition.cboPgb.selectedIndex].value == 0) ||
            (frmDefinition.chkPercentage.checked == false)) {
            checkbox_disable(frmDefinition.chkPerPage, true);
            frmDefinition.chkPerPage.checked = false;
        }
        else {
            checkbox_disable(frmDefinition.chkPerPage, (fViewing == true));
        }
    }

    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
        (fViewing == true)));
	
    // Little dodge to get around a browser bug that
    // does not refresh the display on all controls.
    try
    {
        window.resizeBy(0,-1);
        window.resizeBy(0,1);
    }
    catch(e) {}
}


function formatClick(index)
{
    var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

    checkbox_disable(frmDefinition.chkPreview, ((index == 0) || (fViewing == true)));
    frmDefinition.chkPreview.checked = (index != 0);

    frmDefinition.chkDestination0.checked = false;
    frmDefinition.chkDestination1.checked = false;
    frmDefinition.chkDestination2.checked = false;
    frmDefinition.chkDestination3.checked = false;


    if (index == 1) {
        frmDefinition.chkDestination2.checked = true;
        frmDefinition.cboSaveExisting.length = 0;
        frmDefinition.txtFilename.value = '';			
    }
    else {
        frmDefinition.chkDestination0.checked = true;
    }

    frmUseful.txtChanged.value = 1;
    refreshTab3Controls();
}


function refreshTab3Controls()
{
    var i;
    var iCount;
	
    var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

    with (frmDefinition)
    {
        if (optOutputFormat0.checked == true)		//Data Only
        {
            //disable preview opitons
            chkPreview.checked = false;
            checkbox_disable(chkPreview, true);
			
            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));
            if (chkDestination1.checked == true)
            {
                populatePrinters();
                combo_disable(cboPrinterName, (fViewing == true));
            }
            else
            {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName, true);
            }
			
            //disable save options
            chkDestination2.checked = false;
            checkbox_disable(chkDestination2, true);
            combo_disable(cboSaveExisting, true);
            cboSaveExisting.length = 0;
            txtFilename.value = '';
            text_disable(txtFilename, true);
            button_disable(cmdFilename, true);
			
            //disable email options
            chkDestination3.checked = false;
            checkbox_disable(chkDestination3, true);
            text_disable(txtEmailGroup, true);
            txtEmailGroup.value = '';
            txtEmailGroupID.value = 0;
            button_disable(cmdEmailGroup, true);
            text_disable(txtEmailSubject, true);
            text_disable(txtEmailAttachAs, true);
        }
        else if (optOutputFormat1.checked == true)   //CSV File
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));
			
            //disable display on screen options
            chkDestination0.checked = false;
            checkbox_disable(chkDestination0, true);
			
            //disable printer options
            chkDestination1.checked = false;
            checkbox_disable(chkDestination1, true);
            cboPrinterName.length = 0;
            combo_disable(cboPrinterName, true);
						
            //enable-disable save options
            checkbox_disable(chkDestination2, false);
            if (chkDestination2.checked == true)
            {
                populateSaveExisting();
                combo_disable(cboSaveExisting, false);
                text_disable(txtFilename, false);
                button_disable(cmdFilename, false);
            }	
            else
            {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
                text_disable(txtFilename, true);
                txtFilename.value = '';
                button_disable(cmdFilename, true);
            }
			
            //enable-disable email options
            checkbox_disable(chkDestination3, false);
            if (chkDestination3.checked == true)
            {
                text_disable(txtEmailGroup, false);
                text_disable(txtEmailSubject, false);
                button_disable(cmdEmailGroup, false);
                text_disable(txtEmailAttachAs, false);
            }
            else
            {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                button_disable(cmdEmailGroup, true);
                text_disable(txtEmailSubject, true);
                text_disable(txtEmailAttachAs, true);
            }
        }
        else if (optOutputFormat2.checked == true)		//HTML Document
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));
			
            //disable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //disable printer options
            chkDestination1.checked = false;
            checkbox_disable(chkDestination1, true);
            cboPrinterName.length = 0;
            combo_disable(cboPrinterName, true);
						
            //enable-disable save options
            checkbox_disable(chkDestination2, false);
            if (chkDestination2.checked == true)
            {
                populateSaveExisting();
                combo_disable(cboSaveExisting, false);
                text_disable(txtFilename, false);
                button_disable(cmdFilename, false);
            }	
            else
            {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
                text_disable(txtFilename, true);
                txtFilename.value = '';
                button_disable(cmdFilename, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, false);
            if (chkDestination3.checked == true)
            {
                text_disable(txtEmailGroup, false);
                text_disable(txtEmailSubject, false);
                button_disable(cmdEmailGroup, false);
                text_disable(txtEmailAttachAs, false);
            }
            else
            {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                button_disable(cmdEmailGroup, true);
                text_disable(txtEmailSubject, true);
                text_disable(txtEmailAttachAs, true);
            }
        }
        else if (optOutputFormat3.checked == true)		//Word Document
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));
			
            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));
            if (chkDestination1.checked == true)
            {
                populatePrinters();
                combo_disable(cboPrinterName, (fViewing == true));
            }
            else
            {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName, true);
            }
										
            //enable-disable save options
            checkbox_disable(chkDestination2, false);
            if (chkDestination2.checked == true)
            {
                populateSaveExisting();
                combo_disable(cboSaveExisting, false);
                text_disable(txtFilename, false);
                button_disable(cmdFilename, false);
            }	
            else
            {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
                text_disable(txtFilename, true);
                txtFilename.value = '';
                button_disable(cmdFilename, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, false);
            if (chkDestination3.checked == true)
            {
                text_disable(txtEmailGroup, false);
                text_disable(txtEmailSubject, false);
                button_disable(cmdEmailGroup, false);
                text_disable(txtEmailAttachAs, false);
            }
            else
            {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                button_disable(cmdEmailGroup, true);
                text_disable(txtEmailSubject, true);
                text_disable(txtEmailAttachAs, true);
            }
        }
        else if ((optOutputFormat4.checked == true) ||		//Excel Worksheet
            (optOutputFormat5.checked == true) ||
            (optOutputFormat6.checked == true))
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));
			
            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));
            if (chkDestination1.checked == true)
            {
                populatePrinters();
                combo_disable(cboPrinterName, (fViewing == true));
            }
            else
            {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName, true);
            }
										
            //enable-disable save options
            checkbox_disable(chkDestination2, false);
            if (chkDestination2.checked == true)
            {
                populateSaveExisting();
                combo_disable(cboSaveExisting, false);
                text_disable(txtFilename, false);
                button_disable(cmdFilename, false);
            }	
            else
            {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
                text_disable(txtFilename, true);
                txtFilename.value = '';
                button_disable(cmdFilename, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, false);
            if (chkDestination3.checked == true)
            {
                text_disable(txtEmailGroup, false);
                text_disable(txtEmailSubject, false);
                button_disable(cmdEmailGroup, false);
                text_disable(txtEmailAttachAs, false);
            }
            else
            {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                button_disable(cmdEmailGroup, true);
                text_disable(txtEmailSubject, true);
                text_disable(txtEmailAttachAs, true);
            }
        }
            /*else if (optOutputFormat5.checked == true)		//Excel Chart
            {
            }
        else if (optOutputFormat6.checked == true)		//Excel Pivot Table
            {
            }*/
        else
        {
            optOutputFormat0.checked = true;
            chkDestination0.checked=true;
            refreshTab3Controls();
        }
		
        if (txtEmailSubject.disabled)
        {
            txtEmailSubject.value = '';
        }

        if (txtEmailAttachAs.disabled)
        {
            txtEmailAttachAs.value = '';
        }
        else
        {
            if (txtEmailAttachAs.value == '') {
                if (txtFilename.value != '') {
                    sAttachmentName = new String(txtFilename.value);
                    txtEmailAttachAs.value = sAttachmentName.substr(sAttachmentName.lastIndexOf("\\")+1);
                }
            }
        }

        if (cmdFilename.disabled == true) {
            txtFilename.value = "";
        }

    }

    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
        (fViewing == true)));

    // Little dodge to get around a browser bug that
    // does not refresh the display on all controls.
    try
    {
        window.resizeBy(0,-1);
        window.resizeBy(0,1);
    }
    catch(e) {}
}

function saveFile() {
    dialog.CancelError = true;
    dialog.CancelError = true;
    dialog.DialogTitle = "Output Document";
    dialog.Flags = 2621444;

    if (frmDefinition.optOutputFormat1.checked == true) {
        //CSV
        dialog.Filter = "Comma Separated Values (*.csv)|*.csv";
    }

    else if (frmDefinition.optOutputFormat2.checked == true) {
        //HTML
        dialog.Filter = "HTML Document (*.htm)|*.htm";
    }

    else if (frmDefinition.optOutputFormat3.checked == true) {
        //WORD
        //dialog.Filter = "Word Document (*.doc)|*.doc";
        dialog.Filter = frmDefinition.txtWordFormats.value;
        dialog.FilterIndex = frmDefinition.txtWordFormatDefaultIndex.value;
    }

    else {
        //EXCEL
        //dialog.Filter = "Excel Workbook (*.xls)|*.xls";
        dialog.Filter = frmDefinition.txtExcelFormats.value;
        dialog.FilterIndex = frmDefinition.txtExcelFormatDefaultIndex.value;
    }


    if (frmDefinition.txtFilename.value.length == 0) {
        sKey = new String("documentspath_");
        sKey = sKey.concat(frmDefinition.txtDatabase.value);
        //TODO
        //sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
        dialog.InitDir = sPath;
    }
    else {
        dialog.FileName = frmDefinition.txtFilename.value;
    }


    try {
        dialog.ShowSave();

        if (dialog.FileName.length > 256) {
            OpenHR.MessageBox("Path and file name must not exceed 256 characters in length");
            return;
        }

        frmDefinition.txtFilename.value = dialog.FileName;

    }
    catch(e) {
    }

}

function changeBaseTableRecordOptions()
{
    frmDefinition.txtBasePicklist.value = '';
    frmDefinition.txtBasePicklistID.value = 0;

    frmDefinition.txtBaseFilter.value = '';
    frmDefinition.txtBaseFilterID.value = 0;

    frmSelectionAccess.baseHidden.value = "N";

    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}

function clearBaseTableRecordOptions()
{
    frmDefinition.optRecordSelection1.checked = true;
	
    button_disable(frmDefinition.cmdBasePicklist, true);
    frmDefinition.txtBasePicklist.value = '';
    frmDefinition.txtBasePicklistID.value = 0;
	
    button_disable(frmDefinition.cmdBaseFilter, true);
    frmDefinition.txtBaseFilter.value = '';
    frmDefinition.txtBaseFilterID.value = 0;
	
    frmDefinition.chkPrintFilter.checked = false;

    frmSelectionAccess.baseHidden.value = "N";
}

function selectRecordOption(psTable, psType)
{	
    var sURL;
	
    if (psTable == 'base') {
        iTableID = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;


		
        if (psType == 'picklist') {
            iCurrentID = frmDefinition.txtBasePicklistID.value;
        }
        else {
            iCurrentID = frmDefinition.txtBaseFilterID.value;
        }
    }

    frmRecordSelection.recSelTable.value = psTable;
    frmRecordSelection.recSelType.value = psType;
    frmRecordSelection.recSelTableID.value = iTableID;
    frmRecordSelection.recSelCurrentID.value = iCurrentID; 
	
    var strDefOwner = new String(frmDefinition.txtOwner.value);
    var strCurrentUser = new String(frmUseful.txtUserName.value);
	
    strDefOwner = strDefOwner.toLowerCase();
    strCurrentUser = strCurrentUser.toLowerCase();
	
    if (strDefOwner == strCurrentUser) 
    {
        frmRecordSelection.recSelDefOwner.value = '1';
    }
    else
    {
        frmRecordSelection.recSelDefOwner.value = '0';
    }
    frmRecordSelection.recSelDefType.value = "Cross Tabs";
	
    sURL = "util_recordSelection" +
        "?recSelType=" + escape(frmRecordSelection.recSelType.value) +
        "&recSelTableID=" + escape(frmRecordSelection.recSelTableID.value) + 
        "&recSelCurrentID=" + escape(frmRecordSelection.recSelCurrentID.value) +
        "&recSelTable=" + escape(frmRecordSelection.recSelTable.value) +
        "&recSelDefOwner=" + escape(frmRecordSelection.recSelDefOwner.value) +
        "&recSelDefType=" + escape(frmRecordSelection.recSelDefType.value);
    openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");

    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}

function submitDefinition()
{
    var i;
    var iIndex;
    var sColumnID;
    var sType;
    var sURL;
	
    if (validateTab1() == false) { menu_refreshMenu(); return;}
    if (validateTab2() == false) { menu_refreshMenu(); return;}
    if (validateTab3() == false) { menu_refreshMenu(); return;}
    if (populateSendForm() == false)  {menu_refreshMenu(); return;}

    // Now create the validate popup to check that any filters/calcs
    // etc havent been deleted, or made hidden etc.		

    // first populate the validate fields
    frmValidate.validateBaseFilter.value = frmDefinition.txtBaseFilterID.value;
    frmValidate.validateBasePicklist.value = frmDefinition.txtBasePicklistID.value;
    frmValidate.validateEmailGroup.value = frmDefinition.txtEmailGroupID.value;
    frmValidate.validateName.value = frmDefinition.txtName.value;

    if(frmUseful.txtAction.value.toUpperCase() == "EDIT"){
        frmValidate.validateTimestamp.value = frmOriginalDefinition.txtDefn_Timestamp.value;
        frmValidate.validateUtilID.value = frmUseful.txtUtilID.value;
    }
    else {
        frmValidate.validateTimestamp.value = 0;
        frmValidate.validateUtilID.value = 0;
    }
	
    sHiddenGroups = HiddenGroups(frmDefinition.grdAccess);
    frmValidate.validateHiddenGroups.value = sHiddenGroups;

    sURL = "dialog" +
        "?validateBaseFilter=" +  escape(frmValidate.validateBaseFilter.value) +
        "&validateBasePicklist=" + escape(frmValidate.validateBasePicklist.value) +
        "&validateEmailGroup=" + escape(frmValidate.validateEmailGroup.value) +
        "&validateCalcs=" + escape(frmValidate.validateCalcs.value) +
        "&validateHiddenGroups=" + escape(frmValidate.validateHiddenGroups.value) +
        "&validateName=" + escape(frmValidate.validateName.value) +
        "&validateTimestamp=" + escape(frmValidate.validateTimestamp.value) +
        "&validateUtilID=" + frmValidate.validateUtilID.value +
        "&destination=util_validate_crosstab";
    openDialog(sURL, (screen.width)/2,(screen.height)/3,"no", "no");
}

function cancelClick()
{
    if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
        (definitionChanged() == false)) {
        //todo
        //window.location.href="defsel";
        return;
    }

    answer = OpenHR.MessageBox("You have changed the current definition. Save changes ?",3,"Cross Tabs");
    if (answer == 7) {
        // No
        //todo
        //window.location.href="defsel";
        return (false);
    }
    if (answer == 6) {
        // Yes
        okClick();
    }
}

function okClick()
{
    //window.parent.frames("menuframe").disableMenu();
    disableMenu();
	
    var sAttachmentName = new String(frmDefinition.txtEmailAttachAs.value);
    if ((sAttachmentName.indexOf("/") != -1) || 
        (sAttachmentName.indexOf(":") != -1) || 
        (sAttachmentName.indexOf("?") != -1) || 
        (sAttachmentName.indexOf(String.fromCharCode(34)) != -1) || 
        (sAttachmentName.indexOf("<") != -1) || 
        (sAttachmentName.indexOf(">") != -1) || 
        (sAttachmentName.indexOf("|") != -1) || 
        (sAttachmentName.indexOf("\\") != -1) || 
        (sAttachmentName.indexOf("*") != -1)) {
        OpenHR.MessageBox("The attachment file name can not contain any of the following characters:\n/ : ? " + String.fromCharCode(34) + " < > | \\ *",48,"Cross Tabs");
        return;
    }

    frmSend.txtSend_reaction.value = "CROSSTABS";
    submitDefinition();
}

function saveChanges(psAction, pfPrompt, pfTBOverride)
{
    if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
        (definitionChanged() == false)) {
        return 7; //No to saving the changes, as none have been made.
    }

    answer = OpenHR.MessageBox("You have changed the current definition. Save changes ?",3,"Cross Tabs");
    if (answer == 7) {
        // No
        return 7;
    }
    if (answer == 6) {
        // Yes
        okClick();
    }

    return 2; //Cancel.
}

function definitionChanged()
{
    if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
        return false;
    }
	
    if (frmUseful.txtChanged.value == 1) {
        return true;
    }
    else {
        if (frmUseful.txtAction.value.toUpperCase() != "NEW") {
            // Compare the tab 1 controls with the original values.
            if (frmDefinition.txtName.value != frmOriginalDefinition.txtDefn_Name.value) {
                return true;
            }
		
            if (frmDefinition.txtDescription.value != frmOriginalDefinition.txtDefn_Description.value) {
                return true;
            }
		
            if (frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value != frmOriginalDefinition.txtDefn_BaseTableID.value) {
                return true;
            }

            if (frmOriginalDefinition.txtDefn_PicklistID.value > 0) {
                if (frmDefinition.optRecordSelection2.checked == false) {
                    return true;
                }
                else {
                    if (frmDefinition.txtBasePicklistID.value != frmOriginalDefinition.txtDefn_PicklistID.value) {
                        return true;
                    }
                }				
            }
            else {
                if (frmOriginalDefinition.txtDefn_FilterID.value > 0) {
                    if (frmDefinition.optRecordSelection3.checked == false) {
                        return true;
                    }
                    else {
                        if (frmDefinition.txtBaseFilterID.value != frmOriginalDefinition.txtDefn_FilterID.value) {
                            return true;
                        }
                    }				
                }
                else {
                    if (frmDefinition.optRecordSelection1.checked == false) {
                        return true;
                    }
                }
            }

            if ((frmUseful.txtCurrentHColID.value != frmOriginalDefinition.txtDefn_HColID.value) ||
                (frmUseful.txtCurrentVColID.value != frmOriginalDefinition.txtDefn_VColID.value) ||
                (frmUseful.txtCurrentPColID.value != frmOriginalDefinition.txtDefn_PColID.value) ||
                (frmUseful.txtCurrentIColID.value != frmOriginalDefinition.txtDefn_IColID.value)) {
                return true;
            }

            if ((frmDefinition.txtHorStart.Value != frmOriginalDefinition.txtDefn_HStart.value) ||
                (frmDefinition.txtHorStop.Value != frmOriginalDefinition.txtDefn_HStop.value) ||
                (frmDefinition.txtHorStep.Value != frmOriginalDefinition.txtDefn_HStep.value) ||
                (frmDefinition.txtVerStart.Value != frmOriginalDefinition.txtDefn_VStart.value) ||
                (frmDefinition.txtVerStop.Value != frmOriginalDefinition.txtDefn_VStop.value) ||
                (frmDefinition.txtVerStep.Value != frmOriginalDefinition.txtDefn_VStep.value) ||
                (frmDefinition.txtPgbStart.Value != frmOriginalDefinition.txtDefn_PStart.value) ||
                (frmDefinition.txtPgbStop.Value != frmOriginalDefinition.txtDefn_PStop.value) ||
                (frmDefinition.txtPgbStep.Value != frmOriginalDefinition.txtDefn_PStep.value)) {
                return true;
            }




            // Compare the tab 3 controls with the original values.
            if (frmDefinition.chkPreview.checked.toString().toUpperCase() != frmOriginalDefinition.txtDefn_OutputPreview.value.toUpperCase()){
                return true;
            }

            if (frmDefinition.chkDestination0.checked.toString().toUpperCase() != frmOriginalDefinition.txtDefn_OutputScreen.value.toUpperCase()) {
                return true;
            }

            if (frmDefinition.chkDestination1.checked.toString().toUpperCase() != frmOriginalDefinition.txtDefn_OutputPrinter.value.toUpperCase()) {
                return true;
            }

            if (frmDefinition.chkDestination2.checked.toString().toUpperCase() != frmOriginalDefinition.txtDefn_OutputSave.value.toUpperCase()) {
                return true;
            }

            if (frmDefinition.chkDestination3.checked.toString().toUpperCase() != frmOriginalDefinition.txtDefn_OutputEmail.value.toUpperCase()) {
                return true;
            }

            with (frmDefinition.cboPrinterName) {
                if (options.selectedIndex > -1) {
                    if (options(options.selectedIndex).innerText != frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
                        return true;
                    }
                }
            }

            with (frmDefinition.cboSaveExisting) {
                if (options.selectedIndex > -1) {
                    if (options(options.selectedIndex).value != frmOriginalDefinition.txtDefn_OutputSaveExisting.value) {
                        return true;
                    }
                }
            }

            if (frmDefinition.txtEmailGroup.value != frmOriginalDefinition.txtDefn_OutputEmailAddrName.value) {
                return true;
            }

            if (frmDefinition.txtEmailSubject.value != frmOriginalDefinition.txtDefn_OutputEmailSubject.value) {
                return true;
            }

            if (frmDefinition.txtEmailAttachAs.value != frmOriginalDefinition.txtDefn_OutputEmailAttachAs.value) {
                return true;
            }

            if (frmDefinition.txtFilename.value != frmOriginalDefinition.txtDefn_OutputFilename.value) {
                return true;
            }
        }
		
        return false;
    }
}

function getTableName(piTableID)
{
    var i;
    var sTableName = new String("");
	
    sReqdControlName = new String("txtTableName_");
    sReqdControlName = sReqdControlName.concat(piTableID);

    var dataCollection = frmTables.elements;
    if (dataCollection!=null) {
        for (i=0; i<dataCollection.length; i++)  {
            sControlName = dataCollection.item(i).name;
					
            if (sControlName == sReqdControlName) {
                sTableName = dataCollection.item(i).value;
                return sTableName;
            }
        }
    }	

    return sTableName;
}

function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll)
{
    dlgwinprops = "center:yes;" +
        "dialogHeight:" + pHeight + "px;" +
        "dialogWidth:" + pWidth + "px;" +
        "help:no;" +
        "resizable:" + psResizable + ";" +
        "scroll:" + psScroll + ";" +
        "status:no;";
    window.showModalDialog(pDestination, self, dlgwinprops);
}

function validateTab1()
{
    // check name has been entered
    if (frmDefinition.txtName.value == '') {
        OpenHR.MessageBox("You must enter a name for this definition.",48,"Cross Tabs");
        displayPage(1);
        return (false);
    }
      
    // check base picklist
    if ((frmDefinition.optRecordSelection2.checked == true) &&
        (frmDefinition.txtBasePicklistID.value == 0)) {
        OpenHR.MessageBox("You must select a picklist for the base table.",48,"Cross Tabs");
        displayPage(1);
        return (false);
    }

    // check base filter
    if ((frmDefinition.optRecordSelection3.checked == true) &&
        (frmDefinition.txtBaseFilterID.value == 0)) {
        OpenHR.MessageBox("You must select a filter for the base table.",48,"Cross Tabs");
        displayPage(1);
        return (false);
    }

    return (true);
}

function validateTab2()
{

    if ((frmDefinition.txtHorStart.Value != 0) || (frmDefinition.txtHorStop.Value != 0)) {
        if (frmDefinition.txtHorStop.Value <= frmDefinition.txtHorStart.Value) {
            OpenHR.MessageBox("Horizontal stop value must be greater than Horizontal start value",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
        if (frmDefinition.txtHorStep.Value <= 0) {
            OpenHR.MessageBox("Horizontal increment must be greater than zero",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
        if (((frmDefinition.txtHorStop.Value - frmDefinition.txtHorStart.Value) / frmDefinition.txtHorStep.Value) > 32768) {
            OpenHR.MessageBox("Maximum number of steps between start, stop and increment value for the Horizontal Range\nhas been exceeded. You must either increase the increment value or decrease the stop value.",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
    }

    if ((frmDefinition.txtVerStart.Value != 0) || (frmDefinition.txtVerStop.Value != 0)) {
        if (frmDefinition.txtVerStop.Value <= frmDefinition.txtVerStart.Value) {
            OpenHR.MessageBox("Vertical stop value must be greater than Vertical start value",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
        if (frmDefinition.txtVerStep.Value <= 0) {
            OpenHR.MessageBox("Vertical increment must be greater than zero",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
        if (((frmDefinition.txtVerStop.Value - frmDefinition.txtVerStart.Value) / frmDefinition.txtVerStep.Value) > 32768) {
            OpenHR.MessageBox("Maximum number of steps between start, stop and increment value for the Vertical Range\nhas been exceeded. You must either increase the increment value or decrease the stop value.",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
    }

    if ((frmDefinition.txtPgbStart.Value != 0) || (frmDefinition.txtPgbStop.Value != 0)) {
        if (frmDefinition.txtPgbStop.Value <= frmDefinition.txtPgbStart.Value) {
            OpenHR.MessageBox("Page Break stop value must be greater than Page Break start value",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
        if (frmDefinition.txtPgbStep.Value <= 0) {
            OpenHR.MessageBox("Page Break increment must be greater than zero",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
        if (((frmDefinition.txtPgbStop.Value - frmDefinition.txtPgbStart.Value) / frmDefinition.txtPgbStep.Value) > 32768) {
            OpenHR.MessageBox("Maximum number of steps between start, stop and increment value for the Page Break Range\nhas been exceeded. You must either increase the increment value or decrease the stop value.",48,"Cross Tabs");
            displayPage(2);
            return (false);
        }
    }
	
    return (true);
}  

function validateTab3()
{
    var sErrMsg;
	
    sErrMsg = "";
	
    if (!frmDefinition.chkDestination0.checked 
        && !frmDefinition.chkDestination1.checked 
        && !frmDefinition.chkDestination2.checked 
        && !frmDefinition.chkDestination3.checked)
    {
        sErrMsg = "You must select a destination";
    }

    if ((frmDefinition.txtFilename.value == "") &&
        (frmDefinition.cmdFilename.disabled == false)) {
        sErrMsg = "You must enter a file name";
    }

    if ((frmDefinition.txtEmailGroup.value == "") &&
        (frmDefinition.cmdEmailGroup.disabled == false)) {
        sErrMsg = "You must select an email group";
    }

    if ((frmDefinition.chkDestination3.checked)
        && (frmDefinition.txtEmailAttachAs.value == '')) {
        sErrMsg = "You must enter an email attachment file name.";
    }

    if (frmDefinition.chkDestination3.checked &&
        (frmDefinition.optOutputFormat3.checked || frmDefinition.optOutputFormat4.checked || frmDefinition.optOutputFormat5.checked || frmDefinition.optOutputFormat6.checked) &&
        frmDefinition.txtEmailAttachAs.value.match(/.html$/)) {
        sErrMsg = "You cannot email html output from word or excel.";
    }

    if (sErrMsg.length > 0) 
    {    
        OpenHR.MessageBox(sErrMsg,48,"Cross Tabs");
        displayPage(5);
        return (false);
    }
	
    try 
    {
        var frmRefresh = OpenHR.getForm("workframe", "frmRefresh");
        var testDataCollection = frmRefresh.elements;
        var iDummy = testDataCollection.txtDummy.value;
        OpenHR.submit(frmRefresh);
    }
    catch(e) 
    {
    }
		
    return (true);
}

function populateSendForm()
{
    var i;
    var iIndex;
    var sControlName;
    var iNum;
    var varBookmark;
    var iLoop;
    var sAccess;
	
    // Copy all the header information to frmSend
    frmSend.txtSend_ID.value = frmUseful.txtUtilID.value;
    frmSend.txtSend_name.value = frmDefinition.txtName.value;
    frmSend.txtSend_description.value = frmDefinition.txtDescription.value;
    frmSend.txtSend_userName.value = frmDefinition.txtOwner.value;
	
    sAccess = "";
    frmDefinition.grdAccess.update();
    for(iLoop = 1; iLoop <= (frmDefinition.grdAccess.Rows - 1); iLoop++) {
        varBookmark = frmDefinition.grdAccess.AddItemBookmark(iLoop);
		
        sAccess = sAccess + 
            frmDefinition.grdAccess.Columns("GroupName").CellText(varBookmark) + 
            "	" +
            AccessCode(frmDefinition.grdAccess.Columns("Access").CellText(varBookmark)) + 
            "	";
    }
    frmSend.txtSend_access.value = sAccess;
  
    frmSend.txtSend_baseTable.value = frmDefinition.cboBaseTable.options(frmDefinition.cboBaseTable.options.selectedIndex).value;

    frmSend.txtSend_allRecords.value = "0";
    frmSend.txtSend_picklist.value = "0";
    frmSend.txtSend_filter.value = "0";
    if (frmDefinition.optRecordSelection1.checked == true) {
        frmSend.txtSend_allRecords.value = "1";
    }
    if (frmDefinition.optRecordSelection2.checked == true) {
        frmSend.txtSend_picklist.value = frmDefinition.txtBasePicklistID.value;
    }
    if (frmDefinition.optRecordSelection3.checked == true) {
        frmSend.txtSend_filter.value = frmDefinition.txtBaseFilterID.value;
    }


    frmSend.txtSend_HColID.value = frmDefinition.cboHor.options(frmDefinition.cboHor.options.selectedIndex).value;
    frmSend.txtSend_HStart.value = frmDefinition.txtHorStart.value;
    frmSend.txtSend_HStop.value = frmDefinition.txtHorStop.value;
    frmSend.txtSend_HStep.value = frmDefinition.txtHorStep.value;
    frmSend.txtSend_VColID.value = frmDefinition.cboVer.options(frmDefinition.cboVer.options.selectedIndex).value;
    frmSend.txtSend_VStart.value = frmDefinition.txtVerStart.value;
    frmSend.txtSend_VStop.value = frmDefinition.txtVerStop.value;
    frmSend.txtSend_VStep.value = frmDefinition.txtVerStep.value;
    frmSend.txtSend_PColID.value = frmDefinition.cboPgb.options(frmDefinition.cboPgb.options.selectedIndex).value;
    frmSend.txtSend_PStart.value = frmDefinition.txtPgbStart.value;
    frmSend.txtSend_PStop.value = frmDefinition.txtPgbStop.value;
    frmSend.txtSend_PStep.value = frmDefinition.txtPgbStep.value;
    frmSend.txtSend_IType.value = frmDefinition.cboIntType.options(frmDefinition.cboIntType.options.selectedIndex).value;
    frmSend.txtSend_IColID.value = frmDefinition.cboInt.options(frmDefinition.cboInt.options.selectedIndex).value;


    if (frmDefinition.chkPercentage.checked == true) {
        frmSend.txtSend_Percentage.value = '1';
    }
    else {
        frmSend.txtSend_Percentage.value = '0';
    }

    if (frmDefinition.chkPerPage.checked == true) {
        frmSend.txtSend_PerPage.value = '1';
    }
    else {
        frmSend.txtSend_PerPage.value = '0';
    }

    if (frmDefinition.chkSuppress.checked == true) {
        frmSend.txtSend_Suppress.value = '1';
    }
    else {
        frmSend.txtSend_Suppress.value = '0';
    }

    if (frmDefinition.chkUse1000.checked == true) {
        frmSend.txtSend_Use1000Separator.value = '1';
    }
    else {
        frmSend.txtSend_Use1000Separator.value = '0';
    }

    if (frmDefinition.chkPrintFilter.checked == true) {
        frmSend.txtSend_PrintFilter.value = '1';
    }
    else {
        frmSend.txtSend_PrintFilter.value = '0';
    }


    if (frmDefinition.chkPreview.checked == true)
    {
        frmSend.txtSend_OutputPreview.value = 1;
    }
    else
    {
        frmSend.txtSend_OutputPreview.value = 0;
    }
	
    frmSend.txtSend_OutputFormat.value = 0;
    if (frmDefinition.optOutputFormat1.checked)	frmSend.txtSend_OutputFormat.value = 1;
    if (frmDefinition.optOutputFormat2.checked)	frmSend.txtSend_OutputFormat.value = 2;
    if (frmDefinition.optOutputFormat3.checked)	frmSend.txtSend_OutputFormat.value = 3;
    if (frmDefinition.optOutputFormat4.checked)	frmSend.txtSend_OutputFormat.value = 4;
    if (frmDefinition.optOutputFormat5.checked)	frmSend.txtSend_OutputFormat.value = 5;
    if (frmDefinition.optOutputFormat6.checked)	frmSend.txtSend_OutputFormat.value = 6;

    if (frmDefinition.chkDestination0.checked == true)
    {
        frmSend.txtSend_OutputScreen.value = 1;
    }
    else
    {
        frmSend.txtSend_OutputScreen.value = 0;
    }
	
    if (frmDefinition.chkDestination1.checked == true)
    {
        frmSend.txtSend_OutputPrinter.value = 1;
        frmSend.txtSend_OutputPrinterName.value = frmDefinition.cboPrinterName.options[frmDefinition.cboPrinterName.selectedIndex].innerText;
    }
    else
    {
        frmSend.txtSend_OutputPrinter.value = 0;
        frmSend.txtSend_OutputPrinterName.value = '';
    }
	
    if (frmDefinition.chkDestination2.checked == true)
    {
        frmSend.txtSend_OutputSave.value = 1;
        frmSend.txtSend_OutputSaveExisting.value = frmDefinition.cboSaveExisting.options[frmDefinition.cboSaveExisting.selectedIndex].value;
    }
    else
    {
        frmSend.txtSend_OutputSave.value = 0;
        frmSend.txtSend_OutputSaveExisting.value = 0;
    }
	
    if (frmDefinition.chkDestination3.checked == true)
    {
        frmSend.txtSend_OutputEmail.value = 1;
        frmSend.txtSend_OutputEmailAddr.value = frmDefinition.txtEmailGroupID.value;
        frmSend.txtSend_OutputEmailSubject.value = frmDefinition.txtEmailSubject.value;
        frmSend.txtSend_OutputEmailAttachAs.value = frmDefinition.txtEmailAttachAs.value;
    }
    else
    {
        frmSend.txtSend_OutputEmail.value = 0;
        frmSend.txtSend_OutputEmailAddr.value = 0;
        frmSend.txtSend_OutputEmailSubject.value = '';
        frmSend.txtSend_OutputEmailAttachAs.value = '';
    }
		
    frmSend.txtSend_OutputFilename.value = frmDefinition.txtFilename.value;

}


//function loadAvailableColumns()
//{
//	var i;
//	var sSelectedIDs;
//	var sTemp;
//	var iIndex;
//	var sType;
//	var sID;
	
//	sSelectedIDs = selectedIDs();

//	frmDefinition.ssOleDBGridAvailableColumns.RemoveAll();

//	var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
//	var dataCollection = frmUtilDefForm.elements;

//	if (dataCollection!=null) {
//		for (i=0; i<dataCollection.length; i++)  {
//		  sControlName = dataCollection.item(i).name;
//			sControlName = sControlName.substr(0, 10);
//			if (sControlName=="txtRepCol_") {
//				sTemp = dataCollection.item(i).value;
//				iIndex = sTemp.indexOf("	");
//				if (iIndex >= 0) {
//					sType = sTemp.substr(0, iIndex);
//					sTemp = sTemp.substr(iIndex + 1);

//					iIndex = sTemp.indexOf("	");
//					if (iIndex >= 0) {
//						sTemp = sTemp.substr(iIndex + 1);
						
//						iIndex = sTemp.indexOf("	");
//						if (iIndex >= 0) {
//							sID = sTemp.substr(0, iIndex);
						
//							sTemp = "	" + sType + sID + "	";
//							iIndex = sSelectedIDs.indexOf(sTemp);
//							if (iIndex < 0) {
//								frmDefinition.ssOleDBGridAvailableColumns.AddItem(dataCollection.item(i).value);
//							}
//						}
//					}
//				}
//			}
//		}
//	}	

//	refreshTab2Controls();		  
	
// Get menu to refresh the menu.
//	window.parent.frames("menuframe").refreshMenu();		  
//}

function loadDefinition()
{
    frmDefinition.txtName.value = frmOriginalDefinition.txtDefn_Name.value;

    if((frmUseful.txtAction.value.toUpperCase() == "EDIT") ||
        (frmUseful.txtAction.value.toUpperCase() == "VIEW")) {
        frmDefinition.txtOwner.value = frmOriginalDefinition.txtDefn_Owner.value;
    }
    else {
        frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
    }

    frmDefinition.txtDescription.value= frmOriginalDefinition.txtDefn_Description.value;

    setBaseTable(frmOriginalDefinition.txtDefn_BaseTableID.value);
    changeBaseTable();		

    // Set the basic record selection.
    fRecordOptionSet = false;

    if (frmOriginalDefinition.txtDefn_PicklistID.value > 0) {
        button_disable(frmDefinition.cmdBasePicklist, false);
        frmDefinition.optRecordSelection2.checked = true;
        frmDefinition.txtBasePicklistID.value = frmOriginalDefinition.txtDefn_PicklistID.value;
        frmDefinition.txtBasePicklist.value = frmOriginalDefinition.txtDefn_PicklistName.value;
        fRecordOptionSet = true;
    }
    else {
        if (frmOriginalDefinition.txtDefn_FilterID.value > 0) {
            button_disable(frmDefinition.cmdBaseFilter, false);
            frmDefinition.optRecordSelection3.checked = true;
            frmDefinition.txtBaseFilterID.value = frmOriginalDefinition.txtDefn_FilterID.value;
            frmDefinition.txtBaseFilter.value = frmOriginalDefinition.txtDefn_FilterName.value;
            fRecordOptionSet = true;			
        }
    }
    if (fRecordOptionSet == false) {
        frmDefinition.optRecordSelection1.checked = true;
    }

    if ((frmOriginalDefinition.txtDefn_PicklistHidden.value.toUpperCase() == "TRUE") ||
        (frmOriginalDefinition.txtDefn_FilterHidden.value.toUpperCase() == "TRUE")) {
        frmSelectionAccess.baseHidden.value = "Y";
    }
    frmSelectionAccess.calcsHiddenCount.value = frmOriginalDefinition.txtDefn_HiddenCalcCount.value;

    frmUseful.txtCurrentHColID.value = frmOriginalDefinition.txtDefn_HColID.value;
    frmUseful.txtCurrentVColID.value = frmOriginalDefinition.txtDefn_VColID.value;
    frmUseful.txtCurrentPColID.value = frmOriginalDefinition.txtDefn_PColID.value;
    frmUseful.txtCurrentIColID.value = frmOriginalDefinition.txtDefn_IColID.value;

    populateColumnCombos();

    for (i=0; i<frmDefinition.cboIntType.options.length; i++)  {
        if (frmDefinition.cboIntType.options(i).value == frmOriginalDefinition.txtDefn_IType.value) {
            frmDefinition.cboIntType.selectedIndex = i;
            break;
        }		
    }

    frmDefinition.chkPercentage.checked = (frmOriginalDefinition.txtDefn_Percentage.value != "False");
    frmDefinition.chkPerPage.checked = (frmOriginalDefinition.txtDefn_PerPage.value != "False");
    frmDefinition.chkSuppress.checked = (frmOriginalDefinition.txtDefn_Suppress.value != "False");
    frmDefinition.chkUse1000.checked = (frmOriginalDefinition.txtDefn_Use1000.value != "False");

    // Print Filter Header ?
    frmDefinition.chkPrintFilter.checked = ((frmOriginalDefinition.txtDefn_PrintFilter.value != "False") &&
        ((frmOriginalDefinition.txtDefn_FilterID.value > 0) || (frmOriginalDefinition.txtDefn_PicklistID.value > 0)));


    frmDefinition.chkPreview.checked = (frmOriginalDefinition.txtDefn_OutputPreview.value != "False");
	
    if (frmOriginalDefinition.txtDefn_OutputFormat.value == 0)
    {
        frmDefinition.optOutputFormat0.checked = true;
    }
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 1)
    {
        frmDefinition.optOutputFormat1.checked = true;
    }
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 2)
    {
        frmDefinition.optOutputFormat2.checked = true;
    }
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 3)
    {
        frmDefinition.optOutputFormat3.checked = true;
    }
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 4)
    {
        frmDefinition.optOutputFormat4.checked = true;
    }
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 5)
    {
        frmDefinition.optOutputFormat5.checked = true;
    }
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 6)
    {
        frmDefinition.optOutputFormat6.checked = true;
    }
    else 
    {
        frmDefinition.optOutputFormat0.checked = true;
    }
		
    frmDefinition.chkDestination0.checked = (frmOriginalDefinition.txtDefn_OutputScreen.value != "False");
    frmDefinition.chkDestination1.checked = (frmOriginalDefinition.txtDefn_OutputPrinter.value != "False");

    if (frmDefinition.chkDestination1.checked == true) {
        populatePrinters();
        for (i=0; i<frmDefinition.cboPrinterName.options.length; i++) {
            if (frmDefinition.cboPrinterName.options(i).innerText == frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
                frmDefinition.cboPrinterName.selectedIndex = i;
                break;
            }			
        }
    }


    frmDefinition.chkDestination2.checked = (frmOriginalDefinition.txtDefn_OutputSave.value != "False");
    if (frmDefinition.chkDestination2.checked == true)
    {
        populateSaveExisting();
        frmDefinition.cboSaveExisting.selectedIndex = frmOriginalDefinition.txtDefn_OutputSaveExisting.value;
    }
    frmDefinition.chkDestination3.checked = (frmOriginalDefinition.txtDefn_OutputEmail.value != "False");		
    if (frmDefinition.chkDestination3.checked == true)
    {
        frmDefinition.txtEmailGroupID.value = frmOriginalDefinition.txtDefn_OutputEmailAddr.value;
        frmDefinition.txtEmailGroup.value = frmOriginalDefinition.txtDefn_OutputEmailAddrName.value;
        frmDefinition.txtEmailSubject.value = frmOriginalDefinition.txtDefn_OutputEmailSubject.value;
        frmDefinition.txtEmailAttachAs.value = frmOriginalDefinition.txtDefn_OutputEmailAttachAs.value;
    }	
    frmDefinition.txtFilename.value = frmOriginalDefinition.txtDefn_OutputFilename.value;


    // If its read only, disable everything.
    if(frmUseful.txtAction.value.toUpperCase() == "VIEW")
    {
        disableAll();
    }
		
    //TODO
    //window.parent.frames("menuframe").ASRIntranetFunctions.ClosePopup();

    if (frmDefinition.chkDestination1.checked == true) {
        if (frmOriginalDefinition.txtDefn_OutputPrinterName.value != "") {
            if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.selectedIndex).innerText != frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
                OpenHR.MessageBox("This definition is set to output to printer "+frmOriginalDefinition.txtDefn_OutputPrinterName.value+" which is not set up on your PC.");
                var oOption = document.createElement("OPTION");
                frmDefinition.cboPrinterName.options.add(oOption);
                oOption.innerText = frmOriginalDefinition.txtDefn_OutputPrinterName.value;
                oOption.value = frmDefinition.cboPrinterName.options.length-1;
                frmDefinition.cboPrinterName.selectedIndex = oOption.value;
            }
        }
    }

}

function setFileFormat(piFormat) 
{
    var i;
	
    for (i=0; i<frmDefinition.cboFileFormat.options.length; i++)  {
        if (frmDefinition.cboFileFormat.options(i).value == piFormat) {
            frmDefinition.cboFileFormat.selectedIndex = i;
            return;
        }		
    }
    frmDefinition.cboFileFormat.selectedIndex = 0;
    frmOriginalDefinition.txtDefn_DefaultExportTo.value = 0;
}

function disableAll()
{
    var i;
	
    var dataCollection = frmDefinition.elements;
    if (dataCollection!=null) 
    {
        for (i=0; i<dataCollection.length; i++)  
        {
            var eElem = frmDefinition.elements[i];

            if ("text" == eElem.type)  
            {
                text_disable(eElem, true);
            }
            else if ("TEXTAREA" == eElem.tagName) 
            {
                textarea_disable(eElem, true);
            }
            else if ("checkbox" == eElem.type)  
            {
                checkbox_disable(eElem, true);
            }
            else if ("radio" == eElem.type) 
            {
                radio_disable(eElem, true);
            }
            else if ("button" == eElem.type) 
            {
                if (eElem.value != "Cancel") 
                {
                    button_disable(eElem, true);
                }
            }
            else if ("SELECT" == eElem.tagName) 
            {
                combo_disable(eElem, true);
            }
            else 
            {
                grid_disable(eElem, true);
            }
        }
    }	
}

function populatePrinters()
{

    with (frmDefinition.cboPrinterName)
    {

        strCurrentPrinter = '';
        if (selectedIndex > 0) {
            strCurrentPrinter = options[selectedIndex].innerText;
        }


        length = 0;
        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "<Default Printer>";
        oOption.value = 0;

        //TODO Do this printer loop when we redo asrintrnetFunctions
        //for (iLoop=0; iLoop<window.parent.frames("menuframe").ASRIntranetFunctions.PrinterCount(); iLoop++)  {

        //    var oOption = document.createElement("OPTION");
        //    options.add(oOption);
        //    oOption.innerText = window.parent.frames("menuframe").ASRIntranetFunctions.PrinterName(iLoop);
        //    oOption.value = iLoop+1;

        //    if (oOption.innerText == strCurrentPrinter) {
        //        selectedIndex = iLoop+1;
        //    }
        //}

        if (strCurrentPrinter != '') {
            if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.selectedIndex).innerText != strCurrentPrinter) {
                var oOption = document.createElement("OPTION");
                frmDefinition.cboPrinterName.options.add(oOption);
                oOption.innerText = strCurrentPrinter;
                oOption.value = frmDefinition.cboPrinterName.options.length-1;
                selectedIndex = oOption.value;
            }
        }
    }

}

function populateSaveExisting()
{
    with (frmDefinition.cboSaveExisting)
    {
        lngCurrentOption = 0;
        if (selectedIndex > 0) {
            lngCurrentOption = options[selectedIndex].value;
        }
        length = 0;

        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "Overwrite";
        oOption.value = 0;
		
        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "Do not overwrite";
        oOption.value = 1;
		
        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "Add sequential number to name";
        oOption.value = 2;
		
        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "Append to file";
        oOption.value = 3;
		
        if ((frmDefinition.optOutputFormat4.checked) ||
            (frmDefinition.optOutputFormat5.checked) ||
            (frmDefinition.optOutputFormat6.checked))
        {
            var oOption = document.createElement("OPTION");
            options.add(oOption);
            oOption.innerText = "Create new sheet in workbook";
            oOption.value = 4;
        }

        for (iLoop=0; iLoop<options.length; iLoop++)  {
            if (options(iLoop).value == lngCurrentOption) {
                selectedIndex = iLoop;
                break;
            }
        }

    }
}

function createNew(pPopup)
{
    pPopup.close();
	
    frmUseful.txtUtilID.value = 0;
    frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
    frmUseful.txtAction.value = "new";
	
    submitDefinition();
}

function refreshFile()
{
    var sText;
    var iIndex = frmDefinition.cboFileFormat.options[frmDefinition.cboFileFormat.selectedIndex].value;	
	
    if (frmDefinition.txtFilename.value != "") {
        sText = frmDefinition.txtFilename.value;

        if (iIndex == 0) {
            sText = sText.substr(0, sText.length - 4) + ".htm";	
        }
        if (iIndex == 1) {
            sText = sText.substr(0, sText.length - 4) + ".xls";	
        }
        if (iIndex == 2) {
            sText = sText.substr(0, sText.length - 4) + ".doc";	
        }
    
        frmDefinition.txtFilename.value = sText;
    }
}

function populateAccessGrid()
{
    // Set focus onto the grid. This ensures it is properly loaded before adding items to it.
    frmDefinition.grdAccess.focus();
    frmDefinition.grdAccess.removeAll();
	
    var dataCollection = frmAccess.elements;
    if (dataCollection!=null) {

        frmDefinition.grdAccess.AddItem("(All Groups)");
        for (i=0; i<dataCollection.length; i++)  {
            frmDefinition.grdAccess.AddItem(dataCollection.item(i).value);
        }
    }
}

function setJobsToHide(psJobs)
{
    frmSend.txtSend_jobsToHide.value = psJobs;
    frmSend.txtSend_jobsToHideGroups.value = frmValidate.validateHiddenGroups.value;
}

function changeTab1Control() {
    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}

function changeTab2Control() {
    frmUseful.txtChanged.value = 1;
    refreshTab2Controls();
}

function changeTab3Control() {
    frmUseful.txtChanged.value = 1;
    refreshTab3Controls();
}

    function grdAccess_ComboCloseUp() { 
        frmUseful.txtChanged.value = 1;
        if((frmDefinition.grdAccess.AddItemRowIndex(frmDefinition.grdAccess.Bookmark) == 0) &&
          (frmDefinition.grdAccess.Columns("Access").Text.length > 0)) {
            ForceAccess(frmDefinition.grdAccess, AccessCode(frmDefinition.grdAccess.Columns("Access").Text));
    
            frmDefinition.grdAccess.MoveFirst();
            frmDefinition.grdAccess.Col = 1;
        }
        refreshTab1Controls();
    }
    
function grdAccess_GotFocus() {
    frmDefinition.grdAccess.Col = 1;
}
    
function grdAccess_RowColChange() {
    var fViewing;
    var fIsNotOwner;
    var varBkmk;
		
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

    if (frmDefinition.grdAccess.AddItemRowIndex(frmDefinition.grdAccess.Bookmark) == 0) {
        frmDefinition.grdAccess.Columns("Access").Text = "";
    }

    varBkmk = frmDefinition.grdAccess.SelBookmarks(0);

    if ((fIsNotOwner == true) ||
        (fViewing == true) ||
        (frmSelectionAccess.forcedHidden.value == "Y") ||
        (frmDefinition.grdAccess.Columns("SysSecMgr").CellText(varBkmk) == "1")) {
        frmDefinition.grdAccess.Columns("Access").Style = 0; // 0 = Edit
    }
    else {
        frmDefinition.grdAccess.Columns("Access").Style = 3; // 3 = Combo box
        frmDefinition.grdAccess.Columns("Access").RemoveAll();
        frmDefinition.grdAccess.Columns("Access").AddItem(AccessDescription("RW"));
        frmDefinition.grdAccess.Columns("Access").AddItem(AccessDescription("RO"));
        frmDefinition.grdAccess.Columns("Access").AddItem(AccessDescription("HD"));
    }

    frmDefinition.grdAccess.Col = 1;
}
    
function grdAccess_RowLoaded() {
    var fViewing;
    var fIsNotOwner;
		
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

    if ((fIsNotOwner == true) ||
        (fViewing == true) ||
        (frmSelectionAccess.forcedHidden.value == "Y")) {
        frmDefinition.grdAccess.Columns("GroupName").CellStyleSet("ReadOnly");
        frmDefinition.grdAccess.Columns("Access").CellStyleSet("ReadOnly");
        frmDefinition.grdAccess.ForeColor = "-2147483631";
    }  
    else {
        if (frmDefinition.grdAccess.Columns("SysSecMgr").CellText(Bookmark) == "1") {
            frmDefinition.grdAccess.Columns("GroupName").CellStyleSet("SysSecMgr");
            frmDefinition.grdAccess.Columns("Access").CellStyleSet("SysSecMgr");
            frmDefinition.grdAccess.ForeColor = "0";
        }
        else {
            frmDefinition.grdAccess.ForeColor = "0";
        }
    }
}

