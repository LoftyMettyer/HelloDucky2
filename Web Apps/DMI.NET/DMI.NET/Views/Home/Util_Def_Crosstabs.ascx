<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>

<script type="text/javaScript">

    function util_def_crosstabs_window_onload() {
        var fOK;
        fOK = true;
        debugger;
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
</script>

<script type="text/javaScript" id=scptGeneralFunctions>
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
-->
</script>

<script type="text/javaScript">
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
</script>


<%--<script FOR=grdAccess EVENT=ComboCloseUp LANGUAGE=JavaScript>
<!--
    frmUseful.txtChanged.value = 1;
    if((grdAccess.AddItemRowIndex(grdAccess.Bookmark) == 0) &&
        (grdAccess.Columns("Access").Text.length > 0)) {
        ForceAccess(grdAccess, AccessCode(grdAccess.Columns("Access").Text));
    
        grdAccess.MoveFirst();
        grdAccess.Col = 1;
    }
    refreshTab1Controls();
-->
</script>

<script FOR=grdAccess EVENT=GotFocus LANGUAGE=JavaScript>
<!--
    grdAccess.Col = 1
-->
</script>

<script FORM=grdAccess EVENT=RowColChange(LastRow, LastCol) LANGUAGE=JavaScript>
<!--
    var fViewing;
    var fIsNotOwner;
    var varBkmk;
		
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

    if (grdAccess.AddItemRowIndex(grdAccess.Bookmark) == 0) {
        grdAccess.Columns("Access").Text = "";
    }

    varBkmk = grdAccess.SelBookmarks(0);

    if ((fIsNotOwner == true) ||
        (fViewing == true) ||
        (frmSelectionAccess.forcedHidden.value == "Y") ||
        (grdAccess.Columns("SysSecMgr").CellText(varBkmk) == "1")) {
        grdAccess.Columns("Access").Style = 0; // 0 = Edit
    }
    else {
        grdAccess.Columns("Access").Style = 3; // 3 = Combo box
        grdAccess.Columns("Access").RemoveAll();
        grdAccess.Columns("Access").AddItem(AccessDescription("RW"));
        grdAccess.Columns("Access").AddItem(AccessDescription("RO"));
        grdAccess.Columns("Access").AddItem(AccessDescription("HD"));
    }

    grdAccess.Col = 1;
</script>

<script FOR=grdAccess EVENT=RowLoaded(Bookmark) LANGUAGE=JavaScript>
    var fViewing;
    var fIsNotOwner;
		
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

    if ((fIsNotOwner == true) ||
        (fViewing == true) ||
        (frmSelectionAccess.forcedHidden.value == "Y")) {
        grdAccess.Columns("GroupName").CellStyleSet("ReadOnly");
        grdAccess.Columns("Access").CellStyleSet("ReadOnly");
        grdAccess.ForeColor = "-2147483631";
    }  
    else {
        if (grdAccess.Columns("SysSecMgr").CellText(Bookmark) == "1") {
            grdAccess.Columns("GroupName").CellStyleSet("SysSecMgr");
            grdAccess.Columns("Access").CellStyleSet("SysSecMgr");
            grdAccess.ForeColor = "0";
        }
        else {
            grdAccess.ForeColor = "0";
        }
    }
</script>--%>

<object classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
    id="dialog"
    codebase="cabs/comdlg32.cab#Version=1,0,0,0"
    style="LEFT: 0px; TOP: 0px"
    viewastext>
    <param name="_ExtentX" value="847">
    <param name="_ExtentY" value="847">
    <param name="_Version" value="393216">
    <param name="CancelError" value="0">
    <param name="Color" value="0">
    <param name="Copies" value="1">
    <param name="DefaultExt" value="">
    <param name="DialogTitle" value="">
    <param name="FileName" value="">
    <param name="Filter" value="">
    <param name="FilterIndex" value="0">
    <param name="Flags" value="0">
    <param name="FontBold" value="0">
    <param name="FontItalic" value="0">
    <param name="FontName" value="">
    <param name="FontSize" value="8">
    <param name="FontStrikeThru" value="0">
    <param name="FontUnderLine" value="0">
    <param name="FromPage" value="0">
    <param name="HelpCommand" value="0">
    <param name="HelpContext" value="0">
    <param name="HelpFile" value="">
    <param name="HelpKey" value="">
    <param name="InitDir" value="">
    <param name="Max" value="0">
    <param name="Min" value="0">
    <param name="MaxFileSize" value="260">
    <param name="PrinterDefault" value="1">
    <param name="ToPage" value="0">
    <param name="Orientation" value="1">
</object>

<DIV <%=session("BodyTag")%>>
    <form id=frmTables style="visibility:hidden;display:none">
        <%
            Dim sErrorDescription = ""

            ' Get the table records.
            Dim cmdTables = Server.CreateObject("ADODB.Command")
            cmdTables.CommandText = "sp_ASRIntGetCrossTabTablesInfo"
            cmdTables.CommandType = 4 ' Stored Procedure
            cmdTables.ActiveConnection = Session("databaseConnection")
	
            Response.Write("<B>Set Connection</B>")
	
            Err.Clear()
            Dim rstTablesInfo = cmdTables.Execute
	
            Response.Write("<B>Executed SP</B>")
	
            If (Err.Number <> 0) Then
                sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                ' Dim iCount = 0
                Do While Not rstTablesInfo.EOF
                    Response.Write("<INPUT type='hidden' id=txtTableName_" & rstTablesInfo.fields("tableID").value & " name=txtTableName_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("tableName").value & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtTableType_" & rstTablesInfo.fields("tableID").value & " name=txtTableType_" & rstTablesInfo.fields("tableID").value & " value=" & rstTablesInfo.fields("tableType").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenString").value & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenNames").value & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtTableParents_" & rstTablesInfo.fields("tableID").value & " name=txtTableParents_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("parentsString").value & """>" & vbCrLf)

                    rstTablesInfo.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstTablesInfo.close()
                rstTablesInfo = Nothing
            End If
	
            ' Release the ADO command object.
            cmdTables = Nothing
%>
    </form>
    <form id=frmOriginalDefinition name=frmOriginalDefinition style="visibility:hidden;display:none">
        <%
            Dim sErrMsg = ""
            Dim lngHStart = 0
            Dim lngHStop = 0
            Dim lngHStep = 0
            Dim lngVStart = 0
            Dim lngVStop = 0
            Dim lngVStep = 0
            Dim lngPStart = 0
            Dim lngPStop = 0
            Dim lngPStep = 0

            If Session("action") <> "new" Then
                Dim cmdDefn = Server.CreateObject("ADODB.Command")
                cmdDefn.CommandText = "sp_ASRIntGetCrossTabDefinition"
                cmdDefn.CommandType = 4 ' Stored Procedure
                cmdDefn.ActiveConnection = Session("databaseConnection")
                
                Dim prmUtilDefnID = cmdDefn.CreateParameter("utilid", 3, 1) ' 3=integer, 1=input
                cmdDefn.Parameters.Append(prmUtilDefnID)
                prmUtilDefnID.value = CleanNumeric(Session("utilid"))
                
                Dim prmUser = cmdDefn.CreateParameter("user", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
                cmdDefn.Parameters.Append(prmUser)
                prmUser.value = Session("username")

                Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
                cmdDefn.Parameters.Append(prmAction)
                prmAction.value = Session("action")

                Dim prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmErrMsg)

                Dim prmName = cmdDefn.CreateParameter("name", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmName)

                Dim prmOwner = cmdDefn.CreateParameter("owner", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOwner)

                Dim prmDescription = cmdDefn.CreateParameter("description", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmDescription)

                Dim prmBaseTableID = cmdDefn.CreateParameter("baseTableID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmBaseTableID)

                Dim prmAllRecords = cmdDefn.CreateParameter("allRecords", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmAllRecords)

                Dim prmPicklistID = cmdDefn.CreateParameter("picklistID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmPicklistID)

                Dim prmPicklistName = cmdDefn.CreateParameter("picklistName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmPicklistName)

                Dim prmPicklistHidden = cmdDefn.CreateParameter("picklistHidden", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmPicklistHidden)

                Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmFilterID)

                Dim prmFilterName = cmdDefn.CreateParameter("filterName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmFilterName)

                Dim prmFilterHidden = cmdDefn.CreateParameter("filterHidden", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmFilterHidden)
		
                Dim prmPrintFilter = cmdDefn.CreateParameter("PrintFilter", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmPrintFilter)

                Dim prmHColID = cmdDefn.CreateParameter("HColID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmHColID)

                Dim prmHStart = cmdDefn.CreateParameter("HStart", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmHStart)

                Dim prmHStop = cmdDefn.CreateParameter("HStop", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmHStop)

                Dim prmHStep = cmdDefn.CreateParameter("HStep", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmHStep)

                Dim prmVColID = cmdDefn.CreateParameter("VColID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmVColID)

                Dim prmVStart = cmdDefn.CreateParameter("VStart", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmVStart)

                Dim prmVStop = cmdDefn.CreateParameter("VStop", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmVStop)

                Dim prmVStep = cmdDefn.CreateParameter("VStep", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmVStep)

                Dim prmPColID = cmdDefn.CreateParameter("PColID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmPColID)

                Dim prmPStart = cmdDefn.CreateParameter("PStart", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmPStart)

                Dim prmPStop = cmdDefn.CreateParameter("PStop", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmPStop)

                Dim prmPStep = cmdDefn.CreateParameter("PStep", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmPStep)

                Dim prmIType = cmdDefn.CreateParameter("IType", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmIType)

                Dim prmIColID = cmdDefn.CreateParameter("IColID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmIColID)

                Dim prmPercentage = cmdDefn.CreateParameter("Percentage", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmPercentage)

                Dim prmPerPage = cmdDefn.CreateParameter("PerPage", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmPerPage)

                Dim prmSuppress = cmdDefn.CreateParameter("Suppress", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmSuppress)

                Dim prmThousand = cmdDefn.CreateParameter("Thousand", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmThousand)

                Dim prmOutputPreview = cmdDefn.CreateParameter("outputPreview", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputPreview)
		
                Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmOutputFormat)
		
                Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputScreen)
		
                Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputPrinter)
		
                Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputPrinterName)
		
                Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputSave)
		
                Dim prmOutputSaveExisting = cmdDefn.CreateParameter("outputSaveExisting", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmOutputSaveExisting)
		
                Dim prmOutputEmail = cmdDefn.CreateParameter("outputEmail", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputEmail)
		
                Dim prmOutputEmailAddr = cmdDefn.CreateParameter("outputEmailAddr", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmOutputEmailAddr)

                Dim prmOutputEmailAddrName = cmdDefn.CreateParameter("outputEmailAddrName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputEmailAddrName)

                Dim prmOutputEmailSubject = cmdDefn.CreateParameter("outputEmailSubject", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputEmailSubject)

                Dim prmOutputEmailAttachAs = cmdDefn.CreateParameter("outputEmailAttachAs", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputEmailAttachAs)

                Dim prmOutputFilename = cmdDefn.CreateParameter("outputFilename", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputFilename)

                Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2) ' 3=integer, 2=output
                cmdDefn.Parameters.Append(prmTimestamp)

                Err.Clear()
                cmdDefn.Execute()

                Dim iHiddenCalcCount As Integer = 0
                If (Err.Number <> 0) Then
                    sErrMsg = "'" & Session("utilname") & "' cross tab definition could not be read." & vbCrLf & FormatError(Err.Description)
                Else

                    'rstDefinition.close
                    'set rstDefinition = nothing

                    ' NB. IMPORTANT ADO NOTE.
                    ' When calling a stored procedure which returns a recordset AND has output parameters
                    ' you need to close the recordset and set it to nothing before using the output parameters. 
                    If Len(cmdDefn.Parameters("errMsg").value) > 0 Then
                        sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").value
                    End If

                    lngHStart = cmdDefn.Parameters("HStart").value
                    lngHStop = cmdDefn.Parameters("HStop").value
                    lngHStep = cmdDefn.Parameters("HStep").value
                    lngVStart = cmdDefn.Parameters("VStart").value
                    lngVStop = cmdDefn.Parameters("VStop").value
                    lngVStep = cmdDefn.Parameters("VStep").value
                    lngPStart = cmdDefn.Parameters("PStart").value
                    lngPStop = cmdDefn.Parameters("PStop").value
                    lngPStep = cmdDefn.Parameters("PStep").value

                    Response.Write("<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & cmdDefn.Parameters("baseTableID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_AllRecords name=txtDefn_AllRecords value=" & cmdDefn.Parameters("allRecords").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & cmdDefn.Parameters("picklistID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(cmdDefn.Parameters("picklistName").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & cmdDefn.Parameters("picklistHidden").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & cmdDefn.Parameters("filterID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(cmdDefn.Parameters("filterName").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & cmdDefn.Parameters("filterHidden").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_FilterHeader name=txtDefn_FilterHeader value=" & cmdDefn.Parameters("PrintFilter").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PrintFilter name=txtDefn_PrintFilter value=" & cmdDefn.Parameters("PrintFilter").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HColID name=txtDefn_HColID value=" & cmdDefn.Parameters("HColID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HStart name=txtDefn_HStart value=" & cmdDefn.Parameters("HStart").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HStop name=txtDefn_HStop value=" & cmdDefn.Parameters("HStop").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HStep name=txtDefn_HStep value=" & cmdDefn.Parameters("HStep").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_VColID name=txtDefn_VColID value=" & cmdDefn.Parameters("VColID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_VStart name=txtDefn_VStart value=" & cmdDefn.Parameters("VStart").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_VStop name=txtDefn_VStop value=" & cmdDefn.Parameters("VStop").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_VStep name=txtDefn_VStep value=" & cmdDefn.Parameters("VStep").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PColID name=txtDefn_PColID value=" & cmdDefn.Parameters("PColID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PStart name=txtDefn_PStart value=" & cmdDefn.Parameters("PStart").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PStop name=txtDefn_PStop value=" & cmdDefn.Parameters("PStop").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PStep name=txtDefn_PStep value=" & cmdDefn.Parameters("PStep").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_IType name=txtDefn_IType value=" & cmdDefn.Parameters("IType").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_IColID name=txtDefn_IColID value=" & cmdDefn.Parameters("IColID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Percentage name=txtDefn_Percentage value=" & cmdDefn.Parameters("Percentage").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PerPage name=txtDefn_PerPage value=" & cmdDefn.Parameters("PerPage").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Suppress name=txtDefn_Suppress value=" & cmdDefn.Parameters("Suppress").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Use1000 name=txtDefn_Use1000 value=" & cmdDefn.Parameters("Thousand").value & ">" & vbCrLf)

                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputPreview name=txtDefn_OutputPreview value=" & cmdDefn.Parameters("OutputPreview").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & cmdDefn.Parameters("OutputFormat").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & cmdDefn.Parameters("OutputScreen").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & cmdDefn.Parameters("OutputPrinter").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & cmdDefn.Parameters("OutputPrinterName").value & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & cmdDefn.Parameters("OutputSave").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputSaveExisting name=txtDefn_OutputSaveExisting value=" & cmdDefn.Parameters("OutputSaveExisting").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmail name=txtDefn_OutputEmail value=" & cmdDefn.Parameters("OutputEmail").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddr name=txtDefn_OutputEmailAddr value=" & cmdDefn.Parameters("OutputEmailAddr").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddrName name=txtDefn_OutputEmailName value=""" & Replace(cmdDefn.Parameters("OutputEmailAddrName").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailSubject name=txtDefn_OutputEmailSubject value=""" & Replace(cmdDefn.Parameters("OutputEmailSubject").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAttachAs name=txtDefn_OutputEmailAttachAs value=""" & Replace(cmdDefn.Parameters("OutputEmailAttachAs").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputFilename name=txtDefn_OutputFilename value=""" & cmdDefn.Parameters("OutputFilename").value & """>" & vbCrLf)

                    Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=session_action name=session_action value=" & Session("action") & ">" & vbCrLf)
                    Response.Write("</form>" & vbCrLf)

                End If

                ' Release the ADO command object.
                cmdDefn = Nothing

                If Len(sErrMsg) > 0 Then
                    Session("confirmtext") = sErrMsg
                    Session("confirmtitle") = "OpenHR Intranet"
                    Session("followpage") = "defsel"
                    Session("reaction") = "CROSSTABS"
                    Response.Clear()
                    Response.Redirect("confirmok")
                End If
	
            Else
                Session("childcount") = 0
                Session("hiddenfiltercount") = 0
            End If
%>
    </form>

    <form id=frmDefinition name=frmDefinition>
        <table valign=top align=center class="outline" cellPadding=5 width=100% height=100% cellSpacing=0>
            <tr>
                <td>
                    <TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
                        <tr height=5> 
                            <td colspan=3></td>
                        </tr> 

                        <tr height=10>
                            <td width=10></td>
                            <td>
                                <INPUT type="button" value="Definition" id=btnTab1 name=btnTab1 class="btn btndisabled" disabled="disabled"
                                       onclick="displayPage(1)" 
                                       onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                       onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                       onfocus="try{button_onFocus(this);}catch(e){}"
                                       onblur="try{button_onBlur(this);}catch(e){}" />
                                <INPUT type="button" value="Columns" id=btnTab2 name=btnTab2 class="btn btndisabled" disabled="disabled"
                                       onclick="displayPage(2)" 
                                       onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                       onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                       onfocus="try{button_onFocus(this);}catch(e){}"
                                       onblur="try{button_onBlur(this);}catch(e){}" />
                                <INPUT type="button" value="Output" id=btnTab3 name=btnTab3 class="btn btndisabled" disabled="disabled"
                                       onclick="displayPage(3)" 
                                       onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                       onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                       onfocus="try{button_onFocus(this);}catch(e){}"
                                       onblur="try{button_onBlur(this);}catch(e){}" />
                            </td>
                            <td width=10></td>
                        </tr>

                        <tr height=10> 
                            <td colspan=3></td>
                        </tr> 

                        <tr>
                            <td width=10></td>
                            <td>
                                <!-- First tab -->
                                <DIV id=div1>
                                    <TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
                                        <tr valign=top> 
                                            <td>
                                                <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                    <tr height=10>
                                                        <td width=5>&nbsp;</td>
                                                        <td width=10>Name :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td>
                                                            <INPUT id=txtName name=txtName maxlength="50" style="WIDTH: 100%" class="text"
                                                                   onkeyup="changeTab1Control()">
                                                        </td>
                                                        <td width=20>&nbsp;</td>
                                                        <td width=10>Owner :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%">
                                                            <INPUT id=txtOwner name=txtOwner class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
                                                        </td>
                                                        <td width=5>&nbsp;</td>
                                                    </tr>

                                                    <tr>
                                                        <td colspan=9 height=5></td>
                                                    </tr>

                                                    <tr height=60>
                                                        <td width=5>&nbsp;</td>
                                                        <td width=10 nowrap valign=top>Description :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%" rowspan="3">
                                                            <TEXTAREA id=txtDescription name=txtDescription class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap=VIRTUAL height="0" maxlength="255" 
                                                                      onkeyup="changeTab1Control()" 
                                                                      onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}" 
                                                                      onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
													</TEXTAREA>
                                                        </td>
                                                        <td width=20 nowrap>&nbsp;</td>
                                                        <td width=10 valign=top>Access :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%" rowspan="3" valign=top>
                                                        </td>
                                                        <td width=5>&nbsp;</td>
                                                    </tr>

                                                    <tr height=10>
                                                        <td colspan=7>&nbsp;</td>
                                                    </tr>

                                                    <tr height=10>
                                                        <td colspan=7>&nbsp;</td>
                                                    </tr>

                                                    <tr>
                                                        <td colspan=9><hr></td>
                                                    </tr>

                                                    <tr height=10>
                                                        <td width=5>&nbsp;</td>
                                                        <td width=100 nowrap vAlign=top>Base Table :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%" vAlign=top>
                                                            <select id=cboBaseTable name=cboBaseTable style="WIDTH: 100%" class="combo combodisabled"
                                                                    onchange="changeBaseTable()" disabled="disabled"> 
                                                            </select>
                                                        </td>
                                                        <td width=20 nowrap>&nbsp;</td>
                                                        <td width=10 vAlign=top>Records :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%"> 
                                                            <TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
                                                                <tr>
                                                                    <td width=5>
                                                                        <input CHECKED id=optRecordSelection1 name=optRecordSelection type=radio 
                                                                               onclick="changeBaseTableRecordOptions()"
                                                                               onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                               onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td width=30>
                                                                        <label 
                                                                            tabindex="-1"
                                                                            for="optRecordSelection1"
                                                                            class="radio"
                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
                                                                        All
                                                                    </label>
                                                                    </td>
                                                                    <td>&nbsp;</td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan=6 height=5></td>
                                                                </tr>
                                                                <tr>
                                                                    <td width=5>
                                                                        <input id=optRecordSelection2 name=optRecordSelection type=radio 
                                                                               onclick="changeBaseTableRecordOptions()"
                                                                               onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                               onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td width=20>
                                                                        <label 
                                                                            tabindex="-1"
                                                                            for="optRecordSelection2"
                                                                            class="radio"
                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
                                                                        Picklist
                                                                    </label>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td>
                                                                        <INPUT id=txtBasePicklist name=txtBasePicklist class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
                                                                    </td>
                                                                    <td width=30>
                                                                        <INPUT id=cmdBasePicklist name=cmdBasePicklist style="WIDTH: 100%" type=button disabled="disabled" value="..." class="btn btndisabled"
                                                                               onclick="selectRecordOption('base', 'picklist')"
                                                                               onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{button_onFocus(this);}catch(e){}"
                                                                               onblur="try{button_onBlur(this);}catch(e){}" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan=6 height=5></td>
                                                                </tr>
                                                                <tr>
                                                                    <td width=5>
                                                                        <input id=optRecordSelection3 name=optRecordSelection type=radio
                                                                               onclick=changeBaseTableRecordOptions() 
                                                                               onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                               onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td width=20>
                                                                        <label 
                                                                            tabindex="-1"
                                                                            for="optRecordSelection3"
                                                                            class="radio"
                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
                                                                        Filter
                                                                    </label>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td>
                                                                        <INPUT id=txtBaseFilter name=txtBaseFilter disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
                                                                    </td>
                                                                    <td width=30>
                                                                        <INPUT id=cmdBaseFilter name=cmdBaseFilter style="WIDTH: 100%" type=button class="btn btndisabled" disabled="disabled" value="..."
                                                                               onclick="selectRecordOption('base', 'filter')" 
                                                                               onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{button_onFocus(this);}catch(e){}"
                                                                               onblur="try{button_onBlur(this);}catch(e){}" />
                                                                    </td>
                                                                </tr>
                                                            </TABLE>
                                                        </td>
                                                        <td width=5>&nbsp;</td>
                                                    </tr>
											
                                                    <tr>
                                                        <td colspan=9 height=5>&nbsp;</td>
                                                    </tr>
											

                                                    <tr>
                                                        <td colspan=5>&nbsp;</td>
                                                        <td colspan=3>
                                                            <input name=chkPrintFilter id=chkPrintFilter type=checkbox disabled="disabled" tabindex=-1 
                                                                   onclick="changeTab1Control()"
                                                                   onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                   onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
                                                                for="chkPrintFilter"
                                                                class="checkbox checkboxdisabled"
                                                                tabindex=0 
                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                Display filter or picklist title in the report header
                                                            </label> 
                                                        </td>
                                                        <td width=5>&nbsp;</td>
                                                    </tr>
											
                                                    <tr>
                                                        <td colspan=9 height=5>&nbsp;</td>
                                                    </tr>
                                                </TABLE>
                                            </td>
                                        </tr>
                                    </TABLE>
                                </DIV>
                                <DIV id=div2 style="visibility:hidden;display:none">
                                    <TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>

                                        <tr valign=top> 
                                            <td>
                                                <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>

                                                <tr height=10>
                                                    <td width=5>&nbsp;</td>
                                                    <td colspan=4 vAlign=top><U>Headings & Breaks</U></td>
                                                    <td width="15%" align=Center>Start</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%" align=Center>Stop</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%" align=Center>Increment</td>
                                                    <td>&nbsp;</td>
                                                </tr>

                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>
                                                <tr height=23>
                                                    <td width=5>&nbsp;</td>
                                                    <td width=80 nowrap vAlign=top>Horizontal :</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="40%" vAlign=top>
                                                        <select id=cboHor name=cboHor style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                onchange="cboHor_Change();changeTab2Control(); "> 
                                                        </select>
                                                    </td>
                                                    <td width=15>&nbsp;</td>
                                                    <td>
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtHorStart name=txtHorStart width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngHStart%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtHorStop name=txtHorStop width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngHStop%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtHorStep name=txtHorStep width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngHStep%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>

                                                    <td>&nbsp;</td>
                                                </tr>

                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>
                                                <tr height=23>
                                                    <td width=5>&nbsp;</td>
                                                    <td width=80 nowrap vAlign=top>Vertical :</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="40%" vAlign=top>
                                                        <select id=cboVer name=cboVer style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                onchange="cboVer_Change();changeTab2Control(); "> 
                                                        </select>
                                                    </td>
                                                    <td width=15>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtVerStart name=txtVerStart width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngVStart%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtVerStop name=txtVerStop width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngVStop%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtVerStep name=txtVerStep width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngVStep%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>

                                                    <td>&nbsp;</td>
                                                </tr>
                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>
                                                <tr height=23>
                                                    <td width=5>&nbsp;</td>
                                                    <td width=100 nowrap vAlign=top>Page Break :</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="40%" vAlign=top>
                                                        <select id=cboPgb name=cboPgb style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                onchange="cboPgb_Change();changeTab2Control(); " > 
                                                        </select>
                                                    </td>
                                                    <td width=15>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtPgbStart name=txtPgbStart style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngPStart%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtPgbStop name=txtPgbStop style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngPStop%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtPgbStep name=txtPgbStep style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngPStep%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td>&nbsp;</td>
                                                </tr>

                                                <tr height=40>
                                                    <td colspan=11><hr></td>
                                                </tr>

                                                <tr height=10>
                                                    <td width=5>&nbsp;</td>
                                                    <td width=80 colspan=4 nowrap vAlign=top><U>Intersection</U></td>
                                                </tr>
                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>

                                                <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                    <tr height=0>
                                                        <td width=90></td>
                                                        <td width="40%"></td>
                                                    </tr>

                                                    <td colspan=2>
                                                        <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                            <tr height=10>
                                                                <td width=5>&nbsp;</td>
                                                                <td width=80 nowrap vAlign=top>Column :</td>
                                                                <td width=5>&nbsp;</td>
                                                                <td width="100%" vAlign=top>
                                                                    <select id=cboInt name=cboInt style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                            onchange="cboInt_Change();changeTab2Control(); " > 
                                                                    </select>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td colspan=9 height=5></td>
                                                            </tr>
                                                            <tr height=10>
                                                                <td width=5>&nbsp;</td>
                                                                <td width=80 nowrap vAlign=top>Type :</td>
                                                                <td width=5>&nbsp;</td>
                                                                <td width="100%" vAlign=top>
                                                                    <select id=cboIntType name=cboIntType style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                            onchange="changeTab2Control()" >
                                                                        <option value="1">Average</option>
                                                                        <option value="0" selected>Count</option>
                                                                        <option value="2">Maximum</option>
                                                                        <option value="3">Minimum</option>
                                                                        <option value="4">Total</option>
                                                                    </select>
                                                                </td>
                                                            </tr>
                                                        </TABLE>
                                                    </td>
                                                    <td width=15>&nbsp;</td>
                                                    <td>
                                                        <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                            <tr>
                                                                <td>
                                                                    <INPUT type="checkbox" id=chkPercentage name=chkPercentage tabindex=-1
                                                                           onclick="changeTab2Control()"
                                                                           onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                    <label 
                                                                        for="chkPercentage"
                                                                        class="checkbox"
                                                                        tabindex=0 
                                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    
                                                                        Percentage of Type
                                                                    </label> 
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td height=5></td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <INPUT type="checkbox" id=chkPerPage name=chkPerPage tabindex=-1
                                                                           onclick="changeTab2Control()"                                                                     
                                                                           onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                    <label 
                                                                        for="chkPerPage"
                                                                        class="checkbox"
                                                                        tabindex=0 
                                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    
                                                                        Percentage of Page
                                                                    </label> 
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td height=5></td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <INPUT type="checkbox" id=chkSuppress name=chkSuppress tabindex=-1
                                                                           onclick="changeTab2Control()"
                                                                           onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                    <label 
                                                                        for="chkSuppress"
                                                                        class="checkbox"
                                                                        tabindex=0 
                                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    
                                                                        Suppress Zeros
                                                                    </label> 
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td height=5></td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <INPUT type="checkbox" id=chkUse1000 name=chkUse1000 tabindex=-1
                                                                           onclick="changeTab2Control()" 																    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                    <label 
                                                                        for="chkUse1000"
                                                                        class="checkbox"
                                                                        tabindex=0 
                                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    
                                                                        Use 1000 Separators (,)
                                                                    </label> 
                                                                </td>
                                                            </tr>													
                                                        </TABLE>
                                                    </td>
                                                    <tr>
                                                        <td colspan=9 height=5></td>
                                                    </tr>
                                                </TABLE>
                                            </td>
                                        </tr>
                                    </TABLE>
                                </DIV>

                                <!-- OUTPUT OPTIONS -->
                                <DIV id=div3 style="visibility:hidden;display:none">
                                <TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
                                    <tr valign=top> 
                                        <td>
                                            <TABLE WIDTH="100%" class="invisible" CELLSPACING=10 CELLPADDING=0>
                                                <tr>						
                                                    <td valign=top rowspan=2 width=25% height="100%">
                                                        <table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
                                                            <tr height=10> 
                                                                <td height=10 align=left valign=top>
                                                                    Output Format : <BR><BR>
                                                                                        <TABLE class="invisible" cellspacing="0" cellpadding="0" width="100%">
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat0 value=0
                                                                                                           onClick="formatClick(0);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat0"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Data Only
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat1 value=1
                                                                                                           onClick="formatClick(1);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat1"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    CSV File
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat2 value=2
                                                                                                           onClick="formatClick(2);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat2"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    HTML Document
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat3 value=3
                                                                                                           onClick="formatClick(3);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat3"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Word Document
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat4 value=4
                                                                                                           onClick="formatClick(4);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat4"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Excel Worksheet
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=5>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat5 value=5
                                                                                                           onClick="formatClick(5);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat5"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Excel Chart
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=5>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat6 value=6
                                                                                                           onClick="formatClick(6);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat6"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Excel Pivot Table
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=5> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                        </TABLE>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                    <td valign=top width="75%">
                                                        <table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
                                                            <tr height=10> 
                                                                <td height=10 align=left valign=top>
                                                                    Output Destination(s) : <BR><BR>
                                                                                                <TABLE class="invisible" cellspacing="0" cellpadding="0" width="100%">
                                                                                                    <tr height=20>
                                                                                                        <td width=5>&nbsp</td>
                                                                                                        <td align=left colspan=6 nowrap>
                                                                                                            <input name=chkPreview id=chkPreview type=checkbox disabled="disabled" tabindex=-1 
                                                                                                                   onClick="changeTab3Control();"
                                                                                                                   onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                                                                   onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                                            <label 
                                                                                                                for="chkPreview"
                                                                                                                class="checkbox checkboxdisabled"
                                                                                                                tabindex=0 
                                                                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                                                                Preview on screen
                                                                                                            </label>
                                                                                                        </td>
                                                                                                        <td width=5>&nbsp</td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=10> 
                                                                                                        <td colspan=8></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=20>
                                                                                                        <td></td>
                                                                                                        <td align=left colspan=6 nowrap>
                                                                                                            <input name=chkDestination0 id=chkDestination0 type=checkbox disabled="disabled" tabindex=-1 
                                                                                                                   onClick="changeTab3Control();"
                                                                                                                   onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                                                                   onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                                            <label 
                                                                                                                for="chkDestination0"
                                                                                                                class="checkbox checkboxdisabled"
                                                                                                                tabindex=0 
                                                                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                                                                Display output on screen 
                                                                                                            </label>
                                                                                                        </td>
                                                                                                        <td></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=10> 
                                                                                                        <td colspan=8></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=20>
                                                                                                        <td></td>
                                                                                                        <td align=left nowrap>
                                                                                                            <input name=chkDestination1 id=chkDestination1 type=checkbox disabled="disabled" tabindex=-1  
                                                                                                                   onClick="changeTab3Control();"
                                                                                                                   onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                                                                   onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                                            <label 
                                                                                                                for="chkDestination1"
                                                                                                                class="checkbox checkboxdisabled"
                                                                                                                tabindex=0 
                                                                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                                                                Send to printer 
                                                                                                            </label>
                                                                                                        </td>
                                                                                                        <td width=30 nowrap>&nbsp</td>
                                                                                                        <td align=left nowrap>
                                                                                                            Printer location : 
                                                                                                        </td>
                                                                                                        <td width=15>&nbsp</td>
                                                                                                        <td colspan=2>
                                                                                                            <select id=cboPrinterName name=cboPrinterName width=100% style="WIDTH: 400px" class="combo"
                                                                                                                    onchange="changeTab3Control()">	
                                                                                                            </select>								
                                                                                                        </td>
                                                                                                        <td></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=10> 
                                                                                                        <td colspan=8></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=20>
                                                                                                    <td></td>
                                                                                                    <td align=left nowrap>
                                                                                                        <input name=chkDestination2 id=chkDestination2 type=checkbox disabled="disabled" tabindex=-1 
                                                                                                               onClick="changeTab3Control();"
                                                                                                               onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                                                               onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                                        <label 
                                                                                                            for="chkDestination2"
                                                                                                            class="checkbox checkboxdisabled"
                                                                                                            tabindex=0 
                                                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                                                            Save to file
                                                                                                        </label>
                                                                                                    </td>
                                                                                                    <td></td>
                                                                                                    <td align=left nowrap>
                                                                                                        File name :   
                                                                                                    </td>
                                                                                                    <td></td>
                                                                                                    <td colspan=2>
                                                                                                        <TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 style="WIDTH: 400px">
                                                                                                        <tr>
                                                                                                        <td>
                                                                                                            <INPUT id=txtFilename name=txtFilename class="text textdisabled" disabled="disabled" style="WIDTH: 375px">
                                                                                                        </td>
                                                                                                        <td width=25>
                                                                                                            <INPUT id=cmdFilename name=cmdFilename style="WIDTH: 100%" type=button class="btn" value="..."
                                                                                                                   onClick="saveFile();changeTab3Control();"  			                                
                                                                                                                   onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                                                                   onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                                                                   onfocus="try{button_onFocus(this);}catch(e){}"
                                                                                                                   onblur="try{button_onBlur(this);}catch(e){}" />
                                                                                                        </td>
                                                                                                    </td>
                                                                                                </TABLE>
                                                                </td>
                                                                <td></td>
                                                            </tr>
																	
                                                            <tr height=10> 
                                                                <td colspan=8></td>
                                                            </tr>
																	
                                                            <tr height=20>
                                                                <td colspan=3></td>
                                                                <td align=left nowrap>
                                                                    If existing file :
                                                                </td>
                                                                <td></td>
                                                                <td colspan=2 width=100% nowrap>
                                                                    <select id=cboSaveExisting name=cboSaveExisting width=100% style="WIDTH: 400px" class="combo"
                                                                            onchange="changeTab3Control()">
                                                                    </select>							
                                                                </td>
                                                                <td></td>
                                                            </tr>
																	
                                                            <tr height=10> 
                                                                <td colspan=8></td>
                                                            </tr>
																	
                                                            <tr height=20>
                                                            <td></td>
                                                            <td align=left nowrap>
                                                                <input name=chkDestination3 id=chkDestination3 type=checkbox disabled="disabled" tabindex=-1 
                                                                       onClick="changeTab3Control();"
                                                                       onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                       onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                <label 
                                                                    for="chkDestination3"
                                                                    class="checkbox checkboxdisabled"
                                                                    tabindex=0 
                                                                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                    onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                    onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                    Send as email 
                                                                </label>
                                                            </td>
                                                            <td></td>
                                                            <td align=left nowrap>
                                                                Email group :   
                                                            </td>
                                                            <td></td>
                                                            <td colspan=2>
                                                                <TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 style="WIDTH: 400px">
                                                                <tr>
                                                                <td>
                                                                    <INPUT id=txtEmailGroup name=txtEmailGroup class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
                                                                    <INPUT id=txtEmailGroupID name=txtEmailGroupID type=hidden>
                                                                </td>
                                                                <td width=25>
                                                                    <INPUT id=cmdEmailGroup name=cmdEmailGroup style="WIDTH: 100%" type=button class="btn" value="..."
                                                                           onClick="selectEmailGroup();changeTab3Control();" 
                                                                           onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                           onfocus="try{button_onFocus(this);}catch(e){}"
                                                                           onblur="try{button_onBlur(this);}catch(e){}" />
                                                                </td>
                                                            </td>
                                                        </TABLE>
                                                    </td>
                                                    <td></td>
                                                </tr>
																	
                                                <tr height=10> 
                                                    <td colspan=8></td>
                                                </tr>
																	
                                                <tr height=20>
                                                    <td colspan=3></td>
                                                    <td align=left nowrap>
                                                        Email subject :   
                                                    </td>
                                                    <td></td>
                                                    <td colspan=2 width=100% nowrap>
                                                        <INPUT id=txtEmailSubject disabled="disabled" class="text textdisabled" maxlength=255 name=txtEmailSubject style="WIDTH: 400px" 
                                                               onchange="frmUseful.txtChanged.value = 1;" 
                                                               onkeydown="frmUseful.txtChanged.value = 1;">
                                                    </td>
                                                    <td></td>
                                                </tr>
																	
                                                <tr height=10>
                                                    <td colspan=8></td>
                                                </tr>
																	
                                                <tr height=20>
                                                    <td colspan=3></td>
                                                    <td align=left nowrap>
                                                        Attach as :   
                                                    </td>
                                                    <td></td>
                                                    <td colspan=2 width=100% nowrap>
                                                        <INPUT id=txtEmailAttachAs disabled="disabled" maxlength=255 class="text textdisabled" name=txtEmailAttachAs style="WIDTH: 400px" 
                                                               onchange="frmUseful.txtChanged.value = 1;" 
                                                               onkeydown="frmUseful.txtChanged.value = 1;">
                                                    </td>
                                                    <td></td>
                                                </tr>
																	
                                                <tr height=10>
                                                    <td colspan=8></td>
                                                </tr>
                                            </TABLE>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </TABLE>
                </td>
            </tr>
        </TABLE></form>
</DIV>
    </td>
        <td width=10></td>
    </tr> 

        <tr height=10> 
            <td colspan=3></td>
        </tr> 

        <tr height=10>
            <td width=10></td>
            <td>
                <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                    <tr>
                        <td>&nbsp;</td>
                        <td width=80>
                            <input type=button id=cmdOK name=cmdOK value=OK style="WIDTH: 100%" class="btn"
                                   onclick="okClick()"
                                   onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                   onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                   onfocus="try{button_onFocus(this);}catch(e){}"
                                   onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width=10></td>
                        <td width=80>
                            <input type=button id=cmdCancel name=cmdCancel value=Cancel style="WIDTH: 100%"  class="btn"
                                   onclick="cancelClick()"
                                   onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                   onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                   onfocus="try{button_onFocus(this);}catch(e){}"
                                   onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                    </tr>
                </TABLE>
            </td>
            <td width=10></td>
        </tr> 

        <tr height=5> 
            <td colspan=3></td>
        </tr> 
    </table>
    </td>
    </tr> 
    </table>

        <input type='hidden' id=txtBasePicklistID name=txtBasePicklistID>
        <input type='hidden' id=txtBaseFilterID name=txtBaseFilterID>
        <input type='hidden' id=txtDatabase name=txtDatabase value="<%=session("Database")%>">

        <input type='hidden' id=txtWordVer name=txtWordVer value="<%=Session("WordVer")%>">
        <input type='hidden' id=txtExcelVer name=txtExcelVer value="<%=Session("ExcelVer")%>">
        <input type='hidden' id=txtWordFormats name=txtWordFormats value="<%=Session("WordFormats")%>">
        <input type='hidden' id=txtExcelFormats name=txtExcelFormats value="<%=Session("ExcelFormats")%>">
        <input type='hidden' id=txtWordFormatDefaultIndex name=txtWordFormatDefaultIndex value="<%=Session("WordFormatDefaultIndex")%>">
        <input type='hidden' id=txtExcelFormatDefaultIndex name=txtExcelFormatDefaultIndex value="<%=Session("ExcelFormatDefaultIndex")%>">

    </form>

    <form id=frmAccess>
        <%
            sErrorDescription = ""
	
            ' Get the table records.
            Dim cmdAccess = Server.CreateObject("ADODB.Command")
            cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
            cmdAccess.CommandType = 4 ' Stored Procedure
            cmdAccess.ActiveConnection = Session("databaseConnection")

            Dim prmUtilType = cmdAccess.CreateParameter("utilType", 3, 1) ' 3=integer, 1=input
            cmdAccess.Parameters.Append(prmUtilType)
            prmUtilType.value = 1 ' 1 = cross tabs

            Dim prmUtilID = cmdAccess.CreateParameter("utilID", 3, 1) ' 3=integer, 1=input
            cmdAccess.Parameters.Append(prmUtilID)
            If UCase(Session("action")) = "NEW" Then
                prmUtilID.value = 0
            Else
                prmUtilID.value = CleanNumeric(Session("utilid"))
            End If

            Dim prmFromCopy = cmdAccess.CreateParameter("fromCopy", 3, 1) ' 3=integer, 1=input
            cmdAccess.Parameters.Append(prmFromCopy)
            If UCase(Session("action")) = "COPY" Then
                prmFromCopy.value = 1
            Else
                prmFromCopy.value = 0
            End If

            Err.Clear()
            Dim rstAccessInfo = cmdAccess.Execute
            If (Err.Number <> 0) Then
                sErrorDescription = "The access information could not be retrieved." & vbCrLf & FormatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                Dim iCount = 0
                Do While Not rstAccessInfo.EOF
                    Response.Write("<INPUT type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & rstAccessInfo.fields("accessDefinition").value & """>" & vbCrLf)

                    iCount = iCount + 1
                    rstAccessInfo.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstAccessInfo.close()
                rstAccessInfo = Nothing
            End If
	
            ' Release the ADO command object.
            cmdAccess = Nothing
%>
    </form>

    <FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
        <INPUT type="hidden" id=txtUserName name=txtUserName value="<%=session("username")%>">
        <INPUT type="hidden" id=txtLoading name=txtLoading value="Y">
        <INPUT type="hidden" id=txtCurrentBaseTableID name=txtCurrentBaseTableID>
        <INPUT type="hidden" id=txtCurrentHColID name=txtCurrentHColID>
        <INPUT type="hidden" id=txtCurrentVColID name=txtCurrentVColID>
        <INPUT type="hidden" id=txtCurrentPColID name=txtCurrentPColID>
        <INPUT type="hidden" id=txtCurrentIColID name=txtCurrentIColID>
        <INPUT type="hidden" id=txtTablesChanged name=txtTablesChanged>
        <INPUT type="hidden" id=txtSelectedColumnsLoaded name=txtSelectedColumnsLoaded value=0>
        <INPUT type="hidden" id=txtSortLoaded name=txtSortLoaded value=0>
        <INPUT type="hidden" id=txtSecondTabShown name=txtSecondTabShown value=0>
        <INPUT type="hidden" id=txtRepetitionLoaded name=txtRepetitionLoaded value=0>
        <INPUT type="hidden" id=txtChanged name=txtChanged value=0>
        <INPUT type="hidden" id=txtUtilID name=txtUtilID value=<%=session("utilid")%>>
        <%
            Dim cmdDefinition = Server.CreateObject("ADODB.Command")
            cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
            cmdDefinition.CommandType = 4 ' Stored procedure.
            cmdDefinition.ActiveConnection = Session("databaseConnection")

            Dim prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
            cmdDefinition.Parameters.Append(prmModuleKey)
            prmModuleKey.value = "MODULE_PERSONNEL"

            Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
            cmdDefinition.Parameters.Append(prmParameterKey)
            prmParameterKey.value = "Param_TablePersonnel"

            Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000) '200=varchar, 2=output, 8000=size
            cmdDefinition.Parameters.Append(prmParameterValue)

            Err.Clear()
            cmdDefinition.Execute()

            Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
            cmdDefinition = Nothing

            Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
%>
    </FORM>

    <FORM id=frmValidate name=frmValidate target=validate method=post action=util_validate_crosstab style="visibility:hidden;display:none">
        <INPUT type=hidden id="validateBaseFilter" name=validateBaseFilter value=0>
        <INPUT type=hidden id="validateBasePicklist" name=validateBasePicklist value=0>
        <INPUT type=hidden id="validateEmailGroup" name=validateEmailGroup value=0>
        <INPUT type=hidden id="validateCalcs" name=validateCalcs value = ''>
        <INPUT type=hidden id="validateHiddenGroups" name=validateHiddenGroups value = ''>
        <INPUT type=hidden id="validateName" name=validateName value=''>
        <INPUT type=hidden id="validateTimestamp" name=validateTimestamp value=''>
        <INPUT type=hidden id="validateUtilID" name=validateUtilID value=''>
    </FORM>

    <FORM id=frmSend name=frmSend method=post action=util_def_crosstabs_Submit style="visibility:hidden;display:none">
        <INPUT type="hidden" id=txtSend_ID name=txtSend_ID value=0>
        <INPUT type="hidden" id=txtSend_name name=txtSend_name value=''>
        <INPUT type="hidden" id=txtSend_description name=txtSend_description value=''>
        <INPUT type="hidden" id=txtSend_baseTable name=txtSend_baseTable value=0>
        <INPUT type="hidden" id=txtSend_allRecords name=txtSend_allRecords value=0>
        <INPUT type="hidden" id=txtSend_picklist name=txtSend_picklist value=0>
        <INPUT type="hidden" id=txtSend_filter name=txtSend_filter value=0>
        <INPUT type="hidden" id=txtSend_PrintFilter name=txtSend_PrintFilter value=0>
        <INPUT type="hidden" id=txtSend_access name=txtSend_access value=''>
        <INPUT type="hidden" id=txtSend_userName name=txtSend_userName value=''>

        <INPUT type="hidden" id=txtSend_HColID name=txtSend_HColID value=0>
        <INPUT type="hidden" id=txtSend_HStart name=txtSend_HStart value=''>
        <INPUT type="hidden" id=txtSend_HStop name=txtSend_HStop value=''>
        <INPUT type="hidden" id=txtSend_HStep name=txtSend_HStep value=''>
        <INPUT type="hidden" id=txtSend_VColID name=txtSend_VColID value=0>
        <INPUT type="hidden" id=txtSend_VStart name=txtSend_VStart value=''>
        <INPUT type="hidden" id=txtSend_VStop name=txtSend_VStop value=''>
        <INPUT type="hidden" id=txtSend_VStep name=txtSend_VStep value=''>
        <INPUT type="hidden" id=txtSend_PColID name=txtSend_PColID value=0>
        <INPUT type="hidden" id=txtSend_PStart name=txtSend_PStart value=''>
        <INPUT type="hidden" id=txtSend_PStop name=txtSend_PStop value=''>
        <INPUT type="hidden" id=txtSend_PStep name=txtSend_PStep value=''>
        <INPUT type="hidden" id=txtSend_IType name=txtSend_IType value=0>
        <INPUT type="hidden" id=txtSend_IColID name=txtSend_IColID value=0>
        <INPUT type="hidden" id=txtSend_Percentage name=txtSend_Percentage value=0>
        <INPUT type="hidden" id=txtSend_PerPage name=txtSend_PerPage value=0>
        <INPUT type="hidden" id=txtSend_Suppress name=txtSend_Suppress value=0>
        <INPUT type="hidden" id=txtSend_Use1000Separator name=txtSend_Use1000Separator value=0>

        <INPUT type="hidden" id=txtSend_OutputPreview name=txtSend_OutputPreview>
        <INPUT type="hidden" id=txtSend_OutputFormat name=txtSend_OutputFormat>
        <INPUT type="hidden" id=txtSend_OutputScreen name=txtSend_OutputScreen>
        <INPUT type="hidden" id=txtSend_OutputPrinter name=txtSend_OutputPrinter>
        <INPUT type="hidden" id=txtSend_OutputPrinterName name=txtSend_OutputPrinterName>
        <INPUT type="hidden" id=txtSend_OutputSave name=txtSend_OutputSave>
        <INPUT type="hidden" id=txtSend_OutputSaveExisting name=txtSend_OutputSaveExisting>
        <INPUT type="hidden" id=txtSend_OutputEmail name=txtSend_OutputEmail>
        <INPUT type="hidden" id=txtSend_OutputEmailAddr name=txtSend_OutputEmailAddr>
        <INPUT type="hidden" id=txtSend_OutputEmailSubject name=txtSend_OutputEmailSubject>
        <INPUT type="hidden" id=txtSend_OutputEmailAttachAs name=txtSend_OutputEmailAttachAs>
        <INPUT type="hidden" id=txtSend_OutputFilename name=txtSend_OutputFilename>
	
        <INPUT type="hidden" id=txtSend_reaction name=txtSend_reaction>

        <INPUT type="hidden" id=txtSend_jobsToHide name=txtSend_jobsToHide>
        <INPUT type="hidden" id=txtSend_jobsToHideGroups name=txtSend_jobsToHideGroups>
    </FORM>

    <FORM id=frmRecordSelection name=frmRecordSelection target="recordSelection" action="util_recordSelection" method=post style="visibility:hidden;display:none">
        <INPUT type="hidden" id=recSelType name=recSelType>
        <INPUT type="hidden" id=recSelTableID name=recSelTableID>
        <INPUT type="hidden" id=recSelCurrentID name=recSelCurrentID>
        <INPUT type="hidden" id=recSelTable name=recSelTable>
        <INPUT type="hidden" id=recSelDefOwner name=recSelDefOwner>
        <INPUT type="hidden" id=recSelDefType name=recSelDefType>
    </FORM>

    <FORM id=frmEmailSelection name=frmEmailSelection target="emailSelection" action="util_emailSelection" method=post style="visibility:hidden;display:none">
        <INPUT type="hidden" id=EmailSelCurrentID name=EmailSelCurrentID>
    </FORM>

    <FORM id=frmSelectionAccess name=frmSelectionAccess style="visibility:hidden;display:none">
        <INPUT type="hidden" id=forcedHidden name=forcedHidden value="N">
        <INPUT type="hidden" id=baseHidden name=baseHidden value="N">

        <!-- need the count of hidden child filter access info -->
        <INPUT type="hidden" id=childHidden name=childHidden value="N">
        <INPUT type="hidden" id=calcsHiddenCount name=calcsHiddenCount value=0>
    </FORM>

    <FORM action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
        <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
    </FORM>

    <INPUT type='hidden' id=txtTicker name=txtTicker value=0>
    <INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<script type="text/javascript">
    util_def_crosstabs_window_onload();
    util_def_crosstabs_addhandlers();
</script>
