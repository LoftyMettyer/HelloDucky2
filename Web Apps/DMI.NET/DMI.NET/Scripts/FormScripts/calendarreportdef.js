var frmDefinition = document.getElementById("frmDefinition");
var frmUseful = document.getElementById("frmUseful");
var frmOriginalDefinition = document.getElementById("frmOriginalDefinition");
var frmSortOrder = document.getElementById("frmSortOrder");
var frmSelectionAccess = document.getElementById("frmSelectionAccess");
var frmSend = document.getElementById("frmSend");
var frmAccess = document.getElementById("frmAccess");
var frmValidate = document.getElementById("frmValidate");
var frmEmailSelection = document.getElementById("frmEmailSelection");
var frmTables = document.getElementById("frmTables");
var frmPopup = document.getElementById("frmPopup");
//var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");

var div1 = document.getElementById("div1");
var div2 = document.getElementById("div2");
var div3 = document.getElementById("div3");
var div4 = document.getElementById("div4");
var div5 = document.getElementById("div5");

var sRepDefn;
var sSortDefn;


function util_def_calendarreport_window_onload() {
	    var fOK;
	    fOK = true;
			var sErrMsg = frmUseful.txtErrorDescription.value;
	    if (sErrMsg.length > 0) {
	        fOK = false;
	        OpenHR.messageBox(sErrMsg, 48, "Calendar Reports");
	        //TODO
	        //window.parent.location.replace("login");
	    }

	    if (fOK == true) {
	        setGridFont(frmDefinition.grdAccess);
	        setGridFont(frmDefinition.grdEvents);
	        setGridFont(frmDefinition.ssOleDBGridSortOrder);

	        frmUseful.txtLoading.value = 'Y';

	        // Expand the work frame and hide the option frame.
	        //window.parent.document.all.item("workframeset").cols = "*, 0";
	        $("#workframe").attr("data-framesource", "UTIL_DEF_CALENDARREPORT");
			

	        frmDefinition.cboBaseTable.style.color = 'window';
	        frmDefinition.cboDescription1.style.color = 'window';
	        frmDefinition.cboDescription2.style.color = 'window';
	        frmDefinition.cboRegion.style.color = 'window';

	        populateBaseTableCombo();

	        if (frmUseful.txtAction.value.toUpperCase() == "NEW") {

	            
	               frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
	               frmDefinition.txtDescription.value = "";
	               frmDefinition.txtName.value = "";
	               frmDefinition.optRecordSelection1.checked = true;
	               frmDefinition.optFixedStart.checked = true;
	               frmDefinition.optFixedEnd.checked = true;
	               frmDefinition.chkShadeWeekends.checked = true;
	               frmDefinition.chkStartOnCurrentMonth.checked = true;
	               frmDefinition.chkCaptions.checked = true;
	               frmDefinition.optOutputFormat0.checked = true;
	               frmDefinition.chkDestination0.checked = true;
	            

	            setBaseTable(0);
	            changeBaseTable();

	            frmUseful.txtEventsLoaded.value = 1;
	            frmUseful.txtSortLoaded.value = 1;
	            //frmUseful.txtChanged.value = 1;
	        } else {
	            loadDefinition();
	        }

	        frmUseful.txtAvailableColumnsLoaded.value = 0;

	        populateBaseTableColumns();

	        populateAccessGrid();

	        if ((frmUseful.txtAction.value.toUpperCase() == "EDIT")
				|| (frmUseful.txtAction.value.toUpperCase() == "VIEW")) {
	            frmUseful.txtChanged.value = 0;
	        } else {
	            frmUseful.txtUtilID.value = 0;
	            //frmUseful.txtChanged.value = 1;
	        }

	        if (frmUseful.txtAction.value.toUpperCase() == "COPY") {
	            frmUseful.txtChanged.value = 1;
	        }

	        refreshTab5Controls();
	        displayPage(1);


	        if (frmDefinition.chkDestination1.checked == true) {
	            if (frmOriginalDefinition.txtDefn_OutputPrinterName.value != "") {
	                if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.selectedIndex).innerText != frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
	                    OpenHR.messageBox("This definition is set to output to printer " + frmOriginalDefinition.txtDefn_OutputPrinterName.value + " which is not set up on your PC.");
	                    var oOption = document.createElement("OPTION");
	                    frmDefinition.cboPrinterName.options.add(oOption);
	                    oOption.innerText = frmOriginalDefinition.txtDefn_OutputPrinterName.value;
	                    oOption.value = frmDefinition.cboPrinterName.options.length - 1;
	                    frmDefinition.cboPrinterName.selectedIndex = oOption.value;
	                }
	            }
	        }
	    }
	}

var fValidating;
fValidating = false;

function displayPage(piPageNumber) {

    var iLoop;
		////debugger;
    //window.parent.frames("refreshframe").document.forms("frmRefresh").submit();
    OpenHR.submitForm(window.frmRefresh);

    if (piPageNumber == 1) {
        var lngGridHeight = new Number(0);
        var lngGridWidth = new Number(0);

        div1.style.visibility = "visible";
        div1.style.display = "block";
        div2.style.visibility = "hidden";
        div2.style.display = "none";
        div3.style.visibility = "hidden";
        div3.style.display = "none";
        div4.style.visibility = "hidden";
        div4.style.display = "block";
        lngGridHeight = frmDefinition.ssOleDBGridSortOrder.style.height;
        lngGridWidth = frmDefinition.ssOleDBGridSortOrder.style.width;
        frmDefinition.ssOleDBGridSortOrder.style.height = 0;
        frmDefinition.ssOleDBGridSortOrder.style.width = 0;
        loadSortDefinition();
        frmDefinition.ssOleDBGridSortOrder.style.height = lngGridHeight;
        frmDefinition.ssOleDBGridSortOrder.style.width = lngGridWidth;
        div4.style.visibility = "hidden";
        div4.style.display = "none";
        div5.style.visibility = "hidden";
        div5.style.display = "none";

        button_disable(frmDefinition.btnTab1, true);
        button_disable(frmDefinition.btnTab2, false);
        button_disable(frmDefinition.btnTab3, false);
        button_disable(frmDefinition.btnTab4, false);
        button_disable(frmDefinition.btnTab5, false);

        refreshTab1Controls();

        if (frmDefinition.txtName.disabled == false) {
            try {
                frmDefinition.txtName.focus();
            }
            catch (e) { }
        }

    }

    if (frmUseful.txtAvailableColumnsLoaded.value != 1) {
        return;
    }

    if (piPageNumber == 2) {
        div1.style.visibility = "hidden";
        div1.style.display = "none";
        div2.style.visibility = "visible";
        div2.style.display = "block";
        div3.style.visibility = "hidden";
        div3.style.display = "none";
        div4.style.visibility = "hidden";
        div4.style.display = "none";
        div5.style.visibility = "hidden";
        div5.style.display = "none";

        button_disable(frmDefinition.btnTab1, false);
        button_disable(frmDefinition.btnTab2, true);
        button_disable(frmDefinition.btnTab3, false);
        button_disable(frmDefinition.btnTab4, false);
        button_disable(frmDefinition.btnTab5, false);

        loadEventsDefinition();

        frmDefinition.grdEvents.SelBookmarks.RemoveAll();
        if (frmDefinition.grdEvents.Rows > 0) {
            frmDefinition.grdEvents.MoveFirst();
            frmDefinition.grdEvents.SelBookmarks.Add(frmDefinition.grdEvents.Bookmark);
        }

        if (frmDefinition.grdEvents.Enabled) {
            frmDefinition.grdEvents.focus();
        }

        refreshTab2Controls();
    }

    if (piPageNumber == 3) {
        div1.style.visibility = "hidden";
        div1.style.display = "none";
        div2.style.visibility = "hidden";
        div2.style.display = "none";
        div3.style.visibility = "visible";
        div3.style.display = "block";
        div4.style.visibility = "hidden";
        div4.style.display = "none";
        div5.style.visibility = "hidden";
        div5.style.display = "none";

        button_disable(frmDefinition.btnTab1, false);
        button_disable(frmDefinition.btnTab2, false);
        button_disable(frmDefinition.btnTab3, true);
        button_disable(frmDefinition.btnTab4, false);
        button_disable(frmDefinition.btnTab5, false);

        refreshTab3Controls();
    }

    if (piPageNumber == 4) {
        div1.style.visibility = "hidden";
        div1.style.display = "none";
        div2.style.visibility = "hidden";
        div2.style.display = "none";
        div3.style.visibility = "hidden";
        div3.style.display = "none";
        div4.style.visibility = "visible";
        div4.style.display = "block";
        div5.style.visibility = "hidden";
        div5.style.display = "none";

        button_disable(frmDefinition.btnTab1, false);
        button_disable(frmDefinition.btnTab2, false);
        button_disable(frmDefinition.btnTab3, false);
        button_disable(frmDefinition.btnTab4, true);
        button_disable(frmDefinition.btnTab5, false);

        loadSortDefinition();

        frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();
        if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
            frmDefinition.ssOleDBGridSortOrder.MoveFirst();
            frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
        }

        if (frmDefinition.ssOleDBGridSortOrder.Enabled) {
            frmDefinition.ssOleDBGridSortOrder.focus();
        }

        refreshTab4Controls();
    }

    if (piPageNumber == 5) {
        div1.style.visibility = "hidden";
        div1.style.display = "none";
        div2.style.visibility = "hidden";
        div2.style.display = "none";
        div3.style.visibility = "hidden";
        div3.style.display = "none";
        div4.style.visibility = "hidden";
        div4.style.display = "none";
        div5.style.visibility = "visible";
        div5.style.display = "block";

        button_disable(frmDefinition.btnTab1, false);
        button_disable(frmDefinition.btnTab2, false);
        button_disable(frmDefinition.btnTab3, false);
        button_disable(frmDefinition.btnTab4, false);
        button_disable(frmDefinition.btnTab5, true);

        refreshTab5Controls();
    }
}

function populateBaseTableCombo() {
    var i;

    //Clear the existing data in the base table combo
    while (frmDefinition.cboBaseTable.options.length > 0) {
    	frmDefinition.cboBaseTable.options.remove(0);
    }

    var dataCollection = frmTables.elements;
    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {
            var sControlName = dataCollection.item(i).name;
            var sControlTag = sControlName.substr(0, 13);
            if (sControlTag == "txtTableName_") {
                var sTableID = sControlName.substr(13);
                var oOption = document.createElement("OPTION");
                frmDefinition.cboBaseTable.options.add(oOption);
                oOption.innerText = dataCollection.item(i).value;
                oOption.value = sTableID;
            }
        }
    }
}

function populateBaseTableColumns() {
    // Get the columns/calcs for the current table selection.
    //var frmGetDataForm = window.parent.frames("dataframe").document.forms("frmGetData");
    var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");

    frmUseful.txtAvailableColumnsLoaded.value = 0;

    frmGetDataForm.txtAction.value = "LOADCALENDARREPORTCOLUMNS";
    frmGetDataForm.txtReportBaseTableID.value = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;

    //this should be in scope by now.
	//TODO: NPG
   data_refreshData(); //window.parent.frames("dataframe").refreshData();
	}

function trim(strInput) {
    if (strInput.length < 1) {
        return "";
    }

    while (strInput.substr(strInput.length - 1, 1) == " ") {
        strInput = strInput.substr(0, strInput.length - 1);
    }

    while (strInput.substr(0, 1) == " ") {
        strInput = strInput.substr(1, strInput.length);
    }

    return strInput;
}

function setBaseTable(piTableID) {
    var i;

    if (piTableID == 0) piTableID = frmUseful.txtPersonnelTableID.value;

    if (piTableID > 0) {
        for (i = 0; i < frmDefinition.cboBaseTable.options.length; i++) {
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

function changeBaseTable() {
    var i;

    if (frmUseful.txtLoading.value == 'N') {

        if ((frmDefinition.grdEvents.Rows > 0) ||
            ((frmUseful.txtAction.value.toUpperCase() != "NEW") &&
                (frmUseful.txtEventsLoaded.value == 0))) {

            var iAnswer = OpenHR.messageBox("Warning: Changing the base table will result in all table/column specific aspects of this report definition being cleared. Are you sure you wish to continue?", 36, "Calendar Reports");
            if (iAnswer == 7) {
                // cancel and change back ! (txtcurrentbasetable)
                setBaseTable(frmUseful.txtCurrentBaseTableID.value);
                return;
            }
            else {
                frmUseful.txtEventsLoaded.value = 1;
                frmUseful.txtSortLoaded.value = 1;
                frmUseful.txtChanged.value = 1;
            }
        }
        else {
            frmUseful.txtChanged.value = 1;
        }
    }

    clearBaseTableRecordOptions();

    frmDefinition.cboDescription1.length = 0;
    frmDefinition.cboDescription2.length = 0;
    frmDefinition.txtDescExpr.value = '';
    frmDefinition.txtDescExprID.value = 0;
    frmSelectionAccess.descHidden.value = "N";
    frmDefinition.chkGroupByDesc.checked = false;
    frmDefinition.cboDescriptionSeparator.selectedIndex = 0;

    frmDefinition.cboRegion.length = 0;

    while (frmDefinition.grdEvents.Rows > 0) {
        frmDefinition.grdEvents.RemoveAll();
    }

    while (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
        frmDefinition.ssOleDBGridSortOrder.RemoveAll();
    }

    if (frmUseful.txtLoading.value == 'N') {
        populateBaseTableColumns();
    }

    frmDefinition.chkShadeBHols.checked = false;
    frmDefinition.chkIncludeBHols.checked = false;
    frmDefinition.chkIncludeWorkingDaysOnly.checked = false;
    frmDefinition.chkShadeWeekends.checked = true;
    frmDefinition.chkCaptions.checked = true;
    frmDefinition.chkStartOnCurrentMonth.checked = true;

    var sRelationNames = new String("");

    var dataCollection = frmTables.elements;
    if (dataCollection != null) {
        sReqdControlName = new String("txtTableRelations_");
        sReqdControlName = sReqdControlName.concat(frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);

        for (i = 0; i < dataCollection.length; i++) {
            sControlName = dataCollection.item(i).name;
            if (sControlName == sReqdControlName) {
                sRelationNames = dataCollection.item(i).value;
                frmEventDetails.relationNames.value = sRelationNames;
                break;
            }
        }
    }

    recalcHiddenEventFiltersCount();

    if (frmUseful.txtLoading.value != 'N') {
        refreshTab1Controls();
    }

    frmUseful.txtCurrentBaseTableID.value = frmDefinition.cboBaseTable.options(frmDefinition.cboBaseTable.options.selectedIndex).value;
    frmUseful.txtAvailableColumnsLoaded.value = 0;
}

function refreshTab1Controls() {
    var fIsForcedHidden;
    var fViewing;
    var fIsNotOwner;
    var fAllAlreadyHidden;
    var fSilent;

    fSilent = ((frmUseful.txtAction.value.toUpperCase() == "COPY") &&
        (frmUseful.txtLoading.value == "Y"));

    if (frmUseful.txtAvailableColumnsLoaded.value == 0) {
        return;
    }

    fIsForcedHidden = ((frmSelectionAccess.baseHidden.value == "Y") ||
        (frmSelectionAccess.descHidden.value == "Y") ||
        (frmSelectionAccess.eventHidden.value > 0) ||
        (frmSelectionAccess.calcStartDateHidden.value == "Y") ||
        (frmSelectionAccess.calcEndDateHidden.value == "Y"));

    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());
    fAllAlreadyHidden = AllHiddenAccess(frmDefinition.grdAccess);

    if (fIsForcedHidden == true) {
        if (fAllAlreadyHidden != true) {
            if (fSilent == false) {
                OpenHR.messageBox("This definition will now be made hidden as it contains a hidden picklist/filter/calculation.", 64);
            }
            ForceAccess(frmDefinition.grdAccess, "HD");
            frmUseful.txtChanged.value = 1;
        }
        else {
            if (frmSelectionAccess.forcedHidden.value == "N") {
                //MH20040816 Fault 9047
                //if (fSilent == false) {
                if ((fSilent == false) && (frmUseful.txtLoading.value != "Y")) {
                    OpenHR.messageBox("The definition access cannot be changed as it contains a hidden picklist/filter/calculation.", 64);
                }
            }
        }
        frmSelectionAccess.forcedHidden.value = "Y";

        frmDefinition.grdAccess.Columns("Access").Style = 0; // 0 = Edit
    }
    else {
        try {
            window.resizeBy(0, -1);
            window.resizeBy(0, 1);
            window.resizeBy(0, -1);
            window.resizeBy(0, 1);
        }
        catch (e) { }
        if (frmSelectionAccess.forcedHidden.value == "Y") {
            frmSelectionAccess.forcedHidden.value = "N";
            // No longer forced hidden.
            if (fSilent == false) {
                OpenHR.messageBox("This definition no longer has to be hidden.", 64);
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
    combo_disable(frmDefinition.cboDescription1, (fViewing == true));
    combo_disable(frmDefinition.cboDescription2, (fViewing == true));
    button_disable(frmDefinition.cmdDescExpr, (fViewing == true));
    combo_disable(frmDefinition.cboRegion, (fViewing == true));

    if (frmDefinition.cboDescription1.selectedIndex < 0) frmDefinition.cboDescription1.selectedIndex = 0;
    if (frmDefinition.cboDescription2.selectedIndex < 0) frmDefinition.cboDescription2.selectedIndex = 0;

    var iDescCount = new Number(0);
    if (frmDefinition.cboDescription1.selectedIndex > 0) iDescCount++;
    if (frmDefinition.cboDescription2.selectedIndex > 0) iDescCount++;
    if (frmDefinition.txtDescExprID.value > 0) iDescCount++;
    if ((iDescCount < 2) || (frmDefinition.cboDescriptionSeparator.selectedIndex < 0)) {
        frmDefinition.cboDescriptionSeparator.selectedIndex = 0;
    }

    combo_disable(frmDefinition.cboDescriptionSeparator, ((fViewing == true) || (iDescCount < 2)));

    if (frmDefinition.cboRegion.selectedIndex < 0) frmDefinition.cboRegion.selectedIndex = 0;

    button_disable(frmDefinition.cmdBasePicklist, ((frmDefinition.optRecordSelection2.checked == false)
        || (fViewing == true)));
    button_disable(frmDefinition.cmdBaseFilter, ((frmDefinition.optRecordSelection3.checked == false)
        || (fViewing == true)));

    if (frmDefinition.optRecordSelection2.checked || frmDefinition.optRecordSelection3.checked) {
        checkbox_disable(frmDefinition.chkPrintFilterHeader, (fViewing == true));
    }
    else {
        frmDefinition.chkPrintFilterHeader.checked = false;
        checkbox_disable(frmDefinition.chkPrintFilterHeader, true);
    }

    with (frmDefinition) {
        if (chkIncludeBHols.checked || chkIncludeWorkingDaysOnly.checked || chkShadeBHols.checked
            || (cboRegion.options[cboRegion.selectedIndex].value > 0)) {
            checkbox_disable(chkGroupByDesc, true);
        }
        else {
            checkbox_disable(chkGroupByDesc, fViewing);
        }

        if (chkGroupByDesc.checked) {
            checkbox_disable(chkIncludeBHols, true);
            checkbox_disable(chkIncludeWorkingDaysOnly, true);
            checkbox_disable(chkShadeBHols, true);
            combo_disable(cboRegion, true);
        }
        else {
            checkbox_disable(chkIncludeBHols, fViewing);
            checkbox_disable(chkIncludeWorkingDaysOnly, fViewing);
            checkbox_disable(chkShadeBHols, fViewing);
            combo_disable(cboRegion, fViewing);
        }

        checkbox_disable(chkCaptions, fViewing);
        checkbox_disable(chkShadeWeekends, fViewing);
        checkbox_disable(chkStartOnCurrentMonth, fViewing);
    }

    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
        (fViewing == true)));

}

function refreshTab2Controls() {
	var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

	button_disable(frmDefinition.cmdAddEvent, (fViewing == true));
	button_disable(frmDefinition.cmdEditEvent, ((frmDefinition.grdEvents.Rows < 1)
			|| (frmDefinition.grdEvents.SelBookmarks.Count != 1)
			|| (fViewing == true)));
	button_disable(frmDefinition.cmdRemoveEvent, ((frmDefinition.grdEvents.Rows < 1)
			|| (frmDefinition.grdEvents.SelBookmarks.Count != 1)
			|| (fViewing == true)));
	button_disable(frmDefinition.cmdRemoveAllEvents, ((frmDefinition.grdEvents.Rows < 1)
			|| (fViewing == true)));
	
	recalcHiddenEventFiltersCount();
	refreshTab1Controls();
	frmDefinition.grdEvents.RowHeight = 19;
}

function refreshTab3Controls() {
	var frmUse = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");
	//var frmPopup = document.getElementById("frmPopup");
  var fViewing = (frmUse.txtAction.value.toUpperCase() == "VIEW");

    with (frmDefinition) {
        if (optFixedStart.checked) {
            text_disable(txtFixedStart, fViewing);

            text_disable(txtFreqStart, true);
            txtFreqStart.value = '';
            button_disable(cmdPeriodStartDown, true);
            button_disable(cmdPeriodStartUp, true);

            combo_disable(cboPeriodStart, true);
            cboPeriodStart.selectedIndex = -1;

            button_disable(cmdCustomStart, true);
            txtCustomStart.value = '';
            txtCustomStartID.value = 0;

            radio_disable(optFixedEnd, fViewing);
            radio_disable(optOffsetEnd, fViewing);
            radio_disable(optCurrentEnd, fViewing);
        }
        else if (optCurrentStart.checked) {
            text_disable(txtFixedStart, true);
            txtFixedStart.value = '';

            text_disable(txtFreqStart, true);
            txtFreqStart.value = '';
            button_disable(cmdPeriodStartDown, true);
            button_disable(cmdPeriodStartUp, true);

            combo_disable(cboPeriodStart, true);
            cboPeriodStart.selectedIndex = -1;

            radio_disable(optFixedEnd, fViewing);
            radio_disable(optOffsetEnd, fViewing);
            radio_disable(optCurrentEnd, fViewing);

            button_disable(cmdCustomStart, true);
            txtCustomStart.value = '';
            txtCustomStartID.value = 0;
        }
        else if (optOffsetStart.checked) {
            text_disable(txtFixedStart, true);
            txtFixedStart.value = '';

            if (Number(txtFreqStart.value) > 0) {
                optFixedEnd.checked = false;
                radio_disable(optFixedEnd, true);

                text_disable(txtFixedEnd, true);
                txtFixedEnd.value = '';

                optCurrentEnd.checked = false;
                radio_disable(optCurrentEnd, true);
                //optOffsetEnd.checked = true;
            }
            else {
                radio_disable(optFixedEnd, fViewing);
                text_disable(txtFixedEnd, ((fViewing) || (optFixedEnd.checked == false)));
                radio_disable(optCurrentEnd, fViewing);
            }

            text_disable(txtFreqStart, fViewing);
            if (txtFreqStart.value == '') {
                txtFreqStart.value = 0;
            }
            button_disable(cmdPeriodStartDown, fViewing);
            button_disable(cmdPeriodStartUp, fViewing);
            combo_disable(cboPeriodStart, fViewing);
            if (cboPeriodStart.selectedIndex < 0) {
                cboPeriodStart.selectedIndex = 0;
            }

            radio_disable(optOffsetEnd, fViewing);
            text_disable(txtFreqEnd, ((fViewing) || (optOffsetEnd.checked == false)));
            if (txtFreqEnd.value == '') {
                txtFreqEnd.value = 0;
            }
            button_disable(cmdPeriodEndDown, ((fViewing) || (optOffsetEnd.checked == false)));
            button_disable(cmdPeriodEndUp, ((fViewing) || (optOffsetEnd.checked == false)));

            button_disable(cmdCustomStart, true);
            txtCustomStart.value = '';
            txtCustomStartID.value = 0;
        }
        else if (optCustomStart.checked) {
            text_disable(txtFixedStart, true);
            txtFixedStart.value = '';

            text_disable(txtFreqStart, true);
            txtFreqStart.value = '';
            button_disable(cmdPeriodStartDown, true);
            button_disable(cmdPeriodStartUp, true);

            combo_disable(cboPeriodStart, true);
            cboPeriodStart.selectedIndex = -1;

            radio_disable(optOffsetEnd, fViewing);
            radio_disable(optCurrentEnd, fViewing);

            button_disable(cmdCustomStart, fViewing);
        }

        if (optFixedEnd.checked) {
            text_disable(txtFixedEnd, fViewing);

            text_disable(txtFreqEnd, true);
            txtFreqEnd.value = '';

            button_disable(cmdPeriodEndDown, true);
            button_disable(cmdPeriodEndUp, true);
            combo_disable(cboPeriodEnd, true);
            cboPeriodEnd.selectedIndex = -1;

            radio_disable(optCurrentStart, fViewing);
            radio_disable(optOffsetStart, fViewing);

            button_disable(cmdCustomEnd, true);
            txtCustomEnd.value = '';
            txtCustomEndID.value = 0;
        }
        else if (optCurrentEnd.checked) {
            text_disable(txtFixedEnd, true);
            txtFixedEnd.value = '';

            text_disable(txtFreqEnd, true);
            txtFreqEnd.value = '';
            button_disable(cmdPeriodEndDown, true);
            button_disable(cmdPeriodEndUp, true);
            combo_disable(cboPeriodEnd, true);
            cboPeriodEnd.selectedIndex = -1;

            radio_disable(optOffsetStart, fViewing);
            radio_disable(optCurrentStart, fViewing);

            button_disable(cmdCustomEnd, true);
            txtCustomEnd.value = '';
            txtCustomEndID.value = 0;
        }
        else if (optOffsetEnd.checked) {
            text_disable(txtFixedEnd, true);
            txtFixedEnd.value = '';

            text_disable(txtFreqEnd, fViewing);
            if (txtFreqEnd.value == '') {
                txtFreqEnd.value = 0;
            }
            button_disable(cmdPeriodEndDown, fViewing);
            button_disable(cmdPeriodEndUp, fViewing);
            combo_disable(cboPeriodEnd, fViewing);
            if (cboPeriodEnd.selectedIndex < 0) {
                cboPeriodEnd.selectedIndex = 0;
            }

            radio_disable(optOffsetStart, fViewing);
            text_disable(txtFreqStart, ((fViewing) || (optOffsetStart.checked == false)));
            combo_disable(cboPeriodStart, ((fViewing) || (optOffsetStart.checked == false)));
            radio_disable(optCurrentStart, fViewing);

            button_disable(cmdCustomEnd, true);
            txtCustomEnd.value = '';
            txtCustomEndID.value = 0;
        }
        else if (optCustomEnd.checked) {
            text_disable(txtFixedEnd, true);
            txtFixedEnd.value = '';

            text_disable(txtFreqEnd, true);
            txtFreqEnd.value = '';
            button_disable(cmdPeriodEndDown, true);
            button_disable(cmdPeriodEndUp, true);
            combo_disable(cboPeriodEnd, true);
            cboPeriodEnd.selectedIndex = -1;

            radio_disable(optOffsetStart, fViewing);
            radio_disable(optCurrentStart, fViewing);

            button_disable(cmdCustomEnd, fViewing);
        }

        if (txtFreqStart.disabled) {
            button_disable(cmdPeriodStartDown, true);
            button_disable(cmdPeriodStartUp, true);
        }
        else {
            button_disable(cmdPeriodStartDown, fViewing);
            button_disable(cmdPeriodStartUp, fViewing);
        }

        if (txtFreqEnd.disabled) {
            button_disable(cmdPeriodEndDown, true);
            button_disable(cmdPeriodEndUp, true);
        }
        else {
            button_disable(cmdPeriodEndDown, fViewing);
            button_disable(cmdPeriodEndUp, fViewing);
        }

        var blnPersonnelBaseTable = (frmUse.txtPersonnelTableID.value == frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);
	    var blnRegionSelected = (cboRegion.options[cboRegion.selectedIndex].value > 0);

        if (chkIncludeBHols.checked || chkIncludeWorkingDaysOnly.checked || chkShadeBHols.checked
            || blnRegionSelected) {
            checkbox_disable(chkGroupByDesc, true);
        }
        else {
            checkbox_disable(chkGroupByDesc, fViewing);
        }

        if (chkGroupByDesc.checked) {
            checkbox_disable(chkIncludeBHols, true);
            checkbox_disable(chkIncludeWorkingDaysOnly, true);
            checkbox_disable(chkShadeBHols, true);
            combo_disable(cboRegion, true);
        }
        else {
            checkbox_disable(chkIncludeBHols, ((fViewing) || ((!blnPersonnelBaseTable) && (!blnRegionSelected))));
            checkbox_disable(chkIncludeWorkingDaysOnly, ((fViewing) || (!blnPersonnelBaseTable)));
            checkbox_disable(chkShadeBHols, ((fViewing) || ((!blnPersonnelBaseTable) && (!blnRegionSelected))));
            combo_disable(cboRegion, fViewing);
        }

        checkbox_disable(chkCaptions, fViewing);
        checkbox_disable(chkShadeWeekends, fViewing);
        checkbox_disable(chkStartOnCurrentMonth, fViewing);

        if (!blnPersonnelBaseTable) {
            chkIncludeWorkingDaysOnly.checked = false;
            if (!blnRegionSelected) {
                chkIncludeBHols.checked = false;
                chkShadeBHols.checked = false;
            }
        }
    }

    if (frmDefinition.optCustomStart.checked == false) {
        frmSelectionAccess.calcStartDateHidden.value = "N";
    }
    if (frmDefinition.optCustomEnd.checked == false) {
        frmSelectionAccess.calcEndDateHidden.value = "N";
    }

    refreshTab1Controls();

    button_disable(frmDefinition.cmdOK, ((frmUsel.txtChanged.value == 0) ||
        (fViewing == true)));
}

function refreshTab4Controls() {
	var i;
	var iCount;
	var fSortAddDisabled = false;
	var fSortEditDisabled = false;
	var fSortRemoveDisabled = false;
	var fSortRemoveAllDisabled = false;
	var fSortMoveUpDisabled = false;
	var fSortMoveDownDisabled = false;
	var fViewing;

	fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

	if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count == 1
			&& frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
		// Are we on the top row ?
		if ((frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) == 0)
				|| (frmDefinition.ssOleDBGridSortOrder.rows <= 1)) {
			fSortMoveUpDisabled = true;
		}

		// Are we on the bottom row ?
		if ((frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) == frmDefinition.ssOleDBGridSortOrder.rows - 1)
				|| (frmDefinition.ssOleDBGridSortOrder.rows <= 1)) {
			fSortMoveDownDisabled = true;
		}
	}

	if (frmDefinition.ssOleDBGridSortOrder.Rows < 1
			|| frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count != 1) {
		fSortMoveUpDisabled = true;
		fSortMoveDownDisabled = true;
	}

	if (fViewing) {
		fSortAddDisabled = true;
		fSortMoveUpDisabled = true;
		fSortMoveDownDisabled = true;
	}

	fSortRemoveDisabled = ((frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count != 1) || (fViewing == true));
	fSortRemoveAllDisabled = ((fViewing == true) || (frmDefinition.ssOleDBGridSortOrder.Rows < 1));
	fSortEditDisabled = ((frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count != 1) || (fViewing == true));

	button_disable(frmDefinition.cmdSortAdd, fSortAddDisabled);
	button_disable(frmDefinition.cmdSortEdit, fSortEditDisabled);
	button_disable(frmDefinition.cmdSortRemove, fSortRemoveDisabled);
	button_disable(frmDefinition.cmdSortRemoveAll, fSortRemoveAllDisabled);
	button_disable(frmDefinition.cmdSortMoveUp, fSortMoveUpDisabled);
	button_disable(frmDefinition.cmdSortMoveDown, fSortMoveDownDisabled);

	frmDefinition.ssOleDBGridSortOrder.AllowUpdate = (fViewing == false);
	frmDefinition.ssOleDBGridSortOrder.RowHeight = 19;
	button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
			(fViewing == true)));

}

function refreshTab5Controls() {
    var i;
    var iCount;
	
    var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    with (frmDefinition) {
        if (optOutputFormat0.checked == true)		//Data Only
        {
            //disable preview opitons
            chkPreview.checked = false;
            checkbox_disable(chkPreview, true);

            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));

            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));
            if (chkDestination1.checked == true) {
                populatePrinters();
                combo_disable(cboPrinterName, (fViewing == true));
            }
            else {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName, true);
            }

            //disable save options
            chkDestination2.checked = false;
	        checkbox_disable(chkDestination2, true);
            combo_disable(cboSaveExisting, true);
            cboSaveExisting.length = 0;

            //disable email options
            chkDestination3.checked = false;
            checkbox_disable(chkDestination3, true);
            text_disable(txtEmailGroup, true);
            txtEmailGroup.value = '';
            txtEmailGroupID.value = 0;
            text_disable(txtEmailSubject, true);
            txtEmailSubject.value = '';
            text_disable(cmdEmailGroup, true);
            txtEmailAttachAs.value = '';
            text_disable(txtEmailAttachAs, true);

            //disable filename options
            txtFilename.value = '';
            text_disable(txtFilename, true);
            button_disable(cmdFilename, true);
        }
            /*else if (optOutputFormat1.checked == true)   //CSV File
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
                checkbox_disable(chkDestination2, (fViewing == true));
                if (chkDestination2.checked == true)
                    {
                    populateSaveExisting();
                    combo_disable(cboSaveExisting, (fViewing == true));
                    }	
                else
                    {
                    cboSaveExisting.length = 0;
                    combo_disable(cboSaveExisting, true);
                    }
                
                //enable-disable email options
                checkbox_disable(chkDestination3, (fViewing == true));
                if (chkDestination3.checked == true)
                    {
                    text_disable(txtEmailGroup, (fViewing == true));
                    text_disable(txtEmailSubject, (fViewing == true));
                    button_disable(cmdEmailGroup, (fViewing == true));
                    text_disable(txtEmailAttachAs, (fViewing == true));
                    }
                else
                    {
                    text_disable(txtEmailGroup, true);
                    txtEmailGroup.value = '';
                    txtEmailGroupID.value = 0;
                    text_disable(txtEmailSubject, true);
                    txtEmailSubject.value = '';
                    button_disable(cmdEmailGroup, true);
                    text_disable(txtEmailAttachAs, true);
                    }
    
                //enable-disable filename options
                text_disable(txtFilename, ((fViewing == true) || (!chkDestination2.checked)));
                button_disable(cmdFilename, ((fViewing == true) || (!chkDestination2.checked)));
                }*/
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
            checkbox_disable(chkDestination2, (fViewing == true));
            if (chkDestination2.checked == true) {
                populateSaveExisting();
                combo_disable(cboSaveExisting, (fViewing == true));
            }
            else {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, (fViewing == true));
            if (chkDestination3.checked == true) {
                text_disable(txtEmailGroup, (fViewing == true));
                text_disable(txtEmailSubject, (fViewing == true));
                button_disable(cmdEmailGroup, (fViewing == true));
                text_disable(txtEmailAttachAs, (fViewing == true));
            }
            else {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                text_disable(txtEmailSubject, true);
                txtEmailSubject.value = '';
                button_disable(cmdEmailGroup, true);
                txtEmailAttachAs.value = '';
                text_disable(txtEmailAttachAs, true);
            }

            //enable-disable filename options
            text_disable(txtFilename, ((fViewing == true) || (!chkDestination2.checked)));
            button_disable(cmdFilename, ((fViewing == true) || (!chkDestination2.checked)));
        }
        else if (optOutputFormat3.checked == true)		//Word Document
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));

            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));

            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));
            if (chkDestination1.checked == true) {
                populatePrinters();
                combo_disable(cboPrinterName, (fViewing == true));
            }
            else {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName, true);
            }

            //enable-disable save options
            checkbox_disable(chkDestination2, (fViewing == true));
            if (chkDestination2.checked == true) {
                populateSaveExisting();
                combo_disable(cboSaveExisting, (fViewing == true));
            }
            else {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, (fViewing == true));
            if (chkDestination3.checked == true) {
                text_disable(txtEmailGroup, (fViewing == true));
                text_disable(txtEmailSubject, (fViewing == true));
                button_disable(cmdEmailGroup, (fViewing == true));
                text_disable(txtEmailAttachAs, (fViewing == true));
            }
            else {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                text_disable(txtEmailSubject, true);
                txtEmailSubject.value = '';
                button_disable(cmdEmailGroup, true);
                txtEmailAttachAs.value = '';
                text_disable(txtEmailAttachAs, true);
            }

            //enable-disable filename options
            text_disable(txtFilename, ((fViewing == true) || (!chkDestination2.checked)));
            button_disable(cmdFilename, ((fViewing == true) || (!chkDestination2.checked)));
        }
        else if (optOutputFormat4.checked == true)		//Excel Worksheet
        {
            //enable preview options
            checkbox_disable(chkPreview, (fViewing == true));

            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));

            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));
            if (chkDestination1.checked == true) {
                populatePrinters();
                combo_disable(cboPrinterName, (fViewing == true));
            }
            else {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName, true);
            }

            //enable-disable save options
            checkbox_disable(chkDestination2, (fViewing == true));
            if (chkDestination2.checked == true) {
                populateSaveExisting();
                combo_disable(cboSaveExisting, (fViewing == true));
            }
            else {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, (fViewing == true));
            if (chkDestination3.checked == true) {
                text_disable(txtEmailGroup, (fViewing == true));
                text_disable(txtEmailSubject, (fViewing == true));
                button_disable(cmdEmailGroup, (fViewing == true));
                text_disable(txtEmailAttachAs, (fViewing == true));
            }
            else {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                text_disable(txtEmailSubject, true);
                txtEmailSubject.value = '';
                button_disable(cmdEmailGroup, true);
                txtEmailAttachAs.value = '';
                text_disable(txtEmailAttachAs, true);
            }

            //enable-disable filename options
            text_disable(txtFilename, ((fViewing == true) || (!chkDestination2.checked)));
            button_disable(cmdFilename, ((fViewing == true) || (!chkDestination2.checked)));
        }
            /*else if (optOutputFormat5.checked == true)		//Excel Chart
                {
                }
            else if (optOutputFormat6.checked == true)		//Excel Pivot Table
                {
                }*/
        else {
            optOutputFormat0.checked = true;
            chkDestination0.checked = true;
            refreshTab5Controls();
        }

        if (txtEmailAttachAs.disabled) {
            txtEmailAttachAs.value = '';
        }
        else {
            if (txtEmailAttachAs.value == '') {
                if (txtFilename.value != '') {
                    sAttachmentName = new String(txtFilename.value);
                    txtEmailAttachAs.value = sAttachmentName.substr(sAttachmentName.lastIndexOf("\\") + 1);
                }
            }
        }

        if (cmdFilename.disabled == true) {
            txtFilename.value = "";
        }
    }

    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
        (fViewing == true)));


}

function formatClick(index) {
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
    refreshTab5Controls();
}

function saveFile() {
    dialog.CancelError = true;
    dialog.DialogTitle = "Output Document";
    dialog.Flags = 2621444;

    /*if (frmDefinition.optOutputFormat1.checked == true) {
        //CSV
        dialog.Filter = "Comma Separated Values (*.csv)|*.csv";
    }

    else */if (frmDefinition.optOutputFormat2.checked == true) {
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
        sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
        dialog.InitDir = sPath;
    }
    else {
        dialog.FileName = frmDefinition.txtFilename.value;
    }


    try {
        dialog.ShowSave();

        if (dialog.FileName.length > 256) {
            OpenHR.messageBox("Path and file name must not exceed 256 characters in length");
            return;
        }

        frmDefinition.txtFilename.value = dialog.FileName;

    }
    catch (e) {
    }

}


function changeBaseTableRecordOptions() {
    frmDefinition.txtBasePicklist.value = '';
    frmDefinition.txtBasePicklistID.value = 0;

    frmDefinition.txtBaseFilter.value = '';
    frmDefinition.txtBaseFilterID.value = 0;

    frmSelectionAccess.baseHidden.value = "N";

    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}

function clearBaseTableRecordOptions() {
    frmDefinition.optRecordSelection1.checked = true;

    button_disable(frmDefinition.cmdBasePicklist, true);
    frmDefinition.txtBasePicklist.value = '';
    frmDefinition.txtBasePicklistID.value = 0;

    button_disable(frmDefinition.cmdBaseFilter, true);
    frmDefinition.txtBaseFilter.value = '';
    frmDefinition.txtBaseFilterID.value = 0;

    frmSelectionAccess.baseHidden.value = "N";
}

function setRecordsNumeric(objTextBox) {
    var sConvertedValue;
    var sDecimalSeparator;
    var sThousandSeparator;
    var sPoint;

    sDecimalSeparator = "\\";
    sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
		
    var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

    sThousandSeparator = "\\";
    sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
    var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

    sPoint = "\\.";
    var rePoint = new RegExp(sPoint, "gi");

    if (objTextBox.value == '') {
        objTextBox.value = 0;
    }

    // Convert the value from locale to UK settings for use with the isNaN funtion.
    sConvertedValue = new String(objTextBox.value);

    // Remove any thousand separators.
    sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
    objTextBox.value = sConvertedValue;

    // Convert any decimal separators to '.'.
    if (OpenHR.LocaleDecimalSeparator != ".") {
        // Remove decimal points.
        sConvertedValue = sConvertedValue.replace(rePoint, "A");
        // replace the locale decimal marker with the decimal point.
        sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
    }

    if (isNaN(sConvertedValue) == true) {
        OpenHR.messageBox("Invalid numeric value.", 48, "Calendar Reports");
        objTextBox.value = 0;
    }
    else {
        if (sConvertedValue.indexOf(".") >= 0) {
            OpenHR.messageBox("Invalid integer value.", 48, "Calendar Reports");
            objTextBox.value = 0;
        }
        else {
            if (objTextBox.value > 99) {
                OpenHR.messageBox("The value cannot be greater than 99.", 48, "Calendar Reports");
                objTextBox.value = 99;
            }
        }
    }
}

function replace(sExpression, sFind, sReplace) {
    //gi (global search, ignore case)
    var re = new RegExp(sFind, "gi");
    sExpression = sExpression.replace(re, sReplace);
    return (sExpression);
}

function spinRecords(pfUp, objTextBox) {
    var iRecords = objTextBox.value;
    if (pfUp == true) {
        iRecords = ++iRecords;
    }
    else {
        iRecords = iRecords - 1;
    }
    objTextBox.value = iRecords;
}

function validateOffsets() {
    with (frmDefinition) {
        if ((optOffsetStart.checked == true) && (optOffsetEnd.checked == true)) {
            if (cboPeriodEnd.selectedIndex != cboPeriodStart.selectedIndex) {
                OpenHR.messageBox("The End Date Offset period must be the same as the Start Date Offset period", 48, "Calendar Reports");
                cboPeriodEnd.selectedIndex = cboPeriodStart.selectedIndex;
            }

            if (Number(txtFreqStart.value) > Number(txtFreqEnd.value)) {
                txtFreqEnd.value = txtFreqStart.value;
            }
        }
    }
}

function selectCalc(psCalcType, bRecordIndepend) {
    var iTableID;
    var iCurrentID;
    var sURL;
	
		var frmDefinition = document.getElementById("frmDefinition");
    var frmUseful = document.getElementById("frmUseful");
    var frmCalcSelection = document.getElementById("frmCalcSelection");
	


    if (psCalcType == 'baseDesc') {
        iTableID = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;
        iCurrentID = frmDefinition.txtDescExprID.value;
    }
    else if (psCalcType == 'startDate') {
        iTableID = 0;
        iCurrentID = frmDefinition.txtCustomStartID.value;
    }
    else if (psCalcType == 'endDate') {
        iTableID = 0;
        iCurrentID = frmDefinition.txtCustomEndID.value;
    }

    frmCalcSelection.calcSelRecInd.value = bRecordIndepend;
    frmCalcSelection.calcSelType.value = psCalcType;
    frmCalcSelection.calcSelTableID.value = iTableID;
    frmCalcSelection.calcSelCurrentID.value = iCurrentID;

    var strDefOwner = new String(frmDefinition.txtOwner.value);
    var strCurrentUser = new String(frmUseful.txtUserName.value);

    strDefOwner = strDefOwner.toLowerCase();
    strCurrentUser = strCurrentUser.toLowerCase();

    if (strDefOwner == strCurrentUser) {
        frmCalcSelection.recSelDefOwner.value = '1';
    }
    else {
        frmCalcSelection.recSelDefOwner.value = '0';
    }

    sURL = "util_calcSelection" +
        "?calcSelRecInd=" + frmCalcSelection.calcSelRecInd.value +
        "&calcSelType=" + escape(frmCalcSelection.calcSelType.value) +
        "&calcSelTableID=" + escape(frmCalcSelection.calcSelTableID.value) +
        "&calcSelCurrentID=" + escape(frmCalcSelection.calcSelCurrentID.value) +
        "&recSelDefOwner=" + escape(frmCalcSelection.recSelDefOwner.value) +
        "&destination=util_calcSelection";
    openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");

    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}

function selectEmailGroup() {
    var sURL;

    frmEmailSelection.EmailSelCurrentID.value = frmDefinition.txtEmailGroupID.value;

    sURL = "util_emailSelection" +
        "?EmailSelCurrentID=" + frmEmailSelection.EmailSelCurrentID.value;
    openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");
}

function eventFilterString() {
    var i;
    var pvarbookmark;
    var sEventFilterString = '';

    with (frmDefinition.grdEvents) {
        if (Rows > 0) {
            MoveFirst();
            for (i = 0; i < Rows; i++) {
                pvarbookmark = GetBookmark(i);
                sEventFilterString = sEventFilterString + Columns('FilterID').CellValue(pvarbookmark);
                if (i != Rows - 1) {
                    sEventFilterString = sEventFilterString + "	";
                }
            }
        }
    }
    if (sEventFilterString.length < 1) {
        sEventFilterString = '';
    }
    return sEventFilterString;
}

function submitDefinition() {
    var i;
    var iIndex;
    var sColumnID;
    var sType;
    var iDummy;
    var sURL;

    if (validateTab1() == false) { OpenHR.refreshMenu(); return; }
    if (validateTab2() == false) { OpenHR.refreshMenu(); return; }
    if (validateTab3() == false) { OpenHR.refreshMenu(); return; }
    if (validateTab4() == false) { OpenHR.refreshMenu(); return; }
    if (validateTab5() == false) { OpenHR.refreshMenu(); return; }
    if (populateSendForm() == false) { OpenHR.refreshMenu(); return; }

    // Now create the validate popup to check that any filters/calcs
    // etc havent been deleted, or made hidden etc.		
	
    // first populate the validate fields
    frmValidate.validateBaseFilter.value = frmDefinition.txtBaseFilterID.value;
    frmValidate.validateBasePicklist.value = frmDefinition.txtBasePicklistID.value;
    frmValidate.validateEmailGroup.value = frmDefinition.txtEmailGroupID.value;
    frmValidate.validateEventFilter.value = eventFilterString();
    frmValidate.validateName.value = frmDefinition.txtName.value;
    frmValidate.validateDescExpr.value = frmDefinition.txtDescExprID.value;
	frmValidate.validateCustomStart.value = frmDefinition.txtCustomStartID.value;
	frmValidate.validateCustomEnd.value = frmDefinition.txtCustomEndID.value;
	if (frmUseful.txtAction.value.toUpperCase() == "EDIT") {
        frmValidate.validateTimestamp.value = frmOriginalDefinition.txtDefn_Timestamp.value;
        frmValidate.validateUtilID.value = frmUseful.txtUtilID.value;
    }
    else {
        frmValidate.validateTimestamp.value = 0;
        frmValidate.validateUtilID.value = 0;
    }

		//ND May need to revisit or reinstate this if performance issues ensue
    //try {

    //	var frmRefresh = OpenHR.getForm("workframe", "frmRefresh");
    //	var testDataCollection = frmRefresh.elements;
    //    iDummy = testDataCollection.txtDummy.value;
    //    frmRefresh.submit();
    //}
    //catch (e) {
    //}

    var sHiddenGroups = HiddenGroups(frmDefinition.grdAccess);
    frmValidate.validateHiddenGroups.value = sHiddenGroups;

    sURL = "util_validate_calendarreport" +
        "?validateBaseFilter=" + frmValidate.validateBaseFilter.value +
        "&validateBasePicklist=" + escape(frmValidate.validateBasePicklist.value) +
        "&validateEmailGroup=" + escape(frmValidate.validateEmailGroup.value) +
        "&validateEventFilter=" + escape(frmValidate.validateEventFilter.value) +
        "&validateDescExpr=" + escape(frmValidate.validateDescExpr.value) +
        "&validateCustomStart=" + escape(frmValidate.validateCustomStart.value) +
        "&validateCustomEnd=" + escape(frmValidate.validateCustomEnd.value) +
        "&validateHiddenGroups=" + escape(frmValidate.validateHiddenGroups.value) +
        "&validateName=" + escape(frmValidate.validateName.value) +
        "&validateTimestamp=" + escape(frmValidate.validateTimestamp.value) +
        "&validateUtilID=" + escape(frmValidate.validateUtilID.value) +
        "&destination=util_validate_calendarreport";
	//openDialog(sURL, (screen.width) / 2, (screen.height) / 3, "no", "no");
    openDialog(sURL, (screen.width) / 2, (screen.height) / 3, "no", "no");
}

function cancelClick() {
	self.close();
}

function okClick() {
	menu_disableMenu();
	frmSend.txtSend_reaction.value = "CALENDARREPORTS";
	submitDefinition();
}

function saveChanges(psAction, pfPrompt, pfTBOverride) {
    if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
        (!definitionChanged())) {
        return 7; //No to saving the changes, as none have been made.
    }

    var answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3, "Calendar Reports");
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

function definitionChanged() {
    if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
        return false;
    }

    if (frmUseful.txtAction.value.toUpperCase() == "COPY") {
        return true;
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

            // Compare the tab 3 controls with the original values.
            // Changes to the following data are checked using the frmUseful.txtChanged value :
            //   Selected columns and their order
            //   Column headings, sizes and decimals
            //	 Column aggregates

            // Compare the tab 4 controls with the original values.
            // Changes to the following data are checked using the frmUseful.txtChanged value :
            //   Sort columns, their order, asc/desc, boc, poc, voc and srv values

            // Compare the tab 5 controls with the original values.
            if (frmDefinition.chkPreview.checked.toString().toUpperCase() != frmOriginalDefinition.txtDefn_OutputPreview.value.toUpperCase()) {
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

function getTableName(piTableID) {
    var i;
    var sTableName = new String("");

    sReqdControlName = new String("txtTableName_");
    sReqdControlName = sReqdControlName.concat(piTableID);

    var dataCollection = frmTables.elements;
    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {
            sControlName = dataCollection.item(i).name;

            if (sControlName == sReqdControlName) {
                sTableName = dataCollection.item(i).value;
                return sTableName;
            }
        }
    }

    return sTableName;
}



function getEventKey() {
    var i = 1;

    while (checkUniqueEventKey("EV_" + i)) {
        i = i + 1;
    }

    var sEventKey = new String("EV_" + i);

    return sEventKey;
}

function checkUniqueEventKey(psNewKey) {
    var bm;

    with (frmDefinition.grdEvents) {
        Redraw = false;
        MoveFirst();
        for (var i = 0; i < Rows; i++) {
            bm = AddItemBookmark(i);
            if (psNewKey == trim(Columns("EventKey").CellValue(bm))) {
                Redraw = true;
                return true;
            }
        }
        Redraw = true;
    }
    return false;
}

function eventEdit() {
    var sURL;
	debugger;
    with (frmEventDetails) {
        eventAction.value = "EDIT";
        eventName.value = frmDefinition.grdEvents.Columns("Name").value;
        eventID.value = frmDefinition.grdEvents.Columns("EventKey").value;
        eventTableID.value = frmDefinition.grdEvents.Columns("TableID").value;
        eventTable.value = frmDefinition.grdEvents.Columns("Table").value;
        eventFilterID.value = frmDefinition.grdEvents.Columns("FilterID").value;
        eventFilter.value = frmDefinition.grdEvents.Columns("Filter").value;
        eventFilterHidden.value = frmDefinition.grdEvents.Columns("FilterHidden").value;

        eventStartDateID.value = frmDefinition.grdEvents.Columns("StartDateID").value;
        eventStartDate.value = frmDefinition.grdEvents.Columns("Start Date").value;
        eventStartSessionID.value = frmDefinition.grdEvents.Columns("StartSessionID").value;
        eventStartSession.value = frmDefinition.grdEvents.Columns("Start Session").value;
        eventEndDateID.value = frmDefinition.grdEvents.Columns("EndDateID").value;
        eventEndDate.value = frmDefinition.grdEvents.Columns("End Date").value;
        eventEndSessionID.value = frmDefinition.grdEvents.Columns("EndSessionID").value;
        eventEndSession.value = frmDefinition.grdEvents.Columns("End Session").value;
        eventDurationID.value = frmDefinition.grdEvents.Columns("DurationID").value;
        eventDuration.value = frmDefinition.grdEvents.Columns("Duration").value;

        eventLookupType.value = frmDefinition.grdEvents.Columns("LegendType").value;
        eventKeyCharacter.value = frmDefinition.grdEvents.Columns("Legend").value.substr(0, 2);
        eventLookupTableID.value = frmDefinition.grdEvents.Columns("LegendTableID").value;
        eventLookupColumnID.value = frmDefinition.grdEvents.Columns("LegendColumnID").value;
        eventLookupCodeID.value = frmDefinition.grdEvents.Columns("LegendCodeID").value;
        eventTypeColumnID.value = frmDefinition.grdEvents.Columns("LegendEventTypeID").value;

        eventDesc1ID.value = frmDefinition.grdEvents.Columns("Desc1ID").value;
        eventDesc1.value = frmDefinition.grdEvents.Columns("Description 1").value;
        eventDesc2ID.value = frmDefinition.grdEvents.Columns("Desc2ID").value;
        eventDesc2.value = frmDefinition.grdEvents.Columns("Description 2").value;

        sURL = "util_def_calendarreportdates_main" +
            "?eventAction=" + escape(frmEventDetails.eventAction.value) +
            "&eventName=" + escape(frmEventDetails.eventName.value) +
            "&eventID=" + escape(frmEventDetails.eventID.value) +
            "&eventTableID=" + escape(frmEventDetails.eventTableID.value) +
            "&eventTable=" + escape(frmEventDetails.eventTable.value) +
            "&eventFilterID=" + escape(frmEventDetails.eventFilterID.value) +
            "&eventFilter=" + escape(frmEventDetails.eventFilter.value) +
            "&eventFilterHidden=" + escape(frmEventDetails.eventFilterHidden.value) +
            "&eventStartDateID=" + escape(frmEventDetails.eventStartDateID.value) +
            "&eventStartDate=" + escape(frmEventDetails.eventStartDate.value) +
            "&eventStartSessionID=" + escape(frmEventDetails.eventStartSessionID.value) +
            "&eventStartSession=" + escape(frmEventDetails.eventStartSession.value) +
            "&eventEndDateID=" + escape(frmEventDetails.eventEndDateID.value) +
            "&eventEndDate=" + escape(frmEventDetails.eventEndDate.value) +
            "&eventEndSessionID=" + escape(frmEventDetails.eventEndSessionID.value) +
            "&eventEndSession=" + escape(frmEventDetails.eventEndSession.value) +
            "&eventDurationID=" + escape(frmEventDetails.eventDurationID.value) +
            "&eventDuration=" + escape(frmEventDetails.eventDuration.value) +
            "&eventLookupType=" + escape(frmEventDetails.eventLookupType.value) +
            "&eventKeyCharacter=" + escape(frmEventDetails.eventKeyCharacter.value) +
            "&eventLookupTableID=" + escape(frmEventDetails.eventLookupTableID.value) +
            "&eventLookupColumnID=" + escape(frmEventDetails.eventLookupColumnID.value) +
            "&eventLookupCodeID=" + escape(frmEventDetails.eventLookupCodeID.value) +
            "&eventTypeColumnID=" + escape(frmEventDetails.eventTypeColumnID.value) +
            "&eventDesc1ID=" + escape(frmEventDetails.eventDesc1ID.value) +
            "&eventDesc1=" + escape(frmEventDetails.eventDesc1.value) +
            "&eventDesc2ID=" + escape(frmEventDetails.eventDesc2ID.value) +
            "&eventDesc2=" + escape(frmEventDetails.eventDesc2.value) +
            "&relationNames=" + escape(frmEventDetails.relationNames.value);
        openDialog(sURL, 650, 500, "yes", "yes");
        frmUseful.txtChanged.value = 1;
    }
    refreshTab2Controls();
}


function eventRemove() {
    var lRow;
    var lngSelectedEvent;

    with (frmDefinition.grdEvents) {
        if (Rows < 1) return;

        lRow = AddItemRowIndex(Bookmark);
        lngSelectedEvent = Columns('EventKey').CellValue(lRow);

        var bContinueRemoval;

        bContinueRemoval = true;

        if (!bContinueRemoval) return;

        if (Rows == 1) {
            RemoveAll();
        }
        else {
            RemoveItem(lRow);
            if (Rows != 0) {
                if (lRow < Rows) {
                    Bookmark = lRow;
                }
                else {
                    Bookmark = (Rows - 1);
                }
                SelBookmarks.Add(Bookmark);
            }
        }
    }
    frmUseful.txtChanged.value = 1;

    refreshTab2Controls();
}

function eventRemoveAll() {
    var i;
    var pvarbookmark;
    var bContinueRemoval;
    var lngSelectedEvent;
    var lngRowCount;

    bContinueRemoval = true;

    if (!bContinueRemoval) return;

    with (frmDefinition.grdEvents) {
        Redraw = false;
        lngRowCount = Rows;
        for (i = 0; i < lngRowCount; i++) {
            MoveFirst();
            pvarbookmark = AddItemBookmark(i);
            lngSelectedEvent = Columns('EventKey').CellValue(pvarbookmark);
        }
        Redraw = true;
        RemoveAll();
        SelBookmarks.RemoveAll();
    }
    frmUseful.txtChanged.value = 1;

    refreshTab2Controls();
}


function sortAdd() {
    var i;
    var iCalcsCount = 0;
    var iColumnsCount = 0;
    var sURL;
    var bm;
		// Loop through the columns added and populate the 
    // sort order text boxes to pass to util_sortorderselection
    frmSortOrder.txtSortInclude.value = '';
    frmSortOrder.txtSortExclude.value = '';
    frmSortOrder.txtSortEditing.value = 'false';
    frmSortOrder.txtSortColumnID.value = frmDefinition.ssOleDBGridSortOrder.Columns(0).text;
    frmSortOrder.txtSortColumnName.value = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
    frmSortOrder.txtSortOrder.value = frmDefinition.ssOleDBGridSortOrder.Columns(2).text;

    if (frmDefinition.cboDescription1.options.length > 0) {
        for (i = 0; i < frmDefinition.cboDescription1.options.length; i++) {
            iColumnsCount++;
            if (frmSortOrder.txtSortInclude.value != '') {
                frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
            }
            frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.cboDescription1.options[i].value;
        }
    }
    else {
        OpenHR.messageBox("No columns on the base table.", 48, "Calendar Reports");
    }

    if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
        frmDefinition.ssOleDBGridSortOrder.Redraw = false;
        frmDefinition.ssOleDBGridSortOrder.movefirst();

        for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) {
            bm = frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(i);
            if (frmSortOrder.txtSortExclude.value != '') {
                frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + ',';
            }

            frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + frmDefinition.ssOleDBGridSortOrder.Columns(0).CellValue(bm);

            //frmDefinition.ssOleDBGridSortOrder.movenext();
        }

        frmDefinition.ssOleDBGridSortOrder.Redraw = true;
    }

    if (frmSortOrder.txtSortInclude.value == frmSortOrder.txtSortExclude.value) {
        OpenHR.messageBox("You have selected all base columns in the sort order.", 48, "Calendar Reports");
    }
    else if ((frmDefinition.ssOleDBGridSortOrder.Rows - iColumnsCount) == 0) {
        OpenHR.messageBox("You have selected all base columns in the sort order.", 48, "Calendar Reports");
    }
    else {
    	if (frmSortOrder.txtSortInclude.value != '') {
    		sURL = "util_sortorderselection" +
                "?txtSortInclude=" + escape(frmSortOrder.txtSortInclude.value) +
                "&txtSortExclude=" + escape(frmSortOrder.txtSortExclude.value) +
                "&txtSortEditing=" + escape(frmSortOrder.txtSortEditing.value) +
                "&txtSortColumnID=" + escape(frmSortOrder.txtSortColumnID.value) +
                "&txtSortColumnName=" + escape(frmSortOrder.txtSortColumnName.value) +
                "&txtSortOrder=" + escape(frmSortOrder.txtSortOrder.value);
            openDialog(sURL, 600, 275, "yes", "yes");

            frmUseful.txtChanged.value = 1;
        }
    }
}

function sortEdit() {
    var i;
    var iIndex;
    var sDefn;
    var sColumnID;
    var sURL;

    frmSortOrder.txtSortInclude.value = '';
    frmSortOrder.txtSortExclude.value = '';
    frmSortOrder.txtSortEditing.value = 'true';
    frmSortOrder.txtSortColumnID.value = frmDefinition.ssOleDBGridSortOrder.Columns(0).text;
    frmSortOrder.txtSortColumnName.value = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
    frmSortOrder.txtSortOrder.value = frmDefinition.ssOleDBGridSortOrder.Columns(2).text;

    for (i = 0; i < frmDefinition.cboDescription1.options.length; i++) {
        if (frmSortOrder.txtSortInclude.value != '') {
            frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
        }
        frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.cboDescription1.options[i].value;
    }

    frmDefinition.ssOleDBGridSortOrder.Redraw = false;
    var rowNum = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark);
    frmDefinition.ssOleDBGridSortOrder.MoveFirst();

    for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) {
        if (frmDefinition.ssOleDBGridSortOrder.columns(0).text != frmSortOrder.txtSortColumnID.value) {
            if (frmSortOrder.txtSortExclude.value != '') {
                frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + ',';
            }
            frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + frmDefinition.ssOleDBGridSortOrder.columns(0).text;
        }

        frmDefinition.ssOleDBGridSortOrder.movenext();
    }

    frmDefinition.ssOleDBGridSortOrder.Redraw = true;
    frmDefinition.ssOleDBGridSortOrder.bookmark = frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(rowNum);
    sURL = "util_sortorderselection" +
        "?txtSortInclude=" + escape(frmSortOrder.txtSortInclude.value) +
        "&txtSortExclude=" + escape(frmSortOrder.txtSortExclude.value) +
        "&txtSortEditing=" + escape(frmSortOrder.txtSortEditing.value) +
        "&txtSortColumnID=" + escape(frmSortOrder.txtSortColumnID.value) +
        "&txtSortColumnName=" + escape(frmSortOrder.txtSortColumnName.value) +
        "&txtSortOrder=" + escape(frmSortOrder.txtSortOrder.value);
    openDialog(sURL, 500, 275, "yes", "yes");

    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
}

function sortRemove() {
    if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count() == 0) {
        OpenHR.messageBox("You must select a column to remove.", 48, "Calendar Reports");
        return;
    }

    frmDefinition.ssOleDBGridSortOrder.RemoveItem(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.bookmark));

    if (frmDefinition.ssOleDBGridSortOrder.Rows != 0) {
        frmDefinition.ssOleDBGridSortOrder.MoveLast();
        frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
    }

    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
}

function sortRemoveAll() {
    frmDefinition.ssOleDBGridSortOrder.RemoveAll();
    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
}

function sortMove(pfUp) {
    var sAddline = '';

    if (pfUp == true) {
        iNewIndex = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) - 1;
        iOldIndex = iNewIndex + 2;
        iSelectIndex = iNewIndex;
    }
    else {
        iNewIndex = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) + 2;
        iOldIndex = iNewIndex - 2;
        iSelectIndex = iNewIndex - 1;
    }

    sAddline = frmDefinition.ssOleDBGridSortOrder.columns(0).text +
        '	' + frmDefinition.ssOleDBGridSortOrder.columns(1).text +
        '	' + frmDefinition.ssOleDBGridSortOrder.columns(2).text;
	frmDefinition.ssOleDBGridSortOrder.additem(sAddline, iNewIndex);
    frmDefinition.ssOleDBGridSortOrder.RemoveItem(iOldIndex);

    frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();
    frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(iSelectIndex));
    frmDefinition.ssOleDBGridSortOrder.Bookmark = frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(iSelectIndex);

    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
}

function validateTab1() {
    // check name has been entered
    if (frmDefinition.txtName.value == '') {
        OpenHR.messageBox("You must enter a name for this definition.", 48, "Calendar Reports");
        displayPage(1);
        return (false);
    }

    // check base picklist
    if ((frmDefinition.optRecordSelection2.checked == true) &&
        (frmDefinition.txtBasePicklistID.value == 0)) {
        OpenHR.messageBox("You must select a picklist for the base table.", 48, "Calendar Reports");
        displayPage(1);
        return (false);
    }

    // check base filter
    if ((frmDefinition.optRecordSelection3.checked == true) &&
        (frmDefinition.txtBaseFilterID.value == 0)) {
        OpenHR.messageBox("You must select a filter for the base table.", 48, "Calendar Reports");
        displayPage(1);
        return (false);
    }

    // Check that a valid description column or a valid calculation has been selected
    if ((frmDefinition.cboDescription1.options[frmDefinition.cboDescription1.selectedIndex].value < 1) && (frmDefinition.txtDescExprID.value < 1) && (frmDefinition.cboDescription2.options[frmDefinition.cboDescription2.selectedIndex].value < 1)) {
        OpenHR.messageBox("You must select at least one base description column or calculation for the report.", 48, "Calendar Reports");
        displayPage(1);
        return (false);
    }

    return (true);
}

function validateTab2() {
    var i;
    var sErrMsg;
    var iCount;
    var sDefn;
    var sControlName;
    var frmRefresh = OpenHR.getForm("workframe", "frmRefresh");
    var iDummy;

    sErrMsg = "";

    //check at least one column defined as sort order
    if (frmUseful.txtEventsLoaded.value == 1) {
        if (frmDefinition.grdEvents.Rows <= 0) {
            sErrMsg = "You must select at least one event to report on.";
        }
    }
    else {
        iCount = 0;
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (i = 0; i < dataCollection.length; i++) {
					
                sControlName = dataCollection.item(i).name;
                sControlName = sControlName.substr(0, 19);
                if (sControlName == "txtReportDefnEvent_") {
                    sDefn = new String(dataCollection.item(i).value);

                    iCount = iCount + 1;
                }
            }
        }

        if (iCount == 0) {
            sErrMsg = "You must select at least one event to report on.";
        }
    }

    if (sErrMsg.length > 0) {
        OpenHR.messageBox(sErrMsg, 48, "Calendar Reports");
        displayPage(2);
        return (false);
    }

	//ND May need to revisit or reinstate this if performance issues ensue
    //try {
	  //  var testDataCollection = frmRefresh.elements;
	  //  iDummy = testDataCollection.txtDummy.value;
	  //  iDummy = frmRefresh.elements.txtDummy.value;
	    
    //    frmRefresh.submit();
    //}
    //catch (e) {
    //}

    return (true);
}

function validateTab3() {
    with (frmDefinition) {
        if (optFixedStart.checked && (trim(txtFixedStart.value) == '')) {
            OpenHR.messageBox("You must select a Fixed Start Date for the report.", 48, "Calendar Reports");
            displayPage(3);
            return (false);
        }
        if (optFixedStart.checked && (!validateDate(txtFixedStart))) {
            OpenHR.messageBox("You must enter a valid Fixed Start Date for the report.", 48, "Calendar Reports");
            displayPage(3);
            return (false);
        }

        if (optFixedEnd.checked && (trim(txtFixedEnd.value) == '')) {
            OpenHR.messageBox("You must select a Fixed End Date for the report.", 48, "Calendar Reports");
            displayPage(3);
            return (false);
        }
        if (optFixedEnd.checked && (!validateDate(txtFixedEnd))) {
            OpenHR.messageBox("You must enter a valid Fixed End Date for the report.", 48, "Calendar Reports");
            displayPage(3);
            return (false);
        }

        if (optFixedStart.checked && optFixedEnd.checked) {
            var dtStartDate = convertLocaleDateToDateObject(txtFixedStart.value);
            var dtEndDate = convertLocaleDateToDateObject(txtFixedEnd.value);

            if (dtEndDate.getTime() < dtStartDate.getTime()) {
                OpenHR.messageBox("You must select a Fixed End Date later than or equal to the Fixed Start Date.", 48, "Calendar Reports");
                displayPage(3);
                return (false);
            }
        }

        if (optOffsetStart.checked && optFixedEnd.checked) {
            if (Number(txtFreqEnd.value) < 0) {
                OpenHR.messageBox("You must select an End Date Offset greater than or equal to zero.", 48, "Calendar Reports");
                displayPage(3);
                return (false);
            }
        }

        if (optCurrentStart.checked && optOffsetEnd.checked) {
            if (Number(txtFreqEnd.value) < 0) {
                OpenHR.messageBox("You must select an End Date Offset greater than or equal to zero.", 48, "Calendar Reports");
                displayPage(3);
                return (false);
            }
        }

        if (optOffsetStart.checked && (optFixedEnd.checked || optCurrentEnd.checked)) {
            if (Number(txtFreqStart.value) > 0) {
                OpenHR.messageBox("You must select a Start Date Offset less than or equal to zero.", 48, "Calendar Reports");
                displayPage(3);
                return (false);
            }
        }

        if (optOffsetStart.checked && optOffsetEnd.checked) {
            if (cboPeriodStart.selectedIndex != cboPeriodEnd.selectedIndex) {
                OpenHR.messageBox("You must select the same End Date Offset period as Start Date Offset period.", 48, "Calendar Reports");
                displayPage(3);
                return (false);
            }

            if (Number(txtFreqEnd.value) < Number(txtFreqStart.value)) {
                OpenHR.messageBox("You must select an End Date Offset greater than or equal to the Start Date Offset.", 48, "Calendar Reports");
                displayPage(3);
                return (false);
            }
        }

        if (optCustomStart.checked && (txtCustomStartID.value < 1)) {
            OpenHR.messageBox("You must select a calculation for the Report Start Date.", 48, "Calendar Reports");
            displayPage(3);
            return (false);
        }

        if (optCustomEnd.checked && (txtCustomEndID.value < 1)) {
            OpenHR.messageBox("You must select a calculation for the Report End Date.", 48, "Calendar Reports");
            displayPage(3);
            return (false);
        }
    }

    return (true);
}



function validateTab4() {
	var i;
	var sErrMsg;
	var iCount;
	var sDefn;
	var sControlName;

	sErrMsg = "";

	//check at least one column defined as sort order
	var testDataCollection;
	if (frmUseful.txtSortLoaded.value == 1) {
		with (frmDefinition.ssOleDBGridSortOrder) {
			if (frmDefinition.ssOleDBGridSortOrder.Rows <= 0) {
				sErrMsg = "You must select at least one column to order the report by.";
			}
			else {
				if ((frmDefinition.chkGroupByDesc.checked) && (frmDefinition.txtDescExprID.value < 1)) {
					var lngDesc1ID = frmDefinition.cboDescription1.options[frmDefinition.cboDescription1.selectedIndex].value;
					var lngDesc2ID = frmDefinition.cboDescription2.options[frmDefinition.cboDescription2.selectedIndex].value;

					var strDesc = new String(lngDesc1ID);
					if (lngDesc2ID > 0) {
						strDesc = strDesc + '	' + lngDesc2ID;
					}
					var strTemp = new String('');
					frmDefinition.ssOleDBGridSortOrder.Redraw = false;
					frmDefinition.ssOleDBGridSortOrder.MoveFirst();
					for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.Rows; i++) {
						if (frmDefinition.ssOleDBGridSortOrder.Columns('ColumnID').Text > 0) {
							if (i > 0) {
								strTemp = strTemp + '	';
							}
							strTemp = strTemp + frmDefinition.ssOleDBGridSortOrder.Columns('ColumnID').Value;
						}
						if (i >= 1) {
							break;
						}
						frmDefinition.ssOleDBGridSortOrder.MoveNext();
					}
					frmDefinition.ssOleDBGridSortOrder.MoveFirst();

					if (strTemp != strDesc) {
						sErrMsg = "The sort order does not reflect the selected Group By Description columns. Do you wish to continue?";
						if (OpenHR.messageBox(sErrMsg, 36, "Calendar Report Definition") == 7) {
							frmDefinition.ssOleDBGridSortOrder.Redraw = true;
							displayPage(4);
							return (false);
						}
						else {
							frmDefinition.ssOleDBGridSortOrder.Redraw = true;
							sErrMsg = '';
						}
					}
				}
			}
		}
	}
	else {
		iCount = 0;
		var dataCollection = frmOriginalDefinition.elements;
		if (dataCollection != null) {
			for (i = 0; i < dataCollection.length; i++) {
				sControlName = dataCollection.item(i).name;
				sControlName = sControlName.substr(0, 19);
				if (sControlName == "txtReportDefnOrder_") {
					sDefn = new String(dataCollection.item(i).value);

					iCount = iCount + 1;
				}
			}
		}

		if (iCount == 0) {
			sErrMsg = "You must select at least one column to order the report by.";
		}
	}

	if (sErrMsg.length > 0) {
		OpenHR.messageBox(sErrMsg, 48, "Calendar Reports");
		displayPage(4);
		return (false);
	}
	//ND May need to revisit or reinstate this if performance issues ensue
	//try {
	//	testDataCollection = frmRefresh.elements;
	//	iDummy = testDataCollection.txtDummy.value;
	//	frmRefresh.submit();
	//}
	//catch (e) {
	//}

	return (true);
}

function validateTab5() {
    var sErrMsg;

    sErrMsg = "";

    if (!frmDefinition.chkDestination0.checked
        && !frmDefinition.chkDestination1.checked
        && !frmDefinition.chkDestination2.checked
        && !frmDefinition.chkDestination3.checked) {
        sErrMsg = "You must select a destination";
    }

    if ((frmDefinition.txtFilename.value == "")
        && (frmDefinition.cmdFilename.disabled == false)) {
        sErrMsg = "You must enter a file name";
    }

    if ((frmDefinition.txtEmailGroup.value == "")
        && (frmDefinition.cmdEmailGroup.disabled == false)) {
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

    if (sErrMsg.length > 0) {
        OpenHR.messageBox(sErrMsg, 48, "Calendar Reports");
        displayPage(5);
        return (false);
    }

    return (true);
}

function populateSendForm() {
    var i;
    var iIndex;
    var sControlName;
    var iNum;
    var frmRefresh = OpenHR.getForm("workframe", "frmRefresh");
    var iDummy;
    var varBookmark;
    var iLoop;
    var sAccess;

    /******************** TAB 1 - DEFINITION *********************/

    // Copy all the header information to frmSend
    frmSend.txtSend_ID.value = frmUseful.txtUtilID.value;
    frmSend.txtSend_name.value = frmDefinition.txtName.value;
    frmSend.txtSend_description.value = frmDefinition.txtDescription.value;
    frmSend.txtSend_userName.value = frmDefinition.txtOwner.value;

    sAccess = "";
    frmDefinition.grdAccess.update();
    for (iLoop = 1; iLoop <= (frmDefinition.grdAccess.Rows - 1) ; iLoop++) {
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

    if (frmDefinition.chkPrintFilterHeader.checked == true) {
        frmSend.txtSend_printFilterHeader.value = '1';
    }
    else {
        frmSend.txtSend_printFilterHeader.value = '0';
    }

    frmSend.txtSend_desc1.value = frmDefinition.cboDescription1.options(frmDefinition.cboDescription1.options.selectedIndex).value;
    frmSend.txtSend_desc2.value = frmDefinition.cboDescription2.options(frmDefinition.cboDescription2.options.selectedIndex).value;

    if (frmDefinition.txtDescExprID.value > 0) {
        frmSend.txtSend_descExpr.value = frmDefinition.txtDescExprID.value;
    }
    else {
        frmSend.txtSend_descExpr.value = 0;
    }

    frmSend.txtSend_region.value = frmDefinition.cboRegion.options(frmDefinition.cboRegion.options.selectedIndex).value;

    if (frmDefinition.chkGroupByDesc.checked) {
        frmSend.txtSend_groupbydesc.value = 1;
    }
    else {
        frmSend.txtSend_groupbydesc.value = 0;
    }

    if (frmDefinition.cboDescriptionSeparator.selectedIndex < 0) {
        frmSend.txtSend_descseparator.value = ', ';
    }
    else {
        frmSend.txtSend_descseparator.value = frmDefinition.cboDescriptionSeparator.options[frmDefinition.cboDescriptionSeparator.selectedIndex].value;
    }

    /*************************************************************/
    /******************* TAB 2 - EVENT DETAILS *******************/
    // now go through the columns grid (and sort order grid)(and the repetition grid)
    var sEvents = '';
	//var testDataCollection;
	//try {
	//	testDataCollection = frmRefresh.elements;
	//	iDummy = testDataCollection.txtDummy.value;
  //      frmRefresh.submit();
  //  }
  //  catch (e) {
  //  }

    frmUseful.txtLockGridEvents.value = 1;
	var dataCollection;
	if (frmUseful.txtEventsLoaded.value == 1) {
        frmDefinition.grdEvents.Redraw = false;
        frmDefinition.grdEvents.movefirst();

        for (i = 0; i < frmDefinition.grdEvents.rows; i++) {
				
            sEvents = sEvents +
                trim(frmDefinition.grdEvents.columns("EventKey").text) +
                '||' + frmDefinition.grdEvents.columns("Name").text +
                '||' + frmDefinition.grdEvents.columns("TableID").text +
                '||' + frmDefinition.grdEvents.columns("FilterID").text +
                '||' + frmDefinition.grdEvents.columns("StartDateID").text +
                '||' + frmDefinition.grdEvents.columns("StartSessionID").text +
                '||' + frmDefinition.grdEvents.columns("EndDateID").text +
                '||' + frmDefinition.grdEvents.columns("EndSessionID").text +
                '||' + frmDefinition.grdEvents.columns("DurationID").text;
	        if (frmDefinition.grdEvents.columns("LegendType").text == '1') {
                sEvents = sEvents +
                    '||' + '1' +
                    '||' + '' +
                    '||' + frmDefinition.grdEvents.columns("LegendTableID").text +
                    '||' + frmDefinition.grdEvents.columns("LegendColumnID").text +
                    '||' + frmDefinition.grdEvents.columns("LegendCodeID").text +
                    '||' + frmDefinition.grdEvents.columns("LegendEventTypeID").text;
	        }
            else {
                sEvents = sEvents +
                    '||' + '0' +
                    '||' + replace(frmDefinition.grdEvents.columns("Legend").text, "'", "''") +
                    '||' + 0 +
                    '||' + 0 +
                    '||' + 0 +
                    '||' + 0;
	        }

            sEvents = sEvents +
                '||' + frmDefinition.grdEvents.columns("Desc1ID").text +
                '||' + frmDefinition.grdEvents.columns("Desc2ID").text +
                '||';

            sEvents = sEvents + '**';

            frmDefinition.grdEvents.movenext();
        }
        frmDefinition.grdEvents.Redraw = true;
    }
    else {
		dataCollection = frmOriginalDefinition.elements;
		if (dataCollection != null) {
            iNum = 0;
            for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                sControlName = dataCollection.item(iIndex).name;
                sControlName = sControlName.substr(0, 19);
                if (sControlName == "txtReportDefnEvent_") {
                    sEvents = sEvents +
                        trim(selectedEventParameter(dataCollection.item(iIndex).value, "EVENTKEY")) +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "NAME") +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "TABLEID") +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "FILTERID") +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "STARTDATEID") +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "STARTSESSIONID") +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "ENDDATEID") +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "ENDSESSIONID") +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "DURATIONID");
	                if (selectedEventParameter(dataCollection.item(iIndex).value, "LEGENDTYPE") == '1') {
                        sEvents = sEvents +
                            '||' + '1' +
                            '||' + '' +
                            '||' + selectedEventParameter(dataCollection.item(iIndex).value, "LEGENDLOOKUPTABLEID") +
                            '||' + selectedEventParameter(dataCollection.item(iIndex).value, "LEGENDLOOKUPCOLUMNID") +
                            '||' + selectedEventParameter(dataCollection.item(iIndex).value, "LEGENDLOOKUPCODEID") +
                            '||' + selectedEventParameter(dataCollection.item(iIndex).value, "LEGENDEVENTCOLUMNID");
	                }
                    else {
                        sEvents = sEvents +
                            '||' + '0' +
                            '||' + replace(selectedEventParameter(dataCollection.item(iIndex).value, "LEGEND"), "'", "''") +
                            '||' + '0' +
                            '||' + '0' +
                            '||' + '0' +
                            '||' + '0';
	                }

                    sEvents = sEvents +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "DESC1COLUMNID") +
                        '||' + selectedEventParameter(dataCollection.item(iIndex).value, "DESC2COLUMNID") +
                        '||';

                    sEvents = sEvents + '**';
                }
            }
        }
    }

    frmSend.txtSend_Events.value = sEvents.substr(0, 8000);
    frmSend.txtSend_Events2.value = sEvents.substr(8000, 8000);

    /*************************************************************/

    /******************* TAB 3 - REPORT DETAILS ******************/

    if (frmDefinition.optFixedStart.checked == true) {
        frmSend.txtSend_StartType.value = 0;
        frmSend.txtSend_FixedStart.value = convertLocaleDateToSQL(frmDefinition.txtFixedStart.value);
        frmSend.txtSend_StartFrequency.value = 0;
        frmSend.txtSend_StartPeriod.value = -1;
        frmSend.txtSend_CustomStart.value = 0;
    }
    else if (frmDefinition.optCurrentStart.checked == true) {
        frmSend.txtSend_StartType.value = 1;
        frmSend.txtSend_FixedStart.value = "";
	    frmSend.txtSend_StartFrequency.value = 0;
        frmSend.txtSend_StartPeriod.value = -1;
        frmSend.txtSend_CustomStart.value = 0;
    }
    else if (frmDefinition.optOffsetStart.checked == true) {
        frmSend.txtSend_StartType.value = 2;
        frmSend.txtSend_FixedStart.value = "";
	    frmSend.txtSend_StartFrequency.value = frmDefinition.txtFreqStart.value;
        frmSend.txtSend_StartPeriod.value = frmDefinition.cboPeriodStart.options[frmDefinition.cboPeriodStart.selectedIndex].value;
        frmSend.txtSend_CustomStart.value = 0;
    }
    else if (frmDefinition.optCustomStart.checked == true) {
        frmSend.txtSend_StartType.value = 3;
        frmSend.txtSend_FixedStart.value = "";
	    frmSend.txtSend_StartFrequency.value = 0;
        frmSend.txtSend_StartPeriod.value = -1;
        frmSend.txtSend_CustomStart.value = frmDefinition.txtCustomStartID.value;
    }

    if (frmDefinition.optFixedEnd.checked == true) {
        frmSend.txtSend_EndType.value = 0;
        frmSend.txtSend_FixedEnd.value = convertLocaleDateToSQL(frmDefinition.txtFixedEnd.value);
        frmSend.txtSend_EndFrequency.value = 0;
        frmSend.txtSend_EndPeriod.value = -1;
        frmSend.txtSend_CustomEnd.value = 0;
    }
    else if (frmDefinition.optCurrentEnd.checked == true) {
        frmSend.txtSend_EndType.value = 1;
        frmSend.txtSend_FixedEnd.value = "";
        frmSend.txtSend_EndFrequency.value = 0;
        frmSend.txtSend_EndPeriod.value = -1;
        frmSend.txtSend_CustomEnd.value = 0;
    }
    else if (frmDefinition.optOffsetEnd.checked == true) {
        frmSend.txtSend_EndType.value = 2;
        frmSend.txtSend_FixedEnd.value = "";
        frmSend.txtSend_EndFrequency.value = frmDefinition.txtFreqEnd.value;
        frmSend.txtSend_EndPeriod.value = frmDefinition.cboPeriodEnd.options[frmDefinition.cboPeriodEnd.selectedIndex].value;
        frmSend.txtSend_CustomEnd.value = 0;
    }
    else if (frmDefinition.optCustomEnd.checked == true) {
        frmSend.txtSend_EndType.value = 3;
        frmSend.txtSend_FixedEnd.value = "";
        frmSend.txtSend_EndFrequency.value = 0;
        frmSend.txtSend_EndPeriod.value = -1;
        frmSend.txtSend_CustomEnd.value = frmDefinition.txtCustomEndID.value;
    }

    if (frmDefinition.chkIncludeBHols.checked == true) {
        frmSend.txtSend_IncludeBHols.value = '1';
    }
    else {
        frmSend.txtSend_IncludeBHols.value = '0';
    }

    if (frmDefinition.chkIncludeWorkingDaysOnly.checked == true) {
        frmSend.txtSend_IncludeWorkingDaysOnly.value = '1';
    }
    else {
        frmSend.txtSend_IncludeWorkingDaysOnly.value = '0';
    }

    if (frmDefinition.chkShadeBHols.checked == true) {
        frmSend.txtSend_ShadeBHols.value = '1';
    }
    else {
        frmSend.txtSend_ShadeBHols.value = '0';
    }

    if (frmDefinition.chkCaptions.checked == true) {
        frmSend.txtSend_Captions.value = '1';
    }
    else {
        frmSend.txtSend_Captions.value = '0';
    }

    if (frmDefinition.chkShadeWeekends.checked == true) {
        frmSend.txtSend_ShadeWeekends.value = '1';
    }
    else {
        frmSend.txtSend_ShadeWeekends.value = '0';
    }

    if (frmDefinition.chkStartOnCurrentMonth.checked == true) {
        frmSend.txtSend_StartOnCurrentMonth.value = '1';
    }
    else {
        frmSend.txtSend_StartOnCurrentMonth.value = '0';
    }

    /*************************************************************/

    /********************* TAB 4 - SORT ORDER ********************/

    /*now use the txtSend_OrderString to hold the string of selected order information*/
    if (frmUseful.txtSortLoaded.value == 1) {
        sOrders = '';
        i = 0;
			
        with (frmDefinition.ssOleDBGridSortOrder) {
            if (Rows > 0) {
                Redraw = false;
                MoveFirst();

                for (i = 0; i < Rows; i++) {
                    sOrders = sOrders + Columns('columnID').text + '||';
                    sOrders = sOrders + i + '||';
                    sOrders = sOrders + Columns('order').text + '||';
                    sOrders = sOrders + '**';
                    MoveNext();
                }
                Redraw = true;
            }
        }
        frmSend.txtSend_OrderString.value = sOrders;
    }
    else {
	    dataCollection = frmOriginalDefinition.elements;
	    var sOrders = '';
        var sDefnString = '';

        if (dataCollection != null) {
            for (i = 1; i <= frmUseful.txtOrderCount.value; i++) {
                sDefnString = document.getElementById('txtReportDefnOrder_' + i).value;
                sOrders = sOrders + sortColumnParameter(sDefnString, "COLUMNID") + '||';
                sOrders = sOrders + i + '||';
                sOrders = sOrders + sortColumnParameter(sDefnString, "ORDER") + '||';
                sOrders = sOrders + '**';
            }
        }
        frmSend.txtSend_OrderString.value = sOrders;
    }

    /*************************************************************/

    /****************** TAB 5 - OUTPUT OPTIONS *******************/

    if (frmDefinition.chkPreview.checked == true) {
        frmSend.txtSend_OutputPreview.value = 1;
    }
    else {
        frmSend.txtSend_OutputPreview.value = 0;
    }

    if (frmDefinition.optOutputFormat0.checked) frmSend.txtSend_OutputFormat.value = 0;
    //if (frmDefinition.optOutputFormat1.checked)	frmSend.txtSend_OutputFormat.value = 1;
    if (frmDefinition.optOutputFormat2.checked) frmSend.txtSend_OutputFormat.value = 2;
    if (frmDefinition.optOutputFormat3.checked) frmSend.txtSend_OutputFormat.value = 3;
    if (frmDefinition.optOutputFormat4.checked) frmSend.txtSend_OutputFormat.value = 4;

    if (frmDefinition.chkDestination0.checked == true) {
        frmSend.txtSend_OutputScreen.value = 1;
    }
    else {
        frmSend.txtSend_OutputScreen.value = 0;
    }

    if (frmDefinition.chkDestination1.checked == true) {
        frmSend.txtSend_OutputPrinter.value = 1;
        frmSend.txtSend_OutputPrinterName.value = frmDefinition.cboPrinterName.options[frmDefinition.cboPrinterName.selectedIndex].innerText;
    }
    else {
        frmSend.txtSend_OutputPrinter.value = 0;
        frmSend.txtSend_OutputPrinterName.value = '';
    }

    if (frmDefinition.chkDestination2.checked == true) {
        frmSend.txtSend_OutputSave.value = 1;
        frmSend.txtSend_OutputSaveExisting.value = frmDefinition.cboSaveExisting.options[frmDefinition.cboSaveExisting.selectedIndex].value;
    }
    else {
        frmSend.txtSend_OutputSave.value = 0;
        frmSend.txtSend_OutputSaveExisting.value = 0;
    }

    if (frmDefinition.chkDestination3.checked == true) {
        frmSend.txtSend_OutputEmail.value = 1;
        frmSend.txtSend_OutputEmailAddr.value = frmDefinition.txtEmailGroupID.value;
        frmSend.txtSend_OutputEmailSubject.value = frmDefinition.txtEmailSubject.value;
        frmSend.txtSend_OutputEmailAttachAs.value = frmDefinition.txtEmailAttachAs.value;
    }
    else {
        frmSend.txtSend_OutputEmail.value = 0;
        frmSend.txtSend_OutputEmailAddr.value = 0;
        frmSend.txtSend_OutputEmailSubject.value = '';
        frmSend.txtSend_OutputEmailAttachAs.value = '';
    }

    frmSend.txtSend_OutputFilename.value = frmDefinition.txtFilename.value;

    /*************************************************************/

    frmUseful.txtLockGridEvents.value = 0;

	//ND May need to revisit or reinstate this if performance issues ensue
	//try {
	//	testDataCollection = frmRefresh.elements;
	//   iDummy = testDataCollection.txtDummy.value;
  //      frmRefresh.submit();
  //  }
  //  catch (e) {
  //  }

    if (sEvents.length > 16000) {
        OpenHR.messageBox("Too many events selected.", 48, "Calendar Reports");
        return false;
    }
    else {
        return true;
    }

}

function getOrderString() {
    return true;
    var sOrders = '';
    var i;
    var pvarbookmark;

    with (frmDefinition.ssOleDBGridSortOrder) {
        if (Rows > 0) {
            MoveFirst();
            for (var i = 0; i < Rows; i++) {

                sOrders = sOrders + frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value + '||';
                sOrders = sOrders + Columns('columnID').text + '||';
                sOrders = sOrders + i + '||';
                sOrders = sOrders + Columns('order').text + '||';

                sOrders = sOrders + '**';
                MoveNext();
            }
            return sOrders;
        }
        else {
            return '';
        }
    }
}

function loadAvailableColumns() {
    var i;
    var blnPersonnelBaseTable = (frmUseful.txtPersonnelTableID.value == frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);
	frmUseful.txtLockGridEvents.value = 1;

    frmDefinition.cboDescription1.length = 0;
    frmDefinition.cboDescription2.length = 0;
    frmDefinition.cboRegion.length = 0;

    var oOption = document.createElement("OPTION");
    frmDefinition.cboDescription1.options.add(oOption);
    oOption.innerText = "<None>";
    oOption.value = 0;
    if (frmUseful.txtAction.value.toUpperCase() == "NEW") oOption.selected = true;

    oOption = document.createElement("OPTION");
    frmDefinition.cboDescription2.options.add(oOption);
    oOption.innerText = "<None>";
    oOption.value = 0;
    if (frmUseful.txtAction.value.toUpperCase() == "NEW") oOption.selected = true;

    oOption = document.createElement("OPTION");
    frmDefinition.cboRegion.options.add(oOption);


    if (blnPersonnelBaseTable) {
        oOption.innerText = "<Default>";
    }
    else {
        oOption.innerText = "<None>";
    }
    oOption.value = 0;
    if (frmUseful.txtAction.value.toUpperCase() == "NEW") oOption.selected = true;

    //var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
    var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
    var dataCollection = frmUtilDefForm.elements;

    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {

            sControlName = dataCollection.item(i).name;

            if (sControlName.substr(0, 10) == "txtRepCol_") {
                sColumnID = sControlName.substring(10, sControlName.length);

                oOption = document.createElement("OPTION");
                frmDefinition.cboDescription1.options.add(oOption);
                oOption.innerText = dataCollection.item(i).value;
                oOption.value = sColumnID;

                if ((frmUseful.txtAction.value.toUpperCase() != "NEW")
                    && (frmUseful.txtFirstLoad.value == 'Y')) {
                    if (sColumnID == frmOriginalDefinition.txtDefn_Desc1ID.value) {
                        oOption.selected = true;
                    }
                }

                oOption = document.createElement("OPTION");
                frmDefinition.cboDescription2.options.add(oOption);
                oOption.innerText = dataCollection.item(i).value;
                oOption.value = sColumnID;

                if ((frmUseful.txtAction.value.toUpperCase() != "NEW")
                    && (frmUseful.txtFirstLoad.value == 'Y')) {
                    if (sColumnID == frmOriginalDefinition.txtDefn_Desc2ID.value) {
                        oOption.selected = true;
                    }
                }

                /* Only add varchar columns to the region column. */
                var sDataTypeControlName = "txtRepColDataType_" + sColumnID;
                var iDataTypeControlValue = frmUtilDefForm.elements(sDataTypeControlName).value;

                if (iDataTypeControlValue == 12) {
                    oOption = document.createElement("OPTION");
                    frmDefinition.cboRegion.options.add(oOption);
                    oOption.innerText = dataCollection.item(i).value;
                    oOption.value = sColumnID;

                    if ((frmUseful.txtAction.value.toUpperCase() != "NEW")
                        && (frmUseful.txtFirstLoad.value == 'Y')) {
                        if (sColumnID == frmOriginalDefinition.txtDefn_RegionID.value) {
                            oOption.selected = true;
                        }
                    }
                }
            }
        }
    }

    frmUseful.txtLockGridEvents.value = 0;
    frmUseful.txtAvailableColumnsLoaded.value = 1;
    frmUseful.txtFirstLoad.value = 'N';

    frmDefinition.cboBaseTable.style.color = 'windowtext';
    frmDefinition.cboDescription1.style.color = 'windowtext';
    frmDefinition.cboDescription2.style.color = 'windowtext';
    frmDefinition.cboRegion.style.color = 'windowtext';

    if (frmUseful.txtAvailableColumnsLoaded.value == 1) {
        if (frmDefinition.txtName.disabled == false) {
            frmDefinition.focus();
            try {
                frmDefinition.txtName.focus();
            }
            catch (e) { }
        }
    }

    // Get menu to refresh the menu.
    OpenHR.refreshMenu();
		refreshTab1Controls();
    frmUseful.txtLoading.value = 'N';

    if (frmDefinition.txtName.disabled == false) {
        frmDefinition.focus();
        try {
            frmDefinition.txtName.focus();
        }
        catch (e) { }
    }
}

function loadDefinition() {
    frmDefinition.txtName.value = frmOriginalDefinition.txtDefn_Name.value;

    if ((frmUseful.txtAction.value.toUpperCase() == "EDIT") ||
        (frmUseful.txtAction.value.toUpperCase() == "VIEW")) {
        frmDefinition.txtOwner.value = frmOriginalDefinition.txtDefn_Owner.value;
    }
    else {
        frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
    }

    frmDefinition.txtDescription.value = frmOriginalDefinition.txtDefn_Description.value;

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

    // Print Filter Header ?
    frmDefinition.chkPrintFilterHeader.checked = ((frmOriginalDefinition.txtDefn_PrintFilterHeader.value != "False") &&
        ((frmOriginalDefinition.txtDefn_FilterID.value > 0) ||
            (frmOriginalDefinition.txtDefn_PicklistID.value > 0)));

    frmDefinition.txtDescExpr.value = frmOriginalDefinition.txtDefn_DescExprName.value;
    frmDefinition.txtDescExprID.value = frmOriginalDefinition.txtDefn_DescExprID.value;

    frmDefinition.chkGroupByDesc.checked = (frmOriginalDefinition.txtDefn_GroupByDesc.value != "False");

    for (var i = 0; i < frmDefinition.cboDescriptionSeparator.options.length; i++) {
        if (frmDefinition.cboDescriptionSeparator.options[i].value == frmOriginalDefinition.txtDefn_DescSeparator.value) {
            frmDefinition.cboDescriptionSeparator.selectedIndex = i;
            break;
        }
    }

    if (frmOriginalDefinition.txtDefn_StartType.value == "0") {
        frmDefinition.optFixedStart.checked = true;
        frmDefinition.txtFixedStart.value = frmOriginalDefinition.txtDefn_FixedStart.value;
        frmDefinition.cboPeriodStart.selectedIndex = 0;
        frmDefinition.txtFreqStart.value = 0;
        frmDefinition.txtCustomStart.value = '';
        frmDefinition.txtCustomStartID.value = 0;
    }
    else if (frmOriginalDefinition.txtDefn_StartType.value == "1") {
        frmDefinition.optCurrentStart.checked = true;
        frmDefinition.txtFixedStart.value = '';
        frmDefinition.cboPeriodStart.selectedIndex = 0;
        frmDefinition.txtFreqStart.value = 0;
        frmDefinition.txtCustomStart.value = '';
        frmDefinition.txtCustomStartID.value = 0;
    }
    else if (frmOriginalDefinition.txtDefn_StartType.value == "2") {
        frmDefinition.optOffsetStart.checked = true;
        frmDefinition.txtFixedStart.value = '';
        frmDefinition.cboPeriodStart.value = frmOriginalDefinition.txtDefn_StartPeriod.value;
        frmDefinition.txtFreqStart.value = frmOriginalDefinition.txtDefn_StartFrequency.value;
        frmDefinition.txtCustomStart.value = '';
        frmDefinition.txtCustomStartID.value = 0;
    }
    else if (frmOriginalDefinition.txtDefn_StartType.value == "3") {
        frmDefinition.optCustomStart.checked = true;
        frmDefinition.txtFixedStart.value = '';
        frmDefinition.cboPeriodStart.selectedIndex = 0;
        frmDefinition.txtFreqStart.value = 0;
        frmDefinition.txtCustomStart.value = frmOriginalDefinition.txtDefn_CustomStartName.value;
        frmDefinition.txtCustomStartID.value = frmOriginalDefinition.txtDefn_CustomStartID.value;
    }
    else {
        frmDefinition.optFixedStart.checked = true;
        frmDefinition.txtFixedStart.value = '';
        frmDefinition.cboPeriodStart.selectedIndex = 0;
        frmDefinition.txtFreqStart.value = 0;
        frmDefinition.txtCustomStart.value = '';
        frmDefinition.txtCustomStartID.value = 0;
    }

    if (frmOriginalDefinition.txtDefn_EndType.value == "0") {
        frmDefinition.optFixedEnd.checked = true;
        frmDefinition.txtFixedEnd.value = frmOriginalDefinition.txtDefn_FixedEnd.value;
        frmDefinition.cboPeriodEnd.selectedIndex = 0;
        frmDefinition.txtFreqEnd.value = 0;
        frmDefinition.txtCustomEnd.value = '';
        frmDefinition.txtCustomEndID.value = 0;
    }
    else if (frmOriginalDefinition.txtDefn_EndType.value == "1") {
        frmDefinition.optCurrentEnd.checked = true;
        frmDefinition.txtFixedEnd.value = '';
        frmDefinition.cboPeriodEnd.selectedIndex = 0;
        frmDefinition.txtFreqEnd.value = 0;
        frmDefinition.txtCustomEnd.value = '';
        frmDefinition.txtCustomEndID.value = 0;
    }
    else if (frmOriginalDefinition.txtDefn_EndType.value == "2") {
        frmDefinition.optOffsetEnd.checked = true;
        frmDefinition.txtFixedEnd.value = '';
        frmDefinition.cboPeriodEnd.selectedIndex = frmOriginalDefinition.txtDefn_EndPeriod.value;
        frmDefinition.txtFreqEnd.value = frmOriginalDefinition.txtDefn_EndFrequency.value;
        frmDefinition.txtCustomEnd.value = '';
        frmDefinition.txtCustomEndID.value = 0;
    }
    else if (frmOriginalDefinition.txtDefn_EndType.value == "3") {
        frmDefinition.optCustomEnd.checked = true;
        frmDefinition.txtFixedEnd.value = '';
        frmDefinition.cboPeriodEnd.selectedIndex = 0;
        frmDefinition.txtFreqEnd.value = 0;
        frmDefinition.txtCustomEnd.value = frmOriginalDefinition.txtDefn_CustomEndName.value;
        frmDefinition.txtCustomEndID.value = frmOriginalDefinition.txtDefn_CustomEndID.value;
    }
    else {
        frmDefinition.optFixedEnd.checked = true;
        frmDefinition.txtFixedEnd.value = '';
        frmDefinition.cboPeriodEnd.selectedIndex = 0;
        frmDefinition.txtFreqEnd.value = 0;
        frmDefinition.txtCustomEnd.value = '';
        frmDefinition.txtCustomEndID.value = 0;
    }

    //Display Options	
    frmDefinition.chkShadeBHols.checked = (frmOriginalDefinition.txtDefn_ShadeBHols.value != "False");
    frmDefinition.chkCaptions.checked = (frmOriginalDefinition.txtDefn_ShowCaptions.value != "False");
    frmDefinition.chkShadeWeekends.checked = (frmOriginalDefinition.txtDefn_ShadeWeekends.value != "False");
    frmDefinition.chkStartOnCurrentMonth.checked = (frmOriginalDefinition.txtDefn_StartOnCurrentMonth.value != "False");
    frmDefinition.chkIncludeWorkingDaysOnly.checked = (frmOriginalDefinition.txtDefn_IncludeWorkingDaysOnly.value != "False");
    frmDefinition.chkIncludeBHols.checked = (frmOriginalDefinition.txtDefn_IncludeBHols.value != "False");

    if ((frmOriginalDefinition.txtDefn_PicklistHidden.value.toUpperCase() == "TRUE") ||
        (frmOriginalDefinition.txtDefn_FilterHidden.value.toUpperCase() == "TRUE")) {
        frmSelectionAccess.baseHidden.value = "Y";
    }

    if (frmOriginalDefinition.txtDefn_DescExprHidden.value.toUpperCase() == "TRUE") {
        frmSelectionAccess.descHidden.value = "Y";
    }

    if (frmOriginalDefinition.txtDefn_DescExprHidden.value.toUpperCase() == "TRUE") {
        frmSelectionAccess.descHidden.value = "Y";
    }

    if (frmOriginalDefinition.txtDefn_CustomStartCalcHidden.value.toUpperCase() == "TRUE") {
        frmSelectionAccess.calcStartDateHidden.value = "Y";
    }

    if (frmOriginalDefinition.txtDefn_CustomEndCalcHidden.value.toUpperCase() == "TRUE") {
        frmSelectionAccess.calcEndDateHidden.value = "Y";
    }

    frmSelectionAccess.eventHidden.value = frmUseful.txtHiddenEventFilterCount.value;

    //OUTPUT OPTIONS?

    frmDefinition.chkPreview.checked = (frmOriginalDefinition.txtDefn_OutputPreview.value != "False");

    if (frmOriginalDefinition.txtDefn_OutputFormat.value == 0) {
        frmDefinition.optOutputFormat0.checked = true;
    }
        /*else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 1)
            {
            frmDefinition.optOutputFormat1.checked = true;
            }*/
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 2) {
        frmDefinition.optOutputFormat2.checked = true;
    }
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 3) {
        frmDefinition.optOutputFormat3.checked = true;
    }
    else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 4) {
        frmDefinition.optOutputFormat4.checked = true;
    }
    else {
        frmDefinition.optOutputFormat0.checked = true;
    }

    frmDefinition.chkDestination0.checked = (frmOriginalDefinition.txtDefn_OutputScreen.value != "False");
    frmDefinition.chkDestination1.checked = (frmOriginalDefinition.txtDefn_OutputPrinter.value != "False");

    if (frmDefinition.chkDestination1.checked == true) {
        populatePrinters();
        for (i = 0; i < frmDefinition.cboPrinterName.options.length; i++) {
            if (frmDefinition.cboPrinterName.options(i).innerText == frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
                frmDefinition.cboPrinterName.selectedIndex = i;
                break;
            }
        }
    }

    frmDefinition.chkDestination2.checked = (frmOriginalDefinition.txtDefn_OutputSave.value != "False");
    if (frmDefinition.chkDestination2.checked == true) {
        populateSaveExisting();
        frmDefinition.cboSaveExisting.selectedIndex = frmOriginalDefinition.txtDefn_OutputSaveExisting.value;
    }
    frmDefinition.chkDestination3.checked = (frmOriginalDefinition.txtDefn_OutputEmail.value != "False");
    if (frmDefinition.chkDestination3.checked == true) {
        frmDefinition.txtEmailGroupID.value = frmOriginalDefinition.txtDefn_OutputEmailAddr.value;
        frmDefinition.txtEmailGroup.value = frmOriginalDefinition.txtDefn_OutputEmailAddrName.value;
        frmDefinition.txtEmailSubject.value = frmOriginalDefinition.txtDefn_OutputEmailSubject.value;
        frmDefinition.txtEmailAttachAs.value = frmOriginalDefinition.txtDefn_OutputEmailAttachAs.value;
    }
    frmDefinition.txtFilename.value = frmOriginalDefinition.txtDefn_OutputFilename.value;


    frmDefinition.grdEvents.MoveFirst();
    frmDefinition.grdEvents.FirstRow = frmDefinition.grdEvents.Bookmark;

    frmDefinition.ssOleDBGridSortOrder.Movefirst();
    frmDefinition.ssOleDBGridSortOrder.FirstRow = frmDefinition.ssOleDBGridSortOrder.bookmark;

    // If its read only, disable everything.
    if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
        disableAll();
    }

}

function loadEventsDefinition() {
    var iIndex;
    var sDefnString;
    var frmRefresh;

    if (frmUseful.txtEventsLoaded.value == 0) {
        frmDefinition.grdEvents.focus();
        frmDefinition.grdEvents.Redraw = false;

        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {

                sControlName = dataCollection.item(iIndex).name;
                sControlName = sControlName.substr(0, 19);
                if (sControlName == "txtReportDefnEvent_") {
                    sDefnString = new String(dataCollection.item(iIndex).value);

                    if (sDefnString.length > 0) {
                        frmDefinition.grdEvents.AddItem(sDefnString);
                    }
                }
            }
        }
        frmDefinition.grdEvents.Redraw = true;
        frmUseful.txtEventsLoaded.value = 1;
    }
}

function loadSortDefinition() {
    var iIndex;

    if (frmUseful.txtSortLoaded.value == 0) {
        frmDefinition.ssOleDBGridSortOrder.focus();
        frmDefinition.ssOleDBGridSortOrder.Redraw = false;

        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                var sControlName = dataCollection.item(iIndex).name;
                sControlName = sControlName.substr(0, 19);
                if (sControlName == "txtReportDefnOrder_") {
                    var sDefnString = new String(dataCollection.item(iIndex).value);

                    if (sDefnString.length > 0) {
                        frmDefinition.ssOleDBGridSortOrder.AddItem(sDefnString);
                    }
                }
            }
        }
        frmDefinition.ssOleDBGridSortOrder.Redraw = true;
        frmUseful.txtSortLoaded.value = 1;
    }
}

function convertLocaleDateToSQL(psDateString) {
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

    sDateFormat = OpenHR.LocaleDateFormat;

    sDays = "";
    sMonths = "";
    sYears = "";
    iValuePos = 0;

    // Trim leading spaces.
    sTempValue = psDateString.substr(iValuePos, 1);
    while (sTempValue.charAt(0) == " ") {
        iValuePos = iValuePos + 1;
        sTempValue = psDateString.substr(iValuePos, 1);
    }

    for (iLoop = 0; iLoop < sDateFormat.length; iLoop++) {
        if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'D') && (sDays.length == 0)) {
            sDays = psDateString.substr(iValuePos, 1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sDays = sDays.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
        }

        if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'M') && (sMonths.length == 0)) {
            sMonths = psDateString.substr(iValuePos, 1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sMonths = sMonths.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
        }

        if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'Y') && (sYears.length == 0)) {
            sYears = psDateString.substr(iValuePos, 1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
        }

        // Skip non-numerics
        sTempValue = psDateString.substr(iValuePos, 1);
        while (isNaN(sTempValue) == true) {
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);
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

function convertLocaleDateToDateObject(psDateString) {
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

    sDateFormat = OpenHR.LocaleDateFormat;

    sDays = "";
    sMonths = "";
    sYears = "";
    iValuePos = 0;

    // Skip non-numerics
    sTempValue = psDateString.substr(iValuePos, 1);
    while (isNaN(sTempValue) == true) {
        iValuePos = iValuePos + 1;
        sTempValue = psDateString.substr(iValuePos, 1);
    }

    for (iLoop = 0; iLoop < sDateFormat.length; iLoop++) {
        if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'D') && (sDays.length == 0)) {
            sDays = psDateString.substr(iValuePos, 1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sDays = sDays.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
        }

        if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'M') && (sMonths.length == 0)) {
            sMonths = psDateString.substr(iValuePos, 1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sMonths = sMonths.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
        }

        if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'Y') && (sYears.length == 0)) {
            sYears = psDateString.substr(iValuePos, 1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);
            }
            iValuePos = iValuePos + 1;
        }

        // Skip non-numerics
        sTempValue = psDateString.substr(iValuePos, 1);
        while (isNaN(sTempValue) == true) {
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos, 1);
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
    return newDateObj;
}

function populatePrinters() {
    with (frmDefinition.cboPrinterName) {

        strCurrentPrinter = '';
        if (selectedIndex > 0) {
            strCurrentPrinter = options[selectedIndex].innerText;
        }

        length = 0;
        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "<Default Printer>";
        oOption.value = 0;

        for (iLoop = 0; iLoop < OpenHR.PrinterCount() ; iLoop++) {

            var oOption = document.createElement("OPTION");
            options.add(oOption);
            oOption.innerText = OpenHR.PrinterName(iLoop);
            oOption.value = iLoop + 1;

            if (oOption.innerText == strCurrentPrinter) {
                selectedIndex = iLoop + 1;
            }
        }

        if (strCurrentPrinter != '') {
            if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.selectedIndex).innerText != strCurrentPrinter) {
                var oOption = document.createElement("OPTION");
                frmDefinition.cboPrinterName.options.add(oOption);
                oOption.innerText = strCurrentPrinter;
                oOption.value = frmDefinition.cboPrinterName.options.length - 1;
                selectedIndex = oOption.value;
            }
        }
    }
}

function populateSaveExisting() {
    with (frmDefinition.cboSaveExisting) {
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
            (frmDefinition.optOutputFormat6.checked)) {
            var oOption = document.createElement("OPTION");
            options.add(oOption);
            oOption.innerText = "Create new sheet in workbook";
            oOption.value = 4;
        }

        for (iLoop = 0; iLoop < options.length; iLoop++) {
            if (options(iLoop).value == lngCurrentOption) {
                selectedIndex = iLoop;
	            break;
            }
        }

    }

}

function validateDate(pobjDateControl) {
    // Date column.
    // Ensure that the value entered is a date.

    var sValue = pobjDateControl.value;

    if (sValue.length == 0) {
        return true;
    }
    else {
        // Convert the date to SQL format (use this as a validation check).
        // An empty string is returned if the date is invalid.
        sValue = convertLocaleDateToSQL(sValue);
        if (sValue.length == 0) {
            return false;
        }
        else {
            return true;
        }
    }
}

function disableAll() {
    var i;

    var dataCollection = frmDefinition.elements;
    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {
            var eElem = frmDefinition.elements[i];

            if ("text" == eElem.type) {
                text_disable(eElem, true);
            }
            else if ("TEXTAREA" == eElem.tagName) {
                textarea_disable(eElem, true);
            }
            else if ("checkbox" == eElem.type) {
                checkbox_disable(eElem, true);
            }
            else if ("radio" == eElem.type) {
                radio_disable(eElem, true);
            }
            else if ("button" == eElem.type) {
                if (eElem.value != "Cancel") {
                    button_disable(eElem, true);
                }
            }
            else if ("SELECT" == eElem.tagName) {
                combo_disable(eElem, true);
            }
            else {
                grid_disable(eElem, true);
            }
        }
    }
}

function selectedEventParameter(psDefnString, psParameter) {
    var iCharIndex;
    var sDefn;

    sDefn = new String(psDefnString);
    psParameter = psParameter.toUpperCase();

    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) {
        if (psParameter == "NAME") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) {
            if (psParameter == "TABLEID") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            iCharIndex = sDefn.indexOf("	");
            if (iCharIndex >= 0) {
                if (psParameter == "TABLE") return sDefn.substr(0, iCharIndex);
                sDefn = sDefn.substr(iCharIndex + 1);
                iCharIndex = sDefn.indexOf("	");
                if (iCharIndex >= 0) {
                    if (psParameter == "FILTERID") return sDefn.substr(0, iCharIndex);
                    sDefn = sDefn.substr(iCharIndex + 1);
                    iCharIndex = sDefn.indexOf("	");
                    if (iCharIndex >= 0) {
                        if (psParameter == "FILTER") return sDefn.substr(0, iCharIndex);
                        sDefn = sDefn.substr(iCharIndex + 1);
                        iCharIndex = sDefn.indexOf("	");
                        if (iCharIndex >= 0) {
                            if (psParameter == "STARTDATEID") return sDefn.substr(0, iCharIndex);
                            sDefn = sDefn.substr(iCharIndex + 1);
                            iCharIndex = sDefn.indexOf("	");
                            if (iCharIndex >= 0) {
                                if (psParameter == "STARTDATE") return sDefn.substr(0, iCharIndex);
                                sDefn = sDefn.substr(iCharIndex + 1);
                                iCharIndex = sDefn.indexOf("	");
                                if (iCharIndex >= 0) {
                                    if (psParameter == "STARTSESSIONID") return sDefn.substr(0, iCharIndex);
                                    sDefn = sDefn.substr(iCharIndex + 1);
                                    iCharIndex = sDefn.indexOf("	");
                                    if (iCharIndex >= 0) {
                                        if (psParameter == "STARTSESSION") return sDefn.substr(0, iCharIndex);
                                        sDefn = sDefn.substr(iCharIndex + 1);
                                        iCharIndex = sDefn.indexOf("	");
                                        if (iCharIndex >= 0) {
                                            if (psParameter == "ENDDATEID") return sDefn.substr(0, iCharIndex);
                                            sDefn = sDefn.substr(iCharIndex + 1);
                                            iCharIndex = sDefn.indexOf("	");
                                            if (iCharIndex >= 0) {
                                                if (psParameter == "ENDDATE") return sDefn.substr(0, iCharIndex);
                                                sDefn = sDefn.substr(iCharIndex + 1);
                                                iCharIndex = sDefn.indexOf("	");
                                                if (iCharIndex >= 0) {
                                                    if (psParameter == "ENDSESSIONID") return sDefn.substr(0, iCharIndex);
                                                    sDefn = sDefn.substr(iCharIndex + 1);
                                                    iCharIndex = sDefn.indexOf("	");
                                                    if (iCharIndex >= 0) {
                                                        if (psParameter == "ENDSESSION") return sDefn.substr(0, iCharIndex);
                                                        sDefn = sDefn.substr(iCharIndex + 1);
                                                        iCharIndex = sDefn.indexOf("	");
                                                        if (iCharIndex >= 0) {
                                                            if (psParameter == "DURATIONID") return sDefn.substr(0, iCharIndex);
                                                            sDefn = sDefn.substr(iCharIndex + 1);
                                                            iCharIndex = sDefn.indexOf("	");
                                                            if (iCharIndex >= 0) {
                                                                if (psParameter == "DURATION") return sDefn.substr(0, iCharIndex);
                                                                sDefn = sDefn.substr(iCharIndex + 1);
                                                                iCharIndex = sDefn.indexOf("	");
                                                                if (iCharIndex >= 0) {
                                                                    if (psParameter == "LEGENDTYPE") return sDefn.substr(0, iCharIndex);
                                                                    sDefn = sDefn.substr(iCharIndex + 1);
                                                                    iCharIndex = sDefn.indexOf("	");
                                                                    if (iCharIndex >= 0) {
                                                                        if (psParameter == "LEGEND") return sDefn.substr(0, iCharIndex);
                                                                        sDefn = sDefn.substr(iCharIndex + 1);
                                                                        iCharIndex = sDefn.indexOf("	");
                                                                        if (iCharIndex >= 0) {
                                                                            if (psParameter == "LEGENDLOOKUPTABLEID") return sDefn.substr(0, iCharIndex);
                                                                            sDefn = sDefn.substr(iCharIndex + 1);
                                                                            iCharIndex = sDefn.indexOf("	");
                                                                            if (iCharIndex >= 0) {
                                                                                if (psParameter == "LEGENDLOOKUPCOLUMNID") return sDefn.substr(0, iCharIndex);
                                                                                sDefn = sDefn.substr(iCharIndex + 1);
                                                                                iCharIndex = sDefn.indexOf("	");
                                                                                if (iCharIndex >= 0) {
                                                                                    if (psParameter == "LEGENDLOOKUPCODEID") return sDefn.substr(0, iCharIndex);
                                                                                    sDefn = sDefn.substr(iCharIndex + 1);
                                                                                    iCharIndex = sDefn.indexOf("	");
                                                                                    if (iCharIndex >= 0) {
                                                                                        if (psParameter == "LEGENDEVENTCOLUMNID") return sDefn.substr(0, iCharIndex);
                                                                                        sDefn = sDefn.substr(iCharIndex + 1);
                                                                                        iCharIndex = sDefn.indexOf("	");
                                                                                        if (iCharIndex >= 0) {
                                                                                            if (psParameter == "DESC1COLUMNID") return sDefn.substr(0, iCharIndex);
                                                                                            sDefn = sDefn.substr(iCharIndex + 1);
                                                                                            iCharIndex = sDefn.indexOf("	");
                                                                                            if (iCharIndex >= 0) {
                                                                                                if (psParameter == "DESC1COLUMN") return sDefn.substr(0, iCharIndex);
                                                                                                sDefn = sDefn.substr(iCharIndex + 1);
                                                                                                iCharIndex = sDefn.indexOf("	");
                                                                                                if (iCharIndex >= 0) {
                                                                                                    if (psParameter == "DESC2COLUMNID") return sDefn.substr(0, iCharIndex);
                                                                                                    sDefn = sDefn.substr(iCharIndex + 1);
                                                                                                    iCharIndex = sDefn.indexOf("	");
                                                                                                    if (iCharIndex >= 0) {
                                                                                                        if (psParameter == "DESC2COLUMN") return sDefn.substr(0, iCharIndex);
                                                                                                        sDefn = sDefn.substr(iCharIndex + 1);
                                                                                                        iCharIndex = sDefn.indexOf("	");
                                                                                                        if (iCharIndex >= 0) {
                                                                                                            if (psParameter == "EVENTKEY") return sDefn.substr(0, iCharIndex);
                                                                                                            sDefn = sDefn.substr(iCharIndex + 1);

                                                                                                            if (psParameter == "HIDDENFILTER") return sDefn;
                                                                                                        }
                                                                                                    }
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    return "";
}

function sortColumnParameter(psDefnString, psParameter) {
    var iCharIndex;
    var sDefn;

    sDefn = new String(psDefnString);
    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) {
        if (psParameter == "COLUMNID") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) {
            if (psParameter == "COLUMNNAME") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);

            if (psParameter == "ORDER") return sDefn;
        }
    }

    return "";
}

function removeSortColumn(piColumnID, piTableID) {
    // Remove the column (if columnID given), 
    // or all columns for a table (if tableID given),
    // or all columns (if no columnID or tableID given).
    // from the sort columns definition.
    var iCount;
    var i;
    var fRemoveRow;

    if (frmUseful.txtSortLoaded.value == 1) {
        if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
            frmDefinition.ssOleDBGridSortOrder.Redraw = false;
            frmDefinition.ssOleDBGridSortOrder.MoveFirst();

            iCount = frmDefinition.ssOleDBGridSortOrder.rows;
            for (i = 0; i < iCount; i++) {
                fRemoveRow = true;

                if (piColumnID > 0) {
                    fRemoveRow = (piColumnID == frmDefinition.ssOleDBGridSortOrder.Columns("id").Text);
                }
                if (piTableID > 0) {
                    fRemoveRow = (piTableID == frmDefinition.ssOleDBGridSortOrder.Columns("tableID").Text);
                }

                if (fRemoveRow == true) {
                    if (frmDefinition.ssOleDBGridSortOrder.rows == 1) {
                        frmDefinition.ssOleDBGridSortOrder.RemoveAll();
                    }
                    else {
                        frmDefinition.ssOleDBGridSortOrder.RemoveItem(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark));
                    }
                }
                else {
                    frmDefinition.ssOleDBGridSortOrder.MoveNext();
                }
            }

            frmDefinition.ssOleDBGridSortOrder.Redraw = true;
            frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();

            if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
                frmDefinition.ssOleDBGridSortOrder.MoveFirst();
                frmDefinition.ssOleDBGridSortOrder.selbookmarks.add(frmDefinition.ssOleDBGridSortOrder.bookmark);
            }
        }
    }
    else {
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                sControlName = dataCollection.item(iIndex).name;
                sControlName = sControlName.substr(0, 19);
                if (sControlName == "txtReportDefnOrder_") {
                    fRemoveRow = true;
                    if (piColumnID > 0) {
                        fRemoveRow = (piColumnID == sortColumnParameter(dataCollection.item(iIndex).value, "COLUMNID"));
                    }
                    if (piTableID > 0) {
                        fRemoveRow = (piTableID == sortColumnParameter(dataCollection.item(iIndex).value, "TABLEID"));
                    }

                    if (fRemoveRow == true) {
                        dataCollection.item(iIndex).value = "";
                    }
                }
            }
        }
    }
}

function removeCalcs(psCalcs) {
    var iCharIndex;
    var sCalcs;
    var sCalcID;
    var fDone;

    sCalcs = new String(psCalcs);

    // Remove the given calcs from the selected columns list.
    while (sCalcs.length > 0) {
        iCharIndex = sCalcs.indexOf(",");

        if (iCharIndex >= 0) {
            sCalcID = sCalcs.substr(0, iCharIndex);
            sCalcs = sCalcs.substr(iCharIndex + 1);
        }
        else {
            sCalcID = sCalcs;
            sCalcs = "";
        }

        fDone = false;

        /* Check if we're removing the description calculation first, then Custom Date1 then Custom Date2. */
        if ((fDone == false) && (frmDefinition.txtDescExprID.value == sCalcID)) {
            frmDefinition.txtDescExpr.value = '';
            frmDefinition.txtDescExprID.value = 0;
            frmSelectionAccess.descHidden.value = "N";
            fDone = true;
        }

        if ((fDone == false) && (frmDefinition.txtCustomStartID.value == sCalcID)) {
            frmDefinition.txtCustomStart.value = '';
            frmDefinition.txtCustomStartID.value = 0;
            frmSelectionAccess.calcStartDateHidden.value = "N";
            fDone = true;
        }

        if ((fDone == false) && (frmDefinition.txtCustomEndID.value == sCalcID)) {
            frmDefinition.txtCustomEnd.value = '';
            frmDefinition.txtCustomEndID.value = 0;
            frmSelectionAccess.calcEndDateHidden.value = "N";
            fDone = true;
        }
    }

    refreshTab1Controls();
    refreshTab3Controls();
}

function removePicklists(psPicklists) {
    var iCharIndex;
    var sPicklists;
    var sPicklistID;
    var fDone;

    sPicklists = new String(psPicklists);

    // Remove the given calcs from the selected columns list.
    while (sPicklists.length > 0) {
        iCharIndex = sPicklists.indexOf(",");

        if (iCharIndex >= 0) {
            sPicklistID = sPicklists.substr(0, iCharIndex);
            sPicklists = sPicklists.substr(iCharIndex + 1);
        }
        else {
            sPicklistID = sPicklists;
            sPicklists = "";
        }

        fDone = false;

        /* Check if we're removing the base table first, then paretn1 then parent 2, and then the children. */
        if ((fDone == false) && (frmDefinition.txtBasePicklistID.value == sPicklistID)) {
            frmDefinition.txtBasePicklist.value = '';
            frmDefinition.txtBasePicklistID.value = 0;
            frmSelectionAccess.baseHidden.value = "N";
            fDone = true;
        }
    }

    refreshTab1Controls();
}

function removeFilters(psEventFilters) {
    var iCharIndex;
    var sEventFilters;
    var sEventFilterID;
    var sGridEventFilterID;

    sEventFilters = new String(psEventFilters);

    // Remove the given calcs from the selected columns list.
    while (sEventFilters.length > 0) {
        iCharIndex = sEventFilters.indexOf(",");

        if (iCharIndex >= 0) {
            sEventFilterID = sEventFilters.substr(0, iCharIndex);
            sEventFilters = sEventFilters.substr(iCharIndex + 1);
        }
        else {
            sEventFilterID = sEventFilters;
            sEventFilters = "";
        }

        if (frmUseful.txtChildsLoaded.value == 1) {
            if (frmDefinition.grdEvents.Rows > 0) {
                frmDefinition.grdEvents.Redraw = false;
                frmDefinition.grdEvents.movefirst();

                for (i = 0; i < frmDefinition.grdEvents.rows; i++) {
                    sGridEventFilterID = frmDefinition.grdEvents.Columns("FilterID").Text;

                    if (sGridEventFilterID == sEventFilterID) {
                        frmDefinition.grdEvents.RemoveItem(frmDefinition.grdEvents.AddItemRowIndex(frmDefinition.grdEvents.Bookmark));
                        break;
                    }

                    frmDefinition.grdEvents.movenext();
                }

                frmDefinition.grdEvents.Redraw = true;
            }
        }
    }
}

function createNew(pPopup) {
    pPopup.close();

    frmUseful.txtUtilID.value = 0;
    frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
    frmUseful.txtAction.value = "new";

    submitDefinition();
}

function locateRecord(psSearchFor) {
	var fFound;
	fFound = false;
	return;
	frmDefinition.ssOleDBGridAvailableColumns.redraw = false;

	frmDefinition.ssOleDBGridAvailableColumns.MoveLast();
	frmDefinition.ssOleDBGridAvailableColumns.MoveFirst();

	frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.removeall();

	for (iIndex = 1; iIndex <= frmDefinition.ssOleDBGridAvailableColumns.rows; iIndex++) {
		var sGridValue = new String(frmDefinition.ssOleDBGridAvailableColumns.Columns(3).value);
		sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
		if (sGridValue == psSearchFor.toUpperCase()) {
			frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.Add(frmDefinition.ssOleDBGridAvailableColumns.Bookmark);
			fFound = true;
			break;
		}

		if (iIndex < frmDefinition.ssOleDBGridAvailableColumns.rows) {
			frmDefinition.ssOleDBGridAvailableColumns.MoveNext();
		}
		else {
			break;
		}
	}

	if ((fFound == false) && (frmDefinition.ssOleDBGridAvailableColumns.rows > 0)) {
		// Select the top row.
		frmDefinition.ssOleDBGridAvailableColumns.MoveFirst();
		frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.Add(frmDefinition.ssOleDBGridAvailableColumns.Bookmark);
	}

	frmDefinition.ssOleDBGridAvailableColumns.redraw = true;
}

function populateAccessGrid() {
    frmDefinition.grdAccess.focus();
    frmDefinition.grdAccess.removeAll();

    var dataCollection = frmAccess.elements;
    if (dataCollection != null) {

        frmDefinition.grdAccess.AddItem("(All Groups)");
        for (i = 0; i < dataCollection.length; i++) {
            frmDefinition.grdAccess.AddItem(dataCollection.item(i).value);
        }
    }
}

function setJobsToHide(psJobs) {
    frmSend.txtSend_jobsToHide.value = psJobs;
    frmSend.txtSend_jobsToHideGroups.value = frmValidate.validateHiddenGroups.value;
}

function changeTab1Control() {
    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}

function changeTab3Control() {

    frmUseful.txtChanged.value = 1;
    refreshTab3Controls();
}

function changeTab5Control() {
    frmUseful.txtChanged.value = 1;
    refreshTab5Controls();
}

function recalcHiddenEventFiltersCount() {
    var iCount;
    var vBM;

    iCount = 0;

    for (var i = 0; i < frmDefinition.grdEvents.Rows; i++) {
        vBM = frmDefinition.grdEvents.AddItemBookmark(i);
	    if (frmDefinition.grdEvents.Columns("FilterHidden").CellValue(vBM) == "Y") {
            iCount = iCount + 1;
        }
    }
    frmSelectionAccess.eventHidden.value = iCount;
}
	
function util_def_calendarreport_addhandlers() {
	    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "beforeupdate", ssOleDBGridSortOrderbeforeupdate);
	    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "afterinsert", ssOleDBGridSortOrderafterinsert);
	    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "rowColChange", ssOleDBGridSortOrderRowColChange);
	    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "Change", ssOleDBGridSortOrderchange);
	  
	    OpenHR.addActiveXHandler("grdEvents", "click", grdEventsclick);
	    OpenHR.addActiveXHandler("grdEvents", "DblClick", grdEventsDblClick);
	    OpenHR.addActiveXHandler("grdEvents", "rowcolchange", grdEventsrowcolchange);
		
	    OpenHR.addActiveXHandler("grdAccess", "ComboCloseUp", grdAccessComboCloseUp);
	    OpenHR.addActiveXHandler("grdAccess", "GotFocus", grdAccessGotFocus);
	    OpenHR.addActiveXHandler("grdAccess", "RowColChange", grdAccessRowColChange);
	    OpenHR.addActiveXHandler("grdAccess", "RowLoaded", grdAccessRowLoaded);
	}
function ssOleDBGridSortOrderbeforeupdate() {
		//<script FOR=ssOleDBGridSortOrder EVENT=beforeupdate LANGUAGE=JavaScript>
    if ((frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Asc') && (frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Desc'))
    {
    	frmDefinition.ssOleDBGridSortOrder.Columns(2).text = 'Asc';
    }
}
function ssOleDBGridSortOrderafterinsert() {
	//<script FOR=ssOleDBGridSortOrder EVENT=afterinsert LANGUAGE=JavaScript>
	refreshTab4Controls();
}
function ssOleDBGridSortOrderRowColChange() {
	//<script FOR=ssOleDBGridSortOrder EVENT=rowcolchange LANGUAGE=JavaScript>
    frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
    frmDefinition.ssOleDBGridSortOrder.columns(1).cellstyleset("ssetSelected", frmDefinition.ssOleDBGridSortOrder.row);
    frmSortOrder.txtSortColumnID.value = frmDefinition.ssOleDBGridSortOrder.Columns(0).text;
    frmSortOrder.txtSortColumnName.value = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
    frmSortOrder.txtSortOrder.value = frmDefinition.ssOleDBGridSortOrder.Columns(2).text;
    refreshTab4Controls();
}
function ssOleDBGridSortOrderchange() {
	//<SCRIPT FOR=ssOleDBGridSortOrder EVENT=change LANGUAGE=JavaScript>
	frmUseful.txtChanged.value = 1;
}
function grdEventsSortOrderchange() {
	//<script for=grdEvents event=change language=JavaScript>
	frmUseful.txtChanged.value = 1;
}
function grdEventsclick() {
	//<script for=grdEvents event=click language=JavaScript>
	refreshTab2Controls();
}
function grdEventsDblClick() {
    if (frmUseful.txtAction.value.toUpperCase() != "VIEW") {
        if (frmDefinition.grdEvents.Rows > 0
            && frmDefinition.grdEvents.SelBookmarks.Count == 1) {
            eventEdit();
        }
        else {
            eventAdd();
        }
    }
}
function grdEventsrowcolchange() {
	//<script FOR=grdEvents EVENT=rowcolchange LANGUAGE=JavaScript>
	frmDefinition.grdEvents.SelBookmarks.Add(frmDefinition.grdEvents.Bookmark);
    frmDefinition.grdEvents.columns('Table').cellstyleset("ssetSelected", frmDefinition.grdEvents.row);
    refreshTab2Controls();
}

function grdAccessComboCloseUp() {
	//<script FOR=grdAccess EVENT=ComboCloseUp LANGUAGE=JavaScript>
	frmUseful.txtChanged.value = 1;
    if ((frmDefinition.grdAccess.AddItemRowIndex(frmDefinition.grdAccess.Bookmark) == 0) &&
        (frmDefinition.grdAccess.Columns("Access").Text.length > 0)) {
        ForceAccess(frmDefinition.grdAccess, AccessCode(frmDefinition.grdAccess.Columns("Access").Text));

        frmDefinition.grdAccess.MoveFirst();
        frmDefinition.grdAccess.Col = 1;
    }
    refreshTab1Controls();
}
function grdAccessGotFocus() {
	//<script FOR=grdAccess EVENT=GotFocus LANGUAGE=JavaScript>
	frmDefinition.grdAccess.Col = 1;
}
function grdAccessRowColChange(LastRow, LastCol) {
	//<script FOR=grdAccess EVENT=RowColChange(LastRow, LastCol) LANGUAGE=JavaScript>
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
    } else {
    	frmDefinition.grdAccess.Columns("Access").Style = 3; // 3 = Combo box
    	frmDefinition.grdAccess.Columns("Access").RemoveAll();
    	frmDefinition.grdAccess.Columns("Access").AddItem(AccessDescription("RW"));
    	frmDefinition.grdAccess.Columns("Access").AddItem(AccessDescription("RO"));
    	frmDefinition.grdAccess.Columns("Access").AddItem(AccessDescription("HD"));
    }

    frmDefinition.grdAccess.Col = 1;
}
function grdAccessRowLoaded(Bookmark) {
	//<script FOR=grdAccess EVENT=RowLoaded(Bookmark) LANGUAGE=JavaScript>
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
    } else {
    	if (frmDefinition.grdAccess.Columns("SysSecMgr").CellText(Bookmark) == "1") {
    		frmDefinition.grdAccess.Columns("GroupName").CellStyleSet("SysSecMgr");
    		frmDefinition.grdAccess.Columns("Access").CellStyleSet("SysSecMgr");
    		frmDefinition.grdAccess.ForeColor = "0";
        } else {
    		frmDefinition.grdAccess.ForeColor = "0";
        }
    }
}

function eventAdd() {
	var sURL;
	var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");
	debugger;
	frmEvent.eventAction.value = "NEW";
	frmEvent.eventID.value = getEventKey();
	frmEvent.eventFilterHidden.value = "";

	if (frmDefinition.grdEvents.Rows < 999) {
		sURL = "util_def_calendarreportdates_main" +
				"?eventAction=" + escape(frmEvent.eventAction.value) +
				"&eventName=" + escape(frmEvent.eventName.value) +
				"&eventID=" + escape(frmEvent.eventID.value) +
				"&eventTableID=" + escape(frmEvent.eventTableID.value) +
				"&eventTable=" + escape(frmEvent.eventTable.value) +
				"&eventFilterID=" + escape(frmEvent.eventFilterID.value) +
				"&eventFilter=" + escape(frmEvent.eventFilter.value) +
				"&eventFilterHidden=" + escape(frmEvent.eventFilterHidden.value) +
				"&eventStartDateID=" + escape(frmEvent.eventStartDateID.value) +
				"&eventStartDate=" + escape(frmEvent.eventStartDate.value) +
				"&eventStartSessionID=" + escape(frmEvent.eventStartSessionID.value) +
				"&eventStartSession=" + escape(frmEvent.eventStartSession.value) +
				"&eventEndDateID=" + escape(frmEvent.eventEndDateID.value) +
				"&eventEndDate=" + escape(frmEvent.eventEndDate.value) +
				"&eventEndSessionID=" + escape(frmEvent.eventEndSessionID.value) +
				"&eventEndSession=" + escape(frmEvent.eventEndSession.value) +
				"&eventDurationID=" + escape(frmEvent.eventDurationID.value) +
				"&eventDuration=" + escape(frmEvent.eventDuration.value) +
				"&eventLookupType=" + escape(frmEvent.eventLookupType.value) +
				"&eventKeyCharacter=" + escape(frmEvent.eventKeyCharacter.value) +
				"&eventLookupTableID=" + escape(frmEvent.eventLookupTableID.value) +
				"&eventLookupColumnID=" + escape(frmEvent.eventLookupColumnID.value) +
				"&eventLookupCodeID=" + escape(frmEvent.eventLookupCodeID.value) +
				"&eventTypeColumnID=" + escape(frmEvent.eventTypeColumnID.value) +
				"&eventDesc1ID=" + escape(frmEvent.eventDesc1ID.value) +
				"&eventDesc1=" + escape(frmEvent.eventDesc1.value) +
				"&eventDesc2ID=" + escape(frmEvent.eventDesc2ID.value) +
				"&eventDesc2=" + escape(frmEvent.eventDesc2.value) +
				"&relationNames=" + escape(frmEvent.relationNames.value);
		openDialog(sURL, 650, 500, "yes", "yes");
		frmUseful.txtChanged.value = 1;
	}
	else {
		var sMessage = "";
		sMessage = "The maximum of 999 events has been selected.";
		OpenHR.messageBox(sMessage, 64, "Calendar Reports");
	}
	refreshTab2Controls();
}

function loadAvailableEventColumns() {
	var i;
	var frmPopup = document.getElementById("frmPopup");
	frmPopup.cboStartDate.length = 0;
	frmPopup.cboStartSession.length = 0;
	frmPopup.cboEndDate.length = 0;
	frmPopup.cboEndSession.length = 0;
	frmPopup.cboDuration.length = 0;
	frmPopup.cboEventType.length = 0;
	frmPopup.cboEventDesc1.length = 0;
	frmPopup.cboEventDesc2.length = 0;
	frmPopup.cboEventType.length = 0;

	var oOption = document.createElement("OPTION");
	frmPopup.cboStartSession.options.add(oOption);
	oOption.innerText = "<None>";
	oOption.value = 0;

	oOption = document.createElement("OPTION");
	frmPopup.cboEndDate.options.add(oOption);
	oOption.innerText = "<None>";
	oOption.value = 0;

	oOption = document.createElement("OPTION");
	frmPopup.cboEndSession.options.add(oOption);
	oOption.innerText = "<None>";
	oOption.value = 0;

	oOption = document.createElement("OPTION");
	frmPopup.cboDuration.options.add(oOption);
	oOption.innerText = "<None>";
	oOption.value = 0;

	oOption = document.createElement("OPTION");
	frmPopup.cboEventDesc1.options.add(oOption);
	oOption.innerText = "<None>";
	oOption.value = 0;

	oOption = document.createElement("OPTION");
	frmPopup.cboEventDesc2.options.add(oOption);
	oOption.innerText = "<None>";
	oOption.value = 0;

	combo_disable(frmPopup.cboStartDate, true);
	combo_disable(frmPopup.cboStartSession, true);
	combo_disable(frmPopup.cboEndDate, true);
	combo_disable(frmPopup.cboEndSession, true);
	combo_disable(frmPopup.cboDuration, true);
	combo_disable(frmPopup.cboEventDesc1, true);
	combo_disable(frmPopup.cboEventDesc2, true);
	combo_disable(frmPopup.cboEventType, true);

	var frmUtilDefForm = OpenHR.getForm("calendardataframe", "frmCalendarData");
	var dataCollection = frmUtilDefForm.elements;

	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {

			var sControlName = dataCollection.item(i).name;

			frmPopup.cboStartDate.selectedIndex = -1;
			frmPopup.cboStartSession.selectedIndex = -1;
			frmPopup.cboEndDate.selectedIndex = -1;
			frmPopup.cboEndSession.selectedIndex = -1;
			frmPopup.cboDuration.selectedIndex = -1;
			frmPopup.cboEventType.selectedIndex = -1;
			frmPopup.cboEventDesc1.selectedIndex = -1;
			frmPopup.cboEventDesc2.selectedIndex = -1;
			frmPopup.cboEventType.selectedIndex = -1;

			if (sControlName.substr(0, 10) == "txtRepCol_") {
				var sColumnID = sControlName.substring(10, sControlName.length);
				var sTableIDControlName = "txtRepColTableID_" + sColumnID;
				var iTableIDControlValue = frmUtilDefForm.elements(sTableIDControlName).value;
				var sTableNameControlName = "txtRepColTableName_" + sColumnID;
				var sTableNameControlValue = frmUtilDefForm.elements(sTableNameControlName).value;

				if (iTableIDControlValue == frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].value) {
					var sDataTypeControlName = "txtRepColDataType_" + sColumnID;
					var iDataTypeControlValue = frmUtilDefForm.elements(sDataTypeControlName).value;
					var sSizeControlName = "txtRepColSize_" + sColumnID;
					var iSizeControlValue = frmUtilDefForm.elements(sSizeControlName).value;
					var sTypeControlName = "txtRepColType_" + sColumnID;
					var iTypeControlValue = frmUtilDefForm.elements(sTypeControlName).value;

					if (iDataTypeControlValue == 11) {
						oOption = document.createElement("OPTION");
						frmPopup.cboStartDate.options.add(oOption);
						oOption.innerText = dataCollection.item(i).value;
						oOption.value = sColumnID;

						oOption = document.createElement("OPTION");
						frmPopup.cboEndDate.options.add(oOption);
						oOption.innerText = dataCollection.item(i).value;
						oOption.value = sColumnID;
					}

					if (iDataTypeControlValue == 12 && iSizeControlValue == 2) {
						oOption = document.createElement("OPTION");
						frmPopup.cboStartSession.options.add(oOption);
						oOption.innerText = dataCollection.item(i).value;
						oOption.value = sColumnID;

						oOption = document.createElement("OPTION");
						frmPopup.cboEndSession.options.add(oOption);
						oOption.innerText = dataCollection.item(i).value;
						oOption.value = sColumnID;
					}

					if (iDataTypeControlValue == 2 || iDataTypeControlValue == 4) {
						oOption = document.createElement("OPTION");
						frmPopup.cboDuration.options.add(oOption);
						oOption.innerText = dataCollection.item(i).value;
						oOption.value = sColumnID;
					}

					if (iTypeControlValue == 1) {
						oOption = document.createElement("OPTION");
						frmPopup.cboEventType.options.add(oOption);
						oOption.innerText = dataCollection.item(i).value;
						oOption.value = sColumnID;
					}
				}

				oOption = document.createElement("OPTION");
				frmPopup.cboEventDesc1.options.add(oOption);
				oOption.innerText = sTableNameControlValue + '.' + dataCollection.item(i).value;
				oOption.value = sColumnID;

				oOption = document.createElement("OPTION");
				frmPopup.cboEventDesc2.options.add(oOption);
				oOption.innerText = sTableNameControlValue + '.' + dataCollection.item(i).value;
				oOption.value = sColumnID;
			}
		}
	}

	if ((frmPopup.cboStartDate.selectedIndex < 0)
			&& (frmPopup.cboStartDate.length > 0)) {
		frmPopup.cboStartDate.selectedIndex = 0;
	}

	var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");

	if (frmPopup.cboStartDate.length < 1) {
		OpenHR.MessageBox("The selected event table has no date columns. Please select an event table that contains date columns.", 48, "Calendar Reports");
		frmPopup.txtNoDateColumns.value = 1;
		refreshEventControls();
		refreshLegendControls();
	}
	else {

		frmPopup.txtNoDateColumns.value = 0;

		if ((frmEvent.eventAction.value.toUpperCase() == "EDIT")
				&& (frmPopup.txtFirstLoad_Event.value == 1)) {
			setEventValues();
			frmPopup.txtFirstLoad_Event.value = 0;
		}
		else {
			refreshEventControls();
			refreshLegendControls();
		}
	}

	if (frmEvent.eventLookupType.value != 1) {
		button_disable(frmPopup.cmdCancel, false);
	}

	frmPopup.txtEventColumnsLoaded.value = 1;
}

function populateLookupColumns() {
	var frmPopup = document.getElementById("frmPopup");
	var frmGetDataForm = OpenHR.getForm("calendardataframe", "frmGetCalendarData");

	if (frmPopup.txtLookupColumnsLoaded.value == 0) {
		// Get the columns/calcs for the current table selection.
		if ((frmPopup.cboLegendTable.options.length > 0) && (frmPopup.cboLegendTable.selectedIndex < 0))
			{
				frmPopup.cboLegendTable.selectedIndex = 0;
			}
		frmGetDataForm.txtCalendarAction.value = "LOADCALENDAREVENTKEYLOOKUPCOLUMNS";
		frmGetDataForm.txtCalendarLookupTableID.value = frmPopup.cboLegendTable.options[frmPopup.cboLegendTable.selectedIndex].value;
		//window.parent.frames("calendardataframe").refreshData();
		data_refreshData();
		//OpenHR.getForm("calendardataframe").refreshData();
	}
	else {
		return;
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

function loadAvailableLookupColumns() {
	var i;
	var sSelectedIDs;
	var sTemp;
	var iIndex;
	var sType;
	var sID;
	var iDummy;
	var frmRefresh;
	var frmPopup = document.getElementById("frmPopup");
	frmPopup.cboLegendColumn.length = 0;
	frmPopup.cboLegendCode.length = 0;

	var frmUtilDefForm = OpenHR.getForm("calendardataframe", "frmCalendarData");
	var dataCollection = frmUtilDefForm.elements;

	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			var sControlName = dataCollection.item(i).name;

			if (sControlName.substr(0, 10) == "txtRepCol_") {
				var sColumnID = sControlName.substring(10, sControlName.length);
				var sDataTypeControlName = "txtRepColDataType_" + sColumnID;
				var iDataTypeControlValue = frmUtilDefForm.elements(sDataTypeControlName).value;

				if (iDataTypeControlValue == 12) {
					var oOption = document.createElement("OPTION");
					frmPopup.cboLegendColumn.options.add(oOption);
					oOption.innerText = dataCollection.item(i).value;
					oOption.value = sColumnID;

					oOption = document.createElement("OPTION");
					frmPopup.cboLegendCode.options.add(oOption);
					oOption.innerText = dataCollection.item(i).value;
					oOption.value = sColumnID;
				}
			}
		}
	}
	//document.parentWindow.parent.window.dialogArguments.window.refreshTab3Controls();		  
	///refreshTab3Controls();

	var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");

	if ((frmEvent.eventAction.value.toUpperCase() == "EDIT") && (frmPopup.txtFirstLoad_Lookup.value == 1)) {
		setLookupValues();
		frmPopup.txtFirstLoad_Lookup.value = 0;
	}
	else {
		refreshLegendControls();
	}

	button_disable(frmPopup.cmdCancel, false);
	frmPopup.txtLookupColumnsLoaded.value = 1;
}

function setEventTable(piTableID) {
	var i;
	var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboEventTable.options.length; i++) {
        if (frmPopup.cboEventTable.options(i).value == piTableID) {
            frmPopup.cboEventTable.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboEventTable.selectedIndex = 0;
}

function setStartDate(piColumnID) {
	var i;
	var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboStartDate.options.length; i++) {
        if (frmPopup.cboStartDate.options(i).value == piColumnID) {
            frmPopup.cboStartDate.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboStartDate.selectedIndex = 0;
}

function setStartSession(piColumnID) {
	var frmPopup = document.getElementById("frmPopup");
    var i;
    for (i = 0; i < frmPopup.cboStartSession.options.length; i++) {
        if (frmPopup.cboStartSession.options(i).value == piColumnID) {
            frmPopup.cboStartSession.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboStartSession.selectedIndex = 0;
}4

function setEndDate(piColumnID) {
	var i;
	var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboEndDate.options.length; i++) {
        if (frmPopup.cboEndDate.options(i).value == piColumnID) {
            frmPopup.cboEndDate.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboEndDate.selectedIndex = 0;
}

function setEndSession(piColumnID) {
	var i;
	var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboEndSession.options.length; i++) {
        if (frmPopup.cboEndSession.options(i).value == piColumnID) {
            frmPopup.cboEndSession.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboEndSession.selectedIndex = 0;
}

function setDuration(piColumnID) {
	var i;
	var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboDuration.options.length; i++) {
        if (frmPopup.cboDuration.options(i).value == piColumnID) {
            frmPopup.cboDuration.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboDuration.selectedIndex = 0;
}

function setLookupTable(piColumnID) {
	var i;
	var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboLegendTable.options.length; i++) {
        if (frmPopup.cboLegendTable.options(i).value == piColumnID) {
            frmPopup.cboLegendTable.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboLegendTable.selectedIndex = 0;
}

function setLookupColumn(piColumnID) {
	var i;
	var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboLegendColumn.options.length; i++) {
        if (frmPopup.cboLegendColumn.options(i).value == piColumnID) {
            frmPopup.cboLegendColumn.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboLegendColumn.selectedIndex = 0;
}

function setLookupCode(piColumnID) {
	var i;
	var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboLegendCode.options.length; i++) {
        if (frmPopup.cboLegendCode.options(i).value == piColumnID) {
            frmPopup.cboLegendCode.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboLegendCode.selectedIndex = 0;
}

function setEventTypeColumn(piColumnID) {
	var i; var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboEventType.options.length; i++) {
        if (frmPopup.cboEventType.options(i).value == piColumnID) {
            frmPopup.cboEventType.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboEventType.selectedIndex = 0;
}

function setDesc1Column(piColumnID) {
	var i; var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboEventDesc1.options.length; i++) {
        if (frmPopup.cboEventDesc1.options(i).value == piColumnID) {
            frmPopup.cboEventDesc1.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboEventDesc1.selectedIndex = 0;
}

function setDesc2Column(piColumnID) {
	var i; var frmPopup = document.getElementById("frmPopup");
    for (i = 0; i < frmPopup.cboEventDesc2.options.length; i++) {
        if (frmPopup.cboEventDesc2.options(i).value == piColumnID) {
            frmPopup.cboEventDesc2.selectedIndex = i;
            return;
        }
    }
    frmPopup.cboEventDesc2.selectedIndex = 0;
}

function setLookupValues() {
	var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");
	var frmPopup = document.getElementById("frmPopup");
    if (frmPopup.txtHaveSetLookupValues.value == 1) {
        return;
    }

    with (frmEvent) {
        if (eventLookupType.value == 1) {
            frmPopup.optLegendLookup.checked = true;

            setLookupTable(eventLookupTableID.value);
            setLookupColumn(eventLookupColumnID.value);
            setLookupCode(eventLookupCodeID.value);
            setEventTypeColumn(eventTypeColumnID.value);
        }
        else {
            frmPopup.optCharacter.checked = true;
            frmPopup.txtCharacter.value = eventKeyCharacter.value;
        }
    }
    frmPopup.txtHaveSetLookupValues.value = 1;
}

function setEventValues() {
	//var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");
	//var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");
	var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");
	var frmPopup = document.getElementById("frmPopup");
	setStartDate(frmEvent.eventStartDateID.value);
	if (frmEvent.eventStartSessionID.value > 0) {
		setStartSession(frmEvent.eventStartSessionID.value);
	}
	if (frmEvent.eventEndDateID.value > 0) {
		frmPopup.optEndDate.checked = true;
		setEndDate(frmEvent.eventEndDateID.value);
		if (frmEvent.eventEndSessionID.value > 0) {
			setEndSession(frmEvent.eventEndSessionID.value);
		}
	}
	else if (frmEvent.eventDurationID.value > 0) {
		frmPopup.optDuration.checked = true;
		setDuration(frmEvent.eventDurationID.value);
	}
	else {
		frmPopup.optNoEnd.checked = true;
	}

	if (frmEvent.eventDesc1ID.value > 0) {
		setDesc1Column(frmEvent.eventDesc1ID.value);
	}

	if (frmEvent.eventDesc2ID.value > 0) {
		setDesc2Column(frmEvent.eventDesc2ID.value);
	}

	refreshEventControls();

	if ((frmEvent.eventLookupType.value == 1)) {
		setLookupTable(frmEvent.eventLookupTableID.value);

		populateLookupColumns();
	}
	else {
		setLookupValues();
	}
	refreshLegendControls();
}

function getTableName(piTableID) {
    var i;
    var sTableName = new String("");
    var frmTab = OpenHR.getForm("workframe", "frmTables");
    var sReqdControlName = new String("txtTableName_");
    sReqdControlName = sReqdControlName.concat(piTableID);

    var dataCollection = frmTab.elements;
    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {
            var sControlName = dataCollection.item(i).name;

            if (sControlName == sReqdControlName) {
                sTableName = dataCollection.item(i).value;
                return sTableName;
            }
        }
    }
}

function trim(strInput) {
    if (strInput.length < 1) {
        return "";
    }

    while (strInput.substr(strInput.length - 1, 1) == " ") {
        strInput = strInput.substr(0, strInput.length - 1);
    }

    while (strInput.substr(0, 1) == " ") {
        strInput = strInput.substr(1, strInput.length);
    }

    return strInput;
}



function checkUniqueEventName(psEventName) {
    //CODE REQUIRED TO CHECK THAT THE EVENT NAME IS UNIQUE 
    return true;
}

function refreshEventControls() {
	//var frmUse = OpenHR.getForm("workframe", "frmUseful");
	var frmUse = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");
	var fViewing = (frmUse.txtAction.value.toUpperCase() == "VIEW");
	var frmPopup = document.getElementById("frmPopup");

	with (frmPopup) {

		if (txtNoDateColumns.value == 1) {
			button_disable(cmdEventFilter, true);
			cboStartDate.length = 0;
			cboStartSession.length = 0;
			cboEndDate.length = 0;
			cboEndSession.length = 0;
			cboDuration.length = 0;
			cboEventDesc1.length = 0;
			cboEventDesc2.length = 0;
		}

		/* Event Frame */
		if (fViewing) {
			text_disable(txtEventName, true);
			combo_disable(cboEventTable, true);
			text_disable(txtEventFilter, true);
			button_disable(cmdEventFilter, true);
		}
		else {
			text_disable(txtEventName, false);
			combo_disable(cboEventTable, false);
			text_disable(txtEventFilter, true);
			button_disable(cmdEventFilter, (txtNoDateColumns.value == 1));
		}

		/* Event Start Frame */
		if (fViewing) {
			combo_disable(cboStartDate, true);
			combo_disable(cboStartSession, true);
		}
		else {
			if (txtNoDateColumns.value == 1) {
				combo_disable(cboStartDate, true);
				combo_disable(cboStartSession, true);
			}
			else {
				combo_disable(cboStartDate, false);
				combo_disable(cboStartSession, false);
			}
		}

		if ((cboStartDate.options.length > 0) && (cboStartDate.selectedIndex < 0)) {
			cboStartDate.selectedIndex = 0;
		}
		if ((cboStartSession.options.length > 0) && (cboStartSession.selectedIndex < 0)) {
			cboStartSession.selectedIndex = 0;
		}

		/* Event End Frame */
		if (fViewing) {
			radio_disable(optNoEnd, true);
			radio_disable(optEndDate, true);
			radio_disable(optDuration, true);
		}
		else {
			radio_disable(optNoEnd, (txtNoDateColumns.value == 1));
			radio_disable(optEndDate, (txtNoDateColumns.value == 1));
			radio_disable(optDuration, (txtNoDateColumns.value == 1));
		}

		if (optNoEnd.checked == true) {
			combo_disable(cboEndDate, true);
			cboEndDate.selectedIndex = -1;

			combo_disable(cboEndSession, true);
			cboEndSession.selectedIndex = -1;

			combo_disable(cboDuration, true);
			cboDuration.selectedIndex = -1;
		}
		else if (optEndDate.checked == true) {
			combo_disable(cboEndDate, fViewing);
			if ((cboEndDate.options.length > 0) && (cboEndDate.selectedIndex < 0)) {
				cboEndDate.selectedIndex = 0;
			}

			combo_disable(cboEndSession, fViewing);
			if ((cboEndSession.options.length > 0) && (cboEndSession.selectedIndex < 0)) {
				cboEndSession.selectedIndex = 0;
			}

			combo_disable(cboDuration, true);
			cboDuration.selectedIndex = -1;
		}
		else if (optDuration.checked == true) {
			combo_disable(cboEndDate, true);
			cboEndDate.selectedIndex = -1;

			combo_disable(cboEndSession, true);
			cboEndSession.selectedIndex = -1;

			combo_disable(cboDuration, fViewing);
			if ((cboDuration.options.length > 0) && (cboDuration.selectedIndex < 0)) {
				cboDuration.selectedIndex = 0;
			}
		}
		else {
			combo_disable(cboEndDate, true);
			cboEndDate.selectedIndex = -1;

			combo_disable(cboEndSession, true);
			cboEndSession.selectedIndex = -1;

			combo_disable(cboDuration, true);
			cboDuration.selectedIndex = -1;
		}

		/* Event Description Frame */
		if (fViewing) {
			combo_disable(cboEventDesc1, true);
			combo_disable(cboEventDesc2, true);
		}
		else {
			if (txtNoDateColumns.value == 1) {
				combo_disable(cboEventDesc1, true);
				combo_disable(cboEventDesc2, true);
			}
			else {
				combo_disable(cboEventDesc1, false);
				combo_disable(cboEventDesc2, false);
			}
		}

		if ((cboEventDesc1.options.length > 0) && (cboEventDesc1.selectedIndex < 0)) {
			cboEventDesc1.selectedIndex = 0;
		}
		if ((cboEventDesc2.options.length > 0) && (cboEventDesc2.selectedIndex < 0)) {
			cboEventDesc2.selectedIndex = 0;
		}
	}
}

function refreshLegendControls() {
	//var frmUse = OpenHR.getForm("workframe", "frmUseful");
	var frmUse = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");
	var fViewing = (frmUse.txtAction.value.toUpperCase() == "VIEW");
	var frmPopup = document.getElementById("frmPopup");

    with (frmPopup) {

        if ((cboEventType.options.length < 1)) {
            optLegendLookup.checked = false;
            optCharacter.checked = true;
        }

        if (fViewing) {
            radio_disable(optCharacter, true);
            radio_disable(optLegendLookup, true);
        }
        else {
            radio_disable(optCharacter, (txtNoDateColumns.value == 1));
            radio_disable(optLegendLookup, ((txtNoDateColumns.value == 1) || (cboEventType.options.length < 1)));
        }

        if (optCharacter.checked == true) {
            if (txtNoDateColumns.value == 1) {
                text_disable(txtCharacter, true);
            }
            else {
                text_disable(txtCharacter, fViewing);
            }

            combo_disable(cboEventType, true);
            cboEventType.selectedIndex = -1;

            combo_disable(cboLegendTable, true);
            cboLegendTable.selectedIndex = -1;

            combo_disable(cboLegendColumn, true);
            cboLegendColumn.selectedIndex = -1;

            combo_disable(cboLegendCode, true);
            cboLegendCode.selectedIndex = -1;
        }
        else {
            if (frmPopup.txtLookupColumnsLoaded.value == 0) {
                populateLookupColumns();
            }

            txtCharacter.value = '';
            text_disable(txtCharacter, true);

            if (cboEventType.options.length < 1) {
                combo_disable(cboEventType, true);
                cboEventType.selectedIndex = -1;

                combo_disable(cboLegendTable, true);
                cboLegendTable.selectedIndex = -1;

                combo_disable(cboLegendColumn, true);
                cboLegendColumn.selectedIndex = -1;

                combo_disable(cboLegendCode, true);
                cboLegendCode.selectedIndex = -1;
            }
            else {
                combo_disable(cboEventType, fViewing);
                if ((cboEventType.length > 0) && (cboEventType.selectedIndex < 0)) {
                    cboEventType.selectedIndex = 0;
                }

                combo_disable(cboLegendTable, fViewing);
                if ((cboLegendTable.length > 0) && (cboLegendTable.selectedIndex < 0)) {
                    cboLegendTable.selectedIndex = 0;
                }

                combo_disable(cboLegendColumn, fViewing);
                if ((cboLegendColumn.length > 0) && (cboLegendColumn.selectedIndex < 0)) {
                    cboLegendColumn.selectedIndex = 0;
                }

                if (cboLegendColumn.length < 1) {
                    combo_disable(cboLegendColumn, true);
                    cboLegendColumn.selectedIndex = -1;
                }

                combo_disable(cboLegendCode, fViewing);
                if ((cboLegendCode.length > 0) && (cboLegendCode.selectedIndex < 0)) {
                    cboLegendCode.selectedIndex = 0;
                }

                if (cboLegendCode.length < 1) {
                    combo_disable(cboLegendCode, true);
                    cboLegendCode.selectedIndex = -1;
                }
            }
        }
    }
}

function removeEventTable(piChildTableID) {
    var i;
    var iCount;
    var iTableID;
    var fChildColumnsSelected;
    var iIndex;
    var iCharIndex;
    var sControlName;
    var sDefn;

    var frmUse = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");
    var frmDefinition = OpenHR.getForm("workframe", "frmDefinition");
    var frmOriginalDefinition = OpenHR.getForm("workframe", "frmOriginalDefinition");

    frmUseful.txtCurrentChildTableID.value = piChildTableID;

    if (frmUseful.txtLoading.value == 'N') {
        if ((frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) ||
            ((frmUseful.txtAction.value.toUpperCase() != "NEW") &&
                (frmUseful.txtSelectedColumnsLoaded.value == 0))) {
            if (frmUseful.txtCurrentEventTableID.value != 0) {
                // Check if there are any child columns in the selected columns list.
                fChildColumnsSelected = false;
                if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
                    if (frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) {
                        frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
                        frmDefinition.ssOleDBGridSelectedColumns.movefirst();

                        for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
                            iTableID = frmDefinition.ssOleDBGridSelectedColumns.Columns("tableID").Text;

                            //if (window.dialogArguments.window.isSelectedChildTable(iTableID)) {
                            if (OpenHR.isSelectedChildTable(iTableID)) {
                                fChildColumnsSelected = true;
                                break;
                            }

                            if (iTableID == frmUseful.txtCurrentChildTableID.value) {
                                fChildColumnsSelected = true;
                                break;
                            }

                            frmDefinition.ssOleDBGridSelectedColumns.movenext();
                        }

                        frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
                        frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();
                        frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.add(frmDefinition.ssOleDBGridSelectedColumns.bookmark);
                    }
                }
                else {
                    var dataCollection = frmOriginalDefinition.elements;
                    if (dataCollection != null) {
                        for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                            sControlName = dataCollection.item(iIndex).name;
                            sControlName = sControlName.substr(0, 20);
                            if (sControlName == "txtReportDefnColumn_") {
                                //iTableID = document.parentWindow.parent.window.dialogArguments.window.selectedColumnParameter(dataCollection.item(iIndex).value, "TABLEID");
                                iTableID = OpenHR.selectedColumnParameter(dataCollection.item(iIndex).value, "TABLEID");
                                //if (document.parentWindow.parent.window.dialogArguments.window.isSelectedChildTable(iTableID))
                                if (OpenHR.isSelectedChildTable(iTableID)) {
                                    fChildColumnsSelected = true;
                                    break;
                                }

                                if (iTableID == frmUseful.txtCurrentChildTableID.value) {
                                    fChildColumnsSelected = true;
                                    break;
                                }

                            }
                        }
                    }
                }

                if (fChildColumnsSelected == true) {
                    var iAnswer = OpenHR.messageBox("One or more columns from the child table have been included in the report definition. Changing the child table will remove these columns from the report definition. Do you wish to continue ?", 36, "Calendar Reports");

                    if (iAnswer == 7) {
                        // cancel and change back !
                        return false;
                    }
                    else {
                        // Remove the child table's columns from the selected columns collection.
                        if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
                            if (frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) {
                                frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
                                frmDefinition.ssOleDBGridSelectedColumns.MoveFirst();

                                iCount = frmDefinition.ssOleDBGridSelectedColumns.rows;
                                for (i = 0; i < iCount; i++) {
                                    iTableID = frmDefinition.ssOleDBGridSelectedColumns.Columns("tableID").Text;
                                    if (iTableID == frmUseful.txtCurrentChildTableID.value) {
                                        if (frmDefinition.ssOleDBGridSelectedColumns.rows == 1) {
                                            frmDefinition.ssOleDBGridSelectedColumns.RemoveAll();
                                        }
                                        else {
                                            frmDefinition.ssOleDBGridSelectedColumns.RemoveItem(frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark));
                                        }
                                    }
                                    frmDefinition.ssOleDBGridSelectedColumns.MoveNext();
                                }

                                frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
                                frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();
                                frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.add(frmDefinition.ssOleDBGridSelectedColumns.bookmark);
                            }
                        }
                        else {
                            var dataCollection = frmOriginalDefinition.elements;
                            if (dataCollection != null) {
                                for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                                    sControlName = dataCollection.item(iIndex).name;
                                    sControlName = sControlName.substr(0, 20);
                                    if (sControlName == "txtReportDefnColumn_") {
                                        iTableID = OpenHR.selectedColumnParameter(dataCollection.item(iIndex).value, "TABLEID");
                                        if (iTableID == frmUseful.txtCurrentChildTableID.value) {
                                            dataCollection.item(iIndex).value = "";
                                        }
                                    }
                                }
                            }
                        }

                        // Remove the child table's columns from the sort order collection.
                        OpenHR.removeSortColumn(0, frmUseful.txtCurrentChildTableID.value);
                    }
                }
            }
        }
        frmUseful.txtChanged.value = 1;
    }

    OpenHR.refreshTab2Controls();
    frmUseful.txtTablesChanged.value = 1;
    //TM 24/07/02 Fault 4215
    frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();
    return true;
}



function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll) {
	var dlgwinprops = "center:yes;" +
			"dialogHeight:" + pHeight + "px;" +
			"dialogWidth:" + pWidth + "px;" +
			"help:no;" +
			"resizable:" + psResizable + ";" +
			"scroll:" + psScroll + ";" +
			"status:no;";
	//calendarreport.js at lineCap 5114
	showModalDialog(pDestination, self, dlgwinprops);
}

function AccessCode(psDescription) {
	if (psDescription == "Read / Write") {
		return "RW";
	}
	if (psDescription == "Read Only") {
		return "RO";
	}
	if (psDescription == "Hidden") {
		return "HD";
	}
	return "";
}

function AccessDescription(psCode) {
	if (psCode == "RW") {
		return "Read / Write";
	}
	if (psCode == "RO") {
		return "Read Only";
	}
	if (psCode == "HD") {
		return "Hidden";
	}
	return "Unknown";
}

function ForceAccess(pgrdAccess, psAccess) {
	var iLoop;
	var varBookmark;

	pgrdAccess.redraw = false;
	for (iLoop = 0; iLoop <= (pgrdAccess.Rows - 1) ; iLoop++) {
		varBookmark = pgrdAccess.AddItemBookmark(iLoop);
		pgrdAccess.Bookmark = varBookmark;

		if (iLoop == 0) {
			pgrdAccess.Columns("Access").Text = "";
		}
		else {
			if (pgrdAccess.Columns("SysSecMgr").CellText(varBookmark) != "1") {
				pgrdAccess.Columns("Access").Text = AccessDescription(psAccess);
			}
		}
	}
	pgrdAccess.redraw = true;

	pgrdAccess.MoveFirst();
}

function AllHiddenAccess(pgrdAccess) {
	var iLoop;
	var varBookmark;

	for (iLoop = 1; iLoop <= (pgrdAccess.Rows - 1) ; iLoop++) {
		varBookmark = pgrdAccess.AddItemBookmark(iLoop);

		if (pgrdAccess.Columns("SysSecMgr").CellText(varBookmark) != "1") {
			if (pgrdAccess.Columns("Access").CellText(varBookmark) != AccessDescription("HD")) {
				return (false);
			}
		}
	}

	return (true);
}

function HiddenGroups(pgrdAccess) {
	var iLoop;
	var varBookmark;
	var sHiddenGroups;

	sHiddenGroups = "";

	pgrdAccess.Update();
	for (iLoop = 1; iLoop <= (pgrdAccess.Rows - 1) ; iLoop++) {
		varBookmark = pgrdAccess.AddItemBookmark(iLoop);

		if (pgrdAccess.Columns("SysSecMgr").CellText(varBookmark) != "1") {
			if (pgrdAccess.Columns("Access").CellText(varBookmark) == AccessDescription("HD")) {
				sHiddenGroups = sHiddenGroups + pgrdAccess.Columns("GroupName").CellText(varBookmark) + "	";
			}
		}
	}

	if (sHiddenGroups.length > 0) {
		sHiddenGroups = "	" + sHiddenGroups;
	}

	return (sHiddenGroups);
}