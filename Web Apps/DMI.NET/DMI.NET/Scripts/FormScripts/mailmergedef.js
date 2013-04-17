
var sRepDefn;
var sSortDefn;

var frmDefinition = document.getElementById("frmDefinition");
var frmUseful = document.getElementById("frmUseful");
var frmOriginalDefinition = document.getElementById("frmOriginalDefinition");
var frmCustomReportChilds = document.getElementById("frmCustomReportChilds");
var frmSortOrder = document.getElementById("frmSortOrder");
var frmSelectionAccess = document.getElementById("frmSelectionAccess");
var frmSend = document.getElementById("frmSend");
var frmAccess = document.getElementById("frmAccess");
var frmValidate = document.getElementById("frmValidate");
var frmEmailSelection = document.getElementById("frmEmailSelection");
var frmTables = document.getElementById("frmTables");

var div1 = document.getElementById("div1");
var div2 = document.getElementById("div2");
var div3 = document.getElementById("div3");
var div4 = document.getElementById("div4");
var div5 = document.getElementById("div5");

function util_def_mailmerge_window_onload() {
	    var fOK;
	    fOK = true;
	    //debugger;
	    var sErrMsg = frmUseful.txtErrorDescription.value;
	    if (sErrMsg.length > 0) {
	        fOK = false;
	        OpenHR.messageBox(sErrMsg);
	        //TODO
	        //window.parent.location.replace("login");
	    }

	    setGridFont(frmDefinition.grdAccess);
	    setGridFont(frmDefinition.ssOleDBGridAvailableColumns);
	    setGridFont(frmDefinition.ssOleDBGridSelectedColumns);
	    setGridFont(frmDefinition.ssOleDBGridSortOrder);

	    if (fOK == true) {
	        // Expand the work frame and hide the option frame.
	        //window.parent.document.all.item("workframeset").cols = "*, 0";
	        $("#workframe").attr("data-framesource", "UTIL_DEF_MAILMERGE");

	        populateBaseTableCombo();

	        if (frmUseful.txtAction.value.toUpperCase() == "NEW") {
	            frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
	            setBaseTable(0);
	            changeBaseTable();
	            frmUseful.txtSelectedColumnsLoaded.value = 1;
	            frmUseful.txtSortLoaded.value = 1;
	            frmDefinition.txtDescription.value = "";
	            button_disable(frmDefinition.cmdTemplateClear, true);
	            frmDefinition.chkPause.checked = true;
	            frmDefinition.chkOutputScreen.checked = true;
	            frmDefinition.chkSuppressBlanks.checked = true;
	            frmDefinition.optDestination0.checked = true;
	            refreshDestination();
	            populatePrinters();
	            populateDMEngine();
	        }
	        else {
	            loadDefinition();
	        }

	        populateAccessGrid();

	        if (frmUseful.txtAction.value.toUpperCase() != "EDIT") {
	            frmUseful.txtUtilID.value = 0;
	        }

	        if (frmUseful.txtAction.value.toUpperCase() == "COPY") {
	            frmUseful.txtChanged.value = 1;
	        }

	        displayPage(1);

	        frmUseful.txtLoading.value = 'N';
	        try {
	            frmDefinition.txtName.focus();
	        } catch (e) {
	        }

	        // Get menu.asp to refresh the menu.
	        menu_refreshMenu();

	        //Check that the specified printer exists on this client machine.
	        if ((frmDefinition.optDestination0.checked == true) && (frmUseful.txtAction.value.toUpperCase() != "NEW")) {
	            if (frmOriginalDefinition.txtDefn_OutputPrinterName.value != "") {
	                if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.options.selectedIndex).text != frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
	                    OpenHR.messageBox("This definition is set to output to printer " + frmOriginalDefinition.txtDefn_OutputPrinterName.value + " which is not set up on your PC.");
	                    oOption = document.createElement("OPTION");
	                    frmDefinition.cboPrinterName.options.add(oOption);
	                    oOption.innerText = frmOriginalDefinition.txtDefn_OutputPrinterName.value;
	                    oOption.value = frmDefinition.cboPrinterName.options.length - 1;
	                    frmDefinition.cboPrinterName.selectedIndex = oOption.value;
	                }
	            }
	        }

	        //Check that the specified DMS Output Engine exists on this client machine.
	        if ((frmDefinition.optDestination2.checked == true) && (frmUseful.txtAction.value.toUpperCase() != "NEW")) {
	            if (frmOriginalDefinition.txtDefn_OutputPrinterName.value != "") {
	                if (frmDefinition.cboDMEngine.options(frmDefinition.cboDMEngine.options.selectedIndex).text != frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
	                    OpenHR.messageBox("This definition is set to output to printer " + frmOriginalDefinition.txtDefn_OutputPrinterName.value + " which is not set up on your PC.");
	                    var oOption = document.createElement("OPTION");
	                    frmDefinition.cboDMEngine.options.add(oOption);
	                    oOption.innerText = frmOriginalDefinition.txtDefn_OutputPrinterName.value;
	                    oOption.value = frmDefinition.cboDMEngine.options.length - 1;
	                    frmDefinition.cboDMEngine.selectedIndex = oOption.value;
	                }
	            }
	        }
	    }
	}
function TemplateClear() {

    frmDefinition.txtTemplate.value = "";
    button_disable(frmDefinition.cmdTemplateClear, true);

    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
}
function populateTableAvailable() {
    with (frmDefinition) {
        //Clear the existing data in the child table combo
    	while (frmDefinition.cboTblAvailable.options.length > 0) { frmDefinition.cboTblAvailable.options.remove(0); }

        //add the base table to the available tables list
    	var sTableID = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;
        var oOption = document.createElement("OPTION");
        frmDefinition.cboTblAvailable.options.add(oOption);
        oOption.innerText = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].innerText;
        oOption.value = sTableID;
        oOption.selected = true;

        //add the Parent 1 table to the available tables list (if it exists)
        if (frmDefinition.txtParent1ID.value > 0) {
            sTableID = frmDefinition.txtParent1ID.value;
            oOption = document.createElement("OPTION");
            frmDefinition.cboTblAvailable.options.add(oOption);
            oOption.innerText = frmDefinition.txtParent1.value;
            oOption.value = sTableID;
        }

        //add the Parent 2 table to the available tables list (if it exists)
        if (frmDefinition.txtParent2ID.value > 0) {
            sTableID = frmDefinition.txtParent2ID.value;
            oOption = document.createElement("OPTION");
            frmDefinition.cboTblAvailable.options.add(oOption);
            oOption.innerText = frmDefinition.txtParent2.value;
            oOption.value = sTableID;
        }
    }
}
function refreshAvailableColumns() {
    if (frmUseful.txtLoading.value == 'N') {
        loadAvailableColumns();
    }
}
function TemplateSelect() {
	var sPath = "";
	if (frmDefinition.txtTemplate.value.length == 0) {
		var sKey = new String("documentspath_");
		sKey = sKey.concat(OpenHR.getForm("menuframe", "frmMenuInfo").txtDatabase.value);
		sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
		dialog.InitDir = sPath;
	}
	else {
		dialog.FileName = frmDefinition.txtTemplate.value;
		//sPath = new String(frmDefinition.txtTemplate.value);
	}

	dialog.CancelError = true;
	dialog.DialogTitle = "Mail Merge Template";
	dialog.Filter = "Word Template (*.dot;*.dotx;*.doc;*.docx)|*.dot;*.dotx;*.doc;*.docx";
	dialog.Flags = 2621444;

	try {
		dialog.ShowOpen();
	}
	catch (e) {
	}
	if (dialog.FileName != "") {
		//if (sPath.length > 0) {
		if (OpenHR.ValidateFilePath(sPath) == false) {
			{
				var iResponse = OpenHR.messageBox("Template file does not exist.  Create it now?", 36);
				if (iResponse == 6) {
					frmDefinition.txtTemplate.value = dialog.FileName;
					button_disable(frmDefinition.cmdTemplateClear, false);

					try {
						var sOfficeSaveAsValues = '<%=session("OfficeSaveAsValues")%>';
						OpenHR.SaveAsValues = sOfficeSaveAsValues;
						MM_WORD_CreateTemplateFile(dialog.FileName);
					} catch (e) {
					}
				}
			}
		} else {
			frmDefinition.txtTemplate.value = dialog.FileName;
			button_disable(frmDefinition.cmdTemplateClear, false);
		}
	}

	frmUseful.txtChanged.value = 1;
	refreshTab4Controls();
}

function displayPage(piPageNumber) {
	//window.parent.frames("refreshframe").document.forms("frmRefresh").submit();
	OpenHR.submitForm(window.frmRefresh);
	
    if (piPageNumber == 1) {
        div1.style.visibility = "visible";
        div1.style.display = "block";
        div2.style.visibility = "hidden";
        div2.style.display = "none";
        div3.style.visibility = "hidden";
        div3.style.display = "none";
        div4.style.visibility = "hidden";
        div4.style.display = "none";

        button_disable(frmDefinition.btnTab1, true);
        button_disable(frmDefinition.btnTab2, false);
        button_disable(frmDefinition.btnTab3, false);
        button_disable(frmDefinition.btnTab4, false);

    	try {
            frmDefinition.txtName.focus();
        }
        catch (e) { }


        refreshTab1Controls();
    }

    if (piPageNumber == 2) {
        // Get the columns/calcs for the current tvable selection.
    	var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
    	
        if (frmUseful.txtTablesChanged.value == 1) {
            frmGetDataForm.txtAction.value = "LOADREPORTCOLUMNS";
            frmGetDataForm.txtReportBaseTableID.value = frmUseful.txtCurrentBaseTableID.value;
            frmGetDataForm.txtReportParent1TableID.value = frmDefinition.txtParent1ID.value;
            frmGetDataForm.txtReportParent2TableID.value = frmDefinition.txtParent2ID.value;
            frmGetDataForm.txtReportChildTableID.value = 0;
            window.data_refreshData();

            frmUseful.txtTablesChanged.value = 0;
        }

        div1.style.visibility = "hidden";
        div1.style.display = "none";
        div2.style.visibility = "visible";
        div2.style.display = "block";
        div3.style.visibility = "hidden";
        div3.style.display = "none";
        div4.style.visibility = "hidden";
        div4.style.display = "none";
        button_disable(frmDefinition.btnTab1, false);
        button_disable(frmDefinition.btnTab2, true);
        button_disable(frmDefinition.btnTab3, false);
        button_disable(frmDefinition.btnTab4, false);

        loadSelectedColumnsDefinition();
        CheckExpressionTypes();
        frmDefinition.ssOleDBGridAvailableColumns.focus();

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
        button_disable(frmDefinition.btnTab1, false);
        button_disable(frmDefinition.btnTab2, false);
        button_disable(frmDefinition.btnTab3, true);
        button_disable(frmDefinition.btnTab4, false);

        frmDefinition.ssOleDBGridSortOrder.focus();
        loadSortDefinition();

        frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();
        if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
            frmDefinition.ssOleDBGridSortOrder.MoveFirst();
            frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
        }

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
        button_disable(frmDefinition.btnTab1, false);
        button_disable(frmDefinition.btnTab2, false);
        button_disable(frmDefinition.btnTab3, false);
        button_disable(frmDefinition.btnTab4, true);

        refreshTab4Controls();
    }

    // Little dodge to get around a browser bug that
    // does not refresh the display on all controls.
    try {
        window.resizeBy(0, -1);
        window.resizeBy(0, 1);
    }
    catch (e) { }
}
function populateBaseTableCombo() {
    var i;

    //Clear the existing data in the child table combo
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

    populateTableAvailable();
}
function changeBaseTable() {
    var i;

    if (frmUseful.txtLoading.value == 'N') {
        if ((frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) ||
            ((frmUseful.txtAction.value.toUpperCase() != "NEW") &&
                (frmUseful.txtSelectedColumnsLoaded.value == 0))) {
            var iAnswer = OpenHR.messageBox("Warning: Changing the base table will result in all table/column specific aspects of this mail merge definition being cleared. Are you sure you wish to continue?", 36);
            if (iAnswer == 7) {
                // cancel and change back ! (txtcurrentbasetable)
                setBaseTable(frmUseful.txtCurrentBaseTableID.value);
                return;
            }
            else {
                // clear cols and sort order
                if (frmUseful.txtSelectedColumnsLoaded.value != 0) {
                    frmDefinition.ssOleDBGridSelectedColumns.RemoveAll();
                }
                if (frmUseful.txtSortLoaded.value != 0) {
                    // JPD20020718 Fault 4193
                    if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
                        frmDefinition.ssOleDBGridSortOrder.RemoveAll();
                    }
                }
                frmSelectionAccess.calcsHiddenCount.value = 0;
                frmUseful.txtSelectedColumnsLoaded.value = 1;
                frmUseful.txtSortLoaded.value = 1;
                frmUseful.txtChanged.value = 1;
            }
        } else {
            frmUseful.txtChanged.value = 1;
        }
    }

    clearBaseTableRecordOptions();

    //Empty the parent textboxes
    frmDefinition.txtParent1.value = '';
    frmDefinition.txtParent1ID.value = 0;
    frmDefinition.txtParent2.value = '';
    frmDefinition.txtParent2ID.value = 0;
    var sParents = new String("");
    var dataCollection = frmTables.elements;
    if (dataCollection != null) {
        var sReqdControlName = new String("txtTableParents_");
        sReqdControlName = sReqdControlName.concat(frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);

        for (i = 0; i < dataCollection.length; i++) {
            var sControlName = dataCollection.item(i).name;
            if (sControlName == sReqdControlName) {
                sParents = dataCollection.item(i).value;
                break;
            }
        }
    }

    var iIndex = sParents.indexOf("	");
    if (iIndex > 0) {
        var sParent1ID = sParents.substr(0, iIndex);
        frmDefinition.txtParent1.value = getTableName(sParent1ID);
        frmDefinition.txtParent1ID.value = sParent1ID;
        sParents = sParents.substr(iIndex + 1);
    }
    iIndex = sParents.indexOf("	");
    if (iIndex > 0) {
        var sParent2ID = sParents.substr(0, iIndex);
        frmDefinition.txtParent2.value = getTableName(sParent2ID);
        frmDefinition.txtParent2ID.value = sParent2ID;
    }

    refreshTab1Controls();
    frmUseful.txtCurrentBaseTableID.value = frmDefinition.cboBaseTable.options(frmDefinition.cboBaseTable.options.selectedIndex).value;
    frmUseful.txtTablesChanged.value = 1;

    refreshDestination();
    populateTableAvailable();
}
function refreshTab1Controls() {
    var fIsForcedHidden;
    var fViewing;
    var fIsNotOwner;
    var fAllAlreadyHidden;
    var fSilent;

    fSilent = ((frmUseful.txtAction.value.toUpperCase() == "COPY") &&
        (frmUseful.txtLoading.value == "Y"));

    fIsForcedHidden = ((frmSelectionAccess.baseHidden.value == "Y") ||
        (frmSelectionAccess.p1Hidden.value == "Y") ||
        (frmSelectionAccess.p2Hidden.value == "Y") ||
        (frmSelectionAccess.childHidden.value == "Y") ||
        (frmSelectionAccess.calcsHiddenCount.value > 0));
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());
    fAllAlreadyHidden = AllHiddenAccessMM(frmDefinition.grdAccess);
    
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
                //MH20040816 Fault 9050
                //if (fSilent == false) {
                if ((fSilent == false) && (frmUseful.txtLoading.value != "Y")) {
                    OpenHR.messageBox("The definition access cannot be changed as it contains a hidden picklist/filter/calculation.", 64);
                }
            }
        }
        frmSelectionAccess.forcedHidden.value = "Y";
    }
    else {
        if (frmSelectionAccess.forcedHidden.value == "Y") {
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

    button_disable(frmDefinition.cmdBasePicklist, ((frmDefinition.optRecordSelection2.checked == false)
        || (fViewing == true)));
    button_disable(frmDefinition.cmdBaseFilter, ((frmDefinition.optRecordSelection3.checked == false)
        || (fViewing == true)));

    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
        (fViewing == true)));

    // Little dodge to get around a browser bug that
    // does not refresh the display on all controls.
    try {
        window.resizeBy(0, -1);
        window.resizeBy(0, 1);
    }
    catch (e) { }
}
function refreshTab2Controls() {
    var fAddDisabled;
    var fAddAllDisabled;
    var fRemoveDisabled;
    var fRemoveAllDisabled;
    var fMoveUpDisabled;
    var fMoveDownDisabled;

    var fTableColDisabled;

    var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

    fTableColDisabled = (fViewing == true);

    fAddDisabled = ((frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.Count == 0)
        || (fViewing == true));
    fAddAllDisabled = ((frmDefinition.ssOleDBGridAvailableColumns.Rows == 0)
        || (fViewing == true));
    fRemoveDisabled = ((frmDefinition.ssOleDBGridSelectedColumns.SelBookmarks.Count == 0)
        || (fViewing == true));
    fRemoveAllDisabled = ((frmDefinition.ssOleDBGridSelectedColumns.Rows == 0)
        || (fViewing == true));
    fMoveUpDisabled = fViewing;
    fMoveDownDisabled = fViewing;

    if ((fRemoveDisabled == true) || (frmDefinition.ssOleDBGridSelectedColumns.SelBookmarks.Count != 1)) {
        fMoveUpDisabled = true;
        fMoveDownDisabled = true;
    }
    else {
        // Are we on the top row ?
        if (frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark) == 0) {
            fMoveUpDisabled = true;
        }

        // Are we on the bottom row ?
        if (frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark) == frmDefinition.ssOleDBGridSelectedColumns.rows - 1) {
            fMoveDownDisabled = true;
        }
    }

    combo_disable(frmDefinition.cboTblAvailable, fTableColDisabled);
    radio_disable(frmDefinition.optCalc, fTableColDisabled);
    radio_disable(frmDefinition.optColumns, fTableColDisabled);
    button_disable(frmDefinition.cmdColumnAdd, fAddDisabled);
    button_disable(frmDefinition.cmdColumnAddAll, fAddAllDisabled);
    button_disable(frmDefinition.cmdColumnRemove, fRemoveDisabled);
    button_disable(frmDefinition.cmdColumnRemoveAll, fRemoveAllDisabled);

    var fSizeDisabled = true;
    var sSize = "";
    var fDecPlacesDisabled = true;
    var sDecPlaces = "";

    if (frmDefinition.ssOleDBGridSelectedColumns.SelBookmarks.Count == 1) {
        fSizeDisabled = fViewing;
        sSize = frmDefinition.ssOleDBGridSelectedColumns.Columns(4).text;

        if (frmDefinition.ssOleDBGridSelectedColumns.columns(7).text == '1') {
            // The column is numeric.
            fDecPlacesDisabled = fViewing;
            sDecPlaces = frmDefinition.ssOleDBGridSelectedColumns.Columns(5).text;
        }
    }

    text_disable(frmDefinition.txtSize, fSizeDisabled);
    frmDefinition.txtSize.value = sSize;
    text_disable(frmDefinition.txtDecPlaces, fDecPlacesDisabled);
    frmDefinition.txtDecPlaces.value = sDecPlaces;

    frmDefinition.ssOleDBGridAvailableColumns.RowHeight = 19;
    frmDefinition.ssOleDBGridSelectedColumns.RowHeight = 19;

    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
        (fViewing == true)));
}
function refreshTab3Controls() {
    var i;
    var iCount;

    var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

    var fSortAddDisabled = fViewing;
    var fSortEditDisabled = fViewing;
    var fSortRemoveDisabled = fViewing;
    var fSortRemoveAllDisabled = fViewing;
    var fSortMoveUpDisabled = fViewing;
    var fSortMoveDownDisabled = fViewing;

    if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
        if (frmDefinition.ssOleDBGridSelectedColumns.Rows <= frmDefinition.ssOleDBGridSortOrder.Rows) {
            // Disable 'Add' if there are no more columns to sort by.
            fSortAddDisabled = true;
        }
    }
    else {
        iCount = 0;
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (i = 0; i < dataCollection.length; i++) {
                var sControlName = dataCollection.item(i).name;
                sControlName = sControlName.substr(0, 20);
                if (sControlName == "txtReportDefnColumn_") {
                    iCount = iCount + 1;
                }
            }
        }

        if (iCount <= frmDefinition.ssOleDBGridSortOrder.Rows) {
            // Disable 'Add' if there are no more columns to sort by.
            fSortAddDisabled = true;
        }
    }

    //  if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count == 0) {
    if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count < 1) {
        fSortRemoveDisabled = true;
    }

    if (frmDefinition.ssOleDBGridSortOrder.rows <= 0) {
        fSortRemoveAllDisabled = true;
    }

    if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count != 1) {
        fSortEditDisabled = true;
        fSortMoveDownDisabled = true;
        fSortMoveUpDisabled = true;
    }
    else {
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

    button_disable(frmDefinition.cmdSortAdd, fSortAddDisabled);
    button_disable(frmDefinition.cmdSortEdit, fSortEditDisabled);
    button_disable(frmDefinition.cmdSortRemove, fSortRemoveDisabled);
    button_disable(frmDefinition.cmdSortRemoveAll, fSortRemoveAllDisabled);
    button_disable(frmDefinition.cmdSortMoveUp, fSortMoveUpDisabled);
    button_disable(frmDefinition.cmdSortMoveDown, fSortMoveDownDisabled);

    // frmDefinition.ssOleDBGridSortOrder.AllowUpdate = (fViewing == false);	
    frmDefinition.ssOleDBGridSortOrder.AllowUpdate = false;

    frmDefinition.ssOleDBGridSortOrder.RowHeight = 19;

    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
        (fViewing == true)));
}
function refreshTab4Controls() {
	debugger;
    var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) || (fViewing == true)));
    // Little dodge to get around a browser bug that
	// does not refresh the display on all controls.
	try {
		window.resizeBy(0, -1);
		window.resizeBy(0, 1);
	}
	catch (e) { }

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
function selectRecordOption(psTable, psType) {
    var sURL;
    var iTableID;
    var iCurrentID;
    if (psTable == 'base') {
        iTableID = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;

        if (psType == 'picklist') {
            iCurrentID = frmDefinition.txtBasePicklistID.value;
        }
        else {
            iCurrentID = frmDefinition.txtBaseFilterID.value;
        }
    }

    window.frmRecordSelection.recSelTable.value = psTable;
    window.frmRecordSelection.recSelType.value = psType;
    window.frmRecordSelection.recSelTableID.value = iTableID;
    window.frmRecordSelection.recSelCurrentID.value = iCurrentID;

    var strDefOwner = new String(frmDefinition.txtOwner.value);
    var strCurrentUser = new String(frmUseful.txtUserName.value);

    strDefOwner = strDefOwner.toLowerCase();
    strCurrentUser = strCurrentUser.toLowerCase();

    if (strDefOwner == strCurrentUser) {
        window.frmRecordSelection.recSelDefOwner.value = '1';
    }
    else {
        window.frmRecordSelection.recSelDefOwner.value = '0';
    }
    window.frmRecordSelection.recSelDefType.value = "Mail Merge";

    sURL = "util_recordSelection" +
        "?recSelType=" + escape(window.frmRecordSelection.recSelType.value) +
        "&recSelTableID=" + escape(window.frmRecordSelection.recSelTableID.value) +
        "&recSelCurrentID=" + escape(window.frmRecordSelection.recSelCurrentID.value) +
        "&recSelTable=" + escape(window.frmRecordSelection.recSelTable.value) +
        "&recSelDefOwner=" + escape(window.frmRecordSelection.recSelDefOwner.value) +
        "&recSelDefType=" + escape(window.frmRecordSelection.recSelDefType.value);
    openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");

    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}
function setRecordsNumeric() {
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

    if (frmDefinition.txtChildRecords.value == '') {
        frmDefinition.txtChildRecords.value = 0;
    }

    // Convert the value from locale to UK settings for use with the isNaN funtion.
    sConvertedValue = new String(frmDefinition.txtChildRecords.value);
    // Remove any thousand separators.
    sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
    frmDefinition.txtChildRecords.value = sConvertedValue;

    // Convert any decimal separators to '.'.
    if (OpenHR.LocaleDecimalSeparator != ".") {
        // Remove decimal points.
        sConvertedValue = sConvertedValue.replace(rePoint, "A");
        // replace the locale decimal marker with the decimal point.
        sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
    }

    if (isNaN(sConvertedValue) == true) {
        OpenHR.messageBox("No. of records must be numeric.");
        frmDefinition.txtChildRecords.value = 0;
    }
    else {
        if (sConvertedValue.indexOf(".") >= 0) {
            OpenHR.messageBox("Invalid integer value.");
            frmDefinition.txtChildRecords.value = 0;
        }
        else {
            if (frmDefinition.txtChildRecords.value < 0) {
                OpenHR.messageBox("The value cannot be negative.");
                frmDefinition.txtChildRecords.value = 0;
            }
        }
    }

    refreshTab2Controls();
}
function validateTab2() {
    var  i;
    var iCount;
    var sAllHeadings;
    var sType;
    var sHidden;
    var sErrMsg;
    var sCurrentHeading;
    var sDefn;
    var sControlName;

    sErrMsg = "";
    sAllHeadings = "";

    // Check report columns have been selected
    // Check all cols have a (unique) heading
    // Any hidden calcs included?
    if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
        if (frmDefinition.ssOleDBGridSelectedColumns.Rows == 0) {
            sErrMsg = "No merge columns selected.";
        } else {
            frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
            frmDefinition.ssOleDBGridSelectedColumns.movefirst();

            for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
                //CurrentHeading = frmDefinition.ssOleDBGridSelectedColumns.Columns("heading").Text;

                //if (sCurrentHeading == '') {
                //	sErrMsg = "All columns must have a heading.";
                //	break;
                //}

                //if (sAllHeadings.indexOf('	' + sCurrentHeading + '	') != -1) {
                //	sErrMsg = "All column headings must be unique.";
                //	break;
                //}
                //else {
                sAllHeadings = sAllHeadings + '	' + sCurrentHeading + '	';

                if ((frmDefinition.ssOleDBGridSelectedColumns.columns("type").text == 'E') &&
                    (frmDefinition.ssOleDBGridSelectedColumns.columns("hidden").text == 'Y')) {
                    if (frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) {
                        sErrMsg = "You have selected a hidden calculation but you are not the owner of the definition.";
                        break;
                    }
                    //}
                }

                frmDefinition.ssOleDBGridSelectedColumns.movenext();
            }

            frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
            frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();
            frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.add(frmDefinition.ssOleDBGridSelectedColumns.bookmark);
        }
    } else {
        iCount = 0;
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (i = 0; i < dataCollection.length; i++) {
                sControlName = dataCollection.item(i).name;
                sControlName = sControlName.substr(0, 20);
                if (sControlName == "txtReportDefnColumn_") {

                    sDefn = new String(dataCollection.item(i).value);
                    //sCurrentHeading = selectedColumnParameter(sDefn, "HEADING");					
                    sType = selectedColumnParameter(sDefn, "TYPE");
                    sHidden = selectedColumnParameter(sDefn, "HIDDEN");

                    iCount = iCount + 1;

                    if ((sType == 'E') && (sHidden == 'Y')) {
                        if (frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) {
                            sErrMsg = "You have selected a hidden calculation but you are not the owner of the definition.";
                            break;
                        }
                    }
                }
            }
        }

        if (iCount == 0) {
            sErrMsg = "No merge columns selected.";
        }
    }

    if (sErrMsg.length > 0) {
        OpenHR.messageBox(sErrMsg, 48);
        displayPage(2);
        return (false);
    }

    return (true);
}
function submitDefinition() {
    var i;
    var iIndex;
    var sColumnID;
    var sType;

    if (validateTab1() == false) { OpenHR.refreshMenu(); return; }
    if (validateTab2() == false) { OpenHR.refreshMenu(); return; }
    if (validateTab3() == false) { OpenHR.refreshMenu(); return; }
    if (validateTab4() == false) { OpenHR.refreshMenu(); return; }
    if (populateSendForm() == false) { OpenHR.refreshMenu(); return; }

    // Now create the validate popup to check that any filters/calcs
    // etc havent been deleted, or made hidden etc.		

    // first populate the validate fields
    frmValidate.validateBaseFilter.value = frmDefinition.txtBaseFilterID.value;
    frmValidate.validateBasePicklist.value = frmDefinition.txtBasePicklistID.value;
    //frmValidate.validateP1Filter.value = frmDefinition.txtParent1FilterID.value;
    //frmValidate.validateP1Picklist.value = frmDefinition.txtParent1PicklistID.value;
    //frmValidate.validateP2Filter.value = frmDefinition.txtParent2FilterID.value;
    //frmValidate.validateP2Picklist.value = frmDefinition.txtParent2PicklistID.value;
    //frmValidate.validateChildFilter.value = frmDefinition.txtChildFilterID.value;		
    frmValidate.validateName.value = frmDefinition.txtName.value;

    if (frmUseful.txtAction.value.toUpperCase() == "EDIT") {
        frmValidate.validateTimestamp.value = frmOriginalDefinition.txtDefn_Timestamp.value;
        frmValidate.validateUtilID.value = frmUseful.txtUtilID.value;
    }
    else {
        frmValidate.validateTimestamp.value = 0;
        frmValidate.validateUtilID.value = 0;
    }
    frmValidate.validateCalcs.value = '';

    if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
        frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
        frmDefinition.ssOleDBGridSelectedColumns.movefirst();

        for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
            if (frmDefinition.ssOleDBGridSelectedColumns.Columns("type").Text == 'E') {
                if (frmValidate.validateCalcs.value != '') {
                    frmValidate.validateCalcs.value = frmValidate.validateCalcs.value + ',';
                }
                frmValidate.validateCalcs.value = frmValidate.validateCalcs.value + frmDefinition.ssOleDBGridSelectedColumns.Columns("columnid").Text;
            }
            frmDefinition.ssOleDBGridSelectedColumns.movenext();
        }

        frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
    }
    else {
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                var sControlName = dataCollection.item(iIndex).name;
                sControlName = sControlName.substr(0, 20);
                if (sControlName == "txtReportDefnColumn_") {
                    var sDefnString = new String(dataCollection.item(iIndex).value);
                    if (sDefnString.length > 0) {
                        sType = selectedColumnParameter(sDefnString, "TYPE");
                        sColumnID = selectedColumnParameter(sDefnString, "COLUMNID");

                        if (sType == 'E') {
                            if (frmValidate.validateCalcs.value != '') {
                                frmValidate.validateCalcs.value = frmValidate.validateCalcs.value + ',';
                            }
                            frmValidate.validateCalcs.value = frmValidate.validateCalcs.value + sColumnID;
                        }
                    }
                }
            }
        }
    }

    var sHiddenGroups = HiddenGroups(frmDefinition.grdAccess);
    frmValidate.validateHiddenGroups.value = sHiddenGroups;

    var sURL = "util_validate_mailmerge" +
        "?validateBaseFilter=" + escape(frmValidate.validateBaseFilter.value) +
        "&validateBasePicklist=" + escape(frmValidate.validateBasePicklist.value) +
        "&validateCalcs=" + escape(frmValidate.validateCalcs.value) +
        "&validateHiddenGroups=" + escape(frmValidate.validateHiddenGroups.value) +
        "&validateName=" + escape(frmValidate.validateName.value) +
        "&validateTimestamp=" + escape(frmValidate.validateTimestamp.value) +
        "&validateUtilID=" + escape(frmValidate.validateUtilID.value) +
        "&destination=util_validate_mailmerge";
    //openDialog(sURL, (screen.width) / 2, (screen.height) / 3, "no", "no");
	openDialog(sURL, (screen.width) / 2, (screen.height) / 3);
}


function cancelClick() {
		if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
			(definitionChanged() == false)) {

			menu_loadDefSelPage(9, frmUseful.txtUtilID.value, frmUseful.txtCurrentBaseTableID.value, false);
			return (false);
		}

		var answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3);
		if (answer == 7) {
			// No
			menu_loadDefSelPage(9, frmUseful.txtUtilID.value, frmUseful.txtCurrentBaseTableID.value, false);
			
			return (false);
		}
		if (answer == 6) {
			// Yes
			okClick();
		}
	}

function okClick() {
	menu_disableMenu();
	frmSend.txtSend_reaction.value = "MAILMERGE";
	submitDefinition();
}

function saveChanges(psAction, pfPrompt, pfTBOverride) {
    if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
        (definitionChanged() == false)) {
        return 7; //No to saving the changes, as none have been made.
    }
    var answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3);
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

            // Compare the tab 4 controls with the original values.
            if (frmDefinition.txtTemplate.value != frmOriginalDefinition.txtDefn_TemplateFileName.value) {
                return true;
            }

            if (frmDefinition.chkSave.checked.toString() != frmOriginalDefinition.txtDefn_OutputSave.value) {
                return true;
            }

            if (frmDefinition.txtSaveFile.value != frmOriginalDefinition.txtDefn_OutputFileName.value) {
                return true;
            }

            if (frmDefinition.cboEmail.options.selectedIndex != -1) {
                if (frmDefinition.cboEmail.options(frmDefinition.cboEmail.options.selectedIndex).value != frmOriginalDefinition.txtDefn_EmailAddrID.value) {
                    return true;
                }
            }

            if (frmDefinition.txtSubject.value != frmOriginalDefinition.txtDefn_EmailSubject.value) {
                return true;
            }

            if ((frmDefinition.optDestination1.checked != true) && (frmDefinition.chkOutputScreen.checked.toString() != frmOriginalDefinition.txtDefn_OutputScreen.value)) {
                return true;
            }
            if (frmDefinition.chkAttachment.checked.toString() != frmOriginalDefinition.txtDefn_EmailAsAttachment.value) {
                return true;
            }

            if (frmDefinition.txtAttachmentName.value != frmOriginalDefinition.txtDefn_EmailAttachmentName.value) {
                return true;
            }

            if (frmDefinition.chkSuppressBlanks.checked.toString() != frmOriginalDefinition.txtDefn_SuppressBlanks.value) {
                return true;
            }
            if (frmDefinition.chkPause.checked.toString() != frmOriginalDefinition.txtDefn_PauseBeforeMerge.value) {
                return true;
            }

            if ((frmDefinition.optDestination0.checked == true) && (frmDefinition.cboPrinterName.options.selectedIndex != -1)) {
                if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.options.selectedIndex).text != frmOriginalDefinition.txtDefn_OutputPrinterName.value) {

                    if ((frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.options.selectedIndex).text == "<Default Printer>") &&
                        (frmOriginalDefinition.txtDefn_OutputPrinterName.value.length == 0)) {
                        //<Default Printer> is stored as "", so no change.
                        return false;
                    }
                    else {
                        return true;
                    }
                }
            }

            if ((frmDefinition.optDestination2.checked == true) && (frmDefinition.cboDMEngine.options.selectedIndex != -1)) {
                if (frmDefinition.cboDMEngine.options(frmDefinition.cboDMEngine.options.selectedIndex).text != frmOriginalDefinition.txtDefn_OutputPrinterName.value) {

                    if ((frmDefinition.cboDMEngine.options(frmDefinition.cboDMEngine.options.selectedIndex).text == "<Default Printer>") &&
                        (frmOriginalDefinition.txtDefn_OutputPrinterName.value.length == 0)) {
                        //<Default Printer> is stored as "", so no change.
                        return false;
                    }
                    else {
                        return true;
                    }

                }
            }


        }
        return false;

    }
}
function spinRecords(pfUp) {
    var iRecords = frmDefinition.txtChildRecords.value;
    if (pfUp == true) {
        iRecords = ++iRecords;
    }
    else {
        if (iRecords > 0) {
            iRecords = iRecords - 1;
        }
    }

    frmDefinition.txtChildRecords.value = iRecords;

    refreshTab2Controls();
}
function getTableName(piTableID) {
    var i;
    var sTableName = new String("");

    var sReqdControlName = new String("txtTableName_");
    sReqdControlName = sReqdControlName.concat(piTableID);

    var dataCollection = frmTables.elements;
    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {
            var sControlName = dataCollection.item(i).name;

            if (sControlName == sReqdControlName) {
                sTableName = dataCollection.item(i).value;
                break;
            }
        }
    }

    return sTableName;
}
function columnSwap(pfSelect) {
    var i;
    var iColumnsSwapped;
    var sAddedCalcIDs;

    sAddedCalcIDs = "";
    iColumnsSwapped = 0;

    // Do nothing of the Add button is disabled (read-only mode).
    if (frmUseful.txtAction.value.toUpperCase() == "VIEW") return;
    var grdFrom;
    var grdTo;
    var iCount;
    var iRowIndex;
    var i2;
    if (pfSelect == true) {
        grdFrom = frmDefinition.ssOleDBGridAvailableColumns;
        grdTo = frmDefinition.ssOleDBGridSelectedColumns;
    }
    else {
        grdFrom = frmDefinition.ssOleDBGridSelectedColumns;
        grdTo = frmDefinition.ssOleDBGridAvailableColumns; // Check if the column being removed is in the sort columns collection.
        iCount = grdFrom.selbookmarks.Count();
        for (i = iCount - 1; i >= 0; i--) {
            grdFrom.bookmark = grdFrom.selbookmarks(i);
            iRowIndex = grdFrom.AddItemRowIndex(grdFrom.Bookmark);

            // Remove the column from the sort columns collection.
            if (grdFrom.columns(0).text == "C") {
                var sColumnName;
                var iResponse;
                if (frmUseful.txtSortLoaded.value == 1) {
                    if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
                        frmDefinition.ssOleDBGridSortOrder.Redraw = false;
                        frmDefinition.ssOleDBGridSortOrder.MoveFirst();

                        var iCount2 = frmDefinition.ssOleDBGridSortOrder.rows;
                        for (i2 = 0; i2 < iCount2; i2++) {
                            if (grdFrom.columns(2).text == frmDefinition.ssOleDBGridSortOrder.Columns("id").Text) {
                                // The selected column is a sort column. Prompt the user to confirm the deselection.

                                sColumnName = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
                                if (iCount > 1) {
                                    iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the definition sort order.\n\nDo you still want to remove this column ?", 3, "Mail Merge");
                                }
                                else {
                                    iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the definition sort order.\n\nDo you still want to remove this column ?", 4, "Mail Merge");
                                }

                                if (iResponse == 2) {
                                    // Cancel.
                                    frmDefinition.ssOleDBGridSortOrder.Redraw = true;
                                    return;
                                }

                                if (iResponse == 7) {
                                    // No. 
                                    grdFrom.selbookmarks.remove(i);
                                }

                                break;
                            }
                            else {
                                frmDefinition.ssOleDBGridSortOrder.MoveNext();
                            }
                        }

                        frmDefinition.ssOleDBGridSortOrder.Redraw = true;
                    }
                }
                else {
                    var dataCollection = frmOriginalDefinition.elements;
                    if (dataCollection != null) {
                        for (var iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                            var sControlName = dataCollection.item(iIndex).name;
                            sControlName = sControlName.substr(0, 19);
                            if (sControlName == "txtReportDefnOrder_") {
                                if (grdFrom.columns(2).text == sortColumnParameter(dataCollection.item(iIndex).value, "COLUMNID")) {
                                    // The selected column is a sort column. Prompt the user to confirm the deselection.
                                    sColumnName = grdFrom.columns(3).text;
                                    if (iCount > 1) {
                                        iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the definition sort order.\n\nDo you still want to remove this column ?", 3, "Mail Merge");
                                    }
                                    else {
                                        iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the definition sort order.\n\nDo you still want to remove this column ?", 4, "Mail Merge");
                                    }

                                    if (iResponse == 2) {
                                        // Cancel.
                                        frmDefinition.ssOleDBGridSortOrder.Redraw = true;
                                        return;
                                    }

                                    if (iResponse == 7) {
                                        // No. 
                                        grdFrom.selbookmarks.remove(i);
                                    }

                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    grdFrom.Redraw = false;
    grdTo.Redraw = false;

    grdTo.selbookmarks.removeall();

    if (grdFrom.SelBookmarks.count() > 0) {
        var iHiddenCalcCount = 0;

        for (i = 0; i < grdFrom.selbookmarks.Count() ; i++) {
            grdFrom.bookmark = grdFrom.selbookmarks(i);

            // Check if the user is selecting a hidden calc, but is not the report owner.
            if ((pfSelect == true) &&
                (frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) &&
                (grdFrom.columns(6).text == "Y")) {

                var sCalcName = new String(grdFrom.columns(3).text);
                var iStringIndex = sCalcName.indexOf("<Calc> ");
                if (iStringIndex >= 0) {
                    sCalcName = sCalcName.substring(iStringIndex + 7, sCalcName.length);
                }
                OpenHR.messageBox("Cannot include the '" + sCalcName + "' calculation.\nIts hidden and you are not the creator of this definition.", 64, "Mail Merge");
            }
            else {
                iColumnsSwapped = iColumnsSwapped + 1;
                var sAddline;
                var sTemp;
                var iTemp;
                if (grdFrom.columns(0).text == 'C') {
                    sAddline = grdFrom.columns(0).text +
                        '	' + grdFrom.columns(1).text +
                        '	' + grdFrom.columns(2).text;
                    if (pfSelect == true) {
                        sAddline = sAddline + '	' + getTableName(grdFrom.columns(1).text) + '.' + grdFrom.columns(3).text;
                    }
                    else {
                        sAddline = sAddline + '	' + grdFrom.columns(3).text.substring(grdFrom.columns(3).text.indexOf(".") + 1);
                    }

                    sAddline = sAddline + '	' + grdFrom.columns(4).text +
                        '	' + grdFrom.columns(5).text +
                        '	' + grdFrom.columns(6).text +
                        '	' + grdFrom.columns(7).text;
                }
                else {
                    sAddline = grdFrom.columns(0).text +
                        '	' + grdFrom.columns(1).text +
                        '	' + grdFrom.columns(2).text;

                    if (pfSelect == true) {
                        sAddline = sAddline + '	<Calc> ' + grdFrom.columns(3).text;
                    }
                    else {
                        sTemp = grdFrom.columns(3).text;
                        iTemp = sTemp.indexOf("<Calc> ");
                        if (iTemp >= 0) {
                            sTemp = sTemp.substring(iTemp + 7);
                        }
                        sAddline = sAddline + '	' + sTemp;
                    }

                    sAddline = sAddline +
                        '	' + grdFrom.columns(4).text +
                        '	' + grdFrom.columns(5).text +
                        '	' + grdFrom.columns(6).text +
                        '	' + grdFrom.columns(7).text;
                }

                if (pfSelect == true) {
                    sAddline = sAddline +
                        '	' + grdFrom.columns(3).text +
                        '	' + '0' + '	' + '0' + '	' + '0';

                    // Remember which calcs we are adding to the report so that
                    // we can get there return types below.						
                    if (grdFrom.columns(0).text == "E") {
                        sAddedCalcIDs = sAddedCalcIDs + grdFrom.columns(2).text + ",";
                    }
                }

                if (grdFrom.columns(6).text == "Y") {
                    iHiddenCalcCount = iHiddenCalcCount + 1;
                }

                if (pfSelect == true) {
                    grdTo.MoveLast();
                    grdTo.AddItem(sAddline);
                    grdTo.MoveLast();
                }
                else {
                    /* Find the right spot to add the row. */
                    //grdTo.redraw = false;

                    var sFromType = grdFrom.columns(0).text;
                    var sFromTableID = grdFrom.columns(1).text;

                    sTemp = grdFrom.columns(3).text;
                    iTemp = sTemp.indexOf("<Calc> ");
                    if (iTemp >= 0) {
                        sTemp = sTemp.substring(iTemp + 7);
                    }
                    var sFromDisplay = replace(sTemp, "_", " ");
                    sFromDisplay = sFromDisplay.substring(sFromDisplay.indexOf(".") + 1);
                    sFromDisplay = sFromDisplay.toUpperCase();

                    var fIsFromTblAvailable = (sFromTableID == frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].value);

                    var fIsFromTypeAvailable = (((sFromType == "C") && (frmDefinition.optColumns.checked)) ||
                        ((sFromType == "E") && (frmDefinition.optCalc.checked)));
                    var fFound = true;

                    if (fIsFromTblAvailable && fIsFromTypeAvailable) {
                        fFound = false;
                        grdTo.movefirst();
                        grdTo.Redraw = true;
                        for (i2 = 0; i2 < grdTo.rows() ; i2++) {
                            grdTo.Redraw = false;

                            var sToType = grdTo.columns(0).text;
                            var sToTableID = grdTo.columns(1).text;
                            var sToDisplay = replace(grdTo.columns(3).text.toUpperCase(), "_", " ");

                            if ((sFromType == "C") && (frmDefinition.optColumns.checked)) {
                                // Column
                                if ((sToType == sFromType) && (sToDisplay > sFromDisplay)) {
                                    fFound = true;
                                }
                            }
                            else {
                                // Calculation
                                if ((sToType == sFromType) &&
                                    (sToDisplay > sFromDisplay) &&
                                    (frmDefinition.optCalc.checked)) {
                                    fFound = true;
                                }
                            }

                            if (fFound == true) {
                                grdTo.additem(sAddline, i2);
                                break;
                            }
                            grdTo.movenext();
                        } //end for loop
                    }

                    if (fFound == false) {
                        grdTo.additem(sAddline);
                    }
                }
                if (i == grdFrom.selbookmarks.Count() - 1) {
                    grdTo.Redraw = true;
                }
                grdTo.SelBookmarks.Add(grdTo.Bookmark);
            }
        }
        grdTo.Redraw = true;
        grdFrom.Redraw = true;

        if (iColumnsSwapped > 0) {

            iCount = grdFrom.selbookmarks.Count();
            for (i = iCount - 1; i >= 0; i--) {
                grdFrom.bookmark = grdFrom.selbookmarks(i);
                iRowIndex = grdFrom.AddItemRowIndex(grdFrom.Bookmark);

                //NPG20111010 Fault HRPRO-1798
                if (iRowIndex > (grdFrom.rows - 1)) iRowIndex = grdFrom.rows - 1;

                if ((grdFrom.Rows == 1) && (iRowIndex == 0)) {
                    grdFrom.RemoveAll();
                    if (pfSelect == false) {
                        // Clear the sort columns collection.
                        removeSortColumn(0, 0);
                    }
                }
                else {
                    if (pfSelect == false) {
                        // Remove the column from the sort columns collection.
                        if (grdFrom.columns(0).text == "C") {
                            removeSortColumn(grdFrom.columns(2).text, 0);
                        }

                        grdFrom.RemoveItem(iRowIndex);
                    }
                    else {
                        if ((frmDefinition.txtOwner.value.toUpperCase() == frmUseful.txtUserName.value.toUpperCase()) ||
                            (grdFrom.columns(6).text != "Y")) {
                            grdFrom.RemoveItem(iRowIndex);
                        }
                    }
                }
            }

            if (iHiddenCalcCount > 0) {
                var iOldCalcCount = new Number(frmSelectionAccess.calcsHiddenCount.value);
                if (pfSelect == true) {
                    frmSelectionAccess.calcsHiddenCount.value = iOldCalcCount + iHiddenCalcCount;
                }
                else {
                    frmSelectionAccess.calcsHiddenCount.value = iOldCalcCount - iHiddenCalcCount;
                }

                refreshTab1Controls();
            }
        }
    }

    if (iColumnsSwapped > 0) {
        frmUseful.txtChanged.value = 1;

        if (sAddedCalcIDs.length > 0) {
            // Get the return types of the added calcs.
            var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
            frmGetDataForm.txtAction.value = "GETEXPRESSIONRETURNTYPES";
            frmGetDataForm.txtParam1.value = sAddedCalcIDs;
            window.data_refreshData();
            OpenHR.getForm("dataframe", "");
        }
    }
    grdFrom.Redraw = true;
    grdTo.Redraw = true;
    refreshTab2Controls();
}
function columnSwapAll(pfSelect) {
    var i;
    var iColumnsSwapped;
    var sAddedCalcIDs;

    sAddedCalcIDs = "";
    iColumnsSwapped = 0;
    var grdFrom;
    var grdTo;

    if (pfSelect == true) {
        grdFrom = frmDefinition.ssOleDBGridAvailableColumns;
        grdTo = frmDefinition.ssOleDBGridSelectedColumns;
    }
    else {
        var iSortedColumnCount;
        if (frmUseful.txtSortLoaded.value == 1) {
            iSortedColumnCount = frmDefinition.ssOleDBGridSortOrder.Rows;
        }
        else {
            iSortedColumnCount = 0;
            var dataCollection = frmOriginalDefinition.elements;
            if (dataCollection != null) {
                for (var iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                    var sControlName = dataCollection.item(iIndex).name;
                    sControlName = sControlName.substr(0, 19);
                    if (sControlName == "txtReportDefnOrder_") {
                        iSortedColumnCount = 1;
                        break;
                    }
                }
            }
        }
        var iAnswer;
        if (iSortedColumnCount > 0) {
            iAnswer = OpenHR.messageBox("Removing all columns will remove all sort order columns. \n Are you sure ?", 36, "Mail Merge");
        }
        else {
            iAnswer = 6;
        }

        if (iAnswer == 7) {
            // cancel 
            return;
        }
        grdFrom = frmDefinition.ssOleDBGridSelectedColumns;
        grdTo = frmDefinition.ssOleDBGridAvailableColumns;
    }

    grdFrom.redraw = false;
    grdTo.redraw = false;

    grdTo.selbookmarks.removeall();

    var iHiddenCalcCount = 0;

    grdFrom.movefirst();
    for (i = 0; i < grdFrom.Rows() ; i++) {
        if ((pfSelect == true) &&
            (frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) &&
            (grdFrom.columns(6).text == "Y")) {

            var sCalcName = new String(grdFrom.columns(3).text);
            var iStringIndex = sCalcName.indexOf("> ");
            if (iStringIndex >= 0) {
                sCalcName = sCalcName.substring(iStringIndex, sCalcName.length);
            }
            OpenHR.messageBox("Cannot include the '" + sCalcName + "' calculation.\nIts hidden and you are not the creator of this definition.", 64, "Mail Merge");
        }
        else {
            iColumnsSwapped = iColumnsSwapped + 1;
            var sAddline;
            var sTemp;
            var iTemp;
            if (grdFrom.columns(0).text == 'C') {
                sAddline = grdFrom.columns(0).text +
                    '	' + grdFrom.columns(1).text +
                    '	' + grdFrom.columns(2).text;
                if (pfSelect == true) {
                    sAddline = sAddline + '	' + getTableName(grdFrom.columns(1).text) + '.' + grdFrom.columns(3).text;
                }
                else {
                    sAddline = sAddline + '	' + grdFrom.columns(3).text.substring(grdFrom.columns(3).text.indexOf(".") + 1);
                }

                sAddline = sAddline + '	' + grdFrom.columns(4).text +
                    '	' + grdFrom.columns(5).text +
                    '	' + grdFrom.columns(6).text +
                    '	' + grdFrom.columns(7).text;
            }
            else {
                sAddline = grdFrom.columns(0).text +
                    '	' + grdFrom.columns(1).text +
                    '	' + grdFrom.columns(2).text;

                if (pfSelect == true) {
                    sAddline = sAddline + '	<Calc> ' + grdFrom.columns(3).text;
                }
                else {
                    sTemp = grdFrom.columns(3).text;
                    iTemp = sTemp.indexOf("<Calc> ");
                    if (iTemp >= 0) {
                        sTemp = sTemp.substring(iTemp + 7);
                    }
                    sAddline = sAddline + '	' + sTemp;
                }

                sAddline = sAddline +
                    '	' + grdFrom.columns(4).text +
                    '	' + grdFrom.columns(5).text +
                    '	' + grdFrom.columns(6).text +
                    '	' + grdFrom.columns(7).text;
            }

            if (pfSelect == true) {
                sAddline = sAddline +
                    '	' + grdFrom.columns(3).text +
                    '	' + '0' + '	' + '0' + '	' + '0';

                // Remember which calcs we are adding to the report so that
                // we can get there return types below.						
                if (grdFrom.columns(0).text == "E") {
                    sAddedCalcIDs = sAddedCalcIDs + grdFrom.columns(2).text + ",";
                }
            }

            if (grdFrom.columns(6).text == "Y") {
                iHiddenCalcCount = iHiddenCalcCount + 1;
            }

            if (pfSelect == true) {
                grdTo.additem(sAddline);
            }
            else {
                /* Find the right spot to add the row. */
                var sFromType = grdFrom.columns(0).text;
                var sFromTableID = grdFrom.columns(1).text;

                sTemp = grdFrom.columns(3).text;
                iTemp = sTemp.indexOf("<Calc> ");
                if (iTemp >= 0) {
                    sTemp = sTemp.substring(iTemp + 7);
                }
                var sFromDisplay = replace(sTemp, "_", " ");
                sFromDisplay = sFromDisplay.substring(sFromDisplay.indexOf(".") + 1);
                sFromDisplay = sFromDisplay.toUpperCase();

                var fIsFromTblAvailable = (sFromTableID == frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].value);

                var fIsFromTypeAvailable = (((sFromType == "C") && (frmDefinition.optColumns.checked)) ||
                    ((sFromType == "E") && (frmDefinition.optCalc.checked)));

                var fFound = true;

                if (fIsFromTblAvailable && fIsFromTypeAvailable) {
                    fFound = false;
                    grdTo.movefirst();
                    grdTo.Redraw = true;
                    /* TM 19/06/02 - Fault 4000 */
                    for (var i2 = 0; i2 < grdTo.rows() ; i2++) {
                        grdTo.Redraw = false;

                        var sToType = grdTo.columns(0).text;
                        var sToTableID = grdTo.columns(1).text;
                        var sToDisplay = replace(grdTo.columns(3).text.toUpperCase(), "_", " ");

                        if ((sFromType == "C") && (frmDefinition.optColumns.checked)) {
                            // Column
                            if ((sToType == sFromType) && (sToDisplay > sFromDisplay)) {
                                fFound = true;
                            }
                        }
                        else {
                            // Calculation
                            if ((sToType == sFromType) &&
                                (sToDisplay > sFromDisplay) &&
                                (frmDefinition.optCalc.checked)) {
                                fFound = true;
                            }
                        }

                        if (fFound == true) {
                            grdTo.additem(sAddline, i2);
                            break;
                        }
                        grdTo.movenext();
                    } //end for loop

                    if (fFound == false) {
                        grdTo.additem(sAddline);
                    }
                }
            }
        }

        if (i == grdFrom.Rows() - 1) {
            grdTo.Redraw = true;
        }

        grdTo.SelBookmarks.Add(grdTo.Bookmark);
        grdFrom.MoveNext();
    }

    grdFrom.Redraw = true;
    grdTo.Redraw = true;

    if (iColumnsSwapped > 0) {
        grdFrom.RemoveAll();

        if (pfSelect == false) {
            // Clear the sort columns collection.
            removeSortColumn(0, 0);
            frmUseful.txtSortLoaded.value = 1;
        }

        if (iHiddenCalcCount > 0) {
            var iOldCalcCount = new Number(frmSelectionAccess.calcsHiddenCount.value);
            if (pfSelect == true) {
                frmSelectionAccess.calcsHiddenCount.value = iOldCalcCount + iHiddenCalcCount;
            }
            else {
                frmSelectionAccess.calcsHiddenCount.value = iOldCalcCount - iHiddenCalcCount;
            }

            refreshTab1Controls();
        }
    }

    if (iColumnsSwapped > 0) {
        frmUseful.txtChanged.value = 1;

        if (sAddedCalcIDs.length > 0) {
            // Get the return types of the added calcs.
            var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
            frmGetDataForm.txtAction.value = "GETEXPRESSIONRETURNTYPES";
            frmGetDataForm.txtParam1.value = sAddedCalcIDs;
            window.data_refreshData();
        }
    }

    refreshTab2Controls();
}
function replace(sExpression, sFind, sReplace) {
    //gi (global search, ignore case)
    var re = new RegExp(sFind, "gi");
    sExpression = sExpression.replace(re, sReplace);
    return (sExpression);
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
function columnMove(pfUp) {
    var iNewIndex;
    var iOldIndex;
    var iSelectIndex;
		
    if (pfUp == true) {
        iNewIndex = frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark) - 1;
        iOldIndex = iNewIndex + 2;
        iSelectIndex = iNewIndex;
    }
    else {
        iNewIndex = frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark) + 2;
        iOldIndex = iNewIndex - 2;
        iSelectIndex = iNewIndex - 1;
    }

    var sAddline = frmDefinition.ssOleDBGridSelectedColumns.columns(0).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(1).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(2).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(3).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(4).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(5).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(6).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(7).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(8).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(9).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(10).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(11).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(12).text +
        '	' + frmDefinition.ssOleDBGridSelectedColumns.columns(13).text;

    frmDefinition.ssOleDBGridSelectedColumns.additem(sAddline, iNewIndex);
    frmDefinition.ssOleDBGridSelectedColumns.RemoveItem(iOldIndex);

    frmDefinition.ssOleDBGridSelectedColumns.SelBookmarks.RemoveAll();
    frmDefinition.ssOleDBGridSelectedColumns.SelBookmarks.Add(frmDefinition.ssOleDBGridSelectedColumns.AddItemBookmark(iSelectIndex));
    frmDefinition.ssOleDBGridSelectedColumns.Bookmark = frmDefinition.ssOleDBGridSelectedColumns.AddItemBookmark(iSelectIndex);

    frmUseful.txtChanged.value = 1;
    refreshTab2Controls();
}
function validateColSize() {
    var sConvertedValue;
    var sDecimalSeparator;
    var sThousandSeparator;
    var sPoint;
    var tempValue;

    sDecimalSeparator = "\\";
    sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
    var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

    sThousandSeparator = "\\";
    sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
    var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

    sPoint = "\\.";
    var rePoint = new RegExp(sPoint, "gi");

    if (frmDefinition.txtSize.value == '') {
        frmDefinition.txtSize.value = 0;
    }

    tempValue = parseFloat(frmDefinition.txtSize.value);
    if (isNaN(tempValue) == false) {
        frmDefinition.txtSize.value = String(tempValue);
    }

    // Convert the value from locale to UK settings for use with the isNaN funtion.
    sConvertedValue = new String(frmDefinition.txtSize.value);
    // Remove any thousand separators.
    sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
    //frmDefinition.txtSize.value = sConvertedValue;

    // Convert any decimal separators to '.'.
    if (OpenHR.LocaleDecimalSeparator != ".") {
        {
            // Remove decimal points.
            sConvertedValue = sConvertedValue.replace(rePoint, "A");
            // replace the locale decimal marker with the decimal point.
            sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
        }

        if (isNaN(sConvertedValue) == true) {
            OpenHR.messageBox("Invalid numeric value.", 48, "Mail Merge");
            frmDefinition.txtSize.value = frmDefinition.ssOleDBGridSelectedColumns.columns(4).text;
            frmDefinition.txtSize.focus();
            return false;
        } else {
            if (sConvertedValue.indexOf(".") >= 0) {
                OpenHR.messageBox("Invalid integer value.", 48, "Mail Merge");
                frmDefinition.txtSize.value = frmDefinition.ssOleDBGridSelectedColumns.columns(4).text;
                frmDefinition.txtSize.focus();
                return false;
            } else {
                if (frmDefinition.txtSize.value < 0) {
                    OpenHR.messageBox("The value must be greater than or equal to 0.", 48, "Mail Merge");
                    frmDefinition.txtSize.value = frmDefinition.ssOleDBGridSelectedColumns.columns(4).text;
                    frmDefinition.txtSize.focus();
                    return false;
                }

                if (frmDefinition.txtSize.value > 2147483646) {
                    OpenHR.messageBox("The value must be less than or equal to 2147483646.", 48, "Mail Merge");
                    frmDefinition.txtSize.value = frmDefinition.ssOleDBGridSelectedColumns.columns(4).text;
                    frmDefinition.txtSize.focus();
                    return false;
                }

            }
        }

        frmDefinition.ssOleDBGridSelectedColumns.columns(4).text = frmDefinition.txtSize.value;
        frmUseful.txtChanged.value = 1;
        refreshTab2Controls();
        return true;
    }
    return false;
}
	
function utilDefMailmergeAddActiveXHandlers() {
    OpenHR.addActiveXHandler("ssOleDBGridAvailableColumns", "rowColChange", ssOleDbGridAvailableColumnsRowColChange);
    OpenHR.addActiveXHandler("ssOleDBGridAvailableColumns", "DblClick", ssOleDbGridAvailableColumnsDblClick);
    OpenHR.addActiveXHandler("ssOleDBGridAvailableColumns", "KeyPress(iKeyAscii)", ssOleDbGridAvailableColumnsKeyPress);

    OpenHR.addActiveXHandler("ssOleDBGridSelectedColumns", "rowColChange", ssOleDbGridSelectedColumnsRowColChange);
    OpenHR.addActiveXHandler("ssOleDBGridSelectedColumns", "DblClick", ssOleDbGridSelectedColumnsDblClick);
    OpenHR.addActiveXHandler("ssOleDBGridSelectedColumns", "SelChange", ssOleDbGridSelectedColumnsSelChange);

    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "beforerowcolchange", ssOleDbGridSortOrderBeforerowcolchange);
    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "beforeupdate", ssOleDbGridSortOrderBeforeupdate);
    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "afterinsert", ssOleDbGridSortOrderAfterinsert);
    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "rowcolchange", ssOleDbGridSortOrderRowcolchange);
    OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "change", ssOleDbGridSortOrderChange);

    OpenHR.addActiveXHandler("grdAccess", "ComboCloseUp", grdAccessComboCloseUp);
    OpenHR.addActiveXHandler("grdAccess", "GotFocus", grdAccessGotFocus);
    OpenHR.addActiveXHandler("grdAccess", "RowColChange", grdAccessRowColChange);
    OpenHR.addActiveXHandler("grdAccess", "RowLoaded", grdAccessRowLoaded);
}

//ssOleDBGridAvailableColumns handlers
function ssOleDbGridAvailableColumnsRowColChange() { refreshTab2Controls(); }
function ssOleDbGridAvailableColumnsDblClick() { columnSwap(true); }
function ssOleDbGridAvailableColumnsKeyPress() {
    if ((window.iKeyAscii >= 32) && (window.iKeyAscii <= 255)) {
        var dtTicker = new Date();
        var iThisTick = new Number(dtTicker.getTime());
        var iLastTick;
        if (window.txtLastKeyFind.value.length > 0) {
            iLastTick = new Number(window.txtTicker.value);
        } else {
            iLastTick = new Number("0");
        }

        if (iThisTick > (iLastTick + 1500)) {
            var sFind = String.fromCharCode(window.iKeyAscii);
        } else {
        	sFind = window.txtLastKeyFind.value + String.fromCharCode(window.iKeyAscii);
        }

        window.txtTicker.value = iThisTick;
        window.txtLastKeyFind.value = sFind;

        locateRecord(sFind);
    }
}
//ssOleDBGridSelectedColumns handlers
function ssOleDbGridSelectedColumnsRowColChange() {
    if (frmUseful.txtLockGridEvents.value != 1) {
        refreshTab2Controls();
    }
}
function ssOleDbGridSelectedColumnsDblClick() {
    columnSwap(false);
}
function ssOleDbGridSelectedColumnsSelChange() {
    refreshTab2Controls();
}
//ssOleDBGridSortOrder handlers
function ssOleDbGridSortOrderBeforerowcolchange() {
    //	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
    //		frmDefinition.ssOleDBGridSortOrder.columns(1).cellstyleset("ssetViewDormant", frmDefinition.ssOleDBGridSortOrder.row);
    //	}
    //	else {
    //		frmDefinition.ssOleDBGridSortOrder.columns(1).cellstyleset("ssetDormant", frmDefinition.ssOleDBGridSortOrder.row);
    //	}
}
function ssOleDbGridSortOrderBeforeupdate() {
    if ((frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Asc') && (frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Desc')) {
        frmDefinition.ssOleDBGridSortOrder.Columns(2).text = 'Asc';
    }
}
function ssOleDbGridSortOrderAfterinsert() { refreshTab3Controls(); }
function ssOleDbGridSortOrderRowcolchange() {
    frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
    frmDefinition.ssOleDBGridSortOrder.columns(1).cellstyleset("ssetSelected", frmDefinition.ssOleDBGridSortOrder.row);
    frmSortOrder.txtSortColumnID.value = frmDefinition.ssOleDBGridSortOrder.Columns(0).text;
    frmSortOrder.txtSortColumnName.value = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
    frmSortOrder.txtSortOrder.value = frmDefinition.ssOleDBGridSortOrder.Columns(2).text;
    refreshTab3Controls();
}
function ssOleDbGridSortOrderChange() {
    frmUseful.txtChanged.value = 1;
    refreshTab3Controls();
}
//grdAccess Handlers
function grdAccessComboCloseUp() {
    frmUseful.txtChanged.value = 1;
    if (frmDefinition.grdAccess.AddItemRowIndex(frmDefinition.grdAccess.Bookmark) == 0) and(frmDefinition.grdAccess.Columns("Access").Text.length > 0);
    {
    	ForceAccess(window.grdAccess, AccessCode(frmDefinition.grdAccess.Columns("Access").Text));
        frmDefinition.grdAccess.MoveFirst();
        frmDefinition.grdAccess.Col = 1;
    }
    refreshTab1Controls();
}
function grdAccessGotFocus() {
    window.grdAccess.Col = 1;
}
function grdAccessRowColChange() {
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
function grdAccessRowLoaded() {
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
function locateRecord(psSearchFor) {
    var fFound;
    var iIndex;
		
    fFound = false;
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
        } else {
            break;
        }
    }

    if ((fFound == false) && (frmDefinition.ssOleDBGridAvailableColumns.rows > 0)) {
        // Select the top row.
        frmDefinition.ssOleDBGridAvailableColumns.MoveFirst();
        frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.Add(frmDefinition.ssOleDBGridAvailableColumns.Bookmark);
    }

    frmDefinition.ssOleDBGridAvailableColumns.Redraw = true;
}
function validateColDecimals() {
	    var sConvertedValue;
	    var sDecimalSeparator;
	    var sThousandSeparator;
	    var sPoint;
	    var tempValue;

	    sDecimalSeparator = "\\";
	    sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
	    var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

	    sThousandSeparator = "\\";
	    sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleThousandSeparator);
	    var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

	    sPoint = "\\.";
	    var rePoint = new RegExp(sPoint, "gi");

	    if (frmDefinition.txtDecPlaces.value == '') {
	        frmDefinition.txtDecPlaces.value = 0;
	    }

	    tempValue = parseFloat(frmDefinition.txtDecPlaces.value);
	    if (isNaN(tempValue) == false) {
	        frmDefinition.txtDecPlaces.value = String(tempValue);
	    }

	    // Convert the value from locale to UK settings for use with the isNaN funtion.
	    sConvertedValue = new String(frmDefinition.txtDecPlaces.value);
	    // Remove any thousand separators.
	    sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
	    // Convert any decimal separators to '.'.
	    if (OpenHR.LocaleDecimalSeparator != ".") {
	        // Remove decimal points.
	        sConvertedValue = sConvertedValue.replace(rePoint, "A");
	        // replace the locale decimal marker with the decimal point.
	        sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
	    }

	    if (isNaN(sConvertedValue) == true) {
	        OpenHR.messageBox("Decimal places must be numeric.", 48, "Mail Merge");
	        frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
	        frmDefinition.txtDecPlaces.focus();
	        return false;
	    } else {
	        if (sConvertedValue.indexOf(".") >= 0) {
	            OpenHR.messageBox("Invalid integer value.", 48, "Mail Merge");
	            frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
	            frmDefinition.txtDecPlaces.focus();
	            return false;
	        } else {
	            if (frmDefinition.txtDecPlaces.value < 0) {
	                OpenHR.messageBox("The value cannot be less than 0.", 48, "Mail Merge");
	                frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
	                frmDefinition.txtDecPlaces.focus();
	                return false;
	            }

	            if (frmDefinition.txtDecPlaces.value > 99999) {
	                OpenHR.messageBox("The value must be less than or equal to 99999.", 48, "Mail Merge");
	                frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
	                frmDefinition.txtDecPlaces.focus();
	                return false;
	            }
	        }
	    }

	    if (frmDefinition.txtDecPlaces.value > 999) {
	        OpenHR.messageBox("The decimals must be less than or equal to 999.", 48, "Mail Merge");
	        frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
	        frmDefinition.txtDecPlaces.focus();
	        return false;
	    }

	    frmDefinition.ssOleDBGridSelectedColumns.columns(5).text = frmDefinition.txtDecPlaces.value;
	    frmUseful.txtChanged.value = 1;
	    refreshTab2Controls();
	    return true;
	}
function selectDocType() {
    var sURL;
    window.frmDocTypeSelection.DocTypeSelCurrentID.value = frmDefinition.txtDocTypeID.value;

    sURL = "util_doctypeSelection" +
        "?DocTypeSelCurrentID=" + window.frmDocTypeSelection.DocTypeSelCurrentID.value;
    openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");
}
function fileClear() {
    frmDefinition.txtSaveFile.value = "";
    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
}
function docTypeClear() {
    frmDefinition.txtDocType.value = "";
    frmDefinition.txtDocTypeID.value = 0;
    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
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
                    } else {
                        frmDefinition.ssOleDBGridSortOrder.RemoveItem(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark));
                    }
                } else {
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
    } else {
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (var iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                var sControlName = dataCollection.item(iIndex).name;
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
function removeCalcs(psColumns) {
    var iCharIndex;
    var sColumns;
    var sColumnID;
    var sGridColumnID;

    sColumns = new String(psColumns);

    // Remove the given calcs from the selected columns list.
    while (sColumns.length > 0) {
        iCharIndex = sColumns.indexOf(",");

        if (iCharIndex >= 0) {
            sColumnID = sColumns.substr(0, iCharIndex);
            sColumns = sColumns.substr(iCharIndex + 1);
        } else {
            sColumnID = sColumns;
            sColumns = "";
        }

        if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
            if (frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) {
                frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
                frmDefinition.ssOleDBGridSelectedColumns.movefirst();

                for (var i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
                    sGridColumnID = frmDefinition.ssOleDBGridSelectedColumns.Columns("columnID").Text;

                    if (sGridColumnID == sColumnID) {
                        frmDefinition.ssOleDBGridSelectedColumns.RemoveItem(frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark));
                        break;
                    }

                    frmDefinition.ssOleDBGridSelectedColumns.movenext();
                }

                frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
            }
        } else {
            var dataCollection = frmOriginalDefinition.elements;
            if (dataCollection != null) {
                for (var iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                    var sControlName = dataCollection.item(iIndex).name;
                    sControlName = sControlName.substr(0, 20);
                    if (sControlName == "txtReportDefnColumn_") {
                        sGridColumnID = selectedColumnParameter(dataCollection.item(iIndex).value, "COLUMNID");
                        if (sGridColumnID == sColumnID) {
                            dataCollection.item(iIndex).value = "";
                            break;
                        }
                    }
                }
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
function changeDescription() {
    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}
function changeTab4Control() {
    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
}
function populateAccessGrid() {
    frmDefinition.grdAccess.focus();
    frmDefinition.grdAccess.removeAll();

    var dataCollection = frmAccess.elements;
    if (dataCollection != null) {

        frmDefinition.grdAccess.AddItem("(All Groups)");
        for (var i = 0; i < dataCollection.length; i++) {
            frmDefinition.grdAccess.AddItem(dataCollection.item(i).value);
        }
    }
}
function setJobsToHide(psJobs) {
    frmSend.txtSend_jobsToHide.value = psJobs;
    frmSend.txtSend_jobsToHideGroups.value = frmValidate.validateHiddenGroups.value;
}
function changeName() {
    frmUseful.txtChanged.value = 1;
    refreshTab1Controls();
}
function loadEmailDefs() {

    while (frmDefinition.cboEmail.options.length > 0) {
        frmDefinition.cboEmail.options.remove(0);
    }

    //var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
    var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
    var dataCollection = frmUtilDefForm.elements;
    var i;
    if (dataCollection != null) {

        for (i = 0; i < dataCollection.length; i++) {

            var sControlName = dataCollection.item(i).name;
            sControlName = sControlName.substr(0, 9);
            if (sControlName == "txtEmail_") {
                sControlName = dataCollection.item(i).value;
                var oOption = document.createElement("OPTION");
                frmDefinition.cboEmail.options.add(oOption);
                oOption.innerText = sControlName.substr(12, 999);
                oOption.value = parseInt(sControlName.substr(0, 9));
            }
        }
    }

    try {
        for (i = 0; i < frmDefinition.cboEmail.options.length; i++) {
            if (frmDefinition.cboEmail.options(i).value == frmOriginalDefinition.txtDefn_EmailAddrID.value) {
                frmDefinition.cboEmail.selectedIndex = i;
                break;
            }
        }
    } catch (e) {
    }

}
function refreshFile() {
    var sText;
    var iIndex = frmDefinition.cboFileFormat.options[frmDefinition.cboFileFormat.selectedIndex].value;

    if (frmDefinition.txtSaveFile.value != "") {
        sText = frmDefinition.txtSaveFile.value;

        if (iIndex == 0) {
            sText = sText.substr(0, sText.length - 4) + ".htm";
        }
        if (iIndex == 1) {
            sText = sText.substr(0, sText.length - 4) + ".xls";
        }
        if (iIndex == 2) {
            sText = sText.substr(0, sText.length - 4) + ".doc";
        }

        frmDefinition.txtSaveFile.value = sText;
    }
}
function sortAdd() {
    var i;
    var sURL;

    // Loop through the columns added and populate the 
    // sort order text boxes to pass to util_sortorderselection.asp
    frmSortOrder.txtSortInclude.value = '';
    frmSortOrder.txtSortExclude.value = '';
    frmSortOrder.txtSortEditing.value = 'false';
    frmSortOrder.txtSortColumnID.value = frmDefinition.ssOleDBGridSortOrder.Columns(0).text;
    frmSortOrder.txtSortColumnName.value = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
    frmSortOrder.txtSortOrder.value = frmDefinition.ssOleDBGridSortOrder.Columns(2).text;

    if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
        frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
        frmDefinition.ssOleDBGridSelectedColumns.movefirst();

        for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
            if (frmDefinition.ssOleDBGridSelectedColumns.Columns(0).Text == 'C') {
                if (frmSortOrder.txtSortInclude.value != '') {
                    frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
                }
                frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.ssOleDBGridSelectedColumns.columns(2).text;
            }
            frmDefinition.ssOleDBGridSelectedColumns.movenext();
        }

        frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
    } else {
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (var iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                var sControlName = dataCollection.item(iIndex).name;
                sControlName = sControlName.substr(0, 20);
                if (sControlName == "txtReportDefnColumn_") {
                    var sDefnString = new String(dataCollection.item(iIndex).value);
                    if (sDefnString.length > 0) {
                        var sType = selectedColumnParameter(sDefnString, "TYPE");
                        var sColumnID = selectedColumnParameter(sDefnString, "COLUMNID");

                        if (sType == 'C') {
                            if (frmSortOrder.txtSortInclude.value != '') {
                                frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
                            }
                            frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + sColumnID;
                        }
                    }
                }
            }
        }
    }

    if (frmDefinition.ssOleDBGridSortOrder.rows > 0) {
        frmDefinition.ssOleDBGridSortOrder.Redraw = false;
        frmDefinition.ssOleDBGridSortOrder.movefirst();

        for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) {
            if (frmSortOrder.txtSortExclude.value != '') {
                frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + ',';
            }
            frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + frmDefinition.ssOleDBGridSortOrder.columns(0).text;

            frmDefinition.ssOleDBGridSortOrder.movenext();
        }

        frmDefinition.ssOleDBGridSortOrder.Redraw = true;
    }

    if (frmSortOrder.txtSortInclude.value == frmSortOrder.txtSortExclude.value) {
        OpenHR.messageBox("You must add more columns to the definition before you can add to the sort order.");
    } else {
        if (frmSortOrder.txtSortInclude.value != '') {
            sURL = "util_sortorderselection" +
                "?txtSortInclude=" + escape(frmSortOrder.txtSortInclude.value) +
                "&txtSortExclude=" + escape(frmSortOrder.txtSortExclude.value) +
                "&txtSortEditing=" + escape(frmSortOrder.txtSortEditing.value) +
                "&txtSortColumnID=" + escape(frmSortOrder.txtSortColumnID.value) +
                "&txtSortColumnName=" + escape(frmSortOrder.txtSortColumnName.value) +
                "&txtSortOrder=" + escape(frmSortOrder.txtSortOrder.value) +
                "&txtSortBOC=" + escape(frmSortOrder.txtSortBOC.value) +
                "&txtSortPOC=" + escape(frmSortOrder.txtSortPOC.value) +
                "&txtSortVOC=" + escape(frmSortOrder.txtSortVOC.value) +
                "&txtSortSRV=" + escape(frmSortOrder.txtSortSRV.value);
            openDialog(sURL, 500, 275, "yes", "yes");

            frmUseful.txtChanged.value = 1;
            refreshTab3Controls();
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

    if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
        frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
        frmDefinition.ssOleDBGridSelectedColumns.movefirst();

        for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
            if (frmDefinition.ssOleDBGridSelectedColumns.columns(0).text == "C") {

                if (frmSortOrder.txtSortInclude.value != '') {
                    frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
                }
                frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.ssOleDBGridSelectedColumns.columns(2).text;
            }
            frmDefinition.ssOleDBGridSelectedColumns.movenext();
        }

        frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
    } else {
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (i = 0; i < dataCollection.length; i++) {
                var sControlName = dataCollection.item(i).name;
                sControlName = sControlName.substr(0, 20);
                if (sControlName == "txtReportDefnColumn_") {

                    sColumnID = "";
                    sDefn = new String(dataCollection.item(i).value);
                    if (sDefn.substr(0, 1) == "C") {
                        iIndex = sDefn.indexOf("	");
                        if (iIndex > 0) {
                            sDefn = sDefn.substr(iIndex + 1);
                            iIndex = sDefn.indexOf("	");
                            if (iIndex > 0) {
                                sDefn = sDefn.substr(iIndex + 1);
                                iIndex = sDefn.indexOf("	");
                                if (iIndex > 0) {
                                    sColumnID = sDefn.substr(0, iIndex);
                                }
                            }
                        }

                        if (sColumnID != "") {
                            if (frmSortOrder.txtSortInclude.value != '') {
                                frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
                            }
                            frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + sColumnID;
                        }
                    }
                }
            }
        }
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
        "&txtSortOrder=" + escape(frmSortOrder.txtSortOrder.value) +
        "&txtSortBOC=" + escape(frmSortOrder.txtSortBOC.value) +
        "&txtSortPOC=" + escape(frmSortOrder.txtSortPOC.value) +
        "&txtSortVOC=" + escape(frmSortOrder.txtSortVOC.value) +
        "&txtSortSRV=" + escape(frmSortOrder.txtSortSRV.value);
    openDialog(sURL, 500, 275, "yes", "yes");

    frmUseful.txtChanged.value = 1;
    refreshTab3Controls();
}
function sortRemove() {
    if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count() == 0) {
        OpenHR.messageBox("You must select a column to remove.");
        return;
    }

    frmDefinition.ssOleDBGridSortOrder.RemoveItem(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.bookmark));

    if (frmDefinition.ssOleDBGridSortOrder.Rows != 0) {
        frmDefinition.ssOleDBGridSortOrder.MoveLast();
        frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
    }
    frmUseful.txtChanged.value = 1;
    refreshTab3Controls();
}
function sortRemoveAll() {
    frmDefinition.ssOleDBGridSortOrder.RemoveAll();
    frmUseful.txtChanged.value = 1;
    refreshTab3Controls();
}
function sortMove(pfUp) {
    var iNewIndex;
    var iOldIndex;
    if (pfUp == true) {
        iNewIndex = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) - 1;
        iOldIndex = iNewIndex + 2;
        var iSelectIndex = iNewIndex;
    } else {
        iNewIndex = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) + 2;
        iOldIndex = iNewIndex - 2;
        iSelectIndex = iNewIndex - 1;
    }

    var sAddline = frmDefinition.ssOleDBGridSortOrder.columns(0).text +
        '	' + frmDefinition.ssOleDBGridSortOrder.columns(1).text +
        '	' + frmDefinition.ssOleDBGridSortOrder.columns(2).text +
        '	' + frmDefinition.ssOleDBGridSortOrder.columns(3).text;
		
    frmDefinition.ssOleDBGridSortOrder.additem(sAddline, iNewIndex);
    frmDefinition.ssOleDBGridSortOrder.RemoveItem(iOldIndex);

    frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();
    frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(iSelectIndex));
    frmDefinition.ssOleDBGridSortOrder.Bookmark = frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(iSelectIndex);

    frmUseful.txtChanged.value = 1;
    refreshTab3Controls();
}
function validateTab1() {
    // check name has been entered
    if (frmDefinition.txtName.value == '') {
        OpenHR.messageBox("You must enter a name for this definition.", 48);
        displayPage(1);
        return (false);
    }

    // check base picklist
    if ((frmDefinition.optRecordSelection2.checked == true) &&
        (frmDefinition.txtBasePicklistID.value == 0)) {
        OpenHR.messageBox("You must select a picklist for the base table.", 48);
        displayPage(1);
        return (false);
    }

    // check base filter
    if ((frmDefinition.optRecordSelection3.checked == true) &&
        (frmDefinition.txtBaseFilterID.value == 0)) {
        OpenHR.messageBox("You must select a filter for the base table.", 48);
        displayPage(1);
        return (false);
    }

    return (true);
}
function validateTab3() {
    var sErrMsg;
    var iCount;
    var sControlName;

    sErrMsg = "";

    //check at least one column defined as sort order
    if (frmUseful.txtSortLoaded.value == 1) {
        if (frmDefinition.ssOleDBGridSortOrder.Rows == 0) {
            sErrMsg = "You must select at least 1 column to order the mail merge by";
        }
    } else {
        iCount = 0;
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (var i = 0; i < dataCollection.length; i++) {
                sControlName = dataCollection.item(i).name;
                sControlName = sControlName.substr(0, 19);
                if (sControlName == "txtReportDefnOrder_") {
                    if (dataCollection.item(i).value != "") {
                        iCount = iCount + 1;
                    }
                }
            }
        }

        if (iCount == 0) {
            sErrMsg = "You must select at least 1 column to order the mail merge by";
        }
    }

    if (sErrMsg.length > 0) {
        OpenHR.messageBox(sErrMsg, 48);
        displayPage(3);
        return (false);
    }

    return (true);
}
function validateTab4() {
    var sErrMsg;
	debugger;
    sErrMsg = "";
		var sPath = "";
    if (frmDefinition.txtTemplate.value == "") {
        sErrMsg = "No Template file selected.";
    }

    if (sErrMsg == "") {
    	//if (OpenHR.validateFilePath(frmDefinition.txtTemplate.value) == false) {
    	sPath = frmDefinition.txtTemplate.value;
	    if (OpenHR.ValidateFilePath(sPath)== false) {
            sErrMsg = "Template file not found.";
        }
    }

    if (sErrMsg == "") {
        if (frmDefinition.optDestination0.checked == true) {
            if (frmDefinition.chkSave.checked == true) {
                if (frmDefinition.txtSaveFile.value == "") {
                    sErrMsg = "No save output file name selected.";
                }
            }
        }
    }

    if (sErrMsg == "") {
        if (frmDefinition.optDestination1.checked == true) {
            if (frmDefinition.cboEmail.selectedIndex == -1) {
                sErrMsg = "No email column selected.";
            } else {
                if (frmDefinition.cboEmail.options[frmDefinition.cboEmail.selectedIndex].value == "") {
                    sErrMsg = "No email column selected.";
                }
            }
        }
    }

    if (sErrMsg == "") {
        var sAttachmentName = new String(frmDefinition.txtAttachmentName.value);
        if ((sAttachmentName.indexOf("/") != -1) ||
            (sAttachmentName.indexOf(":") != -1) ||
            (sAttachmentName.indexOf("?") != -1) ||
            (sAttachmentName.indexOf(String.fromCharCode(34)) != -1) ||
            (sAttachmentName.indexOf("<") != -1) ||
            (sAttachmentName.indexOf(">") != -1) ||
            (sAttachmentName.indexOf("|") != -1) ||
            (sAttachmentName.indexOf("\\") != -1) ||
            (sAttachmentName.indexOf("*") != -1)) {
            sErrMsg = "The attachment file name can not contain any of the following characters:\n/ : ? " + String.fromCharCode(34) + " < > | \\ *";
        }
    }

    if (sErrMsg == "") {
        if ((frmDefinition.optDestination0.checked == true) &&
            (frmDefinition.chkOutputScreen.checked == false) &&
            (frmDefinition.chkOutputPrinter.checked == false) &&
            (frmDefinition.chkSave.checked == false)) {
            sErrMsg = "You must select an output destination";
        }
    }

    if (sErrMsg.length > 0) {
        OpenHR.messageBox(sErrMsg, 48);
        displayPage(4);
        return (false);
    }

    return (true);
}
function populateSendForm() {
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

    frmSend.txtSend_selection.value = "0";
    frmSend.txtSend_picklist.value = "0";
    frmSend.txtSend_filter.value = "0";
    if (frmDefinition.optRecordSelection2.checked == true) {
        frmSend.txtSend_selection.value = "1";
        frmSend.txtSend_picklist.value = frmDefinition.txtBasePicklistID.value;
    }
    if (frmDefinition.optRecordSelection3.checked == true) {
        frmSend.txtSend_selection.value = "2";
        frmSend.txtSend_filter.value = frmDefinition.txtBaseFilterID.value;
    }
		
    frmSend.txtSend_templatefilename.value = frmDefinition.txtTemplate.value;
		
    if (frmDefinition.optDestination0.checked == true) {
        frmSend.txtSend_outputformat.value = 0;
    }
    if (frmDefinition.optDestination1.checked == true) {
        frmSend.txtSend_outputformat.value = 1;
    }
    if (frmDefinition.optDestination2.checked == true) {
        frmSend.txtSend_outputformat.value = 2;
    }

    frmSend.txtSend_outputsave.value = "0";
    frmSend.txtSend_outputfilename.value = "";
    frmSend.txtSend_outputscreen.value = "0";
    frmSend.txtSend_emailaddrid.value = "0";
    frmSend.txtSend_emailsubject.value = "";
    frmSend.txtSend_emailasattachment.value = "0";
    frmSend.txtSend_emailattachmentname.value = "";
    frmSend.txtSend_outputprinter.value = "0";
    frmSend.txtSend_outputprintername.value = "";
    frmSend.txtSend_documentmapid.value = "0";
    frmSend.txtSend_manualdocmanheader.value = "0";

    if (frmDefinition.optDestination0.checked == true) {
        if (frmDefinition.chkSave.checked == true) {
            frmSend.txtSend_outputsave.value = "1";
            frmSend.txtSend_outputfilename.value = frmDefinition.txtSaveFile.value;
        }
        if (frmDefinition.chkOutputScreen.checked == true) {
            frmSend.txtSend_outputscreen.value = "1";
        }

        if (frmDefinition.chkOutputPrinter.checked == true) {
            frmSend.txtSend_outputprinter.value = "1";
            if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.options.selectedIndex).text == "<Default Printer>") {
                frmSend.txtSend_outputprintername.value = "";
            } else {
                frmSend.txtSend_outputprintername.value = frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.options.selectedIndex).text;
            }
        }
    }

    if (frmDefinition.optDestination1.checked == true) {

        frmSend.txtSend_emailaddrid.value = frmDefinition.cboEmail.options(frmDefinition.cboEmail.options.selectedIndex).value;
        frmSend.txtSend_emailsubject.value = frmDefinition.txtSubject.value;
        if (frmDefinition.chkAttachment.checked == true) {
            frmSend.txtSend_emailasattachment.value = "1";
            frmSend.txtSend_emailattachmentname.value = frmDefinition.txtAttachmentName.value;
        }
    }

    if (frmDefinition.optDestination2.checked == true) {
        //Document Management stuff......
        if (frmDefinition.cboDMEngine.options(frmDefinition.cboDMEngine.options.selectedIndex).text == "<Default Printer>") {
            frmSend.txtSend_outputprintername.value = "";
        } else {
            frmSend.txtSend_outputprintername.value = frmDefinition.cboDMEngine.options(frmDefinition.cboDMEngine.options.selectedIndex).text;
        }
        if (frmDefinition.chkOutputScreen.checked = true) {
            frmSend.txtSend_outputscreen.value = "1";
        }
    }

    if (frmDefinition.chkSuppressBlanks.checked == true) {
        frmSend.txtSend_suppressblanks.value = "1";
    }

    if (frmDefinition.chkPause.checked == true) {
        frmSend.txtSend_pausebeforemerge.value = "1";
    }


    // now go through the columns grid (and sort order grid)
    var sColumns = '';

    if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
        frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
        frmDefinition.ssOleDBGridSelectedColumns.movefirst();

        for (var i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
            iNum = new Number(i + 1);
            sColumns = sColumns + iNum +
                '||' + frmDefinition.ssOleDBGridSelectedColumns.columns("type").text +
                '||' + frmDefinition.ssOleDBGridSelectedColumns.columns("columnID").text +
                '||' + frmDefinition.ssOleDBGridSelectedColumns.columns("size").text +
                '||' + frmDefinition.ssOleDBGridSelectedColumns.columns("decimals").text +
                '||' + frmDefinition.ssOleDBGridSelectedColumns.columns("numeric").text +
                '||' + getSortOrderString(frmDefinition.ssOleDBGridSelectedColumns.columns("columnID").text) +
                '**';

            frmDefinition.ssOleDBGridSelectedColumns.movenext();
        }
        frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
    } else {
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            iNum = 0;
            for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                sControlName = dataCollection.item(iIndex).name;
                sControlName = sControlName.substr(0, 20);
                if (sControlName == "txtReportDefnColumn_") {
                    iNum = iNum + 1;

                    sColumns = sColumns + iNum +
                        '||' + selectedColumnParameter(dataCollection.item(iIndex).value, "TYPE") +
                        '||' + selectedColumnParameter(dataCollection.item(iIndex).value, "COLUMNID") +
                        '||' + selectedColumnParameter(dataCollection.item(iIndex).value, "SIZE") +
                        '||' + selectedColumnParameter(dataCollection.item(iIndex).value, "DECIMALS") +
                        '||' + selectedColumnParameter(dataCollection.item(iIndex).value, "NUMERIC") +
                        '||' + getSortOrderString(selectedColumnParameter(dataCollection.item(iIndex).value, "COLUMNID")) +
                        '**';
                }
            }
        }
    }

    frmSend.txtSend_columns.value = sColumns.substr(0, 8000);
    frmSend.txtSend_columns2.value = sColumns.substr(8000, 8000);

    if (sColumns.length > 16000) {
        OpenHR.messageBox("Too many columns selected.");
        return false;
    } else {
        return true;
    }
}
function loadDefinition() {
    var i;
		
    frmDefinition.txtName.value = frmOriginalDefinition.txtDefn_Name.value;

    if ((frmUseful.txtAction.value.toUpperCase() == "EDIT") ||
        (frmUseful.txtAction.value.toUpperCase() == "VIEW")) {
        frmDefinition.txtOwner.value = frmOriginalDefinition.txtDefn_Owner.value;
    } else {
        frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
    }

    frmDefinition.txtDescription.value = frmOriginalDefinition.txtDefn_Description.value;

    setBaseTable(frmOriginalDefinition.txtDefn_BaseTableID.value);
    changeBaseTable();
    populatePrinters();
    populateDMEngine();

    // Set the basic record selection.
    var fRecordOptionSet = false;

    if (frmOriginalDefinition.txtDefn_PicklistID.value > 0) {
        button_disable(frmDefinition.cmdBasePicklist, false);
        frmDefinition.optRecordSelection2.checked = true;
        frmDefinition.txtBasePicklistID.value = frmOriginalDefinition.txtDefn_PicklistID.value;
        frmDefinition.txtBasePicklist.value = frmOriginalDefinition.txtDefn_PicklistName.value;
        fRecordOptionSet = true;
    } else {
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

    frmDefinition.txtTemplate.value = frmOriginalDefinition.txtDefn_TemplateFileName.value;
    button_disable(frmDefinition.cmdTemplateClear, false);

    if (frmOriginalDefinition.txtDefn_SuppressBlanks.value == "true") {
        frmDefinition.chkSuppressBlanks.checked = true;
    }

    if (frmOriginalDefinition.txtDefn_PauseBeforeMerge.value == "true") {
        frmDefinition.chkPause.checked = true;
    }

    if (frmOriginalDefinition.txtDefn_OutputFormat.value == 0) {
        frmDefinition.optDestination0.checked = true;
    }

    if (frmOriginalDefinition.txtDefn_OutputFormat.value == 1) {
        frmDefinition.optDestination1.checked = true;
    }

    if (frmOriginalDefinition.txtDefn_OutputFormat.value == 2) {
        frmDefinition.optDestination2.checked = true;
    }

    //Word Document option...
    if (frmOriginalDefinition.txtDefn_OutputFormat.value == 0) {
        //Populate the 'Display output on screen' row...
        if (frmOriginalDefinition.txtDefn_OutputScreen.value == "true") {

            frmDefinition.chkOutputScreen.checked = true;
        }

        //Populate the 'Send to printers' row...    
        if (frmOriginalDefinition.txtDefn_OutputPrinter.value == "true") {

            frmDefinition.chkOutputPrinter.checked = true;
            for (i = 0; i < frmDefinition.cboPrinterName.options.length; i++) {
                if (frmDefinition.cboPrinterName.options(i).innerText == frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
                    frmDefinition.cboPrinterName.selectedIndex = i;
                    break;
                }
            }
        }

        //Populate the 'Save to File' row...
        if (frmOriginalDefinition.txtDefn_OutputSave.value == "true") {
            frmDefinition.chkSave.checked = true;
            frmDefinition.txtSaveFile.value = frmOriginalDefinition.txtDefn_OutputFileName.value;
        }
    }

    //Individual EMails option...	
    /*if (frmUseful.txtEmailPermission.value == 1)
{*/
    if (frmOriginalDefinition.txtDefn_OutputFormat.value == 1) {
        //refreshDestination();
        GetEmailDefs();
        //loadEmailDefs();

        frmDefinition.txtSubject.value = frmOriginalDefinition.txtDefn_EmailSubject.value;
        if (frmOriginalDefinition.txtDefn_EmailAsAttachment.value == "true") {
            frmDefinition.chkAttachment.checked = true;
            frmDefinition.txtAttachmentName.value = frmOriginalDefinition.txtDefn_EmailAttachmentName.value;
        }
    }
    //}

    //Document Management Option...
    if (frmOriginalDefinition.txtDefn_OutputFormat.value == 2) {
        //Populate the 'Engine' dropdown...    
        if (frmOriginalDefinition.txtDefn_OutputPrinterName.value != "") {

            for (i = 0; i < frmDefinition.cboDMEngine.options.length; i++) {
                if (frmDefinition.cboDMEngine.options(i).innerText == frmOriginalDefinition.txtDefn_OutputPrinterName.value) {
                    frmDefinition.cboDMEngine.selectedIndex = i;
                    break;
                }
            }
        }

        //Document Type...
        if (frmOriginalDefinition.txtDefn_DocumentMapID.value > 0) {
            button_disable(frmDefinition.cmdSelDocType, false);
            frmDefinition.txtDocTypeID.value = frmOriginalDefinition.txtDefn_DocumentMapID.value;
            frmDefinition.txtDocType.value = frmOriginalDefinition.txtDefn_DocumentMapName.value;
        }

        //Manual doc header...
        if (frmOriginalDefinition.txtDefn_ManualDocManHeader.value == "true") {
            frmDefinition.chkManualDocManHeader.checked = true;
        }

        //Display output on screen...
        if (frmOriginalDefinition.txtDefn_OutputScreen.value == "true") {
            frmDefinition.chkOutputScreen.checked = true;
        }
    }

    refreshDestination();
		
    if ((frmOriginalDefinition.txtDefn_PicklistHidden.value.toUpperCase() == "TRUE") ||
        (frmOriginalDefinition.txtDefn_FilterHidden.value.toUpperCase() == "TRUE")) {
        frmSelectionAccess.baseHidden.value = "Y";
    }

    frmSelectionAccess.calcsHiddenCount.value = frmOriginalDefinition.txtDefn_HiddenCalcCount.value;

    frmDefinition.ssOleDBGridSelectedColumns.MoveFirst();
    frmDefinition.ssOleDBGridSelectedColumns.FirstRow = frmDefinition.ssOleDBGridSelectedColumns.Bookmark;

    frmDefinition.ssOleDBGridSortOrder.movefirst();
    frmDefinition.ssOleDBGridSortOrder.FirstRow = frmDefinition.ssOleDBGridSortOrder.bookmark;

    // If its read only, disable everything.
    if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
        disableAll();
    }

    if (frmOriginalDefinition.txtDefn_Warning.value.length > 0) {
        OpenHR.messageBox(frmOriginalDefinition.txtDefn_Warning.value);
        frmUseful.txtChanged.value = 1;
    }
}
function setFileFormat(piFormat) {
    var i;
    for (i = 0; i < frmDefinition.cboFileFormat.options.length; i++) {
        if (frmDefinition.cboFileFormat.options(i).value == piFormat) {
            frmDefinition.cboFileFormat.selectedIndex = i;
            return;
        }
    }

    frmDefinition.cboFileFormat.selectedIndex = 0;
    frmOriginalDefinition.txtDefn_DefaultExportTo.value = 0;
}
function loadSelectedColumnsDefinition() {
    var iIndex;
    var sDefnString;

    if (frmUseful.txtSelectedColumnsLoaded.value == 0) {
        frmDefinition.ssOleDBGridSelectedColumns.focus();

        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                var sControlName = dataCollection.item(iIndex).name;
                sControlName = sControlName.substr(0, 20);
                if (sControlName == "txtReportDefnColumn_") {
                    sDefnString = new String(dataCollection.item(iIndex).value);

                    if (sDefnString.length > 0) {
                        frmDefinition.ssOleDBGridSelectedColumns.AddItem(sDefnString);
                    }

                }
            }
        }
        frmUseful.txtSelectedColumnsLoaded.value = 1;
    }
}
function CheckExpressionTypes() {
    // Get the return types of the added calcs.

    var sAddedCalcIDs = "";

    frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
    frmDefinition.ssOleDBGridSelectedColumns.movefirst();

    for (var i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
        if (frmDefinition.ssOleDBGridSelectedColumns.columns(0).text == "E") {
            sAddedCalcIDs = sAddedCalcIDs + frmDefinition.ssOleDBGridSelectedColumns.columns(2).text + ",";
        }
        frmDefinition.ssOleDBGridSelectedColumns.movenext();
    }

    frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;

    var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
    frmGetDataForm.txtAction.value = "GETEXPRESSIONRETURNTYPES";
    frmGetDataForm.txtParam1.value = sAddedCalcIDs;
    window.data_refreshData();

}
function loadSortDefinition() {
    var iIndex;

    if (frmUseful.txtSortLoaded.value == 0) {
        frmDefinition.ssOleDBGridSortOrder.focus();

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

        frmUseful.txtSortLoaded.value = 1;
    }
}
function saveFile() {
    window.dialog.CancelError = true;
    window.dialog.FileName = frmDefinition.txtSaveFile.value;
    window.dialog.DialogTitle = "Mail Merge Output Document";
    window.dialog.Filter = frmDefinition.txtWordFormats.value;
    window.dialog.FilterIndex = frmDefinition.txtWordFormatDefaultIndex.value;
    window.dialog.Flags = 2621446;

    try {
        window.dialog.ShowSave();
    } catch (e) {
    }

    if (window.dialog.FileName.length > 256) {
        OpenHR.messageBox("Path and file name must not exceed 256 characters in length");
        return;
    }

    frmDefinition.txtSaveFile.value = window.dialog.FileName;
    frmUseful.txtChanged.value = 1;
    refreshTab4Controls();
}
function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll) {
    var dlgwinprops = "center:yes;" +
        "dialogHeight:" + pHeight + "px;" +
        "dialogWidth:" + pWidth + "px;" +
        "help:no;" +
        "resizable:" + psResizable + ";" +
        "scroll:" + psScroll + ";" +
        "status:no;";
    window.showModalDialog(pDestination, self, dlgwinprops);
}
function getSortOrderString(piColumnID) {
    var i;
    var iNum;
    var sTemp;

    if (frmUseful.txtSortLoaded.value == 1) {
        frmDefinition.ssOleDBGridSortOrder.movefirst();

        for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) {
            if (frmDefinition.ssOleDBGridSortOrder.Columns(0).text == piColumnID) {
                // its there !
                iNum = new Number(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.bookmark) + 1);

                sTemp = iNum + '||' +
                    frmDefinition.ssOleDBGridSortOrder.columns("order").text + '||';
                return (sTemp);
            }

            frmDefinition.ssOleDBGridSortOrder.movenext();
        }
    } else {
        iNum = 0;
        var dataCollection = frmOriginalDefinition.elements;
        if (dataCollection != null) {
            for (i = 0; i < dataCollection.length; i++) {
                var sControlName = dataCollection.item(i).name;
                sControlName = sControlName.substr(0, 19);
                if (sControlName == "txtReportDefnOrder_") {
                    iNum = iNum + 1;
                    var sDefn = new String(dataCollection.item(i).value);

                    if (sortColumnParameter(sDefn, "COLUMNID") == piColumnID) {
                        // its there !
                        sTemp = iNum + '||' +
                            sortColumnParameter(sDefn, "ORDER") + '||';
                        return (sTemp);
                    }
                }
            }
        }
    }

    return ('0||0||0||0||');
}
function loadAvailableColumns() {
    var i;
    var sSelectedIDs;
    var sTemp;
    var iIndex;
    var sType;
    var sID;
    var sTableID;
    var sAddString;

    frmUseful.txtLockGridEvents.value = 1;

    if (frmUseful.txtTablesChanged.value == 1) {
        // Get the columns/calcs for the current table selection.
        var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");

        frmGetDataForm.txtAction.value = "LOADREPORTCOLUMNS";
        frmGetDataForm.txtReportBaseTableID.value = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;
        frmGetDataForm.txtReportParent1TableID.value = frmDefinition.txtParent1ID.value;
        frmGetDataForm.txtReportParent2TableID.value = frmDefinition.txtParent2ID.value;
        frmGetDataForm.txtReportChildTableID.value = '0';

        window.data_refreshData();
        frmUseful.txtTablesChanged.value = 0;
    }

    sSelectedIDs = selectedIDs();
    frmDefinition.ssOleDBGridAvailableColumns.RemoveAll();

    //var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
    var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
    var dataCollection = frmUtilDefForm.elements;

    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {

            var sControlName = dataCollection.item(i).name;
            sControlName = sControlName.substr(0, 10);
            if (sControlName == "txtRepCol_") {
                sAddString = dataCollection.item(i).value;

                sType = selectedColumnParameter(sAddString, "TYPE");
                sID = selectedColumnParameter(sAddString, "COLUMNID");
                sTableID = selectedColumnParameter(sAddString, "TABLEID");

                sTemp = "	" + sType + sID + "	";

                if (((frmDefinition.optCalc.checked && (sType == 'E'))
                        || (frmDefinition.optColumns.checked && (sType == 'C')))
                    && (sTableID == frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].value)
                    && (sSelectedIDs.indexOf(sTemp) < 0)) {
                    frmDefinition.ssOleDBGridAvailableColumns.AddItem(sAddString);
                }
            }
        }
    }
    frmUseful.txtLockGridEvents.value = 0;
    refreshTab2Controls();
    // Get menu.asp to refresh the menu.
    menu_refreshMenu();
}
function loadExpressionTypes() {
    var sControlName;
    var sValue;
    var sExprID;
    var i;
    var iLoop;
    var asExprIDs = new Array();
    var asTypes = new Array();
    var iFoundCount;

    //var frmExprTypeForm = window.parent.frames("dataframe").document.forms("frmData");
    var frmExprTypeForm = OpenHR.getForm("dataframe", "frmData");
    var dataCollection = frmExprTypeForm.elements;
    iFoundCount = 0;

    frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;

    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {
            sControlName = dataCollection.item(i).name;
            sControlName = sControlName.substr(0, 12);
            if (sControlName == "txtExprType_") {
                sExprID = dataCollection.item(i).name;
                sExprID = sExprID.substr(12);
                sValue = dataCollection.item(i).value;

                asExprIDs[iFoundCount] = sExprID;
                asTypes[iFoundCount] = sValue;
                iFoundCount = iFoundCount + 1;
            }
        }
    }

    frmDefinition.ssOleDBGridSelectedColumns.movefirst();

    for (iLoop = 0; iLoop < frmDefinition.ssOleDBGridSelectedColumns.rows; iLoop++) {
        if (frmDefinition.ssOleDBGridSelectedColumns.Columns("type").Text == 'E') {
            for (i = 0; i < iFoundCount; i++) {
                if (frmDefinition.ssOleDBGridSelectedColumns.Columns("columnID").Text == asExprIDs[i]) {
                    if (asTypes[i] == 2) {
                        frmDefinition.ssOleDBGridSelectedColumns.columns(7).text = '1';
                    } else {
                        frmDefinition.ssOleDBGridSelectedColumns.columns(7).text = '0';
                    }

                    break;
                }
            }
        }

        frmDefinition.ssOleDBGridSelectedColumns.movenext();
    }

    frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;

    refreshTab2Controls();

    // Get menu.asp to refresh the menu.
    menu_refreshMenu();

    //have added this as the available columns data has be wiped.
    frmUseful.txtTablesChanged.value = 1;

    var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
    frmGetDataForm.txtAction.value = "LOADREPORTCOLUMNS";
    frmGetDataForm.txtReportBaseTableID.value = frmUseful.txtCurrentBaseTableID.value;
    frmGetDataForm.txtReportParent1TableID.value = frmDefinition.txtParent1ID.value;
    frmGetDataForm.txtReportParent2TableID.value = frmDefinition.txtParent2ID.value;
    frmGetDataForm.txtReportChildTableID.value = 0;
    window.data_refreshData();

}
function disableAll() {
    var i;

    var dataCollection = frmDefinition.elements;
    if (dataCollection != null) {
        for (i = 0; i < dataCollection.length; i++) {
            var eElem = frmDefinition.elements[i];

            if ("text" == eElem.type) {
                text_disable(eElem, true);
            } else if ("TEXTAREA" == eElem.tagName) {
                textarea_disable(eElem, true);
            } else if ("checkbox" == eElem.type) {
                checkbox_disable(eElem, true);
            } else if ("radio" == eElem.type) {
                radio_disable(eElem, true);
            } else if ("button" == eElem.type) {
                if (eElem.value != "Cancel") {
                    button_disable(eElem, true);
                }
            } else if ("SELECT" == eElem.tagName) {
                combo_disable(eElem, true);
            } else {
                grid_disable(eElem, true);
            }
        }
    }
}
function refreshDestination() {
    window.row1.style.visibility = "hidden";
    window.row1.style.display = "none";
    window.row2.style.visibility = "hidden";
    window.row2.style.display = "none";
    window.row3.style.visibility = "hidden";
    window.row3.style.display = "none";
    window.row4.style.visibility = "hidden";
    window.row4.style.display = "none";
    window.row5.style.visibility = "hidden";
    window.row5.style.display = "none";
    window.row6.style.visibility = "hidden";
    window.row6.style.display = "none";
    window.row7.style.visibility = "hidden";
    window.row7.style.display = "none";
    window.row8.style.visibility = "hidden";
    window.row8.style.display = "none";
    window.row9.style.visibility = "hidden";
    window.row9.style.display = "none";
    window.row10.style.visibility = "hidden";
    window.row10.style.display = "none";

    if (frmDefinition.optDestination0.checked == true) {
        window.row4.style.visibility = "visible";
        window.row4.style.display = "block";
        window.row5.style.visibility = "visible";
        window.row5.style.display = "block";
        window.row6.style.visibility = "visible";
        window.row6.style.display = "block";

        if (frmUseful.txtLoading.value == 'N') {
            frmDefinition.chkOutputScreen.checked = false;
            frmDefinition.chkOutputPrinter.checked = false;
            frmDefinition.chkSave.checked = false;
        }
        chkSave_Click();
        chkOutputPrinter_Click();
    } else if (frmDefinition.optDestination1.checked == true) {
        window.row7.style.visibility = "visible";
        window.row7.style.display = "block";
        window.row8.style.visibility = "visible";
        window.row8.style.display = "block";
        window.row9.style.visibility = "visible";
        window.row9.style.display = "block";
        window.row10.style.visibility = "visible";
        window.row10.style.display = "block";

        if (frmUseful.txtLoading.value == 'N') {
            GetEmailDefs();
            frmDefinition.txtSubject.value = "";
            frmDefinition.chkAttachment.checked = false;
        }
        chkAttachment_Click();
    } else if (frmDefinition.optDestination2.checked == true) {
        window.row1.style.visibility = "visible";
        window.row1.style.display = "block";
        window.row2.style.visibility = "visible";
        window.row2.style.display = "block";
        window.row3.style.visibility = "visible";
        window.row3.style.display = "block";
        window.row4.style.visibility = "visible";
        window.row4.style.display = "block";
    }

    // Get menu.asp to refresh the menu.
    menu_refreshMenu();
}
function populatePrinters() {
    with (frmDefinition.cboPrinterName) {
        var strCurrentPrinter = '';
        selectedIndex.value = 0;
        if (selectedIndex > 0) {
            strCurrentPrinter = options[selectedIndex].innerText;
        }

        oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "<Default Printer>";
        oOption.value = 0;

        for (var iLoop = 0; iLoop < OpenHR.PrinterCount() ; iLoop++) {
            var oOption = document.createElement("OPTION");
            options.add(oOption);
            oOption.innerText = OpenHR.PrinterName(iLoop);
            oOption.value = iLoop + 1;

            if (oOption.innerText == strCurrentPrinter) {
                selectedIndex.value = iLoop + 1;
            }
        }
    }
}
function populateDMEngine() {
    with (frmDefinition.cboDMEngine) {
        var strCurrentDMEngine = '';
        selectedIndex.value = 0;
        if (selectedIndex > 0) {
            strCurrentDMEngine = options[selectedIndex].innerText;
        }

        oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "<Default Printer>";
        oOption.value = 0;

        for (var iLoop = 0; iLoop < OpenHR.PrinterCount() ; iLoop++) {
            var oOption = document.createElement("OPTION");
            options.add(oOption);
            oOption.innerText = OpenHR.PrinterName(iLoop);
            oOption.value = iLoop + 1;

            if (oOption.innerText == strCurrentDMEngine) {
                selectedIndex.value = iLoop + 1;
            }
        }
    }
}
function selectedIDs() {
    var i;
    var sSelectedIDs;

    if (frmDefinition.ssOleDBGridSelectedColumns.rows == 0) return "";

    sSelectedIDs = "	";

    frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
    frmDefinition.ssOleDBGridSelectedColumns.movefirst();

    for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
        sSelectedIDs = sSelectedIDs +
            frmDefinition.ssOleDBGridSelectedColumns.Columns("type").Text +
            frmDefinition.ssOleDBGridSelectedColumns.Columns("columnID").Text +
            "	";

        frmDefinition.ssOleDBGridSelectedColumns.movenext();
    }

    frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
    return sSelectedIDs;
}
function selectedColumnParameter(psDefnString, psParameter) {
    var iCharIndex;
    var sDefn;

    sDefn = new String(psDefnString);

    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) {
        if (psParameter == "TYPE") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) {
            if (psParameter == "TABLEID") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            iCharIndex = sDefn.indexOf("	");
            if (iCharIndex >= 0) {
                if (psParameter == "COLUMNID") return sDefn.substr(0, iCharIndex);
                sDefn = sDefn.substr(iCharIndex + 1);
                iCharIndex = sDefn.indexOf("	");
                if (iCharIndex >= 0) {
                    if (psParameter == "DISPLAY") return sDefn.substr(0, iCharIndex);
                    sDefn = sDefn.substr(iCharIndex + 1);
                    iCharIndex = sDefn.indexOf("	");
                    if (iCharIndex >= 0) {
                        if (psParameter == "SIZE") return sDefn.substr(0, iCharIndex);
                        sDefn = sDefn.substr(iCharIndex + 1);
                        iCharIndex = sDefn.indexOf("	");
                        if (iCharIndex >= 0) {
                            if (psParameter == "DECIMALS") return sDefn.substr(0, iCharIndex);
                            sDefn = sDefn.substr(iCharIndex + 1);
                            iCharIndex = sDefn.indexOf("	");
                            if (iCharIndex >= 0) {
                                if (psParameter == "HIDDEN") return sDefn.substr(0, iCharIndex);
                                sDefn = sDefn.substr(iCharIndex + 1);
                                iCharIndex = sDefn.indexOf("	");
                                if (iCharIndex >= 0) {
                                    if (psParameter == "NUMERIC") return sDefn.substr(0, iCharIndex);
                                    sDefn = sDefn.substr(iCharIndex + 1);
                                    iCharIndex = sDefn.indexOf("	");
                                    if (iCharIndex >= 0) {
                                        if (psParameter == "HEADING") return sDefn.substr(0, iCharIndex);
                                        sDefn = sDefn.substr(iCharIndex + 1);
                                        iCharIndex = sDefn.indexOf("	");
                                        if (iCharIndex >= 0) {
                                            if (psParameter == "AVERAGE") return sDefn.substr(0, iCharIndex);
                                            sDefn = sDefn.substr(iCharIndex + 1);
                                            iCharIndex = sDefn.indexOf("	");
                                            if (iCharIndex >= 0) {
                                                if (psParameter == "COUNT") return sDefn.substr(0, iCharIndex);
                                                sDefn = sDefn.substr(iCharIndex + 1);
                                                iCharIndex = sDefn.indexOf("	");
                                                if (iCharIndex >= 0) {
                                                    if (psParameter == "TOTAL") return sDefn.substr(0, iCharIndex);
                                                    sDefn = sDefn.substr(iCharIndex + 1);
                                                    iCharIndex = sDefn.indexOf("	");
                                                    if (iCharIndex >= 0) {
                                                        if (psParameter == "HIDDEN") return sDefn.substr(0, iCharIndex);
                                                        sDefn = sDefn.substr(iCharIndex + 1);

                                                        if (psParameter == "GROUPWITHNEXT") return sDefn;

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
            iCharIndex = sDefn.indexOf("	");
            if (iCharIndex >= 0) {
                if (psParameter == "ORDER") return sDefn.substr(0, iCharIndex);
                sDefn = sDefn.substr(iCharIndex + 1);
                iCharIndex = sDefn.indexOf("	");
                if (iCharIndex >= 0) {
                    if (psParameter == "BOC") return sDefn.substr(0, iCharIndex);
                    sDefn = sDefn.substr(iCharIndex + 1);
                    iCharIndex = sDefn.indexOf("	");
                    if (iCharIndex >= 0) {
                        if (psParameter == "POC") return sDefn.substr(0, iCharIndex);
                        sDefn = sDefn.substr(iCharIndex + 1);
                        iCharIndex = sDefn.indexOf("	");
                        if (iCharIndex >= 0) {
                            if (psParameter == "VOC") return sDefn.substr(0, iCharIndex);
                            sDefn = sDefn.substr(iCharIndex + 1);
                            iCharIndex = sDefn.indexOf("	");
                            if (iCharIndex >= 0) {
                                if (psParameter == "SRV") return sDefn.substr(0, iCharIndex);
                                sDefn = sDefn.substr(iCharIndex + 1);

                                if (psParameter == "TABLEID") return sDefn;
                            }
                        }
                    }
                }
            }
        }
    }

    return "";
}
function GetEmailDefs() {
    var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");

    frmGetDataForm.txtAction.value = "LOADEMAILDEFINITIONS";
    frmGetDataForm.txtReportBaseTableID.value = frmUseful.txtCurrentBaseTableID.value;
    frmGetDataForm.txtReportParent1TableID.value = 0;
    frmGetDataForm.txtReportParent2TableID.value = 0;
    frmGetDataForm.txtReportChildTableID.value = 0;
    window.data_refreshData();

}
function chkAttachment_Click() {
    var blnDisabled;
    var strTempFileName;

    blnDisabled = (frmDefinition.chkAttachment.checked == false);
    text_disable(frmDefinition.txtAttachmentName, blnDisabled);

    if (blnDisabled == true) {
        frmDefinition.txtAttachmentName.value = "";
    } else {
        if (frmDefinition.txtAttachmentName.value == "") {
            strTempFileName = frmDefinition.txtTemplate.value;
            strTempFileName = strTempFileName.substr(strTempFileName.lastIndexOf("\\") + 1, 255);
            frmDefinition.txtAttachmentName.value = strTempFileName;
        }
    }
}
function chkSave_Click() {
    var blnDisabled;

    blnDisabled = (frmDefinition.chkSave.checked == false);
    text_disable(frmDefinition.txtSaveFile, true);
    button_disable(frmDefinition.cmdSaveFile, blnDisabled);
    button_disable(frmDefinition.cmdClearFile, blnDisabled);
    //checkbox_disable(frmDefinition.chkOutputScreen, blnDisabled);

    if (blnDisabled == true) {
        frmDefinition.txtSaveFile.value = "";
    }
}
function chkOutputPrinter_Click() {
    var blnDisabled;

    blnDisabled = (frmDefinition.chkOutputPrinter.checked == false);
    combo_disable(frmDefinition.cboPrinterName, blnDisabled);
    text_disable(frmDefinition.txtSaveFile, true);

    if (blnDisabled == true) {
        frmDefinition.cboPrinterName.selectedIndex = 0;
    }
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
function AllHiddenAccessMM(pgrdAccess) {
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
