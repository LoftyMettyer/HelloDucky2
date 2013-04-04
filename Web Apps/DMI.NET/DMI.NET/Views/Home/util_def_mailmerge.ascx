<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">

	function util_def_mailmerge_window_onload() {
		var fOK;
		fOK = true;	

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
			} catch(e) {
			}

			// Get menu.asp to refresh the menu.
			OpenHR.getform("workframe", "frmGoto");

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
	function populateTableAvailable(){
		var i;
	
		with (frmDefinition)
		{
			//Clear the existing data in the child table combo
			while (cboTblAvailable.options.length > 0) 
			{
				cboTblAvailable.options.remove(0);
			}
		
			//add the base table to the available tables list
			var sTableID = cboBaseTable.options[cboBaseTable.selectedIndex].value;
			var oOption = document.createElement("OPTION");
			cboTblAvailable.options.add(oOption);
			oOption.innerText = cboBaseTable.options[cboBaseTable.selectedIndex].innerText;
			oOption.value = sTableID;			
			oOption.selected = true;
		
			//add the Parent 1 table to the available tables list (if it exists)
			if (txtParent1ID.value > 0)
			{
				sTableID = txtParent1ID.value;
				oOption = document.createElement("OPTION");
				cboTblAvailable.options.add(oOption);
				oOption.innerText = txtParent1.value;
				oOption.value = sTableID;			
			}
		
			//add the Parent 2 table to the available tables list (if it exists)
			if (txtParent2ID.value > 0)
			{
				sTableID = txtParent2ID.value;
				oOption = document.createElement("OPTION");
				cboTblAvailable.options.add(oOption);
				oOption.innerText = txtParent2.value;
				oOption.value = sTableID;			
			}
		}
	}
	function refreshAvailableColumns()
	{
		if (frmUseful.txtLoading.value == 'N')
		{
			loadAvailableColumns();
		}
	}
	function TemplateSelect() {

		if (frmDefinition.txtTemplate.value.length == 0) {
			var sKey = new String("documentspath_");
			sKey = sKey.concat(OpenHR.getForm("menuframe", "frmMenuInfo").txtDatabase.value);
			OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			window.dialog.InitDir = sPath;
		}
		else {
			dialog.FileName = frmDefinition.txtTemplate.value;
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
			if (OpenHR.validateFilePath(dialogClass.FileName) == false) {
				{
					var iResponse = OpenHR.messageBox("Template file does not exist.  Create it now?", 36);
					if (iResponse == 6) {
						frmDefinition.txtTemplate.value = dialog.FileName;
						button_disable(frmDefinition.cmdTemplateClear, false);


						try {
							var sOfficeSaveAsValues = '<%=session("OfficeSaveAsValues")%>';
							OpenHR.SaveAsValues = sOfficeSaveAsValues;
							//TODO
							debugger;
							//window.parent.frames("menuframe").ASRIntranetFunctions.MM_WORD_CreateTemplateFile(dialog.FileName);
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
	var iLoop;
	
	window.parent.frames("refreshframe").document.forms("frmRefresh").submit();
			
	if (piPageNumber == 1) {
		div1.style.visibility="visible";
		div1.style.display="block";
		div2.style.visibility="hidden";
		div2.style.display="none";
		div3.style.visibility="hidden";
		div3.style.display="none";
		div4.style.visibility="hidden";
		div4.style.display="none";

		button_disable(frmDefinition.btnTab1, true);
		button_disable(frmDefinition.btnTab2, false);
		button_disable(frmDefinition.btnTab3, false);
		button_disable(frmDefinition.btnTab4, false);


		try 
		{
			frmDefinition.txtName.focus();
		}
		catch (e) {}


		refreshTab1Controls();
	}

	if (piPageNumber == 2) {
		// Get the columns/calcs for the current tvable selection.
		var frmGetDataForm = OpenHR.getForm("dataframe","frmGetData");

		if(frmUseful.txtTablesChanged.value == 1) {
			frmGetDataForm.txtAction.value = "LOADREPORTCOLUMNS";
			frmGetDataForm.txtReportBaseTableID.value = frmUseful.txtCurrentBaseTableID.value;
			frmGetDataForm.txtReportParent1TableID.value = frmDefinition.txtParent1ID.value;
			frmGetDataForm.txtReportParent2TableID.value = frmDefinition.txtParent2ID.value;
			frmGetDataForm.txtReportChildTableID.value = 0;			
			window.data_refreshData();

			frmUseful.txtTablesChanged.value = 0;
		}

		div1.style.visibility="hidden";
		div1.style.display="none";
		div2.style.visibility="visible";
		div2.style.display="block";
		div3.style.visibility="hidden";
		div3.style.display="none";
		div4.style.visibility="hidden";
		div4.style.display="none";
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

		div1.style.visibility="hidden";
		div1.style.display="none";
		div2.style.visibility="hidden";
		div2.style.display="none";
		div3.style.visibility="visible";
		div3.style.display="block";
		div4.style.visibility="hidden";
		div4.style.display="none";
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
		div1.style.visibility="hidden";
		div1.style.display="none";
		div2.style.visibility="hidden";
		div2.style.display="none";
		div3.style.visibility="hidden";
		div3.style.display="none";
		div4.style.visibility="visible";
		div4.style.display="block";
		button_disable(frmDefinition.btnTab1, false);
		button_disable(frmDefinition.btnTab2, false);
		button_disable(frmDefinition.btnTab3, false);
		button_disable(frmDefinition.btnTab4, true);
		
		refreshTab4Controls();
	}

	// Little dodge to get around a browser bug that
	// does not refresh the display on all controls.
	try
	{
		window.resizeBy(0,-1);
		window.resizeBy(0,1);
	}
	catch(e) {}
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
			sControlName = dataCollection.item(i).name;
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
	
	populateTableAvailable();
}
	function changeBaseTable() 
{
	var i;
	
	if (frmUseful.txtLoading.value == 'N') {
		if ((frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) ||
			((frmUseful.txtAction.value.toUpperCase() != "NEW") && 
			(frmUseful.txtSelectedColumnsLoaded.value == 0))) {

			iAnswer = OpenHR.messageBox("Warning: Changing the base table will result in all table/column specific aspects of this mail merge definition being cleared. Are you sure you wish to continue?",36);
			if (iAnswer == 7)	{
				// cancel and change back ! (txtcurrentbasetable)
				setBaseTable(frmUseful.txtCurrentBaseTableID.value);
				return;
			}
			else	{
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
		}
		else {
			frmUseful.txtChanged.value = 1;
		}
	}

	clearBaseTableRecordOptions();

	//Empty the parent textboxes
	frmDefinition.txtParent1.value = ''
	frmDefinition.txtParent1ID.value = 0
	frmDefinition.txtParent2.value = ''
	frmDefinition.txtParent2ID.value = 0

	var sParents = new String("");
	var dataCollection = frmTables.elements;
	if (dataCollection!=null) {
		sReqdControlName = new String("txtTableParents_");
		sReqdControlName = sReqdControlName.concat(frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);
				
		for (i=0; i<dataCollection.length; i++)  {
			sControlName = dataCollection.item(i).name;
					
			if (sControlName == sReqdControlName) {
				sParents = dataCollection.item(i).value;
				break;
			}
		}
	}

	iIndex = sParents.indexOf("	");
	if (iIndex > 0) {
		sParent1ID = sParents.substr(0, iIndex);
		frmDefinition.txtParent1.value = getTableName(sParent1ID)
		frmDefinition.txtParent1ID.value = sParent1ID
		sParents = sParents.substr(iIndex + 1);
	}
	iIndex = sParents.indexOf("	");
	if (iIndex > 0) {
		sParent2ID = sParents.substr(0, iIndex);
		frmDefinition.txtParent2.value = getTableName(sParent2ID)
		frmDefinition.txtParent2ID.value = sParent2ID
	}

	refreshTab1Controls();
	frmUseful.txtCurrentBaseTableID.value = frmDefinition.cboBaseTable.options(frmDefinition.cboBaseTable.options.selectedIndex).value;
	frmUseful.txtTablesChanged.value = 1;

	refreshDestination();
	populateTableAvailable();
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
		(frmSelectionAccess.p1Hidden.value == "Y") || 
		(frmSelectionAccess.p2Hidden.value == "Y") || 
		(frmSelectionAccess.childHidden.value == "Y") || 
		(frmSelectionAccess.calcsHiddenCount.value > 0));
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
		else
		{
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
	try
	{
		window.resizeBy(0,-1);
		window.resizeBy(0,1);
	}
	catch(e) {}
}
	function refreshTab2Controls()
{
	var fAddDisabled;
	var fAddAllDisabled;
	var fRemoveDisabled;
	var fRemoveAllDisabled;
	var fMoveUpDisabled;
	var fMoveDownDisabled;

	var fTableColDisabled;

	fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

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
		if (frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark) == 0)	{
			fMoveUpDisabled = true; 
		}

		// Are we on the bottom row ?
		if (frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark) == frmDefinition.ssOleDBGridSelectedColumns.rows-1) {
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

	fSizeDisabled = true;
	sSize = "";
	fDecPlacesDisabled = true;
	sDecPlaces = "";

	if (frmDefinition.ssOleDBGridSelectedColumns.SelBookmarks.Count == 1)	{
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
	function refreshTab3Controls()
{
	var i;
	var iCount;
	
	fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

	fSortAddDisabled = fViewing;
	fSortEditDisabled = fViewing;
	fSortRemoveDisabled = fViewing;
	fSortRemoveAllDisabled = fViewing;
	fSortMoveUpDisabled = fViewing;
	fSortMoveDownDisabled = fViewing;
	
	if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
		if (frmDefinition.ssOleDBGridSelectedColumns.Rows <= frmDefinition.ssOleDBGridSortOrder.Rows)	{
			// Disable 'Add' if there are no more columns to sort by.
			fSortAddDisabled = true;
		}
	}
	else {
		iCount = 0;
		var dataCollection = frmOriginalDefinition.elements;
		if (dataCollection!=null) {
			for (i=0; i<dataCollection.length; i++)  {
				sControlName = dataCollection.item(i).name;
				sControlName = sControlName.substr(0, 20);
				if (sControlName == "txtReportDefnColumn_") {
					iCount = iCount + 1;
				}				
			}	
		}

		if (iCount <= frmDefinition.ssOleDBGridSortOrder.Rows)	{
			// Disable 'Add' if there are no more columns to sort by.
			fSortAddDisabled = true;
		}	
	}
		
	//  if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count == 0) {
	if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count < 1) {
		fSortRemoveDisabled = true;
	}	

	if (frmDefinition.ssOleDBGridSortOrder.rows <= 0)
	{
		fSortRemoveAllDisabled = true;
	}
	
	if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count != 1){
		fSortEditDisabled = true;
		fSortMoveDownDisabled = true;
		fSortMoveUpDisabled = true;
	}	
	else {
		// Are we on the top row ?
		if ((frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) == 0) 
			|| (frmDefinition.ssOleDBGridSortOrder.rows <= 1)){
			fSortMoveUpDisabled = true; 
		}

		// Are we on the bottom row ?
		if ((frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) == frmDefinition.ssOleDBGridSortOrder.rows - 1) 
			|| (frmDefinition.ssOleDBGridSortOrder.rows <= 1)){
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
	function refreshTab4Controls()
{
	var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

	button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
		(fViewing == true)));
}
	function changeBaseTableRecordOptions()
{
	frmDefinition.txtBasePicklist.value = '';
	frmDefinition.txtBasePicklistID.value = 0;

	frmDefinition.txtBaseFilter.value = '';
	frmDefinition.txtBaseFilterID.value = 0;

	frmSelectionAccess.baseHidden.value = "N";

	//frmDefinition.chkPrintFilter.checked = false;

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
	
	//frmDefinition.chkPrintFilter.checked = false;

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
	frmRecordSelection.recSelDefType.value = "Mail Merge";
	
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
	function setRecordsNumeric()
{
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

	if(isNaN(sConvertedValue) == true) {
		OpenHR.messageBox("No. of records must be numeric.");
		frmDefinition.txtChildRecords.value = 0;
	}
	else {
		if (sConvertedValue.indexOf(".") >= 0 ) {
			OpenHR.messageBox("Invalid integer value.");
			frmDefinition.txtChildRecords.value = 0;
		}
		else {
			if (frmDefinition.txtChildRecords.value < 0 ) {
				OpenHR.messageBox("The value cannot be negative.");
				frmDefinition.txtChildRecords.value = 0;
			}
		}
	}

	refreshTab2Controls();
}	
	function validateTab2() {
	var i;
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
	function submitDefinition()
{
	var i;
	var iIndex;
	var sColumnID;
	var sType;
	
	if (validateTab1() == false) {menu_refreshMenu(); return;}
	if (validateTab2() == false) { menu_refreshMenu(); return; }
	if (validateTab3() == false) {menu_refreshMenu(); return;}
	if (validateTab4() == false) {menu_refreshMenu(); return;}
	if (populateSendForm() == false) {menu_refreshMenu(); return;}

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
	
	if(frmUseful.txtAction.value.toUpperCase() == "EDIT"){
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

		for (i=0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
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
		if (dataCollection!=null) {
			for (iIndex=0; iIndex<dataCollection.length; iIndex++)  {
				sControlName = dataCollection.item(iIndex).name;
				sControlName = sControlName.substr(0, 20);
				if (sControlName == "txtReportDefnColumn_") {
					sDefnString = new String(dataCollection.item(iIndex).value);
					
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
	
	sHiddenGroups = HiddenGroups(frmDefinition.grdAccess);
	frmValidate.validateHiddenGroups.value = sHiddenGroups;

	sURL = "dialog" +
		"?validateBaseFilter=" + escape(frmValidate.validateBaseFilter.value) +
		"&validateBasePicklist=" + escape(frmValidate.validateBasePicklist.value) +
		"&validateCalcs=" + escape(frmValidate.validateCalcs.value) +
		"&validateHiddenGroups=" + escape(frmValidate.validateHiddenGroups.value) +
		"&validateName=" + escape(frmValidate.validateName.value) +
		"&validateTimestamp=" + escape(frmValidate.validateTimestamp.value) +
		"&validateUtilID=" + escape(frmValidate.validateUtilID.value) +
		"&destination=util_validate_mailmerge";
	openDialog(sURL, (screen.width)/2,(screen.height)/3, "no", "no");
}
	function cancelClick()
{
	if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
		(definitionChanged() == false)) {
		//todo
		//window.location.href="defsel";
		return(false);
	}

	var answer = OpenHR.messageBox("You have changed the current definition. Save changes ?",3);
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
	OpenHR.disableMenu("menuframe");
	frmSend.txtSend_reaction.value = "MAILMERGE";
	submitDefinition();
}
	function saveChanges(psAction, pfPrompt, pfTBOverride)
{
	if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
		(definitionChanged() == false)) {
		return 7; //No to saving the changes, as none have been made.
	}
	var answer = OpenHR.messageBox("You have changed the current definition. Save changes ?",3);
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
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") 
	{
		return false;
	}

	if (frmUseful.txtChanged.value == 1) 
	{
		return true;
	}
	else {
		if (frmUseful.txtAction.value.toUpperCase() != "NEW") {
			// Compare the tab 1 controls with the original values.
			if (frmDefinition.txtName.value != frmOriginalDefinition.txtDefn_Name.value) 
			{
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
					else
					{
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
	function spinRecords(pfUp)
{
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
	function getTableName(piTableID)
{
	var i;
	var sTableName = new String("");
	
	var sReqdControlName = new String("txtTableName_");
	sReqdControlName = sReqdControlName.concat(piTableID);

	var dataCollection = frmTables.elements;
	if (dataCollection!=null) {
		for (i=0; i<dataCollection.length; i++)  {
			var sControlName = dataCollection.item(i).name;
					
			if (sControlName == sReqdControlName) {
				sTableName = dataCollection.item(i).value;
				break;
			}
		}
	}	

	return sTableName;
}
	function columnSwap(pfSelect)
{
	var i;
	var iColumnsSwapped;
	var sAddedCalcIDs;
	
	sAddedCalcIDs = "";
	iColumnsSwapped = 0;
	
	// Do nothing of the Add button is disabled (read-only mode).
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") return;
	
	if (pfSelect == true) {
		var grdFrom = frmDefinition.ssOleDBGridAvailableColumns;
		var grdTo = frmDefinition.ssOleDBGridSelectedColumns;
	}
	else {
		var grdFrom = frmDefinition.ssOleDBGridSelectedColumns;
		var grdTo = frmDefinition.ssOleDBGridAvailableColumns;

		// Check if the column being removed is in the sort columns collection.
		iCount = grdFrom.selbookmarks.Count();		
		for (i=iCount-1; i >= 0; i--) {
			grdFrom.bookmark = grdFrom.selbookmarks(i);
			iRowIndex = grdFrom.AddItemRowIndex(grdFrom.Bookmark);
					
			// Remove the column from the sort columns collection.
			if (grdFrom.columns(0).text == "C") {
				if (frmUseful.txtSortLoaded.value == 1) {
					if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
						frmDefinition.ssOleDBGridSortOrder.Redraw = false;
						frmDefinition.ssOleDBGridSortOrder.MoveFirst();

						iCount2 = frmDefinition.ssOleDBGridSortOrder.rows;
						for (i2=0;i2<iCount2; i2++) {	
							if (grdFrom.columns(2).text == frmDefinition.ssOleDBGridSortOrder.Columns("id").Text) {
								// The selected column is a sort column. Prompt the user to confirm the deselection.

								sColumnName = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
								if (iCount > 1 ) {
									iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the definition sort order.\n\nDo you still want to remove this column ?",3,"Mail Merge");
								}
								else {
									iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the definition sort order.\n\nDo you still want to remove this column ?",4,"Mail Merge");
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
					if (dataCollection!=null) {
						for (iIndex=0; iIndex<dataCollection.length; iIndex++) {
							sControlName = dataCollection.item(iIndex).name;
							sControlName = sControlName.substr(0, 19);
							if (sControlName == "txtReportDefnOrder_") {
								if (grdFrom.columns(2).text == sortColumnParameter(dataCollection.item(iIndex).value, "COLUMNID")) {
									// The selected column is a sort column. Prompt the user to confirm the deselection.
									sColumnName = grdFrom.columns(3).text;
									if (iCount > 1 ) {
										iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the definition sort order.\n\nDo you still want to remove this column ?",3,"Mail Merge");
									}
									else {
										iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the definition sort order.\n\nDo you still want to remove this column ?",4,"Mail Merge");
									}
								
									if (iResponse == 2) {
										// Cancel.
										frmDefinition.ssOleDBGridSortOrder.Redraw = true;
										return;
									}
								
									if (iResponse == 7)	{
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
		iHiddenCalcCount = 0;
		
		for (i=0; i < grdFrom.selbookmarks.Count(); i++) {
			grdFrom.bookmark = grdFrom.selbookmarks(i);

			// Check if the user is selecting a hidden calc, but is not the report owner.
			if ((pfSelect == true) &&
				(frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) &&
				(grdFrom.columns(6).text == "Y")) {
				
				sCalcName = new String(grdFrom.columns(3).text);
				iStringIndex = sCalcName.indexOf("<Calc> ");
				if (iStringIndex >= 0) {
					sCalcName = sCalcName.substring(iStringIndex + 7, sCalcName.length);
				}
				OpenHR.messageBox("Cannot include the '" + sCalcName + "' calculation.\nIts hidden and you are not the creator of this definition.",64,"Mail Merge");
			}
			else {	
				iColumnsSwapped = iColumnsSwapped + 1;
				
				if (grdFrom.columns(0).text == 'C') {
					sAddline = grdFrom.columns(0).text + 
								'	' + grdFrom.columns(1).text + 
								'	' + grdFrom.columns(2).text
						
					if (pfSelect == true) { 
						sAddline = sAddline + '	' + getTableName(grdFrom.columns(1).text) + '.' + grdFrom.columns(3).text 
					}
					else {
						sAddline = sAddline + '	' + grdFrom.columns(3).text.substring(grdFrom.columns(3).text.indexOf(".")+1) 							
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
						'	' +	grdFrom.columns(3).text +
						'	' + '0' + '	' + '0' + '	' + '0';

					// Remember which calcs we are adding to the report so that
					// we can get there return types below.						
					if (grdFrom.columns(0).text == "E") {
						sAddedCalcIDs = sAddedCalcIDs + grdFrom.columns(2).text + ",";				
					}	
				}
			
				if (grdFrom.columns(6).text == "Y")	{
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
					
					sFromType = grdFrom.columns(0).text;				
					sFromTableID = grdFrom.columns(1).text;				
					
					sTemp = grdFrom.columns(3).text;
					iTemp = sTemp.indexOf("<Calc> ");
					if (iTemp >= 0) {
						sTemp = sTemp.substring(iTemp + 7);
					}
					sFromDisplay = replace(sTemp,"_"," ");
					sFromDisplay = sFromDisplay.substring(sFromDisplay.indexOf(".")+1)
					sFromDisplay = sFromDisplay.toUpperCase();
				
					fIsFromTblAvailable = (sFromTableID == frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].value);
					
					fIsFromTypeAvailable = (((sFromType == "C") && (frmDefinition.optColumns.checked)) ||
						((sFromType == "E") && (frmDefinition.optCalc.checked)))
											
					fFound = true;

					if (fIsFromTblAvailable && fIsFromTypeAvailable) {
						fFound = false;
						grdTo.movefirst();
						grdTo.Redraw = true;
						for(i2=0; i2<grdTo.rows(); i2++) {
							grdTo.Redraw = false;
							
							sToType = grdTo.columns(0).text;				
							sToTableID = grdTo.columns(1).text;				
							sToDisplay = replace(grdTo.columns(3).text.toUpperCase(),"_"," ");				
							
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
								grdTo.additem(sAddline,i2);
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
			for (i=iCount-1; i >= 0; i--) {
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
		
			if (iHiddenCalcCount > 0 ) {
				iOldCalcCount = new Number(frmSelectionAccess.calcsHiddenCount.value);
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

		if(sAddedCalcIDs.length > 0) {
			// Get the return types of the added calcs.
			var frmGetDataForm = OpenHR.getForm("dataframe","frmGetData");
			frmGetDataForm.txtAction.value = "GETEXPRESSIONRETURNTYPES";
			frmGetDataForm.txtParam1.value = sAddedCalcIDs;
			data_refreshData();
			OpenHR.getForm("dataframe","")
		}
	}
	grdFrom.Redraw = true;
	grdTo.Redraw = true;
	refreshTab2Controls();
}
	function columnSwapAll(pfSelect) 
{
	var i;
	var iColumnsSwapped;
	var sAddedCalcIDs;
	
	sAddedCalcIDs = "";	
	iColumnsSwapped = 0;
	
	if (pfSelect == true) {
		var grdFrom = frmDefinition.ssOleDBGridAvailableColumns;
		var grdTo = frmDefinition.ssOleDBGridSelectedColumns;
	}
	else {
		if (frmUseful.txtSortLoaded.value == 1) 
		{
			iSortedColumnCount = frmDefinition.ssOleDBGridSortOrder.Rows;
		}
		else {
			iSortedColumnCount = 0;
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection!=null) {
				for (iIndex=0; iIndex<dataCollection.length; iIndex++)  {
					sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 19);
					if (sControlName == "txtReportDefnOrder_") {
						iSortedColumnCount = 1;
						break;
					}
				}	
			}
		}

		if (iSortedColumnCount > 0)
		{
			iAnswer = OpenHR.messageBox("Removing all columns will remove all sort order columns. \n Are you sure ?",36,"Mail Merge");
		}
		else {
			iAnswer = 6;
		}
  			
		if (iAnswer == 7)	{
			// cancel 
			return;
		}
	
		var grdFrom = frmDefinition.ssOleDBGridSelectedColumns;
		var grdTo = frmDefinition.ssOleDBGridAvailableColumns;
	}
	
	grdFrom.redraw = false;
	grdTo.redraw = false;
	
	grdTo.selbookmarks.removeall();
		
	iHiddenCalcCount = 0;

	grdFrom.movefirst();
	for (i=0; i < grdFrom.Rows(); i++) {
		if ((pfSelect == true) &&
			(frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) &&
			(grdFrom.columns(6).text == "Y")) {
			
			sCalcName = new String(grdFrom.columns(3).text);
			iStringIndex = sCalcName.indexOf("> ");
			if (iStringIndex >= 0) {
				sCalcName = sCalcName.substring(iStringIndex, sCalcName.length);
			}
			OpenHR.messageBox("Cannot include the '" + sCalcName + "' calculation.\nIts hidden and you are not the creator of this definition.",64,"Mail Merge");
		}
		else {
			iColumnsSwapped = iColumnsSwapped + 1;
				
			if (grdFrom.columns(0).text == 'C')	{
				sAddline = grdFrom.columns(0).text + 
							'	' + grdFrom.columns(1).text + 
							'	' + grdFrom.columns(2).text
						
				if (pfSelect == true)	{ 
					sAddline = sAddline + '	' + getTableName(grdFrom.columns(1).text) + '.' + grdFrom.columns(3).text 
				}
				else {
					sAddline = sAddline + '	' + grdFrom.columns(3).text.substring(grdFrom.columns(3).text.indexOf(".")+1) 							
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

			if (grdFrom.columns(6).text == "Y")  {
				iHiddenCalcCount = iHiddenCalcCount + 1;
			}

			if (pfSelect == true) {
				grdTo.additem(sAddline);
			}
			else {
				/* Find the right spot to add the row. */
				sFromType = grdFrom.columns(0).text;				
				sFromTableID = grdFrom.columns(1).text;				
					
				sTemp = grdFrom.columns(3).text;
				iTemp = sTemp.indexOf("<Calc> ");
				if (iTemp >= 0) {
					sTemp = sTemp.substring(iTemp + 7);
				}
				sFromDisplay = replace(sTemp,"_"," ");
				sFromDisplay = sFromDisplay.substring(sFromDisplay.indexOf(".")+1)
				sFromDisplay = sFromDisplay.toUpperCase();

				fIsFromTblAvailable = (sFromTableID == frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].value);

				fIsFromTypeAvailable = (((sFromType == "C") && (frmDefinition.optColumns.checked)) ||
					((sFromType == "E") && (frmDefinition.optCalc.checked)))
										
				fFound = true;

				if (fIsFromTblAvailable && fIsFromTypeAvailable) {
					fFound = false;
					grdTo.movefirst();
					grdTo.Redraw = true;
					/* TM 19/06/02 - Fault 4000 */
					for(i2=0; i2<grdTo.rows(); i2++) {
						grdTo.Redraw = false;
						
						sToType = grdTo.columns(0).text;
						sToTableID = grdTo.columns(1).text;	
						sToDisplay = replace(grdTo.columns(3).text.toUpperCase(),"_"," ");				
						
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
							
						if (fFound == true)	{
							grdTo.additem(sAddline,i2);
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

		if (iHiddenCalcCount > 0 ) {
			iOldCalcCount = new Number(frmSelectionAccess.calcsHiddenCount.value);
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

		if(sAddedCalcIDs.length > 0) {
			// Get the return types of the added calcs.
			var frmGetDataForm = OpenHR.getForm("dataframe","frmGetData");
			frmGetDataForm.txtAction.value = "GETEXPRESSIONRETURNTYPES";
			frmGetDataForm.txtParam1.value = sAddedCalcIDs;
			data_refreshData();
		}
	}
		
	refreshTab2Controls();
}
	function replace(sExpression, sFind, sReplace) {
	//gi (global search, ignore case)
	var re = new RegExp(sFind,"gi");
	sExpression = sExpression.replace(re, sReplace);
	return(sExpression);
}
	function trim(strInput) {
	if (strInput.length < 1){
		return "";
	}
		
	while (strInput.substr(strInput.length-1, 1) == " ") {
		strInput = strInput.substr(0, strInput.length - 1);
	}
	
	while (strInput.substr(0, 1) == " ") {
		strInput = strInput.substr(1, strInput.length);
	}
	
	return strInput;
}
	function columnMove(pfUp)
{
	if (pfUp == true) {
		iNewIndex = frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark) - 1;
		iOldIndex = iNewIndex + 2;
		iSelectIndex =iNewIndex;
	}
	else {
		iNewIndex = frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark) + 2;
		iOldIndex = iNewIndex - 2;
		iSelectIndex =iNewIndex - 1;
	}

	sAddline = frmDefinition.ssOleDBGridSelectedColumns.columns(0).text + 
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
	
	
	function locateRecord(psSearchFor) {
		var fFound;

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

	<%--ND Combos functions stuff was here--%>
	function util_def_mailmerge_addActiveXHandlers() {
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
		if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
			var dtTicker = new Date();
			var iThisTick = new Number(dtTicker.getTime());
			if (txtLastKeyFind.value.length > 0) {
				var iLastTick = new Number(txtTicker.value);
			} else {
				var iLastTick = new Number("0");
			}

			if (iThisTick > (iLastTick + 1500)) {
				var sFind = String.fromCharCode(iKeyAscii);
			} else {
				var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
			}

			txtTicker.value = iThisTick;
			txtLastKeyFind.value = sFind;

			locateRecord(sFind);
		}
	}
	//ssOleDBGridSelectedColumns handlers
	function ssOleDbGridSelectedColumnsRowColChange() {
		if (frmUseful.txtLockGridEvents.value != 1) {
			refreshTab2Controls();
		}
	}
	function ssOleDbGridSelectedColumnsDblClick() { columnSwap(false); }
	function ssOleDbGridSelectedColumnsSelChange() { refreshTab2Controls(); }
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
		if ((frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Asc') &&
			(frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Desc')) {
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
		if (grdAccess.AddItemRowIndex(grdAccess.Bookmark) == 0) and(grdAccess.Columns("Access").Text.length > 0);
		{
			ForceAccess(grdAccess, AccessCode(grdAccess.Columns("Access").Text));
			grdAccess.MoveFirst();
			grdAccess.Col = 1;
		}
		refreshTab1Controls();
	}
	function grdAccessGotFocus() { grdAccess.Col = 1; }
	function grdAccessRowColChange() {
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
		} else {
			grdAccess.Columns("Access").Style = 3; // 3 = Combo box
			grdAccess.Columns("Access").RemoveAll();
			grdAccess.Columns("Access").AddItem(AccessDescription("RW"));
			grdAccess.Columns("Access").AddItem(AccessDescription("RO"));
			grdAccess.Columns("Access").AddItem(AccessDescription("HD"));
		}

		grdAccess.Col = 1
	}
	function grdAccessRowLoaded() {
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
		} else {
			if (grdAccess.Columns("SysSecMgr").CellText(Bookmark) == "1") {
				grdAccess.Columns("GroupName").CellStyleSet("SysSecMgr");
				grdAccess.Columns("Access").CellStyleSet("SysSecMgr");
				grdAccess.ForeColor = "0";
			} else {
				grdAccess.ForeColor = "0";
			}
		}
	}
}

</script>

<%  
	Dim iVersionOneEnabled = 0
	Dim cmdVersionOneModule = CreateObject("ADODB.Command")
	cmdVersionOneModule.CommandText = "spASRIntActivateModule"
	cmdVersionOneModule.CommandType = 4	' Stored Procedure
	cmdVersionOneModule.ActiveConnection = Session("databaseConnection")
	cmdVersionOneModule.CommandTimeout = 300

	Dim prmModuleKey = cmdVersionOneModule.CreateParameter("moduleKey", 200, 1, 50)	'200=varchar, 1=input, 50=size
	cmdVersionOneModule.Parameters.Append(prmModuleKey)
	prmModuleKey.value = "VERSIONONE"

	Dim prmEnabled = cmdVersionOneModule.CreateParameter("enabled", 11, 2) ' 11=bit, 2=output
	cmdVersionOneModule.Parameters.Append(prmEnabled)

	Err.Number = 0
	cmdVersionOneModule.Execute()

	iVersionOneEnabled = CInt(cmdVersionOneModule.Parameters("enabled").Value)
	If iVersionOneEnabled < 0 Then
		iVersionOneEnabled = 1
	End If
	cmdVersionOneModule = Nothing
%>

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

<div <%=session("BodyTag")%>>
	<form id="frmDefinition" name="frmDefinition">

		<table valign="top" align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr height="5">
							<td colspan="3"></td>
						</tr>

						<tr height="10">
							<td width="10"></td>
							<td>
								<input type="button" value="Definition" id="btnTab1" name="btnTab1" class="btn btndisabled" disabled="disabled"
									onclick="displayPage(1)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Columns" id="btnTab2" name="btnTab2" class="btn"
									onclick="displayPage(2)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Sort Order" id="btnTab3" name="btnTab3" class="btn"
									onclick="displayPage(3)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Output" id="btnTab4" name="btnTab4" class="btn"
									onclick="displayPage(4)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
							<td width="10"></td>
						</tr>

						<tr height="10">
							<td colspan="3"></td>
						</tr>

						<tr>
							<td width="10"></td>
							<td>
								<!-- First tab -->
								<div id="div1">
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
														<tr height="10">
															<td width="5">&nbsp;</td>
															<td width="10">Name :</td>
															<td width="5">&nbsp;</td>
															<td>
																<input id="txtName" name="txtName" maxlength="50" style="WIDTH: 100%" class="text"
																	onkeyup="changeName()">
															</td>
															<td width="20">&nbsp;</td>
															<td width="10">Owner :</td>
															<td width="5">&nbsp;</td>
															<td width="40%">
																<input id="txtOwner" name="txtOwner" style="WIDTH: 100%" disabled="disabled" class="text textdisabled">
															</td>
															<td width="5">&nbsp;</td>
														</tr>

														<tr>
															<td colspan="9" height="5"></td>
														</tr>

														<tr height="60">
															<td width="5">&nbsp;</td>
															<td width="10" nowrap valign="top">Description :</td>
															<td width="5">&nbsp;</td>
															<td width="40%" rowspan="3">
																<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" height="0" maxlength="255"
																	onkeyup="changeDescription()"
																	onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}"
																	onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
																</textarea>
															</td>
															<td width="20" nowrap>&nbsp;</td>
															<td width="10" valign="top">Access :</td>
															<td width="5">&nbsp;</td>
															<td width="40%" rowspan="3" valign="top"></td>
															<td width="5">&nbsp;</td>
														</tr>

														<tr height="10">
															<td colspan="7">&nbsp;</td>
														</tr>

														<tr height="10">
															<td colspan="7">&nbsp;</td>
														</tr>

														<tr>
															<td colspan="9">
																<hr>
															</td>
														</tr>

														<tr height="10">
															<td width="5">&nbsp;</td>
															<td width="100" nowrap valign="top">Base Table :</td>
															<td width="5">&nbsp;</td>
															<td width="40%" valign="top">
																<select id="cboBaseTable" name="cboBaseTable" class="combo" style="WIDTH: 100%"
																	onchange="changeBaseTable()">
																</select>
															</td>
															<td width="20" nowrap>&nbsp;</td>
															<td width="10" valign="top">Records :</td>
															<td width="5">&nbsp;</td>
															<td width="40%">
																<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																	<tr>
																		<td width="5">
																			<input checked id="optRecordSelection1" name="optRecordSelection" type="radio"
																				onclick="changeBaseTableRecordOptions()"
																				onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																				onfocus="try{radio_onFocus(this);}catch(e){}"
																				onblur="try{radio_onBlur(this);}catch(e){}" />
																		</td>
																		<td width="5">&nbsp;</td>
																		<td width="30">
																			<label
																				tabindex="-1"
																				for="optRecordSelection1"
																				class="radio"
																				onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																				All
																			</label>
																		</td>
																		<td>&nbsp;</td>
																	</tr>
																	<tr>
																		<td width="5">
																			<input id="optRecordSelection2" name="optRecordSelection" type="radio"
																				onclick="changeBaseTableRecordOptions()"
																				onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																				onfocus="try{radio_onFocus(this);}catch(e){}"
																				onblur="try{radio_onBlur(this);}catch(e){}" />
																		</td>
																		<td width="5">&nbsp;</td>
																		<td width="20">
																			<label
																				tabindex="-1"
																				for="optRecordSelection2"
																				class="radio"
																				onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																				Picklist</label>
																		</td>
																		<td width="5">&nbsp;</td>
																		<td>
																			<input id="txtBasePicklist" name="txtBasePicklist" disabled="disabled" style="WIDTH: 100%" class="text textdisabled">
																		</td>
																		<td width="30">
																			<input id="cmdBasePicklist" name="cmdBasePicklist" style="WIDTH: 100%" type="button" value="..." class="btn"
																				onclick="selectRecordOption('base', 'picklist')"
																				onmouseover="try{button_onMouseOver(this);}catch(e){}"
																				onmouseout="try{button_onMouseOut(this);}catch(e){}"
																				onfocus="try{button_onFocus(this);}catch(e){}"
																				onblur="try{button_onBlur(this);}catch(e){}" />
																		</td>
																	</tr>
																	<tr>
																		<td width="5">
																			<input id="optRecordSelection3" name="optRecordSelection" type="radio"
																				onclick="changeBaseTableRecordOptions()"
																				onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																				onfocus="try{radio_onFocus(this);}catch(e){}"
																				onblur="try{radio_onBlur(this);}catch(e){}" />
																		</td>
																		<td width="5">&nbsp;</td>
																		<td width="20">
																			<label
																				tabindex="-1"
																				for="optRecordSelection3"
																				class="radio"
																				onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																				Filter</label>
																		</td>
																		<td width="5">&nbsp;</td>
																		<td>
																			<input id="txtBaseFilter" name="txtBaseFilter" disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
																		</td>
																		<td width="30">
																			<input id="cmdBaseFilter" name="cmdBaseFilter" style="WIDTH: 100%" type="button" value="..." class="btn"
																				onclick="selectRecordOption('base', 'filter')"
																				onmouseover="try{button_onMouseOver(this);}catch(e){}"
																				onmouseout="try{button_onMouseOut(this);}catch(e){}"
																				onfocus="try{button_onFocus(this);}catch(e){}"
																				onblur="try{button_onBlur(this);}catch(e){}" />
																		</td>
																	</tr>
																</table>
															</td>
															<td width="5">&nbsp;</td>
														</tr>

														<tr height="10">
															<td width="5">&nbsp;</td>
															<td width="90" nowrap>&nbsp;</td>
															<td width="5">&nbsp;</td>
															<td width="40%">
																<input id="txtParent1" name="txtParent1" style="WIDTH: 100%" disabled="disabled" class="text textdisabled" type="hidden">
															</td>
															<td width="20" nowrap>&nbsp;</td>
															<td width="10">&nbsp;</td>
															<td width="5">&nbsp;</td>
															<td width="40%">
																<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																	<tr>
																		<td>&nbsp;</td>
																		<td width="30">&nbsp;</td>
																	</tr>
																</table>
															</td>
															<td width="5">&nbsp;</td>
														</tr>

														<tr height="10">
															<td width="5">&nbsp;</td>
															<td width="90">&nbsp;</td>
															<td width="5">&nbsp;</td>
															<td width="40%">
																<input id="txtParent2" name="txtParent2" style="WIDTH: 100%" disabled="disabled" class="text textdisabled" type="hidden">
															</td>
															<td width="20" nowrap>&nbsp;</td>
															<td width="10">&nbsp;</td>
															<td width="5">&nbsp;</td>
															<td width="40%">
																<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																	<tr>
																		<td>&nbsp;
																		</td>
																		<td width="30">&nbsp;</td>
																	</tr>
																</table>
															</td>
															<td width="5">&nbsp;</td>
														</tr>
													</table>
												</table>
									</table>
								</div>

								<!-- Second tab -->
								<div id="div2" style="visibility: hidden; display: none">
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr height="5">
														<td colspan="7" height="5"></td>
													</tr>

													<tr height="5">
														<td width="5" height="5"></td>
														<td valign="top" height="5">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr height="5">
																	<td height="5" colspan="7" width="100%">
																		<select id="cboTblAvailable" name="cboTblAvailable" disabled="disabled" class="combo combodisabled" style="WIDTH: 100%; HEIGHT: 100%"
																			onchange="refreshAvailableColumns();">
																		</select>
																	</td>
																</tr>
																<tr height="10">
																	<td height="10" colspan="7" width="100%"></td>
																</tr>
																<tr height="5">
																	<td height="5"></td>
																	<td height="5">
																		<input id="optColumns" name="optAvailType" type="radio" checked disabled="disabled"
																			onclick="refreshAvailableColumns();"
																			onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																			onfocus="try{radio_onFocus(this);}catch(e){}"
																			onblur="try{radio_onBlur(this);}catch(e){}" />
																	</td>
																	<td height="5" width="5">
																		<label
																			tabindex="-1"
																			for="optColumns"
																			class="radio radiodisabled"
																			onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Columns</label>
																	</td>
																	<td width="5" height="5"></td>
																	<td height="5">
																		<input id="optCalc" name="optAvailType" type="radio" disabled="disabled"
																			onclick="refreshAvailableColumns();"
																			onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																			onfocus="try{radio_onFocus(this);}catch(e){}"
																			onblur="try{radio_onBlur(this);}catch(e){}" />
																	</td>
																	<td width="5" height="5">
																		<label
																			tabindex="-1"
																			for="optCalc"
																			class="radio radiodisabled"
																			onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Calculations</label>
																	</td>
																	<td height="5"></td>
<%--																	<tr height="10">
																		<td height="10" colspan="7" width="100%"></td>
																	</tr>--%>
																</tr>
															</table>
														</td>
														<td width="10"></td>
														<td width="5" nowrap></td>
														<td width="10"></td>
														<td rowspan="4" width="40%" height="100%">
															<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="100%" height="100%">
																		<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
																			codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
																			id="ssOleDBGridSelectedColumns"
																			name="ssOleDBGridSelectedColumns"
																			style="HEIGHT: 100%; LEFT: 0; TOP: 0; WIDTH: 100%" width="100%" height="100%">
																			<param name="ScrollBars" value="2">
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
																			<param name="Row.Count" value="0">
																			<param name="Col.Count" value="12">
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
																			<param name="Columns.Count" value="14">

																			<param name="Columns(0).Width" value="0">
																			<param name="Columns(0).Visible" value="0">
																			<param name="Columns(0).Columns.Count" value="1">
																			<param name="Columns(0).Caption" value="type">
																			<param name="Columns(0).Name" value="type">
																			<param name="Columns(0).Alignment" value="0">
																			<param name="Columns(0).CaptionAlignment" value="3">
																			<param name="Columns(0).Bound" value="0">
																			<param name="Columns(0).AllowSizing" value="1">
																			<param name="Columns(0).DataField" value="Column 0">
																			<param name="Columns(0).DataType" value="8">
																			<param name="Columns(0).Level" value="0">
																			<param name="Columns(0).NumberFormat" value="">
																			<param name="Columns(0).Case" value="0">
																			<param name="Columns(0).FieldLen" value="4096">
																			<param name="Columns(0).VertScrollBar" value="0">
																			<param name="Columns(0).Locked" value="0">
																			<param name="Columns(0).Style" value="0">
																			<param name="Columns(0).ButtonsAlways" value="0">
																			<param name="Columns(0).RowCount" value="0">
																			<param name="Columns(0).ColCount" value="1">
																			<param name="Columns(0).HasHeadForeColor" value="0">
																			<param name="Columns(0).HasHeadBackColor" value="0">
																			<param name="Columns(0).HasForeColor" value="0">
																			<param name="Columns(0).HasBackColor" value="0">
																			<param name="Columns(0).HeadForeColor" value="0">
																			<param name="Columns(0).HeadBackColor" value="0">
																			<param name="Columns(0).ForeColor" value="0">
																			<param name="Columns(0).BackColor" value="0">
																			<param name="Columns(0).HeadStyleSet" value="">
																			<param name="Columns(0).StyleSet" value="">
																			<param name="Columns(0).Nullable" value="1">
																			<param name="Columns(0).Mask" value="">
																			<param name="Columns(0).PromptInclude" value="0">
																			<param name="Columns(0).ClipMode" value="0">
																			<param name="Columns(0).PromptChar" value="95">

																			<param name="Columns(1).Width" value="0">
																			<param name="Columns(1).Visible" value="0">
																			<param name="Columns(1).Columns.Count" value="1">
																			<param name="Columns(1).Caption" value="tableID">
																			<param name="Columns(1).Name" value="tableID">
																			<param name="Columns(1).Alignment" value="0">
																			<param name="Columns(1).CaptionAlignment" value="3">
																			<param name="Columns(1).Bound" value="0">
																			<param name="Columns(1).AllowSizing" value="1">
																			<param name="Columns(1).DataField" value="Column 1">
																			<param name="Columns(1).DataType" value="8">
																			<param name="Columns(1).Level" value="0">
																			<param name="Columns(1).NumberFormat" value="">
																			<param name="Columns(1).Case" value="0">
																			<param name="Columns(1).FieldLen" value="4096">
																			<param name="Columns(1).VertScrollBar" value="0">
																			<param name="Columns(1).Locked" value="0">
																			<param name="Columns(1).Style" value="0">
																			<param name="Columns(1).ButtonsAlways" value="0">
																			<param name="Columns(1).RowCount" value="0">
																			<param name="Columns(1).ColCount" value="1">
																			<param name="Columns(1).HasHeadForeColor" value="0">
																			<param name="Columns(1).HasHeadBackColor" value="0">
																			<param name="Columns(1).HasForeColor" value="0">
																			<param name="Columns(1).HasBackColor" value="0">
																			<param name="Columns(1).HeadForeColor" value="0">
																			<param name="Columns(1).HeadBackColor" value="0">
																			<param name="Columns(1).ForeColor" value="0">
																			<param name="Columns(1).BackColor" value="0">
																			<param name="Columns(1).HeadStyleSet" value="">
																			<param name="Columns(1).StyleSet" value="">
																			<param name="Columns(1).Nullable" value="1">
																			<param name="Columns(1).Mask" value="">
																			<param name="Columns(1).PromptInclude" value="0">
																			<param name="Columns(1).ClipMode" value="0">
																			<param name="Columns(1).PromptChar" value="95">

																			<param name="Columns(2).Width" value="0">
																			<param name="Columns(2).Visible" value="0">
																			<param name="Columns(2).Columns.Count" value="1">
																			<param name="Columns(2).Caption" value="columnID">
																			<param name="Columns(2).Name" value="columnID">
																			<param name="Columns(2).Alignment" value="0">
																			<param name="Columns(2).CaptionAlignment" value="3">
																			<param name="Columns(2).Bound" value="0">
																			<param name="Columns(2).AllowSizing" value="1">
																			<param name="Columns(2).DataField" value="Column 2">
																			<param name="Columns(2).DataType" value="8">
																			<param name="Columns(2).Level" value="0">
																			<param name="Columns(2).NumberFormat" value="">
																			<param name="Columns(2).Case" value="0">
																			<param name="Columns(2).FieldLen" value="4096">
																			<param name="Columns(2).VertScrollBar" value="0">
																			<param name="Columns(2).Locked" value="0">
																			<param name="Columns(2).Style" value="0">
																			<param name="Columns(2).ButtonsAlways" value="0">
																			<param name="Columns(2).RowCount" value="0">
																			<param name="Columns(2).ColCount" value="1">
																			<param name="Columns(2).HasHeadForeColor" value="0">
																			<param name="Columns(2).HasHeadBackColor" value="0">
																			<param name="Columns(2).HasForeColor" value="0">
																			<param name="Columns(2).HasBackColor" value="0">
																			<param name="Columns(2).HeadForeColor" value="0">
																			<param name="Columns(2).HeadBackColor" value="0">
																			<param name="Columns(2).ForeColor" value="0">
																			<param name="Columns(2).BackColor" value="0">
																			<param name="Columns(2).HeadStyleSet" value="">
																			<param name="Columns(2).StyleSet" value="">
																			<param name="Columns(2).Nullable" value="1">
																			<param name="Columns(2).Mask" value="">
																			<param name="Columns(2).PromptInclude" value="0">
																			<param name="Columns(2).ClipMode" value="0">
																			<param name="Columns(2).PromptChar" value="95">

																			<param name="Columns(3).Width" value="100000">
																			<param name="Columns(3).Visible" value="-1">
																			<param name="Columns(3).Columns.Count" value="1">
																			<param name="Columns(3).Caption" value="Columns / Calculations Selected">
																			<param name="Columns(3).Name" value="display">
																			<param name="Columns(3).Alignment" value="0">
																			<param name="Columns(3).CaptionAlignment" value="3">
																			<param name="Columns(3).Bound" value="0">
																			<param name="Columns(3).AllowSizing" value="1">
																			<param name="Columns(3).DataField" value="Column 3">
																			<param name="Columns(3).DataType" value="8">
																			<param name="Columns(3).Level" value="0">
																			<param name="Columns(3).NumberFormat" value="">
																			<param name="Columns(3).Case" value="0">
																			<param name="Columns(3).FieldLen" value="4096">
																			<param name="Columns(3).VertScrollBar" value="0">
																			<param name="Columns(3).Locked" value="0">
																			<param name="Columns(3).Style" value="0">
																			<param name="Columns(3).ButtonsAlways" value="0">
																			<param name="Columns(3).RowCount" value="0">
																			<param name="Columns(3).ColCount" value="1">
																			<param name="Columns(3).HasHeadForeColor" value="0">
																			<param name="Columns(3).HasHeadBackColor" value="0">
																			<param name="Columns(3).HasForeColor" value="0">
																			<param name="Columns(3).HasBackColor" value="0">
																			<param name="Columns(3).HeadForeColor" value="0">
																			<param name="Columns(3).HeadBackColor" value="0">
																			<param name="Columns(3).ForeColor" value="0">
																			<param name="Columns(3).BackColor" value="0">
																			<param name="Columns(3).HeadStyleSet" value="">
																			<param name="Columns(3).StyleSet" value="">
																			<param name="Columns(3).Nullable" value="1">
																			<param name="Columns(3).Mask" value="">
																			<param name="Columns(3).PromptInclude" value="0">
																			<param name="Columns(3).ClipMode" value="0">
																			<param name="Columns(3).PromptChar" value="95">

																			<param name="Columns(4).Width" value="0">
																			<param name="Columns(4).Visible" value="0">
																			<param name="Columns(4).Columns.Count" value="1">
																			<param name="Columns(4).Caption" value="size">
																			<param name="Columns(4).Name" value="size">
																			<param name="Columns(4).Alignment" value="0">
																			<param name="Columns(4).CaptionAlignment" value="3">
																			<param name="Columns(4).Bound" value="0">
																			<param name="Columns(4).AllowSizing" value="1">
																			<param name="Columns(4).DataField" value="Column 4">
																			<param name="Columns(4).DataType" value="8">
																			<param name="Columns(4).Level" value="0">
																			<param name="Columns(4).NumberFormat" value="">
																			<param name="Columns(4).Case" value="0">
																			<param name="Columns(4).FieldLen" value="4096">
																			<param name="Columns(4).VertScrollBar" value="0">
																			<param name="Columns(4).Locked" value="0">
																			<param name="Columns(4).Style" value="0">
																			<param name="Columns(4).ButtonsAlways" value="0">
																			<param name="Columns(4).RowCount" value="0">
																			<param name="Columns(4).ColCount" value="1">
																			<param name="Columns(4).HasHeadForeColor" value="0">
																			<param name="Columns(4).HasHeadBackColor" value="0">
																			<param name="Columns(4).HasForeColor" value="0">
																			<param name="Columns(4).HasBackColor" value="0">
																			<param name="Columns(4).HeadForeColor" value="0">
																			<param name="Columns(4).HeadBackColor" value="0">
																			<param name="Columns(4).ForeColor" value="0">
																			<param name="Columns(4).BackColor" value="0">
																			<param name="Columns(4).HeadStyleSet" value="">
																			<param name="Columns(4).StyleSet" value="">
																			<param name="Columns(4).Nullable" value="1">
																			<param name="Columns(4).Mask" value="">
																			<param name="Columns(4).PromptInclude" value="0">
																			<param name="Columns(4).ClipMode" value="0">
																			<param name="Columns(4).PromptChar" value="95">

																			<param name="Columns(5).Width" value="0">
																			<param name="Columns(5).Visible" value="0">
																			<param name="Columns(5).Columns.Count" value="1">
																			<param name="Columns(5).Caption" value="decimals">
																			<param name="Columns(5).Name" value="decimals">
																			<param name="Columns(5).Alignment" value="0">
																			<param name="Columns(5).CaptionAlignment" value="3">
																			<param name="Columns(5).Bound" value="0">
																			<param name="Columns(5).AllowSizing" value="1">
																			<param name="Columns(5).DataField" value="Column 5">
																			<param name="Columns(5).DataType" value="8">
																			<param name="Columns(5).Level" value="0">
																			<param name="Columns(5).NumberFormat" value="">
																			<param name="Columns(5).Case" value="0">
																			<param name="Columns(5).FieldLen" value="4096">
																			<param name="Columns(5).VertScrollBar" value="0">
																			<param name="Columns(5).Locked" value="0">
																			<param name="Columns(5).Style" value="0">
																			<param name="Columns(5).ButtonsAlways" value="0">
																			<param name="Columns(5).RowCount" value="0">
																			<param name="Columns(5).ColCount" value="1">
																			<param name="Columns(5).HasHeadForeColor" value="0">
																			<param name="Columns(5).HasHeadBackColor" value="0">
																			<param name="Columns(5).HasForeColor" value="0">
																			<param name="Columns(5).HasBackColor" value="0">
																			<param name="Columns(5).HeadForeColor" value="0">
																			<param name="Columns(5).HeadBackColor" value="0">
																			<param name="Columns(5).ForeColor" value="0">
																			<param name="Columns(5).BackColor" value="0">
																			<param name="Columns(5).HeadStyleSet" value="">
																			<param name="Columns(5).StyleSet" value="">
																			<param name="Columns(5).Nullable" value="1">
																			<param name="Columns(5).Mask" value="">
																			<param name="Columns(5).PromptInclude" value="0">
																			<param name="Columns(5).ClipMode" value="0">
																			<param name="Columns(5).PromptChar" value="95">

																			<param name="Columns(6).Width" value="0">
																			<param name="Columns(6).Visible" value="0">
																			<param name="Columns(6).Columns.Count" value="1">
																			<param name="Columns(6).Caption" value="hidden">
																			<param name="Columns(6).Name" value="hidden">
																			<param name="Columns(6).Alignment" value="0">
																			<param name="Columns(6).CaptionAlignment" value="3">
																			<param name="Columns(6).Bound" value="0">
																			<param name="Columns(6).AllowSizing" value="1">
																			<param name="Columns(6).DataField" value="Column 6">
																			<param name="Columns(6).DataType" value="8">
																			<param name="Columns(6).Level" value="0">
																			<param name="Columns(6).NumberFormat" value="">
																			<param name="Columns(6).Case" value="0">
																			<param name="Columns(6).FieldLen" value="4096">
																			<param name="Columns(6).VertScrollBar" value="0">
																			<param name="Columns(6).Locked" value="0">
																			<param name="Columns(6).Style" value="0">
																			<param name="Columns(6).ButtonsAlways" value="0">
																			<param name="Columns(6).RowCount" value="0">
																			<param name="Columns(6).ColCount" value="1">
																			<param name="Columns(6).HasHeadForeColor" value="0">
																			<param name="Columns(6).HasHeadBackColor" value="0">
																			<param name="Columns(6).HasForeColor" value="0">
																			<param name="Columns(6).HasBackColor" value="0">
																			<param name="Columns(6).HeadForeColor" value="0">
																			<param name="Columns(6).HeadBackColor" value="0">
																			<param name="Columns(6).ForeColor" value="0">
																			<param name="Columns(6).BackColor" value="0">
																			<param name="Columns(6).HeadStyleSet" value="">
																			<param name="Columns(6).StyleSet" value="">
																			<param name="Columns(6).Nullable" value="1">
																			<param name="Columns(6).Mask" value="">
																			<param name="Columns(6).PromptInclude" value="0">
																			<param name="Columns(6).ClipMode" value="0">
																			<param name="Columns(6).PromptChar" value="95">

																			<param name="Columns(7).Width" value="0">
																			<param name="Columns(7).Visible" value="0">
																			<param name="Columns(7).Columns.Count" value="1">
																			<param name="Columns(7).Caption" value="numeric">
																			<param name="Columns(7).Name" value="numeric">
																			<param name="Columns(7).Alignment" value="0">
																			<param name="Columns(7).CaptionAlignment" value="3">
																			<param name="Columns(7).Bound" value="0">
																			<param name="Columns(7).AllowSizing" value="1">
																			<param name="Columns(7).DataField" value="Column 7">
																			<param name="Columns(7).DataType" value="8">
																			<param name="Columns(7).Level" value="0">
																			<param name="Columns(7).NumberFormat" value="">
																			<param name="Columns(7).Case" value="0">
																			<param name="Columns(7).FieldLen" value="4096">
																			<param name="Columns(7).VertScrollBar" value="0">
																			<param name="Columns(7).Locked" value="0">
																			<param name="Columns(7).Style" value="0">
																			<param name="Columns(7).ButtonsAlways" value="0">
																			<param name="Columns(7).RowCount" value="0">
																			<param name="Columns(7).ColCount" value="1">
																			<param name="Columns(7).HasHeadForeColor" value="0">
																			<param name="Columns(7).HasHeadBackColor" value="0">
																			<param name="Columns(7).HasForeColor" value="0">
																			<param name="Columns(7).HasBackColor" value="0">
																			<param name="Columns(7).HeadForeColor" value="0">
																			<param name="Columns(7).HeadBackColor" value="0">
																			<param name="Columns(7).ForeColor" value="0">
																			<param name="Columns(7).BackColor" value="0">
																			<param name="Columns(7).HeadStyleSet" value="">
																			<param name="Columns(7).StyleSet" value="">
																			<param name="Columns(7).Nullable" value="1">
																			<param name="Columns(7).Mask" value="">
																			<param name="Columns(7).PromptInclude" value="0">
																			<param name="Columns(7).ClipMode" value="0">
																			<param name="Columns(7).PromptChar" value="95">

																			<param name="Columns(8).Width" value="0">
																			<param name="Columns(8).Visible" value="0">
																			<param name="Columns(8).Columns.Count" value="1">
																			<param name="Columns(8).Caption" value="heading">
																			<param name="Columns(8).Name" value="heading">
																			<param name="Columns(8).Alignment" value="0">
																			<param name="Columns(8).CaptionAlignment" value="3">
																			<param name="Columns(8).Bound" value="0">
																			<param name="Columns(8).AllowSizing" value="1">
																			<param name="Columns(8).DataField" value="Column 8">
																			<param name="Columns(8).DataType" value="8">
																			<param name="Columns(8).Level" value="0">
																			<param name="Columns(8).NumberFormat" value="">
																			<param name="Columns(8).Case" value="0">
																			<param name="Columns(8).FieldLen" value="4096">
																			<param name="Columns(8).VertScrollBar" value="0">
																			<param name="Columns(8).Locked" value="0">
																			<param name="Columns(8).Style" value="0">
																			<param name="Columns(8).ButtonsAlways" value="0">
																			<param name="Columns(8).RowCount" value="0">
																			<param name="Columns(8).ColCount" value="1">
																			<param name="Columns(8).HasHeadForeColor" value="0">
																			<param name="Columns(8).HasHeadBackColor" value="0">
																			<param name="Columns(8).HasForeColor" value="0">
																			<param name="Columns(8).HasBackColor" value="0">
																			<param name="Columns(8).HeadForeColor" value="0">
																			<param name="Columns(8).HeadBackColor" value="0">
																			<param name="Columns(8).ForeColor" value="0">
																			<param name="Columns(8).BackColor" value="0">
																			<param name="Columns(8).HeadStyleSet" value="">
																			<param name="Columns(8).StyleSet" value="">
																			<param name="Columns(8).Nullable" value="1">
																			<param name="Columns(8).Mask" value="">
																			<param name="Columns(8).PromptInclude" value="0">
																			<param name="Columns(8).ClipMode" value="0">
																			<param name="Columns(8).PromptChar" value="95">

																			<param name="Columns(9).Width" value="0">
																			<param name="Columns(9).Visible" value="0">
																			<param name="Columns(9).Columns.Count" value="1">
																			<param name="Columns(9).Caption" value="average">
																			<param name="Columns(9).Name" value="average">
																			<param name="Columns(9).Alignment" value="0">
																			<param name="Columns(9).CaptionAlignment" value="3">
																			<param name="Columns(9).Bound" value="0">
																			<param name="Columns(9).AllowSizing" value="1">
																			<param name="Columns(9).DataField" value="Column 9">
																			<param name="Columns(9).DataType" value="8">
																			<param name="Columns(9).Level" value="0">
																			<param name="Columns(9).NumberFormat" value="">
																			<param name="Columns(9).Case" value="0">
																			<param name="Columns(9).FieldLen" value="4096">
																			<param name="Columns(9).VertScrollBar" value="0">
																			<param name="Columns(9).Locked" value="0">
																			<param name="Columns(9).Style" value="0">
																			<param name="Columns(9).ButtonsAlways" value="0">
																			<param name="Columns(9).RowCount" value="0">
																			<param name="Columns(9).ColCount" value="1">
																			<param name="Columns(9).HasHeadForeColor" value="0">
																			<param name="Columns(9).HasHeadBackColor" value="0">
																			<param name="Columns(9).HasForeColor" value="0">
																			<param name="Columns(9).HasBackColor" value="0">
																			<param name="Columns(9).HeadForeColor" value="0">
																			<param name="Columns(9).HeadBackColor" value="0">
																			<param name="Columns(9).ForeColor" value="0">
																			<param name="Columns(9).BackColor" value="0">
																			<param name="Columns(9).HeadStyleSet" value="">
																			<param name="Columns(9).StyleSet" value="">
																			<param name="Columns(9).Nullable" value="1">
																			<param name="Columns(9).Mask" value="">
																			<param name="Columns(9).PromptInclude" value="0">
																			<param name="Columns(9).ClipMode" value="0">
																			<param name="Columns(9).PromptChar" value="95">

																			<param name="Columns(10).Width" value="0">
																			<param name="Columns(10).Visible" value="0">
																			<param name="Columns(10).Columns.Count" value="1">
																			<param name="Columns(10).Caption" value="count">
																			<param name="Columns(10).Name" value="count">
																			<param name="Columns(10).Alignment" value="0">
																			<param name="Columns(10).CaptionAlignment" value="3">
																			<param name="Columns(10).Bound" value="0">
																			<param name="Columns(10).AllowSizing" value="1">
																			<param name="Columns(10).DataField" value="Column 10">
																			<param name="Columns(10).DataType" value="8">
																			<param name="Columns(10).Level" value="0">
																			<param name="Columns(10).NumberFormat" value="">
																			<param name="Columns(10).Case" value="0">
																			<param name="Columns(10).FieldLen" value="4096">
																			<param name="Columns(10).VertScrollBar" value="0">
																			<param name="Columns(10).Locked" value="0">
																			<param name="Columns(10).Style" value="0">
																			<param name="Columns(10).ButtonsAlways" value="0">
																			<param name="Columns(10).RowCount" value="0">
																			<param name="Columns(10).ColCount" value="1">
																			<param name="Columns(10).HasHeadForeColor" value="0">
																			<param name="Columns(10).HasHeadBackColor" value="0">
																			<param name="Columns(10).HasForeColor" value="0">
																			<param name="Columns(10).HasBackColor" value="0">
																			<param name="Columns(10).HeadForeColor" value="0">
																			<param name="Columns(10).HeadBackColor" value="0">
																			<param name="Columns(10).ForeColor" value="0">
																			<param name="Columns(10).BackColor" value="0">
																			<param name="Columns(10).HeadStyleSet" value="">
																			<param name="Columns(10).StyleSet" value="">
																			<param name="Columns(10).Nullable" value="1">
																			<param name="Columns(10).Mask" value="">
																			<param name="Columns(10).PromptInclude" value="0">
																			<param name="Columns(10).ClipMode" value="0">
																			<param name="Columns(10).PromptChar" value="95">

																			<param name="Columns(11).Width" value="0">
																			<param name="Columns(11).Visible" value="0">
																			<param name="Columns(11).Columns.Count" value="1">
																			<param name="Columns(11).Caption" value="total">
																			<param name="Columns(11).Name" value="total">
																			<param name="Columns(11).Alignment" value="0">
																			<param name="Columns(11).CaptionAlignment" value="3">
																			<param name="Columns(11).Bound" value="0">
																			<param name="Columns(11).AllowSizing" value="1">
																			<param name="Columns(11).DataField" value="Column 11">
																			<param name="Columns(11).DataType" value="8">
																			<param name="Columns(11).Level" value="0">
																			<param name="Columns(11).NumberFormat" value="">
																			<param name="Columns(11).Case" value="0">
																			<param name="Columns(11).FieldLen" value="4096">
																			<param name="Columns(11).VertScrollBar" value="0">
																			<param name="Columns(11).Locked" value="0">
																			<param name="Columns(11).Style" value="0">
																			<param name="Columns(11).ButtonsAlways" value="0">
																			<param name="Columns(11).RowCount" value="0">
																			<param name="Columns(11).ColCount" value="1">
																			<param name="Columns(11).HasHeadForeColor" value="0">
																			<param name="Columns(11).HasHeadBackColor" value="0">
																			<param name="Columns(11).HasForeColor" value="0">
																			<param name="Columns(11).HasBackColor" value="0">
																			<param name="Columns(11).HeadForeColor" value="0">
																			<param name="Columns(11).HeadBackColor" value="0">
																			<param name="Columns(11).ForeColor" value="0">
																			<param name="Columns(11).BackColor" value="0">
																			<param name="Columns(11).HeadStyleSet" value="">
																			<param name="Columns(11).StyleSet" value="">
																			<param name="Columns(11).Nullable" value="1">
																			<param name="Columns(11).Mask" value="">
																			<param name="Columns(11).PromptInclude" value="0">
																			<param name="Columns(11).ClipMode" value="0">
																			<param name="Columns(11).PromptChar" value="95">

																			<param name="Columns(12).Width" value="0">
																			<param name="Columns(12).Visible" value="0">
																			<param name="Columns(12).Columns.Count" value="1">
																			<param name="Columns(12).Caption" value="hidden">
																			<param name="Columns(12).Name" value="hidden">
																			<param name="Columns(12).Alignment" value="0">
																			<param name="Columns(12).CaptionAlignment" value="3">
																			<param name="Columns(12).Bound" value="0">
																			<param name="Columns(12).AllowSizing" value="1">
																			<param name="Columns(12).DataField" value="Column 12">
																			<param name="Columns(12).DataType" value="8">
																			<param name="Columns(12).Level" value="0">
																			<param name="Columns(12).NumberFormat" value="">
																			<param name="Columns(12).Case" value="0">
																			<param name="Columns(12).FieldLen" value="4096">
																			<param name="Columns(12).VertScrollBar" value="0">
																			<param name="Columns(12).Locked" value="0">
																			<param name="Columns(12).Style" value="0">
																			<param name="Columns(12).ButtonsAlways" value="0">
																			<param name="Columns(12).RowCount" value="0">
																			<param name="Columns(12).ColCount" value="1">
																			<param name="Columns(12).HasHeadForeColor" value="0">
																			<param name="Columns(12).HasHeadBackColor" value="0">
																			<param name="Columns(12).HasForeColor" value="0">
																			<param name="Columns(12).HasBackColor" value="0">
																			<param name="Columns(12).HeadForeColor" value="0">
																			<param name="Columns(12).HeadBackColor" value="0">
																			<param name="Columns(12).ForeColor" value="0">
																			<param name="Columns(12).BackColor" value="0">
																			<param name="Columns(12).HeadStyleSet" value="">
																			<param name="Columns(12).StyleSet" value="">
																			<param name="Columns(12).Nullable" value="1">
																			<param name="Columns(12).Mask" value="">
																			<param name="Columns(12).PromptInclude" value="0">
																			<param name="Columns(12).ClipMode" value="0">
																			<param name="Columns(12).PromptChar" value="95">

																			<param name="Columns(13).Width" value="0">
																			<param name="Columns(13).Visible" value="0">
																			<param name="Columns(13).Columns.Count" value="1">
																			<param name="Columns(13).Caption" value="GroupWithNext">
																			<param name="Columns(13).Name" value="GroupWithNext">
																			<param name="Columns(13).Alignment" value="0">
																			<param name="Columns(13).CaptionAlignment" value="3">
																			<param name="Columns(13).Bound" value="0">
																			<param name="Columns(13).AllowSizing" value="1">
																			<param name="Columns(13).DataField" value="Column 13">
																			<param name="Columns(13).DataType" value="8">
																			<param name="Columns(13).Level" value="0">
																			<param name="Columns(13).NumberFormat" value="">
																			<param name="Columns(13).Case" value="0">
																			<param name="Columns(13).FieldLen" value="4096">
																			<param name="Columns(13).VertScrollBar" value="0">
																			<param name="Columns(13).Locked" value="0">
																			<param name="Columns(13).Style" value="0">
																			<param name="Columns(13).ButtonsAlways" value="0">
																			<param name="Columns(13).RowCount" value="0">
																			<param name="Columns(13).ColCount" value="1">
																			<param name="Columns(13).HasHeadForeColor" value="0">
																			<param name="Columns(13).HasHeadBackColor" value="0">
																			<param name="Columns(13).HasForeColor" value="0">
																			<param name="Columns(13).HasBackColor" value="0">
																			<param name="Columns(13).HeadForeColor" value="0">
																			<param name="Columns(13).HeadBackColor" value="0">
																			<param name="Columns(13).ForeColor" value="0">
																			<param name="Columns(13).BackColor" value="0">
																			<param name="Columns(13).HeadStyleSet" value="">
																			<param name="Columns(13).StyleSet" value="">
																			<param name="Columns(13).Nullable" value="1">
																			<param name="Columns(13).Mask" value="">
																			<param name="Columns(13).PromptInclude" value="0">
																			<param name="Columns(13).ClipMode" value="0">
																			<param name="Columns(13).PromptChar" value="95">

																			<param name="UseDefaults" value="-1">
																			<param name="TabNavigation" value="1">
																			<param name="BatchUpdate" value="0">
																			<param name="_ExtentX" value="4577">
																			<param name="_ExtentY" value="8202">
																			<param name="_StockProps" value="79">
																			<param name="Caption" value="">
																			<param name="ForeColor" value="0">
																			<param name="BackColor" value="16777215">
																			<param name="Enabled" value="-1">
																			<param name="DataMember" value="">
																		</object>
																	</td>
																</tr>
															</table>
														</td>
														<td width="5"></td>
													</tr>

													<tr height="5">
														<td height="5" colspan="6"></td>
													</tr>

													<tr>
														<td width="5"></td>
														<td rowspan="4" width="40%" height="100%">
															<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
																codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
																id="ssOleDBGridAvailableColumns"
																name="ssOleDBGridAvailableColumns"
																style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%" width="100%" height="100%">
																<param name="ScrollBars" value="2">
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
																<param name="Row.Count" value="0">
																<param name="Col.Count" value="8">
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
																<param name="Columns.Count" value="8">

																<param name="Columns(0).Width" value="0">
																<param name="Columns(0).Visible" value="0">
																<param name="Columns(0).Columns.Count" value="1">
																<param name="Columns(0).Caption" value="type">
																<param name="Columns(0).Name" value="type">
																<param name="Columns(0).Alignment" value="0">
																<param name="Columns(0).CaptionAlignment" value="3">
																<param name="Columns(0).Bound" value="0">
																<param name="Columns(0).AllowSizing" value="1">
																<param name="Columns(0).DataField" value="Column 0">
																<param name="Columns(0).DataType" value="8">
																<param name="Columns(0).Level" value="0">
																<param name="Columns(0).NumberFormat" value="">
																<param name="Columns(0).Case" value="0">
																<param name="Columns(0).FieldLen" value="4096">
																<param name="Columns(0).VertScrollBar" value="0">
																<param name="Columns(0).Locked" value="0">
																<param name="Columns(0).Style" value="0">
																<param name="Columns(0).ButtonsAlways" value="0">
																<param name="Columns(0).RowCount" value="0">
																<param name="Columns(0).ColCount" value="1">
																<param name="Columns(0).HasHeadForeColor" value="0">
																<param name="Columns(0).HasHeadBackColor" value="0">
																<param name="Columns(0).HasForeColor" value="0">
																<param name="Columns(0).HasBackColor" value="0">
																<param name="Columns(0).HeadForeColor" value="0">
																<param name="Columns(0).HeadBackColor" value="0">
																<param name="Columns(0).ForeColor" value="0">
																<param name="Columns(0).BackColor" value="0">
																<param name="Columns(0).HeadStyleSet" value="">
																<param name="Columns(0).StyleSet" value="">
																<param name="Columns(0).Nullable" value="1">
																<param name="Columns(0).Mask" value="">
																<param name="Columns(0).PromptInclude" value="0">
																<param name="Columns(0).ClipMode" value="0">
																<param name="Columns(0).PromptChar" value="95">

																<param name="Columns(1).Width" value="0">
																<param name="Columns(1).Visible" value="0">
																<param name="Columns(1).Columns.Count" value="1">
																<param name="Columns(1).Caption" value="tableID">
																<param name="Columns(1).Name" value="tableID">
																<param name="Columns(1).Alignment" value="0">
																<param name="Columns(1).CaptionAlignment" value="3">
																<param name="Columns(1).Bound" value="0">
																<param name="Columns(1).AllowSizing" value="1">
																<param name="Columns(1).DataField" value="Column 1">
																<param name="Columns(1).DataType" value="8">
																<param name="Columns(1).Level" value="0">
																<param name="Columns(1).NumberFormat" value="">
																<param name="Columns(1).Case" value="0">
																<param name="Columns(1).FieldLen" value="4096">
																<param name="Columns(1).VertScrollBar" value="0">
																<param name="Columns(1).Locked" value="0">
																<param name="Columns(1).Style" value="0">
																<param name="Columns(1).ButtonsAlways" value="0">
																<param name="Columns(1).RowCount" value="0">
																<param name="Columns(1).ColCount" value="1">
																<param name="Columns(1).HasHeadForeColor" value="0">
																<param name="Columns(1).HasHeadBackColor" value="0">
																<param name="Columns(1).HasForeColor" value="0">
																<param name="Columns(1).HasBackColor" value="0">
																<param name="Columns(1).HeadForeColor" value="0">
																<param name="Columns(1).HeadBackColor" value="0">
																<param name="Columns(1).ForeColor" value="0">
																<param name="Columns(1).BackColor" value="0">
																<param name="Columns(1).HeadStyleSet" value="">
																<param name="Columns(1).StyleSet" value="">
																<param name="Columns(1).Nullable" value="1">
																<param name="Columns(1).Mask" value="">
																<param name="Columns(1).PromptInclude" value="0">
																<param name="Columns(1).ClipMode" value="0">
																<param name="Columns(1).PromptChar" value="95">

																<param name="Columns(2).Width" value="0">
																<param name="Columns(2).Visible" value="0">
																<param name="Columns(2).Columns.Count" value="1">
																<param name="Columns(2).Caption" value="columnID">
																<param name="Columns(2).Name" value="columnID">
																<param name="Columns(2).Alignment" value="0">
																<param name="Columns(2).CaptionAlignment" value="3">
																<param name="Columns(2).Bound" value="0">
																<param name="Columns(2).AllowSizing" value="1">
																<param name="Columns(2).DataField" value="Column 2">
																<param name="Columns(2).DataType" value="8">
																<param name="Columns(2).Level" value="0">
																<param name="Columns(2).NumberFormat" value="">
																<param name="Columns(2).Case" value="0">
																<param name="Columns(2).FieldLen" value="4096">
																<param name="Columns(2).VertScrollBar" value="0">
																<param name="Columns(2).Locked" value="0">
																<param name="Columns(2).Style" value="0">
																<param name="Columns(2).ButtonsAlways" value="0">
																<param name="Columns(2).RowCount" value="0">
																<param name="Columns(2).ColCount" value="1">
																<param name="Columns(2).HasHeadForeColor" value="0">
																<param name="Columns(2).HasHeadBackColor" value="0">
																<param name="Columns(2).HasForeColor" value="0">
																<param name="Columns(2).HasBackColor" value="0">
																<param name="Columns(2).HeadForeColor" value="0">
																<param name="Columns(2).HeadBackColor" value="0">
																<param name="Columns(2).ForeColor" value="0">
																<param name="Columns(2).BackColor" value="0">
																<param name="Columns(2).HeadStyleSet" value="">
																<param name="Columns(2).StyleSet" value="">
																<param name="Columns(2).Nullable" value="1">
																<param name="Columns(2).Mask" value="">
																<param name="Columns(2).PromptInclude" value="0">
																<param name="Columns(2).ClipMode" value="0">
																<param name="Columns(2).PromptChar" value="95">

																<param name="Columns(3).Width" value="100000">
																<param name="Columns(3).Visible" value="-1">
																<param name="Columns(3).Columns.Count" value="1">
																<param name="Columns(3).Caption" value="Columns / Calculations Available">
																<param name="Columns(3).Name" value="display">
																<param name="Columns(3).Alignment" value="0">
																<param name="Columns(3).CaptionAlignment" value="3">
																<param name="Columns(3).Bound" value="0">
																<param name="Columns(3).AllowSizing" value="1">
																<param name="Columns(3).DataField" value="Column 3">
																<param name="Columns(3).DataType" value="8">
																<param name="Columns(3).Level" value="0">
																<param name="Columns(3).NumberFormat" value="">
																<param name="Columns(3).Case" value="0">
																<param name="Columns(3).FieldLen" value="4096">
																<param name="Columns(3).VertScrollBar" value="0">
																<param name="Columns(3).Locked" value="0">
																<param name="Columns(3).Style" value="0">
																<param name="Columns(3).ButtonsAlways" value="0">
																<param name="Columns(3).RowCount" value="0">
																<param name="Columns(3).ColCount" value="1">
																<param name="Columns(3).HasHeadForeColor" value="0">
																<param name="Columns(3).HasHeadBackColor" value="0">
																<param name="Columns(3).HasForeColor" value="0">
																<param name="Columns(3).HasBackColor" value="0">
																<param name="Columns(3).HeadForeColor" value="0">
																<param name="Columns(3).HeadBackColor" value="0">
																<param name="Columns(3).ForeColor" value="0">
																<param name="Columns(3).BackColor" value="0">
																<param name="Columns(3).HeadStyleSet" value="">
																<param name="Columns(3).StyleSet" value="">
																<param name="Columns(3).Nullable" value="1">
																<param name="Columns(3).Mask" value="">
																<param name="Columns(3).PromptInclude" value="0">
																<param name="Columns(3).ClipMode" value="0">
																<param name="Columns(3).PromptChar" value="95">

																<param name="Columns(4).Width" value="0">
																<param name="Columns(4).Visible" value="0">
																<param name="Columns(4).Columns.Count" value="1">
																<param name="Columns(4).Caption" value="size">
																<param name="Columns(4).Name" value="size">
																<param name="Columns(4).Alignment" value="0">
																<param name="Columns(4).CaptionAlignment" value="3">
																<param name="Columns(4).Bound" value="0">
																<param name="Columns(4).AllowSizing" value="1">
																<param name="Columns(4).DataField" value="Column 4">
																<param name="Columns(4).DataType" value="8">
																<param name="Columns(4).Level" value="0">
																<param name="Columns(4).NumberFormat" value="">
																<param name="Columns(4).Case" value="0">
																<param name="Columns(4).FieldLen" value="4096">
																<param name="Columns(4).VertScrollBar" value="0">
																<param name="Columns(4).Locked" value="0">
																<param name="Columns(4).Style" value="0">
																<param name="Columns(4).ButtonsAlways" value="0">
																<param name="Columns(4).RowCount" value="0">
																<param name="Columns(4).ColCount" value="1">
																<param name="Columns(4).HasHeadForeColor" value="0">
																<param name="Columns(4).HasHeadBackColor" value="0">
																<param name="Columns(4).HasForeColor" value="0">
																<param name="Columns(4).HasBackColor" value="0">
																<param name="Columns(4).HeadForeColor" value="0">
																<param name="Columns(4).HeadBackColor" value="0">
																<param name="Columns(4).ForeColor" value="0">
																<param name="Columns(4).BackColor" value="0">
																<param name="Columns(4).HeadStyleSet" value="">
																<param name="Columns(4).StyleSet" value="">
																<param name="Columns(4).Nullable" value="1">
																<param name="Columns(4).Mask" value="">
																<param name="Columns(4).PromptInclude" value="0">
																<param name="Columns(4).ClipMode" value="0">
																<param name="Columns(4).PromptChar" value="95">

																<param name="Columns(5).Width" value="0">
																<param name="Columns(5).Visible" value="0">
																<param name="Columns(5).Columns.Count" value="1">
																<param name="Columns(5).Caption" value="decimals">
																<param name="Columns(5).Name" value="decimals">
																<param name="Columns(5).Alignment" value="0">
																<param name="Columns(5).CaptionAlignment" value="3">
																<param name="Columns(5).Bound" value="0">
																<param name="Columns(5).AllowSizing" value="1">
																<param name="Columns(5).DataField" value="Column 5">
																<param name="Columns(5).DataType" value="8">
																<param name="Columns(5).Level" value="0">
																<param name="Columns(5).NumberFormat" value="">
																<param name="Columns(5).Case" value="0">
																<param name="Columns(5).FieldLen" value="4096">
																<param name="Columns(5).VertScrollBar" value="0">
																<param name="Columns(5).Locked" value="0">
																<param name="Columns(5).Style" value="0">
																<param name="Columns(5).ButtonsAlways" value="0">
																<param name="Columns(5).RowCount" value="0">
																<param name="Columns(5).ColCount" value="1">
																<param name="Columns(5).HasHeadForeColor" value="0">
																<param name="Columns(5).HasHeadBackColor" value="0">
																<param name="Columns(5).HasForeColor" value="0">
																<param name="Columns(5).HasBackColor" value="0">
																<param name="Columns(5).HeadForeColor" value="0">
																<param name="Columns(5).HeadBackColor" value="0">
																<param name="Columns(5).ForeColor" value="0">
																<param name="Columns(5).BackColor" value="0">
																<param name="Columns(5).HeadStyleSet" value="">
																<param name="Columns(5).StyleSet" value="">
																<param name="Columns(5).Nullable" value="1">
																<param name="Columns(5).Mask" value="">
																<param name="Columns(5).PromptInclude" value="0">
																<param name="Columns(5).ClipMode" value="0">
																<param name="Columns(5).PromptChar" value="95">

																<param name="Columns(6).Width" value="0">
																<param name="Columns(6).Visible" value="0">
																<param name="Columns(6).Columns.Count" value="1">
																<param name="Columns(6).Caption" value="hidden">
																<param name="Columns(6).Name" value="hidden">
																<param name="Columns(6).Alignment" value="0">
																<param name="Columns(6).CaptionAlignment" value="3">
																<param name="Columns(6).Bound" value="0">
																<param name="Columns(6).AllowSizing" value="1">
																<param name="Columns(6).DataField" value="Column 6">
																<param name="Columns(6).DataType" value="8">
																<param name="Columns(6).Level" value="0">
																<param name="Columns(6).NumberFormat" value="">
																<param name="Columns(6).Case" value="0">
																<param name="Columns(6).FieldLen" value="4096">
																<param name="Columns(6).VertScrollBar" value="0">
																<param name="Columns(6).Locked" value="0">
																<param name="Columns(6).Style" value="0">
																<param name="Columns(6).ButtonsAlways" value="0">
																<param name="Columns(6).RowCount" value="0">
																<param name="Columns(6).ColCount" value="1">
																<param name="Columns(6).HasHeadForeColor" value="0">
																<param name="Columns(6).HasHeadBackColor" value="0">
																<param name="Columns(6).HasForeColor" value="0">
																<param name="Columns(6).HasBackColor" value="0">
																<param name="Columns(6).HeadForeColor" value="0">
																<param name="Columns(6).HeadBackColor" value="0">
																<param name="Columns(6).ForeColor" value="0">
																<param name="Columns(6).BackColor" value="0">
																<param name="Columns(6).HeadStyleSet" value="">
																<param name="Columns(6).StyleSet" value="">
																<param name="Columns(6).Nullable" value="1">
																<param name="Columns(6).Mask" value="">
																<param name="Columns(6).PromptInclude" value="0">
																<param name="Columns(6).ClipMode" value="0">
																<param name="Columns(6).PromptChar" value="95">

																<param name="Columns(7).Width" value="0">
																<param name="Columns(7).Visible" value="0">
																<param name="Columns(7).Columns.Count" value="1">
																<param name="Columns(7).Caption" value="numeric">
																<param name="Columns(7).Name" value="numeric">
																<param name="Columns(7).Alignment" value="0">
																<param name="Columns(7).CaptionAlignment" value="3">
																<param name="Columns(7).Bound" value="0">
																<param name="Columns(7).AllowSizing" value="1">
																<param name="Columns(7).DataField" value="Column 7">
																<param name="Columns(7).DataType" value="8">
																<param name="Columns(7).Level" value="0">
																<param name="Columns(7).NumberFormat" value="">
																<param name="Columns(7).Case" value="0">
																<param name="Columns(7).FieldLen" value="4096">
																<param name="Columns(7).VertScrollBar" value="0">
																<param name="Columns(7).Locked" value="0">
																<param name="Columns(7).Style" value="0">
																<param name="Columns(7).ButtonsAlways" value="0">
																<param name="Columns(7).RowCount" value="0">
																<param name="Columns(7).ColCount" value="1">
																<param name="Columns(7).HasHeadForeColor" value="0">
																<param name="Columns(7).HasHeadBackColor" value="0">
																<param name="Columns(7).HasForeColor" value="0">
																<param name="Columns(7).HasBackColor" value="0">
																<param name="Columns(7).HeadForeColor" value="0">
																<param name="Columns(7).HeadBackColor" value="0">
																<param name="Columns(7).ForeColor" value="0">
																<param name="Columns(7).BackColor" value="0">
																<param name="Columns(7).HeadStyleSet" value="">
																<param name="Columns(7).StyleSet" value="">
																<param name="Columns(7).Nullable" value="1">
																<param name="Columns(7).Mask" value="">
																<param name="Columns(7).PromptInclude" value="0">
																<param name="Columns(7).ClipMode" value="0">
																<param name="Columns(7).PromptChar" value="95">

																<param name="UseDefaults" value="-1">
																<param name="TabNavigation" value="1">
																<param name="BatchUpdate" value="0">
																<param name="_ExtentX" value="4577">
																<param name="_ExtentY" value="8202">
																<param name="_StockProps" value="79">
																<param name="Caption" value="">
																<param name="ForeColor" value="0">
																<param name="BackColor" value="16777215">
																<param name="Enabled" value="-1">
																<param name="DataMember" value="">
															</object>
														</td>
														<td width="10" nowrap></td>
														<td height="5" valign="top" align="center">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr height="25">
																	<td>&nbsp</td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnAdd" id="cmdColumnAdd" value="Add..." style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwap(true)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td>&nbsp</td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnAddAll" id="cmdColumnAddAll" value="Add All" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwapAll(true)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td></td>
																</tr>
																<tr height="15">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnRemove" id="cmdColumnRemove" value="Remove" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwap(false)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td></td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnRemoveAll" id="cmdColumnRemoveAll" value="Remove All" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwapAll(false)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td></td>
																</tr>
															</table>
														</td>
														<td width="10" nowrap></td>
														<td width="5"></td>
													</tr>

													<tr>
														<td colspan="5"></td>
													</tr>

													<tr height="5">
														<td colspan="6" height="5"></td>
													</tr>

													<tr height="5">
														<td width="5"></td>
														<td width="10"></td>
														<td width="80"></td>
														<td width="10"></td>
														<td valign="top">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="125">Size :</td>
																	<td width="5"></td>
																	<td>
																		<input id="txtSize" name="txtSize" maxlength="50" style="WIDTH: 100%" class="text"
																			onchange="validateColSize();"
																			onkeyup="validateColSize();">
																	</td>
																</tr>
																<tr>
																	<td width="125">Decimals :</td>
																	<td width="5"></td>
																	<td>
																		<input id="txtDecPlaces" name="txtDecPlaces" maxlength="50" style="WIDTH: 100%" class="text"
																			onchange="validateColDecimals();"
																			onkeyup="validateColDecimals();">
																	</td>
																</tr>
															</table>
														</td>
														<td width="5"></td>
													</tr>

													<tr height="5">
														<td colspan="7" height="5"></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>

								<!-- Third tab -->
								<div id="div3" style="visibility: hidden; display: none">
									<table width="100%" height="80%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="5" height="5"></td>
													</tr>

													<tr height="20">
														<td width="5">&nbsp;</td>
														<td colspan="3">Sort Order :</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td rowspan="12">
															<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
																codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
																height="100%"
																id="ssOleDBGridSortOrder"
																name="ssOleDBGridSortOrder"
																style="BACKGROUND-COLOR: threedface; HEIGHT: 100%; VISIBILITY: visible; WIDTH: 100%"
																width="100%">
																<param name="ScrollBars" value="2">
																<param name="_Version" value="196617">
																<param name="DataMode" value="2">
																<param name="Cols" value="0">
																<param name="Rows" value="0">
																<param name="BorderStyle" value="1">
																<param name="RecordSelectors" value="0">
																<param name="GroupHeaders" value="-1">
																<param name="ColumnHeaders" value="-1">
																<param name="GroupHeadLines" value="1">
																<param name="HeadLines" value="1">
																<param name="FieldDelimiter" value="(None)">
																<param name="FieldSeparator" value="(Tab)">
																<param name="Row.Count" value="0">
																<param name="Col.Count" value="1">
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
																<param name="AllowUpdate" value="-1">
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
																<param name="RowNavigation" value="2">
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
																<param name="Columns.Count" value="4">

																<param name="Columns(0).Width" value="3200">
																<param name="Columns(0).Visible" value="0">
																<param name="Columns(0).Columns.Count" value="1">
																<param name="Columns(0).Caption" value="id">
																<param name="Columns(0).Name" value="columnID">
																<param name="Columns(0).Alignment" value="0">
																<param name="Columns(0).CaptionAlignment" value="3">
																<param name="Columns(0).Bound" value="0">
																<param name="Columns(0).AllowSizing" value="1">
																<param name="Columns(0).DataField" value="Column 0">
																<param name="Columns(0).DataType" value="8">
																<param name="Columns(0).Level" value="0">
																<param name="Columns(0).NumberFormat" value="">
																<param name="Columns(0).Case" value="0">
																<param name="Columns(0).FieldLen" value="256">
																<param name="Columns(0).VertScrollBar" value="0">
																<param name="Columns(0).Locked" value="0">
																<param name="Columns(0).Style" value="0">
																<param name="Columns(0).ButtonsAlways" value="0">
																<param name="Columns(0).RowCount" value="0">
																<param name="Columns(0).ColCount" value="1">
																<param name="Columns(0).HasHeadForeColor" value="0">
																<param name="Columns(0).HasHeadBackColor" value="0">
																<param name="Columns(0).HasForeColor" value="0">
																<param name="Columns(0).HasBackColor" value="0">
																<param name="Columns(0).HeadForeColor" value="0">
																<param name="Columns(0).HeadBackColor" value="0">
																<param name="Columns(0).ForeColor" value="0">
																<param name="Columns(0).BackColor" value="0">
																<param name="Columns(0).HeadStyleSet" value="">
																<param name="Columns(0).StyleSet" value="">
																<param name="Columns(0).Nullable" value="1">
																<param name="Columns(0).Mask" value="">
																<param name="Columns(0).PromptInclude" value="0">
																<param name="Columns(0).ClipMode" value="0">
																<param name="Columns(0).PromptChar" value="95">

																<param name="Columns(1).Width" value="5292">
																<param name="Columns(1).Visible" value="-1">
																<param name="Columns(1).Columns.Count" value="1">
																<param name="Columns(1).Caption" value="Column">
																<param name="Columns(1).Name" value="column">
																<param name="Columns(1).Alignment" value="0">
																<param name="Columns(1).CaptionAlignment" value="3">
																<param name="Columns(1).Bound" value="0">
																<param name="Columns(1).AllowSizing" value="1">
																<param name="Columns(1).DataField" value="Column 1">
																<param name="Columns(1).DataType" value="8">
																<param name="Columns(1).Level" value="0">
																<param name="Columns(1).NumberFormat" value="">
																<param name="Columns(1).Case" value="0">
																<param name="Columns(1).FieldLen" value="256">
																<param name="Columns(1).VertScrollBar" value="0">
																<param name="Columns(1).Locked" value="1">
																<param name="Columns(1).Style" value="0">
																<param name="Columns(1).ButtonsAlways" value="0">
																<param name="Columns(1).RowCount" value="0">
																<param name="Columns(1).ColCount" value="1">
																<param name="Columns(1).HasHeadForeColor" value="0">
																<param name="Columns(1).HasHeadBackColor" value="0">
																<param name="Columns(1).HasForeColor" value="0">
																<param name="Columns(1).HasBackColor" value="0">
																<param name="Columns(1).HeadForeColor" value="0">
																<param name="Columns(1).HeadBackColor" value="0">
																<param name="Columns(1).ForeColor" value="0">
																<param name="Columns(1).BackColor" value="0">
																<param name="Columns(1).HeadStyleSet" value="">
																<param name="Columns(1).StyleSet" value="">
																<param name="Columns(1).Nullable" value="1">
																<param name="Columns(1).Mask" value="">
																<param name="Columns(1).PromptInclude" value="0">
																<param name="Columns(1).ClipMode" value="0">
																<param name="Columns(1).PromptChar" value="95">

																<param name="Columns(2).Width" value="3000">
																<param name="Columns(2).Visible" value="-1">
																<param name="Columns(2).Columns.Count" value="1">
																<param name="Columns(2).Caption" value="Sort Order">
																<param name="Columns(2).Name" value="order">
																<param name="Columns(2).Alignment" value="0">
																<param name="Columns(2).CaptionAlignment" value="3">
																<param name="Columns(2).Bound" value="0">
																<param name="Columns(2).AllowSizing" value="1">
																<param name="Columns(2).DataField" value="Column 2">
																<param name="Columns(2).DataType" value="8">
																<param name="Columns(2).Level" value="0">
																<param name="Columns(2).NumberFormat" value="">
																<param name="Columns(2).Case" value="0">
																<param name="Columns(2).FieldLen" value="256">
																<param name="Columns(2).VertScrollBar" value="0">
																<param name="Columns(2).Locked" value="-1">
																<param name="Columns(2).Style" value="3">
																<param name="Columns(2).ButtonsAlways" value="0">
																<param name="Columns(2).Row.Count" value="2">
																<param name="Columns(2).Col.Count" value="2">
																<param name="Columns(2).Row(0).Col(0)" value="Asc">
																<param name="Columns(2).Row(0).Col(1)" value="">
																<param name="Columns(2).Row(1).Col(0)" value="Desc">
																<param name="Columns(2).Row(1).Col(1)" value="">
																<param name="Columns(2).HasHeadForeColor" value="0">
																<param name="Columns(2).HasHeadBackColor" value="0">
																<param name="Columns(2).HasForeColor" value="0">
																<param name="Columns(2).HasBackColor" value="0">
																<param name="Columns(2).HeadForeColor" value="0">
																<param name="Columns(2).HeadBackColor" value="0">
																<param name="Columns(2).ForeColor" value="0">
																<param name="Columns(2).BackColor" value="0">
																<param name="Columns(2).HeadStyleSet" value="">
																<param name="Columns(2).StyleSet" value="">
																<param name="Columns(2).Nullable" value="1">
																<param name="Columns(2).Mask" value="">
																<param name="Columns(2).PromptInclude" value="0">
																<param name="Columns(2).ClipMode" value="0">
																<param name="Columns(2).PromptChar" value="95">

																<param name="Columns(3).Width" value="3200">
																<param name="Columns(3).Visible" value="0">
																<param name="Columns(3).Columns.Count" value="1">
																<param name="Columns(3).Caption" value="tableID">
																<param name="Columns(3).Name" value="tableID">
																<param name="Columns(3).Alignment" value="0">
																<param name="Columns(3).CaptionAlignment" value="3">
																<param name="Columns(3).Bound" value="0">
																<param name="Columns(3).AllowSizing" value="1">
																<param name="Columns(3).DataField" value="Column 7">
																<param name="Columns(3).DataType" value="8">
																<param name="Columns(3).Level" value="0">
																<param name="Columns(3).NumberFormat" value="">
																<param name="Columns(3).Case" value="0">
																<param name="Columns(3).FieldLen" value="256">
																<param name="Columns(3).VertScrollBar" value="0">
																<param name="Columns(3).Locked" value="0">
																<param name="Columns(3).Style" value="0">
																<param name="Columns(3).ButtonsAlways" value="0">
																<param name="Columns(3).RowCount" value="0">
																<param name="Columns(3).ColCount" value="1">
																<param name="Columns(3).HasHeadForeColor" value="0">
																<param name="Columns(3).HasHeadBackColor" value="0">
																<param name="Columns(3).HasForeColor" value="0">
																<param name="Columns(3).HasBackColor" value="0">
																<param name="Columns(3).HeadForeColor" value="0">
																<param name="Columns(3).HeadBackColor" value="0">
																<param name="Columns(3).ForeColor" value="0">
																<param name="Columns(3).BackColor" value="0">
																<param name="Columns(3).HeadStyleSet" value="">
																<param name="Columns(3).StyleSet" value="">
																<param name="Columns(3).Nullable" value="1">
																<param name="Columns(3).Mask" value="">
																<param name="Columns(3).PromptInclude" value="0">
																<param name="Columns(3).ClipMode" value="0">
																<param name="Columns(3).PromptChar" value="95">

																<param name="UseDefaults" value="-1">
																<param name="TabNavigation" value="1">
																<param name="BatchUpdate" value="0">
																<param name="_ExtentX" value="11298">
																<param name="_ExtentY" value="3969">
																<param name="_StockProps" value="79">
																<param name="Caption" value="">
																<param name="ForeColor" value="0">
																<param name="BackColor" value="-2147483633">
																<param name="Enabled" value="-1">
																<param name="DataMember" value="">
															</object>
														</td>

														<td width="10">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortAdd" name="cmdSortAdd" value="Add..." style="WIDTH: 100%" class="btn"
																onclick="sortAdd()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortEdit" name="cmdSortEdit" value="Edit..." style="WIDTH: 100%" class="btn"
																onclick="sortEdit()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortRemove" name="cmdSortRemove" value="Remove" style="WIDTH: 100%" class="btn"
																onclick="sortRemove()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>
													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortRemoveAll" name="cmdSortRemoveAll" value="Remove All" style="WIDTH: 100%" class="btn"
																onclick="sortRemoveAll()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortMoveUp" name="cmdSortMoveUp" value="Move Up" style="WIDTH: 100%" class="btn"
																onclick="sortMove(true)"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortMoveDown" name="cmdSortMoveDown" value="Move Down" style="WIDTH: 100%" class="btn"
																onclick="sortMove(false)"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="5" height="5"></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>

								<!-- Fourth tab -->
								<div id="div4" style="visibility: hidden; display: none">
									<table width="100%" height="80%" class="outline" cellspacing="0" cellpadding="0">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="4">

													<tr height="5">
														<td colspan="9"></td>
													</tr>

													<tr>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td nowrap width="100">Template :</td>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td width="20">
															<input id="txtTemplate" name="txtTemplate" style="width: 400px" class="text textdisabled" disabled="disabled">
														</td>
														<td width="30">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td>
																		<input type="button" value="..." id="cmdTemplateSelect" name="cmdTemplateSelect" style="WIDTH: 25px" class="btn"
																			onclick="TemplateSelect()"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td>
																		<input type="button" value="Clear" id="cmdTemplateClear" name="cmdTemplateClear" style="WIDTH: 50px" class="btn"
																			onclick="TemplateClear()"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																</tr>
															</table>
														</td>
														<td width="80">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
														<td nowrap>
															<input type="checkbox" id="chkPause" name="chkPause" tabindex="-1"
																onclick="changeTab4Control()"
																onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" /><label
																for="chkPause"
																class="checkbox"
																tabindex="0"
																onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Pause before mail merge</label>
														</td>
														<td width="100%">&nbsp;&nbsp;&nbsp;</td>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
													</tr>

													<tr>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td nowrap width="100"></td>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td width="420"></td>
														<td width="30"></td>
														<td width="80">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>

														<td nowrap>
															<input type="checkbox" id="chkSuppressBlanks" name="chkSuppressBlanks" tabindex="-1"
																onclick="changeTab4Control()"
																onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" /><label
																for="chkSuppressBlanks"
																class="checkbox"
																tabindex="0"
																onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Suppress blank lines</label>
														</td>

														<td colspan="2"></td>
													</tr>

													<tr height="5">
														<td></td>
														<td colspan="7">
															<hr>
														</td>
														<td></td>
													</tr>

												</table>

												<table width="100%" class="invisible" cellspacing="0" cellpadding="0" height="100%">

													<tr style="height: 100%">
														<td></td>
														<td colspan="6">

															<table style="width: 100%; height: 100%">
																<tr>
																	<td width="20">&nbsp;&nbsp;&nbsp;</td>
																	<td width="220px" valign="top">
																		<table style="vertical-align: text-top" class="outline" cellspacing="0" cellpadding="4" width="100%" height="200px">
																			<tr style="height: 20px">
																				<td colspan="4" align="left" style="vertical-align: text-top">Output Format :
																					<br>
																				</td>
																			</tr>

																			<tr style="height: 20px">
																				<td width="5" style="vertical-align: text-top">
																					<input checked id="optDestination0" name="optDestination" type="radio"
																						onclick="refreshDestination();changeTab4Control()"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="5">&nbsp;</td>
																				<td width="130px" style="vertical-align: text-top">
																					<label tabindex="-1"
																						for="optDestination0"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																						Word Document</label>
																				</td>
																				<td>&nbsp;</td>
																			</tr>
																			<tr style="height: 20px">
																				<td width="5">
																					<input id="optDestination1" name="optDestination" type="radio"
																						onclick="refreshDestination();changeTab4Control()"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="5">&nbsp;</td>
																				<td width="130">
																					<label tabindex="-1"
																						for="optDestination1"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																						Individual Emails</label>
																				</td>
																				<td width="5">&nbsp;</td>

																			</tr>
																			<%If iVersionOneEnabled = 0 Then%>
																			<tr style="height: 20px; visibility: hidden; display: none">
																				<%Else%>
																			<tr style="height: 20px; visibility: visible; display: block">
																				<%End If%>
																				<td width="5">
																					<input id="optDestination2" name="optDestination" type="radio"
																						onclick="refreshDestination();changeTab4Control()"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="5">&nbsp;</td>
																				<td nowrap>
																					<label tabindex="-1"
																						for="optDestination2"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																						Document Management</label>
																				</td>
																			</tr>
																			<tr></tr>
																		</table>


																		<td style="width: 5px"></td>

																	<td valign="top">
																		<table class="outline" cellspacing="0" cellpadding="4" style="width: 100%; height: 200px; vertical-align: top">
																			<tr style="height: 20px">
																				<td colspan="4" align="left">Output Destinations :
																					<br>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row1" id="row1">
																				<td width="150px" nowrap>Engine :</td>
																				<td width="5px"></td>
																				<td colspan="2">
																					<select id="cboDMEngine" name="cboDMEngine" style="WIDTH: 400px" class="combo"
																						onchange="changeTab4Control()">
																					</select>
																				</td>
																			</tr>


																			<tr style="height: 20px" name="row4" id="row4">
																				<td nowrap colspan="2">
																					<input type="checkbox" id="chkOutputScreen" name="chkOutputScreen" tabindex="-1"
																						onclick="changeTab4Control()"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkOutputScreen"
																						class="checkbox"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Display output on screen
																					</label>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row2" id="row2">
																				<td nowrap></td>
																				<td></td>
																				<td style="width: 30px" colspan="3">
																					<table class="invisible" cellspacing="0" cellpadding="0">
																						<tr>
																							<td width="20"></td>
																							<td style="padding-right: 0; vertical-align: middle"></td>
																							<td></td>
																						</tr>
																					</table>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row3" id="row3">
																				<td nowrap colspan="6"></td>
																			</tr>

																			<tr style="height: 20px" name="row5" id="row5">
																				<td nowrap>
																					<input type="checkbox" id="chkOutputPrinter" name="chkOutputPrinter" tabindex="-1"
																						onclick="chkOutputPrinter_Click();changeTab4Control()"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkOutputPrinter"
																						class="checkbox"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Send to printer
																					</label>
																				</td>
																				<td class="text">Printer location :</td>
																				<td colspan="2">
																					<select id="cboPrinterName" name="cboPrinterName" style="WIDTH: 400px" class="combo"
																						onchange="changeTab4Control()">
																					</select>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row6" id="row6">
																				<td nowrap>
																					<input type="checkbox" id="chkSave" name="chkSave" tabindex="-1"
																						onclick="chkSave_Click();changeTab4Control()"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkSave"
																						class="checkbox"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Save to file
																					</label>
																				</td>
																				<td class="text">File name :</td>
																				<td colspan="2">
																					<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																						<tr>
																							<td width="20">
																								<input id="  " name="txtSaveFile" style="WIDTH: 325px" disabled="disabled" class="text textdisabled">
																							</td>
																							<td width="20">
																								<input type="button" value="..." id="cmdSaveFile" name="cmdSaveFile" style="WIDTH: 25px" class="btn"
																									onclick="saveFile()"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
																							</td>
																							<td>
																								<input type="button" value="Clear" id="cmdClearFile" name="cmdClearFile" style="WIDTH: 50px" class="btn"
																									onclick="fileClear()"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
																							</td>
																						</tr>
																					</table>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row7" id="row7">
																				<td width="150px" nowrap>Email Address :</td>
																				<td width="5px"></td>
																				<td>
																					<select id="cboEmail" name="cboEmail" style="WIDTH: 400px" class="combo"
																						onchange="changeTab4Control()">
																					</select>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row8" id="row8">
																				<td nowrap>Subject :</td>
																				<td width="5px"></td>
																				<td colspan="2">
																					<input id="txtSubject" name="txtSubject" style="WIDTH: 400px" maxlength="255" class="text"
																						onkeyup="changeTab4Control()">
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row9" id="row9">
																				<td nowrap colspan="3">
																					<input type="checkbox" id="chkAttachment" name="chkAttachment" tabindex="-1"
																						onclick="chkAttachment_Click();changeTab4Control()"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkAttachment"
																						class="checkbox"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Send as attachment
																					</label>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row10" id="row10">
																				<td nowrap>Attach as :</td>
																				<td></td>
																				<td colspan="2">
																					<input id="txtAttachmentName" name="txtAttachmentName" maxlength="255" style="WIDTH: 400px" class="text"
																						onkeyup="changeTab4Control()" />
																				</td>
																			</tr>
																			<tr height="100%"></tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
													</tr>

												</table>
											</td>
										</tr>
									</table>
								</div>

							</td>
							<td width="10"></td>
						</tr>

						<tr height="10">
							<td colspan="3"></td>
						</tr>

						<tr height="10">
							<td width="10"></td>
							<td>
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td>&nbsp;</td>
										<td width="80">
											<input type="button" id="cmdOK" name="cmdOK" value="OK" style="WIDTH: 100%" class="btn"
												onclick="okClick()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="10"></td>
										<td width="80">
											<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" style="WIDTH: 100%" class="btn"
												onclick="cancelClick()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td width="10"></td>
						</tr>

						<tr height="5">
							<td colspan="3"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>

		<input type='hidden' id="txtBasePicklistID" name="txtBasePicklistID">
		<input type='hidden' id="txtBaseFilterID" name="txtBaseFilterID">

		<input type='hidden' id="txtParent1ID" name="txtParent1ID">
		<input type='hidden' id="txtParent2ID" name="txtParent2ID">
		<input type='hidden' id="txtParent1FilterID" name="txtParent1FilterID">
		<input type='hidden' id="txtParent1PicklistID" name="txtParent1PicklistID">
		<input type='hidden' id="txtParent2FilterID" name="txtParent2FilterID">
		<input type='hidden' id="txtParent2PicklistID" name="txtParent2PicklistID">

		<input type='hidden' id="txtChildFilterID" name="txtChildFilterID">

		<input type="hidden" id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
		<input type="hidden" id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
	</form>

	<form id="frmTables">
		<%
			Dim sErrorDescription = ""
	
			' Get the table records.
			Dim cmdTables = CreateObject("ADODB.Command")
			cmdTables.CommandText = "sp_ASRIntGetTablesInfo"
			cmdTables.CommandType = 4	' Stored Procedure
			cmdTables.ActiveConnection = Session("databaseConnection")

			Err.Number = 0
			Dim rstTablesInfo = cmdTables.Execute
			If (Err.Number <> 0) Then
				sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				Dim iCount = 0
				Do While Not rstTablesInfo.EOF
					Response.Write("<INPUT type='hidden' id=txtTableName_" & rstTablesInfo.fields("tableID").value & " name=txtTableName_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("tableName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableType_" & rstTablesInfo.fields("tableID").value & " name=txtTableType_" & rstTablesInfo.fields("tableID").value & " value=" & rstTablesInfo.fields("tableType").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenString").value & """>" & vbCrLf)
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

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
	</form>

	<form id="frmOriginalDefinition">
		<%
			Dim sErrMsg = ""
			Dim prmUtilID As Object
	
			If Session("action") <> "new" Then
				Dim cmdDefn = CreateObject("ADODB.Command")
				cmdDefn.CommandText = "sp_ASRIntGetMailMergeDefinition"
				cmdDefn.CommandType = 4	' Stored Procedure
				cmdDefn.ActiveConnection = Session("databaseConnection")
		
				prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1)
				' 3=integer, 1=input
				cmdDefn.Parameters.Append(prmUtilID)
				prmUtilID.value = CleanNumeric(Session("utilid"))

				Dim prmUser = cmdDefn.CreateParameter("user", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefn.Parameters.Append(prmUser)
				prmUser.value = Session("username")

				Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefn.Parameters.Append(prmAction)
				prmAction.value = Session("action")

				Dim prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmErrMsg)

				Dim prmName = cmdDefn.CreateParameter("name", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmName)

				Dim prmOwner = cmdDefn.CreateParameter("owner", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOwner)

				Dim prmDescription = cmdDefn.CreateParameter("description", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmDescription)

				Dim prmBaseTableID = cmdDefn.CreateParameter("baseTableID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmBaseTableID)

				Dim prmSelection = cmdDefn.CreateParameter("selection", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmSelection)

				Dim prmPicklistID = cmdDefn.CreateParameter("picklistID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmPicklistID)

				Dim prmPicklistName = cmdDefn.CreateParameter("picklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmPicklistName)

				Dim prmPicklistHidden = cmdDefn.CreateParameter("picklistHidden", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmPicklistHidden)

				Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmFilterID)

				Dim prmFilterName = cmdDefn.CreateParameter("filterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmFilterName)

				Dim prmFilterHidden = cmdDefn.CreateParameter("filterHidden", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmFilterHidden)

				Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmOutputFormat)

				Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputSave)

				Dim prmOutputFileName = cmdDefn.CreateParameter("outputFileName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputFileName)

				Dim prmEmailAddrID = cmdDefn.CreateParameter("EmailAddrID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmEmailAddrID)

				Dim prmEmailSubject = cmdDefn.CreateParameter("EmailSubject", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmEmailSubject)

				Dim prmTemplateFileName = cmdDefn.CreateParameter("TemplateFileName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmTemplateFileName)

				Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputScreen)

				Dim prmEmailAsAttachment = cmdDefn.CreateParameter("EmailAsAttachment", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmEmailAsAttachment)

				Dim prmEmailAttachmentName = cmdDefn.CreateParameter("EmailAttachmentName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmEmailAttachmentName)

				Dim prmSuppressBlanks = cmdDefn.CreateParameter("SuppressBlanks", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmSuppressBlanks)

				Dim prmPauseBeforeMerge = cmdDefn.CreateParameter("PauseBeforeMerge", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmPauseBeforeMerge)

				Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputPrinter)
		
				Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 255) '200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputPrinterName)

				Dim prmDocumentMapID = cmdDefn.CreateParameter("documentMapID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmDocumentMapID)

				Dim prmManualDocManHeader = cmdDefn.CreateParameter("manualDocManHeader", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmManualDocManHeader)

				Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2)	' 3=integer, 2=output
				cmdDefn.Parameters.Append(prmTimestamp)

				Dim prmWarningMsg = cmdDefn.CreateParameter("warningMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmWarningMsg)

				Err.Number = 0
				Dim rstDefinition = cmdDefn.Execute
				Dim iHiddenCalcCount = 0
		
				If (Err.Number <> 0) Then
					sErrMsg = CType(("'" & Session("utilname") & "' definition could not be read." & vbCrLf & FormatError(Err.Description)), String)
				Else
					If rstDefinition.state <> 0 Then
						' Read recordset values.
						Dim iCount = 0
						Do While Not rstDefinition.EOF
							iCount = iCount + 1
							If rstDefinition.fields("definitionType").value = "ORDER" Then
								Response.Write("<INPUT type='hidden' id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & rstDefinition.fields("definitionString").value & """>" & vbCrLf)
							Else
								Response.Write("<INPUT type='hidden' id=txtReportDefnColumn_" & iCount & " name=txtReportDefnColumn_" & iCount & " value=""" & Replace(rstDefinition.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
	
								' Check if the report column is a hidden calc.
								If rstDefinition.fields("hidden").value = "Y" Then
									iHiddenCalcCount = iHiddenCalcCount + 1
								End If
							End If
							rstDefinition.MoveNext()
						Loop

						' Release the ADO recordset object.
						rstDefinition.close()
					End If
					rstDefinition = Nothing
			
					' NB. IMPORTANT ADO NOTE.
					' When calling a stored procedure which returns a recordset AND has output parameters
					' you need to close the recordset and set it to nothing before using the output parameters. 
					If Len(cmdDefn.Parameters("errMsg").value) > 0 Then
						sErrMsg = CType(("'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").value), String)
					End If

					Response.Write("<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").value, """", "&quot;") & """>" & vbCrLf)
					'			Response.Write "<INPUT type='hidden' id=txtDefn_Access name=txtDefn_Access value=""" & cmdDefn.Parameters("access").value & """>" & vbcrlf
					Response.Write("<INPUT type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & cmdDefn.Parameters("baseTableID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Selection name=txtDefn_Selection value=" & cmdDefn.Parameters("Selection").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & cmdDefn.Parameters("picklistID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(cmdDefn.Parameters("picklistName").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & LCase(cmdDefn.Parameters("picklistHidden").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & cmdDefn.Parameters("filterID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(cmdDefn.Parameters("filterName").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & LCase(cmdDefn.Parameters("filterHidden").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & cmdDefn.Parameters("OutputFormat").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & LCase(cmdDefn.Parameters("OutputSave").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputFileName name=txtDefn_OutputFileName value=""" & cmdDefn.Parameters("OutputFileName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_EmailAddrID name=txtDefn_EmailAddrID value=" & cmdDefn.Parameters("EmailAddrID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_EmailSubject name=txtDefn_EmailSubject value=""" & Replace(cmdDefn.Parameters("EmailSubject").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_TemplateFileName name=txtDefn_TemplateFileName value=""" & cmdDefn.Parameters("TemplateFileName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & LCase(cmdDefn.Parameters("OutputScreen").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_EmailAsAttachment name=txtDefn_EmailAsAttachment value=" & Replace(LCase(cmdDefn.Parameters("EmailAsAttachment").value), """", "&quot;") & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_EmailAttachmentName name=txtDefn_EmailAttachmentName value=""" & cmdDefn.Parameters("EmailAttachmentName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_SuppressBlanks name=txtDefn_SuppressBlanks value=" & LCase(cmdDefn.Parameters("SuppressBlanks").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PauseBeforeMerge name=txtDefn_PauseBeforeMerge value=" & LCase(cmdDefn.Parameters("PauseBeforeMerge").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & LCase(cmdDefn.Parameters("OutputPrinter").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & cmdDefn.Parameters("OutputPrinterName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_DocumentMapID name=txtDefn_DocumentMapID value=" & cmdDefn.Parameters("DocumentMapID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_ManualDocManHeader name=txtDefn_ManualDocManHeader value=" & LCase(cmdDefn.Parameters("ManualDocManHeader").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Warning name=txtDefn_Warning value=""" & Replace(cmdDefn.Parameters("warningMsg").value, """", "&quot;") & """>" & vbCrLf)
				End If
    
				Dim fDocManagement = False
				Dim lngDocumentMapID = 0
				If cmdDefn.Parameters("DocumentMapID").value > 0 Then
					fDocManagement = True
					lngDocumentMapID = CInt(cmdDefn.Parameters("DocumentMapID").value)
				End If

				' Release the ADO command object.
				cmdDefn = Nothing

				If Len(sErrMsg) > 0 Then
					Session("confirmtext") = sErrMsg
					Session("confirmtitle") = "OpenHR Intranet"
					Session("followpage") = "defsel"
					Session("reaction") = "MAILMERGE"
					Response.Clear()
					Response.Redirect("confirmok")
				End If
    
				If fDocManagement = True Then
					' Get the Document Type 'Name' (only the ID is stored in the table)
					Dim cmdDocManRecords = CreateObject("ADODB.Command")
					cmdDocManRecords.CommandText = "spASRIntGetDocumentManagementTypes"
					cmdDocManRecords.CommandType = 4 ' Stored Procedure
					cmdDocManRecords.ActiveConnection = Session("databaseConnection")
					Err.Number = 0
					Dim rstDocManRecords = cmdDocManRecords.Execute
	    
					Dim lngCount = 1
					Do While Not rstDocManRecords.EOF
						If CInt(rstDocManRecords.Fields(0).Value) = lngDocumentMapID Then
							Response.Write("<INPUT type='hidden' id=txtDefn_DocumentMapName name=txtDefn_DocumentMapName value=""" & Replace(rstDocManRecords.Fields(1).Value, """", "&quot;") & """>" & vbCrLf)
						End If

						rstDocManRecords.MoveNext()
						lngCount = lngCount + 1
					Loop
        
					cmdDocManRecords = Nothing
				End If

			End If
		%>
	</form>

	<form id="frmAccess">
		<%
			sErrorDescription = ""
	
			' Get the table records.
			Dim cmdAccess = CreateObject("ADODB.Command")
			cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
			cmdAccess.CommandType = 4	' Stored Procedure
			cmdAccess.ActiveConnection = Session("databaseConnection")

			Dim prmUtilType = cmdAccess.CreateParameter("utilType", 3, 1)	' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmUtilType)
			prmUtilType.value = 9	' 9 = mail merge

			prmUtilID = cmdAccess.CreateParameter("utilID", 3, 1)	' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmUtilID)
			If UCase(Session("action")) = "NEW" Then
				prmUtilID.value = 0
			Else
				prmUtilID.value = CleanNumeric(Session("utilid"))
			End If

			Dim prmFromCopy = cmdAccess.CreateParameter("fromCopy", 3, 1)	' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmFromCopy)
			If UCase(Session("action")) = "COPY" Then
				prmFromCopy.value = 1
			Else
				prmFromCopy.value = 0
			End If

			Err.Number = 0
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

	<form id="frmUseful" name="frmUseful">
		<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
		<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
		<input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
		<input type="hidden" id="txtCurrentChildTableID" name="txtCurrentChildTableID" value="0">
		<input type="hidden" id="txtTablesChanged" name="txtTablesChanged">
		<input type="hidden" id="txtSelectedColumnsLoaded" name="txtSelectedColumnsLoaded" value="0">
		<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
		<input type="hidden" id="txtChanged" name="txtChanged" value="0">
		<input type="hidden" id="txtLockGridEvents" name="txtLockGridEvents" value="0">
		<input type="hidden" id="txtUtilID" name="txtUtilID" value='<% =session("utilid")%>'>
		<input type="hidden" id="txtEmailPermission" name="txtEmailPermission">
		<%
			Dim cmdDefinition = CreateObject("ADODB.Command")
			cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
			cmdDefinition.CommandType = 4	' Stored procedure.
			cmdDefinition.ActiveConnection = Session("databaseConnection")

			prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefinition.Parameters.Append(prmModuleKey)
			prmModuleKey.value = "MODULE_PERSONNEL"

			Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefinition.Parameters.Append(prmParameterKey)
			prmParameterKey.value = "Param_TablePersonnel"

			Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefinition.Parameters.Append(prmParameterValue)

			Err.Number = 0
			cmdDefinition.Execute()

			Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
			cmdDefinition = Nothing

			Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
		%>
	</form>

	<form id="frmValidate" name="frmValidate" target="validate" method="post" action="util_validate_mailmerge">
		<input type="hidden" id="validateBaseFilter" name="validateBaseFilter" value="0">
		<input type="hidden" id="validateBasePicklist" name="validateBasePicklist" value="0">
		<input type="hidden" id="validateCalcs" name="validateCalcs" value=''>
		<input type="hidden" id="validateHiddenGroups" name="validateHiddenGroups" value=''>
		<input type="hidden" id="validateName" name="validateName" value=''>
		<input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
		<input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
	</form>

	<form id="frmSend" name="frmSend" method="post" action="util_def_mailmerge_Submit">
		<input type="hidden" id="txtSend_ID" name="txtSend_ID">
		<input type="hidden" id="txtSend_name" name="txtSend_name">
		<input type="hidden" id="txtSend_description" name="txtSend_description">
		<input type="hidden" id="txtSend_baseTable" name="txtSend_baseTable">
		<input type="hidden" id="txtSend_selection" name="txtSend_selection">
		<input type="hidden" id="txtSend_picklist" name="txtSend_picklist">
		<input type="hidden" id="txtSend_filter" name="txtSend_filter">
		<input type="hidden" id="txtSend_outputformat" name="txtSend_outputformat">
		<input type="hidden" id="txtSend_outputsave" name="txtSend_outputsave">
		<input type="hidden" id="txtSend_outputfilename" name="txtSend_outputfilename">
		<input type="hidden" id="txtSend_emailaddrid" name="txtSend_emailaddrid">
		<input type="hidden" id="txtSend_emailsubject" name="txtSend_emailsubject">
		<input type="hidden" id="txtSend_templatefilename" name="txtSend_templatefilename">
		<input type="hidden" id="txtSend_outputscreen" name="txtSend_outputscreen">
		<input type="hidden" id="txtSend_access" name="txtSend_access">
		<input type="hidden" id="txtSend_userName" name="txtSend_userName">
		<input type="hidden" id="txtSend_emailasattachment" name="txtSend_emailasattachment">
		<input type="hidden" id="txtSend_emailattachmentname" name="txtSend_emailattachmentname">
		<input type="hidden" id="txtSend_suppressblanks" name="txtSend_suppressblanks" value="0">
		<input type="hidden" id="txtSend_pausebeforemerge" name="txtSend_pausebeforemerge" value="0">
		<input type="hidden" id="txtSend_outputprinter" name="txtSend_outputprinter">
		<input type="hidden" id="txtSend_outputprintername" name="txtSend_outputprintername">
		<input type="hidden" id="txtSend_documentmapid" name="txtSend_documentmapid">
		<input type="hidden" id="txtSend_manualdocmanheader" name="txtSend_manualdocmanheader">

		<input type="hidden" id="txtSend_columns" name="txtSend_columns">
		<input type="hidden" id="txtSend_columns2" name="txtSend_columns2">

		<input type="hidden" id="txtSend_reaction" name="txtSend_reaction">

		<input type="hidden" id="txtSend_jobsToHide" name="txtSend_jobsToHide">
		<input type="hidden" id="txtSend_jobsToHideGroups" name="txtSend_jobsToHideGroups">
	</form>

	<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection" method="post">
		<input type="hidden" id="recSelType" name="recSelType">
		<input type="hidden" id="recSelTableID" name="recSelTableID">
		<input type="hidden" id="recSelCurrentID" name="recSelCurrentID">
		<input type="hidden" id="recSelTable" name="recSelTable">
		<input type="hidden" id="recSelDefOwner" name="recSelDefOwner">
		<input type="hidden" id="recSelDefType" name="recSelDefType">
	</form>

	<form id="frmDocTypeSelection" name="frmDocTypeSelection" target="doctypeSelection" action="util_doctypeSelection" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="DocTypeSelCurrentID" name="DocTypeSelCurrentID">
	</form>

	<form id="frmSortOrder" name="frmSortOrder" action="util_sortorderselection" target="sortorderselection" method="post">
		<input type="hidden" id="txtSortInclude" name="txtSortInclude">
		<input type="hidden" id="txtSortExclude" name="txtSortExclude">
		<input type="hidden" id="txtSortEditing" name="txtSortEditing">
		<input type="hidden" id="txtSortColumnID" name="txtSortColumnID">
		<input type="hidden" id="txtSortColumnName" name="txtSortColumnName">
		<input type="hidden" id="txtSortOrder" name="txtSortOrder">
		<input type="hidden" id="txtSortBOC" name="txtSortBOC">
		<input type="hidden" id="txtSortPOC" name="txtSortPOC">
		<input type="hidden" id="txtSortVOC" name="txtSortVOC">
		<input type="hidden" id="txtSortSRV" name="txtSortSRV">
	</form>


	<form id="frmSelectionAccess" name="frmSelectionAccess">
		<input type="hidden" id="forcedHidden" name="forcedHidden" value="N">
		<input type="hidden" id="baseHidden" name="baseHidden" value="N">
		<input type="hidden" id="p1Hidden" name="p1Hidden" value="N">
		<input type="hidden" id="p2Hidden" name="p2Hidden" value="N">
		<input type="hidden" id="childHidden" name="childHidden" value="N">
		<input type="hidden" id="calcsHiddenCount" name="calcsHiddenCount" value="0">
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">
</div>

<script type="text/javascript">
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

				for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
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
				for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
					sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 20);
					if (sControlName == "txtReportDefnColumn_") {
						sGridColumnID = selectedColumnParameter(dataCollection.item(iIndex).value, "COLUMNID")
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
		for (i = 0; i < dataCollection.length; i++) {
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

	var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
	var dataCollection = frmUtilDefForm.elements;
	if (dataCollection != null) {

		for (i = 0; i < dataCollection.length; i++) {

			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 9);
			if (sControlName == "txtEmail_") {
				sControlName = dataCollection.item(i).value
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
	// frmSortOrder.txtSortBOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(3).text; 
	// frmSortOrder.txtSortPOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(4).text; 
	// frmSortOrder.txtSortVOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(5).text; 
	// frmSortOrder.txtSortSRV.value = frmDefinition.ssOleDBGridSortOrder.Columns(6).text;

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
			for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
				sControlName = dataCollection.item(iIndex).name;
				sControlName = sControlName.substr(0, 20);
				if (sControlName == "txtReportDefnColumn_") {
					sDefnString = new String(dataCollection.item(iIndex).value);

					if (sDefnString.length > 0) {
						sType = selectedColumnParameter(sDefnString, "TYPE");
						sColumnID = selectedColumnParameter(sDefnString, "COLUMNID");

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
	// frmSortOrder.txtSortBOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(3).text;
	// frmSortOrder.txtSortPOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(4).text;
	// frmSortOrder.txtSortVOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(5).text;
	// frmSortOrder.txtSortSRV.value = frmDefinition.ssOleDBGridSortOrder.Columns(6).text;

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
				sControlName = dataCollection.item(i).name;
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
	if (pfUp == true) {
		iNewIndex = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) - 1;
		iOldIndex = iNewIndex + 2;
		iSelectIndex = iNewIndex;
	} else {
		iNewIndex = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) + 2;
		iOldIndex = iNewIndex - 2;
		iSelectIndex = iNewIndex - 1;
	}

	sAddline = frmDefinition.ssOleDBGridSortOrder.columns(0).text +
		'	' + frmDefinition.ssOleDBGridSortOrder.columns(1).text +
		'	' + frmDefinition.ssOleDBGridSortOrder.columns(2).text +
		'	' + frmDefinition.ssOleDBGridSortOrder.columns(3).text

	// + 
	// '	' + frmDefinition.ssOleDBGridSortOrder.columns(4).text + 
	// '	' + frmDefinition.ssOleDBGridSortOrder.columns(5).text + 
	// '	' + frmDefinition.ssOleDBGridSortOrder.columns(6).text + 
	// '	' + frmDefinition.ssOleDBGridSortOrder.columns(7).text + 
	// '	' + frmDefinition.ssOleDBGridSortOrder.columns(8).text

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
	var i;
	var sErrMsg;
	var iIndex;
	var iCount;
	var sPage;
	var sBreak;
	var sDefn;
	var sControlName;

	sErrMsg = "";

	//check at least one column defined as sort order
	if (frmUseful.txtSortLoaded.value == 1) {
		if (frmDefinition.ssOleDBGridSortOrder.Rows == 0) {
			sErrMsg = "You must select at least 1 column to order the mail merge by";
		}
		// else {
		//	frmDefinition.ssOleDBGridSortOrder.movefirst();

		// check boc and poc not both selected
		//	for (i=0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) {
		//		if ((frmDefinition.ssOleDBGridSortOrder.Columns("break").Text == -1) && 
		//			(frmDefinition.ssOleDBGridSortOrder.Columns("page").Text == -1)) {
		//			sErrMsg = "You cannot select break on change and page on change for the same column.";
		//			break;
		//		}

		//		frmDefinition.ssOleDBGridSortOrder.movenext();
		//	}     
		//}
	} else {
		iCount = 0;
		var dataCollection = frmOriginalDefinition.elements;
		if (dataCollection != null) {
			for (i = 0; i < dataCollection.length; i++) {
				sControlName = dataCollection.item(i).name;
				sControlName = sControlName.substr(0, 19);
				if (sControlName == "txtReportDefnOrder_") {
					// sDefn = new String(dataCollection.item(i).value);
					// sPage = sortColumnParameter(sDefn, "POC");					
					// sBreak = sortColumnParameter(sDefn, "BOC")

					// if ((sBreak == "-1") && (sPage == "-1")) {
					//	sErrMsg = "You cannot select break on change and page on change for the same column.";
					//	break;
					//}

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

	sErrMsg = "";

	if (frmDefinition.txtTemplate.value == "") {
		sErrMsg = "No Template file selected.";
	}

	if (sErrMsg == "") {
		if (OpenHR.validateFilePath(frmDefinition.txtTemplate.value) == false) {
			sErrMsg = "Template file not found.";
		}
	}

	if (sErrMsg == "") {
		//if (frmDefinition.cboDestination.selectedIndex == 1) {
		if (frmDefinition.optDestination0.checked == true) {
			if (frmDefinition.chkSave.checked == true) {
				if (frmDefinition.txtSaveFile.value == "") {
					sErrMsg = "No save output file name selected."
				}
			}
		}
	}

	if (sErrMsg == "") {
		//if (frmDefinition.cboDestination.selectedIndex == 0) {
		if (frmDefinition.optDestination1.checked == true) {
			if (frmDefinition.cboEmail.selectedIndex == -1) {
				sErrMsg = "No email column selected."
			} else {
				if (frmDefinition.cboEmail.options[frmDefinition.cboEmail.selectedIndex].value == "") {
					sErrMsg = "No email column selected."
				}
			}
		}
	}

	if (sErrMsg == "") {
		sAttachmentName = new String(frmDefinition.txtAttachmentName.value);
		if ((sAttachmentName.indexOf("/") != -1) ||
			(sAttachmentName.indexOf(":") != -1) ||
			(sAttachmentName.indexOf("?") != -1) ||
			(sAttachmentName.indexOf(String.fromCharCode(34)) != -1) ||
			(sAttachmentName.indexOf("<") != -1) ||
			(sAttachmentName.indexOf(">") != -1) ||
			(sAttachmentName.indexOf("|") != -1) ||
			(sAttachmentName.indexOf("\\") != -1) ||
			(sAttachmentName.indexOf("*") != -1)) {
			sErrMsg = "The attachment file name can not contain any of the following characters:\n/ : ? " + String.fromCharCode(34) + " < > | \\ *"
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
	//frmSend.txtSend_output.value = frmDefinition.cboDestination.options(frmDefinition.cboDestination.options.selectedIndex).value;
	if (frmDefinition.optDestination0.checked == true) {
		frmSend.txtSend_outputformat.value = 0
	}
	if (frmDefinition.optDestination1.checked == true) {
		frmSend.txtSend_outputformat.value = 1
	}
	if (frmDefinition.optDestination2.checked == true) {
		frmSend.txtSend_outputformat.value = 2
	}

	frmSend.txtSend_outputsave.value = "0";
	frmSend.txtSend_outputfilename.value = "";
	frmSend.txtSend_outputscreen.value = "0";
	frmSend.txtSend_emailaddrid.value = "0";
	frmSend.txtSend_emailsubject.value = "";
	frmSend.txtSend_emailasattachment.value = "0";
	frmSend.txtSend_emailattachmentname.value = "";
	//	frmSend.txtSend_suppressblanks.value = "0";
	//	frmSend.txtSend_pausebeforemerge.value = "0";
	frmSend.txtSend_outputprinter.value = "0";
	frmSend.txtSend_outputprintername.value = "";
	frmSend.txtSend_documentmapid.value = "0";
	frmSend.txtSend_manualdocmanheader.value = "0";

	//if (frmDefinition.cboDestination.options(frmDefinition.cboDestination.options.selectedIndex).value == 0) {
	if (frmDefinition.optDestination0.checked == true) {
		if (frmDefinition.chkSave.checked == true) {
			frmSend.txtSend_outputsave.value = "1";
			frmSend.txtSend_outputfilename.value = frmDefinition.txtSaveFile.value;
			//MH20050110 Fault 9410
			//			if (frmDefinition.chkSave.checked == true) {
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

	//if (frmDefinition.cboDestination.options(frmDefinition.cboDestination.options.selectedIndex).value == 2) {
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

		for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
			var iNum = new Number(i + 1)
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
	var blnChecked;

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
	fRecordOptionSet = false;

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

	//	for (i=0; i<frmDefinition.cboDestination.options.length; i++)  {
	//		if (frmDefinition.cboDestination.options(i).value == frmOriginalDefinition.txtDefn_Output.value) {
	//			frmDefinition.cboDestination.selectedIndex = i;
	//			break;
	//		}
	//  }

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
				sControlName = dataCollection.item(iIndex).name;
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

	for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
		if (frmDefinition.ssOleDBGridSelectedColumns.columns(0).text == "E") {
			sAddedCalcIDs = sAddedCalcIDs + frmDefinition.ssOleDBGridSelectedColumns.columns(2).text + ",";
		}
		frmDefinition.ssOleDBGridSelectedColumns.movenext();
	}

	frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;

	var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
	frmGetDataForm.txtAction.value = "GETEXPRESSIONRETURNTYPES";
	frmGetDataForm.txtParam1.value = sAddedCalcIDs;
	data_refreshData();

}
function loadSortDefinition() {
	var iIndex;

	if (frmUseful.txtSortLoaded.value == 0) {
		frmDefinition.ssOleDBGridSortOrder.focus();

		var dataCollection = frmOriginalDefinition.elements;
		if (dataCollection != null) {
			for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
				sControlName = dataCollection.item(iIndex).name;
				sControlName = sControlName.substr(0, 19);
				if (sControlName == "txtReportDefnOrder_") {
					sDefnString = new String(dataCollection.item(iIndex).value);

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
	dialog.CancelError = true;
	dialog.FileName = frmDefinition.txtSaveFile.value;
	dialog.DialogTitle = "Mail Merge Output Document";
	//dialog.Filter = "Word Document (*.doc)|*.doc";
	dialog.Filter = frmDefinition.txtWordFormats.value;
	dialog.FilterIndex = frmDefinition.txtWordFormatDefaultIndex.value;
	dialog.Flags = 2621446;

	try {
		dialog.ShowSave();
	} catch (e) {
	}

	if (dialog.FileName.length > 256) {
		OpenHR.messageBox("Path and file name must not exceed 256 characters in length");
		return;
	}

	frmDefinition.txtSaveFile.value = dialog.FileName;
	frmUseful.txtChanged.value = 1;
	refreshTab4Controls();
}
function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll) {
	dlgwinprops = "center:yes;" +
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
	var sTemp = '';

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
				sControlName = dataCollection.item(i).name;
				sControlName = sControlName.substr(0, 19);
				if (sControlName == "txtReportDefnOrder_") {
					iNum = iNum + 1;
					sDefn = new String(dataCollection.item(i).value);

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

		data_refreshData();

		frmUseful.txtTablesChanged.value = 0;
	}

	sSelectedIDs = selectedIDs();

	frmDefinition.ssOleDBGridAvailableColumns.RemoveAll();

	var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
	var dataCollection = frmUtilDefForm.elements;

	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {

			sControlName = dataCollection.item(i).name;
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

	var frmExprTypeForm = window.parent.frames("dataframe").document.forms("frmData");
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
	data_refreshData();

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
	var i;

	row1.style.visibility = "hidden";
	row1.style.display = "none";
	row2.style.visibility = "hidden";
	row2.style.display = "none";
	row3.style.visibility = "hidden";
	row3.style.display = "none";
	row4.style.visibility = "hidden";
	row4.style.display = "none";
	row5.style.visibility = "hidden";
	row5.style.display = "none";
	row6.style.visibility = "hidden";
	row6.style.display = "none";
	row7.style.visibility = "hidden";
	row7.style.display = "none";
	row8.style.visibility = "hidden";
	row8.style.display = "none";
	row9.style.visibility = "hidden";
	row9.style.display = "none";
	row10.style.visibility = "hidden";
	row10.style.display = "none";

	if (frmDefinition.optDestination0.checked == true) {
		row4.style.visibility = "visible";
		row4.style.display = "block";
		row5.style.visibility = "visible";
		row5.style.display = "block";
		row6.style.visibility = "visible";
		row6.style.display = "block";

		if (frmUseful.txtLoading.value == 'N') {
			frmDefinition.chkOutputScreen.checked = false;
			frmDefinition.chkOutputPrinter.checked = false;
			frmDefinition.chkSave.checked = false;
		}
		chkSave_Click();
		chkOutputPrinter_Click();
	} else if (frmDefinition.optDestination1.checked == true) {
		row7.style.visibility = "visible";
		row7.style.display = "block";
		row8.style.visibility = "visible";
		row8.style.display = "block";
		row9.style.visibility = "visible";
		row9.style.display = "block";
		row10.style.visibility = "visible";
		row10.style.display = "block";

		if (frmUseful.txtLoading.value == 'N') {
			GetEmailDefs();
			frmDefinition.txtSubject.value = "";
			frmDefinition.chkAttachment.checked = false;
		}
		chkAttachment_Click();
	} else if (frmDefinition.optDestination2.checked == true) {
		row1.style.visibility = "visible";
		row1.style.display = "block";
		row2.style.visibility = "visible";
		row2.style.display = "block";
		row3.style.visibility = "visible";
		row3.style.display = "block";
		row4.style.visibility = "visible";
		row4.style.display = "block";
	}

	// Get menu.asp to refresh the menu.
	menu_refreshMenu();
}
function populatePrinters() {
	with (frmDefinition.cboPrinterName) {
		var strCurrentPrinter = '';
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
				selectedIndex = iLoop + 1
			}
		}
	}
}
function populateDMEngine() {
	with (frmDefinition.cboDMEngine) {
		var strCurrentDMEngine = '';
		if (selectedIndex > 0) {
			strCurrentDMEngine = options[selectedIndex].innerText;
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

			if (oOption.innerText == strCurrentDMEngine) {
				selectedIndex = iLoop + 1
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
	frmGetDataForm.txtReportBaseTableID.value = frmUseful.txtCurrentBaseTableID.value
	frmGetDataForm.txtReportParent1TableID.value = 0;
	frmGetDataForm.txtReportParent2TableID.value = 0;
	frmGetDataForm.txtReportChildTableID.value = 0;
	data_refreshData();

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
		//frmDefinition.chkOutputScreen.checked = false;
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

</script>
<script type="text/javascript">
	util_def_mailmerge_window_onload();
	util_def_mailmerge_addActiveXHandlers();
</script>
