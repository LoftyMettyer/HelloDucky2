
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
	
	function util_def_customreports_onload() {
		
		var fOk;
		fOk = true;

		var sErrMsg = frmUseful.txtErrorDescription.value;
		if (sErrMsg.length > 0) 
		{
			fOk = false;
			OpenHR.messageBox(sErrMsg,48,"Custom Reports");
			//TODO
			//window.parent.location.replace("login");
		}
	
		if (fOk == true) 
		{
			setGridFont(frmDefinition.grdAccess);

			// Expand the work frame and hide the option frame.
			//TODO
			//window.parent.document.all.item("workframeset").cols = "*, 0";	
	
			populateBaseTableCombo();
		
			if (frmUseful.txtAction.value.toUpperCase() == "NEW"){
				frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
				setBaseTable(0);
				changeBaseTable();	
				frmUseful.txtSelectedColumnsLoaded.value = 1;	
				frmUseful.txtSortLoaded.value = 1;	
				frmDefinition.txtDescription.value = "";
			}
			else 
			{
				loadDefinition();
			}			
		
			setGridFont(frmDefinition.ssOleDBGridAvailableColumns);

			setGridFont(frmDefinition.ssOleDBGridChildren);
			grid_disable(frmDefinition.ssOleDBGridChildren, true);
		
			setGridFont(frmDefinition.ssOleDBGridSelectedColumns);
			setGridFont(frmDefinition.ssOleDBGridSortOrder);
			setGridFont(frmDefinition.ssOleDBGridRepetition);
		
			populateAccessGrid();
		
			if (frmUseful.txtAction.value.toUpperCase() != "EDIT")
			{
				frmUseful.txtUtilID.value = 0;
			}
		
			if (frmUseful.txtAction.value.toUpperCase() == "EDIT")
			{
				// Get the columns/calcs for the current table selection.
				var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");

				frmGetDataForm.txtReportBaseTableID.value = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;
				frmGetDataForm.txtReportParent1TableID.value = frmDefinition.txtParent1ID.value;
				frmGetDataForm.txtReportParent2TableID.value = frmDefinition.txtParent2ID.value;
				// assign the tab delimited string of selected child table ids 
				frmGetDataForm.txtReportChildTableID.value = childTableString();
			}
			
			if (frmUseful.txtAction.value.toUpperCase() == "COPY")
			{
				frmUseful.txtChanged.value = 1;
			}
			displayPage(1);

			refreshTab5Controls();
		
			frmUseful.txtLoading.value = 'N';

			// Get menu.asp to refresh the menu.
			OpenHR.refreshMenu();		

			if ((frmUseful.txtAction.value.toUpperCase() != "NEW") &&
				(frmOriginalDefinition.txtDefn_Info.value != "")) 
			{
				OpenHR.messageBox(frmOriginalDefinition.txtDefn_Info.value,48,"Custom Reports");
				frmUseful.txtChanged.value = 1;
				refreshTab1Controls();
			}

			if (frmDefinition.chkDestination1.checked == true) 
			{
				if (frmOriginalDefinition.txtDefn_OutputPrinterName.value != "") 
				{
					if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.selectedIndex).innerText != frmOriginalDefinition.txtDefn_OutputPrinterName.value) 
					{
						OpenHR.messageBox("This definition is set to output to printer "+frmOriginalDefinition.txtDefn_OutputPrinterName.value+" which is not set up on your PC.");
						var oOption = document.createElement("OPTION");
						frmDefinition.cboPrinterName.options.add(oOption);
						oOption.innerText = frmOriginalDefinition.txtDefn_OutputPrinterName.value;
						oOption.value = frmDefinition.cboPrinterName.options.length-1;
						frmDefinition.cboPrinterName.selectedIndex = oOption.value;
					}
				}
			}
		}
	}

	function util_def_customreports_addhandlers() {
		
		OpenHR.addActiveXHandler("ssOleDBGridAvailableColumns", "RowColChange", ssOleDBGridAvailableColumns_RowColChange);
		OpenHR.addActiveXHandler("ssOleDBGridAvailableColumns", "DblClick", ssOleDBGridAvailableColumns_DblClick);
		OpenHR.addActiveXHandler("ssOleDBGridAvailableColumns", "KeyPress", ssOleDBGridAvailableColumns_KeyPress);

		OpenHR.addActiveXHandler("ssOleDBGridSelectedColumns", "RowColChange", ssOleDBGridSelectedColumns_RowColChange);
		OpenHR.addActiveXHandler("ssOleDBGridSelectedColumns", "DblClick", ssOleDBGridSelectedColumns_DblClick);
		OpenHR.addActiveXHandler("ssOleDBGridSelectedColumns", "SelChange", ssOleDBGridSelectedColumns_SelChange);

		OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "BeforeUpdate", ssOleDBGridSortOrder_BeforeUpdate);
		OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "AfterInsert", ssOleDBGridSortOrder_AfterInsert);
		OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "RowLoaded", ssOleDBGridSortOrder_RowLoaded);
		OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "RowColChange", ssOleDBGridSortOrder_RowColChange);
		OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "Change", ssOleDBGridSortOrder_Change);
		OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "KeyUp", ssOleDBGridSortOrder_KeyUp);
		OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "Click", ssOleDBGridSortOrder_Click);
		OpenHR.addActiveXHandler("ssOleDBGridSortOrder", "DblClick", ssOleDBGridSortOrder_DblClick);
		
		OpenHR.addActiveXHandler("ssOleDBGridRepetition", "AfterInsert", ssOleDBGridRepetition_AfterInsert);
		OpenHR.addActiveXHandler("ssOleDBGridRepetition", "Change", ssOleDBGridRepetition_Change);
		OpenHR.addActiveXHandler("ssOleDBGridRepetition", "Click", ssOleDBGridRepetition_Click);
		OpenHR.addActiveXHandler("ssOleDBGridRepetition", "DblClick", ssOleDBGridRepetition_DblClick);
		OpenHR.addActiveXHandler("ssOleDBGridRepetition", "KeyUp", ssOleDBGridRepetition_KeyUp);
		OpenHR.addActiveXHandler("ssOleDBGridRepetition", "RowColChange", ssOleDBGridRepetition_RowColChange);
		OpenHR.addActiveXHandler("ssOleDBGridRepetition", "RowLoaded", ssOleDBGridRepetition_RowLoaded);

		OpenHR.addActiveXHandler("ssOleDBGridChildren", "Change", ssOleDBGridChildren_Change);
		OpenHR.addActiveXHandler("ssOleDBGridChildren", "Click", ssOleDBGridChildren_Click);
		OpenHR.addActiveXHandler("ssOleDBGridChildren", "AfterInsert", ssOleDBGridChildren_AfterInsert);
		OpenHR.addActiveXHandler("ssOleDBGridChildren", "DblClick", ssOleDBGridChildren_DblClick);
		OpenHR.addActiveXHandler("ssOleDBGridChildren", "RowColChange", ssOleDBGridChildren_RowColChange);
		
		OpenHR.addActiveXHandler("grdAccess", "ComboCloseUp", grdAccess_ComboCloseUp);
		OpenHR.addActiveXHandler("grdAccess", "GotFocus", grdAccess_GotFocus);
		OpenHR.addActiveXHandler("grdAccess", "RowColChange", grdAccess_RowColChange);
		OpenHR.addActiveXHandler("grdAccess", "RowLoaded", grdAccess_RowLoaded);
	}

	function addGridCol(psKey) {
		with (grdColProps) {
			AddItem(psKey);
			Refresh();
		}
	}

	function removeGridCol(psKey) {
		with (grdColProps) {
			Redraw = false;
			MoveFirst();
			for (var i = 0; i < Rows; i++) {
				if (Columns('ColumnID').Value == psKey) {
					RemoveItem(i);
					Redraw = true;
					return true;
				}
				MoveNext();
			}
			Refresh();
			Redraw = true;
			return false;
		}
	}

	function setGirdCol(psKey) {
		with (grdColProps) {
			if (Columns('ColumnID').Value == psKey) {
				return true;
			}

			MoveFirst();
			for (var i = 0; i < Rows; i++) {
				if (Columns('ColumnID').Value == psKey) {
					return true;
				}
				MoveNext();
			}
		}
		return false;
	}

	function displayPage(piPageNumber) {

		if (piPageNumber == 1) {
			div1.style.visibility = "visible";
			div1.style.display = "block";

			div2.style.visibility = "hidden";
			div2.style.display = "block";
			loadChildTables();
			div2.style.visibility = "hidden";
			div2.style.display = "none";

			div3.style.visibility = "hidden";
			div3.style.display = "block";
			loadSelectedColumnsDefinition();
			div3.style.visibility = "hidden";
			div3.style.display = "none";

			div4.style.visibility = "hidden";
			div4.style.display = "block";
			loadSortDefinition();
			loadRepetitionDefinition();
			div4.style.visibility = "hidden";
			div4.style.display = "none";

			div5.style.visibility = "hidden";
			div5.style.display = "none";

			button_disable(frmDefinition.btnTab1, true);
			button_disable(frmDefinition.btnTab2, false);
			button_disable(frmDefinition.btnTab3, false);
			button_disable(frmDefinition.btnTab4, false);
			button_disable(frmDefinition.btnTab5, false);

			if (frmDefinition.txtName.disabled == false) {
				try {
					frmDefinition.txtName.focus();
				}
				catch (e) { }
			}
			refreshTab1Controls();
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

			frmDefinition.ssOleDBGridChildren.SelBookmarks.RemoveAll();
			if (frmDefinition.ssOleDBGridChildren.Rows > 0) {
				frmDefinition.ssOleDBGridChildren.MoveFirst();
				frmDefinition.ssOleDBGridChildren.SelBookmarks.Add(frmDefinition.ssOleDBGridChildren.Bookmark);
			}

			if (frmDefinition.ssOleDBGridChildren.Enabled) {
				frmDefinition.ssOleDBGridChildren.focus();
			}

			refreshTab2Controls();
		}

		if (piPageNumber == 3) {
			// Get the columns/calcs for the current table selection.
			var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");

			if (frmUseful.txtTablesChanged.value == 1) {

				frmGetDataForm.txtAction.value = "LOADREPORTCOLUMNS";
				frmGetDataForm.txtReportBaseTableID.value = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;
				frmGetDataForm.txtReportParent1TableID.value = frmDefinition.txtParent1ID.value;
				frmGetDataForm.txtReportParent2TableID.value = frmDefinition.txtParent2ID.value;
				// assign the tab delimited string of selected child table ids 
				frmGetDataForm.txtReportChildTableID.value = childTableString();

				data_refreshData();

				frmUseful.txtTablesChanged.value = 0;
			}

			div1.style.visibility = "hidden";
			div1.style.display = "none";
			div2.style.visibility = "hidden";
			div2.style.display = "none";
			div3.style.visibility = "visible";
			div3.style.display = "block";

			//We need in initialise the repetition grid here but to do that the
			//container div must not be blocked. 
			// JPD 20020708 Fault 4040
			/*div4.style.visibility="visible";
			div4.style.display="block";*/
			//loadRepetitionDefinition();
			div4.style.visibility = "hidden";
			div4.style.display = "none";

			div5.style.visibility = "hidden";
			div5.style.display = "none";

			button_disable(frmDefinition.btnTab1, false);
			button_disable(frmDefinition.btnTab2, false);
			button_disable(frmDefinition.btnTab3, true);
			button_disable(frmDefinition.btnTab4, false);
			button_disable(frmDefinition.btnTab5, false);

			//loadSelectedColumnsDefinition();

			//frmDefinition.ssOleDBGridAvailableColumns.focus();
			if (frmDefinition.cboTblAvailable.disabled == false) {
				frmDefinition.cboTblAvailable.focus();
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
			div5.style.visibility = "hidden";
			div5.style.display = "none";

			button_disable(frmDefinition.btnTab1, false);
			button_disable(frmDefinition.btnTab2, false);
			button_disable(frmDefinition.btnTab3, false);
			button_disable(frmDefinition.btnTab4, true);
			button_disable(frmDefinition.btnTab5, false);

			frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();
			if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
				frmDefinition.ssOleDBGridSortOrder.MoveFirst();
				frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
			}

			if (frmDefinition.ssOleDBGridSortOrder.Enabled == true) {
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

			if (frmDefinition.chkSummary.disabled == false) {
				frmDefinition.chkSummary.focus();
			}
		}
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

	function populateTableAvailable() {
		with (frmDefinition) {
			//Clear the existing data in the child table combo
			while (cboTblAvailable.options.length > 0) {
				cboTblAvailable.options.remove(0);
			}

			//add the base table to the available tables list
			sTableID = cboBaseTable.options[cboBaseTable.selectedIndex].value;
			var oOption = document.createElement("OPTION");
			cboTblAvailable.options.add(oOption);
			oOption.innerText = cboBaseTable.options[cboBaseTable.selectedIndex].innerText;
			oOption.value = sTableID;
			oOption.selected = true;

			//add the Parent 1 table to the available tables list (if it exists)
			if (txtParent1ID.value > 0) {
				sTableID = txtParent1ID.value;
				var oOption = document.createElement("OPTION");
				cboTblAvailable.options.add(oOption);
				oOption.innerText = txtParent1.value;
				oOption.value = sTableID;
			}

			//add the Parent 2 table to the available tables list (if it exists)
			if (txtParent2ID.value > 0) {
				sTableID = txtParent2ID.value;
				var oOption = document.createElement("OPTION");
				cboTblAvailable.options.add(oOption);
				oOption.innerText = txtParent2.value;
				oOption.value = sTableID;
			}

			//add the child tables to the available tables list (if applicable)
			if (frmUseful.txtChildsLoaded.value == 1) {
				//get the child tables from the grid.

				var pvarbookmark;

				for (var i = 0; i < ssOleDBGridChildren.Rows; i++) {
					pvarbookmark = ssOleDBGridChildren.AddItemBookmark(i);

					sTableID = ssOleDBGridChildren.Columns("TableID").CellValue(pvarbookmark);
					var oOption = document.createElement("OPTION");
					cboTblAvailable.options.add(oOption);
					oOption.innerText = ssOleDBGridChildren.Columns("Table").CellValue(pvarbookmark);
					oOption.value = sTableID;
				}
			}
			else {
				//get the child tables from the original definition elemenets.
				// eg. "2	Absence	12195	Holiday	0	 	0"

				var iChildTableID;
				var sChildTableName;
				var iIndex;
				var iChildCounter;

				for (var i = 0; i < frmUseful.txtChildCount.value; i++) {
					iChildCounter = i + 1;
					var sTemp = new String(document.getElementById('txtReportDefnChildGridString_' + iChildCounter).value);

					iIndex = sTemp.indexOf("	");
					if (iIndex >= 0) {
						iChildTableID = sTemp.substr(0, iIndex);
						sTemp = sTemp.substr(iIndex + 1);

						iIndex = sTemp.indexOf("	");
						if (iIndex >= 0) {
							sChildTableName = sTemp.substr(0, iIndex);
							sTemp = sTemp.substr(iIndex + 1);

							var oOption = document.createElement("OPTION");
							cboTblAvailable.options.add(oOption);
							oOption.innerText = sChildTableName;
							oOption.value = iChildTableID;

						}
					}
				}
			}
		}
	}

	//Fault 4383 - remove leading and trailing spaces from the col heading.
	function trimColHeading() {
		frmDefinition.txtColHeading.value = trim(frmDefinition.txtColHeading.value);
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

		populateTableAvailable();
	}

	function changeBaseTable() {
		var i;

		frmDefinition.ssOleDBGridAvailableColumns.Enabled = false;

		if (frmUseful.txtLoading.value == 'N') {
			if ((frmDefinition.txtBaseFilterID.value > 0) ||
				(frmDefinition.txtBasePicklistID.value > 0) ||
					(frmDefinition.txtParent1FilterID.value > 0) ||
						(frmDefinition.txtParent1PicklistID.value > 0) ||
							(frmDefinition.txtParent2FilterID.value > 0) ||
								(frmDefinition.txtParent2PicklistID.value > 0) ||
									(frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) ||
										((frmUseful.txtAction.value.toUpperCase() != "NEW") &&
											(frmUseful.txtSelectedColumnsLoaded.value == 0))) {

				iAnswer = OpenHR.messageBox("Warning: Changing the base table will result in all table/column specific aspects of this report definition being cleared. Are you sure you wish to continue?", 36, "Custom Reports");
				if (iAnswer == 7) {
					// cancel and change back ! (txtcurrentbasetable)
					setBaseTable(frmUseful.txtCurrentBaseTableID.value);
					frmDefinition.ssOleDBGridAvailableColumns.Enabled = true;
					return;
				}
				else {
					// clear cols and sort order
					if ((frmUseful.txtSelectedColumnsLoaded.value != 0) && (frmDefinition.ssOleDBGridSelectedColumns.Rows > 0)) {
						frmDefinition.ssOleDBGridSelectedColumns.RemoveAll();
					}
					if ((frmUseful.txtSortLoaded.value != 0) && (frmDefinition.ssOleDBGridSortOrder.Rows > 0)) {
						frmDefinition.ssOleDBGridSortOrder.RemoveAll();
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

		while (frmDefinition.ssOleDBGridChildren.Rows > 0) {
			frmDefinition.ssOleDBGridChildren.RemoveAll();
		}

		while (frmDefinition.ssOleDBGridRepetition.Rows > 0) {
			frmDefinition.ssOleDBGridRepetition.RemoveAll();
		}

		var sChildren = new String("");
		var sChildrenNames = new String("");

		var dataCollection = frmTables.elements;
		if (dataCollection != null) {
			sReqdControlName = new String("txtTableChildren_");
			sReqdControlName = sReqdControlName.concat(frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);

			for (i = 0; i < dataCollection.length; i++) {
				sControlName = dataCollection.item(i).name;
				if (sControlName == sReqdControlName) {
					sChildren = dataCollection.item(i).value;
					frmCustomReportChilds.childrenString.value = sChildren;
					break;
				}
			}

			sReqdControlName = new String("txtTableChildrenNames_");
			sReqdControlName = sReqdControlName.concat(frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);

			for (i = 0; i < dataCollection.length; i++) {
				sControlName = dataCollection.item(i).name;
				if (sControlName == sReqdControlName) {
					sChildrenNames = dataCollection.item(i).value;
					frmCustomReportChilds.childrenNames.value = sChildrenNames;
					break;
				}
			}
		}


		iCurrentChildCount = 0;
		iIndex = sChildren.indexOf("	");
		while (iIndex > 0) {
			iCurrentChildCount++;
			sChildren = sChildren.substr(iIndex + 1);
			iIndex = sChildren.indexOf("	");
		}

		if (iCurrentChildCount < 1) {
			with (frmDefinition) {
				grid_disable(ssOleDBGridChildren, true);
				ssOleDBGridChildren.Enabled = false;
				button_disable(cmdAddChild, true);
				button_disable(cmdEditChild, true);
				button_disable(cmdRemoveChild, true);
				button_disable(cmdRemoveAllChilds, true);
			}
		}
		else {
			with (frmDefinition) {
				grid_disable(ssOleDBGridChildren, false);
				ssOleDBGridChildren.Enabled = true;
				button_disable(cmdAddChild, false);
				button_disable(cmdEditChild, false);
				button_disable(cmdRemoveChild, false);
				button_disable(cmdRemoveAllChilds, false);
			}
		}

		frmDefinition.txtBaseTableChildCount.value = iCurrentChildCount;

		//Empty the parent textboxes
		frmDefinition.txtParent1.value = '';
		frmDefinition.txtParent1ID.value = 0;
		frmDefinition.txtParent2.value = '';
		frmDefinition.txtParent2ID.value = 0;

		var sParents = new String("");
		var dataCollection = frmTables.elements;
		if (dataCollection != null) {
			sReqdControlName = new String("txtTableParents_");
			sReqdControlName = sReqdControlName.concat(frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);

			for (i = 0; i < dataCollection.length; i++) {
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
			frmDefinition.txtParent1.value = getTableName(sParent1ID);
			frmDefinition.txtParent1ID.value = sParent1ID;
			sParents = sParents.substr(iIndex + 1);
		}
		iIndex = sParents.indexOf("	");
		if (iIndex > 0) {
			sParent2ID = sParents.substr(0, iIndex);
			frmDefinition.txtParent2.value = getTableName(sParent2ID);
			frmDefinition.txtParent2ID.value = sParent2ID;
		}

		clearParent1RecordOptions();
		clearParent2RecordOptions();

		recalcHiddenChildFiltersCount();

		refreshTab1Controls();
		frmUseful.txtCurrentBaseTableID.value = frmDefinition.cboBaseTable.options(frmDefinition.cboBaseTable.options.selectedIndex).value;
		frmUseful.txtTablesChanged.value = 1;
		frmDefinition.ssOleDBGridAvailableColumns.Enabled = true;
		populateTableAvailable();
	}

	function removeChildTable(piChildTableID, bAutoYes) {
		var i;
		var iCount;
		var iTableID;
		var fChildColumnsSelected;
		var iIndex;
		var iCharIndex;
		var sControlName;
		var sDefn;
		var sMessage;
		var iRowIndex;

		frmUseful.txtCurrentChildTableID.value = piChildTableID;

		if (frmUseful.txtLoading.value == 'N') {
			if ((frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) ||
				((frmUseful.txtAction.value.toUpperCase() != "NEW") &&
					(frmUseful.txtSelectedColumnsLoaded.value == 0))) {
				if (frmUseful.txtCurrentChildTableID.value != 0) {

					// Check if there are any child columns in the selected columns list.
					fChildColumnsSelected = false;
					if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
						if (frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) {
							//frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
							frmDefinition.ssOleDBGridSelectedColumns.movefirst();

							for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
								iTableID = frmDefinition.ssOleDBGridSelectedColumns.Columns("tableID").Text;

								if (isSelectedChildTable(iTableID)) {
									fChildColumnsSelected = true;
									break;
								}

								if (iTableID == frmUseful.txtCurrentChildTableID.value) {
									fChildColumnsSelected = true;
									break;
								}

								frmDefinition.ssOleDBGridSelectedColumns.movenext();
							}

							//frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
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
									iTableID = selectedColumnParameter(dataCollection.item(iIndex).value, "TABLEID");

									if (isSelectedChildTable(iTableID)) {
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
						if (!bAutoYes) {
							sMessage = "One or more columns from the '" + getTableName(piChildTableID) + "' table have been included in the report definition.\n"
								+ "Changing the child table will remove these columns from the report definition.\n"
									+ "Do you wish to continue ?";

							iAnswer = OpenHR.messageBox(sMessage, 36, "Custom Reports");
						}
						else {
							iAnswer = 6;
						}

						if (iAnswer == 7) {
							// cancel and change back !
							return false;
						}
						else {
							// Remove the child table's columns from the selected columns collection.
							if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
								if (frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) {
									//frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
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

									//frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
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
											iTableID = selectedColumnParameter(dataCollection.item(iIndex).value, "TABLEID");
											if (iTableID == frmUseful.txtCurrentChildTableID.value) {
												dataCollection.item(iIndex).value = "";
											}
										}
									}
								}
							}

							// Remove the child table's columns from the sort order collection.
							removeSortColumn(0, frmUseful.txtCurrentChildTableID.value);
						}
					}

				}
			}
			frmUseful.txtChanged.value = 1;
		}

		refreshTab2Controls();
		frmUseful.txtTablesChanged.value = 1;
		frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();

		populateTableAvailable();

		return true;
	}

	function anyChildColumnsSelected() {
		var vBM;

		for (var i = 0; i < frmDefinition.ssOleDBGridChildren.Rows; i++) {
			vBM = frmDefinition.ssOleDBGridChildren.AddItemBookmark(i);

			if (isChildColumnSelected(frmDefinition.ssOleDBGridChildren.Columns(0).CellValue(vBM))) {
				return true;
			}
		}
		return false;
	}

	function isSelectedChildTable(piChildID) {
		var i;
		var pvarbookmark;

		with (frmDefinition.ssOleDBGridChildren) {
			if (Rows > 0) {
				MoveFirst();
				for (var i = 0; i < Rows; i++) {
					pvarbookmark = GetBookmark(i);
					if (Columns('TableID').CellValue(pvarbookmark) == piChildID) {
						return true;
					}
				}
			}
		}
		return false;
	}

	function childTableString() {

		var i;
		var pvarbookmark;
		var sChildString = "";
		var plngRow;

		if (frmUseful.txtChildsLoaded.value == 1) {
			with (frmDefinition.ssOleDBGridChildren) {
				plngRow = AddItemRowIndex(Bookmark);
				if (Rows > 0) {
					MoveFirst();
					for (var i = 0; i < Rows; i++) {
						pvarbookmark = GetBookmark(i);
						sChildString = sChildString + Columns('TableID').CellValue(pvarbookmark);
						if (i != Rows - 1) {
							sChildString = sChildString + "	";
						}
					}
				}
			}
			if (sChildString.length < 1) {
				sChildString = "0";
			}
		}
		else {
			var dataCollection = frmTables.elements;
			if (dataCollection != null) {
				for (var i = 0; i < frmUseful.txtChildCount.value; i++) {
					iChildCounter = i + 1;
					sChildString = sChildString + document.getElementById('txtReportDefnChildTableID_' + iChildCounter).value;
					if (iChildCounter < frmUseful.txtChildCount.value) {
						sChildString = sChildString + "	";
					}
				}
			}
		}

		return sChildString;
	}

	function childFilterString() {

		var i;
		var pvarbookmark;
		var sChildString = "";

		with (frmDefinition.ssOleDBGridChildren) {
			if (Rows > 0) {
				MoveFirst();
				for (var i = 0; i < Rows; i++) {
					pvarbookmark = GetBookmark(i);
					sChildString = sChildString + Columns('FilterID').CellValue(pvarbookmark);
					if (i != Rows - 1) {
						sChildString = sChildString + "	";
					}
				}
			}
		}
		if (sChildString.length < 1) {
			sChildString = "";
		}
		return sChildString;
	}

	function childOrderString() {

		var i;
		var pvarbookmark;
		var sChildString = "";

		with (frmDefinition.ssOleDBGridChildren) {
			if (Rows > 0) {
				MoveFirst();
				for (var i = 0; i < Rows; i++) {
					pvarbookmark = GetBookmark(i);
					sChildString = sChildString + Columns('OrderID').CellValue(pvarbookmark);
					if (i != Rows - 1) {
						sChildString = sChildString + "	";
					}
				}
			}
		}
		if (sChildString.length < 1) {
			sChildString = "";
		}
		return sChildString;
	}

	function refreshTab1Controls() {
		
		var fIsForcedHidden;
		var fViewing;
		var fIsNotOwner;
		var fAllAlreadyHidden;
		var fSilent;

		fSilent = ((frmUseful.txtAction.value.toUpperCase() == "COPY") && (frmUseful.txtLoading.value == "Y"));

		fIsForcedHidden = ((frmSelectionAccess.baseHidden.value == "Y") ||
			(frmSelectionAccess.p1Hidden.value == "Y") ||
				(frmSelectionAccess.p2Hidden.value == "Y") ||
					(frmSelectionAccess.childHidden.value > 0) ||
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
			else {
				if (frmSelectionAccess.forcedHidden.value == "N") {
					//MH20040816 Fault 9048
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

		text_disable(frmDefinition.txtName, (fViewing == true));
		text_disable(frmDefinition.txtDescription, (fViewing == true));
		combo_disable(frmDefinition.cboBaseTable, (fViewing == true));
		button_disable(frmDefinition.cmdBasePicklist, ((frmDefinition.optRecordSelection2.checked == false)
			|| (fViewing == true)));
		button_disable(frmDefinition.cmdBaseFilter, ((frmDefinition.optRecordSelection3.checked == false)
			|| (fViewing == true)));

		if (frmDefinition.optRecordSelection2.checked || frmDefinition.optRecordSelection3.checked) {
			checkbox_disable(frmDefinition.chkPrintFilter, (fViewing == true));
		}
		else {
			frmDefinition.chkPrintFilter.checked = false;
			checkbox_disable(frmDefinition.chkPrintFilter, true);
		}

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

		var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");

		button_disable(frmDefinition.cmdParent1Picklist, ((frmDefinition.txtParent1.value == '')
			|| (frmDefinition.optParent1RecordSelection2.checked == false)
				|| (fViewing == true)));
		button_disable(frmDefinition.cmdParent1Filter, ((frmDefinition.txtParent1.value == '')
			|| (frmDefinition.optParent1RecordSelection3.checked == false)
				|| (fViewing == true)));
		radio_disable(frmDefinition.optParent1RecordSelection1, ((frmDefinition.txtParent1.value == '')
			|| (fViewing == true)));
		radio_disable(frmDefinition.optParent1RecordSelection2, ((frmDefinition.txtParent1.value == '')
			|| (fViewing == true)));
		radio_disable(frmDefinition.optParent1RecordSelection3, ((frmDefinition.txtParent1.value == '')
			|| (fViewing == true)));

		button_disable(frmDefinition.cmdParent2Picklist, ((frmDefinition.txtParent2.value == '')
			|| (frmDefinition.optParent2RecordSelection2.checked == false)
				|| (fViewing == true)));
		button_disable(frmDefinition.cmdParent2Filter, ((frmDefinition.txtParent2.value == '')
			|| (frmDefinition.optParent2RecordSelection3.checked == false)
				|| (fViewing == true)));
		radio_disable(frmDefinition.optParent2RecordSelection1, ((frmDefinition.txtParent2.value == '')
			|| (fViewing == true)));
		radio_disable(frmDefinition.optParent2RecordSelection2, ((frmDefinition.txtParent2.value == '')
			|| (fViewing == true)));
		radio_disable(frmDefinition.optParent2RecordSelection3, ((frmDefinition.txtParent2.value == '')
			|| (fViewing == true)));

		button_disable(frmDefinition.cmdAddChild, ((iCurrentChildCount < 1) || (fViewing == true)));
		button_disable(frmDefinition.cmdEditChild, ((frmDefinition.ssOleDBGridChildren.Rows < 1) || (frmDefinition.ssOleDBGridChildren.SelBookmarks.Count != 1) || (fViewing == true)));
		button_disable(frmDefinition.cmdRemoveChild, ((frmDefinition.ssOleDBGridChildren.Rows < 1) || (frmDefinition.ssOleDBGridChildren.SelBookmarks.Count != 1) || (fViewing == true)));
		button_disable(frmDefinition.cmdRemoveAllChilds, ((frmDefinition.ssOleDBGridChildren.Rows < 1) || (fViewing == true)));

		frmDefinition.ssOleDBGridChildren.Enabled = true;

		if (fViewing || (frmDefinition.txtBaseTableChildCount.value < 1)) {
			grid_disable(frmDefinition.ssOleDBGridChildren, true);
		}
		else {
			grid_disable(frmDefinition.ssOleDBGridChildren, false);
			frmDefinition.ssOleDBGridChildren.SelBookmarks.RemoveAll();
			frmDefinition.ssOleDBGridChildren.SelBookmarks.Add(frmDefinition.ssOleDBGridChildren.Bookmark);
		}

		recalcHiddenChildFiltersCount();
		refreshTab1Controls();

		button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) || (fViewing == true)));

		// Little dodge to get around a browser bug that
		// does not refresh the display on all controls.
		try {
			window.resizeBy(0, -1);
			window.resizeBy(0, 1);
		}
		catch (e) { }
	}

	function refreshTab3Controls() {
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
		button_disable(frmDefinition.cmdColumnMoveDown, fMoveDownDisabled);
		button_disable(frmDefinition.cmdColumnMoveUp, fMoveUpDisabled);

		var blnHeadingDisabled = true;
		var sHeading = "";
		var fSizeDisabled = true;
		var sSize = "";
		var blnDecPlacesDisabled = true;
		var sDecPlaces = "";
		var blnAverageDisabled = true;
		var blnAverageChecked = false;
		var blnCountDisabled = true;
		var blnCountChecked = false;
		var blnTotalDisabled = true;
		var blnTotalChecked = false;
		var blnHiddenDisabled = true;
		var blnHiddenChecked = false;
		var blnGroupDisabled = true;
		var blnGroupChecked = false;
		var blnIsNumeric = false;
		var blnRep = false;
		var blnStatus = false;
		var blnResetAll = false;
		var blnLastColumn = false;
		var iCount = 0;
		var blnSRV = false;
		var blnVOC = false;
		var iPrevIndex = 0;
		var blnPrevGroupChecked = false;
		var sKey = '';

		with (frmDefinition.ssOleDBGridSelectedColumns) {
			iCount = SelBookmarks.Count;

			if (iCount < 1) {
				blnResetAll = true;
				blnStatus = false;
			}
			else {
				blnResetAll = false;

				if ((iCount > 1) || (iCount < 1)) {
					blnStatus = false;
					blnResetAll = true;
					blnLastColumn = true;
				}
				else {
					// only one row selected;
					blnStatus = true;
					blnLastColumn = (AddItemRowIndex(SelBookmarks(0)) == (Rows() - 1));

					if (Row > 0) {
						iPrevIndex = (Row - 1);
						if (Columns("GroupWithNext").CellValue(AddItemBookmark(iPrevIndex)) == '1') {
							blnPrevGroupChecked = true;
						}
						else {
							blnPrevGroupChecked = false;
						}
					}
					else {
						blnPrevGroupChecked = false;
					}

					sKey = Columns(0).Text + Columns(2).Text;
					if (setGirdCol(sKey)) {
						if (blnPrevGroupChecked) {
							Columns(9).Value = false; 	//Average
							Columns(10).Value = false; //Count
							Columns(11).Value = false; //Total
							Columns(12).Value = false;  //Hidden
							updateCurrentColProp('hidden', false);
						}
						//Average 
						if (Columns(9).Value == '1') {
							blnAverageChecked = true;
						}
						else {
							blnAverageChecked = false;
						}

						//Count
						if (Columns(10).Value == '1') {
							blnCountChecked = true;
						}
						else {
							blnCountChecked = false;
						}

						//Total
						if (Columns(11).Value == '1') {
							blnTotalChecked = true;
						}
						else {
							blnTotalChecked = false;
						}

						blnHiddenChecked = getCurrentColProp('hidden');
						blnGroupChecked = ((Columns(13).text == '1') && (!blnLastColumn));
						frmDefinition.chkColGroup.checked = blnGroupChecked;
						blnIsNumeric = (Columns(7).text == '1');
						blnRep = getCurrentColProp('repetition');
						blnSRV = getCurrentColProp('hide');
						blnVOC = getCurrentColProp('value');
					}

					sHeading = Columns(8).text;
					sSize = Columns(4).text;

					if (blnLastColumn) {
						Columns('GroupWithNext').Value = false; 	//Group with Next
					}

					if (blnIsNumeric) {
						sDecPlaces = Columns(5).text;
					}
					else {
						sDecPlaces = '';
					}

					frmDefinition.txtColHeading.value = sHeading;
					frmDefinition.txtSize.value = sSize;
					frmDefinition.txtDecPlaces.value = sDecPlaces;
					frmDefinition.chkColAverage.checked = blnAverageChecked;
					frmDefinition.chkColCount.checked = blnCountChecked;
					frmDefinition.chkColTotal.checked = blnTotalChecked;
					frmDefinition.chkColGroup.checked = blnGroupChecked;
					frmDefinition.chkColHidden.checked = blnHiddenChecked;
				}
			}
		} //end with

		blnStatus = ((blnStatus) && (!fViewing));

		if ((blnResetAll) && (!fViewing)) {
			frmDefinition.txtColHeading.value = '';
			frmDefinition.txtDecPlaces.value = '';
			frmDefinition.txtSize.value = '';
			frmDefinition.chkColAverage.checked = false;
			frmDefinition.chkColCount.checked = false;
			frmDefinition.chkColTotal.checked = false;
			frmDefinition.chkColGroup.checked = false;
			frmDefinition.chkColHidden.checked = false;
		}
		else if (blnPrevGroupChecked == true) {
			frmDefinition.chkColAverage.checked = false;
			frmDefinition.chkColCount.checked = false;
			frmDefinition.chkColTotal.checked = false;
			frmDefinition.chkColHidden.checked = false;
		}

		text_disable(frmDefinition.txtColHeading, ((!blnStatus) || (blnHiddenChecked)));
		text_disable(frmDefinition.txtSize, ((!blnStatus) || (blnHiddenChecked)));
		if ((blnStatus) && (blnIsNumeric) && (!blnHiddenChecked)) {
			text_disable(frmDefinition.txtDecPlaces, false);
		}
		else {
			text_disable(frmDefinition.txtDecPlaces, true);
		}

		if ((blnStatus) && (blnIsNumeric) && (!blnHiddenChecked) && (!blnGroupChecked) && (!blnPrevGroupChecked)) {
			checkbox_disable(frmDefinition.chkColAverage, false);
		}
		else {
			checkbox_disable(frmDefinition.chkColAverage, true);
		}

		if ((blnStatus) && (!blnHiddenChecked) && (!blnGroupChecked) && (!blnPrevGroupChecked)) {
			checkbox_disable(frmDefinition.chkColCount, false);
		}
		else {
			checkbox_disable(frmDefinition.chkColCount, true);
		}

		if ((blnStatus) && (blnIsNumeric) && (!blnHiddenChecked) && (!blnGroupChecked) && (!blnPrevGroupChecked)) {
			checkbox_disable(frmDefinition.chkColTotal, false);
		}
		else {
			checkbox_disable(frmDefinition.chkColTotal, true);
		}

		if ((blnStatus) && (!blnAverageChecked) && (!blnCountChecked) && (!blnTotalChecked) && (!blnGroupChecked) && (!blnRep) && (!blnSRV) && (!blnVOC) && (!blnPrevGroupChecked)) {
			checkbox_disable(frmDefinition.chkColHidden, false);
		}
		else {
			checkbox_disable(frmDefinition.chkColHidden, true);
		}

		if ((blnStatus) && (!blnAverageChecked) && (!blnCountChecked) && (!blnTotalChecked) && (!blnHiddenChecked) && (!blnLastColumn)) {
			checkbox_disable(frmDefinition.chkColGroup, false);
		}
		else {
			checkbox_disable(frmDefinition.chkColGroup, true);
		}

		frmDefinition.ssOleDBGridAvailableColumns.RowHeight = 19;
		frmDefinition.ssOleDBGridSelectedColumns.RowHeight = 19;

		button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
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

		with (frmDefinition.ssOleDBGridSortOrder) {
			Enabled = true;
			AllowUpdate = (!fViewing);
			Columns(1).Locked = true;
			CheckBox3D = false;

			if (fViewing) {
				grid_disable(frmDefinition.ssOleDBGridSortOrder, true);
				SelectTypeCol = 0;
				RowNavigation = 3;
			}
			else {
				grid_disable(frmDefinition.ssOleDBGridSortOrder, false);
				SelectByCell = false;
				SelectTypeCol = 0;

				SelBookmarks.RemoveAll();
				SelBookmarks.Add(Bookmark);
			}
		}

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
					sControlName = dataCollection.item(i).name;
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

		fSortRemoveDisabled = ((frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count != 1)
			|| (fViewing == true) || (frmDefinition.ssOleDBGridSortOrder.Rows < 1));
		fSortRemoveAllDisabled = ((fViewing == true) || (frmDefinition.ssOleDBGridSortOrder.Rows < 1));
		fSortEditDisabled = ((frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count != 1)
			|| (fViewing == true) || (frmDefinition.ssOleDBGridSortOrder.Rows < 1));

		button_disable(frmDefinition.cmdSortAdd, fSortAddDisabled);
		button_disable(frmDefinition.cmdSortEdit, fSortEditDisabled);
		button_disable(frmDefinition.cmdSortRemove, fSortRemoveDisabled);
		button_disable(frmDefinition.cmdSortRemoveAll, fSortRemoveAllDisabled);
		button_disable(frmDefinition.cmdSortMoveUp, fSortMoveUpDisabled);
		button_disable(frmDefinition.cmdSortMoveDown, fSortMoveDownDisabled);

		frmDefinition.ssOleDBGridSortOrder.AllowUpdate = (fViewing == false);

		//********************************************************************

		if (anyChildColumnsSelected()) {
			frmUseful.txtChildColumnSelected.value = 1;
		}
		else {
			frmUseful.txtChildColumnSelected.value = 0;
		}

		with (frmDefinition.ssOleDBGridRepetition) {
			Enabled = true;
			AllowUpdate = ((fViewing == false) && (frmUseful.txtChildColumnSelected.value == 1));
			Columns(1).Locked = true;
			CheckBox3D = false;

			if ((fViewing) || (frmUseful.txtChildColumnSelected.value == 0)) {
				if (frmUseful.txtChildColumnSelected.value == 0) {
					clearRepetition();
				}

				grid_disable(frmDefinition.ssOleDBGridRepetition, true);
				RowNavigation = 3;
				SelectByCell = false;
				SelectTypeCol = 0;
			}
			else {
				grid_disable(frmDefinition.ssOleDBGridRepetition, false);
				SelectTypeRow = 0;
				SelectByCell = false;
				SelectTypeCol = 0;

				SelBookmarks.RemoveAll();
				SelBookmarks.Add(Bookmark);
			}
		}

		//********************************************************************

		frmDefinition.ssOleDBGridSortOrder.RowHeight = 19;
		frmDefinition.ssOleDBGridRepetition.RowHeight = 19;

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
				txtFilename.value = '';
				//text_disable(txtFilename, true);
				button_disable(cmdFilename, true);

				//disable email options
				chkDestination3.checked = false;
				checkbox_disable(chkDestination3, true);
				//text_disable(txtEmailGroup, true);
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
				combo_disable(cboPrinterName, true);
				cboPrinterName.length = 0;

				//enable-disable save options
				checkbox_disable(chkDestination2, (fViewing == true));
				if (chkDestination2.checked == true) {
					populateSaveExisting();
					combo_disable(cboSaveExisting, (fViewing == true));
					//text_disable(txtFilename, (fViewing == true));
					button_disable(cmdFilename, (fViewing == true));
				}
				else {
					cboSaveExisting.length = 0;
					combo_disable(cboSaveExisting, true);
					txtFilename.value = '';
					//text_disable(txtFilename, true);
					button_disable(cmdFilename, true);
				}

				//enable-disable email options
				checkbox_disable(chkDestination3, (fViewing == true));
				if (chkDestination3.checked == true) {
					//text_disable(txtEmailGroup, (fViewing == true));
					text_disable(txtEmailSubject, (fViewing == true));
					button_disable(cmdEmailGroup, (fViewing == true));
					text_disable(txtEmailAttachAs, (fViewing == true));
				}
				else {
					//text_disable(txtEmailGroup, true);
					txtEmailGroup.value = '';
					txtEmailGroupID.value = 0;
					text_disable(txtEmailSubject, true);
					button_disable(cmdEmailGroup, true);
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
				combo_disable(cboPrinterName, true);
				cboPrinterName.length = 0;

				//enable-disable save options
				checkbox_disable(chkDestination2, (fViewing == true));
				if (chkDestination2.checked == true) {
					populateSaveExisting();
					combo_disable(cboSaveExisting, (fViewing == true));
					//text_disable(txtFilename, (fViewing == true));
					button_disable(cmdFilename, (fViewing == true));
				}
				else {
					cboSaveExisting.length = 0;
					combo_disable(cboSaveExisting, true);
					//text_disable(txtFilename, true);
					txtFilename.value = '';
					button_disable(cmdFilename, true);
				}

				//enable-disable email options
				checkbox_disable(chkDestination3, (fViewing == true));
				if (chkDestination3.checked == true) {
					//text_disable(txtEmailGroup, (fViewing == true));
					text_disable(txtEmailSubject, (fViewing == true));
					button_disable(cmdEmailGroup, (fViewing == true));
					text_disable(txtEmailAttachAs, (fViewing == true));
				}
				else {
					//text_disable(txtEmailGroup, true);
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
					//text_disable(txtFilename, (fViewing == true));
					button_disable(cmdFilename, (fViewing == true));
				}
				else {
					cboSaveExisting.length = 0;
					combo_disable(cboSaveExisting, true);
					//text_disable(txtFilename, true);
					txtFilename.value = '';
					button_disable(cmdFilename, true);
				}

				//enable-disable email options
				checkbox_disable(chkDestination3, (fViewing == true));
				if (chkDestination3.checked == true) {
					//text_disable(txtEmailGroup, (fViewing == true));
					text_disable(txtEmailSubject, (fViewing == true));
					button_disable(cmdEmailGroup, (fViewing == true));
					text_disable(txtEmailAttachAs, (fViewing == true));
				}
				else {
					//text_disable(txtEmailGroup, true);
					txtEmailGroup.value = '';
					txtEmailGroupID.value = 0;
					button_disable(cmdEmailGroup, true);
					text_disable(txtEmailSubject, true);
					text_disable(txtEmailAttachAs, true);
				}
			}
			else if ((optOutputFormat4.checked == true) || 	//Excel Worksheet
				(optOutputFormat5.checked == true) ||
					(optOutputFormat6.checked == true)) {
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
					//text_disable(txtFilename, (fViewing == true));
					button_disable(cmdFilename, (fViewing == true));
				}
				else {
					cboSaveExisting.length = 0;
					combo_disable(cboSaveExisting, true);
					//text_disable(txtFilename, true);
					txtFilename.value = '';
					button_disable(cmdFilename, true);
				}

				//enable-disable email options
				checkbox_disable(chkDestination3, (fViewing == true));
				if (chkDestination3.checked == true) {
					//text_disable(txtEmailGroup, (fViewing == true));
					text_disable(txtEmailSubject, (fViewing == true));
					button_disable(cmdEmailGroup, (fViewing == true));
					text_disable(txtEmailAttachAs, (fViewing == true));
				}
				else {
					//text_disable(txtEmailGroup, true);
					txtEmailGroup.value = '';
					txtEmailGroupID.value = 0;
					button_disable(cmdEmailGroup, true);
					text_disable(txtEmailSubject, true);
					text_disable(txtEmailAttachAs, true);
				}
			}
			else {
				optOutputFormat0.checked = true;
				chkDestination0.checked = true;
				refreshTab5Controls();
			}

			if (txtEmailSubject.disabled) {
				txtEmailSubject.value = '';
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

		// Little dodge to get around a browser bug that
		// does not refresh the display on all controls.
		try {
			window.resizeBy(0, -1);
			window.resizeBy(0, 1);
		}
		catch (e) { }
	}

	function clearRepetition() {
		with (grdColProps) {
			Redraw = false;
			if (Rows > 0) {
				MoveFirst();
			}

			for (var i = 0; i < Rows; i++) {
				updateCurrentColProp('repetition', false);
				MoveNext();
			}

			if (Rows > 0) {
				MoveFirst();
			}
			Redraw = true;
		}

		with (frmDefinition.ssOleDBGridRepetition) {
			Redraw = false;
			if (Rows > 0) {
				MoveFirst();
			}

			for (var i2 = 0; i2 < Rows; i2++) {
				Columns("repetition").Value = false;
				MoveNext();
			}

			if (Rows > 0) {
				MoveFirst();
			}
			Redraw = true;
		}
	}

	function clearSortColumnProps() {
		with (grdColProps) {
			Redraw = false;
			if (Rows > 0) {
				MoveFirst();
			}

			for (var i = 0; i < Rows; i++) {
				updateCurrentColProp('break', false);
				updateCurrentColProp('page', false);
				updateCurrentColProp('value', false);
				updateCurrentColProp('hide', false);
				MoveNext();
			}

			if (Rows > 0) {
				MoveFirst();
			}
			Redraw = true;
		}
	}

	function breakingType(psKey) {
		with (frmDefinition.ssOleDBGridSortOrder) {
			var bm;

			Redraw = false;

			for (var iLoop = 0; iLoop < Rows; iLoop++) {
				bm = AddItemBookmark(iLoop);
				if ((Columns("ColumnID").CellValue(bm) == psKey) &&
					(Columns(3).CellValue(bm) == "-1")) {
					Redraw = true;
					return 1;
				}
				else if ((Columns("ColumnID").CellValue(bm) == psKey) &&
					(Columns(4).CellValue(bm) == "-1")) {
					Redraw = true;
					return 2;
				}
			}
			Redraw = true;
		}
		return 0;
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


	function changeBaseTableRecordOptions() {

		frmDefinition.txtBasePicklist.value = '';
		frmDefinition.txtBasePicklistID.value = 0;

		frmDefinition.txtBaseFilter.value = '';
		frmDefinition.txtBaseFilterID.value = 0;

		frmSelectionAccess.baseHidden.value = "N";

		//frmDefinition.chkPrintFilter.checked = false;

		frmUseful.txtChanged.value = 1;
		refreshTab1Controls();
	}

	function changeParent1TableRecordOptions() {
		
		frmDefinition.txtParent1Picklist.value = '';
		frmDefinition.txtParent1PicklistID.value = 0;

		frmDefinition.txtParent1Filter.value = '';
		frmDefinition.txtParent1FilterID.value = 0;

		frmSelectionAccess.p1Hidden.value = "N";

		frmUseful.txtChanged.value = 1;
		refreshTab2Controls();
	}

	function changeParent2TableRecordOptions() {
		
		frmDefinition.txtParent2Picklist.value = '';
		frmDefinition.txtParent2PicklistID.value = 0;

		frmDefinition.txtParent2Filter.value = '';
		frmDefinition.txtParent2FilterID.value = 0;

		frmSelectionAccess.p2Hidden.value = "N";

		frmUseful.txtChanged.value = 1;
		refreshTab2Controls();
	}

	function clearBaseTableRecordOptions() {
		
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

	function clearParent1RecordOptions() {

		frmDefinition.optParent1RecordSelection1.checked = true;

		button_disable(frmDefinition.cmdParent1Picklist, true);
		frmDefinition.txtParent1Picklist.value = '';
		frmDefinition.txtParent1PicklistID.value = 0;

		button_disable(frmDefinition.cmdParent1Filter, true);
		frmDefinition.txtParent1Filter.value = '';
		frmDefinition.txtParent1FilterID.value = 0;

		frmSelectionAccess.p1Hidden.value = 'N';
	}

	function clearParent2RecordOptions() {
				
		frmDefinition.optParent2RecordSelection1.checked = true;

		button_disable(frmDefinition.cmdParent2Picklist, true);
		frmDefinition.txtParent2Picklist.value = '';
		frmDefinition.txtParent2PicklistID.value = 0;

		button_disable(frmDefinition.cmdParent2Filter, true);
		frmDefinition.txtParent2Filter.value = '';
		frmDefinition.txtParent2FilterID.value = 0;

		frmSelectionAccess.p2Hidden.value = 'N';
	}

	function checkHiddenOptions() {
				
		var sKey = frmDefinition.ssOleDBGridSortOrder.Columns("ColumnID").Value;
		var colHidden;
		colHidden = isColumnHidden('C' + sKey);
		if ((colHidden == true) && ((frmDefinition.ssOleDBGridSortOrder.Columns("Value").Value == -1)
			|| (frmDefinition.ssOleDBGridSortOrder.Columns("hide").Value == -1))) {
			frmDefinition.ssOleDBGridSortOrder.Columns("Value").Value = 0;
			frmDefinition.ssOleDBGridSortOrder.Columns("hide").Value = 0;
			var sMessage = "You cannot select 'Value on Change' or 'Suppress Repeated Values' for a hidden column.";
			OpenHR.messageBox(sMessage, 64, "Custom Reports");
		}
	}

	function selectEmailGroup() {
		var sUrl;
		
		frmEmailSelection.EmailSelCurrentID.value = frmDefinition.txtEmailGroupID.value;

		sUrl = "util_emailSelection" +
			"?EmailSelCurrentID=" + frmEmailSelection.EmailSelCurrentID.value;
		openDialog(sUrl, (screen.width) / 3, (screen.height) / 2, "yes", "yes");
	}

	function selectRecordOption(psTable, psType) {
		var sURL;
		var frmDefinition = document.getElementById("frmDefinition");
		var frmUseful = document.getElementById("frmUseful");
		var frmRecordSelection = document.getElementById("frmRecordSelection");

	    debugger;

		if (psTable == 'base') {
			iTableID = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;

			if (psType == 'picklist') {
				iCurrentID = frmDefinition.txtBasePicklistID.value;
			}
			else {
				iCurrentID = frmDefinition.txtBaseFilterID.value;
			}
		}
		if (psTable == 'p1') {
			iTableID = frmDefinition.txtParent1ID.value;

			if (psType == 'picklist') {
				iCurrentID = frmDefinition.txtParent1PicklistID.value;
			}
			else {
				iCurrentID = frmDefinition.txtParent1FilterID.value;
			}
		}
		if (psTable == 'p2') {
			iTableID = frmDefinition.txtParent2ID.value;

			if (psType == 'picklist') {
				iCurrentID = frmDefinition.txtParent2PicklistID.value;
			}
			else {
				iCurrentID = frmDefinition.txtParent2FilterID.value;
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

		if (strDefOwner == strCurrentUser) {
			frmRecordSelection.recSelDefOwner.value = '1';
		}
		else {
			frmRecordSelection.recSelDefOwner.value = '0';
		}
		frmRecordSelection.recSelDefType.value = "Custom Reports";

		sURL = "util_recordSelection" +
			"?recSelType=" + escape(frmRecordSelection.recSelType.value) +
				"&recSelTableID=" + escape(frmRecordSelection.recSelTableID.value) +
					"&recSelCurrentID=" + escape(frmRecordSelection.recSelCurrentID.value) +
						"&recSelTable=" + escape(frmRecordSelection.recSelTable.value) +
							"&recSelDefOwner=" + escape(frmRecordSelection.recSelDefOwner.value) +
								"&recSelDefType=" + escape(frmRecordSelection.recSelDefType.value);
		openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");

		frmUseful.txtChanged.value = 1;
		if (psTable == 'base') {
			refreshTab1Controls();
		}
		else {
			refreshTab2Controls();
		}
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
		if (validateTab5() == false) { OpenHR.refreshMenu(); return; }
		if (populateSendForm() == false) { OpenHR.refreshMenu(); return; }

		// Now create the validate popup to check that any filters/calcs
		// etc havent been deleted, or made hidden etc.		

		// first populate the validate fields
		frmValidate.validateBaseFilter.value = frmDefinition.txtBaseFilterID.value;
		frmValidate.validateBasePicklist.value = frmDefinition.txtBasePicklistID.value;
		frmValidate.validateEmailGroup.value = frmDefinition.txtEmailGroupID.value;
		frmValidate.validateP1Filter.value = frmDefinition.txtParent1FilterID.value;
		frmValidate.validateP1Picklist.value = frmDefinition.txtParent1PicklistID.value;
		frmValidate.validateP2Filter.value = frmDefinition.txtParent2FilterID.value;
		frmValidate.validateP2Picklist.value = frmDefinition.txtParent2PicklistID.value;
		frmValidate.validateChildFilter.value = childFilterString();
		frmValidate.validateChildOrders.value = childOrderString();
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

		sURL = "util_validate_customreports" +
			"?validateBaseFilter=" + escape(frmValidate.validateBaseFilter.value) +
				"&validateBasePicklist=" + escape(frmValidate.validateBasePicklist.value) +
					"&validateEmailGroup=" + escape(frmValidate.validateEmailGroup.value) +
						"&validateP1Filter=" + escape(frmValidate.validateP1Filter.value) +
							"&validateP1Picklist=" + escape(frmValidate.validateP1Picklist.value) +
								"&validateP2Filter=" + escape(frmValidate.validateP2Filter.value) +
									"&validateP2Picklist=" + escape(frmValidate.validateP2Picklist.value) +
										"&validateChildFilter=" + escape(frmValidate.validateChildFilter.value) +
											"&validateChildOrders=" + escape(frmValidate.validateChildOrders.value) +
												"&validateCalcs=" + escape(frmValidate.validateCalcs.value) +
													"&validateHiddenGroups=" + escape(frmValidate.validateHiddenGroups.value) +
														"&validateName=" + escape(frmValidate.validateName.value) +
															"&validateTimestamp=" + escape(frmValidate.validateTimestamp.value) +
																"&validateUtilID=" + frmValidate.validateUtilID.value +
																	"&destination=util_validate_customreports";
		openDialog(sURL, (screen.width) / 2, (screen.height) / 3, "no", "no");
	}

	function cancelClick() {
		if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
			(definitionChanged() == false)) {
			//TODO
			//window.location.href = "defsel";
			return;
		}

		answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3, "Custom Reports");
		if (answer == 7) {
			// No
			//TODO
			//window.location.href = "defsel";
			return (false);
		}
		if (answer == 6) {
			// Yes
			okClick();
		}
	}

	function okClick() {
		OpenHR.disableMenu();

		sAttachmentName = new String(frmDefinition.txtEmailAttachAs.value);
		if ((sAttachmentName.indexOf("/") != -1) ||
			(sAttachmentName.indexOf(":") != -1) ||
				(sAttachmentName.indexOf("?") != -1) ||
					(sAttachmentName.indexOf(String.fromCharCode(34)) != -1) ||
						(sAttachmentName.indexOf("<") != -1) ||
							(sAttachmentName.indexOf(">") != -1) ||
								(sAttachmentName.indexOf("|") != -1) ||
									(sAttachmentName.indexOf("\\") != -1) ||
										(sAttachmentName.indexOf("*") != -1)) {
			OpenHR.messageBox("The attachment file name can not contain any of the following characters:\n/ : ? " + String.fromCharCode(34) + " < > | \\ *", 48, "Custom Reports");
			return;
		}

		frmSend.txtSend_reaction.value = "CUSTOMREPORTS";
		submitDefinition();
	}

	function saveChanges(psAction, pfPrompt, pfTBOverride) {
		if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
			(definitionChanged() == false)) {
			return 7; //No to saving the changes, as none have been made.
		}

		answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3, "Custom Reports");
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
			// Compare the tab 1 controls with the original values.
			if (frmUseful.txtAction.value.toUpperCase() != "NEW") {
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

				if ((frmOriginalDefinition.txtDefn_PrintFilterHeader.value != "False") != frmDefinition.chkPrintFilter.checked) {
					return true;
				}

				// Ignore Zeros
				if ((frmOriginalDefinition.txtDefn_IgnoreZeros.value != "False") != frmDefinition.chkIgnoreZeros.checked) {
					return true;
				}

				// Compare the tab 2 controls with the original values.
				if (frmOriginalDefinition.txtDefn_Parent1TableID.value != frmDefinition.txtParent1ID.value) {
					return true;
				}

				if (frmOriginalDefinition.txtDefn_Parent1PicklistID.value > 0) {
					if (frmDefinition.optParent1RecordSelection2.checked == false) {
						return true;
					}
					else {
						if (frmDefinition.txtParent1PicklistID.value != frmOriginalDefinition.txtDefn_Parent1PicklistID.value) {
							return true;
						}
					}
				}
				else {
					if (frmOriginalDefinition.txtDefn_Parent1FilterID.value > 0) {
						if (frmDefinition.optParent1RecordSelection3.checked == false) {
							return true;
						}
						else {
							if (frmDefinition.txtParent1FilterID.value != frmOriginalDefinition.txtDefn_Parent1FilterID.value) {
								return true;
							}
						}
					}
					else {
						if (frmDefinition.optParent1RecordSelection1.checked == false) {
							return true;
						}
					}
				}

				if (frmOriginalDefinition.txtDefn_Parent2TableID.value != frmDefinition.txtParent2ID.value) {
					return true;
				}

				if (frmOriginalDefinition.txtDefn_Parent2PicklistID.value > 0) {
					if (frmDefinition.optParent2RecordSelection2.checked == false) {
						return true;
					}
					else {
						if (frmDefinition.txtParent2PicklistID.value != frmOriginalDefinition.txtDefn_Parent2PicklistID.value) {
							return true;
						}
					}
				}
				else {
					if (frmOriginalDefinition.txtDefn_Parent2FilterID.value > 0) {
						if (frmDefinition.optParent2RecordSelection3.checked == false) {
							return true;
						}
						else {
							if (frmDefinition.txtParent2FilterID.value != frmOriginalDefinition.txtDefn_Parent2FilterID.value) {
								return true;
							}
						}
					}
					else {
						if (frmDefinition.optParent2RecordSelection1.checked == false) {
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

		var sReqdControlName = new String("txtTableName_");
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

	function replace(sExpression, sFind, sReplace) {
		//gi (global search, ignore case)
		var re = new RegExp(sFind, "gi");
		sExpression = sExpression.replace(re, sReplace);
		return (sExpression);
	}

	function columnSwap(pfSelect) {
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
			for (i = iCount - 1; i >= 0; i--) {
				grdFrom.bookmark = grdFrom.selbookmarks(i);
				iRowIndex = grdFrom.AddItemRowIndex(grdFrom.Bookmark);

				// Remove the column from the sort columns collection.
				if (grdFrom.columns(0).text == "C") {
					if (frmUseful.txtSortLoaded.value == 1) {
						if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
							frmDefinition.ssOleDBGridSortOrder.Redraw = false;
							frmDefinition.ssOleDBGridSortOrder.MoveFirst();

							iCount2 = frmDefinition.ssOleDBGridSortOrder.rows;
							for (i2 = 0; i2 < iCount2; i2++) {
								if (grdFrom.columns(2).text == frmDefinition.ssOleDBGridSortOrder.Columns("id").Text) {
									// The selected column is a sort column. Prompt the user to confirm the deselection.

									sColumnName = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
									if (iCount > 1) {
										iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the report sort order.\n\nDo you still want to remove this column ?", 3, "Custom Reports");
									}
									else {
										iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the report sort order.\n\nDo you still want to remove this column ?", 4, "Custom Reports");
									}

									if (iResponse == 2) {
										// Cancel.
										frmDefinition.ssOleDBGridSortOrder.Redraw = true;
										grdFrom.Redraw = true;
										grdTo.Redraw = true;
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
							for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
								sControlName = dataCollection.item(iIndex).name;
								sControlName = sControlName.substr(0, 19);
								if (sControlName == "txtReportDefnOrder_") {
									if (grdFrom.columns(2).text == sortColumnParameter(dataCollection.item(iIndex).value, "COLUMNID")) {
										// The selected column is a sort column. Prompt the user to confirm the deselection.
										sColumnName = grdFrom.columns(3).text;
										if (iCount > 1) {
											iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the report sort order.\n\nDo you still want to remove this column ?", 3, "Custom Reports");
										}
										else {
											iResponse = OpenHR.messageBox("Removing the '" + sColumnName + "' column will also remove it from the report sort order.\n\nDo you still want to remove this column ?", 4, "Custom Reports");
										}

										if (iResponse == 2) {
											// Cancel.
											frmDefinition.ssOleDBGridSortOrder.Redraw = true;
											grdFrom.Redraw = true;
											grdTo.Redraw = true;
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
			iHiddenCalcCount = 0;

			for (i = 0; i < grdFrom.selbookmarks.Count(); i++) {
				grdFrom.bookmark = grdFrom.selbookmarks(i);

				// Check if the user is selecting a hidden calc, but is not the report owner.
				if ((pfSelect == true) &&
					(frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) &&
						(grdFrom.columns(6).text == "Y")) {

					sCalcName = new String(grdFrom.columns(3).text);
					OpenHR.messageBox("Cannot include the '" + sCalcName + "' calculation.\nIts hidden and you are not the creator of this definition.", 64, "Custom Reports");
				}
				else {
					iColumnsSwapped = iColumnsSwapped + 1;

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
							sAddline = sAddline + '	<' + getTableName(grdFrom.columns(1).text) + ' Calc> ' + grdFrom.columns(3).text;
						}
						else {
							sTemp = grdFrom.columns(3).text;
							iTemp = sTemp.indexOf(" Calc> ");
							if (iTemp >= 0) {
								sTemp = sTemp.substring(iTemp + 7);
							}
							sAddline = sAddline + '	' + sTemp;
						}

						sAddline = sAddline + '	' + grdFrom.columns(4).text +
							'	' + grdFrom.columns(5).text +
								'	' + grdFrom.columns(6).text +
									'	' + grdFrom.columns(7).text;

					}

					if (pfSelect == true) {
						sAddline = sAddline +
							'	';

						sAddline = sAddline +
							grdFrom.columns(3).text;

						sAddline = sAddline +
							'	' + '0' + '	' + '0' + '	' + '0' + '	' + '0' + '	' + '0';

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

						//Add the column/calc to the repetition grid if it is a base or parent column/calc.
						if (grdFrom.columns(1).text == frmUseful.txtCurrentBaseTableID.value
							|| grdFrom.columns(1).text == frmDefinition.txtParent1ID.value
								|| grdFrom.columns(1).text == frmDefinition.txtParent2ID.value) {
							var sRepeatAddStr;

							if (frmDefinition.optCalc.checked) {
								sRepeatAddStr = grdFrom.columns(0).text + grdFrom.columns(2).text +
									'	' + '<' + frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].innerText + ' Calc> ' + grdFrom.columns(3).text +
										'	' + 0 +
											'	' + grdFrom.columns(1).text;
							}
							else {
								sRepeatAddStr = grdFrom.columns(0).text + grdFrom.columns(2).text +
									'	' + frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].innerText + '.' + grdFrom.columns(3).text +
										'	' + 0 +
											'	' + grdFrom.columns(1).text;
							}

							frmDefinition.ssOleDBGridRepetition.AddItem(sRepeatAddStr);
						}
						/* TM 19/06/02 - Fault 3991 */
						grdTo.MoveLast();
						grdTo.AddItem(sAddline);
						grdTo.MoveLast();
						addGridCol(grdFrom.columns(0).text + grdFrom.columns(2).text);

					}
					else {
						/* Find the right spot to add the row. */

						sFromType = grdFrom.columns(0).text;
						sFromTableID = grdFrom.columns(1).text;

						sTemp = grdFrom.columns(3).text;
						iTemp = sTemp.indexOf(" Calc> ");
						if (iTemp >= 0) {
							sTemp = sTemp.substring(iTemp + 7);
						}
						sFromDisplay = replace(sTemp, "_", " ");
						sFromDisplay = sFromDisplay.substring(sFromDisplay.indexOf(".") + 1);
						sFromDisplay = sFromDisplay.toUpperCase();

						fIsFromTblAvailable = (sFromTableID == frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].value);

						fIsFromTypeAvailable = (((sFromType == "C") && (frmDefinition.optColumns.checked)) ||
							((sFromType == "E") && (frmDefinition.optCalc.checked)));

						fFound = true;

						if (fIsFromTblAvailable && fIsFromTypeAvailable) {

							fFound = false;
							grdTo.movefirst();
							grdTo.Redraw = true;
							for (i2 = 0; i2 < grdTo.rows(); i2++) {
								grdTo.Redraw = false;

								sToType = grdTo.columns(0).text;
								sToTableID = grdTo.columns(1).text;
								sToDisplay = replace(grdTo.columns(3).text.toUpperCase(), "_", " ");

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
							grdColProps.RemoveAll();
							// Clear the sort columns collection.
							removeSortColumn(0, 0);
							removeRepetitionColumn(0, 0);
						}
					}
					else {
						if (pfSelect == false) {
							// Remove the column from the sort columns collection.
							if (grdFrom.columns(0).text == "C") {
								removeSortColumn(grdFrom.columns(2).text, 0);
							}

							// Remove the column from the repetition columns collection.
							removeRepetitionColumn(grdFrom.columns(2).text, 0);

							grdFrom.RemoveItem(iRowIndex);
							removeGridCol(grdFrom.columns(0).text + grdFrom.columns(2).text);
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

			if (sAddedCalcIDs.length > 0) {
				// Get the return types of the added calcs.

				var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
				frmGetDataForm.txtAction.value = "GETEXPRESSIONRETURNTYPES";
				frmGetDataForm.txtParam1.value = sAddedCalcIDs;
				data_refreshData();
			}
		}

		grdFrom.Redraw = true;
		grdTo.Redraw = true;
		refreshTab3Controls();
		refreshTab4Controls();
	}

	function columnSwapAll(pfSelect) {
		var iColumnsSwapped;
		var sAddedCalcIDs;

		sAddedCalcIDs = "";
		iColumnsSwapped = 0;

		if (pfSelect == true) {
			var grdFrom = frmDefinition.ssOleDBGridAvailableColumns;
			var grdTo = frmDefinition.ssOleDBGridSelectedColumns;
		}
		else {
			if (frmUseful.txtSortLoaded.value == 1) {
				iSortedColumnCount = frmDefinition.ssOleDBGridSortOrder.Rows;
			}
			else {
				iSortedColumnCount = 0;
				var dataCollection = frmOriginalDefinition.elements;
				if (dataCollection != null) {
					for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
						sControlName = dataCollection.item(iIndex).name;
						sControlName = sControlName.substr(0, 19);
						if (sControlName == "txtReportDefnOrder_") {
							iSortedColumnCount = 1;
							break;
						}
					}
				}
			}

			if (iSortedColumnCount > 0) {
				var iAnswer = OpenHR.messageBox("Removing all columns will remove all sort order columns. \n Are you sure ?", 36, "Custom Reports");
			}
			else {
				iAnswer = 6;
			}

			if (iAnswer == 7) {
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
		for (var i = 0; i < grdFrom.Rows(); i++) {
			if ((pfSelect == true) &&
				(frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) &&
					(grdFrom.columns(6).text == "Y")) {

				sCalcName = new String(grdFrom.columns(3).text);
				iStringIndex = sCalcName.indexOf("> ");
				if (iStringIndex >= 0) {
					sCalcName = sCalcName.substring(iStringIndex, sCalcName.length);
				}
				OpenHR.messageBox("Cannot include the '" + sCalcName + "' calculation.\nIts hidden and you are not the creator of this definition.", 64, "Custom Reports");
			}
			else {
				iColumnsSwapped = iColumnsSwapped + 1;

				if (grdFrom.columns(0).text == 'C') {
					sAddline = grdFrom.columns(0).text +
						'	' + grdFrom.columns(1).text +
							'	' + grdFrom.columns(2).text

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
						sAddline = sAddline + '	<' + getTableName(grdFrom.columns(1).text) + ' Calc> ' + grdFrom.columns(3).text;
					}
					else {
						sTemp = grdFrom.columns(3).text;
						iTemp = sTemp.indexOf(" Calc> ");
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
						'	';

					sAddline = sAddline +
						grdFrom.columns(3).text;

					sAddline = sAddline +
						'	' + '0' + '	' + '0' + '	' + '0' + '	' + '0' + '	' + '0';

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
					//NPG20110929 - Fault HRPRO-1790
					//				//Add the column/calc to the repetition grid if it is a base or parent column/calc.
					//				if (grdFrom.columns(1).text == frmUseful.txtCurrentBaseTableID.value  
					//					|| grdFrom.columns(1).text == frmDefinition.txtParent1ID.value
					//					|| grdFrom.columns(1).text == frmDefinition.txtParent2ID.value) 
					//					{
					//					var sRepeatAddStr;
					//					
					//					if (frmDefinition.optCalc.checked)
					//						{
					//						sRepeatAddStr = grdFrom.columns(0).text + grdFrom.columns(2).text +
					//									'	' + '<' + frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].innerText + ' Calc> ' + grdFrom.columns(3).text +
					//									'	' + 0 +
					//									'	' + grdFrom.columns(1).text ;
					//						}
					//					else
					//						{
					//						sRepeatAddStr = grdFrom.columns(0).text + grdFrom.columns(2).text +
					//									'	' + frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].innerText + '.' + grdFrom.columns(3).text +
					//									'	' + 0 +
					//									'	' + grdFrom.columns(1).text ;
					//						}									
					//						
					//					frmDefinition.ssOleDBGridRepetition.AddItem(sRepeatAddStr);
					//					}
					//						
					grdTo.additem(sAddline);
					//				addGridCol(grdFrom.columns(0).text + grdFrom.columns(2).text);
				}
				else {
					/* Find the right spot to add the row. */
					sFromType = grdFrom.columns(0).text;
					sFromTableID = grdFrom.columns(1).text;

					sTemp = grdFrom.columns(3).text;
					iTemp = sTemp.indexOf(" Calc> ");
					if (iTemp >= 0) {
						sTemp = sTemp.substring(iTemp + 7);
					}
					sFromDisplay = replace(sTemp, "_", " ");
					sFromDisplay = sFromDisplay.substring(sFromDisplay.indexOf(".") + 1);
					sFromDisplay = sFromDisplay.toUpperCase();

					fIsFromTblAvailable = (sFromTableID == frmDefinition.cboTblAvailable.options[frmDefinition.cboTblAvailable.selectedIndex].value);

					fIsFromTypeAvailable = (((sFromType == "C") && (frmDefinition.optColumns.checked)) ||
						((sFromType == "E") && (frmDefinition.optCalc.checked)));

					fFound = true;

					if (fIsFromTblAvailable && fIsFromTypeAvailable) {
						fFound = false;
						grdTo.movefirst();
						grdTo.Redraw = true;
						/* TM 19/06/02 - Fault 4000 */
						for (i2 = 0; i2 < grdTo.rows(); i2++) {
							grdTo.Redraw = false;

							sToType = grdTo.columns(0).text;
							sToTableID = grdTo.columns(1).text;
							sToDisplay = replace(grdTo.columns(3).text.toUpperCase(), "_", " ");

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
				grdColProps.RemoveAll();
				// Clear the sort columns collection.
				removeSortColumn(0, 0);
				removeRepetitionColumn(0, 0);
				frmUseful.txtSortLoaded.value = 1;
			}

			if (iHiddenCalcCount > 0) {
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

			if (sAddedCalcIDs.length > 0) {
				// Get the return types of the added calcs.

				var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
				frmGetDataForm.txtAction.value = "GETEXPRESSIONRETURNTYPES";
				frmGetDataForm.txtParam1.value = sAddedCalcIDs;
				data_refreshData();
			}
		}

		refreshTab3Controls();
		refreshTab4Controls();
	}

	function columnMove(pfUp) {
		
		var iNewIndex, iOldIndex, iSelectIndex;
		
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
		refreshTab3Controls();
	}

	function setAggregate(piKey) {
		// piKey 0 = Average, 1 = Count, 2 = Total, 3 = Hidden, 4 = Group with Next
		
		if (piKey == 0) {
			if (frmDefinition.chkColAverage.checked == true) {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(9).text = '1';
			}
			else {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(9).text = '0';
			}
		}

		if (piKey == 1) {
			if (frmDefinition.chkColCount.checked == true) {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(10).text = '1';
			}
			else {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(10).text = '0';
			}
		}

		if (piKey == 2) {
			if (frmDefinition.chkColTotal.checked == true) {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(11).text = '1';
			}
			else {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(11).text = '0';
			}
		}

		if (piKey == 3) {
			var sKey = frmDefinition.ssOleDBGridSelectedColumns.Columns(0).text +
				frmDefinition.ssOleDBGridSelectedColumns.Columns(2).text;

			if (setGirdCol(sKey)) {
				updateCurrentColProp('hidden', frmDefinition.chkColHidden.checked);
			}

			if (frmDefinition.chkColHidden.checked == true) {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(12).text = '1';
			}
			else {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(12).text = '0';
			}
		}

		if (piKey == 4) {
			if (frmDefinition.chkColGroup.checked == true) {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(13).text = '1';
			}
			else {
				frmDefinition.ssOleDBGridSelectedColumns.Columns(13).text = '0';
			}
		}

		frmUseful.txtChanged.value = 1;

		refreshTab3Controls();
		frmDefinition.ssOleDBGridSortOrder.Refresh();
		frmDefinition.ssOleDBGridRepetition.Refresh();
	}

	function validateColHeading() {
		
		frmDefinition.ssOleDBGridSelectedColumns.columns(8).text = frmDefinition.txtColHeading.value;
		frmUseful.txtChanged.value = 1;
		refreshTab3Controls();
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

		// Convert any decimal separators to '.'.
		if (OpenHR.LocaleDecimalSeparator != ".") {
			// Remove decimal points.
			sConvertedValue = sConvertedValue.replace(rePoint, "A");
			// replace the locale decimal marker with the decimal point.
			sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
		}

		if (isNaN(sConvertedValue) == true) {
			OpenHR.messageBox("Invalid numeric value.", 48, "Custom Reports");
			frmDefinition.txtSize.value = frmDefinition.ssOleDBGridSelectedColumns.columns(4).text;
			frmDefinition.txtSize.focus();
			return false;
		}
		else {
			if (sConvertedValue.indexOf(".") >= 0) {
				OpenHR.messageBox("Invalid integer value.", 48, "Custom Reports");
				frmDefinition.txtSize.value = frmDefinition.ssOleDBGridSelectedColumns.columns(4).text;
				frmDefinition.txtSize.focus();
				return false;
			}
			else {
				if (frmDefinition.txtSize.value < 0) {
					OpenHR.messageBox("The value must be greater than or equal to 0.", 48, "Custom Reports");
					frmDefinition.txtSize.value = frmDefinition.ssOleDBGridSelectedColumns.columns(4).text;
					frmDefinition.txtSize.focus();
					return false;
				}

				if (frmDefinition.txtSize.value > 2147483646) {
					OpenHR.messageBox("The value must be less than or equal to 2147483646.", 48, "Custom Reports");
					frmDefinition.txtSize.value = frmDefinition.ssOleDBGridSelectedColumns.columns(4).text;
					frmDefinition.txtSize.focus();
					return false;
				}

			}
		}

		frmDefinition.ssOleDBGridSelectedColumns.columns(4).text = frmDefinition.txtSize.value;
		frmUseful.txtChanged.value = 1;
		refreshTab3Controls();
		return true;
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
		sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
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
		//frmDefinition.txtDecPlaces.value = sConvertedValue;

		// Convert any decimal separators to '.'.
		if (OpenHR.LocaleDecimalSeparator != ".") {
			// Remove decimal points.
			sConvertedValue = sConvertedValue.replace(rePoint, "A");
			// replace the locale decimal marker with the decimal point.
			sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
		}

		if (isNaN(sConvertedValue) == true) {
			OpenHR.messageBox("Decimal places must be numeric.", 48, "Custom Reports");
			frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
			frmDefinition.txtDecPlaces.focus();
			return false;
		}
		else {
			if (sConvertedValue.indexOf(".") >= 0) {
				OpenHR.messageBox("Invalid integer value.", 48, "Custom Reports");
				frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
				frmDefinition.txtDecPlaces.focus();
				return false;
			}
			else {
				if (frmDefinition.txtDecPlaces.value < 0) {
					OpenHR.messageBox("The value cannot be less than 0.", 48, "Custom Reports");
					frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
					frmDefinition.txtDecPlaces.focus();
					return false;
				}

				if (frmDefinition.txtDecPlaces.value > 999) {
					OpenHR.messageBox("The value must be less than or equal to 999.", 48, "Custom Reports");
					frmDefinition.txtDecPlaces.value = frmDefinition.ssOleDBGridSelectedColumns.columns(5).text;
					frmDefinition.txtDecPlaces.focus();
					return false;
				}
			}
		}
		frmDefinition.ssOleDBGridSelectedColumns.columns(5).text = frmDefinition.txtDecPlaces.value;
		frmUseful.txtChanged.value = 1;
		refreshTab3Controls();
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
		window.showModalDialog(pDestination, self, dlgwinprops);
	}

	function childAdd() {
		
		frmCustomReportChilds.childTableID.value = 0;
		frmCustomReportChilds.childTable.value = '';
		frmCustomReportChilds.childFilterID.value = 0;
		frmCustomReportChilds.childFilter.value = '';
		frmCustomReportChilds.childOrderID.value = 0;
		frmCustomReportChilds.childOrder.value = '';
		frmCustomReportChilds.childRecords.value = 0;

		frmCustomReportChilds.childAction.value = "NEW";

		if (frmDefinition.ssOleDBGridChildren.Rows < frmCustomReportChilds.childMax.value) {
			if (frmDefinition.ssOleDBGridChildren.Rows == frmDefinition.txtBaseTableChildCount.value) {
				OpenHR.messageBox("All child tables for the current base table have been added to the report definition.", 64, "Custom Reports");
			}
			else {
				var sURL = "util_customreportchilds" +
					"?childTableID=" + escape(frmCustomReportChilds.childTableID.value) +
						"&childTable=" + escape(frmCustomReportChilds.childTable.value) +
							"&childFilterID=" + escape(frmCustomReportChilds.childFilterID.value) +
								"&childFilter=" + escape(frmCustomReportChilds.childFilter.value) +
									"&childOrderID=" + escape(frmCustomReportChilds.childOrderID.value) +
										"&childOrder=" + escape(frmCustomReportChilds.childOrder.value) +
											"&childRecords=" + escape(frmCustomReportChilds.childRecords.value) +
												"&childrenString=" + escape(frmCustomReportChilds.childrenString.value) +
													"&childrenNames=" + escape(frmCustomReportChilds.childrenNames.value) +
														"&selectedChildString=" + escape(frmCustomReportChilds.selectedChildString.value) +
															"&childAction=" + escape(frmCustomReportChilds.childAction.value) +
																"&childMax=" + escape(frmCustomReportChilds.childMax.value);
				openDialog(sURL, 365, 275, "no", "no");

				frmUseful.txtChanged.value = 1;
				frmUseful.txtTablesChanged.value = 1;
			}
		}
		else {
			OpenHR.messageBox("The maximum of five child tables has been selected.", 64, "Custom Reports");
		}
		refreshTab2Controls();

	}

	function childEdit() {
		var sALL_RECORDS = "All Records";
				
		frmCustomReportChilds.childTableID.value = frmDefinition.ssOleDBGridChildren.Columns('TableID').text;
		frmCustomReportChilds.childTable.value = frmDefinition.ssOleDBGridChildren.Columns('Table').text;
		frmCustomReportChilds.childFilterID.value = frmDefinition.ssOleDBGridChildren.Columns('FilterID').text;
		frmCustomReportChilds.childFilter.value = frmDefinition.ssOleDBGridChildren.Columns('Filter').text;
		frmCustomReportChilds.childOrderID.value = frmDefinition.ssOleDBGridChildren.Columns('OrderID').text;
		frmCustomReportChilds.childOrder.value = frmDefinition.ssOleDBGridChildren.Columns('Order').text;

		if (frmDefinition.ssOleDBGridChildren.Columns("Records").text == sALL_RECORDS) {
			frmCustomReportChilds.childRecords.value = 0;
		}
		else {
			frmCustomReportChilds.childRecords.value = frmDefinition.ssOleDBGridChildren.Columns("Records").text;
		}

		frmCustomReportChilds.childAction.value = "EDIT";

		var sURL = "util_customreportchilds" +
			"?childTableID=" + escape(frmCustomReportChilds.childTableID.value) +
				"&childTable=" + escape(frmCustomReportChilds.childTable.value) +
					"&childFilterID=" + escape(frmCustomReportChilds.childFilterID.value) +
						"&childFilter=" + escape(frmCustomReportChilds.childFilter.value) +
							"&childOrderID=" + escape(frmCustomReportChilds.childOrderID.value) +
								"&childOrder=" + escape(frmCustomReportChilds.childOrder.value) +
									"&childRecords=" + escape(frmCustomReportChilds.childRecords.value) +
										"&childrenString=" + escape(frmCustomReportChilds.childrenString.value) +
											"&childrenNames=" + escape(frmCustomReportChilds.childrenNames.value) +
												"&selectedChildString=" + escape(frmCustomReportChilds.selectedChildString.value) +
													"&childAction=" + escape(frmCustomReportChilds.childAction.value) +
														"&childMax=" + escape(frmCustomReportChilds.childMax.value);
		openDialog(sURL, 365, 275, "no", "no");

		frmUseful.txtTablesChanged.value = 1;
		frmUseful.txtChanged.value = 1;

		refreshTab2Controls();

		populateTableAvailable();
	}

	function childRemove() {
		var lRow;
		var lngSelectedChild;
		var varBookmark;

		with (frmDefinition.ssOleDBGridChildren) {
			if (Rows < 1) return;

			lRow = AddItemRowIndex(Bookmark);
			//lngSelectedChild = Columns('TableID').CellValue(lRow);
			varBookmark = Bookmark();
			lngSelectedChild = Columns('TableID').CellValue(varBookmark);

			//' Check if any columns in the report definition are from the table that was
			//' previously selected in the child combo box. If so, prompt user for action.
			var bContinueRemoval;

			frmDefinition.ssOleDBGridChildren.Redraw = false;
			bContinueRemoval = removeChildTable(lngSelectedChild, false);
			frmDefinition.ssOleDBGridChildren.Redraw = true;

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
		frmUseful.txtTablesChanged.value = 1;
		frmUseful.txtChanged.value = 1;

		refreshTab2Controls();

		populateTableAvailable();
	}

	function isChildColumnSelected(plngChildTableID) {
		var lngTempTableID = 0;

		if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
			with (frmDefinition.ssOleDBGridSelectedColumns) {
				for (var i = 0; i < Rows(); i++) {

					lngTempTableID = Columns("tableid").CellText(AddItemBookmark(i));

					if (lngTempTableID == plngChildTableID) {
						return true;
					}
				}
			}
		}
		else {
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				var iNum = 0;

				for (var iIndex = 0; iIndex < dataCollection.length; iIndex++) {

					var sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 20);
					if (sControlName == "txtReportDefnColumn_") {
						iNum = iNum + 1;

						lngTempTableID = selectedColumnParameter(dataCollection.item(iIndex).value, "TABLEID");

						if (lngTempTableID == plngChildTableID) {
							return true;
						}

					}
				}
			}
		}

		return false;
	}

	function childRemoveAll() {
		var i;
		var pvarbookmark;
		var bContinueRemoval;
		var lngSelectedChild;
		var lngRowCount;
		var blnChildColumnSelected = false;

		with (frmDefinition.ssOleDBGridChildren) {

			var sMessage = "Removing all the child tables will remove all child table columns "
				+ "included in the report definition.\n"
					+ "Do you wish to continue ?";


			if (Rows < 1) return;

			Redraw = false;
			lngRowCount = Rows;
			for (i = 0; i < lngRowCount; i++) {
				MoveFirst();
				pvarbookmark = AddItemBookmark(i);
				lngSelectedChild = Columns('TableID').CellValue(pvarbookmark);

				blnChildColumnSelected = isChildColumnSelected(lngSelectedChild);
			}

			if (blnChildColumnSelected) {
				bContinueRemoval = (OpenHR.messageBox(sMessage, 4 + 48, "Custom Reports") == 6);
			}
			else {
				bContinueRemoval = true;
			}

			if (!bContinueRemoval) {
				Redraw = true;
				return;
			}

			Redraw = false;
			lngRowCount = Rows;
			for (i = 0; i < lngRowCount; i++) {
				MoveFirst();
				pvarbookmark = AddItemBookmark(i);
				lngSelectedChild = Columns('TableID').CellValue(pvarbookmark);

				removeChildTable(lngSelectedChild, true);
			}
			Redraw = true;
			RemoveAll();
			SelBookmarks.RemoveAll();
		}

		frmUseful.txtTablesChanged.value = 1;
		frmUseful.txtChanged.value = 1;

		refreshTab2Controls();

		populateTableAvailable();
	}

	function refreshAvailableColumns() {
		if (frmUseful.txtLoading.value == 'N') {
			loadAvailableColumns();
		}
	}

	function sortAdd() {
		var i;
		var iCalcsCount = 0;
		var iColumnsCount = 0;
		var sURL;
		
		// Loop through the columns added and populate the 
		// sort order text boxes to pass to util_sortorderselection.asp
		frmSortOrder.txtSortInclude.value = '';
		frmSortOrder.txtSortExclude.value = '';
		frmSortOrder.txtSortEditing.value = 'false';
		frmSortOrder.txtSortColumnID.value = frmDefinition.ssOleDBGridSortOrder.Columns(0).text;
		frmSortOrder.txtSortColumnName.value = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
		frmSortOrder.txtSortOrder.value = frmDefinition.ssOleDBGridSortOrder.Columns(2).text;
		frmSortOrder.txtSortBOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(3).text;
		frmSortOrder.txtSortPOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(4).text;
		frmSortOrder.txtSortVOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(5).text;
		frmSortOrder.txtSortSRV.value = frmDefinition.ssOleDBGridSortOrder.Columns(6).text;

		if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
			frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
			frmDefinition.ssOleDBGridSelectedColumns.movefirst();

			for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
				if (frmDefinition.ssOleDBGridSelectedColumns.Columns(0).Text == 'C') {
					iColumnsCount++;
					if (frmSortOrder.txtSortInclude.value != '') {
						frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
					}
					frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.ssOleDBGridSelectedColumns.columns(2).text;
				}
				else {
					iCalcsCount++;
				}
				frmDefinition.ssOleDBGridSelectedColumns.movenext();
			}

			frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
		}
		else {
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
								iColumnsCount++;

								if (frmSortOrder.txtSortInclude.value != '') {
									frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
								}
								frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + sColumnID;
							}
							else {
								iCalcsCount++;
							}
						}
					}
				}
			}
		}

		if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
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
			OpenHR.messageBox("You must add more columns to the report before you can add to the sort order.", 48, "Custom Reports");
		}
		else if ((frmDefinition.ssOleDBGridSortOrder.Rows - iColumnsCount) == 0) {
			OpenHR.messageBox("You must add more columns to the report before you can add to the sort order.", 48, "Custom Reports");
		}
		else {
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
				openDialog(sURL, 600, 275, "yes", "yes");

				frmUseful.txtChanged.value = 1;
			}
		}

		frmDefinition.ssOleDBGridSortOrder.Refresh();
		frmDefinition.ssOleDBGridRepetition.Refresh();

		refreshTab4Controls();
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
		frmSortOrder.txtSortBOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(3).text;
		frmSortOrder.txtSortPOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(4).text;
		frmSortOrder.txtSortVOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(5).text;
		frmSortOrder.txtSortSRV.value = frmDefinition.ssOleDBGridSortOrder.Columns(6).text;

		if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
			frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
			frmDefinition.ssOleDBGridSelectedColumns.MoveFirst();

			for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.Rows(); i++) {
				if (frmDefinition.ssOleDBGridSelectedColumns.Columns(0).Text == "C") {
					if (frmSortOrder.txtSortInclude.value != '') {
						frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
					}
					frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.ssOleDBGridSelectedColumns.Columns(2).Text;
				}

				frmDefinition.ssOleDBGridSelectedColumns.MoveNext();
			}

			frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
		}
		else {
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {
					var sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 20);
					if (sControlName == "txtReportDefnColumn_") {
						sColumnID = "";
						sDefn = new String(dataCollection.item(i).value);
						if (sDefn.substr(0, 1) == "C") {
							if (frmSortOrder.txtSortInclude.value != '') {
								frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
							}
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

		for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.Rows(); i++) {
			if (frmDefinition.ssOleDBGridSortOrder.columns(0).text != frmSortOrder.txtSortColumnID.value) {
				if (frmSortOrder.txtSortExclude.value != '') {
					frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + ',';
				}
				frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + frmDefinition.ssOleDBGridSortOrder.Columns(0).Text;
			}

			frmDefinition.ssOleDBGridSortOrder.MoveNext();
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

		frmDefinition.ssOleDBGridSortOrder.Refresh();
		frmDefinition.ssOleDBGridRepetition.Refresh();

		refreshTab4Controls();
	}

	function sortRemove() {

		if ((frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count() == 0) || (frmDefinition.ssOleDBGridSortOrder.Rows() < 1)) {
			OpenHR.messageBox("You must select a column to remove.", 48, "Custom Reports");
			return;
		}

		var sKey = new String('C' + frmDefinition.ssOleDBGridSortOrder.Columns(0).value);

		if (setGirdCol(sKey)) {
			updateCurrentColProp('break', false);
			updateCurrentColProp('page', false);
			updateCurrentColProp('value', false);
			updateCurrentColProp('hide', false);
		}

		frmDefinition.ssOleDBGridSortOrder.RemoveItem(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.bookmark));

		if (frmDefinition.ssOleDBGridSortOrder.Rows != 0) {
			frmDefinition.ssOleDBGridSortOrder.MoveLast();
			frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
		}
		frmUseful.txtChanged.value = 1;

		frmDefinition.ssOleDBGridSortOrder.Refresh();
		frmDefinition.ssOleDBGridRepetition.Refresh();

		refreshTab4Controls();
	}

	function sortRemoveAll() {

		frmDefinition.ssOleDBGridSortOrder.RemoveAll();
		clearSortColumnProps();
		frmUseful.txtChanged.value = 1;

		frmDefinition.ssOleDBGridSortOrder.Refresh();
		frmDefinition.ssOleDBGridRepetition.Refresh();

		refreshTab4Controls();
	}

	function sortMove(pfUp) {

		var frmDefinition = document.getElementById("frmDefinition");
		var frmUseful = document.getElementById("frmUseful");
		var iNewIndex, iOldIndex, iSelectIndex;
		
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

		var sAddline = frmDefinition.ssOleDBGridSortOrder.columns(0).text +
			'	' + frmDefinition.ssOleDBGridSortOrder.columns(1).text +
				'	' + frmDefinition.ssOleDBGridSortOrder.columns(2).text +
					'	' + frmDefinition.ssOleDBGridSortOrder.columns(3).text +
						'	' + frmDefinition.ssOleDBGridSortOrder.columns(4).text +
							'	' + frmDefinition.ssOleDBGridSortOrder.columns(5).text +
								'	' + frmDefinition.ssOleDBGridSortOrder.columns(6).text +
									'	' + frmDefinition.ssOleDBGridSortOrder.columns(7).text;

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
			OpenHR.messageBox("You must enter a name for this definition.", 48, "Custom Reports");
			displayPage(1);
			return (false);
		}

		// check base picklist
		if ((frmDefinition.optRecordSelection2.checked == true) && (frmDefinition.txtBasePicklistID.value == 0)) {
			OpenHR.messageBox("You must select a picklist for the base table.", 48, "Custom Reports");
			displayPage(1);
			return (false);
		}

		// check base filter
		if ((frmDefinition.optRecordSelection3.checked == true) && (frmDefinition.txtBaseFilterID.value == 0)) {
			OpenHR.messageBox("You must select a filter for the base table.", 48, "Custom Reports");
			displayPage(1);
			return (false);
		}

		return (true);
	}

	function validateTab2() {

		// check parent 1 picklist
		if ((frmDefinition.optParent1RecordSelection2.checked == true) &&
			(frmDefinition.txtParent1PicklistID.value == 0)) {
			OpenHR.messageBox("You must select a picklist for the first parent table.", 48, "Custom Reports");
			displayPage(2);
			return (false);
		}

		// check Parent 1 filter
		if ((frmDefinition.optParent1RecordSelection3.checked == true) &&
			(frmDefinition.txtParent1FilterID.value == 0)) {
			OpenHR.messageBox("You must select a filter for the first parent table.", 48, "Custom Reports");
			displayPage(2);
			return (false);
		}

		// check parent 2 picklist
		if ((frmDefinition.optParent2RecordSelection2.checked == true) &&
			(frmDefinition.txtParent2PicklistID.value == 0)) {
			OpenHR.messageBox("You must select a picklist for the second parent table.", 48, "Custom Reports");
			displayPage(2);
			return (false);
		}

		// check Parent 2 filter
		if ((frmDefinition.optParent2RecordSelection3.checked == true) &&
			(frmDefinition.txtParent2FilterID.value == 0)) {
			OpenHR.messageBox("You must select a filter for the second parent table.", 48, "Custom Reports");
			displayPage(2);
			return (false);
		}

		return (true);
	}

	function validateTab3() {
		var i;
		var iCount;
		var sAllHeadings;
		var sType;
		var sHidden;
		var sErrMsg;
		var sCurrentHeading;
		var sDefn;
		var sControlName;
		var sColName;
		var sLowerCaseHeading;
		var blnHasAggregate;
		var blnHasNumericAggregate;
		
		blnHasAggregate = false;
		blnHasNumericAggregate = false;

		sColName = "";
		sLowerCaseHeading = "";

		sErrMsg = "";
		sAllHeadings = "";

		// Check report columns have been selected
		// Check all cols have a (unique) heading
		// Any hidden calcs included?
		if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
			if (frmDefinition.ssOleDBGridSelectedColumns.Rows == 0) {
				sErrMsg = "You must select columns for the report.";
			}
			else {
				frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
				frmDefinition.ssOleDBGridSelectedColumns.movefirst();

				for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {

					//Disable Group With Next for the last column in the report.
					if (i == (frmDefinition.ssOleDBGridSelectedColumns.rows - 1)) {
						frmDefinition.ssOleDBGridSelectedColumns.columns("GroupWithNext").value = false;
					}

					frmDefinition.ssOleDBGridSelectedColumns.columns("heading").text = trim(frmDefinition.ssOleDBGridSelectedColumns.Columns("heading").Text);
					sCurrentHeading = frmDefinition.ssOleDBGridSelectedColumns.Columns("heading").Text.toLowerCase();
					sColName = frmDefinition.ssOleDBGridSelectedColumns.Columns("display").Text.toLowerCase();
					sHiddenColumn = frmDefinition.ssOleDBGridSelectedColumns.Columns("hiddencolumn").Text.toLowerCase();

					if ((sCurrentHeading == '') && (sHiddenColumn == "0")) {
						sErrMsg = "The '" + frmDefinition.ssOleDBGridSelectedColumns.Columns("display").Text + "' column has a blank column heading.";
						break;
					}
					if (sAllHeadings.indexOf('	' + sCurrentHeading + '	') != -1) {
						if (sHiddenColumn == "0") {
							sErrMsg = "One or more columns in the report have a heading of '" + frmDefinition.ssOleDBGridSelectedColumns.columns("heading").text + "'. All column headings must be unique.";
							break;
						}
					}

					if (sCurrentHeading.substr(0, 3) == "?id") {
						sErrMsg = "The '" + sColName + "' column has a heading beginning '?ID'. '?ID' is a reserved word and cannot be used at the beginning of a column heading.";
						break;
					}

					else {
						if (sHiddenColumn == "0") {
							sAllHeadings = sAllHeadings + '	' + sCurrentHeading.toLowerCase() + '	';
						}

						if ((frmDefinition.ssOleDBGridSelectedColumns.columns("type").text == 'E') &&
							(frmDefinition.ssOleDBGridSelectedColumns.columns("hidden").text == 'Y')) {
							if (frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) {
								sErrMsg = "You have selected a hidden calculation but you are not the owner of the definition.";
								break;
							}
						}
					}

					if ((frmDefinition.ssOleDBGridSelectedColumns.columns("average").text == '1')
						|| (frmDefinition.ssOleDBGridSelectedColumns.columns("count").text == '1')
							|| (frmDefinition.ssOleDBGridSelectedColumns.columns("total").text == '1')) {
						blnHasAggregate = true;
					}

					if ((frmDefinition.ssOleDBGridSelectedColumns.columns("numeric").text == '1')
						&& ((frmDefinition.ssOleDBGridSelectedColumns.columns("average").text == '1')
							|| (frmDefinition.ssOleDBGridSelectedColumns.columns("count").text == '1')
								|| (frmDefinition.ssOleDBGridSelectedColumns.columns("total").text == '1'))) {
						blnHasNumericAggregate = true;
					}

					frmDefinition.ssOleDBGridSelectedColumns.movenext();
				}

				frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
				frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();
				frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.add(frmDefinition.ssOleDBGridSelectedColumns.bookmark);
			}
		}
		else {
			iCount = 0;
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {

					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 20);
					if (sControlName == "txtReportDefnColumn_") {

						sDefn = new String(dataCollection.item(i).value);

						if (trim(sDefn) != '') {
							sCurrentHeading = trim(selectedColumnParameter(sDefn, "HEADING"));
							sType = selectedColumnParameter(sDefn, "TYPE");
							sHidden = selectedColumnParameter(sDefn, "HIDDEN");
							sColName = selectedColumnParameter(sDefn, "DISPLAY");
							sLowerCaseHeading = sCurrentHeading.toLowerCase();
							sHiddenColumn = selectedColumnParameter(sDefn, "HIDDENCOLUMN");

							iCount = iCount + 1;

							if ((sLowerCaseHeading == '') && (sHiddenColumn == "0")) {
								sErrMsg = "All columns must have a heading.";
								break;
							}

							if (sAllHeadings.indexOf('	' + sLowerCaseHeading + '	') != -1) {
								if (sHiddenColumn == "0") {
									sErrMsg = "One or more columns in the report have a heading of '" + sCurrentHeading + "'. All column headings must be unique.";
									break;
								}
							}

							if (sLowerCaseHeading.substr(0, 3) == '?id') {
								sErrMsg = "The '" + sColName + "' column has a heading beginning '?ID'. '?ID' is a reserved word and cannot be used at the beginning of a column heading.";
								break;
							}

							sAllHeadings = sAllHeadings + '	' + sLowerCaseHeading + '	';

							if ((sType == 'E') && (sHidden == 'Y')) {
								if (frmDefinition.txtOwner.value.toUpperCase() != frmUseful.txtUserName.value.toUpperCase()) {
									sErrMsg = "You have selected a hidden calculation but you are not the owner of the definition.";
									break;
								}
							}

							if ((selectedColumnParameter(sDefn, "AVERAGE") == '1')
								|| (selectedColumnParameter(sDefn, "COUNT") == '1')
									|| (selectedColumnParameter(sDefn, "TOTAL") == '1')) {
								blnHasAggregate = true;
							}

							if ((selectedColumnParameter(sDefn, "NUMERIC") == '1')
								&& ((selectedColumnParameter(sDefn, "AVERAGE") == '1')
									|| (selectedColumnParameter(sDefn, "COUNT") == '1')
										|| (selectedColumnParameter(sDefn, "TOTAL") == '1'))) {
								blnHasNumericAggregate = true;
							}

						}
					}
				}
			}

			if (iCount == 0) {
				sErrMsg = "You must select columns for the report.";
			}
		}

		if (frmDefinition.chkSummary.checked) {
			if (!blnHasAggregate) {
				sErrMsg = "You have defined this report as a summary report but have not selected to show aggregates for any of the columns.";
			}
		}

		if (frmDefinition.chkIgnoreZeros.checked) {
			if (!blnHasNumericAggregate) {
				sErrMsg = "You have chosen to ignore zeros when calculating aggregates, but have not selected to show aggregates for any numeric columns.";
			}
		}

		if (sErrMsg.length > 0) {
			OpenHR.messageBox(sErrMsg, 48, "Custom Reports");
			displayPage(3);
			return (false);
		}

		return (true);
	}


	function validateTab4() {
		var i;
		var sErrMsg;
		var iIndex;
		var iCount;
		var sPage;
		var sBreak;
		var sDefn;
		var sControlName;
		var blnHasVOC;
		
		blnHasVOC = false;

		sErrMsg = "";

		//check at least one column defined as sort order
		if (frmUseful.txtSortLoaded.value == 1) {
			if (frmDefinition.ssOleDBGridSortOrder.Rows == 0) {
				sErrMsg = "You must select a column to order the report by.";
			}
			else {
				frmDefinition.ssOleDBGridSortOrder.redraw = false;
				frmDefinition.ssOleDBGridSortOrder.movefirst();

				// check boc and poc not both selected
				for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) {

					if ((frmDefinition.ssOleDBGridSortOrder.Columns("break").Text == -1) &&
						(frmDefinition.ssOleDBGridSortOrder.Columns("page").Text == -1)) {
						sErrMsg = "You cannot select break on change and page on change for the same column.";
						break;
					}

					if ((frmDefinition.ssOleDBGridSortOrder.Columns("value").Text == -1)) {
						blnHasVOC = true;
					}

					frmDefinition.ssOleDBGridSortOrder.movenext();
				}
				frmDefinition.ssOleDBGridSortOrder.redraw = true;
			}
		}
		else {
			iCount = 0;
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {

					if (dataCollection.item(i).value != "") {
						sControlName = dataCollection.item(i).name;
						sControlName = sControlName.substr(0, 19);
						if (sControlName == "txtReportDefnOrder_") {
							sDefn = new String(dataCollection.item(i).value);
							sPage = sortColumnParameter(sDefn, "POC");
							sBreak = sortColumnParameter(sDefn, "BOC");

							if ((sBreak == "-1") && (sPage == "-1")) {
								sErrMsg = "You cannot select break on change and page on change for the same column.";
								break;
							}

							if ((sortColumnParameter(sDefn, "VOC") == "1")) {
								blnHasVOC = true;
							}

							iCount = iCount + 1;
						}
					}
				}
			}

			if (iCount == 0) {
				sErrMsg = "You must select a column to order the report by.";
			}
		}


		if (frmDefinition.chkSummary.checked) {
			if (!blnHasVOC) {
				sErrMsg = "You have defined this report as a summary report but have not set a column as 'Value on Change'.\n\nDo you wish to continue?";
				if ((OpenHR.messageBox(sErrMsg, 36, "Custom Reports") == 7)) {
					displayPage(4);
					return (false);
				}
			}
		}
		else {
			if (sErrMsg.length > 0) {
				OpenHR.messageBox(sErrMsg, 48, "Custom Reports");
				displayPage(4);
				return (false);
			}
		}

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

		if (sErrMsg.length > 0) {
			OpenHR.messageBox(sErrMsg, 48, "Custom Reports");
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
		for (iLoop = 1; iLoop <= (frmDefinition.grdAccess.Rows - 1); iLoop++) {
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

		frmSend.txtSend_parent1Table.value = frmDefinition.txtParent1ID.value;

		frmSend.txtSend_parent1AllRecords.value = "0";
		frmSend.txtSend_parent1Picklist.value = "0";
		frmSend.txtSend_parent1Filter.value = "0";
		if (frmDefinition.optParent1RecordSelection1.checked == true) {
			frmSend.txtSend_parent1AllRecords.value = "1";
		}
		if (frmDefinition.optParent1RecordSelection2.checked == true) {
			frmSend.txtSend_parent1Picklist.value = frmDefinition.txtParent1PicklistID.value;
		}
		if (frmDefinition.optParent1RecordSelection3.checked == true) {
			frmSend.txtSend_parent1Filter.value = frmDefinition.txtParent1FilterID.value;
		}

		frmSend.txtSend_parent2Table.value = frmDefinition.txtParent2ID.value;

		frmSend.txtSend_parent2AllRecords.value = "0";
		frmSend.txtSend_parent2Picklist.value = "0";
		frmSend.txtSend_parent2Filter.value = "0";
		if (frmDefinition.optParent2RecordSelection1.checked == true) {
			frmSend.txtSend_parent2AllRecords.value = "1";
		}
		if (frmDefinition.optParent2RecordSelection2.checked == true) {
			frmSend.txtSend_parent2Picklist.value = frmDefinition.txtParent2PicklistID.value;
		}
		if (frmDefinition.optParent2RecordSelection3.checked == true) {
			frmSend.txtSend_parent2Filter.value = frmDefinition.txtParent2FilterID.value;
		}

		/*now use the txtSend_childTable to hold the string of selected child tables*/
		if (frmUseful.txtChildsLoaded.value == 1) {
			frmSend.txtSend_childTable.value = getChildString();
		}
		else {
			var dataCollection = frmOriginalDefinition.elements;
			var iChildTableID = 0;
			var iChildFilterID = 0;
			var iChildRecords = 0;
			var sChildren = "";
			var iChildCounter = 0;

			if (dataCollection != null) {
				for (var i = 0; i < frmUseful.txtChildCount.value; i++) {

					iChildCounter = i + 1;
					iChildTableID = document.getElementById("txtReportDefnChildTableID_" + iChildCounter).value;
					iChildFilterID = document.getElementById("txtReportDefnChildFilterID_" + iChildCounter).value;
					iChildRecords = document.getElementById("txtReportDefnChildRecords_" + iChildCounter).value;

					sChildren = sChildren +
						iChildTableID + '||' +
							iChildFilterID + '||' +
								iChildRecords + '||';

					sChildren = sChildren + '**';
				}
			}
			frmSend.txtSend_childTable.value = sChildren;
		}

		// Selected columns, in order, with heading, size, decs, aggregates are done below.
		// Sort columns, in order, with asc/desc, boc, poc, voc, srv are done below.

		if (frmDefinition.chkPrintFilter.checked == true) {
			frmSend.txtSend_printFilterHeader.value = '1';
		}
		else {
			frmSend.txtSend_printFilterHeader.value = '0';
		}

		if (frmDefinition.chkSummary.checked == true) {
			frmSend.txtSend_summary.value = '1';
		}
		else {
			frmSend.txtSend_summary.value = '0';
		}

		// Ignore Zeros
		if (frmDefinition.chkIgnoreZeros.checked == true) {
			frmSend.txtSend_IgnoreZeros.value = '1';
		}
		else {
			frmSend.txtSend_IgnoreZeros.value = '0';
		}

		if (frmDefinition.chkPreview.checked == true) {
			frmSend.txtSend_OutputPreview.value = 1;
		}
		else {
			frmSend.txtSend_OutputPreview.value = 0;
		}

		if (frmDefinition.optOutputFormat0.checked) frmSend.txtSend_OutputFormat.value = 0;
		if (frmDefinition.optOutputFormat1.checked) frmSend.txtSend_OutputFormat.value = 1;
		if (frmDefinition.optOutputFormat2.checked) frmSend.txtSend_OutputFormat.value = 2;
		if (frmDefinition.optOutputFormat3.checked) frmSend.txtSend_OutputFormat.value = 3;
		if (frmDefinition.optOutputFormat4.checked) frmSend.txtSend_OutputFormat.value = 4;
		if (frmDefinition.optOutputFormat5.checked) frmSend.txtSend_OutputFormat.value = 5;
		if (frmDefinition.optOutputFormat6.checked) frmSend.txtSend_OutputFormat.value = 6;

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

		// now go through the columns grid (and sort order grid)(and the repetition grid)
		var sColumns = '';

		frmUseful.txtLockGridEvents.value = 1;
		readRepetitionStrings();
		readSortOrderStrings();

		if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
			frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
			frmDefinition.ssOleDBGridSelectedColumns.movefirst();

			for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {

				var iNum = new Number(i + 1);
				sColumns = sColumns + iNum +
					'||' + frmDefinition.ssOleDBGridSelectedColumns.columns("type").text +
						'||' + frmDefinition.ssOleDBGridSelectedColumns.columns("columnID").text +
							'||' + frmDefinition.ssOleDBGridSelectedColumns.columns("heading").text +
								'||' + frmDefinition.ssOleDBGridSelectedColumns.columns("size").text +
									'||' + frmDefinition.ssOleDBGridSelectedColumns.columns("decimals").text +
										'||' + frmDefinition.ssOleDBGridSelectedColumns.columns("numeric").text +
											'||';

				if (frmDefinition.ssOleDBGridSelectedColumns.columns("average").text == '1') {
					sColumns = sColumns + '1' + '||';
				}
				else {
					sColumns = sColumns + '0' + '||';
				}

				if (frmDefinition.ssOleDBGridSelectedColumns.columns("count").text == '1') {
					sColumns = sColumns + '1' + '||';
				}
				else {
					sColumns = sColumns + '0' + '||';
				}

				if (frmDefinition.ssOleDBGridSelectedColumns.columns("total").text == '1') {
					sColumns = sColumns + '1' + '||';
				}
				else {
					sColumns = sColumns + '0' + '||';
				}

				if (frmDefinition.ssOleDBGridSelectedColumns.columns("HiddenColumn").text == '1') {
					sColumns = sColumns + '1' + '||';
				}
				else {
					sColumns = sColumns + '0' + '||';
				}

				if (frmDefinition.ssOleDBGridSelectedColumns.columns("GroupWithNext").text == '1') {
					sColumns = sColumns + '1' + '||';
				}
				else {
					sColumns = sColumns + '0' + '||';
				}

				sColumns = sColumns + getSortOrderString(frmDefinition.ssOleDBGridSelectedColumns.columns("columnID").text);
				sColumns = sColumns + getRepetitionString(frmDefinition.ssOleDBGridSelectedColumns.Columns("type").Text + frmDefinition.ssOleDBGridSelectedColumns.columns("columnID").text) +
					'**';

				frmDefinition.ssOleDBGridSelectedColumns.movenext();
			}
			frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
		}
		else {
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
									'||' + selectedColumnParameter(dataCollection.item(iIndex).value, "HEADING") +
										'||' + selectedColumnParameter(dataCollection.item(iIndex).value, "SIZE") +
											'||' + selectedColumnParameter(dataCollection.item(iIndex).value, "DECIMALS") +
												'||' + selectedColumnParameter(dataCollection.item(iIndex).value, "NUMERIC") +
													'||';

						if (selectedColumnParameter(dataCollection.item(iIndex).value, "AVERAGE") == '1') {
							sColumns = sColumns + '1' + '||';
						}
						else {
							sColumns = sColumns + '0' + '||';
						}

						if (selectedColumnParameter(dataCollection.item(iIndex).value, "COUNT") == '1') {
							sColumns = sColumns + '1' + '||';
						}
						else {
							sColumns = sColumns + '0' + '||';
						}

						if (selectedColumnParameter(dataCollection.item(iIndex).value, "TOTAL") == '1') {
							sColumns = sColumns + '1' + '||';
						}
						else {
							sColumns = sColumns + '0' + '||';
						}

						if (selectedColumnParameter(dataCollection.item(iIndex).value, "HIDDENCOLUMN") == '1') {
							sColumns = sColumns + '1' + '||';
						}
						else {
							sColumns = sColumns + '0' + '||';
						}

						if (selectedColumnParameter(dataCollection.item(iIndex).value, "GROUPWITHNEXT") == '1') {
							sColumns = sColumns + '1' + '||';
						}
						else {
							sColumns = sColumns + '0' + '||';
						}

						sColumns = sColumns + getSortOrderString(selectedColumnParameter(dataCollection.item(iIndex).value, "COLUMNID"));
						sColumns = sColumns + getRepetitionString(selectedColumnParameter(dataCollection.item(iIndex).value, "TYPE") + selectedColumnParameter(dataCollection.item(iIndex).value, "COLUMNID")) +
							'**';

					}
				}
			}
		}

		frmUseful.txtLockGridEvents.value = 0;

		frmSend.txtSend_columns.value = sColumns.substr(0, 8000);
		frmSend.txtSend_columns2.value = sColumns.substr(8000, 8000);

		if (sColumns.length > 16000) {
			OpenHR.messageBox("Too many columns selected.", 48, "Custom Reports");
			return false;
		}
		else {
			return true;
		}
	}

	function populatePrinters() {

		var strCurrentPrinter = '';
		if (frmDefinition.cboPrinterName.selectedIndex > 0) {
			strCurrentPrinter = options[selectedIndex].innerText;
		}

		frmDefinition.cboPrinterName.length = 0;
		var oOption = document.createElement("OPTION");
		options.add(oOption);
		oOption.innerText = "<Default Printer>";
		oOption.value = 0;

		for (var iLoop = 0; iLoop < OpenHR.PrinterCount(); iLoop++) {

			oOption = document.createElement("OPTION");
			options.add(oOption);
			oOption.innerText = OpenHR.PrinterName(iLoop);
			oOption.value = iLoop + 1;

			if (oOption.innerText == strCurrentPrinter) {
				frmDefinition.cboPrinterName.selectedIndex = iLoop + 1;
			}
		}


		if (strCurrentPrinter != '') {
			if (frmDefinition.cboPrinterName.options(frmDefinition.cboPrinterName.selectedIndex).innerText != strCurrentPrinter) {
				oOption = document.createElement("OPTION");
				frmDefinition.cboPrinterName.options.add(oOption);
				oOption.innerText = strCurrentPrinter;
				oOption.value = frmDefinition.cboPrinterName.options.length - 1;
				frmDefinition.cboPrinterName.selectedIndex = oOption.value;
			}
		}
	}

	function populateSaveExisting() {
		
		var lngCurrentOption = 0;
		if (frmDefinition.cboSaveExisting.selectedIndex > 0) {
			lngCurrentOption = options[selectedIndex].value;
		}
		frmDefinition.cboSaveExisting.length = 0;

		var oOption = document.createElement("OPTION");
		options.add(oOption);
		oOption.innerText = "Overwrite";
		oOption.value = 0;

		oOption = document.createElement("OPTION");
		options.add(oOption);
		oOption.innerText = "Do not overwrite";
		oOption.value = 1;

		oOption = document.createElement("OPTION");
		options.add(oOption);
		oOption.innerText = "Add sequential number to name";
		oOption.value = 2;

		oOption = document.createElement("OPTION");
		options.add(oOption);
		oOption.innerText = "Append to file";
		oOption.value = 3;

		if ((frmDefinition.optOutputFormat4.checked) || (frmDefinition.optOutputFormat5.checked) || (frmDefinition.optOutputFormat6.checked)) {
			oOption = document.createElement("OPTION");
			options.add(oOption);
			oOption.innerText = "Create new sheet in workbook";
			oOption.value = 4;
		}

		for (var iLoop = 0; iLoop < options.length; iLoop++) {
			if (options(iLoop).value == lngCurrentOption) {
				frmDefinition.cboSaveExisting.selectedIndex = iLoop;
				break;
			}
		}
	}

	function getChildString() {
		var sChilds = "";
		var i;
		var pvarbookmark;
		var sALL_RECORDS = "All Records";

		with (frmDefinition.ssOleDBGridChildren) {
			if (Rows > 0) {
				MoveFirst();
				for (var i = 0; i < Rows; i++) {

					sChilds = sChilds + Columns('TableID').text + '||';

					//add the FilterID to the string.
					if (Columns('Table').text != '') {
						sChilds = sChilds + Columns('FilterID').text + '||';
					}
					else {
						sChilds = sChilds + 0 + '||';
					}

					//add the OrderID to the string.
					if (Columns('Table').text != '') {
						sChilds = sChilds + Columns('OrderID').text + '||';
					}
					else {
						sChilds = sChilds + 0 + '||';
					}

					//add the Records to the string.
					if (Columns('Records').text != sALL_RECORDS) {
						sChilds = sChilds + Columns('Records').text + '||';
					}
					else {
						sChilds = sChilds + 0 + '||';
					}

					sChilds = sChilds + '**';
					MoveNext();
				}
				return sChilds;
			}
			else {
				return "";
			}
		}
	}

	function readRepetitionStrings() {
		var i;
		var iNum;
		var sTemp = '';
		var bm;

		sRepDefn = "||";

		if (frmUseful.txtRepetitionLoaded.value == 1) {
			frmDefinition.ssOleDBGridRepetition.Redraw = false;
			frmDefinition.ssOleDBGridRepetition.movefirst();

			for (i = 0; i < frmDefinition.ssOleDBGridRepetition.rows; i++) {

				bm = frmDefinition.ssOleDBGridRepetition.GetBookmark(i);

				var sColumnID = new String(frmDefinition.ssOleDBGridRepetition.Columns(0).celltext(bm));

				if (frmDefinition.ssOleDBGridRepetition.columns(2).cellvalue(bm) == '-1'
					|| frmDefinition.ssOleDBGridRepetition.columns(2).cellvalue(bm) == '1') {
					sTemp = '1' + '||';
				}
				else {
					sTemp = '0' + '||';
				}
				sRepDefn = sRepDefn + "|" + sColumnID + "|||" + sTemp;
			}

			frmDefinition.ssOleDBGridRepetition.movefirst();
			frmDefinition.ssOleDBGridRepetition.redraw = true;
		}
		else {
			iNum = 0;
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {

					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 24);
					if (sControlName == "txtReportDefnRepetition_") {
						iNum = iNum + 1;
						sDefn = new String(dataCollection.item(i).value);
						var sColumnID = new String(repetitionColumnParameter(sDefn, "COLUMNID"));

						if (repetitionColumnParameter(sDefn, "REPETITION") == '1') {
							sTemp = '1' + '||';
						}
						else {
							sTemp = '0' + '||';
						}

						sRepDefn = sRepDefn + "|" + sColumnID + "|||" + sTemp;
					}
				}
			}
		}
	}

	function getRepetitionString(piColumnID) {
		var sTemp;
		var sColID = new String(piColumnID);
		var iIndex;
		var iIndex2;

		sTemp = "|||" + sColID + "|||";

		iIndex = sRepDefn.indexOf(sTemp);

		if (iIndex >= 0) {
			iIndex2 = sRepDefn.indexOf("|||", iIndex + sTemp.length);
			if (iIndex2 >= 0) {
				sTemp = sRepDefn.substr(iIndex + sTemp.length, iIndex2 - iIndex - sTemp.length + 2);
			}
			else {
				sTemp = sRepDefn.substr(iIndex + sTemp.length);
			}
			return (sTemp);
		}
		else {
			return ('-1||');
		}
	}

	function readSortOrderStrings() {
		var i;
		var iNum;
		var sTemp = '';

		sSortDefn = "||";

		if (frmUseful.txtSortLoaded.value == 1) {
			frmDefinition.ssOleDBGridSortOrder.redraw = false;
			frmDefinition.ssOleDBGridSortOrder.movefirst();

			for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) {

				var iNum = new Number(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.bookmark) + 1);

				sTemp = iNum + '||' +
					frmDefinition.ssOleDBGridSortOrder.columns("order").text + '||';

				if (frmDefinition.ssOleDBGridSortOrder.columns("break").text == '-1') {
					sTemp = sTemp + '1' + '||';
				}
				else {
					sTemp = sTemp + '0' + '||';
				}

				if (frmDefinition.ssOleDBGridSortOrder.columns("page").text == '-1') {
					sTemp = sTemp + '1' + '||';
				}
				else {
					sTemp = sTemp + '0' + '||';
				}

				if (frmDefinition.ssOleDBGridSortOrder.columns("value").text == '-1') {
					sTemp = sTemp + '1' + '||';
				}
				else {
					sTemp = sTemp + '0' + '||';
				}

				if (frmDefinition.ssOleDBGridSortOrder.columns("hide").text == '-1') {
					sTemp = sTemp + '1' + '||';
				}
				else {
					sTemp = sTemp + '0' + '||';
				}

				sSortDefn = sSortDefn + "|" + frmDefinition.ssOleDBGridSortOrder.Columns(0).text + "|||" + sTemp;

				frmDefinition.ssOleDBGridSortOrder.movenext();
			}

			frmDefinition.ssOleDBGridSortOrder.redraw = true;
		}
		else {
			iNum = 0;
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {
				
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 19);
					if (sControlName == "txtReportDefnOrder_") {
						iNum = iNum + 1;
						sDefn = new String(dataCollection.item(i).value);

						sTemp = iNum + '||' +
							sortColumnParameter(sDefn, "ORDER") + '||';

						if (sortColumnParameter(sDefn, "BOC") == '-1') {
							sTemp = sTemp + '1' + '||';
						}
						else {
							sTemp = sTemp + '0' + '||';
						}

						if (sortColumnParameter(sDefn, "POC") == '-1') {
							sTemp = sTemp + '1' + '||';
						}
						else {
							sTemp = sTemp + '0' + '||';
						}

						if (sortColumnParameter(sDefn, "VOC") == '-1') {
							sTemp = sTemp + '1' + '||';
						}
						else {
							sTemp = sTemp + '0' + '||';
						}

						if (sortColumnParameter(sDefn, "SRV") == '-1') {
							sTemp = sTemp + '1' + '||';
						}
						else {
							sTemp = sTemp + '0' + '||';
						}

						sSortDefn = sSortDefn + "|" + sortColumnParameter(sDefn, "COLUMNID") + "|||" + sTemp;
					}
				}
			}
		}
	}

	function getSortOrderString(piColumnID) {
		var sTemp;
		var sColID = new String(piColumnID);
		var iIndex;
		var iIndex2;

		sTemp = "|||" + sColID + "|||";

		iIndex = sSortDefn.indexOf(sTemp);

		if (iIndex >= 0) {
			iIndex2 = sSortDefn.indexOf("|||", iIndex + sTemp.length);
			if (iIndex2 >= 0) {
				sTemp = sSortDefn.substr(iIndex + sTemp.length, iIndex2 - iIndex - sTemp.length + 2);
			}
			else {
				sTemp = sSortDefn.substr(iIndex + sTemp.length);
			}

			return (sTemp);
		}
		else {
			return ('0||0||0||0||0||0||');
		}
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
			// assign the tab delimited string of selected child table ids 
			frmGetDataForm.txtReportChildTableID.value = childTableString();

			data_refreshData();

			frmUseful.txtTablesChanged.value = 0;
		}

		sSelectedIDs = selectedIDs();

		frmDefinition.ssOleDBGridAvailableColumns.RemoveAll();

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

		refreshTab3Controls();

		// Get menu.asp to refresh the menu.
		OpenHR.refreshMenu();
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
						}
						else {
							frmDefinition.ssOleDBGridSelectedColumns.columns(7).text = '0';
						}

						break;
					}
				}
			}

			frmDefinition.ssOleDBGridSelectedColumns.movenext();
		}

		frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;

		refreshTab3Controls();

		// Get menu.asp to refresh the menu.
		OpenHR.refreshMenu();

		//have added this as the available columns data has be wiped.
		frmUseful.txtTablesChanged.value = 1;
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
		frmDefinition.chkPrintFilter.checked = ((frmOriginalDefinition.txtDefn_PrintFilterHeader.value != "False") &&
			((frmOriginalDefinition.txtDefn_FilterID.value > 0) || (frmOriginalDefinition.txtDefn_PicklistID.value > 0)));

		// Set the parent1 selection
		fRecordOptionSet = false;

		if (frmOriginalDefinition.txtDefn_Parent1TableID.value > 0) {
			frmDefinition.txtParent1ID.value = frmOriginalDefinition.txtDefn_Parent1TableID.value;
			frmDefinition.txtParent1.value = frmOriginalDefinition.txtDefn_Parent1TableName.value;
			button_disable(frmDefinition.cmdParent1Filter, false);

			if (frmOriginalDefinition.txtDefn_Parent1PicklistID.value > 0) {
				button_disable(frmDefinition.cmdParent1Picklist, false);
				frmDefinition.optParent1RecordSelection2.checked = true;
				frmDefinition.txtParent1PicklistID.value = frmOriginalDefinition.txtDefn_Parent1PicklistID.value;
				frmDefinition.txtParent1Picklist.value = frmOriginalDefinition.txtDefn_Parent1PicklistName.value;
				fRecordOptionSet = true;
			}
			else {
				if (frmOriginalDefinition.txtDefn_Parent1FilterID.value > 0) {
					button_disable(frmDefinition.cmdParent1Filter, false);
					frmDefinition.optParent1RecordSelection3.checked = true;
					frmDefinition.txtParent1FilterID.value = frmOriginalDefinition.txtDefn_Parent1FilterID.value;
					frmDefinition.txtParent1Filter.value = frmOriginalDefinition.txtDefn_Parent1FilterName.value;
					fRecordOptionSet = true;
				}
			}
		}
		if (fRecordOptionSet == false) {
			frmDefinition.optParent1RecordSelection1.checked = true;
		}

		// Set the parent2 selection
		fRecordOptionSet = false;

		if (frmOriginalDefinition.txtDefn_Parent2TableID.value > 0) {
			frmDefinition.txtParent2ID.value = frmOriginalDefinition.txtDefn_Parent2TableID.value;
			frmDefinition.txtParent2.value = frmOriginalDefinition.txtDefn_Parent2TableName.value;
			button_disable(frmDefinition.cmdParent2Filter, false);

			if (frmOriginalDefinition.txtDefn_Parent2PicklistID.value > 0) {
				button_disable(frmDefinition.cmdParent2Picklist, false);
				frmDefinition.optParent2RecordSelection2.checked = true;
				frmDefinition.txtParent2PicklistID.value = frmOriginalDefinition.txtDefn_Parent2PicklistID.value;
				frmDefinition.txtParent2Picklist.value = frmOriginalDefinition.txtDefn_Parent2PicklistName.value;
				fRecordOptionSet = true;
			}
			else {
				if (frmOriginalDefinition.txtDefn_Parent2FilterID.value > 0) {
					button_disable(frmDefinition.cmdParent2Filter, false);
					frmDefinition.optParent2RecordSelection3.checked = true;
					frmDefinition.txtParent2FilterID.value = frmOriginalDefinition.txtDefn_Parent2FilterID.value;
					frmDefinition.txtParent2Filter.value = frmOriginalDefinition.txtDefn_Parent2FilterName.value;
					fRecordOptionSet = true;
				}
			}
		}
		if (fRecordOptionSet == false) {
			frmDefinition.optParent2RecordSelection1.checked = true;
		}

		if ((frmOriginalDefinition.txtDefn_PicklistHidden.value.toUpperCase() == "TRUE") ||
			(frmOriginalDefinition.txtDefn_FilterHidden.value.toUpperCase() == "TRUE")) {
			frmSelectionAccess.baseHidden.value = "Y";
		}
		if ((frmOriginalDefinition.txtDefn_Parent1PicklistHidden.value.toUpperCase() == "TRUE") ||
			(frmOriginalDefinition.txtDefn_Parent1FilterHidden.value.toUpperCase() == "TRUE")) {
			frmSelectionAccess.p1Hidden.value = "Y";
		}
		if ((frmOriginalDefinition.txtDefn_Parent2PicklistHidden.value.toUpperCase() == "TRUE") ||
			(frmOriginalDefinition.txtDefn_Parent2FilterHidden.value.toUpperCase() == "TRUE")) {
			frmSelectionAccess.p2Hidden.value = "Y";
		}

		frmSelectionAccess.childHidden.value = frmUseful.txtHiddenChildFilterCount.value;

		frmSelectionAccess.calcsHiddenCount.value = frmOriginalDefinition.txtDefn_HiddenCalcCount.value;

		// Summary report ?
		frmDefinition.chkSummary.checked = (frmOriginalDefinition.txtDefn_Summary.value != "False");

		// Ignore Zeros ?
		frmDefinition.chkIgnoreZeros.checked = (frmOriginalDefinition.txtDefn_IgnoreZeros.value != "False");

		//OUTPUT OPTIONS?

		frmDefinition.chkPreview.checked = (frmOriginalDefinition.txtDefn_OutputPreview.value != "False");

		if (frmOriginalDefinition.txtDefn_OutputFormat.value == 0) {
			frmDefinition.optOutputFormat0.checked = true;
		}
		else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 1) {
			frmDefinition.optOutputFormat1.checked = true;
		}
		else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 2) {
			frmDefinition.optOutputFormat2.checked = true;
		}
		else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 3) {
			frmDefinition.optOutputFormat3.checked = true;
		}
		else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 4) {
			frmDefinition.optOutputFormat4.checked = true;
		}
		else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 5) {
			frmDefinition.optOutputFormat5.checked = true;
		}
		else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 6) {
			frmDefinition.optOutputFormat6.checked = true;
		}
		else {
			frmDefinition.optOutputFormat0.checked = true;
		}

		frmDefinition.chkDestination0.checked = (frmOriginalDefinition.txtDefn_OutputScreen.value != "False");
		frmDefinition.chkDestination1.checked = (frmOriginalDefinition.txtDefn_OutputPrinter.value != "False");

		if (frmDefinition.chkDestination1.checked == true) {
			// OpenHR.messageBox("Printer functionality not yet added",48,"Calendar Reports");
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

		frmDefinition.ssOleDBGridSelectedColumns.MoveFirst();
		frmDefinition.ssOleDBGridSelectedColumns.FirstRow = frmDefinition.ssOleDBGridSelectedColumns.Bookmark;

		frmDefinition.ssOleDBGridSortOrder.movefirst();
		frmDefinition.ssOleDBGridSortOrder.FirstRow = frmDefinition.ssOleDBGridSortOrder.bookmark;

		// If its read only, disable everything.
		if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
			disableAll();
		}

	}

	function loadChildTables() {
		
		if (frmUseful.txtChildsLoaded.value == 0) {
			var iChildCounter = 0;

			with (frmDefinition.ssOleDBGridChildren) {
				if (Enabled) {
					focus();
				}
				for (var i = 0; i < frmUseful.txtChildCount.value; i++) {
					iChildCounter = i + 1;
					var sAdd = new String(document.getElementById('txtReportDefnChildGridString_' + iChildCounter).value);
					AddItem(sAdd);
				}
			}
			frmUseful.txtChildsLoaded.value = 1;
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
		var sKey;
		
		if (frmUseful.txtSelectedColumnsLoaded.value == 0) {

			frmDefinition.ssOleDBGridSelectedColumns.focus();

			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
					
					var sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 20);
					if (sControlName == "txtReportDefnColumn_") {
						sDefnString = new String(dataCollection.item(iIndex).value);
						frmDefinition.ssOleDBGridSelectedColumns.AddItem(sDefnString);

						if (sDefnString.length > 0) {
							sKey = selectedColumnParameter(sDefnString, 'TYPE') + selectedColumnParameter(sDefnString, 'COLUMNID');
							if (selectedColumnParameter(sDefnString, 'HIDDENCOLUMN') == "1") {
								addGridCol(sKey + '	-1	0	0	0	0	0');
							}
							else {
								addGridCol(sKey + '	0	0	0	0	0	0');
							}

						}
					}
				}
			}

			frmUseful.txtSelectedColumnsLoaded.value = 1;
		}
	}

	function loadSortDefinition() {
		var iIndex;
		var sKey;
		
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
							sKey = 'C' + sortColumnParameter(sDefnString, 'COLUMNID');
							frmDefinition.ssOleDBGridSortOrder.AddItem(sDefnString);

							if (setGirdCol(sKey)) {
								if (sortColumnParameter(sDefnString, "BOC") == "-1") {
									updateCurrentColProp('break', true);
								}
								else {
									updateCurrentColProp('break', false);
								}
								if (sortColumnParameter(sDefnString, "POC") == "-1") {
									updateCurrentColProp('page', true);
								}
								else {
									updateCurrentColProp('page', false);
								}
								if (sortColumnParameter(sDefnString, "VOC") == "-1") {
									updateCurrentColProp('value', true);
								}
								else {
									updateCurrentColProp('value', false);
								}
								if (sortColumnParameter(sDefnString, "SRV") == "-1") {
									updateCurrentColProp('hide', true);
								}
								else {
									updateCurrentColProp('hide', false);
								}
							}
						}
					}
				}
			}

			frmUseful.txtSortLoaded.value = 1;
		}
	}

	function loadRepetitionDefinition() {
		var iIndex;
		var sKey;

		if (frmUseful.txtRepetitionLoaded.value == 0) {
			frmDefinition.ssOleDBGridRepetition.focus();
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
				
					var sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 24);
					if (sControlName == "txtReportDefnRepetition_") {
						var sDefnString = new String(dataCollection.item(iIndex).value);

						if (sDefnString.length > 0) {
							sKey = sortColumnParameter(sDefnString, 'COLUMNID');

							frmDefinition.ssOleDBGridRepetition.AddItem(sDefnString);

							if (setGirdCol(sKey)) {
								if (repetitionColumnParameter(sDefnString, "REPETITION") == "1") {
									updateCurrentColProp('repetition', true);
								}
								else {
									updateCurrentColProp('repetition', false);
								}
							}
						}
					}
				}
			}
			frmUseful.txtRepetitionLoaded.value = 1;
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

	function saveFile() {
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
			dialog.Filter = frmDefinition.txtWordFormats.value;
			dialog.FilterIndex = frmDefinition.txtWordFormatDefaultIndex.value;

		}

		else {
			//EXCEL
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
		refreshTab5Controls();

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
															if (psParameter == "HIDDENCOLUMN") return sDefn.substr(0, iCharIndex);
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

	function repetitionColumnParameter(psDefnString, psParameter) {
		var iCharIndex;
		var sDefn;

		sDefn = new String(psDefnString);
		iCharIndex = sDefn.indexOf("	");
		if (iCharIndex >= 0) {
			if (psParameter == "COLUMNID") return sDefn.substr(0, iCharIndex);
			sDefn = sDefn.substr(iCharIndex + 1);
			iCharIndex = sDefn.indexOf("	");
			if (iCharIndex >= 0) {
				if (psParameter == "COLUMN") return sDefn.substr(0, iCharIndex);
				sDefn = sDefn.substr(iCharIndex + 1);
				iCharIndex = sDefn.indexOf("	");
				if (iCharIndex >= 0) {
					if (psParameter == "REPETITION") return sDefn.substr(0, iCharIndex);
					sDefn = sDefn.substr(iCharIndex + 1);
					iCharIndex = sDefn.indexOf("	");
					if (iCharIndex >= 0) {
						if (psParameter == "TABLEID") return sDefn.substr(0, iCharIndex);
						sDefn = sDefn.substr(iCharIndex + 1);

						if (psParameter == "HIDDEN") return sDefn;
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

	function removeRepetitionColumn(piColumnID, piTableID) {
		// Remove the column (if columnID given), 
		// or all columns for a table (if tableID given),
		// or all columns (if no columnID or tableID given).
		// from the repetition columns definition.
		var iCount;
		var i;
		var fRemoveRow;

		if (frmUseful.txtRepetitionLoaded.value == 1) {
			if (frmDefinition.ssOleDBGridRepetition.Rows > 0) {
				frmDefinition.ssOleDBGridRepetition.Redraw = false;
				frmDefinition.ssOleDBGridRepetition.MoveFirst();

				iCount = frmDefinition.ssOleDBGridRepetition.Rows;
				for (i = 0; i < iCount; i++) {
					fRemoveRow = true;

					if (piColumnID > 0) {
						var sColID = new String(frmDefinition.ssOleDBGridRepetition.Columns("ColumnID").Text);
						fRemoveRow = (piColumnID == sColID.substring(1));
					}

					if (piTableID > 0) {
						fRemoveRow = (piTableID == frmDefinition.ssOleDBGridRepetition.Columns("tableID").Text);
					}

					if (fRemoveRow == true) {
						if (frmDefinition.ssOleDBGridRepetition.rows == 1) {
							frmDefinition.ssOleDBGridRepetition.RemoveAll();
						}
						else {
							frmDefinition.ssOleDBGridRepetition.RemoveItem(frmDefinition.ssOleDBGridRepetition.AddItemRowIndex(frmDefinition.ssOleDBGridRepetition.Bookmark));
						}
					}
					else {
						frmDefinition.ssOleDBGridRepetition.MoveNext();
					}
				}

				frmDefinition.ssOleDBGridRepetition.Redraw = true;
				frmDefinition.ssOleDBGridRepetition.SelBookmarks.RemoveAll();

				if (frmDefinition.ssOleDBGridRepetition.Rows > 0) {
					frmDefinition.ssOleDBGridRepetition.MoveFirst();
				}
			}
		}
		else {
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection != null) {
				for (var iIndex = 0; iIndex < dataCollection.length; iIndex++) {
					var sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 24);
					if (sControlName == "txtReportDefnRepetition_") {
						fRemoveRow = true;
						if (piColumnID > 0) {
							var sColID = new String(repetitionColumnParameter(dataCollection.item(iIndex).value, "COLUMNID"));
							fRemoveRow = (piColumnID == sColID.substring(1));
						}

						if (piTableID > 0) {
							fRemoveRow = (piTableID == repetitionColumnParameter(dataCollection.item(iIndex).value, "TABLEID"));
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
			}
			else {
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
			}
			else {
				var dataCollection = frmOriginalDefinition.elements;
				if (dataCollection != null) {
					for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
						sControlName = dataCollection.item(iIndex).name;
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


	function removeFilters(psChildFilters) {
		var iCharIndex;
		var sChildFilters;
		var sChildFilterID;
		var sGridChildFilterID;
		var fDone;

		sChildFilters = new String(psChildFilters);

		// Remove the given calcs from the selected columns list.
		while (sChildFilters.length > 0) {
			iCharIndex = sChildFilters.indexOf(",");

			if (iCharIndex >= 0) {
				sChildFilterID = sChildFilters.substr(0, iCharIndex);
				sChildFilters = sChildFilters.substr(iCharIndex + 1);
			}
			else {
				sChildFilterID = sChildFilters;
				sChildFilters = "";
			}

			fDone = false;

			/* Check if we're removing the base table first, then paretn1 then parent 2, and then the children. */
			if ((fDone == false) && (frmDefinition.txtBaseFilterID.value == sChildFilterID)) {
				frmDefinition.txtBaseFilter.value = '';
				frmDefinition.txtBaseFilterID.value = 0;
				frmSelectionAccess.baseHidden.value = "N";
				fDone = true;
			}

			if ((fDone == false) && (frmDefinition.txtParent1FilterID.value == sChildFilterID)) {
				frmDefinition.txtParent1Filter.value = '';
				frmDefinition.txtParent1FilterID.value = 0;
				frmSelectionAccess.p1Hidden.value = 'N';
				fDone = true;
			}

			if ((fDone == false) && (frmDefinition.txtParent2FilterID.value == sChildFilterID)) {
				frmDefinition.txtParent2Filter.value = '';
				frmDefinition.txtParent2FilterID.value = 0;
				frmSelectionAccess.p2Hidden.value = 'N';
				fDone = true;
			}

			if ((fDone == false) && (frmUseful.txtChildsLoaded.value == 1)) {
				if (frmDefinition.ssOleDBGridChildren.Rows > 0) {
					frmDefinition.ssOleDBGridChildren.Redraw = false;
					frmDefinition.ssOleDBGridChildren.movefirst();

					for (i = 0; i < frmDefinition.ssOleDBGridChildren.rows; i++) {
						sGridChildFilterID = frmDefinition.ssOleDBGridChildren.Columns("FilterID").Text;

						if (sGridChildFilterID == sChildFilterID) {
							frmDefinition.ssOleDBGridChildren.Columns("FilterID").Text = 0;
							frmDefinition.ssOleDBGridChildren.Columns("Filter").Text = "";
							frmDefinition.ssOleDBGridChildren.Columns("FilterHidden").Text == "Y";
							break;
						}

						frmDefinition.ssOleDBGridChildren.movenext();
					}

					frmDefinition.ssOleDBGridChildren.Redraw = true;
				}
			}
		}

		refreshTab2Controls();
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

			if ((fDone == false) && (frmDefinition.txtParent1PicklistID.value == sPicklistID)) {
				frmDefinition.txtParent1Picklist.value = '';
				frmDefinition.txtParent1PicklistID.value = 0;
				frmSelectionAccess.p1Hidden.value = 'N';
				fDone = true;
			}

			if ((fDone == false) && (frmDefinition.txtParent2PicklistID.value == sPicklistID)) {
				frmDefinition.txtParent2Picklist.value = '';
				frmDefinition.txtParent2PicklistID.value = 0;
				frmSelectionAccess.p2Hidden.value = 'N';
				fDone = true;
			}
		}

		refreshTab2Controls();
	}

	function removeOrders(psChildOrders) {
		var iCharIndex;
		var sChildOrders;
		var sChildOrderID;
		var sGridChildOrderID;

		sChildOrders = new String(psChildOrders);

		// Remove the given calcs from the selected columns list.
		while (sChildOrders.length > 0) {
			iCharIndex = sChildFilters.indexOf(",");

			if (iCharIndex >= 0) {
				sChildOrderID = sChildOrders.substr(0, iCharIndex);
				sChildOrders = sChildOrders.substr(iCharIndex + 1);
			}
			else {
				sChildOrderID = sChildOrders;
				sChildOrders = "";
			}

			if (frmUseful.txtChildsLoaded.value == 1) {
				if (frmDefinition.ssOleDBGridChildren.Rows > 0) {
					frmDefinition.ssOleDBGridChildren.Redraw = false;
					frmDefinition.ssOleDBGridChildren.movefirst();

					for (i = 0; i < frmDefinition.ssOleDBGridChildren.rows; i++) {
						sGridChildOrderID = frmDefinition.ssOleDBGridChildren.Columns("OrderID").Text;

						if (sGridChildOrderID == sChildOrderID) {
							frmDefinition.ssOleDBGridChildren.RemoveItem(frmDefinition.ssOleDBGridChildren.AddItemRowIndex(frmDefinition.ssOleDBGridChildren.Bookmark));

							frmDefinition.ssOleDBGridChildren.Columns("OrderID").Text = 0;
							frmDefinition.ssOleDBGridChildren.Columns("Order").Text = "";
							break;
						}

						frmDefinition.ssOleDBGridChildren.movenext();
					}

					frmDefinition.ssOleDBGridChildren.Redraw = true;
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

		frmDefinition.ssOleDBGridAvailableColumns.redraw = false;

		frmDefinition.ssOleDBGridAvailableColumns.MoveLast();
		frmDefinition.ssOleDBGridAvailableColumns.MoveFirst();

		frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.removeall();

		for (var iIndex = 1; iIndex <= frmDefinition.ssOleDBGridAvailableColumns.rows; iIndex++) {
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

	function fileClear() {
		return;
		frmDefinition.txtSaveFile.value = "";
		refreshTab5Controls();
	}

	function refreshFile() {
		return;
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

	function recalcHiddenChildFiltersCount() {
		var iCount;
		var vBM;

		iCount = 0;

		for (var i = 0; i < frmDefinition.ssOleDBGridChildren.Rows; i++) {
			vBM = frmDefinition.ssOleDBGridChildren.AddItemBookmark(i);

			if (frmDefinition.ssOleDBGridChildren.Columns("FilterHidden").CellValue(vBM) == "Y") {
				iCount = iCount + 1;
			}
		}
		frmSelectionAccess.childHidden.value = iCount;
	}
	
	function ssOleDBGridAvailableColumns_RowColChange(LastRow,LastCol) {
		refreshTab3Controls();	
	}
	
	function ssOleDBGridAvailableColumns_DblClick() {
		columnSwap(true);	
	}
	
	function ssOleDBGridAvailableColumns_KeyPress(iKeyAscii) {
		
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

	function ssOleDBGridSelectedColumns_RowColChange(LastRow,LastCol) {
		if (frmUseful.txtLockGridEvents.value != 1) {
			refreshTab3Controls();
		}
	}

	function ssOleDBGridSelectedColumns_DblClick() {
		columnSwap(false);	
	}
	
	function ssOleDBGridSelectedColumns_SelChange() {
		refreshTab3Controls();		
	}
	
	function ssOleDBGridSortOrder_BeforeUpdate() {
		
		if ((frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Asc') && (frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Desc')) {
			frmDefinition.ssOleDBGridSortOrder.Columns(2).text = 'Asc';
		}
	}
	
	function ssOleDBGridSortOrder_AfterInsert() {
		refreshTab4Controls();		
	}
	
	function ssOleDBGridSortOrder_RowLoaded(Bookmark) {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if (fViewing) 
		{
			return;
		}
	
		with (frmDefinition.ssOleDBGridSortOrder)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol('C' + sKey))
			{
				var bBreak = getCurrentColProp('Break');
				var bPage = getCurrentColProp('Page');
				var bHidden = getCurrentColProp('Hidden');
				var bRepetition = getCurrentColProp('Repetition');

				if (bBreak == true)
				{
					Columns("page").CellStyleSet("ssetFixDataDisabled");
				}
				else
				{
					Columns("page").CellStyleSet("ssetFixData");
				}

				if (bPage == true)
				{
					Columns("break").CellStyleSet("ssetFixDataDisabled");
				}
				else
				{
					Columns("break").CellStyleSet("ssetFixData");
				}
			
				if (bHidden == true)
				{
					Columns("value").CellStyleSet("ssetFixDataDisabled", AddItemRowIndex(Bookmark));
				}
				else
				{
					Columns("value").CellStyleSet("ssetFixData", AddItemRowIndex(Bookmark));
				}
								
				if ((bHidden == true) || (bRepetition == true))
				{
					Columns("hide").CellStyleSet("ssetFixDataDisabled", AddItemRowIndex(Bookmark));
				}
				else
				{
					Columns("hide").CellStyleSet("ssetFixData", AddItemRowIndex(Bookmark));
				}		
			}
		} 
	}
	
	function ssOleDBGridSortOrder_RowColChange(LastRow,LastCol) {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if (fViewing) 
		{
			return;
		}
    
		var fSortAddDisabled = false;
		var fSortEditDisabled = false;
		var fSortRemoveDisabled = false;
		var fSortRemoveAllDisabled = false;
		var fSortMoveUpDisabled = false;
		var fSortMoveDownDisabled = false;

		with (frmDefinition.ssOleDBGridSortOrder)
		{
			frmSortOrder.txtSortColumnID.value = Columns(0).text;
			frmSortOrder.txtSortColumnName.value = Columns(1).text;
			frmSortOrder.txtSortOrder.value = Columns(2).text;
			frmSortOrder.txtSortBOC.value = Columns(3).text;
			frmSortOrder.txtSortPOC.value = Columns(4).text;
			frmSortOrder.txtSortVOC.value = Columns(5).text;
			frmSortOrder.txtSortSRV.value = Columns(6).text;

			for (var i=0; i<Rows; i++)
			{
				Columns("column").CellStyleSet("ssetFixData", i);
				Columns("order").CellStyleSet("ssetFixData", i);
			
				var sKey = Columns("ColumnID").CellValue(AddItemBookmark(i));
				if (setGirdCol('C' + sKey))
				{
					var bBreak = getCurrentColProp('Break');
					var bPage = getCurrentColProp('Page');
					var bHidden = getCurrentColProp('Hidden');
					var bRepetition = getCurrentColProp('Repetition');
			
					if (bBreak == true)
					{
						Columns("page").CellStyleSet("ssetFixDataDisabled", i);
					}
					else
					{
						Columns("page").CellStyleSet("ssetFixData", i);
					}

					if (bPage == true)
					{
						Columns("break").CellStyleSet("ssetFixDataDisabled", i);
					}
					else
					{
						Columns("break").CellStyleSet("ssetFixData", i);
					}
			
					if (bHidden == true)
					{
						Columns("value").CellStyleSet("ssetFixDataDisabled", i);
					}
					else
					{
						Columns("value").CellStyleSet("ssetFixData", i);
					}
								
					if ((bHidden == true) || (bRepetition == true))
					{
						Columns("hide").CellStyleSet("ssetFixDataDisabled", i);
					}
					else
					{
						Columns("hide").CellStyleSet("ssetFixData", i);
					}
				}
			}

			Refresh();    
		
			/*if (Col == 1)
				{
				Col = 0;
				}*/
			
			SelBookmarks.RemoveAll();
			SelBookmarks.Add(Bookmark);
		
			if (frmUseful.txtSelectedColumnsLoaded.value == 1) 
			{
				if (frmDefinition.ssOleDBGridSelectedColumns.Rows <= Rows)	
				{
					//Disable 'Add' if there are no more columns to sort by.
					fSortAddDisabled = true;
				}
			}
			else 
			{
				iCount = 0;
				var dataCollection = frmOriginalDefinition.elements;
				if (dataCollection!=null) 
				{
					for (i=0; i<dataCollection.length; i++)  
					{
						sControlName = dataCollection.item(i).name;
						sControlName = sControlName.substr(0, 20);
						if (sControlName == "txtReportDefnColumn_") 
						{
							iCount = iCount + 1;
						}				
					}	
				}

				if (iCount <= Rows)	
				{
					// Disable 'Add' if there are no more columns to sort by.
					fSortAddDisabled = true;
				}	
			}
	
			if ((SelBookmarks.Count == 1) 
				&& (Rows > 0))
			{
				// Are we on the top row ?
				if ((AddItemRowIndex(Bookmark) == 0) 
					|| (Rows <= 1))
				{
					fSortMoveUpDisabled = true; 
				}

				// Are we on the bottom row ?
				if ((AddItemRowIndex(Bookmark) == Rows - 1) 
					|| (Rows <= 1))
				{
					fSortMoveDownDisabled = true; 
				}
			}	

			if ((Rows < 1) 
				|| (SelBookmarks.Count != 1))
			{
				fSortMoveUpDisabled = true; 
				fSortMoveDownDisabled = true; 
			}

			if (fViewing)
			{
				fSortAddDisabled = true;
				fSortMoveUpDisabled = true; 
				fSortMoveDownDisabled = true; 
			}

			fSortRemoveDisabled = ((SelBookmarks.Count!=1) 
				|| (fViewing == true)
					|| (Rows<1));
			fSortRemoveAllDisabled = ((fViewing == true)
				|| (Rows<1));
			fSortEditDisabled = ((SelBookmarks.Count!=1) 
				|| (fViewing == true)
					|| (Rows<1));

			button_disable(frmDefinition.cmdSortAdd, fSortAddDisabled);
			button_disable(frmDefinition.cmdSortEdit, fSortEditDisabled);
			button_disable(frmDefinition.cmdSortRemove, fSortRemoveDisabled);
			button_disable(frmDefinition.cmdSortRemoveAll, fSortRemoveAllDisabled);
			button_disable(frmDefinition.cmdSortMoveUp, fSortMoveUpDisabled);
			button_disable(frmDefinition.cmdSortMoveDown, fSortMoveDownDisabled);
	
			AllowUpdate = (fViewing == false);
		}
	}
	
	function ssOleDBGridSortOrder_Change() {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}
  
		var sMessage = '';
  
		with (frmDefinition.ssOleDBGridSortOrder)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol('C' + sKey))
			{
				var bBreak = getCurrentColProp('Break');
				var bPage = getCurrentColProp('Page');
				var bHidden = getCurrentColProp('Hidden');
				var bRepetition = getCurrentColProp('Repetition');
			
				if (frmUseful.txtGridActionCancelled.value == 1)
				{
				
					if (Col == 3)
					{
						if ((Columns("break").Value == "-1") && (bPage == true))
						{
							frmUseful.txtGridChangeRecursive.value = 1;
							Columns("break").Value = 0;
							sMessage = "You cannot select both 'Break on Change' and 'Page on Change' for the same column.";
							OpenHR.messageBox(sMessage,48,"Custom Reports");					
						}
					}
					if (Col == 4)
					{
						if ((bBreak == true) && (Columns("page").Value == "-1"))
						{
							frmUseful.txtGridChangeRecursive.value = 1;
							Columns("page").Value = 0;
							sMessage = "You cannot select both 'Break on Change' and 'Page on Change' for the same column.";
							OpenHR.messageBox(sMessage,48,"Custom Reports");					
						}
					}
					if (Col == 5)
					{
						if ((bHidden == true) && (Columns("value").Value == "-1"))
						{
							frmUseful.txtGridChangeRecursive.value = 1;
							Columns("value").Value = 0;
							sMessage = "You cannot select 'Value on Change' for a hidden column.";
							OpenHR.messageBox(sMessage,48,"Custom Reports");					
						}
					}
					if (Col == 6)
					{
						if ((bHidden == true) && (Columns("hide").Value == "-1"))
						{
							frmUseful.txtGridChangeRecursive.value = 1;
							Columns("hide").Value = 0;
							sMessage = "You cannot select 'Suppress Repeated Values' for a hidden column.";
							OpenHR.messageBox(sMessage,48,"Custom Reports");					
						}
						else if ((bRepetition == true) && (Columns("hide").Value == "-1"))
						{
							frmUseful.txtGridChangeRecursive.value = 1;
							Columns("hide").Value = 0;
							sMessage = "You cannot select both 'Suppress Repeated Values' and 'Repetition' for the same column.";
							OpenHR.messageBox(sMessage,48,"Custom Reports");					
						}
					}
				}	
				if (Columns("break").Value == "-1")
				{
					updateCurrentColProp('Break', true);
				}
				else
				{
					updateCurrentColProp('Break', false);
				}
				
				if (Columns("page").Value == "-1")
				{
					updateCurrentColProp('Page', true);
				}
				else
				{
					updateCurrentColProp('Page', false);
				}
				
				if (Columns("value").Value == "-1")
				{
					updateCurrentColProp('Value', true);
				}
				else
				{
					updateCurrentColProp('Value', false);
				}
				
				if (Columns("hide").Value == "-1")
				{			
					updateCurrentColProp('Hide', true);	
				}
				else
				{
					updateCurrentColProp('Hide', false);
				}
			}
		}

		if (frmUseful.txtGridActionCancelled.value == 0 &&
			frmUseful.txtGridChangeRecursive.value == 0)
		{
			frmUseful.txtChanged.value = 1;
		}

		frmDefinition.ssOleDBGridSortOrder.Refresh();
		frmDefinition.ssOleDBGridRepetition.Refresh();
	
		refreshTab4Controls();
	}
	
	function ssOleDBGridSortOrder_KeyUp(KeyCode,Shift) {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}
	
		frmUseful.txtGridActionCancelled.value = 0;
		frmUseful.txtGridChangeRecursive.value = 0;
	
		with (frmDefinition.ssOleDBGridSortOrder)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol('C' + sKey))
			{
				var bBreak = getCurrentColProp('Break');
				var bPage = getCurrentColProp('Page');
				var bHidden = getCurrentColProp('Hidden');
				var bRepetition = getCurrentColProp('Repetition');
		
				if (Col == 1) 
				{
					return;
				}
				else if (Col == 3)
				{
					if (bPage == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 4)
				{
					if (bBreak == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 5)
				{
					if (bHidden == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 6)
				{
					if ((bHidden == true) || (bRepetition == true))
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}		
			}
		}
	}
	
	function ssOleDBGridSortOrder_Click() {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}
	
		frmUseful.txtGridActionCancelled.value = 0;
		frmUseful.txtGridChangeRecursive.value = 0;
	
		with (frmDefinition.ssOleDBGridSortOrder)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol('C' + sKey))
			{
				var bBreak = getCurrentColProp('Break');
				var bPage = getCurrentColProp('Page');
				var bHidden = getCurrentColProp('Hidden');
				var bRepetition = getCurrentColProp('Repetition');
		
				if (Col == 1) 
				{
					return;
				}
				else if (Col == 3)
				{
					if (bPage == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 4)
				{
					if (bBreak == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 5)
				{
					if (bHidden == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 6)
				{
					if ((bHidden == true) || (bRepetition == true))
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}		
			}
		}
	}
	
	function ssOleDBGridSortOrder_DblClick() {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}
	
		frmUseful.txtGridActionCancelled.value = 0;
		frmUseful.txtGridChangeRecursive.value = 0;
	
		with (frmDefinition.ssOleDBGridSortOrder)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol('C' + sKey))
			{
				var bBreak = getCurrentColProp('Break');
				var bPage = getCurrentColProp('Page');
				var bHidden = getCurrentColProp('Hidden');
				var bRepetition = getCurrentColProp('Repetition');
		
				if (Col == 1) 
				{
					return;
				}
				else if (Col == 3)
				{
					if (bPage == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 4)
				{
					if (bBreak == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 5)
				{
					if (bHidden == true)
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
				else if (Col == 6)
				{
					if ((bHidden == true) || (bRepetition == true))
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}		
			}
		}
	}

	function ssOleDBGridRepetition_AfterInsert() {
	
		if(frmUseful.txtRepetitionLoaded.value == 1) 
		{
			refreshTab4Controls();
		}
	}
	
	function ssOleDBGridRepetition_Change() {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}
  
		var sMessage = '';
  
		with (frmDefinition.ssOleDBGridRepetition)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol(sKey))
			{
				var bHidden = getCurrentColProp('Hidden');
				var bSurpress = getCurrentColProp('Hide');
		
				if (frmUseful.txtGridActionCancelled.value == 1)
				{
					if (Col == 2)
					{
						if (bHidden && (Columns("repetition").Value == "-1"))
						{
							frmUseful.txtGridChangeRecursive.value = 1;
							Columns("repetition").Value = 0;
							sMessage = "You cannot select 'Repetition' for a hidden column.";
							OpenHR.messageBox(sMessage,48,"Custom Reports");					
						}
						else if (bSurpress && (Columns("repetition").Value == "-1"))
						{
							frmUseful.txtGridChangeRecursive.value = 1;
							Columns("repetition").Value = 0;
							sMessage = "You cannot select both 'Suppress Repeated Values' and 'Repetition' for the same column.";
							OpenHR.messageBox(sMessage,48,"Custom Reports");					
						}
					}
				}
		
				setGirdCol(sKey);
				if (Columns("repetition").Value == "-1") 
				{
					updateCurrentColProp('Repetition',true);
				}
				else
				{
					updateCurrentColProp('Repetition',false);
				}
			}
		}

		if (frmUseful.txtGridActionCancelled.value == 0 &&
			frmUseful.txtGridChangeRecursive.value == 0)
		{
			frmUseful.txtChanged.value = 1;
		}

		frmDefinition.ssOleDBGridRepetition.Refresh();
		frmDefinition.ssOleDBGridSortOrder.Refresh();
		
		refreshTab4Controls();
	}

	function ssOleDBGridRepetition_Click() {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}
	 
		if ((frmUseful.txtChildColumnSelected.value == 0) && (frmDefinition.ssOleDBGridRepetition.Rows > 0))
		{
			var sMessage = "Repetition cannot be selected until a child table column or calculation has been added to the report.";
			OpenHR.messageBox(sMessage,48,"Custom Reports");
			return;
		}

		frmUseful.txtGridActionCancelled.value = 0;
		frmUseful.txtGridChangeRecursive.value = 0;
	
		with (frmDefinition.ssOleDBGridRepetition)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol(sKey))
			{
				var bHidden = getCurrentColProp('Hidden');
				var bSurpress = getCurrentColProp('Hide');
		
				if (Col == 1) 
				{
					return;
				}
				else if (Col == 2)
				{
					if ((bHidden == true)  || (bSurpress == true))
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
			}
		}
	}
	
	function ssOleDBGridRepetition_DblClick() {

		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing)) {
			return;
		}

		frmUseful.txtGridActionCancelled.value = 0;
		frmUseful.txtGridChangeRecursive.value = 0;

		with (frmDefinition.ssOleDBGridRepetition)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol(sKey)) {
				var bHidden = getCurrentColProp('Hidden');
				var bSurpress = getCurrentColProp('Hide');

				if (Col == 1) {
					return;
				} else if (Col == 2) {
					if ((bHidden == true) || (bSurpress == true)) {
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
			}
		}
	}

	function ssOleDBGridRepetition_KeyUp (KeyCode, Shift) {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}

		frmUseful.txtGridActionCancelled.value = 0;
		frmUseful.txtGridChangeRecursive.value = 0;
	
		with (frmDefinition.ssOleDBGridRepetition)
		{
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol(sKey))
			{
				var bHidden = getCurrentColProp('Hidden');
				var bSurpress = getCurrentColProp('Hide');
		
				if (Col == 1) 
				{
					return;
				}
				else if (Col == 2)
				{
					if ((bHidden == true) || (bSurpress == true))
					{
						frmUseful.txtGridActionCancelled.value = 1;
						return;
					}
				}
			}
		}
	}
	
	function ssOleDBGridRepetition_RowColChange(LastRow, LastCol) {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}

		with (frmDefinition.ssOleDBGridRepetition)
		{
			for (var i=0; i<Rows; i++)
			{
				if (frmUseful.txtChildColumnSelected.value == 0) 
				{
					Columns("Column").CellStyleSet("ssetFixDataDisabled", i);
				}
				else
				{
					Columns("Column").CellStyleSet("ssetFixData", i);
				}
				
				var sKey = Columns("ColumnID").Value;
				if (setGirdCol(sKey))
				{
					var bHidden = getCurrentColProp('Hidden');
					var bSurpress = getCurrentColProp('Hide');
				
					if ((bHidden == true) || (bSurpress == true) || (frmUseful.txtChildColumnSelected.value == 0))
					{
						Columns("repetition").CellStyleSet("ssetFixDataDisabled", i);
					}
					else
					{
						Columns("repetition").CellStyleSet("ssetFixData", i);
					}
				}
			}
		
			Refresh();
		
			if (Col == 1)
			{
				Col = 0;
			}
		}
	}
	
	function ssOleDBGridRepetition_RowLoaded(Bookmark) {
		
		var fViewing;
		fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
		if ((fViewing))
		{
			return;
		}

		with (frmDefinition.ssOleDBGridRepetition)
		{
			if (frmUseful.txtChildColumnSelected.value == 0) 
			{
				Columns("Column").CellStyleSet("ssetFixDataDisabled", AddItemRowIndex(Bookmark));
			}
			else
			{
				Columns("Column").CellStyleSet("ssetFixData", AddItemRowIndex(Bookmark));
			}
			
			var sKey = Columns("ColumnID").Value;
			if (setGirdCol(sKey))
			{
				var bHidden = getCurrentColProp('Hidden');
				var bSurpress = getCurrentColProp('Hide');
    
				if ((bHidden == true) || (bSurpress == true) || (frmUseful.txtChildColumnSelected.value == 0))
				{
					Columns("repetition").CellStyleSet("ssetFixDataDisabled", AddItemRowIndex(Bookmark));
				}
				else
				{
					Columns("repetition").CellStyleSet("ssetFixData", AddItemRowIndex(Bookmark));
				}
			}
			if (Col == 1)
			{
				Col = 0;
			}
		}
	}
		
	function ssOleDBGridChildren_Change() {
		frmUseful.txtChanged.value = 1;	
	}

	function ssOleDBGridChildren_Click() {
		refreshTab2Controls();	
	}

	function ssOleDBGridChildren_AfterInsert() {
		self.focus();
		refreshTab2Controls();
		populateTableAvailable();
	}
	
	function ssOleDBGridChildren_DblClick() {
		
		if (frmUseful.txtAction.value.toUpperCase() != "VIEW")
		{
			if (frmDefinition.ssOleDBGridChildren.Rows > 0 
				&& frmDefinition.ssOleDBGridChildren.SelBookmarks.Count == 1)
			{
				childEdit(); 
			}
			else
			{
				if(cmdAddChild.disabled==false) childAdd();			  
			}
		}
	}

	function ssOleDBGridChildren_RowColChange() {
		frmDefinition.ssOleDBGridChildren.SelBookmarks.RemoveAll();
		frmDefinition.ssOleDBGridChildren.SelBookmarks.Add(frmDefinition.ssOleDBGridChildren.Bookmark);
		frmDefinition.ssOleDBGridChildren.columns('Table').cellstyleset("ssetSelected", frmDefinition.ssOleDBGridChildren.row);
		frmCustomReportChilds.childTableID.value = frmDefinition.ssOleDBGridChildren.Columns('TableID').text;
		frmCustomReportChilds.childTable.value = frmDefinition.ssOleDBGridChildren.Columns('Table').text;
		frmCustomReportChilds.childFilterID.value = frmDefinition.ssOleDBGridChildren.Columns('FilterID').text;
		frmCustomReportChilds.childFilter.value = frmDefinition.ssOleDBGridChildren.Columns('Filter').text;
		frmCustomReportChilds.childRecords.value = frmDefinition.ssOleDBGridChildren.Columns('Records').text;
		refreshTab2Controls();
	}

	function grdAccess_ComboCloseUp() {
		frmUseful.txtChanged.value = 1;
		if ((frmDefinition.grdAccess.AddItemRowIndex(frmDefinition.grdAccess.Bookmark) == 0) && (frmDefinition.grdAccess.Columns("Access").Text.length > 0)) {
			ForceAccess(frmDefinition.grdAccess, AccessCode(frmDefinition.grdAccess.Columns("Access").Text));

			frmDefinition.grdAccess.MoveFirst();
			frmDefinition.grdAccess.Col = 1;
		}
		refreshTab1Controls();
	}
	
	function grdAccess_GotFocus() {
		frmDefinition.grdAccess.Col = 1;	
	}
	
	function grdAccess_RowColChange(LastRow, LastCol) {
		
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

	function grdAccess_RowLoaded(Bookmark) {
		
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
		for (iLoop = 0; iLoop <= (pgrdAccess.Rows - 1); iLoop++) {
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

		for (iLoop = 1; iLoop <= (pgrdAccess.Rows - 1); iLoop++) {
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
		for (iLoop = 1; iLoop <= (pgrdAccess.Rows - 1); iLoop++) {
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
