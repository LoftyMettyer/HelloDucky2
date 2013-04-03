<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">

	function util_def_calendarreport_window_onload() {
		var fOK;
		fOK = true;

		var sErrMsg = frmUseful.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg, 48, "Calendar Reports");
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			setGridFont(frmDefinition.grdAccess);
			setGridFont(frmDefinition.grdEvents);
			setGridFont(frmDefinition.ssOleDBGridSortOrder);

			frmUseful.txtLoading.value = 'Y';

			// Expand the work frame and hide the option frame.
			window.parent.document.all.item("workframeset").cols = "*, 0";

			frmDefinition.cboBaseTable.style.color = 'window';
			frmDefinition.cboDescription1.style.color = 'window';
			frmDefinition.cboDescription2.style.color = 'window';
			frmDefinition.cboRegion.style.color = 'window';

			populateBaseTableCombo();

			if (frmUseful.txtAction.value.toUpperCase() == "NEW") {

				with (frmDefinition)
				{
					txtOwner.value = frmUseful.txtUserName.value;
					txtDescription.value = "";
					txtName.value = "";
					optRecordSelection1.checked = true;
					optFixedStart.checked = true;
					optFixedEnd.checked = true;
					chkShadeWeekends.checked = true;
					chkStartOnCurrentMonth.checked = true;
					chkCaptions.checked = true;
					optOutputFormat0.checked = true;
					chkDestination0.checked = true;
				}

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
						window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("This definition is set to output to printer " + frmOriginalDefinition.txtDefn_OutputPrinterName.value + " which is not set up on your PC.");
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
</script>

<script LANGUAGE=JavaScript id=scptGeneralFunctions>
<!--
	var sRepDefn;
	var sSortDefn;
	var fValidating;
	fValidating = false;

	function displayPage(piPageNumber) 
	{
	
		var iLoop;

		window.parent.frames("refreshframe").document.forms("frmRefresh").submit();
			
		if (piPageNumber == 1) 
		{
			var lngGridHeight = new Number(0);
			var lngGridWidth = new Number(0);
		
			div1.style.visibility="visible";
			div1.style.display="block";
			div2.style.visibility="hidden";
			div2.style.display="none";
			div3.style.visibility="hidden";
			div3.style.display="none";
			div4.style.visibility="hidden";
			div4.style.display="block";	
			lngGridHeight = frmDefinition.ssOleDBGridSortOrder.style.height;
			lngGridWidth = frmDefinition.ssOleDBGridSortOrder.style.width;
			frmDefinition.ssOleDBGridSortOrder.style.height = 0;
			frmDefinition.ssOleDBGridSortOrder.style.width = 0;
			loadSortDefinition();
			frmDefinition.ssOleDBGridSortOrder.style.height = lngGridHeight;
			frmDefinition.ssOleDBGridSortOrder.style.width = lngGridWidth;
			div4.style.visibility="hidden";
			div4.style.display="none";
			div5.style.visibility="hidden";
			div5.style.display="none";
		
			button_disable(frmDefinition.btnTab1, true);
			button_disable(frmDefinition.btnTab2, false);
			button_disable(frmDefinition.btnTab3, false);
			button_disable(frmDefinition.btnTab4, false);
			button_disable(frmDefinition.btnTab5, false);
		
			refreshTab1Controls();

			if (frmDefinition.txtName.disabled == false) 
			{
				try 
				{
					frmDefinition.txtName.focus();
				}
				catch (e) {}
			}

		}

		if (frmUseful.txtAvailableColumnsLoaded.value != 1)
		{
			return;
		}
	
		if (piPageNumber == 2) 
		{
			div1.style.visibility="hidden";
			div1.style.display="none";
			div2.style.visibility="visible";
			div2.style.display="block";
			div3.style.visibility="hidden";
			div3.style.display="none";
			div4.style.visibility="hidden";
			div4.style.display="none";
			div5.style.visibility="hidden";
			div5.style.display="none";
		
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

			if (frmDefinition.grdEvents.Enabled)
			{
				frmDefinition.grdEvents.focus();
			}
			
			refreshTab2Controls();
		}

		if (piPageNumber == 3) 
		{
			div1.style.visibility="hidden";
			div1.style.display="none";
			div2.style.visibility="hidden";
			div2.style.display="none";
			div3.style.visibility="visible";
			div3.style.display="block";
			div4.style.visibility="hidden";
			div4.style.display="none";
			div5.style.visibility="hidden";
			div5.style.display="none";
		
			button_disable(frmDefinition.btnTab1, false);
			button_disable(frmDefinition.btnTab2, false);
			button_disable(frmDefinition.btnTab3, true);
			button_disable(frmDefinition.btnTab4, false);
			button_disable(frmDefinition.btnTab5, false);

			refreshTab3Controls();
		}

		if (piPageNumber == 4) 
		{
			div1.style.visibility="hidden";
			div1.style.display="none";
			div2.style.visibility="hidden";
			div2.style.display="none";
			div3.style.visibility="hidden";
			div3.style.display="none";
			div4.style.visibility="visible";
			div4.style.display="block";
			div5.style.visibility="hidden";
			div5.style.display="none";
		
			button_disable(frmDefinition.btnTab1, false);
			button_disable(frmDefinition.btnTab2, false);
			button_disable(frmDefinition.btnTab3, false);
			button_disable(frmDefinition.btnTab4, true);
			button_disable(frmDefinition.btnTab5, false);

			loadSortDefinition();
		
			frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();
			if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) 
			{
				frmDefinition.ssOleDBGridSortOrder.MoveFirst();
				frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
			}

			if (frmDefinition.ssOleDBGridSortOrder.Enabled)
			{
				frmDefinition.ssOleDBGridSortOrder.focus();
			}
					
			refreshTab4Controls();
		}

		if (piPageNumber == 5) 
		{
			div1.style.visibility="hidden";
			div1.style.display="none";
			div2.style.visibility="hidden";
			div2.style.display="none";
			div3.style.visibility="hidden";
			div3.style.display="none";
			div4.style.visibility="hidden";
			div4.style.display="none";
			div5.style.visibility="visible";
			div5.style.display="block";

			button_disable(frmDefinition.btnTab1, false);
			button_disable(frmDefinition.btnTab2, false);
			button_disable(frmDefinition.btnTab3, false);
			button_disable(frmDefinition.btnTab4, false);
			button_disable(frmDefinition.btnTab5, true);

			refreshTab5Controls();
		}
	}

	function populateBaseTableCombo()
	{
		var i;
	
		//Clear the existing data in the base table combo
		while (frmDefinition.cboBaseTable.options.length > 0) 
		{
			frmDefinition.cboBaseTable.options.remove(0);
		}

		var dataCollection = frmTables.elements;
		if (dataCollection!=null) 
		{
			for (i=0; i<dataCollection.length; i++)  
			{
				sControlName = dataCollection.item(i).name;
				sControlTag = sControlName.substr(0, 13);
				if (sControlTag == "txtTableName_") 
				{
					sTableID = sControlName.substr(13);
					var oOption = document.createElement("OPTION");
					frmDefinition.cboBaseTable.options.add(oOption);
					oOption.innerText = dataCollection.item(i).value;
					oOption.value = sTableID;			
				}
			}
		}	
	}

	function populateBaseTableColumns()
	{
		// Get the columns/calcs for the current table selection.
		var frmGetDataForm = window.parent.frames("dataframe").document.forms("frmGetData");

		frmUseful.txtAvailableColumnsLoaded.value = 0;

		frmGetDataForm.txtAction.value = "LOADCALENDARREPORTCOLUMNS";
		frmGetDataForm.txtReportBaseTableID.value = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;
			
		window.parent.frames("dataframe").refreshData();
	}
	
	function trim(strInput)
	{
		if (strInput.length < 1)
		{
			return "";
		}
		
		while (strInput.substr(strInput.length-1, 1) == " ") 
		{
			strInput = strInput.substr(0, strInput.length - 1);
		}
	
		while (strInput.substr(0, 1) == " ") 
		{
			strInput = strInput.substr(1, strInput.length);
		}
	
		return strInput;
	}
                        
	function setBaseTable(piTableID) 
	{
		var i;
	
		if (piTableID == 0) piTableID = frmUseful.txtPersonnelTableID.value;

		if (piTableID > 0) 
		{
			for (i=0; i<frmDefinition.cboBaseTable.options.length; i++)  
			{
				if (frmDefinition.cboBaseTable.options(i).value == piTableID) 
				{
					frmDefinition.cboBaseTable.selectedIndex = i;
					frmUseful.txtCurrentBaseTableID.value = piTableID;
					break;
				}			
			}
		}
		else 
		{
			if (frmDefinition.cboBaseTable.options.length > 0) 
			{
				frmDefinition.cboBaseTable.selectedIndex = 0;
				frmUseful.txtCurrentBaseTableID.value = frmDefinition.cboBaseTable.options(0).value;
			}
		}
	}

	function changeBaseTable() 
	{
		var i;
		var iPollCounter;
		var iPollPeriod;
		var frmRefresh;
		var iDummy;
	
		iPollPeriod = 100;
		iPollCounter = iPollPeriod;
		frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");	
	
		if (frmUseful.txtLoading.value == 'N') 
		{
	
			if ((frmDefinition.grdEvents.Rows > 0) ||
				((frmUseful.txtAction.value.toUpperCase() != "NEW") && 
					(frmUseful.txtEventsLoaded.value == 0))) 
			{

				iAnswer = window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Warning: Changing the base table will result in all table/column specific aspects of this report definition being cleared. Are you sure you wish to continue?",36,"Calendar Reports");
				if (iAnswer == 7)	
				{
					// cancel and change back ! (txtcurrentbasetable)
					setBaseTable(frmUseful.txtCurrentBaseTableID.value);
					return;
				}
				else	
				{
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
					
		while (frmDefinition.grdEvents.Rows > 0) 
		{
			frmDefinition.grdEvents.RemoveAll();
		}

		while (frmDefinition.ssOleDBGridSortOrder.Rows > 0) 
		{
			frmDefinition.ssOleDBGridSortOrder.RemoveAll();
		}
	
		if (frmUseful.txtLoading.value == 'N')
		{ 
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
		if (dataCollection!=null) 
		{
			sReqdControlName = new String("txtTableRelations_");
			sReqdControlName = sReqdControlName.concat(frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);
				
			for (i=0; i<dataCollection.length; i++)  
			{
				sControlName = dataCollection.item(i).name;		
				if (sControlName == sReqdControlName) 
				{
					sRelationNames = dataCollection.item(i).value;
					frmEventDetails.relationNames.value = sRelationNames;
					break;
				}
			}
		}

		recalcHiddenEventFiltersCount();
	
		if (frmUseful.txtLoading.value != 'N')
		{ 
			refreshTab1Controls();
		}
	
		frmUseful.txtCurrentBaseTableID.value = frmDefinition.cboBaseTable.options(frmDefinition.cboBaseTable.options.selectedIndex).value;
		frmUseful.txtAvailableColumnsLoaded.value = 0;
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
	
		if (frmUseful.txtAvailableColumnsLoaded.value == 0)
		{
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
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("This definition will now be made hidden as it contains a hidden picklist/filter/calculation.", 64);
				}
				ForceAccess(frmDefinition.grdAccess, "HD");
				frmUseful.txtChanged.value = 1;
			}
			else
			{
				if (frmSelectionAccess.forcedHidden.value == "N") {
					//MH20040816 Fault 9047
					//if (fSilent == false) {
					if ((fSilent == false) && (frmUseful.txtLoading.value != "Y")) {
						window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("The definition access cannot be changed as it contains a hidden picklist/filter/calculation.", 64);
					}
				}
			}
			frmSelectionAccess.forcedHidden.value = "Y";

			frmDefinition.grdAccess.Columns("Access").Style = 0; // 0 = Edit
		}
		else {
			try
			{
				window.resizeBy(0,-1);
				window.resizeBy(0,1);
				window.resizeBy(0,-1);
				window.resizeBy(0,1);
			}
			catch(e) {}
			if (frmSelectionAccess.forcedHidden.value == "Y") {
				frmSelectionAccess.forcedHidden.value = "N";
				// No longer forced hidden.
				if (fSilent == false) {
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("This definition no longer has to be hidden.", 64);
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
		if ((iDescCount < 2) || (frmDefinition.cboDescriptionSeparator.selectedIndex < 0)) 
		{
			frmDefinition.cboDescriptionSeparator.selectedIndex = 0;
		}

		combo_disable(frmDefinition.cboDescriptionSeparator, ((fViewing == true) || (iDescCount<2)));
		
		if (frmDefinition.cboRegion.selectedIndex < 0) frmDefinition.cboRegion.selectedIndex = 0;

		button_disable(frmDefinition.cmdBasePicklist, ((frmDefinition.optRecordSelection2.checked == false)
											 || (fViewing == true)));
		button_disable(frmDefinition.cmdBaseFilter, ((frmDefinition.optRecordSelection3.checked == false)
											 || (fViewing == true)));

		if (frmDefinition.optRecordSelection2.checked || frmDefinition.optRecordSelection3.checked)
		{
			checkbox_disable(frmDefinition.chkPrintFilterHeader, (fViewing == true));
		} 
		else
		{
			frmDefinition.chkPrintFilterHeader.checked = false;
			checkbox_disable(frmDefinition.chkPrintFilterHeader, true);
		}

		with (frmDefinition)
		{
			if (chkIncludeBHols.checked || chkIncludeWorkingDaysOnly.checked || chkShadeBHols.checked
					|| (cboRegion.options[cboRegion.selectedIndex].value > 0))
			{
				checkbox_disable(chkGroupByDesc, true);
			}
			else
			{
				checkbox_disable(chkGroupByDesc, fViewing);
			}
				
			if (chkGroupByDesc.checked)
			{
				checkbox_disable(chkIncludeBHols, true);
				checkbox_disable(chkIncludeWorkingDaysOnly, true);
				checkbox_disable(chkShadeBHols, true);
				combo_disable(cboRegion, true);
			}
			else
			{
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
		
		with (frmDefinition)
		{
			button_disable(cmdAddEvent, (fViewing == true));
			button_disable(cmdEditEvent, ((grdEvents.Rows<1) 
								|| (grdEvents.SelBookmarks.Count!=1) 
								|| (fViewing == true)));
			button_disable(cmdRemoveEvent, ((grdEvents.Rows<1) 
									|| (grdEvents.SelBookmarks.Count!=1) 
									|| (fViewing == true)));
			button_disable(cmdRemoveAllEvents, ((grdEvents.Rows<1) 
										|| (fViewing == true)));
		}
	
		recalcHiddenEventFiltersCount();

		refreshTab1Controls();

		frmDefinition.grdEvents.RowHeight = 19;
	
		// Little dodge to get around a browser bug that
		// does not refresh the display on all controls.
		try
		{
			window.resizeBy(0,-1);
			window.resizeBy(0,1);
		}
		catch(e) {}
	}

	function refreshTab3Controls()
	{
		var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
	
		with (frmDefinition)
		{
			if (optFixedStart.checked) 
			{
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
			else if (optCurrentStart.checked)
			{
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
			else if (optOffsetStart.checked)
			{
				text_disable(txtFixedStart, true);
				txtFixedStart.value = '';
			
				if (Number(txtFreqStart.value) > 0)
				{
					optFixedEnd.checked = false;
					radio_disable(optFixedEnd, true);
				
					text_disable(txtFixedEnd, true);
					txtFixedEnd.value = '';
				
					optCurrentEnd.checked = false;
					radio_disable(optCurrentEnd, true);
					//optOffsetEnd.checked = true;
				} 
				else
				{
					radio_disable(optFixedEnd, fViewing);
					text_disable(txtFixedEnd, ((fViewing) || (optFixedEnd.checked == false)));
					radio_disable(optCurrentEnd, fViewing);
				}
			
				text_disable(txtFreqStart, fViewing);
				if (txtFreqStart.value == '') 
				{
					txtFreqStart.value = 0;
				}
				button_disable(cmdPeriodStartDown, fViewing);
				button_disable(cmdPeriodStartUp, fViewing);
				combo_disable(cboPeriodStart, fViewing);
				if (cboPeriodStart.selectedIndex < 0)
				{
					cboPeriodStart.selectedIndex = 0;
				}
			
				radio_disable(optOffsetEnd, fViewing);
				text_disable(txtFreqEnd, ((fViewing) || (optOffsetEnd.checked == false)));
				if (txtFreqEnd.value == '') 
				{
					txtFreqEnd.value = 0;
				}
				button_disable(cmdPeriodEndDown, ((fViewing) || (optOffsetEnd.checked == false)));
				button_disable(cmdPeriodEndUp, ((fViewing) || (optOffsetEnd.checked == false)));
			
				button_disable(cmdCustomStart, true);
				txtCustomStart.value = '';
				txtCustomStartID.value = 0;
			}
			else if (optCustomStart.checked)
			{
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
		
			if (optFixedEnd.checked)
			{
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
			else if (optCurrentEnd.checked)
			{
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
			else if (optOffsetEnd.checked)
			{
				text_disable(txtFixedEnd, true);
				txtFixedEnd.value = '';
			
				text_disable(txtFreqEnd, fViewing);
				if (txtFreqEnd.value == '') 
				{
					txtFreqEnd.value = 0;
				}
				button_disable(cmdPeriodEndDown, fViewing);
				button_disable(cmdPeriodEndUp, fViewing);
				combo_disable(cboPeriodEnd, fViewing);
				if (cboPeriodEnd.selectedIndex < 0)
				{
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
			else if (optCustomEnd.checked)
			{
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
		
			if (txtFreqStart.disabled)
			{
				button_disable(cmdPeriodStartDown, true);
				button_disable(cmdPeriodStartUp, true);
			}
			else
			{
				button_disable(cmdPeriodStartDown, fViewing);
				button_disable(cmdPeriodStartUp, fViewing);
			}

			if (txtFreqEnd.disabled)
			{
				button_disable(cmdPeriodEndDown, true);
				button_disable(cmdPeriodEndUp, true);
			}
			else
			{
				button_disable(cmdPeriodEndDown, fViewing);
				button_disable(cmdPeriodEndUp, fViewing);
			}
		
			var blnPersonnelBaseTable = (frmUseful.txtPersonnelTableID.value == frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value)
			var blnRegionSelected = (cboRegion.options[cboRegion.selectedIndex].value > 0);
		
			if (chkIncludeBHols.checked || chkIncludeWorkingDaysOnly.checked || chkShadeBHols.checked
					|| blnRegionSelected)
			{
				checkbox_disable(chkGroupByDesc, true);
			}
			else
			{
				checkbox_disable(chkGroupByDesc, fViewing);
			}
			
			if (chkGroupByDesc.checked)
			{
				checkbox_disable(chkIncludeBHols, true);
				checkbox_disable(chkIncludeWorkingDaysOnly, true);
				checkbox_disable(chkShadeBHols, true);
				combo_disable(cboRegion, true);
			}
			else
			{
				checkbox_disable(chkIncludeBHols, ((fViewing) || ((!blnPersonnelBaseTable) && (!blnRegionSelected))));
				checkbox_disable(chkIncludeWorkingDaysOnly, ((fViewing) || (!blnPersonnelBaseTable)));
				checkbox_disable(chkShadeBHols, ((fViewing) || ((!blnPersonnelBaseTable) && (!blnRegionSelected))));
				combo_disable(cboRegion, fViewing);
			}
			
			checkbox_disable(chkCaptions, fViewing);
			checkbox_disable(chkShadeWeekends, fViewing);
			checkbox_disable(chkStartOnCurrentMonth, fViewing);
		
			if (!blnPersonnelBaseTable)
			{
				chkIncludeWorkingDaysOnly.checked = false;
				if (!blnRegionSelected)
				{
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

	function refreshTab4Controls()
	{
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
			&& frmDefinition.ssOleDBGridSortOrder.Rows > 0)
		{
			// Are we on the top row ?
			if ((frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) == 0) 
				|| (frmDefinition.ssOleDBGridSortOrder.rows <= 1))
			{
				fSortMoveUpDisabled = true; 
			}

			// Are we on the bottom row ?
			if ((frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) == frmDefinition.ssOleDBGridSortOrder.rows - 1) 
				|| (frmDefinition.ssOleDBGridSortOrder.rows <= 1))
			{
				fSortMoveDownDisabled = true; 
			}
		
		}	

		if (frmDefinition.ssOleDBGridSortOrder.Rows < 1 
			|| frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count != 1)
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

		fSortRemoveDisabled = ((frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count!=1) || (fViewing == true));
		fSortRemoveAllDisabled = ((fViewing == true) || (frmDefinition.ssOleDBGridSortOrder.Rows < 1));
		fSortEditDisabled = ((frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count!=1) || (fViewing == true));

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

		// Little dodge to get around a browser bug that
		// does not refresh the display on all controls.
		try
		{
			window.resizeBy(0,-1);
			window.resizeBy(0,1);
		}
		catch(e) {}
	}





	function refreshTab5Controls()
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
				chkDestination2.checked = false
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
				checkbox_disable(chkDestination2, (fViewing == true));
				if (chkDestination2.checked ==true)
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
			else
			{
				optOutputFormat0.checked = true;
				chkDestination0.checked = true;
				refreshTab5Controls();
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
		refreshTab5Controls();
	}

	function saveFile()
	{
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
			sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			dialog.InitDir = sPath;
		}
		else {
			dialog.FileName = frmDefinition.txtFilename.value;
		}


		try {
			dialog.ShowSave();

			if (dialog.FileName.length > 256) {
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Path and file name must not exceed 256 characters in length");
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
	
		frmSelectionAccess.baseHidden.value = "N";
	}

	function setRecordsNumeric(objTextBox)
	{
		var sConvertedValue;
		var sDecimalSeparator;
		var sThousandSeparator;
		var sPoint;
			
		sDecimalSeparator = "\\";
		sDecimalSeparator = sDecimalSeparator.concat(window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDecimalSeparator);
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(window.parent.frames("menuframe").ASRIntranetFunctions.LocaleThousandSeparator);
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");
			
		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");
		
		if (objTextBox.value == '') 
		{
			objTextBox.value = 0;
		}
			
		// Convert the value from locale to UK settings for use with the isNaN funtion.
		sConvertedValue = new String(objTextBox.value);
		
		// Remove any thousand separators.
		sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
		objTextBox.value = sConvertedValue;

		// Convert any decimal separators to '.'.
		if (window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDecimalSeparator != ".") 
		{
			// Remove decimal points.
			sConvertedValue = sConvertedValue.replace(rePoint, "A");
			// replace the locale decimal marker with the decimal point.
			sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
		}
		
		if(isNaN(sConvertedValue) == true) 
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Invalid numeric value.",48,"Calendar Reports");
			objTextBox.value = 0;
		}
		else 
		{
			if (sConvertedValue.indexOf(".") >= 0 ) 
			{
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Invalid integer value.",48,"Calendar Reports");
				objTextBox.value = 0;
			}
			else 
			{
				if (objTextBox.value > 99) 
				{
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("The value cannot be greater than 99.",48,"Calendar Reports");
					objTextBox.value = 99;
				}
			}
		}
	}

	function replace(sExpression, sFind, sReplace)
	{
		//gi (global search, ignore case)
		var re = new RegExp(sFind,"gi");
		sExpression = sExpression.replace(re, sReplace);
		return(sExpression);
	}
	
	function spinRecords(pfUp, objTextBox) 
	{ 
		var iRecords = objTextBox.value; 
		if (pfUp == true) 
		{ 
			iRecords = ++iRecords; 
		} 
		else 
		{ 
			iRecords = iRecords - 1; 
		} 
		objTextBox.value = iRecords; 
	}

	function validateOffsets()
	{
		with (frmDefinition)
		{
			if ((optOffsetStart.checked == true) && (optOffsetEnd.checked == true))
			{
				if (cboPeriodEnd.selectedIndex != cboPeriodStart.selectedIndex)
				{
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("The End Date Offset period must be the same as the Start Date Offset period",48,"Calendar Reports");
					cboPeriodEnd.selectedIndex = cboPeriodStart.selectedIndex;
				}
				
				if (Number(txtFreqStart.value) > Number(txtFreqEnd.value))
				{
					txtFreqEnd.value = txtFreqStart.value;
				}
			}
		} 
	}

	function selectCalc(psCalcType, bRecordIndepend)
	{	
		var iTableID;
		var iCurrentID;
		var sURL;
	
		if (psCalcType == 'baseDesc')
		{
			iTableID = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;
			iCurrentID = frmDefinition.txtDescExprID.value;
		}
		else if (psCalcType == 'startDate')
		{
			iTableID = 0;
			iCurrentID = frmDefinition.txtCustomStartID.value;
		}
		else if (psCalcType == 'endDate')
		{
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
	
		if (strDefOwner == strCurrentUser) 
		{
			frmCalcSelection.recSelDefOwner.value = '1';
		}
		else
		{
			frmCalcSelection.recSelDefOwner.value = '0';
		}
			
		sURL = "dialog" +
			"?calcSelRecInd=" + frmCalcSelection.calcSelRecInd.value +
			"&calcSelType=" + escape(frmCalcSelection.calcSelType.value) +
			"&calcSelTableID=" + escape(frmCalcSelection.calcSelTableID.value) +
			"&calcSelCurrentID=" + escape(frmCalcSelection.calcSelCurrentID.value) +
			"&recSelDefOwner=" + escape(frmCalcSelection.recSelDefOwner.value) +
			"&destination=util_calcSelection";
		openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");
	
		frmUseful.txtChanged.value = 1;
		refreshTab1Controls();
	}		
	
	function selectEmailGroup()
	{
		var sURL;
	
		frmEmailSelection.EmailSelCurrentID.value = frmDefinition.txtEmailGroupID.value; 

		sURL = "util_emailSelection" +
			"?EmailSelCurrentID=" + frmEmailSelection.EmailSelCurrentID.value;
		openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");
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
	
		if (strDefOwner == strCurrentUser) 
		{
			frmRecordSelection.recSelDefOwner.value = '1';
		}
		else
		{
			frmRecordSelection.recSelDefOwner.value = '0';
		}
		
		sURL = "util_recordSelection" +
			"?recSelType=" + escape(frmRecordSelection.recSelType.value) +
			"&recSelTableID=" + escape(frmRecordSelection.recSelTableID.value) + 
			"&recSelCurrentID=" + escape(frmRecordSelection.recSelCurrentID.value) +
			"&recSelTable=" + escape(frmRecordSelection.recSelTable.value) +
			"&recSelDefOwner=" + escape(frmRecordSelection.recSelDefOwner.value);
		openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");

		frmUseful.txtChanged.value = 1;
		refreshTab1Controls();
	}

	function eventFilterString()
	{
		var i;
		var pvarbookmark;
		var sEventFilterString = '';
	
		with (frmDefinition.grdEvents)
		{
			if (Rows > 0)
			{
				MoveFirst();
				for (var i=0; i < Rows; i++)
				{
					pvarbookmark = GetBookmark(i);			
					sEventFilterString = sEventFilterString + Columns('FilterID').CellValue(pvarbookmark);
					if (i != Rows-1)
					{
						sEventFilterString = sEventFilterString + "	";
					}
				}
			}
		}
		if (sEventFilterString.length < 1)
		{
			sEventFilterString = '';
		}
		return sEventFilterString;
	}
	
	function submitDefinition()
	{
		var i;
		var iIndex;
		var sColumnID;
		var sType;
		var iPollCounter;
		var iPollPeriod;
		var frmRefresh;
		var iDummy;
		var sURL;
	
		iPollPeriod = 100;
		iPollCounter = iPollPeriod;
		frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");	

		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}

		if (validateTab1() == false) {window.parent.frames("menuframe").refreshMenu(); return;}
		if (validateTab2() == false) {window.parent.frames("menuframe").refreshMenu(); return;}
		if (validateTab3() == false) {window.parent.frames("menuframe").refreshMenu(); return;}
		if (validateTab4() == false) {window.parent.frames("menuframe").refreshMenu(); return;}
		if (validateTab5() == false) {window.parent.frames("menuframe").refreshMenu(); return;}
		if (populateSendForm() == false) {window.parent.frames("menuframe").refreshMenu(); return;}

		// Now create the validate popup to check that any filters/calcs
		// etc havent been deleted, or made hidden etc.		

		// first populate the validate fields
		frmValidate.validateBaseFilter.value = frmDefinition.txtBaseFilterID.value;
		frmValidate.validateBasePicklist.value = frmDefinition.txtBasePicklistID.value;
		frmValidate.validateEmailGroup.value = frmDefinition.txtEmailGroupID.value;
		frmValidate.validateEventFilter.value = eventFilterString();		
		frmValidate.validateName.value = frmDefinition.txtName.value;
		frmValidate.validateDescExpr.value = frmDefinition.txtDescExprID.value 
		frmValidate.validateCustomStart.value = frmDefinition.txtCustomStartID.value 
		frmValidate.validateCustomEnd.value = frmDefinition.txtCustomEndID.value 
	
		if(frmUseful.txtAction.value.toUpperCase() == "EDIT")
		{
			frmValidate.validateTimestamp.value = frmOriginalDefinition.txtDefn_Timestamp.value;
			frmValidate.validateUtilID.value = frmUseful.txtUtilID.value;
		}
		else 
		{
			frmValidate.validateTimestamp.value = 0;
			frmValidate.validateUtilID.value = 0;
		}
	
		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}

		sHiddenGroups = HiddenGroups(frmDefinition.grdAccess);
		frmValidate.validateHiddenGroups.value = sHiddenGroups;

		sURL = "dialog" +
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
		openDialog(sURL, (screen.width)/2,(screen.height)/3, "no", "no");
	}

	function cancelClick()
	{
		if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
			(definitionChanged() == false)) {
			window.location.href="defsel";
			return;
		}

		answer = window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You have changed the current definition. Save changes ?",3,"Calendar Reports");
		if (answer == 7) {
			// No
			window.location.href="defsel";
			return (false);
		}
		if (answer == 6) {
			// Yes
			okClick();
		}
	}

	function okClick()
	{
		window.parent.frames("menuframe").disableMenu();
	
		frmSend.txtSend_reaction.value = "CALENDARREPORTS";
		submitDefinition();
	}

	function saveChanges(psAction, pfPrompt, pfTBOverride)
	{
		if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
			(!definitionChanged())) 
		{
			return 7; //No to saving the changes, as none have been made.
		}

		answer = window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You have changed the current definition. Save changes ?",3,"Calendar Reports");
		if (answer == 7) 
		{
			// No
			return 7;
		}
		if (answer == 6) 
		{
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
	
		if (frmUseful.txtAction.value.toUpperCase() == "COPY") 
		{
			return true;
		}
	
		if (frmUseful.txtChanged.value == 1) 
		{
			return true;
		}
		else 
		{
			if (frmUseful.txtAction.value.toUpperCase() != "NEW") {
				// Compare the tab 1 controls with the original values.
				if (frmDefinition.txtName.value != frmOriginalDefinition.txtDefn_Name.value) 
				{
					return true;
				}
		
				if (frmDefinition.txtDescription.value != frmOriginalDefinition.txtDefn_Description.value) 
				{
					return true;
				}
					
				if (frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value != frmOriginalDefinition.txtDefn_BaseTableID.value) 
				{
					return true;
				}

				if (frmOriginalDefinition.txtDefn_PicklistID.value > 0) 
				{
					if (frmDefinition.optRecordSelection2.checked == false) 
					{
						return true;
					}
					else 
					{
						if (frmDefinition.txtBasePicklistID.value != frmOriginalDefinition.txtDefn_PicklistID.value) 
						{
							return true;
						}
					}				
				}
				else 
				{
					if (frmOriginalDefinition.txtDefn_FilterID.value > 0) 
					{
						if (frmDefinition.optRecordSelection3.checked == false) 
						{
							return true;
						}
						else 
						{
							if (frmDefinition.txtBaseFilterID.value != frmOriginalDefinition.txtDefn_FilterID.value) 
							{
								return true;
							}
						}				
					}
					else 
					{
						if (frmDefinition.optRecordSelection1.checked == false) 
						{
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
		if (dataCollection!=null) 
		{
			for (i=0; i<dataCollection.length; i++)  
			{
				sControlName = dataCollection.item(i).name;
					
				if (sControlName == sReqdControlName) 
				{
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

	function getEventKey()
	{
		var i = 1;
	
		while (checkUniqueEventKey("EV_" + i))
		{
			i = i + 1;
		}
	
		var sEventKey = new String("EV_" + i);
	
		return sEventKey;
	}
	
	function checkUniqueEventKey(psNewKey)
	{
		var bm;
	
		with (frmDefinition.grdEvents)
		{
			Redraw = false;
			MoveFirst();
			for (var i=0; i<Rows; i++)
			{
				bm = AddItemBookmark(i);
				if (psNewKey == trim(Columns("EventKey").CellValue(bm)))
				{
					Redraw = true;
					return true;
				}
			}
			Redraw = true;
		}
		return false;
	}
	
	function eventAdd()
	{
		var sURL;
	
		with (frmEventDetails)
		{
			eventAction.value  = "NEW";
			eventID.value = getEventKey();
			eventFilterHidden.value = "";
		
			if (frmDefinition.grdEvents.Rows < 999)
			{
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
				openDialog(sURL, 650,500, "yes", "yes");

				frmUseful.txtChanged.value = 1;
			}
			else 
			{
				var sMessage = "";
				sMessage = "The maximum of 999 events has been selected.";
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMessage,64,"Calendar Reports");		
			}
		}

		refreshTab2Controls();
	}
	
	function eventEdit()
	{
		var sURL;
	
		with (frmEventDetails)
		{
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
			eventKeyCharacter.value = frmDefinition.grdEvents.Columns("Legend").value.substr(0,2);
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
	
	
	function eventRemove()
	{
		var lRow;
		var lngSelectedEvent;
	
		with (frmDefinition.grdEvents)
		{
			if (Rows < 1) return;
	
			lRow = AddItemRowIndex(Bookmark);
			lngSelectedEvent = Columns('EventKey').CellValue(lRow);

			var bContinueRemoval;

			bContinueRemoval = true;
		
			if (!bContinueRemoval) return;
		
			if (Rows == 1)
			{
				RemoveAll();
			}
			else
			{
				RemoveItem(lRow);
				if (Rows != 0)
				{
					if (lRow < Rows)
					{
						Bookmark = lRow;
					}
					else
					{
						Bookmark = (Rows - 1);
					}
					SelBookmarks.Add(Bookmark);
				}
			}
		}
		frmUseful.txtChanged.value = 1;		
	
		refreshTab2Controls();
	}

	function eventRemoveAll()
	{
		var i;
		var pvarbookmark;
		var bContinueRemoval;
		var lngSelectedEvent;
		var lngRowCount;
	
		bContinueRemoval = true;
	
		if (!bContinueRemoval) return;
	
		with (frmDefinition.grdEvents)
		{
			Redraw = false;
			lngRowCount = Rows;
			for (i=0; i<lngRowCount; i++)
			{
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

	
	function sortAdd()
	{
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
	
		if (frmDefinition.cboDescription1.options.length > 0)
		{
			for (i=0; i < frmDefinition.cboDescription1.options.length; i++) 
			{
				iColumnsCount++;
				if (frmSortOrder.txtSortInclude.value != '') 
				{
					frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
				}
				frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.cboDescription1.options[i].value;
			}		 
		}
		else
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("No columns on the base table.",48,"Calendar Reports");
		}
		
		if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) 
		{
			frmDefinition.ssOleDBGridSortOrder.Redraw = false;
			frmDefinition.ssOleDBGridSortOrder.movefirst();

			for (i=0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) 
			{
				bm = frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(i);
				if (frmSortOrder.txtSortExclude.value != '') 
				{
					frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + ',';
				}
			
				frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + frmDefinition.ssOleDBGridSortOrder.Columns(0).CellValue(bm);

				//frmDefinition.ssOleDBGridSortOrder.movenext();
			}		 

			frmDefinition.ssOleDBGridSortOrder.Redraw = true;
		}
	
		if (frmSortOrder.txtSortInclude.value == frmSortOrder.txtSortExclude.value) 
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You have selected all base columns in the sort order.",48,"Calendar Reports");
		}
		else if ((frmDefinition.ssOleDBGridSortOrder.Rows - iColumnsCount) == 0)
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You have selected all base columns in the sort order.",48,"Calendar Reports");
		}
		else
		{
			if (frmSortOrder.txtSortInclude.value != '') 
			{
				sURL = "util_sortorderselection" +
					"?txtSortInclude=" + escape(frmSortOrder.txtSortInclude.value) +
					"&txtSortExclude=" + escape(frmSortOrder.txtSortExclude.value) + 
					"&txtSortEditing=" + escape(frmSortOrder.txtSortEditing.value) +
					"&txtSortColumnID=" + escape(frmSortOrder.txtSortColumnID.value) +
					"&txtSortColumnName=" + escape(frmSortOrder.txtSortColumnName.value) +
					"&txtSortOrder=" + escape(frmSortOrder.txtSortOrder.value);
				openDialog(sURL, 600,275, "yes", "yes");

				frmUseful.txtChanged.value = 1;
			}
		}	
	}

	function sortEdit()
	{
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

		for (i=0; i < frmDefinition.cboDescription1.options.length; i++) 
		{
			if (frmSortOrder.txtSortInclude.value != '') 
			{
				frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
			}
			frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.cboDescription1.options[i].value;
		}		 

		frmDefinition.ssOleDBGridSortOrder.Redraw = false;	
		var rowNum = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark);
		frmDefinition.ssOleDBGridSortOrder.MoveFirst();

		for (i=0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) 
		{
			if (frmDefinition.ssOleDBGridSortOrder.columns(0).text != frmSortOrder.txtSortColumnID.value) 
			{
				if (frmSortOrder.txtSortExclude.value != '') 
				{
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
		openDialog(sURL, 500,275, "yes", "yes");

		frmUseful.txtChanged.value = 1;
		refreshTab4Controls();	
	}

	function sortRemove()
	{
		if (frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Count() == 0) 
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a column to remove.",48,"Calendar Reports");
			return;
		}

		frmDefinition.ssOleDBGridSortOrder.RemoveItem(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.bookmark));

		if (frmDefinition.ssOleDBGridSortOrder.Rows !=0) 
		{
			frmDefinition.ssOleDBGridSortOrder.MoveLast();
			frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
		}
	
		frmUseful.txtChanged.value = 1;
		refreshTab4Controls();	
	}

	function sortRemoveAll()
	{
		frmDefinition.ssOleDBGridSortOrder.RemoveAll();
		frmUseful.txtChanged.value = 1;
		refreshTab4Controls();	
	}

	function sortMove(pfUp)
	{
		var sAddline = '';
	
		if (pfUp == true) 
		{
			iNewIndex = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) - 1;
			iOldIndex = iNewIndex + 2;
			iSelectIndex =iNewIndex;
		}
		else 
		{
			iNewIndex = frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark) + 2;
			iOldIndex = iNewIndex - 2;
			iSelectIndex =iNewIndex - 1;
		}

		sAddline = frmDefinition.ssOleDBGridSortOrder.columns(0).text + 
			'	' + frmDefinition.ssOleDBGridSortOrder.columns(1).text + 
			'	' + frmDefinition.ssOleDBGridSortOrder.columns(2).text 
	
		frmDefinition.ssOleDBGridSortOrder.additem(sAddline, iNewIndex);
		frmDefinition.ssOleDBGridSortOrder.RemoveItem(iOldIndex);

		frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();
		frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(iSelectIndex));
		frmDefinition.ssOleDBGridSortOrder.Bookmark = frmDefinition.ssOleDBGridSortOrder.AddItemBookmark(iSelectIndex);

		frmUseful.txtChanged.value = 1;
		refreshTab4Controls();
	}

	function validateTab1()
	{
		// check name has been entered
		if (frmDefinition.txtName.value == '') 
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must enter a name for this definition.",48,"Calendar Reports");
			displayPage(1);
			return (false);
		}
      
		// check base picklist
		if ((frmDefinition.optRecordSelection2.checked == true) &&
			(frmDefinition.txtBasePicklistID.value == 0)) 
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a picklist for the base table.",48,"Calendar Reports");
			displayPage(1);
			return (false);
		}

		// check base filter
		if ((frmDefinition.optRecordSelection3.checked == true) &&
			(frmDefinition.txtBaseFilterID.value == 0)) 
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a filter for the base table.",48,"Calendar Reports");
			displayPage(1);
			return (false);
		}
	
		// Check that a valid description column or a valid calculation has been selected
		if ((frmDefinition.cboDescription1.options[frmDefinition.cboDescription1.selectedIndex].value < 1) && (frmDefinition.txtDescExprID.value < 1) && (frmDefinition.cboDescription2.options[frmDefinition.cboDescription2.selectedIndex].value < 1))
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select at least one base description column or calculation for the report.",48,"Calendar Reports");
			displayPage(1);
			return (false);
		}
	
		return (true);
	}

	function validateTab2()
	{
		var i;
		var sErrMsg;
		var iIndex;
		var iCount;
		var sDefn;
		var sControlName;
		var iPollCounter;
		var iPollPeriod;
		var frmRefresh;
		var iDummy;

		iPollPeriod = 100;
		iPollCounter = iPollPeriod;
		frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");	
	
		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}
	
		sErrMsg = "";
	
		//check at least one column defined as sort order
		if (frmUseful.txtEventsLoaded.value == 1) 
		{
			if (frmDefinition.grdEvents.Rows <= 0) 
			{
				sErrMsg = "You must select at least one event to report on.";
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
					if (i==iPollCounter) 
					{			
						try 
						{
							var testDataCollection = frmRefresh.elements;
							iDummy = testDataCollection.txtDummy.value;
							frmRefresh.submit();
							iPollCounter = iPollCounter + iPollPeriod;
						}
						catch(e) 
						{
						}
					}

					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 19);
					if (sControlName == "txtReportDefnEvent_") 
					{
						sDefn = new String(dataCollection.item(i).value);

						iCount = iCount + 1;
					}
				}	
			}
		
			if (iCount == 0) 
			{
				sErrMsg = "You must select at least one event to report on.";
			}
		}
	
		if (sErrMsg.length > 0) 
		{    
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg,48,"Calendar Reports");
			displayPage(2);
			return (false);
		}
	
		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}
		
		return (true);
	}

	function validateTab3()
	{
		with (frmDefinition)
		{
			if (optFixedStart.checked && (trim(txtFixedStart.value) == ''))
			{
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a Fixed Start Date for the report.",48,"Calendar Reports");
				displayPage(3);
				return (false);
			}
			if (optFixedStart.checked && (!validateDate(txtFixedStart)))
			{
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must enter a valid Fixed Start Date for the report.",48,"Calendar Reports");
				displayPage(3);
				return (false);
			}
			
			if (optFixedEnd.checked && (trim(txtFixedEnd.value) == ''))
			{
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a Fixed End Date for the report.",48,"Calendar Reports");
				displayPage(3);
				return (false);
			}
			if (optFixedEnd.checked && (!validateDate(txtFixedEnd)))
			{
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must enter a valid Fixed End Date for the report.",48,"Calendar Reports");
				displayPage(3);
				return (false);
			}
		
			if (optFixedStart.checked && optFixedEnd.checked)
			{
				var dtStartDate = convertLocaleDateToDateObject(txtFixedStart.value);
				var dtEndDate = convertLocaleDateToDateObject(txtFixedEnd.value);

				if (dtEndDate.getTime() < dtStartDate.getTime())
				{
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a Fixed End Date later than or equal to the Fixed Start Date.",48,"Calendar Reports");
					displayPage(3);
					return (false);
				}
			}
		
			if (optOffsetStart.checked && optFixedEnd.checked)
			{
				if (Number(txtFreqEnd.value) < 0)
				{
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select an End Date Offset greater than or equal to zero.",48,"Calendar Reports");
					displayPage(3);
					return (false);
				}
			}
		
			if (optCurrentStart.checked && optOffsetEnd.checked)
			{
				if (Number(txtFreqEnd.value) < 0)
				{
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select an End Date Offset greater than or equal to zero.",48,"Calendar Reports");
					displayPage(3);
					return (false);
				}
			}
		
			if (optOffsetStart.checked && (optFixedEnd.checked  || optCurrentEnd.checked))
			{
				if (Number(txtFreqStart.value) > 0)
				{
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a Start Date Offset less than or equal to zero.",48,"Calendar Reports");
					displayPage(3);
					return (false);
				}
			}
		
			if (optOffsetStart.checked && optOffsetEnd.checked)
			{
				if (cboPeriodStart.selectedIndex != cboPeriodEnd.selectedIndex)
				{
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select the same End Date Offset period as Start Date Offset period.",48,"Calendar Reports");
					displayPage(3);
					return (false);
				}
			
				if (Number(txtFreqEnd.value) < Number(txtFreqStart.value))
				{
					window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select an End Date Offset greater than or equal to the Start Date Offset.",48,"Calendar Reports");
					displayPage(3);
					return (false);
				}
			}
		
			if (optCustomStart.checked && (txtCustomStartID.value < 1))
			{
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a calculation for the Report Start Date.",48,"Calendar Reports");
				displayPage(3);
				return (false);
			}
		
			if (optCustomEnd.checked && (txtCustomEndID.value < 1))
			{
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You must select a calculation for the Report End Date.",48,"Calendar Reports");
				displayPage(3);
				return (false);
			}
		}

		return (true);
	}  


	function validateTab4()
	{
		var i;
		var sErrMsg;
		var iIndex;
		var iCount;
		var sDefn;
		var sControlName;
		var iPollCounter;
		var iPollPeriod;
		var frmRefresh;
		var iDummy;

		iPollPeriod = 100;
		iPollCounter = iPollPeriod;
		frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");	
	
		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}
	
		sErrMsg = "";
	
		//check at least one column defined as sort order
		if (frmUseful.txtSortLoaded.value == 1) 
		{
			with (frmDefinition.ssOleDBGridSortOrder)
			{
				if (frmDefinition.ssOleDBGridSortOrder.Rows <= 0) 
				{
					sErrMsg = "You must select at least one column to order the report by.";
				}
				else
				{
					if ((frmDefinition.chkGroupByDesc.checked) && (frmDefinition.txtDescExprID.value < 1))
					{
						var lngDesc1ID = frmDefinition.cboDescription1.options[frmDefinition.cboDescription1.selectedIndex].value;
						var lngDesc2ID = frmDefinition.cboDescription2.options[frmDefinition.cboDescription2.selectedIndex].value;
					
						var strDesc = new String(lngDesc1ID);
						if (lngDesc2ID > 0)
						{
							strDesc = strDesc + '	' + lngDesc2ID;
						}
						var strTemp = new String('');
						Redraw = false;
						MoveFirst();
						for (var i=0; i<Rows; i++)
						{
							if (Columns('ColumnID').Text > 0)
							{
								if (i > 0)
								{
									strTemp = strTemp + '	';
								}
								strTemp = strTemp + Columns('ColumnID').Value;
							}
							if (i >= 1)
							{
								break;
							}
							MoveNext();
						}
						MoveFirst();
					
						if (strTemp != strDesc)
						{
							sErrMsg = "The sort order does not reflect the selected Group By Description columns. Do you wish to continue?";
							if (window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg,36,"Calendar Report Definition") == 7)
							{
								Redraw = true;
								displayPage(4);
								return (false);
							}
							else
							{
								Redraw = true;
								sErrMsg = '';
							}
						}					
					}
				}
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
					if (i==iPollCounter) 
					{			
						try 
						{
							var testDataCollection = frmRefresh.elements;
							iDummy = testDataCollection.txtDummy.value;
							frmRefresh.submit();
							iPollCounter = iPollCounter + iPollPeriod;
						}
						catch(e) 
						{
						}
					}

					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 19);
					if (sControlName == "txtReportDefnOrder_") 
					{
						sDefn = new String(dataCollection.item(i).value);

						iCount = iCount + 1;
					}
				}	
			}
		
			if (iCount == 0) 
			{
				sErrMsg = "You must select at least one column to order the report by.";
			}
		}
	
		if (sErrMsg.length > 0) 
		{    
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg,48,"Calendar Reports");
			displayPage(4);
			return (false);
		}
	
		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}
		
		return (true);
	}

	function validateTab5()
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

		if ((frmDefinition.txtFilename.value == "") 
			&& (frmDefinition.cmdFilename.disabled == false)) 
		{
			sErrMsg = "You must enter a file name";
		}

		if ((frmDefinition.txtEmailGroup.value == "") 
			&& (frmDefinition.cmdEmailGroup.disabled == false)) 
		{
			sErrMsg = "You must select an email group";
		}
	
		if ((frmDefinition.chkDestination3 .checked) 
			&& (frmDefinition.txtEmailAttachAs.value == ''))
		{
			sErrMsg = "You must enter an email attachment file name.";
		}

		if (frmDefinition.chkDestination3.checked &&
			(frmDefinition.optOutputFormat3.checked || frmDefinition.optOutputFormat4.checked || frmDefinition.optOutputFormat5.checked || frmDefinition.optOutputFormat6.checked) &&
			frmDefinition.txtEmailAttachAs.value.match(/.html$/)) {
			sErrMsg = "You cannot email html output from word or excel.";
		}

		if (sErrMsg.length > 0) 
		{    
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg,48,"Calendar Reports");
			displayPage(5);
			return (false);
		}
	
		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
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
		var iPollCounter;
		var iPollPeriod;
		var frmRefresh;
		var iDummy;
		var varBookmark;
		var iLoop;
		var sAccess;

		iPollPeriod = 100;
		iPollCounter = iPollPeriod;
		frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");	
	
		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}

		/******************** TAB 1 - DEFINITION *********************/

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
		if (frmDefinition.optRecordSelection1.checked == true) 
		{
			frmSend.txtSend_allRecords.value = "1";
		}
		if (frmDefinition.optRecordSelection2.checked == true) 
		{
			frmSend.txtSend_picklist.value = frmDefinition.txtBasePicklistID.value;
		}
		if (frmDefinition.optRecordSelection3.checked == true) 
		{
			frmSend.txtSend_filter.value = frmDefinition.txtBaseFilterID.value;
		}

		if (frmDefinition.chkPrintFilterHeader.checked == true) 
		{
			frmSend.txtSend_printFilterHeader.value = '1';
		}
		else 
		{
			frmSend.txtSend_printFilterHeader.value = '0';
		}
		
		frmSend.txtSend_desc1.value = frmDefinition.cboDescription1.options(frmDefinition.cboDescription1.options.selectedIndex).value;
		frmSend.txtSend_desc2.value = frmDefinition.cboDescription2.options(frmDefinition.cboDescription2.options.selectedIndex).value;
	
		if (frmDefinition.txtDescExprID.value > 0)
		{
			frmSend.txtSend_descExpr.value = frmDefinition.txtDescExprID.value;
		}
		else
		{
			frmSend.txtSend_descExpr.value = 0;
		}
		
		frmSend.txtSend_region.value = frmDefinition.cboRegion.options(frmDefinition.cboRegion.options.selectedIndex).value;
	
		if (frmDefinition.chkGroupByDesc.checked)
		{
			frmSend.txtSend_groupbydesc.value = 1;
		}
		else
		{
			frmSend.txtSend_groupbydesc.value = 0;
		}
		
		if (frmDefinition.cboDescriptionSeparator.selectedIndex < 0)
		{
			frmSend.txtSend_descseparator.value = ', ';
		}
		else
		{
			frmSend.txtSend_descseparator.value = frmDefinition.cboDescriptionSeparator.options[frmDefinition.cboDescriptionSeparator.selectedIndex].value;
		}

		/*************************************************************/

		/******************* TAB 2 - EVENT DETAILS *******************/

		// now go through the columns grid (and sort order grid)(and the repetition grid)
		var sEvents = '';

		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}

		frmUseful.txtLockGridEvents.value = 1;

		if (frmUseful.txtEventsLoaded.value == 1) 
		{
			frmDefinition.grdEvents.Redraw = false;
			frmDefinition.grdEvents.movefirst();

			iPollCounter = iPollPeriod;
			for (var i=0; i < frmDefinition.grdEvents.rows; i++) 
			{
				if (i==iPollCounter) 
				{			
					try 
					{
						var testDataCollection = frmRefresh.elements;
						iDummy = testDataCollection.txtDummy.value;
						frmRefresh.submit();
						iPollCounter = iPollCounter + iPollPeriod;
					}
					catch(e) 
					{
					}
				}

				sEvents = sEvents + 
										 trim(frmDefinition.grdEvents.columns("EventKey").text) +
							'||' + frmDefinition.grdEvents.columns("Name").text +
							'||' + frmDefinition.grdEvents.columns("TableID").text +
							'||' + frmDefinition.grdEvents.columns("FilterID").text + 
							'||' + frmDefinition.grdEvents.columns("StartDateID").text +
							'||' + frmDefinition.grdEvents.columns("StartSessionID").text +
							'||' + frmDefinition.grdEvents.columns("EndDateID").text +
							'||' + frmDefinition.grdEvents.columns("EndSessionID").text +
							'||' + frmDefinition.grdEvents.columns("DurationID").text
						
				if (frmDefinition.grdEvents.columns("LegendType").text == '1')
				{
					sEvents = sEvents +
						'||' + '1' + 
						'||' + '' +
						'||' + frmDefinition.grdEvents.columns("LegendTableID").text +
						'||' + frmDefinition.grdEvents.columns("LegendColumnID").text +
						'||' + frmDefinition.grdEvents.columns("LegendCodeID").text +
						'||' + frmDefinition.grdEvents.columns("LegendEventTypeID").text
				}
				else
				{
					sEvents = sEvents +
						'||' + '0' + 
						'||' + replace(frmDefinition.grdEvents.columns("Legend").text,"'","''") +
						'||' + 0 +
						'||' + 0 +
						'||' + 0 +
						'||' + 0
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
		else 
		{
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection!=null) 
			{
				iNum = 0;
				iPollCounter = iPollPeriod;

				for (iIndex=0; iIndex<dataCollection.length; iIndex++)  
				{
					if (iIndex==iPollCounter) 
					{			
						try 
						{
							var testDataCollection = frmRefresh.elements;
							iDummy = testDataCollection.txtDummy.value;
							frmRefresh.submit();
							iPollCounter = iPollCounter + iPollPeriod;
						}
						catch(e) 
						{
						}
					}

					sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 19);
					if (sControlName == "txtReportDefnEvent_") 
					{
						sEvents = sEvents + 
											 trim(selectedEventParameter(dataCollection.item(iIndex).value,"EVENTKEY")) +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"NAME") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"TABLEID") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"FILTERID") + 
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"STARTDATEID") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"STARTSESSIONID") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"ENDDATEID") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"ENDSESSIONID") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"DURATIONID")
						
						if (selectedEventParameter(dataCollection.item(iIndex).value,"LEGENDTYPE") == '1') 
						{
							sEvents = sEvents + 
								'||' + '1' +
								'||' + '' +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"LEGENDLOOKUPTABLEID") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"LEGENDLOOKUPCOLUMNID") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"LEGENDLOOKUPCODEID") +
								'||' + selectedEventParameter(dataCollection.item(iIndex).value,"LEGENDEVENTCOLUMNID")
						}
						else
						{
							sEvents = sEvents + 
								'||' + '0' +
								'||' + replace(selectedEventParameter(dataCollection.item(iIndex).value,"LEGEND"),"'","''") +
								'||' + '0' +
								'||' + '0' +
								'||' + '0' +
								'||' + '0'
						}
						
						sEvents = sEvents + 
							'||' + selectedEventParameter(dataCollection.item(iIndex).value,"DESC1COLUMNID") +
							'||' + selectedEventParameter(dataCollection.item(iIndex).value,"DESC2COLUMNID") +
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

		if (frmDefinition.optFixedStart.checked == true)
		{
			frmSend.txtSend_StartType.value = 0;
			frmSend.txtSend_FixedStart.value = convertLocaleDateToSQL(frmDefinition.txtFixedStart.value);
			frmSend.txtSend_StartFrequency.value = 0;
			frmSend.txtSend_StartPeriod.value = -1;
			frmSend.txtSend_CustomStart.value = 0;
		}
		else if (frmDefinition.optCurrentStart.checked == true)
		{
			frmSend.txtSend_StartType.value = 1;
			frmSend.txtSend_FixedStart.value = ""
			frmSend.txtSend_StartFrequency.value = 0;
			frmSend.txtSend_StartPeriod.value = -1;
			frmSend.txtSend_CustomStart.value = 0;
		}
		else if (frmDefinition.optOffsetStart.checked == true)
		{
			frmSend.txtSend_StartType.value = 2;
			frmSend.txtSend_FixedStart.value = ""
			frmSend.txtSend_StartFrequency.value = frmDefinition.txtFreqStart.value;
			frmSend.txtSend_StartPeriod.value = frmDefinition.cboPeriodStart.options[frmDefinition.cboPeriodStart.selectedIndex].value;
			frmSend.txtSend_CustomStart.value = 0;
		}
		else if (frmDefinition.optCustomStart.checked == true)
		{
			frmSend.txtSend_StartType.value = 3;
			frmSend.txtSend_FixedStart.value = ""
			frmSend.txtSend_StartFrequency.value = 0;
			frmSend.txtSend_StartPeriod.value = -1;
			frmSend.txtSend_CustomStart.value = frmDefinition.txtCustomStartID.value;
		}
	
		if (frmDefinition.optFixedEnd.checked == true)
		{
			frmSend.txtSend_EndType.value = 0;
			frmSend.txtSend_FixedEnd.value = convertLocaleDateToSQL(frmDefinition.txtFixedEnd.value);
			frmSend.txtSend_EndFrequency.value = 0;
			frmSend.txtSend_EndPeriod.value = -1;
			frmSend.txtSend_CustomEnd.value = 0;
		}
		else if (frmDefinition.optCurrentEnd.checked == true)
		{
			frmSend.txtSend_EndType.value = 1;
			frmSend.txtSend_FixedEnd.value = "";
			frmSend.txtSend_EndFrequency.value = 0;
			frmSend.txtSend_EndPeriod.value = -1;
			frmSend.txtSend_CustomEnd.value = 0;
		}
		else if (frmDefinition.optOffsetEnd.checked == true)
		{
			frmSend.txtSend_EndType.value = 2;
			frmSend.txtSend_FixedEnd.value = "";
			frmSend.txtSend_EndFrequency.value = frmDefinition.txtFreqEnd.value;
			frmSend.txtSend_EndPeriod.value = frmDefinition.cboPeriodEnd.options[frmDefinition.cboPeriodEnd.selectedIndex].value;
			frmSend.txtSend_CustomEnd.value = 0;
		}
		else if (frmDefinition.optCustomEnd.checked == true)
		{
			frmSend.txtSend_EndType.value = 3;
			frmSend.txtSend_FixedEnd.value = "";
			frmSend.txtSend_EndFrequency.value = 0;
			frmSend.txtSend_EndPeriod.value = -1;
			frmSend.txtSend_CustomEnd.value = frmDefinition.txtCustomEndID.value;
		}
		
		if (frmDefinition.chkIncludeBHols.checked == true) 
		{
			frmSend.txtSend_IncludeBHols.value = '1';
		}
		else 
		{
			frmSend.txtSend_IncludeBHols.value = '0';
		}
		
		if (frmDefinition.chkIncludeWorkingDaysOnly.checked == true) 
		{
			frmSend.txtSend_IncludeWorkingDaysOnly.value = '1';
		}
		else 
		{
			frmSend.txtSend_IncludeWorkingDaysOnly.value = '0';
		}
	
		if (frmDefinition.chkShadeBHols.checked == true) 
		{
			frmSend.txtSend_ShadeBHols.value = '1';
		}
		else 
		{
			frmSend.txtSend_ShadeBHols.value = '0';
		}
	
		if (frmDefinition.chkCaptions.checked == true) 
		{
			frmSend.txtSend_Captions.value = '1';
		}
		else 
		{
			frmSend.txtSend_Captions.value = '0';
		}
	
		if (frmDefinition.chkShadeWeekends.checked == true) 
		{
			frmSend.txtSend_ShadeWeekends.value = '1';
		}
		else 
		{
			frmSend.txtSend_ShadeWeekends.value = '0';
		}

		if (frmDefinition.chkStartOnCurrentMonth.checked == true) 
		{
			frmSend.txtSend_StartOnCurrentMonth.value = '1';
		}
		else 
		{
			frmSend.txtSend_StartOnCurrentMonth.value = '0';
		}
		
		/*************************************************************/

		/********************* TAB 4 - SORT ORDER ********************/

		/*now use the txtSend_OrderString to hold the string of selected order information*/
		if (frmUseful.txtSortLoaded.value == 1) 
		{ 
			var sOrders = '';
			var i;
			var pvarbookmark;
	
			with (frmDefinition.ssOleDBGridSortOrder)
			{
				if (Rows > 0)
				{
					Redraw = false;
					MoveFirst();
					
					for (var i=0; i < Rows; i++)
					{
						if (i==iPollCounter) 
						{			
							try 
							{
								var testDataCollection = frmRefresh.elements;
								iDummy = testDataCollection.txtDummy.value;
								frmRefresh.submit();
								iPollCounter = iPollCounter + iPollPeriod;
							}
							catch(e) 
							{
							}
						}
					
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
		else
		{
			var dataCollection = frmOriginalDefinition.elements;
		
			var sOrders = '';
			var sDefnString = '';
		
			if (dataCollection!=null) 
			{
				for (var i=1; i<=frmUseful.txtOrderCount.value; i++)  
				{
					if (i==iPollCounter) 
					{			
						try 
						{
							var testDataCollection = frmRefresh.elements;
							iDummy = testDataCollection.txtDummy.value;
							frmRefresh.submit();
							iPollCounter = iPollCounter + iPollPeriod;
						}
						catch(e) 
						{
						}
					}
				
					sDefnString = document.getElementById('txtReportDefnOrder_'+i).value;
				
					sOrders = sOrders + sortColumnParameter(sDefnString,"COLUMNID") + '||';
					sOrders = sOrders + i + '||';
					sOrders = sOrders + sortColumnParameter(sDefnString,"ORDER") + '||';
						
					sOrders = sOrders + '**';
				
				}
			}
			frmSend.txtSend_OrderString.value = sOrders;
		}
	
		/*************************************************************/

		/****************** TAB 5 - OUTPUT OPTIONS *******************/
	
		if (frmDefinition.chkPreview.checked == true)
		{
			frmSend.txtSend_OutputPreview.value = 1;
		}
		else
		{
			frmSend.txtSend_OutputPreview.value = 0;
		}
	
		if (frmDefinition.optOutputFormat0.checked)	frmSend.txtSend_OutputFormat.value = 0;
		//if (frmDefinition.optOutputFormat1.checked)	frmSend.txtSend_OutputFormat.value = 1;
		if (frmDefinition.optOutputFormat2.checked)	frmSend.txtSend_OutputFormat.value = 2;
		if (frmDefinition.optOutputFormat3.checked)	frmSend.txtSend_OutputFormat.value = 3;
		if (frmDefinition.optOutputFormat4.checked)	frmSend.txtSend_OutputFormat.value = 4;
	
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
	
		/*************************************************************/
	
		frmUseful.txtLockGridEvents.value = 0;

		try 
		{
			var testDataCollection = frmRefresh.elements;
			iDummy = testDataCollection.txtDummy.value;
			frmRefresh.submit();
		}
		catch(e) 
		{
		}

		if (sEvents.length > 16000) 
		{
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Too many events selected.",48,"Calendar Reports");
			return false;
		}
		else 
		{
			return true;
		}

	}

	function getOrderString()
	{
		return true;
		var sOrders = '';
		var i;
		var pvarbookmark;
	
		with (frmDefinition.ssOleDBGridSortOrder)
		{
			if (Rows > 0)
			{
				MoveFirst();
				for (var i=0; i < Rows; i++)
				{
				
					sOrders = sOrders + frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value + '||';
					sOrders = sOrders + Columns('columnID').text + '||';
					sOrders = sOrders + i + '||';
					sOrders = sOrders + Columns('order').text + '||';
					
					sOrders = sOrders + '**';
					MoveNext();
				}
				return sOrders;
			}
			else
			{
				return '';
			}
		}
	}

	function loadAvailableColumns()
	{
		var i;
		var sSelectedIDs;
		var sTemp;
		var iIndex;
		var sType;
		var sID;
		var iPollPeriod;
		var iPollCounter;
		var iDummy;
		var frmRefresh;

		var blnPersonnelBaseTable = (frmUseful.txtPersonnelTableID.value == frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value)

		frmUseful.txtLockGridEvents.value = 1;
	
		iPollPeriod = 100;
		iPollCounter = iPollPeriod;
		frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");	

		frmDefinition.cboDescription1.length = 0;
		frmDefinition.cboDescription2.length = 0;
		frmDefinition.cboRegion.length = 0;

		var oOption = document.createElement("OPTION");
		frmDefinition.cboDescription1.options.add(oOption);
		oOption.innerText = "<None>";
		oOption.value = 0;
		if (frmUseful.txtAction.value.toUpperCase() == "NEW") oOption.selected = true;

		var oOption = document.createElement("OPTION");
		frmDefinition.cboDescription2.options.add(oOption);
		oOption.innerText = "<None>";
		oOption.value = 0;
		if (frmUseful.txtAction.value.toUpperCase() == "NEW") oOption.selected = true;
								
		var oOption = document.createElement("OPTION");
		frmDefinition.cboRegion.options.add(oOption);
	

		if (blnPersonnelBaseTable) 
		{
			oOption.innerText = "<Default>";
		}
		else 
		{
			oOption.innerText = "<None>";
		}
		oOption.value = 0;
		if (frmUseful.txtAction.value.toUpperCase() == "NEW") oOption.selected = true;
	
		var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
		var dataCollection = frmUtilDefForm.elements;
	
		if (dataCollection!=null) 
		{
			for (i=0; i<dataCollection.length; i++)  
			{
				if (i==iPollCounter) 
				{			
					try 
					{
						var testDataCollection = frmRefresh.elements;
						iDummy = testDataCollection.txtDummy.value;
						frmRefresh.submit();
						iPollCounter = iPollCounter + iPollPeriod;
					}
					catch(e) 
					{
					}
				}

				sControlName = dataCollection.item(i).name;
			
				if (sControlName.substr(0, 10) == "txtRepCol_") 
				{
					sColumnID = sControlName.substring(10, sControlName.length);
					
					var oOption = document.createElement("OPTION");
					frmDefinition.cboDescription1.options.add(oOption);
					oOption.innerText = dataCollection.item(i).value;
					oOption.value = sColumnID;
			
					if ((frmUseful.txtAction.value.toUpperCase() != "NEW")
							&& (frmUseful.txtFirstLoad.value == 'Y'))
					{
						if (sColumnID == frmOriginalDefinition.txtDefn_Desc1ID.value)
						{
							oOption.selected = true;
						}
					}
				
					var oOption = document.createElement("OPTION");
					frmDefinition.cboDescription2.options.add(oOption);
					oOption.innerText = dataCollection.item(i).value;
					oOption.value = sColumnID;

					if ((frmUseful.txtAction.value.toUpperCase() != "NEW")
							&& (frmUseful.txtFirstLoad.value == 'Y'))
					{
						if (sColumnID == frmOriginalDefinition.txtDefn_Desc2ID.value)
						{
							oOption.selected = true;
						}
					}
				
					/* Only add varchar columns to the region column. */
					var sDataTypeControlName = "txtRepColDataType_" + sColumnID;
					var iDataTypeControlValue = frmUtilDefForm.elements(sDataTypeControlName).value;

					if (iDataTypeControlValue == 12)
					{
						var oOption = document.createElement("OPTION");
						frmDefinition.cboRegion.options.add(oOption);
						oOption.innerText = dataCollection.item(i).value;
						oOption.value = sColumnID;

						if ((frmUseful.txtAction.value.toUpperCase() != "NEW")
								&& (frmUseful.txtFirstLoad.value == 'Y'))
						{
							if (sColumnID == frmOriginalDefinition.txtDefn_RegionID.value)
							{
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
		
		// Little dodge to get around a browser bug that
		// does not refresh the display on all controls.
		try
		{
			window.resizeBy(0,-1);
			window.resizeBy(0,1);
		}
		catch(e) {}
	
		if (frmUseful.txtAvailableColumnsLoaded.value == 1)
		{
			if (frmDefinition.txtName.disabled == false) 
			{
				frmDefinition.focus();
				try 
				{
					frmDefinition.txtName.focus();
				}
				catch (e) {}
			}
		}

		// Get menu to refresh the menu.
		window.parent.frames("menuframe").refreshMenu();		 
		refreshTab1Controls();		  
		frmUseful.txtLoading.value = 'N';

		if (frmDefinition.txtName.disabled == false) 
		{
			frmDefinition.focus();
			try 
			{
				frmDefinition.txtName.focus();
			}
			catch (e) {}
		}
	}

	function loadDefinition()
	{
		frmDefinition.txtName.value = frmOriginalDefinition.txtDefn_Name.value;
	
		if((frmUseful.txtAction.value.toUpperCase() == "EDIT") ||
			(frmUseful.txtAction.value.toUpperCase() == "VIEW")) 
		{
			frmDefinition.txtOwner.value = frmOriginalDefinition.txtDefn_Owner.value;
		}
		else 
		{
			frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
		}

		frmDefinition.txtDescription.value= frmOriginalDefinition.txtDefn_Description.value;

		setBaseTable(frmOriginalDefinition.txtDefn_BaseTableID.value);
		changeBaseTable();	

		// Set the basic record selection.
		fRecordOptionSet = false;

		if (frmOriginalDefinition.txtDefn_PicklistID.value > 0) 
		{
			button_disable(frmDefinition.cmdBasePicklist, false);
			frmDefinition.optRecordSelection2.checked = true;
			frmDefinition.txtBasePicklistID.value = frmOriginalDefinition.txtDefn_PicklistID.value;
			frmDefinition.txtBasePicklist.value = frmOriginalDefinition.txtDefn_PicklistName.value;
			fRecordOptionSet = true;
		}
		else 
		{
			if (frmOriginalDefinition.txtDefn_FilterID.value > 0) 
			{
				button_disable(frmDefinition.cmdBaseFilter, false);
				frmDefinition.optRecordSelection3.checked = true;
				frmDefinition.txtBaseFilterID.value = frmOriginalDefinition.txtDefn_FilterID.value;
				frmDefinition.txtBaseFilter.value = frmOriginalDefinition.txtDefn_FilterName.value;
				fRecordOptionSet = true;			
			}
		}
	
		if (fRecordOptionSet == false) 
		{
			frmDefinition.optRecordSelection1.checked = true;
		}	

		// Print Filter Header ?
		frmDefinition.chkPrintFilterHeader.checked = ((frmOriginalDefinition.txtDefn_PrintFilterHeader.value != "False") &&
																							((frmOriginalDefinition.txtDefn_FilterID.value > 0) || 
																								(frmOriginalDefinition.txtDefn_PicklistID.value > 0)));
		
		frmDefinition.txtDescExpr.value = frmOriginalDefinition.txtDefn_DescExprName.value;
		frmDefinition.txtDescExprID.value = frmOriginalDefinition.txtDefn_DescExprID.value;
	
		frmDefinition.chkGroupByDesc.checked = (frmOriginalDefinition.txtDefn_GroupByDesc.value != "False");
	
		for (var i=0; i<frmDefinition.cboDescriptionSeparator.options.length; i++)
		{
			if (frmDefinition.cboDescriptionSeparator.options[i].value == frmOriginalDefinition.txtDefn_DescSeparator.value)
			{
				frmDefinition.cboDescriptionSeparator.selectedIndex = i;
				break;
			}
		}

		if (frmOriginalDefinition.txtDefn_StartType.value == "0")
		{
			frmDefinition.optFixedStart.checked = true;
			frmDefinition.txtFixedStart.value = frmOriginalDefinition.txtDefn_FixedStart.value;
			frmDefinition.cboPeriodStart.selectedIndex = 0;
			frmDefinition.txtFreqStart.value = 0;
			frmDefinition.txtCustomStart.value = '';
			frmDefinition.txtCustomStartID.value = 0;
		}
		else if (frmOriginalDefinition.txtDefn_StartType.value == "1")
		{
			frmDefinition.optCurrentStart.checked = true;
			frmDefinition.txtFixedStart.value = '';
			frmDefinition.cboPeriodStart.selectedIndex = 0;
			frmDefinition.txtFreqStart.value = 0;
			frmDefinition.txtCustomStart.value = '';
			frmDefinition.txtCustomStartID.value = 0;
		}
		else if (frmOriginalDefinition.txtDefn_StartType.value == "2")
		{
			frmDefinition.optOffsetStart.checked = true;
			frmDefinition.txtFixedStart.value = '';
			frmDefinition.cboPeriodStart.value = frmOriginalDefinition.txtDefn_StartPeriod.value;
			frmDefinition.txtFreqStart.value = frmOriginalDefinition.txtDefn_StartFrequency.value;
			frmDefinition.txtCustomStart.value = '';
			frmDefinition.txtCustomStartID.value = 0;
		}
		else if (frmOriginalDefinition.txtDefn_StartType.value == "3")
		{
			frmDefinition.optCustomStart.checked = true;
			frmDefinition.txtFixedStart.value = '';
			frmDefinition.cboPeriodStart.selectedIndex = 0;
			frmDefinition.txtFreqStart.value = 0;
			frmDefinition.txtCustomStart.value = frmOriginalDefinition.txtDefn_CustomStartName.value;
			frmDefinition.txtCustomStartID.value = frmOriginalDefinition.txtDefn_CustomStartID.value;
		}
		else 
		{
			frmDefinition.optFixedStart.checked = true;
			frmDefinition.txtFixedStart.value = '';
			frmDefinition.cboPeriodStart.selectedIndex = 0;
			frmDefinition.txtFreqStart.value = 0;
			frmDefinition.txtCustomStart.value = '';
			frmDefinition.txtCustomStartID.value = 0;
		}

		if (frmOriginalDefinition.txtDefn_EndType.value == "0")
		{
			frmDefinition.optFixedEnd.checked = true;
			frmDefinition.txtFixedEnd.value = frmOriginalDefinition.txtDefn_FixedEnd.value;
			frmDefinition.cboPeriodEnd.selectedIndex = 0;
			frmDefinition.txtFreqEnd.value = 0;
			frmDefinition.txtCustomEnd.value = '';
			frmDefinition.txtCustomEndID.value = 0;
		}
		else if (frmOriginalDefinition.txtDefn_EndType.value == "1")
		{
			frmDefinition.optCurrentEnd.checked = true;
			frmDefinition.txtFixedEnd.value = '';
			frmDefinition.cboPeriodEnd.selectedIndex = 0;
			frmDefinition.txtFreqEnd.value = 0;
			frmDefinition.txtCustomEnd.value = '';
			frmDefinition.txtCustomEndID.value = 0;
		}
		else if (frmOriginalDefinition.txtDefn_EndType.value == "2")
		{
			frmDefinition.optOffsetEnd.checked = true;
			frmDefinition.txtFixedEnd.value = '';
			frmDefinition.cboPeriodEnd.selectedIndex = frmOriginalDefinition.txtDefn_EndPeriod.value;
			frmDefinition.txtFreqEnd.value = frmOriginalDefinition.txtDefn_EndFrequency.value;
			frmDefinition.txtCustomEnd.value = '';
			frmDefinition.txtCustomEndID.value = 0;
		}
		else if (frmOriginalDefinition.txtDefn_EndType.value == "3")
		{
			frmDefinition.optCustomEnd.checked = true;
			frmDefinition.txtFixedEnd.value = '';
			frmDefinition.cboPeriodEnd.selectedIndex = 0;
			frmDefinition.txtFreqEnd.value = 0;
			frmDefinition.txtCustomEnd.value = frmOriginalDefinition.txtDefn_CustomEndName.value;
			frmDefinition.txtCustomEndID.value = frmOriginalDefinition.txtDefn_CustomEndID.value;
		}
		else 
		{
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
	
		if (frmOriginalDefinition.txtDefn_OutputFormat.value == 0)
		{
			frmDefinition.optOutputFormat0.checked = true;
		}
			/*else if (frmOriginalDefinition.txtDefn_OutputFormat.value == 1)
				{
				frmDefinition.optOutputFormat1.checked = true;
				}*/
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
		else 
		{
			frmDefinition.optOutputFormat0.checked = true;
		}
		
		frmDefinition.chkDestination0.checked = (frmOriginalDefinition.txtDefn_OutputScreen.value != "False");
		frmDefinition.chkDestination1.checked = (frmOriginalDefinition.txtDefn_OutputPrinter.value != "False");

		if (frmDefinition.chkDestination1.checked == true)
		{
			populatePrinters();
			for (i=0; i<frmDefinition.cboPrinterName.options.length; i++)  
			{
				if (frmDefinition.cboPrinterName.options(i).innerText == frmOriginalDefinition.txtDefn_OutputPrinterName.value)
				{
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
	

		frmDefinition.grdEvents.MoveFirst();
		frmDefinition.grdEvents.FirstRow = frmDefinition.grdEvents.Bookmark;
	
		frmDefinition.ssOleDBGridSortOrder.Movefirst();
		frmDefinition.ssOleDBGridSortOrder.FirstRow = frmDefinition.ssOleDBGridSortOrder.bookmark;
	
		// If its read only, disable everything.
		if(frmUseful.txtAction.value.toUpperCase() == "VIEW")
		{
			disableAll();
		}
		
	}

	function loadEventsDefinition()
	{
		var iIndex;
		var sDefnString;
		var iPollCounter;
		var iPollPeriod;
		var frmRefresh;

		iPollPeriod = 100;
		iPollCounter = iPollPeriod;
		frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");	
	
		if (frmUseful.txtEventsLoaded.value == 0) 
		{
			frmDefinition.grdEvents.focus();
			frmDefinition.grdEvents.Redraw = false;
		
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection!=null) 
			{
				for (iIndex=0; iIndex<dataCollection.length; iIndex++)  
				{
					if (iIndex==iPollCounter) 
					{		
						try 
						{
							var testDataCollection = frmRefresh.elements;
							iDummy = testDataCollection.txtDummy.value;
							frmRefresh.submit();
							iPollCounter = iPollCounter + iPollPeriod;
						}
						catch(e) 
						{
						}
					}

					sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 19);
					if (sControlName == "txtReportDefnEvent_") 
					{
						sDefnString = new String(dataCollection.item(iIndex).value);
					
						if (sDefnString.length > 0) 
						{
							frmDefinition.grdEvents.AddItem(sDefnString);
						}
					}				
				}	
			}
			frmDefinition.grdEvents.Redraw = true;
			frmUseful.txtEventsLoaded.value = 1;
		}
	}

	function loadSortDefinition()
	{
		var iIndex;
	
		if (frmUseful.txtSortLoaded.value == 0) 
		{
			frmDefinition.ssOleDBGridSortOrder.focus();
			frmDefinition.ssOleDBGridSortOrder.Redraw = false;
		
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection!=null) 
			{
				for (iIndex=0; iIndex<dataCollection.length; iIndex++)  
				{
					sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 19);
					if (sControlName == "txtReportDefnOrder_") 
					{
						sDefnString = new String(dataCollection.item(iIndex).value);
					
						if (sDefnString.length > 0) 
						{
							frmDefinition.ssOleDBGridSortOrder.AddItem(sDefnString);
						}
					}
				}	
			}
		
			frmDefinition.ssOleDBGridSortOrder.Redraw = true;
			frmUseful.txtSortLoaded.value = 1;
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
		
		sDateFormat = window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDateFormat;

		sDays="";
		sMonths="";
		sYears="";
		iValuePos = 0;

		// Trim leading spaces.
		sTempValue = psDateString.substr(iValuePos,1);
		while (sTempValue.charAt(0) == " ") {
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
	
		sValue = window.parent.frames("menuframe").ASRIntranetFunctions.ConvertSQLDateToLocale(sTempValue);

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

	function convertLocaleDateToDateObject(psDateString)
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
		
		sDateFormat = window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDateFormat;

		sDays="";
		sMonths="";
		sYears="";
		iValuePos = 0;

		// Skip non-numerics
		sTempValue = psDateString.substr(iValuePos,1);
		while (isNaN(sTempValue) == true) {
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
	
		sValue = window.parent.frames("menuframe").ASRIntranetFunctions.ConvertSQLDateToLocale(sTempValue);

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

			for (iLoop=0; iLoop<window.parent.frames("menuframe").ASRIntranetFunctions.PrinterCount(); iLoop++)  {

				var oOption = document.createElement("OPTION");
				options.add(oOption);
				oOption.innerText = window.parent.frames("menuframe").ASRIntranetFunctions.PrinterName(iLoop);
				oOption.value = iLoop+1;

				if (oOption.innerText == strCurrentPrinter) {
					selectedIndex = iLoop+1
				}
			}

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
					(frmDefinition.optOutputFormat6.checked)) {
				var oOption = document.createElement("OPTION");
				options.add(oOption);
				oOption.innerText = "Create new sheet in workbook";
				oOption.value = 4;
			}

			for (iLoop=0; iLoop<options.length; iLoop++)  {
				if (options(iLoop).value == lngCurrentOption) {
					selectedIndex = iLoop
					break;
				}
			}

		}

	}

	function validateDate(pobjDateControl)
	{
		// Date column.
		// Ensure that the value entered is a date.

		var sValue = pobjDateControl.value;
	
		if (sValue.length == 0) 
		{
			return true;
		}
		else 
		{
			// Convert the date to SQL format (use this as a validation check).
			// An empty string is returned if the date is invalid.
			sValue = convertLocaleDateToSQL(sValue);
			if (sValue.length == 0) 
			{
				return false;
			}
			else 
			{
				return true;
			}
		}
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

	function selectedEventParameter(psDefnString, psParameter)
	{
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
	
	function sortColumnParameter(psDefnString, psParameter)
	{
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

	function removeSortColumn(piColumnID, piTableID)
	{
		// Remove the column (if columnID given), 
		// or all columns for a table (if tableID given),
		// or all columns (if no columnID or tableID given).
		// from the sort columns definition.
		var iCount;
		var i;
		var fRemoveRow;

		if (frmUseful.txtSortLoaded.value == 1) 
		{
			if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) 
			{
				frmDefinition.ssOleDBGridSortOrder.Redraw = false;
				frmDefinition.ssOleDBGridSortOrder.MoveFirst();

				iCount = frmDefinition.ssOleDBGridSortOrder.rows;
				for (i=0;i<iCount; i++) 
				{			
					fRemoveRow = true;
				
					if (piColumnID > 0) 
					{
						fRemoveRow = (piColumnID == frmDefinition.ssOleDBGridSortOrder.Columns("id").Text);
					}		
					if (piTableID > 0) 
					{
						fRemoveRow = (piTableID == frmDefinition.ssOleDBGridSortOrder.Columns("tableID").Text);
					}
				
					if (fRemoveRow == true) 
					{
						if (frmDefinition.ssOleDBGridSortOrder.rows == 1) 
						{
							frmDefinition.ssOleDBGridSortOrder.RemoveAll();
						}
						else 
						{
							frmDefinition.ssOleDBGridSortOrder.RemoveItem(frmDefinition.ssOleDBGridSortOrder.AddItemRowIndex(frmDefinition.ssOleDBGridSortOrder.Bookmark));
						}					
					}
					else 
					{
						frmDefinition.ssOleDBGridSortOrder.MoveNext();
					}
				}     

				frmDefinition.ssOleDBGridSortOrder.Redraw = true;
				frmDefinition.ssOleDBGridSortOrder.SelBookmarks.RemoveAll();

				if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) 
				{
					frmDefinition.ssOleDBGridSortOrder.MoveFirst();
					frmDefinition.ssOleDBGridSortOrder.selbookmarks.add(frmDefinition.ssOleDBGridSortOrder.bookmark);
				}
			}
		}
		else 
		{
			var dataCollection = frmOriginalDefinition.elements;
			if (dataCollection!=null) 
			{
				for (iIndex=0; iIndex<dataCollection.length; iIndex++)  
				{
					sControlName = dataCollection.item(iIndex).name;
					sControlName = sControlName.substr(0, 19);
					if (sControlName == "txtReportDefnOrder_") 
					{
						fRemoveRow = true;
						if (piColumnID > 0) 
						{
							fRemoveRow = (piColumnID == sortColumnParameter(dataCollection.item(iIndex).value, "COLUMNID"));
						}		
						if (piTableID > 0) 
						{
							fRemoveRow = (piTableID == sortColumnParameter(dataCollection.item(iIndex).value, "TABLEID"));
						}
										
						if (fRemoveRow == true) 
						{
							dataCollection.item(iIndex).value = "";
						}
					}
				}
			}
		}
	}

	function removeCalcs(psCalcs)
	{
		var iCharIndex;
		var sCalcs;
		var sCalcID;
		var fDone;
	
		sCalcs = new String(psCalcs);

		// Remove the given calcs from the selected columns list.
		while (sCalcs.length > 0) 
		{
			iCharIndex = sCalcs.indexOf(",");
	
			if (iCharIndex >= 0) 
			{
				sCalcID = sCalcs.substr(0, iCharIndex);
				sCalcs = sCalcs.substr(iCharIndex + 1);
			}
			else 
			{
				sCalcID = sCalcs;
				sCalcs = "";
			}
		
			fDone = false;
		
			/* Check if we're removing the description calculation first, then Custom Date1 then Custom Date2. */
			if ((fDone == false) && (frmDefinition.txtDescExprID.value == sCalcID)) 
			{
				frmDefinition.txtDescExpr.value = '';
				frmDefinition.txtDescExprID.value = 0;
				frmSelectionAccess.descHidden.value = "N";
				fDone = true;
			}
			
			if ((fDone == false) && (frmDefinition.txtCustomStartID.value == sCalcID)) 
			{
				frmDefinition.txtCustomStart.value = '';
				frmDefinition.txtCustomStartID.value = 0;
				frmSelectionAccess.calcStartDateHidden.value = "N";
				fDone = true;
			}
		
			if ((fDone == false) && (frmDefinition.txtCustomEndID.value == sCalcID)) 
			{
				frmDefinition.txtCustomEnd.value = '';
				frmDefinition.txtCustomEndID.value = 0;
				frmSelectionAccess.calcEndDateHidden.value = "N";
				fDone = true;
			}
		}

		refreshTab1Controls();
		refreshTab3Controls();
	}

	function removePicklists(psPicklists)	
	{
		var iCharIndex;
		var sPicklists;
		var sPicklistID;
		var fDone;
	
		sPicklists = new String(psPicklists);
	
		// Remove the given calcs from the selected columns list.
		while (sPicklists.length > 0) 
		{
			iCharIndex = sPicklists.indexOf(",");
	
			if (iCharIndex >= 0) 
			{
				sPicklistID = sPicklists.substr(0, iCharIndex);
				sPicklists = sPicklists.substr(iCharIndex + 1);
			}
			else 
			{
				sPicklistID = sPicklists;
				sPicklists = "";
			}
		
			fDone = false;
		
			/* Check if we're removing the base table first, then paretn1 then parent 2, and then the children. */
			if ((fDone == false) && (frmDefinition.txtBasePicklistID.value == sPicklistID)) 
			{
				frmDefinition.txtBasePicklist.value = '';
				frmDefinition.txtBasePicklistID.value = 0;
				frmSelectionAccess.baseHidden.value = "N";
				fDone = true;
			}
		}

		refreshTab1Controls();
	}

	function removeFilters(psEventFilters)
	{
		var iCharIndex;
		var sEventFilters;
		var sEventFilterID;
		var sGridEventFilterID;

		sEventFilters = new String(psEventFilters);
	
		// Remove the given calcs from the selected columns list.
		while (sEventFilters.length > 0) 
		{
			iCharIndex = sEventFilters.indexOf(",");
	
			if (iCharIndex >= 0) 
			{
				sEventFilterID = sEventFilters.substr(0, iCharIndex);
				sEventFilters = sEventFilters.substr(iCharIndex + 1);
			}
			else 
			{
				sEventFilterID = sEventFilters;
				sEventFilters = "";
			}

			if (frmUseful.txtChildsLoaded.value == 1) 
			{
				if (frmDefinition.grdEvents.Rows > 0) 
				{
					frmDefinition.grdEvents.Redraw = false;
					frmDefinition.grdEvents.movefirst();

					for (i=0; i < frmDefinition.grdEvents.rows; i++) 
					{			
						sGridEventFilterID = frmDefinition.grdEvents.Columns("FilterID").Text;
							
						if (sGridEventFilterID == sEventFilterID) 
						{
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

	function createNew(pPopup)
	{	
		pPopup.close();
	
		frmUseful.txtUtilID.value = 0;
		frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
		frmUseful.txtAction.value = "new";
	
		submitDefinition();
	}

	function locateRecord(psSearchFor)
	{  
		var fFound
	
		return;
	
		fFound = false;
	
		frmDefinition.ssOleDBGridAvailableColumns.redraw = false;

		frmDefinition.ssOleDBGridAvailableColumns.MoveLast();
		frmDefinition.ssOleDBGridAvailableColumns.MoveFirst();

		frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.removeall();
	
		for (iIndex = 1; iIndex <= frmDefinition.ssOleDBGridAvailableColumns.rows; iIndex++) 
		{	
			var sGridValue = new String(frmDefinition.ssOleDBGridAvailableColumns.Columns(3).value);
			sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
			if (sGridValue == psSearchFor.toUpperCase()) 
			{
				frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.Add(frmDefinition.ssOleDBGridAvailableColumns.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < frmDefinition.ssOleDBGridAvailableColumns.rows) 
			{
				frmDefinition.ssOleDBGridAvailableColumns.MoveNext();
			}
			else 
			{
				break;
			}
		}

		if ((fFound == false) && (frmDefinition.ssOleDBGridAvailableColumns.rows > 0)) 
		{
			// Select the top row.
			frmDefinition.ssOleDBGridAvailableColumns.MoveFirst();
			frmDefinition.ssOleDBGridAvailableColumns.SelBookmarks.Add(frmDefinition.ssOleDBGridAvailableColumns.Bookmark);
		}

		frmDefinition.ssOleDBGridAvailableColumns.redraw = true;
	}

	function populateAccessGrid()
	{
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
	
		for (var i=0; i<frmDefinition.grdEvents.Rows; i++) {
			vBM = frmDefinition.grdEvents.AddItemBookmark(i)
		
			if (frmDefinition.grdEvents.Columns("FilterHidden").CellValue(vBM) == "Y") {
				iCount = iCount + 1;
			}
		}
		frmSelectionAccess.eventHidden.value = iCount;
	}
</script>

<script FOR=ssOleDBGridSortOrder EVENT=beforerowcolchange LANGUAGE=JavaScript>
	//	if (frmUseful.txtLockGridEvents.value != 1) 
	//		{
	//		if (frmUseful.txtAction.value.toUpperCase() == "VIEW") 
	//			{
	//			frmDefinition.ssOleDBGridSortOrder.columns(1).cellstyleset("ssetViewDormant", frmDefinition.ssOleDBGridSortOrder.row);
	//			}
	//		else 
	//			{
	//			frmDefinition.ssOleDBGridSortOrder.columns(1).cellstyleset("ssetDormant", frmDefinition.ssOleDBGridSortOrder.row);
	//			}
	//		}
	</script>

<script FOR=ssOleDBGridSortOrder EVENT=beforeupdate LANGUAGE=JavaScript>
<!--
	if ((frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Asc') &&
		(frmDefinition.ssOleDBGridSortOrder.Columns(2).text != 'Desc')) 
	{
		frmDefinition.ssOleDBGridSortOrder.Columns(2).text = 'Asc';
	}
	-->
</script>

<script FOR=ssOleDBGridSortOrder EVENT=afterinsert LANGUAGE=JavaScript>
<!--
	refreshTab4Controls();
	-->
</script>

<script FOR=ssOleDBGridSortOrder EVENT=rowcolchange LANGUAGE=JavaScript>
<!--
	frmDefinition.ssOleDBGridSortOrder.SelBookmarks.Add(frmDefinition.ssOleDBGridSortOrder.Bookmark);
	frmDefinition.ssOleDBGridSortOrder.columns(1).cellstyleset("ssetSelected", frmDefinition.ssOleDBGridSortOrder.row);
	frmSortOrder.txtSortColumnID.value = frmDefinition.ssOleDBGridSortOrder.Columns(0).text;
	frmSortOrder.txtSortColumnName.value = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
	frmSortOrder.txtSortOrder.value = frmDefinition.ssOleDBGridSortOrder.Columns(2).text;
	
	refreshTab4Controls();
	-->
</script>

<script FOR=ssOleDBGridSortOrder EVENT=change LANGUAGE=JavaScript>
<!--
	frmUseful.txtChanged.value = 1;
	-->
</script>

<script for=grdEvents event=change language=JavaScript>
<!--
	frmUseful.txtChanged.value = 1;
	-->
</script>

<script for=grdEvents event=click language=JavaScript>
<!--
	refreshTab2Controls();
	-->
</script>

<script for=grdEvents event=beforerowcolchange language=JavaScript>
<!--
	//	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") 
	//		{
	//		frmDefinition.grdEvents.columns(1).cellstyleset("ssetViewDormant", frmDefinition.grdEvents.row);
	//		}
	//	else 
	//		{
	//		frmDefinition.grdEvents.columns(1).cellstyleset("ssetDormant", frmDefinition.grdEvents.row);
	//		}
	-->
</script>

<script FOR=grdEvents EVENT=beforeupdate LANGUAGE=JavaScript>
<!--
	-->
</script>

<script FOR=grdEvents EVENT=afterinsert LANGUAGE=JavaScript>
<!--
	//refreshTab2Controls();
	-->
</script>

<script FOR=grdEvents EVENT=DblClick LANGUAGE=JavaScript>
<!--
	if (frmUseful.txtAction.value.toUpperCase() != "VIEW")
	{
		if (frmDefinition.grdEvents.Rows > 0 
			&& frmDefinition.grdEvents.SelBookmarks.Count == 1)
		{
			eventEdit(); 
		}
		else
		{
			eventAdd();
		}
	}
	-->
</script>

<script FOR=grdEvents EVENT=rowcolchange LANGUAGE=JavaScript>
	frmDefinition.grdEvents.SelBookmarks.Add(frmDefinition.grdEvents.Bookmark);
	frmDefinition.grdEvents.columns('Table').cellstyleset("ssetSelected", frmDefinition.grdEvents.row);
	refreshTab2Controls();
</script>

<script FOR=grdAccess EVENT=ComboCloseUp LANGUAGE=JavaScript>
	frmUseful.txtChanged.value = 1;
	if((grdAccess.AddItemRowIndex(grdAccess.Bookmark) == 0) &&
    (grdAccess.Columns("Access").Text.length > 0)) {
		ForceAccess(grdAccess, AccessCode(grdAccess.Columns("Access").Text));
    
		grdAccess.MoveFirst();
		grdAccess.Col = 1;
	}
	refreshTab1Controls();
</script>

<script FOR=grdAccess EVENT=GotFocus LANGUAGE=JavaScript>
	grdAccess.Col = 1
</script>

<script FOR=grdAccess EVENT=RowColChange(LastRow, LastCol) LANGUAGE=JavaScript>
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

	grdAccess.Col = 1
      
	-->
</script>

<script FOR=grdAccess EVENT=RowLoaded(Bookmark) LANGUAGE=JavaScript>
<!--
	var fViewing;
	var fIsNotOwner;
		
	fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
	fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

	if ((fIsNotOwner == true) ||
		(fViewing == true) ||
		(frmSelectionAccess.forcedHidden.value == "Y")) 
	{
		grdAccess.Columns("GroupName").CellStyleSet("ReadOnly");
		grdAccess.Columns("Access").CellStyleSet("ReadOnly");
		grdAccess.ForeColor = "-2147483631";
	}  
	else 
	{
		if (grdAccess.Columns("SysSecMgr").CellText(Bookmark) == "1") 
		{
			grdAccess.Columns("GroupName").CellStyleSet("SysSecMgr");
			grdAccess.Columns("Access").CellStyleSet("SysSecMgr");
			grdAccess.ForeColor = "0";
		}
		else 
		{
			grdAccess.ForeColor = "0";
		}
	}
	-->
</script>

<OBJECT classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB" 
	id=dialog 
  codebase="cabs/comdlg32.cab#Version=1,0,0,0"
	style="LEFT: 0px; TOP: 0px" 
	VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="847">
	<PARAM NAME="_ExtentY" VALUE="847">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="CancelError" VALUE="0">
	<PARAM NAME="Color" VALUE="0">
	<PARAM NAME="Copies" VALUE="1">
	<PARAM NAME="DefaultExt" VALUE="">
	<PARAM NAME="DialogTitle" VALUE="">
	<PARAM NAME="FileName" VALUE="">
	<PARAM NAME="Filter" VALUE="">
	<PARAM NAME="FilterIndex" VALUE="0">
	<PARAM NAME="Flags" VALUE="0">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="FontName" VALUE="">
	<PARAM NAME="FontSize" VALUE="8">
	<PARAM NAME="FontStrikeThru" VALUE="0">
	<PARAM NAME="FontUnderLine" VALUE="0">
	<PARAM NAME="FromPage" VALUE="0">
	<PARAM NAME="HelpCommand" VALUE="0">
	<PARAM NAME="HelpContext" VALUE="0">
	<PARAM NAME="HelpFile" VALUE="">
	<PARAM NAME="HelpKey" VALUE="">
	<PARAM NAME="InitDir" VALUE="">
	<PARAM NAME="Max" VALUE="0">
	<PARAM NAME="Min" VALUE="0">
	<PARAM NAME="MaxFileSize" VALUE="260">
	<PARAM NAME="PrinterDefault" VALUE="1">
	<PARAM NAME="ToPage" VALUE="0">
	<PARAM NAME="Orientation" VALUE="1"></OBJECT>

<div <%=session("BodyTag")%>>
<form id=frmDefinition name=frmDefinition>

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
				<tr height=5> 
					<td colspan=3></td>
				</tr> 

				<tr height=10>
					<TD width=10></td>
					<td>
						<INPUT type="button" value="Definition" id=btnTab1 name=btnTab1 disabled="disabled" class="btn btndisabled"
						    onclick="displayPage(1)" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
						<INPUT type="button" value="Event Details" id=btnTab2 name=btnTab2  class="btn"
						    onclick="displayPage(2)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
						<INPUT type="button" value="Report Details" id=btnTab3 name=btnTab3  class="btn"
						    onclick="displayPage(3)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
						<INPUT type="button" value="Sort Order" id=btnTab4 name=btnTab4  class="btn"
						    onclick="displayPage(4)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
						<INPUT type="button" value="Output" id=btnTab5 name=btnTab5  class="btn"
						    onclick="displayPage(5)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</td>
					<TD width=10></td>
				</tr> 
				
				<tr height=10> 
					<td colspan=3></td>
				</tr> 

				<tr> 
					<TD width=10></td>
					<td>
						<!-- First tab -->
						<DIV id=div1>
							<TABLE WIDTH="100%" height="80%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD colspan=10 height=5></TD>
											</TR>

											<TR height=10>
												<TD width=5></TD>
												<TD width=10>Name :</TD>
												<TD width=5>&nbsp;</TD>
												<TD>
													<INPUT id=txtName name=txtName maxlength="50" style="WIDTH: 100%" class="text"
													    onkeyup="changeTab1Control()">
												</TD>
												<TD width=20></TD>
												<TD width=10>Owner :</TD>
												<TD width=5>&nbsp;</TD>
												<TD>
													<INPUT id=txtOwner name=txtOwner style="WIDTH: 100%" class="text textdisabled"
													    disabled="disabled">
												</TD>
												<TD width=5></TD>
											</TR>
											
											<TR>
												<TD colspan=9 height=5></TD>
											</TR>
											
											<TR height=60>
												<TD width=5></TD>
												<TD width=10 nowrap valign=top>Description :</TD>
												<TD width=5></TD>
												<TD width="40%" rowspan="1" colspan=1>
													<TEXTAREA id=txtDescription name=txtDescription class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap=VIRTUAL height="0" maxlength="255"  
													    onkeyup="changeTab1Control()" 
													    onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}" 
													    onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
													</TEXTAREA>
												</TD>
												<TD width=20 nowrap></TD>
												<TD width=10 valign=top nowrap>Access :</TD>
												<TD width=5></TD>
												<TD width="40%" rowspan="1" valign=top height=78>
<!--#include file="include\ctl_grdAccess.txt"-->                  
												</TD>
												<TD width=5></TD>
											</TR>
											
											<TR height=20>
												<TD width=5>&nbsp;</TD>
												<TD colspan=7><hr></TD>
												<TD width=5>&nbsp;</TD>
											</TR>
											
											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=85 nowrap vAlign=top>Base Table :</TD>
												<TD width=5></TD>
												<TD width=40% vAlign=top colspan=1>
													<select id=cboBaseTable name=cboBaseTable style="WIDTH: 100%" class="combo combodisabled"
													    onchange="changeBaseTable()" disabled="disabled"> 
													</select>
												</TD>
												<TD width=20 nowrap></TD>
												<TD width=10 vAlign=top>Records :</TD>												
												<TD width="40%" colspan=2> 
													<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
														<TR>
															<TD width=5 nowrap>
																<input CHECKED id=optRecordSelection1 name=optRecordSelection type=radio 
																    onclick="changeBaseTableRecordOptions()"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD></TD>
															<TD colspan=4>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optRecordSelection1"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                        />
																    All
																</label>
															</TD>
															<TD colspan=3>&nbsp;</TD>
														</TR>
														<TR>
															<TD colspan=6 height=5></TD>
														</TR>						
														<TR>
															<TD width=5 nowrap>
																<input id=optRecordSelection2 name=optRecordSelection type=radio 
																    onclick="changeBaseTableRecordOptions()"
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD></TD>
															<TD width=5>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optRecordSelection2"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                        />
    																Picklist
                                        	    		        </label>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD Width=100%>
																<INPUT id=txtBasePicklist name=txtBasePicklist disabled="disabled" class="text textdisabled" style="WIDTH: 100%"> 
															</TD>
															<TD>
																<INPUT id=cmdBasePicklist name=cmdBasePicklist style="WIDTH: 30px" type=button disabled="disabled" class="btn btndisabled" value="..." 
																    onclick="selectRecordOption('base', 'picklist')"
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
														<TR>
															<TD colspan=6 height=5></TD>
														</TR>						
														<TR>
															<TD width=5 nowrap>
																<input id=optRecordSelection3 name=optRecordSelection type=radio
																    onclick=changeBaseTableRecordOptions() 
		                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD></TD>
															<TD width=5>
                                                                <label 
                                                                    tabindex="-1"
	                                                                for="optRecordSelection3"
	                                                                class="radio"
		                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                        />
    																Filter
                                        	    		        </label>
															</TD>
															<TD width=5></TD>
															<TD Width=100%>
																<INPUT id=txtBaseFilter name=txtBaseFilter class="text textdisabled" disabled="disabled" style="WIDTH: 100%"> 
															</TD>
															<TD>
																<INPUT id=cmdBaseFilter name=cmdBaseFilter style="WIDTH: 30px" type=button disabled="disabled" value="..." class="btn btndisabled"
																    onclick="selectRecordOption('base', 'filter')" 
                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
														<TR>
															<TD colspan=6 >
																<input name=chkPrintFilterHeader id=chkPrintFilterHeader type=checkbox disabled="disabled" tabindex="-1" 
	                                                                onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
	                                                                onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" 
																    onClick="changeTab1Control();">
                                                                <label 
                                                                    id="lblPrintFilterHeader"
                                                                    name="lblPrintFilterHeader"
				                                                    for="chkPrintFilterHeader"
				                                                    class="checkbox checkboxdisabled"
				                                                    tabindex=0 
				                                                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">

    																Display filter or picklist title in the report header 
                                    		    		        </label>
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
											<TR>
												<TD colspan=9 height=5></TD>
											</TR>
											
											<TR height=5>
												<TD width=5></TD>
												<TD nowrap vAlign=top>Description 1 :</TD>
												<TD width=5></TD>
												<TD vAlign=top colspan=1>
													<select id=cboDescription1 name=cboDescription1 style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
													    onchange="changeTab1Control();"> 
													</select>
												</TD>
												<TD width=20 nowrap></TD>
												<TD width=10 vAlign=top>Region :</TD>
												<TD width=5></TD>
												<TD> 
													<select id=cboRegion name=cboRegion style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
													    onchange="changeTab1Control();refreshTab3Controls();"> 
													</select>
												</TD>
												<TD width=5></TD>
											</TR>
											<TR>
												<TD colspan=9 height=3></TD>
											</TR>
											
											<TR height=10>
												<TD width=5></TD>
												<TD nowrap vAlign=top>Description 2 :</TD>
												<TD width=5>&nbsp;</TD>
												<TD vAlign=top colspan=1>
													<select id=cboDescription2 name=cboDescription2 style="WIDTH: 100%" disabled="disabled" class="combo combodisabled"
													    onchange="changeTab1Control();" > 
													</select>
												</TD>
												<TD width=20 nowrap></TD>
												<TD width=10 vAlign=top colspan=3></TD>
												<TD width=5></TD>
											</TR>
											<TR>
												<TD colspan=9 height=3></TD>
											</TR>
											<TR height=10>
												<TD width=5></TD>
												<TD nowrap vAlign=top>Description 3 :</TD>
												<TD width=5 nowrap>&nbsp;</TD>
												<TD>
													<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
														<TR>
															<TD WIDTH=100%>
																<INPUT id=txtDescExpr name=txtDescExpr disabled="disabled" class="text textdisabled" style="WIDTH: 100%"> 
															</TD>
															<TD width=30 nowrap>
																<INPUT id=cmdDescExpr name=cmdDescExpr style="WIDTH: 30px" type=button disabled="disabled" class="btn btndisabled" value="..."
																    onclick="selectCalc('baseDesc', false)"
																    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=5></TD>
												<TD width=10 vAlign=top colspan=3></TD>
												<TD width=5></TD>											
											</TR>
											<TR>
												<TD colspan=9 height=3></TD>
											</TR>
											
											<TR height=10>
												<TD colspan=4 align=left>
													<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD width=5>&nbsp;</TD>
															<TD align=left width=15 valign=center id="qq">
																<input valign=center name=chkGroupByDesc id=chkGroupByDesc type=checkbox disabled="disabled" tabindex="-1" 
	                                                                onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
	                                                                onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" 
																    onClick="changeTab1Control();refreshTab3Controls();">
															</TD>
															<TD align=left width=150 valign=center>
                                                                <label 
				                                                    for="chkGroupByDesc"
				                                                    class="checkbox checkboxdisabled"
				                                                    tabindex=0 
				                                                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
    																Group By Description
                                    		    		        </label>
															</TD>
															<TD>&nbsp;</TD>
															<TD width=20 valign=center>Separator : </TD>
															<TD width=5>&nbsp;</TD>
															<TD width=85 valign=center>
																<SELECT name=cboDescriptionSeparator id=cboDescriptionSeparator style="WIDTH: 100%" disabled="disabled" class="combo combodisabled"
																    onchange="changeTab1Control();">
																	<OPTION value="">&lt;None&gt;
																	<OPTION value=" ">&lt;Space&gt;
																	<OPTION value=", ">, 
																	<OPTION value=".  ">.  
																	<OPTION value=" - "> - 
																	<OPTION value=" : "> : 
																	<OPTION value=" ; "> ; 
																	<OPTION value=" / "> / 
																	<OPTION value=" \ "> \ 
																	<OPTION value=" # "> # 
																	<OPTION value=" ~ "> ~ 
																	<OPTION value=" ^ "> ^       
																</SELECT>
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>		
											<TR>
												<TD colspan=9 height=5></TD>
											</TR>								
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>

						<!-- Second tab -->
						<DIV id=div2 style="visibility:hidden;display:none">
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD colspan=3 height=5></TD>
											</TR>
																						
											<TR height=10>
												<TD width=5>&nbsp;</TD>
												<TD width=90 nowrap colspan=7>Events :</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
											<TR>
												<TD width=5>&nbsp;</TD>
												<TD colspan=7>
													<TABLE WIDTH="100%"  height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD colspan=3 height=5></TD>
														</TR>
														<TR height=5>
															<TD rowspan=8>

<%

	Dim avColumnDef(26, 4)
	
	avColumnDef(0, 0) = "Name"			 'name
	avColumnDef(0, 1) = "Name"			 'caption
	avColumnDef(0, 2) = "2000"			 'width
	avColumnDef(0, 3) = "-1"				 'visible
	
	avColumnDef(1, 0) = "TableID"			 'name
	avColumnDef(1, 1) = "TableID"			 'caption
	avColumnDef(1, 2) = "1814"			 'width
	avColumnDef(1, 3) = "0"					 'visible

	avColumnDef(2, 0) = "Table"				 'name
	avColumnDef(2, 1) = "Table"				 'caption
	avColumnDef(2, 2) = "2000"			 'width
	avColumnDef(2, 3) = "-1"				 'visible

	avColumnDef(3, 0) = "FilterID"		 'name
	avColumnDef(3, 1) = "FilterID"		 'caption
	avColumnDef(3, 2) = "1814"			 'width
	avColumnDef(3, 3) = "0"					 'visible

	avColumnDef(4, 0) = "Filter"			 'name
	avColumnDef(4, 1) = "Filter"			 'caption
	avColumnDef(4, 2) = "2000"			 'width
	avColumnDef(4, 3) = "-1"				 'visible

	avColumnDef(5, 0) = "StartDateID"		 'name
	avColumnDef(5, 1) = "StartDateID"		 'caption
	avColumnDef(5, 2) = "1814"			 'width
	avColumnDef(5, 3) = "0"					 'visible

	avColumnDef(6, 0) = "Start Date"		 'name
	avColumnDef(6, 1) = "Start Date"		 'caption
	avColumnDef(6, 2) = "2100"			 'width
	avColumnDef(6, 3) = "-1"				 'visible

	avColumnDef(7, 0) = "StartSessionID"	 'name
	avColumnDef(7, 1) = "StartSessionID"	 'caption
	avColumnDef(7, 2) = "1814"			 'width
	avColumnDef(7, 3) = "0"					 'visible
	
	avColumnDef(8, 0) = "Start Session"		 'name
	avColumnDef(8, 1) = "Start Session"		 'caption
	avColumnDef(8, 2) = "2600"			 'width
	avColumnDef(8, 3) = "-1"				 'visible
	
	avColumnDef(9, 0) = "EndDateID"			 'name
	avColumnDef(9, 1) = "EndDateID"			 'caption
	avColumnDef(9, 2) = "1814"			 'width
	avColumnDef(9, 3) = "0"					 'visible
	
	avColumnDef(10, 0) = "End Date"			 'name
	avColumnDef(10, 1) = "End Date"			 'caption
	avColumnDef(10, 2) = "2100"				 'width
	avColumnDef(10, 3) = "-1"				 'visible
	
	avColumnDef(11, 0) = "EndSessionID"		 'name
	avColumnDef(11, 1) = "EndSessionID"		 'caption
	avColumnDef(11, 2) = "1814"				 'width
	avColumnDef(11, 3) = "0"				 'visible
	
	avColumnDef(12, 0) = "End Session"	 'name
	avColumnDef(12, 1) = "End Session"	 'caption
	avColumnDef(12, 2) = "2600"				 'width
	avColumnDef(12, 3) = "-1"				 'visible
	
	avColumnDef(13, 0) = "DurationID"		 'name
	avColumnDef(13, 1) = "DurationID"		 'caption
	avColumnDef(13, 2) = "1814"				 'width
	avColumnDef(13, 3) = "0"				 'visible
	
	avColumnDef(14, 0) = "Duration"			 'name
	avColumnDef(14, 1) = "Duration"			 'caption
	avColumnDef(14, 2) = "2000"				 'width
	avColumnDef(14, 3) = "-1"				 'visible
	
	avColumnDef(15, 0) = "LegendType"		 'name
	avColumnDef(15, 1) = "LegendType"		 'caption
	avColumnDef(15, 2) = "1814"				 'width
	avColumnDef(15, 3) = "0"				 'visible
	
	avColumnDef(16, 0) = "Legend"			 'name
	avColumnDef(16, 1) = "Key"			 'caption
	avColumnDef(16, 2) = "2000"				 'width
	avColumnDef(16, 3) = "-1"				 'visible
	
	avColumnDef(17, 0) = "LegendTableID"	 'name
	avColumnDef(17, 1) = "LegendTableID"	 'caption
	avColumnDef(17, 2) = "1814"				 'width
	avColumnDef(17, 3) = "0"				 'visible
	
	avColumnDef(18, 0) = "LegendColumnID"	 'name
	avColumnDef(18, 1) = "LegendColumnID"	 'caption
	avColumnDef(18, 2) = "1814"				 'width
	avColumnDef(18, 3) = "0"				 'visible
	
	avColumnDef(19, 0) = "LegendCodeID"		 'name
	avColumnDef(19, 1) = "LegendCodeID"		 'caption
	avColumnDef(19, 2) = "1814"				 'width
	avColumnDef(19, 3) = "0"				 'visible
	
	avColumnDef(20, 0) = "LegendEventTypeID" 'name
	avColumnDef(20, 1) = "LegendEventTypeID" 'caption
	avColumnDef(20, 2) = "1814"				 'width
	avColumnDef(20, 3) = "0"				 'visible
	
	avColumnDef(21, 0) = "Desc1ID"		 'name
	avColumnDef(21, 1) = "Desc1ID"		 'caption
	avColumnDef(21, 2) = "1814"				 'width
	avColumnDef(21, 3) = "0"				 'visible
	
	avColumnDef(22, 0) = "Description 1"	 'name
	avColumnDef(22, 1) = "Description 1"	 'caption
	avColumnDef(22, 2) = "2600"				 'width
	avColumnDef(22, 3) = "-1"				 'visible
	
	avColumnDef(23, 0) = "Desc2ID"		 'name
	avColumnDef(23, 1) = "Desc2ID"		 'caption
	avColumnDef(23, 2) = "1814"				 'width
	avColumnDef(23, 3) = "0"				 'visible
	
	avColumnDef(24, 0) = "Description 2"	 'name
	avColumnDef(24, 1) = "Description 2"	 'caption
	avColumnDef(24, 2) = "2600"				 'width
	avColumnDef(24, 3) = "-1"				 'visible
	
	avColumnDef(25, 0) = "EventKey"			 'name
	avColumnDef(25, 1) = "EventKey"			 'caption
	avColumnDef(25, 2) = "1395"				 'width
	avColumnDef(25, 3) = "0"				 'visible
	
	avColumnDef(26, 0) = "FilterHidden"			 'name
	avColumnDef(26, 1) = "FilterHidden"			 'caption
	avColumnDef(26, 2) = "1395"				 'width
	avColumnDef(26, 3) = "0"				 'visible
	
			
	Response.Write("											<OBJECT classid=clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" & vbCrLf)
	Response.Write("													 codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6""" & vbCrLf)
	Response.Write("													height=""100%""" & vbCrLf)
	Response.Write("													id=grdEvents" & vbCrLf)
	Response.Write("													name=grdEvents" & vbCrLf)
	Response.Write("													style=""HEIGHT: 100%; VISIBILITY: visible; WIDTH: 100%""" & vbCrLf)
	Response.Write("													width=""100%"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""GroupHeaders"" VALUE=""-1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
	Response.Write("												<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""SelectTypeRow"" VALUE=""3"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""RowNavigation"" VALUE=""2"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
	Response.Write("												<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
	Response.Write("												<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
	Response.Write("												<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Columns.Count"" VALUE=""" & (UBound(avColumnDef) + 1) & """>" & vbCrLf)
	
	For i = 0 To UBound(avColumnDef) Step 1
		Response.Write("												<!--" & avColumnDef(i, 0) & "-->  " & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Width"" VALUE=""" & avColumnDef(i, 2) & """>" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Visible"" VALUE=""" & avColumnDef(i, 3) & """>" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Caption"" VALUE=""" & avColumnDef(i, 1) & """>" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Name"" VALUE=""" & avColumnDef(i, 0) & """>" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Alignment"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Bound"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").DataField"" VALUE=""Column " & i & """>" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").DataType"" VALUE=""8"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Level"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").NumberFormat"" VALUE="""">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Case"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").FieldLen"" VALUE=""256"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Locked"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Style"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").RowCount"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").ColCount"" VALUE=""1"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").ForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").BackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").StyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Nullable"" VALUE=""1"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").Mask"" VALUE="""">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").ClipMode"" VALUE=""0"">" & vbCrLf)
		Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptChar"" VALUE=""95"">" & vbCrLf)
	Next
		
	Response.Write("												<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""_ExtentX"" VALUE=""11298"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""_ExtentY"" VALUE=""3969"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
	Response.Write("												<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""BackColor"" VALUE=""-2147483633"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
	Response.Write("												<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

	Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
	Response.Write("											</OBJECT>" & vbCrLf)
	
%>											

															</TD>

															<TD width=10>&nbsp;</TD>
															<TD width=80>
																<input type="button" id=cmdAddEvent name=cmdAddEvent value="Add..." style="WIDTH: 100%" class="btn"
																    onclick="eventAdd()"
		                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                                            onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<TD width=5>&nbsp;</TD>
														</TR>

														<TR height=5>
															<TD colspan=3></TD>
														</TR>
																	
														<TR height=5>
															
															<TD width=5>&nbsp;</TD>
															<TD width=80>
																<input type="button" id=cmdEditEvent name=cmdChildEvent value="Edit..." style="WIDTH: 100%" class="btn"
																    onclick="eventEdit()"
		                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                                            onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<TD width=5>&nbsp;</TD>
														</TR>
																	
														<TR height=5>
															<TD colspan=3></TD>
														</TR>

														<TR height=5>
															
															<TD width=5>&nbsp;</TD>
															<TD width=80>
																<input type="button" id=cmdRemoveEvent name=cmdRemoveEvent value="Remove" style="WIDTH: 100%" class="btn"
																    onclick="eventRemove()"
		                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                                            onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<TD width=5>&nbsp;</TD>
														</TR>
																	
														<TR height=5>
															<TD colspan=3></TD>
														</TR>
																	
														<TR height=5>
															
															<TD width=5>&nbsp;</TD>
															<TD width=80>
																<input type="button" id=cmdRemoveAllEvents name=cmdRemoveAllEvents value="Remove All" style="WIDTH: 100%" class="btn"
																    onclick="eventRemoveAll()"
		                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                                            onblur="try{button_onBlur(this);}catch(e){}" />
															</TD>
															<TD width=5>&nbsp;</TD>
														</TR>
														
														<TR>
															<TD colspan=3></TD>
														</TR>
														
													</TABLE>
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
<!---------------------------------------------------------------------------------------------->

											<TR>
												<TD colspan=9 height=5></TD>
											</TR>
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>

						<!-- Third tab -->
						<DIV id=div3 style="visibility:hidden;display:none">
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=10 CELLPADDING=0>
											<tr>						
												<td valign=top rowspan=1 width=50% height=100%>
													<table class="outline" cellspacing="0" cellpadding="4" width=100% height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																Start Date : <BR><BR>
																<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width=100%>
																	<TR height=20>	
																		<td width=5>&nbsp</td>
																		<TD align=left width=15>
																			<INPUT type=radio name=optStart id=optFixedStart 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</TD>
																		<td width=10>&nbsp</td>
																		<TD align=left width=15 nowrap>
                                                                            <label 
                                                                                tabindex="-1"
	                                                                            for="optFixedStart"
	                                                                            class="radio"
		                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                    />
	        																	Fixed
                                        	    		                    </label>
    																	</TD>
																		<td width=30 nowrap>&nbsp</td>
																		<TD align=left width=100%>
																			<input type=text id=txtFixedStart name=txtFixedStart value="" style="WIDTH: 100%" class="text"
																			    onkeyup="changeTab3Control()">
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=5>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR height=20>	
																		<td width=5>&nbsp</td>
																		<TD align=left width=15>
																			<INPUT type=radio name=optStart id=optCurrentStart 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</TD>
																		<td width=10>&nbsp</td>
																		<TD colspan=3 align=left nowrap>
                                                                            <label 
                                                                                tabindex="-1"
	                                                                            for="optCurrentStart"
	                                                                            class="radio"
		                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                    />
    																			Current Date
                                        	    		                    </label>
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=5>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR height=20>	
																		<td width=5>&nbsp</td>
																		<TD align=left width=15 nowrap>
																			<INPUT type=radio name=optStart id=optOffsetStart 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</TD>
																		<td width=10>&nbsp</td>
																		<TD align=left width=15 nowrap>
                                                                            <label 
                                                                                tabindex="-1"
	                                                                            for="optOffsetStart"
	                                                                            class="radio"
		                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                    />
    																			Offset
                                        	    		                    </label>
																		</TD>
																		<td width=30 nowrap>&nbsp</td>
																		<TD align=left width=100%>
																			<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD width=40>
																						<INPUT id=txtFreqStart maxlength=4 name=txtFreqStart style="WIDTH: 40px" width="40" value="0" class="text"
																						    onkeyup="setRecordsNumeric(frmDefinition.txtFreqStart);changeTab3Control();validateOffsets();" 
																						    onchange="setRecordsNumeric(frmDefinition.txtFreqStart);changeTab3Control();validateOffsets();">
																					</TD>
																					<TD>
																						<input style="WIDTH: 15px" type="button" value="+" id="cmdPeriodStartUp" name="cmdPeriodStartUp" class="btn"
																						    onclick="spinRecords(true,frmDefinition.txtFreqStart);setRecordsNumeric(frmDefinition.txtFreqStart);changeTab3Control();validateOffsets();"
                                                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{button_onFocus(this);}catch(e){}"
                                                                                            onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																					<TD>
																						<input style="WIDTH: 15px" type="button" value="-" id="cmdPeriodStartDown" name="cmdPeriodStartDown" class="btn"
																						    onclick="spinRecords(false,frmDefinition.txtFreqStart);setRecordsNumeric(frmDefinition.txtFreqStart);changeTab3Control();validateOffsets();"
                                                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{button_onFocus(this);}catch(e){}"
                                                                                            onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																					<TD width=10>&nbsp;</TD>
																					<TD width=100%>
																						<SELECT name=cboPeriodStart id=cboPeriodStart style="WIDTH: 100%" width=100% class="combo"
																						    onChange="changeTab3Control();validateOffsets();">
																							<OPTION name=Day value=0 selected>Day(s)
																							<OPTION name=Week value=1>Week(s)
																							<OPTION name=Month value=2>Month(s)
																							<OPTION name=Year value=3>Year(s)
																						</SELECT>
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=5>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR height=20>	
																		<td width=5>&nbsp</td>
																		<TD align=left width=15 nowrap>
																			<INPUT type=radio name=optStart id=optCustomStart 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</TD>
																		<td width=10>&nbsp</td>
																		<TD align=left width=15 nowrap>
                                                                            <label 
                                                                                tabindex="-1"
	                                                                            for="optCustomStart"
	                                                                            class="radio"
		                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                    />
    																			Custom
                                        	    		                    </label>
																		</TD>
																		<td width=30 nowrap>&nbsp</td>
																		<TD align=left width=100%>
																			<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD>
																						<INPUT id=txtCustomStart name=txtCustomStart disabled="disabled" style="WIDTH: 100%" class="text textdisabled"> 
																					</TD>
																					<TD width=30>
																						<INPUT id=cmdCustomStart name=cmdCustomStart style="WIDTH: 100%" type=button disabled="disabled" value="..." class="btn btndisabled"
																						    onclick="selectCalc('startDate', true)"
                                                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{button_onFocus(this);}catch(e){}"
                                                                                            onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=5>	
																		<TD colspan=7></TD>
																	</TR>
																</TABLE>
															</td>
														</tr>
													</table>
												</td>
												<td valign=top width=50%>
													<table class="outline" cellspacing="0" cellpadding="2" width=100%  height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																End Date : <BR><BR>
																<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width=100%>
																	<TR height=20>	
																		<td width=5>&nbsp</td>
																		<TD align=left width=15>
																			<INPUT type=radio name=optEnd id=optFixedEnd 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</TD>
																		<td width=10>&nbsp</td>
																		<TD align=left width=15 nowrap>
                                                                            <label 
                                                                                tabindex="-1"
	                                                                            for="optFixedEnd"
	                                                                            class="radio"
		                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                    />
    																			Fixed
                                        	    		                    </label>
																		</TD>
																		<td width=30 nowrap>&nbsp</td>
																		<TD align=left width=100%>
																			<input type=text id=txtFixedEnd name=txtFixedEnd value="" style="WIDTH: 100%" class="text"
																			    onkeyup="changeTab3Control()">
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=5>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR height=20>	
																		<td width=5>&nbsp</td>
																		<TD align=left width=15>
																			<INPUT type=radio name=optEnd id=optCurrentEnd 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</TD>
																		<td width=10>&nbsp</td>
																		<TD colspan=3 align=left nowrap>
                                                                            <label 
                                                                                tabindex="-1"
	                                                                            for="optCurrentEnd"
	                                                                            class="radio"
		                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                    />
    																			Current Date
                                        	    		                    </label>
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=5>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR height=20>	
																		<td width=5>&nbsp</td>
																		<TD align=left width=15 nowrap>
																			<INPUT type=radio name=optEnd id=optOffsetEnd 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</TD>
																		<td width=10>&nbsp</td>
																		<TD align=left width=15 nowrap>
                                                                            <label 
                                                                                tabindex="-1"
	                                                                            for="optOffsetEnd"
	                                                                            class="radio"
		                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                    />
																			    Offset
                                        	    		                    </label>
																		</TD>
																		<td width=30 nowrap>&nbsp</td>
																		<TD align=left width=100%>
																			<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD width=40>
																						<INPUT id=txtFreqEnd maxlength=4 name=txtFreqEnd style="WIDTH: 40px" width="40" value="0" class="text"
																						    onkeyup="setRecordsNumeric(frmDefinition.txtFreqEnd);changeTab3Control();validateOffsets();" 
																						    onchange="setRecordsNumeric(frmDefinition.txtFreqEnd);changeTab3Control();validateOffsets();">
																					</TD>
																					<TD>
																						<input style="WIDTH: 15px" type="button" value="+" id="cmdPeriodEndUp" name="cmdPeriodEndUp" class="btn"
																						    onclick="spinRecords(true,frmDefinition.txtFreqEnd);setRecordsNumeric(frmDefinition.txtFreqEnd);changeTab3Control();validateOffsets();"
                                                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{button_onFocus(this);}catch(e){}"
                                                                                            onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																					<TD>
																						<input style="WIDTH: 15px" type="button" value="-" id="cmdPeriodEndDown" name="cmdPeriodEndDown" class="btn"
																						    onclick="spinRecords(false,frmDefinition.txtFreqEnd);setRecordsNumeric(frmDefinition.txtFreqEnd);changeTab3Control();validateOffsets();"
                                                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{button_onFocus(this);}catch(e){}"
                                                                                            onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																					<TD width=10>&nbsp;</TD>
																					<TD width=100%>
																						<SELECT name=cboPeriodEnd id=cboPeriodEnd style="WIDTH: 100%" width=100% class="combo"
																						    onChange="changeTab3Control();validateOffsets();">
																							<OPTION name=Day value=0 selected>Day(s)
																							<OPTION name=Week value=1>Week(s)
																							<OPTION name=Month value=2>Month(s)
																							<OPTION name=Year value=3>Year(s)
																						</SELECT>
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=5>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR height=20>	
																		<td width=5>&nbsp</td>
																		<TD align=left width=15 nowrap>
																			<INPUT type=radio name=optEnd id=optCustomEnd 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</TD>
																		<td width=10>&nbsp</td>
																		<TD align=left width=15 nowrap>
                                                                            <label 
                                                                                tabindex="-1"
	                                                                            for="optCustomEnd"
	                                                                            class="radio"
		                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                    />
    																			Custom
                                        	    		                    </label>
																		</TD>
																		<td width=30 nowrap>&nbsp</td>
																		<TD align=left width=100%>
																			<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD>
																						<INPUT id=txtCustomEnd name=txtCustomEnd disabled="disabled" class="text textdisabled" style="WIDTH: 100%"> 
																					</TD>
																					<TD width=30>
																						<INPUT id=cmdCustomEnd name=cmdCustomEnd style="WIDTH: 100%" type=button disabled="disabled" value='...' class="btn btndisabled"
																						    onclick="selectCalc('endDate', true)"
                                                                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{button_onFocus(this);}catch(e){}"
                                                                                            onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=5>	
																		<TD colspan=7></TD>
																	</TR>
																</TABLE>
															</td>
														</tr>
													</table>
												</td>
											</tr>
											<tr>						
												<td valign=top width=100% colspan=2>
													<table class="outline" cellspacing="0" cellpadding="2" width=100% height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top width=100%>
																Default Display Options : <BR><BR>
																<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
																	<TR>
																		<td width=5>&nbsp</td>
																		<TD>
																			<input name=chkIncludeBHols id=chkIncludeBHols type=checkbox disabled="disabled" tabindex=-1 
																			    onClick="changeTab3Control();" 
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkIncludeBHols"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																			    Include Bank Holidays 
																			</label> 
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=2>	
																		<TD colspan=3></TD>
																	</TR>
																	<TR>
																		<td width=5>&nbsp</td>
																		<TD>
																			<input name=chkIncludeWorkingDaysOnly id=chkIncludeWorkingDaysOnly type=checkbox disabled="disabled" tabindex=-1  
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkIncludeWorkingDaysOnly"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
	    																		Working Days Only 
    																		</label> 
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=2>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR>
																		<td width=5>&nbsp</td>
																		<TD>
																			<input name=chkShadeBHols id=chkShadeBHols type=checkbox disabled="disabled" tabindex=-1 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkShadeBHols"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
    																			Show Bank Holidays 
    																		</label> 
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=2>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR>
																		<td width=5>&nbsp</td>
																		<TD>
																			<input name=chkCaptions id=chkCaptions type=checkbox disabled="disabled" tabindex=-1  
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkCaptions"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
    																			Show Calendar Captions
    																		</label> 
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=2>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR>
																		<td width=5>&nbsp</td>
																		<TD>
																			<input name=chkShadeWeekends id=chkShadeWeekends type=checkbox disabled="disabled" tabindex=-1  
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkShadeWeekends"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
    																			Show Weekends 
    																		</label> 
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=2>	
																		<TD colspan=7></TD>
																	</TR>
																	<TR>
																		<td width=5>&nbsp</td>
																		<TD>
																			<input name=chkStartOnCurrentMonth id=chkStartOnCurrentMonth type=checkbox disabled="disabled" tabindex=-1 
																			    onClick="changeTab3Control();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkStartOnCurrentMonth"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
    																			Start on Current Month 
    																		</label> 
																		</TD>
																		<td width=5>&nbsp</td>
																	</TR>
																	<TR height=2>	
																		<TD colspan=7></TD>
																	</TR>
																</TABLE>															
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>
					
						<!-- Fourth tab -->
						<DIV id=div4 style="visibility:hidden;display:none">
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD colspan=5 height=5></TD>
											</TR>
											
											<TR height=20>
												<TD width=5>&nbsp;</TD>
												<TD width=90 colspan=3 nowrap>Sort Order :</TD>
												<TD width=5>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD rowspan=12>
													<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
														  codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" 
														height="100%" 
														id=ssOleDBGridSortOrder 
														name=ssOleDBGridSortOrder 
														style="HEIGHT: 100%; VISIBILITY: visible; display: block; WIDTH: 100%" 
														width="100%">
												    <PARAM NAME="ScrollBars" VALUE="4">
												    <PARAM NAME="_Version" VALUE="196617">
												    <PARAM NAME="DataMode" VALUE="2">
												    <PARAM NAME="Cols" VALUE="0">
												    <PARAM NAME="Rows" VALUE="0">
												    <PARAM NAME="BorderStyle" VALUE="1">
												    <PARAM NAME="RecordSelectors" VALUE="0">
												    <PARAM NAME="GroupHeaders" VALUE="-1">
												    <PARAM NAME="ColumnHeaders" VALUE="-1">
												    <PARAM NAME="GroupHeadLines" VALUE="1">
												    <PARAM NAME="HeadLines" VALUE="1">
												    <PARAM NAME="FieldDelimiter" VALUE="(None)">
												    <PARAM NAME="FieldSeparator" VALUE="(Tab)">
												    <PARAM NAME="Row.Count" VALUE="0">
												    <PARAM NAME="Col.Count" VALUE="1">
												    <PARAM NAME="stylesets.count" VALUE="0">
												    <PARAM NAME="TagVariant" VALUE="EMPTY">
												    <PARAM NAME="UseGroups" VALUE="0">
												    <PARAM NAME="HeadFont3D" VALUE="0">
												    <PARAM NAME="Font3D" VALUE="0">
												    <PARAM NAME="DividerType" VALUE="3">
												    <PARAM NAME="DividerStyle" VALUE="1">
												    <PARAM NAME="DefColWidth" VALUE="0">
												    <PARAM NAME="BeveColorScheme" VALUE="2">
												    <PARAM NAME="BevelColorFrame" VALUE="-2147483642">
												    <PARAM NAME="BevelColorHighlight" VALUE="-2147483628">
												    <PARAM NAME="BevelColorShadow" VALUE="-2147483632">
												    <PARAM NAME="BevelColorFace" VALUE="-2147483633">
												    <PARAM NAME="CheckBox3D" VALUE="-1">
												    <PARAM NAME="AllowAddNew" VALUE="0">
												    <PARAM NAME="AllowDelete" VALUE="0">
												    <PARAM NAME="AllowUpdate" VALUE="-1">
												    <PARAM NAME="MultiLine" VALUE="0">
												    <PARAM NAME="ActiveCellStyleSet" VALUE="">
												    <PARAM NAME="RowSelectionStyle" VALUE="0">
												    <PARAM NAME="AllowRowSizing" VALUE="0">
												    <PARAM NAME="AllowGroupSizing" VALUE="0">
												    <PARAM NAME="AllowColumnSizing" VALUE="-1">
												    <PARAM NAME="AllowGroupMoving" VALUE="0">
												    <PARAM NAME="AllowColumnMoving" VALUE="0">
												    <PARAM NAME="AllowGroupSwapping" VALUE="0">
												    <PARAM NAME="AllowColumnSwapping" VALUE="0">
												    <PARAM NAME="AllowGroupShrinking" VALUE="0">
												    <PARAM NAME="AllowColumnShrinking" VALUE="0">
												    <PARAM NAME="AllowDragDrop" VALUE="0">
												    <PARAM NAME="UseExactRowCount" VALUE="-1">
												    <PARAM NAME="SelectTypeCol" VALUE="0">
												    <PARAM NAME="SelectTypeRow" VALUE="1">
												    <PARAM NAME="SelectByCell" VALUE="-1">
												    <PARAM NAME="BalloonHelp" VALUE="0">
												    <PARAM NAME="RowNavigation" VALUE="2">
												    <PARAM NAME="CellNavigation" VALUE="0">
												    <PARAM NAME="MaxSelectedRows" VALUE="1">
												    <PARAM NAME="HeadStyleSet" VALUE="">
												    <PARAM NAME="StyleSet" VALUE="">
												    <PARAM NAME="ForeColorEven" VALUE="0">
												    <PARAM NAME="ForeColorOdd" VALUE="0">
												    <PARAM NAME="BackColorEven" VALUE="16777215">
												    <PARAM NAME="BackColorOdd" VALUE="16777215">
												    <PARAM NAME="Levels" VALUE="1">
												    <PARAM NAME="RowHeight" VALUE="503">
												    <PARAM NAME="ExtraHeight" VALUE="0">
												    <PARAM NAME="ActiveRowStyleSet" VALUE="">
												    <PARAM NAME="CaptionAlignment" VALUE="2">
												    <PARAM NAME="SplitterPos" VALUE="0">
												    <PARAM NAME="SplitterVisible" VALUE="0">
												    <PARAM NAME="Columns.Count" VALUE="3">
													        
												    <PARAM NAME="Columns(0).Width" VALUE="3200">
												    <PARAM NAME="Columns(0).Visible" VALUE="0">
												    <PARAM NAME="Columns(0).Columns.Count" VALUE="1">
												    <PARAM NAME="Columns(0).Caption" VALUE="id">
												    <PARAM NAME="Columns(0).Name" VALUE="columnID">
												    <PARAM NAME="Columns(0).Alignment" VALUE="0">
												    <PARAM NAME="Columns(0).CaptionAlignment" VALUE="3">
												    <PARAM NAME="Columns(0).Bound" VALUE="0">
												    <PARAM NAME="Columns(0).AllowSizing" VALUE="1">
												    <PARAM NAME="Columns(0).DataField" VALUE="Column 0">
												    <PARAM NAME="Columns(0).DataType" VALUE="8">
												    <PARAM NAME="Columns(0).Level" VALUE="0">
												    <PARAM NAME="Columns(0).NumberFormat" VALUE="">
												    <PARAM NAME="Columns(0).Case" VALUE="0">
												    <PARAM NAME="Columns(0).FieldLen" VALUE="256">
												    <PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
												    <PARAM NAME="Columns(0).Locked" VALUE="0">
												    <PARAM NAME="Columns(0).Style" VALUE="0">
												    <PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
												    <PARAM NAME="Columns(0).RowCount" VALUE="0">
												    <PARAM NAME="Columns(0).ColCount" VALUE="1">
												    <PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
												    <PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
												    <PARAM NAME="Columns(0).HasForeColor" VALUE="0">
												    <PARAM NAME="Columns(0).HasBackColor" VALUE="0">
												    <PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
												    <PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
												    <PARAM NAME="Columns(0).ForeColor" VALUE="0">
												    <PARAM NAME="Columns(0).BackColor" VALUE="0">
												    <PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
												    <PARAM NAME="Columns(0).StyleSet" VALUE="">
												    <PARAM NAME="Columns(0).Nullable" VALUE="1">
												    <PARAM NAME="Columns(0).Mask" VALUE="">
												    <PARAM NAME="Columns(0).PromptInclude" VALUE="0">
												    <PARAM NAME="Columns(0).ClipMode" VALUE="0">
												    <PARAM NAME="Columns(0).PromptChar" VALUE="95">
													        
												    <PARAM NAME="Columns(1).Width" VALUE="8000">
												    <PARAM NAME="Columns(1).Visible" VALUE="-1">
												    <PARAM NAME="Columns(1).Columns.Count" VALUE="1">
												    <PARAM NAME="Columns(1).Caption" VALUE="Column">
												    <PARAM NAME="Columns(1).Name" VALUE="column">
												    <PARAM NAME="Columns(1).Alignment" VALUE="0">
												    <PARAM NAME="Columns(1).CaptionAlignment" VALUE="3">
												    <PARAM NAME="Columns(1).Bound" VALUE="0">
												    <PARAM NAME="Columns(1).AllowSizing" VALUE="1">
												    <PARAM NAME="Columns(1).DataField" VALUE="Column 1">
												    <PARAM NAME="Columns(1).DataType" VALUE="8">
												    <PARAM NAME="Columns(1).Level" VALUE="0">
												    <PARAM NAME="Columns(1).NumberFormat" VALUE="">
												    <PARAM NAME="Columns(1).Case" VALUE="0">
												    <PARAM NAME="Columns(1).FieldLen" VALUE="256">
												    <PARAM NAME="Columns(1).VertScrollBar" VALUE="0">
												    <PARAM NAME="Columns(1).Locked" VALUE="1">
												    <PARAM NAME="Columns(1).Style" VALUE="0">
												    <PARAM NAME="Columns(1).ButtonsAlways" VALUE="0">
												    <PARAM NAME="Columns(1).RowCount" VALUE="0">
												    <PARAM NAME="Columns(1).ColCount" VALUE="1">
												    <PARAM NAME="Columns(1).HasHeadForeColor" VALUE="0">
												    <PARAM NAME="Columns(1).HasHeadBackColor" VALUE="0">
												    <PARAM NAME="Columns(1).HasForeColor" VALUE="0">
												    <PARAM NAME="Columns(1).HasBackColor" VALUE="0">
												    <PARAM NAME="Columns(1).HeadForeColor" VALUE="0">
												    <PARAM NAME="Columns(1).HeadBackColor" VALUE="0">
												    <PARAM NAME="Columns(1).ForeColor" VALUE="0">
												    <PARAM NAME="Columns(1).BackColor" VALUE="0">
												    <PARAM NAME="Columns(1).HeadStyleSet" VALUE="">
												    <PARAM NAME="Columns(1).StyleSet" VALUE="">
												    <PARAM NAME="Columns(1).Nullable" VALUE="1">
												    <PARAM NAME="Columns(1).Mask" VALUE="">
												    <PARAM NAME="Columns(1).PromptInclude" VALUE="0">
												    <PARAM NAME="Columns(1).ClipMode" VALUE="0">
												    <PARAM NAME="Columns(1).PromptChar" VALUE="95">
													        
												    <PARAM NAME="Columns(2).Width" VALUE="1402">
												    <PARAM NAME="Columns(2).Visible" VALUE="-1">
												    <PARAM NAME="Columns(2).Columns.Count" VALUE="1">
												    <PARAM NAME="Columns(2).Caption" VALUE="Sort Order">
												    <PARAM NAME="Columns(2).Name" VALUE="order">
												    <PARAM NAME="Columns(2).Alignment" VALUE="0">
												    <PARAM NAME="Columns(2).CaptionAlignment" VALUE="3">
												    <PARAM NAME="Columns(2).Bound" VALUE="0">
												    <PARAM NAME="Columns(2).AllowSizing" VALUE="1">
												    <PARAM NAME="Columns(2).DataField" VALUE="Column 2">
												    <PARAM NAME="Columns(2).DataType" VALUE="8">
												    <PARAM NAME="Columns(2).Level" VALUE="0">
												    <PARAM NAME="Columns(2).NumberFormat" VALUE="">
												    <PARAM NAME="Columns(2).Case" VALUE="0">
												    <PARAM NAME="Columns(2).FieldLen" VALUE="256">
												    <PARAM NAME="Columns(2).VertScrollBar" VALUE="0">
												    <PARAM NAME="Columns(2).Locked" VALUE="-1">
												    <PARAM NAME="Columns(2).Style" VALUE="3">
												    <PARAM NAME="Columns(2).ButtonsAlways" VALUE="0">
												    <PARAM NAME="Columns(2).Row.Count" VALUE="2">
												    <PARAM NAME="Columns(2).Col.Count" VALUE="2">
												    <PARAM NAME="Columns(2).Row(0).Col(0)" VALUE="Asc">
												    <PARAM NAME="Columns(2).Row(0).Col(1)" VALUE="">
												    <PARAM NAME="Columns(2).Row(1).Col(0)" VALUE="Desc">
												    <PARAM NAME="Columns(2).Row(1).Col(1)" VALUE="">
												    <PARAM NAME="Columns(2).HasHeadForeColor" VALUE="0">
												    <PARAM NAME="Columns(2).HasHeadBackColor" VALUE="0">
												    <PARAM NAME="Columns(2).HasForeColor" VALUE="0">
												    <PARAM NAME="Columns(2).HasBackColor" VALUE="0">
												    <PARAM NAME="Columns(2).HeadForeColor" VALUE="0">
												    <PARAM NAME="Columns(2).HeadBackColor" VALUE="0">
												    <PARAM NAME="Columns(2).ForeColor" VALUE="0">
												    <PARAM NAME="Columns(2).BackColor" VALUE="0">
												    <PARAM NAME="Columns(2).HeadStyleSet" VALUE="">
												    <PARAM NAME="Columns(2).StyleSet" VALUE="">
												    <PARAM NAME="Columns(2).Nullable" VALUE="1">
												    <PARAM NAME="Columns(2).Mask" VALUE="">
												    <PARAM NAME="Columns(2).PromptInclude" VALUE="0">
												    <PARAM NAME="Columns(2).ClipMode" VALUE="0">
												    <PARAM NAME="Columns(2).PromptChar" VALUE="95">
																										        
												    <PARAM NAME="UseDefaults" VALUE="-1">
												    <PARAM NAME="TabNavigation" VALUE="1">
												    <PARAM NAME="BatchUpdate" VALUE="0">
												    <PARAM NAME="_ExtentX" VALUE="11298">
												    <PARAM NAME="_ExtentY" VALUE="3969">
												    <PARAM NAME="_StockProps" VALUE="79">
												    <PARAM NAME="Caption" VALUE="">
												    <PARAM NAME="ForeColor" VALUE="0">
												    <PARAM NAME="BackColor" VALUE="-2147483633">
												    <PARAM NAME="Enabled" VALUE="-1">
												    <PARAM NAME="DataMember" VALUE="">
												   </OBJECT>
												</TD>

												<TD width=10>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortAdd name=cmdSortAdd value="Add..." style="WIDTH: 100%" class="btn"
													    onclick="sortAdd()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortEdit name=cmdSortEdit value="Edit..." style="WIDTH: 100%" class="btn"
													    onclick="sortEdit()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD width=5>&nbsp;</TD>

												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortRemove name=cmdSortRemove value="Remove" style="WIDTH: 100%" class="btn"
													    onclick="sortRemove()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD width=5>&nbsp;</TD>

												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortRemoveAll name=cmdSortRemoveAll value="Remove All" style="WIDTH: 100%" class="btn"
													    onclick="sortRemoveAll()"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
																									
											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortMoveUp name=cmdSortMoveUp value="Move Up" style="WIDTH: 100%" class="btn"
													    onclick="sortMove(true)"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>
												<TD width=5>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD colspan=4>&nbsp;</TD>
											</TR>
														
											<TR height=5>
												<TD width=5>&nbsp;</TD>
												<TD width=5>&nbsp;</TD>
												<TD width=100>
													<input type="button" id=cmdSortMoveDown name=cmdSortMoveDown value="Move Down" style="WIDTH: 100%" class="btn"
													    onclick="sortMove(false)"
                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                                        onblur="try{button_onBlur(this);}catch(e){}" />
												</TD>

												<TD width=5>&nbsp;</TD>
											</TR>
											
											<TR height=5>
												<TD colspan=5></TD>
											</TR>
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>

						<!-- Fifth tab -->
						<!-- OUTPUT OPTIONS -->
						<DIV id=div5 style="visibility:hidden;display:none">
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=10 CELLPADDING=0>
											<tr>						
												<td valign=top rowspan=2 width=25% height=100%>
													<table class="outline" cellspacing="0" cellpadding="4" width=100% height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																Output Format : <BR><BR>
																<TABLE class="invisible" cellspacing="0" cellpadding="0" width=100%>
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
                                                                                tabindex=-1
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
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat2 value=2
																			    onClick="formatClick(2);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td align=left nowrap>
                                                                            <label 
                                                                                tabindex=-1
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
                                                                                tabindex=-1
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
                                                                                tabindex=-1
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
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat1 value=1 Style="visibility: hidden"
																			    onClick="formatClick(1);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td align=left nowrap>
																			<!--
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat1"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
	        																	CSV File
                                                       	    		        </label>
																			-->
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=4></td>
																	</tr>
																	<tr height=5>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat5 value=5 Style="visibility: hidden"
																			    onClick="formatClick(5);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td>
																			<!--
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat5"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																			    Excel Chart
                                                       	    		        </label>
																			-->
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=4></td>
																	</tr>
																	<tr height=5>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat6 value=6 Style="visibility: hidden"
																			    onClick="formatClick(6);" 
                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td>
																			<!--
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat6"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																			    Excel Pivot Table
                                                       	    		        </label>
																			-->
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
												<td valign=top width=75%>
													<table class="outline" cellspacing="0" cellpadding="4" width=100%  height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																Output Destination(s) : <BR><BR>
																<TABLE class="invisible" cellspacing="0" cellpadding="0" width=100%>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left colspan=6 nowrap>
																			<input name=chkPreview id=chkPreview type=checkbox disabled="disabled" tabindex="-1"
																			    onClick="changeTab5Control();"
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
																			<input name=chkDestination0 id=chkDestination0 type=checkbox disabled="disabled" tabindex="-1" 
																			    onClick="changeTab5Control();"
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
																			<input name=chkDestination1 id=chkDestination1 type=checkbox disabled="disabled" tabindex="-1" 
																			    onClick="changeTab5Control();"
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
																			<select id=cboPrinterName name=cboPrinterName class="combo"width=100% style="WIDTH: 400px" 
																			    onchange="changeTab5Control()">	
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
																			<input name=chkDestination2 id=chkDestination2 type=checkbox disabled="disabled" tabindex="-1"
																			    onClick="changeTab5Control();"
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
																				<TR>
																					<TD>
																						<INPUT id=txtFilename width=100% name=txtFilename class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 375px">
																					</TD>
																					<TD width=25>
																						<INPUT id=cmdFilename width=100% name=cmdFilename class="btn btndisabled" style="WIDTH: 100%" type=button value='...' disabled="disabled"
																						    onclick="saveFile();changeTab5Control();"  
	                                                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                                                                        onfocus="try{button_onFocus(this);}catch(e){}"
	                                                                                        onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>	
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
																			<select id=cboSaveExisting name=cboSaveExisting class="combo" style="WIDTH: 400px" 
																			    onchange="changeTab5Control()">	
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
																			<input name=chkDestination3 id=chkDestination3 type=checkbox disabled="disabled" tabindex="-1"
																			    onClick="changeTab5Control();"
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
																				<TR>
																					<TD>
																						<INPUT id=txtEmailGroup name=txtEmailGroup class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 100%">
																						<INPUT id=txtEmailGroupID name=txtEmailGroupID type=hidden class="text textdisabled" disabled="disabled" tabindex="-1">
																					</TD>
																					<TD width=25>
																						<INPUT id=cmdEmailGroup name=cmdEmailGroup style="WIDTH: 100%" type=button value='...' disabled="disabled" class="btn btndisabled"
																						    onClick="selectEmailGroup();changeTab5Control();" 
	                                                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                                                                        onfocus="try{button_onFocus(this);}catch(e){}"
	                                                                                        onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TD>
																			</TABLE>
																		</TD>
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
																		<TD colspan=2 width=100% nowrap>
																			<INPUT id=txtEmailSubject class="text textdisabled" disabled="disabled" maxlength=255 name=txtEmailSubject style="WIDTH: 400px" 
																			    onchange="frmUseful.txtChanged.value = 1;" 
																			    onkeydown="frmUseful.txtChanged.value = 1;">
																		</TD>
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
																		<TD colspan=2 width=100% nowrap>
																			<INPUT id=txtEmailAttachAs maxlength=255 class="text textdisabled" disabled="disabled" name=txtEmailAttachAs style="WIDTH: 400px" 
																			    onchange="frmUseful.txtChanged.value = 1;" 
																			    onkeydown="frmUseful.txtChanged.value = 1;">
																		</TD>
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
							</TABLE>
						</DIV>
					</td>
					<TD width=10></td>
				</tr> 

				<tr height=10> 
					<td colspan=3></td>
				</tr> 

				<TR height=10>
					<TD width=10></td>
					<TD>
						<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=80>
									<input type=button id=cmdOK name=cmdOK value=OK style="WIDTH: 100%" class="btn"
									    onclick="okClick()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10></TD>
								<TD width=80>
									<input type=button id=cmdCancel name=cmdCancel value=Cancel style="WIDTH: 100%" class="btn"  
									    onclick="cancelClick()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</TR>
						</TABLE>
					</td>
					<TD width=10></td>
				</tr> 

				<tr height=5> 
					<td colspan=3></td>
				</tr> 
			</TABLE>
		</td>
	</tr> 
</TABLE>

<INPUT type='hidden' id=txtBasePicklistID name=txtBasePicklistID>
<INPUT type='hidden' id=txtBaseFilterID name=txtBaseFilterID>
<INPUT type='hidden' id=txtDescExprID name=txtDescExprID>
<INPUT type='hidden' id=txtCustomStartID name=txtCustomStartID>
<INPUT type='hidden' id=txtCustomEndID name=txtCustomEndID>
<INPUT type='hidden' id=txtDatabase name=txtDatabase value="<%=session("Database")%>">

<INPUT type='hidden' id=txtWordVer name=txtWordVer value="<%=Session("WordVer")%>">
<INPUT type='hidden' id=txtExcelVer name=txtExcelVer value="<%=Session("ExcelVer")%>">
<INPUT type='hidden' id=txtWordFormats name=txtWordFormats value="<%=Session("WordFormats")%>">
<INPUT type='hidden' id=txtExcelFormats name=txtExcelFormats value="<%=Session("ExcelFormats")%>">
<INPUT type='hidden' id=txtWordFormatDefaultIndex name=txtWordFormatDefaultIndex value="<%=Session("WordFormatDefaultIndex")%>">
<INPUT type='hidden' id=txtExcelFormatDefaultIndex name=txtExcelFormatDefaultIndex value="<%=Session("ExcelFormatDefaultIndex")%>">

</form>


<form id=frmTables style="visibility:hidden;display:none">
<%
	dim sErrorDescription = ""
	
	' Get the table records.
	Dim cmdTables = Server.CreateObject("ADODB.Command")
	cmdTables.CommandText = "sp_ASRIntGetTablesInfo"
	cmdTables.CommandType = 4	' Stored Procedure
	cmdTables.ActiveConnection = session("databaseConnection")
	
	Response.Write("<B>Set Connection</B>")
	
	Err.Number = 0
	Dim rstTablesInfo = cmdTables.Execute
	
	Response.Write("<B>Executed SP</B>")
	
	If (Err.number <> 0) Then
		sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
	End If

	If Len(sErrorDescription) = 0 Then
		dim iCount = 0
		Do While Not rstTablesInfo.EOF
			Response.Write("<INPUT type='hidden' id=txtTableName_" & rstTablesInfo.fields("tableID").value & " name=txtTableName_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("tableName").value & """>" & vbcrlf)
			Response.Write("<INPUT type='hidden' id=txtTableType_" & rstTablesInfo.fields("tableID").value & " name=txtTableType_" & rstTablesInfo.fields("tableID").value & " value=" & rstTablesInfo.fields("tableType").value & ">" & vbcrlf)
			Response.Write("<INPUT type='hidden' id=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenString").value & """>" & vbcrlf)
			Response.Write("<INPUT type='hidden' id=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenNames").value & """>" & vbcrlf)
			Response.Write("<INPUT type='hidden' id=txtTableParents_" & rstTablesInfo.fields("tableID").value & " name=txtTableParents_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("parentsString").value & """>" & vbcrlf)
			Response.Write("<INPUT type='hidden' id=txtTableRelations_" & rstTablesInfo.fields("tableID").value & " name=txtTableRelations_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("relatedString").value & """>" & vbcrlf)
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

<FORM action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
</FORM>

<form id=frmOriginalDefinition name=frmOriginalDefinition style="visibility:hidden;display:none">
<%
	Dim sErrMsg = ""
	Dim prmUtilID
	
	If LCase(Session("action")) <> "new" Then
		Dim cmdDefn = Server.CreateObject("ADODB.Command")
		cmdDefn.CommandText = "spASRIntGetCalendarReportDefinition"
		cmdDefn.CommandType = 4	' Stored Procedure
		cmdDefn.ActiveConnection = Session("databaseConnection")
		
		prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1)	' 3=integer, 1=input
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

		Dim prmAllRecords = cmdDefn.CreateParameter("allRecords", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmAllRecords)

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
		
		Dim prmPrintFilterHeader = cmdDefn.CreateParameter("printFilterHeader", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmPrintFilterHeader)
		
		Dim prmDesc1ID = cmdDefn.CreateParameter("desc1ID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmDesc1ID)

		Dim prmDesc2ID = cmdDefn.CreateParameter("desc2ID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmDesc2ID)
		
		Dim prmDescExprID = cmdDefn.CreateParameter("descExprID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmDescExprID)
		
		Dim prmDescExprName = cmdDefn.CreateParameter("descExprName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmDescExprName)
		
		Dim prmDescCalcHidden = cmdDefn.CreateParameter("descCalcHidden", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmDescCalcHidden)
		
		Dim prmRegionID = cmdDefn.CreateParameter("regionID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmRegionID)

		Dim prmGroupByDesc = cmdDefn.CreateParameter("groupByDesc", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmGroupByDesc)
		
		Dim prmDescSeparator = cmdDefn.CreateParameter("descSeparator", 200, 2, 8000)	'11=bit, 2=output
		cmdDefn.Parameters.Append(prmDescSeparator)
		
		'-----------------------------------------
		Dim prmStartType = cmdDefn.CreateParameter("startType", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmStartType)

		Dim prmFixedStart = cmdDefn.CreateParameter("fixedStart", 135, 2)	'135=datetime, 2=output
		cmdDefn.Parameters.Append(prmFixedStart)
		
		Dim prmStartFrequency = cmdDefn.CreateParameter("startFrequency", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmStartFrequency)
		
		Dim prmStartPeriod = cmdDefn.CreateParameter("startPeriod", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmStartPeriod)
		
		Dim prmCustomStartID = cmdDefn.CreateParameter("customStartID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmCustomStartID)
		
		Dim prmCustomStartName = cmdDefn.CreateParameter("customStartName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmCustomStartName)
		
		Dim prmStartDateCalcHidden = cmdDefn.CreateParameter("startDateCalcHidden", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmStartDateCalcHidden)
		
		Dim prmEndType = cmdDefn.CreateParameter("endType", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmEndType)

		Dim prmFixedEnd = cmdDefn.CreateParameter("fixedEnd", 135, 2)	'135=datetime, 2=output
		cmdDefn.Parameters.Append(prmFixedEnd)
		
		Dim prmEndFrequency = cmdDefn.CreateParameter("endFrequency", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmEndFrequency)
		
		Dim prmEndPeriod = cmdDefn.CreateParameter("endPeriod", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmEndPeriod)
		
		Dim prmCustomEndID = cmdDefn.CreateParameter("customEndID", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmCustomEndID)

		Dim prmCustomEndName = cmdDefn.CreateParameter("customEndName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmCustomEndName)
		
		Dim prmEndDateCalcHidden = cmdDefn.CreateParameter("endDateCalcHidden", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmEndDateCalcHidden)
		
		Dim prmShadeBHols = cmdDefn.CreateParameter("shadeBHols", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmShadeBHols)
		
		Dim prmShowCaptions = cmdDefn.CreateParameter("showCaptions", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmShowCaptions)
		
		Dim prmShadeWeekends = cmdDefn.CreateParameter("shadeWeekends", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmShadeWeekends)
		
		Dim prmStartOnCurrentMonth = cmdDefn.CreateParameter("startOnCurrentMonth", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmStartOnCurrentMonth)
		
		Dim prmIncludeWorkingDaysOnly = cmdDefn.CreateParameter("includeWorkingDaysOnly", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmIncludeWorkingDaysOnly)
		
		Dim prmIncludeBHols = cmdDefn.CreateParameter("includeBHols", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmIncludeBHols)
		'-----------------------------------------
		Dim prmOutputPreview = cmdDefn.CreateParameter("outputPreview", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputPreview)
		
		Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmOutputFormat)
		
		Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputScreen)
		
		Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputPrinter)
		
		Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputPrinterName)
		
		Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputSave)
		
		Dim prmOutputSaveExisting = cmdDefn.CreateParameter("outputSaveExisting", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmOutputSaveExisting)
		
		Dim prmOutputEmail = cmdDefn.CreateParameter("outputEmail", 11, 2) '11=bit, 2=output
		cmdDefn.Parameters.Append(prmOutputEmail)
		
		Dim prmOutputEmailAddr = cmdDefn.CreateParameter("outputEmailAddr", 3, 2)	'3=integer, 2=output
		cmdDefn.Parameters.Append(prmOutputEmailAddr)
		
		Dim prmOutputEmailAddrName = cmdDefn.CreateParameter("outputEmailAddrName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputEmailAddrName)
		
		Dim prmOutputEmailSubject = cmdDefn.CreateParameter("outputEmailSubject", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputEmailSubject)

		Dim prmOutputEmailAttachAs = cmdDefn.CreateParameter("outputEmailAttachAs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputEmailAttachAs)

		Dim prmOutputFilename = cmdDefn.CreateParameter("outputFilename", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefn.Parameters.Append(prmOutputFilename)
		'-----------------------------------------
	
		Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2)	' 3=integer, 2=output
		cmdDefn.Parameters.Append(prmTimestamp)

		Err.Number = 0
		Dim rstDefinition = cmdDefn.Execute
		
		Dim iHiddenEventFilterCount = 0
		dim iCount = 0
		If (Err.Number <> 0) Then
			sErrMsg = CType(("'" & Session("utilname") & "' report definition could not be read." & vbCrLf) & FormatError(Err.Description), String)
		Else
			If rstDefinition.state <> 0 Then
				' Read recordset values.
				iCount = 0
				Do While Not rstDefinition.EOF
					iCount = iCount + 1
					
					Response.Write("<INPUT type='hidden' id=txtReportDefnEvent_" & iCount & " name=txtReportDefnEvent_" & iCount & " value=""" & Replace(rstDefinition.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
					
					If rstDefinition.fields("FilterHidden").value = "Y" Then
						iHiddenEventFilterCount = iHiddenEventFilterCount + 1
					End If
					
					rstDefinition.MoveNext()
				Loop

				' Release the ADO recordset object.
				rstDefinition.close()
			End If
			rstDefinition = Nothing
			
			Session("hiddenfiltercount") = iHiddenEventFilterCount
			Session("CalendarEventCount") = iCount
			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			If Len(cmdDefn.Parameters("errMsg").value) > 0 Then
				sErrMsg = CType(("'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").value), String)
			End If

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
			Response.Write("<INPUT type='hidden' id=txtDefn_PrintFilterHeader name=txtDefn_PrintFilterHeader value=" & cmdDefn.Parameters("PrintFilterHeader").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Desc1ID name=txtDefn_Desc1ID value=" & cmdDefn.Parameters("Desc1ID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_Desc2ID name=txtDefn_Desc2ID value=" & cmdDefn.Parameters("Desc2ID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_DescExprID name=txtDefn_DescExprID value=" & cmdDefn.Parameters("descExprID").value & ">" & vbCrLf)
			If IsDBNull(cmdDefn.Parameters("descExprName").value) Then
				Response.Write("<INPUT type='hidden' id=txtDefn_DescExprName name=txtDefn_DescExprName value="""">" & vbCrLf)
			Else
				Response.Write("<INPUT type='hidden' id=txtDefn_DescExprName name=txtDefn_DescExprName value=""" & Replace(cmdDefn.Parameters("descExprName").value, """", "&quot;") & """>" & vbCrLf)
			End If
			Response.Write("<INPUT type='hidden' id=txtDefn_DescExprHidden name=txtDefn_DescExprHidden value=" & cmdDefn.Parameters("descCalcHidden").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_RegionID name=txtDefn_RegionID value=" & cmdDefn.Parameters("RegionID").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_GroupByDesc name=txtDefn_GroupByDesc value=" & cmdDefn.Parameters("GroupByDesc").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_DescSeparator name=txtDefn_DescSeparator value=""" & cmdDefn.Parameters("DescSeparator").value & """>" & vbCrLf)

			Response.Write("<INPUT type='hidden' id=txtDefn_StartType name=txtDefn_StartType value=" & cmdDefn.Parameters("StartType").value & ">" & vbCrLf)
			
			If IsDBNull(cmdDefn.Parameters("FixedStart").value) Then
				Response.Write("<INPUT type='hidden' id=txtDefn_FixedStart name=txtDefn_FixedStart value="""">" & vbCrLf)
			Else
				Response.Write("<INPUT type='hidden' id=txtDefn_FixedStart name=txtDefn_FixedStart value=" & convertSQLDateToLocale(cmdDefn.Parameters("FixedStart").value) & ">" & vbCrLf)
			End If
			Response.Write("<INPUT type='hidden' id=txtDefn_StartFrequency name=txtDefn_StartFrequency value=" & cmdDefn.Parameters("StartFrequency").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_StartPeriod name=txtDefn_StartPeriod value=" & cmdDefn.Parameters("StartPeriod").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_CustomStartID name=txtDefn_CustomStartID value=" & cmdDefn.Parameters("customStartID").value & ">" & vbCrLf)
			If IsDBNull(cmdDefn.Parameters("customStartName").value) Then
				Response.Write("<INPUT type='hidden' id=txtDefn_CustomStartName name=txtDefn_CustomStartName value="""">" & vbCrLf)
			Else
				Response.Write("<INPUT type='hidden' id=txtDefn_CustomStartName name=txtDefn_CustomStartName value=""" & Replace(cmdDefn.Parameters("customStartName").value, """", "&quot;") & """>" & vbCrLf)
			End If
			Response.Write("<INPUT type='hidden' id=txtDefn_CustomStartCalcHidden name=txtDefn_CustomStartCalcHidden value=" & cmdDefn.Parameters("startDateCalcHidden").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_EndType name=txtDefn_EndType value=" & cmdDefn.Parameters("EndType").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_FixedEnd name=txtDefn_FixedEnd value=" & convertSQLDateToLocale(cmdDefn.Parameters("FixedEnd").value) & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_EndFrequency name=txtDefn_EndFrequency value=" & cmdDefn.Parameters("EndFrequency").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_EndPeriod name=txtDefn_EndPeriod value=" & cmdDefn.Parameters("EndPeriod").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_CustomEndID name=txtDefn_CustomEndID value=" & cmdDefn.Parameters("customEndID").value & ">" & vbCrLf)
			If IsDBNull(cmdDefn.Parameters("customEndName").value) Then
				Response.Write("<INPUT type='hidden' id=txtDefn_CustomEndName name=txtDefn_CustomEndName value="""">" & vbCrLf)
			Else
				Response.Write("<INPUT type='hidden' id=txtDefn_CustomEndName name=txtDefn_CustomEndName value=""" & Replace(cmdDefn.Parameters("customEndName").value, """", "&quot;") & """>" & vbCrLf)
			End If
			Response.Write("<INPUT type='hidden' id=txtDefn_CustomEndCalcHidden name=txtDefn_CustomEndCalcHidden value=" & cmdDefn.Parameters("endDateCalcHidden").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_ShadeBHols name=txtDefn_ShadeBHols value=" & cmdDefn.Parameters("ShadeBHols").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_ShowCaptions name=txtDefn_ShowCaptions value=" & cmdDefn.Parameters("ShowCaptions").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_ShadeWeekends name=txtDefn_ShadeWeekends value=" & cmdDefn.Parameters("ShadeWeekends").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_StartOnCurrentMonth name=txtDefn_StartOnCurrentMonth value=" & cmdDefn.Parameters("StartOnCurrentMonth").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_IncludeWorkingDaysOnly name=txtDefn_IncludeWorkingDaysOnly value=" & cmdDefn.Parameters("IncludeWorkingDaysOnly").value & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtDefn_IncludeBHols name=txtDefn_IncludeBHols value=" & cmdDefn.Parameters("IncludeBHols").value & ">" & vbCrLf)
			
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

			'********************************************************************************

			Dim cmdReportOrder = Server.CreateObject("ADODB.Command")
			cmdReportOrder.CommandText = "spASRIntGetCalendarReportOrder"
			cmdReportOrder.CommandType = 4	'Stored Procedure
			cmdReportOrder.ActiveConnection = Session("databaseConnection")
		
			Dim prmUtilID2 = cmdReportOrder.CreateParameter("utilID2", 3, 1) ' 3=integer, 1=input
			cmdReportOrder.Parameters.Append(prmUtilID2)
			prmUtilID2.value = CleanNumeric(Session("utilid"))
		
			Dim prmErrMsg2 = cmdReportOrder.CreateParameter("errMsg2", 200, 2, 8000) '200=varchar, 2=output, 8000=size
			cmdReportOrder.Parameters.Append(prmErrMsg2)

			Err.Number = 0
			Dim rstOrder = cmdReportOrder.Execute
		
			iCount = 0
			If (Err.Number <> 0) Then
				sErrMsg = "'" & Session("utilname") & "' report order definition could not be read." & vbCrLf & FormatError(Err.Description)
			Else
				If rstOrder.state <> 0 Then
					' Read recordset values.
			
					Do While Not rstOrder.EOF
						iCount = iCount + 1
						Response.Write("<INPUT type='hidden' id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & rstOrder.fields("orderString").value & """>" & vbCrLf)

						rstOrder.MoveNext()
					Loop
					' Release the ADO recordset object.
					rstOrder.close()
				End If
				rstOrder = Nothing
			End If

			Session("CalendarOrderCount") = iCount

			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			If Len(cmdReportOrder.Parameters("errMsg2").value) > 0 Then
				sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg2").value
			End If

			cmdReportOrder = Nothing

			'********************************************************************************

		End If

		' Release the ADO command object.
		cmdDefn = Nothing

		If Len(sErrMsg) > 0 Then
			Session("confirmtext") = sErrMsg
			Session("confirmtitle") = "OpenHR Intranet"
	                                    	    
			Session("followpage") = "defsel"
	                                    	    
			Session("reaction") = "CALENDARREPORTS"
	                                    	    			Response.Clear()
	                                    	    			Response.Redirect("confirmok")
		End If
	
	Else
		Session("CalendarEventCount") = 0
		Session("CalendarOrderCount") = 0
		Session("hiddenfiltercount") = 0
	End If
	
%>
</form>

<form id=frmAccess>
<%
	sErrorDescription = ""
	
	' Get the table records.
	Dim cmdAccess = Server.CreateObject("ADODB.Command")
	cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
	cmdAccess.CommandType = 4	' Stored Procedure
	cmdAccess.ActiveConnection = Session("databaseConnection")

	Dim prmUtilType = cmdAccess.CreateParameter("utilType", 3, 1)	' 3=integer, 1=input
	cmdAccess.Parameters.Append(prmUtilType)
	prmUtilType.value = 17 ' 17 = calendar report

	prmUtilID = cmdAccess.Cr0eateParameter("utilID", 3, 1)	' 3=integer, 1=input
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
			Response.Write("<INPUT type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & rstAccessInfo.fields("accessDefinition").value & """>" & vbcrlf)

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
	<INPUT type="hidden" id=txtFirstLoad name=txtFirstLoad value="Y">
	<INPUT type="hidden" id=txtCurrentBaseTableID name=txtCurrentBaseTableID>
	<INPUT type="hidden" id=txtAvailableColumnsLoaded name=txtAvailableColumnsLoaded value=0>
	<INPUT type="hidden" id=txtEventsLoaded name=txtEventsLoaded value=0>
	<INPUT type="hidden" id=txtSortLoaded name=txtSortLoaded value=0>
	<INPUT type="hidden" id=txtChanged name=txtChanged value=0>
	<INPUT type="hidden" id=txtUtilID name=txtUtilID value=<%=session("utilid")%>>
	<INPUT type="hidden" id=txtEventCount name=txtEventCount value=<%=session("CalendarEventCount")%>>
	<INPUT type="hidden" id=txtOrderCount name=txtOrderCount value=<%=session("CalendarOrderCount")%>>
	<INPUT type="hidden" id=txtHiddenEventFilterCount name=txtHiddenEventFilterCount value=<%=session("hiddenfiltercount")%>>
	<INPUT type="hidden" id=txtLockGridEvents name=txtLockGridEvents value=0>
<%
	Dim cmdDefinition = Server.CreateObject("ADODB.Command")
	cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
	cmdDefinition.CommandType = 4	' Stored procedure.
	cmdDefinition.ActiveConnection = Session("databaseConnection")

	Dim prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
	cmdDefinition.Parameters.Append(prmModuleKey)
	prmModuleKey.value = "MODULE_PERSONNEL"

	Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
	cmdDefinition.Parameters.Append(prmParameterKey)
	prmParameterKey.value = "Param_TablePersonnel"

	Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
	cmdDefinition.Parameters.Append(prmParameterValue)

	Err.Number = 0
	cmdDefinition.Execute()

	Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbcrlf)
	
	cmdDefinition = Nothing

	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbcrlf)
	Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & session("action") & ">" & vbcrlf)
%>
</FORM>

<FORM id=frmValidate name=frmValidate target=validate method=post action=util_validate_calendarreport style="visibility:hidden;display:none">
	<INPUT type=hidden id="validateBaseFilter" name=validateBaseFilter value=0>
	<INPUT type=hidden id="validateBasePicklist" name=validateBasePicklist value=0>
	<INPUT type=hidden id="validateEmailGroup" name=validateEmailGroup value=0>
	<INPUT type=hidden id="validateEventFilter" name=validateEventFilter value=0> 
	<INPUT type=hidden id="validateDescExpr" name=validateDescExpr value=0>
	<INPUT type=hidden id="validateCustomStart" name=validateCustomStart value=0>
	<INPUT type=hidden id="validateCustomEnd" name=validateCustomEnd value=0>
	<INPUT type=hidden id="validateHiddenGroups" name=validateHiddenGroups value = ''>
	<INPUT type=hidden id="validateName" name=validateName value=''>
	
	<INPUT type=hidden id="validateTimestamp" name=validateTimestamp value=''>
	<INPUT type=hidden id="validateUtilID" name=validateUtilID value=''>
</FORM>

<FORM id=frmSend name=frmSend method=post action=util_def_calendarreport_Submit style="visibility:hidden;display:none">

	<INPUT type="hidden" id=txtSend_ID name=txtSend_ID>	
	<INPUT type="hidden" id=txtSend_name name=txtSend_name>
	<INPUT type="hidden" id=txtSend_description name=txtSend_description>
	<INPUT type="hidden" id=txtSend_access name=txtSend_access>
	<INPUT type="hidden" id=txtSend_userName name=txtSend_userName>
	<INPUT type="hidden" id=txtSend_baseTable name=txtSend_baseTable>
	<INPUT type="hidden" id=txtSend_allRecords name=txtSend_allRecords>
	<INPUT type="hidden" id=txtSend_picklist name=txtSend_picklist>
	<INPUT type="hidden" id=txtSend_filter name=txtSend_filter>
	<INPUT type="hidden" id=txtSend_printFilterHeader name=txtSend_printFilterHeader>
	<INPUT type="hidden" id=txtSend_desc1 name=txtSend_desc1>
	<INPUT type="hidden" id=txtSend_desc2 name=txtSend_desc2>
	<INPUT type="hidden" id=txtSend_descExpr name=txtSend_descExpr>
	<INPUT type="hidden" id=txtSend_region name=txtSend_region>
	<INPUT type="hidden" id=txtSend_groupbydesc name=txtSend_groupbydesc>
	<INPUT type="hidden" id=txtSend_descseparator name=txtSend_descseparator>
	
	<INPUT type="hidden" id=txtSend_StartType name=txtSend_StartType>
	<INPUT type="hidden" id=txtSend_FixedStart name=txtSend_FixedStart>
	<INPUT type="hidden" id=txtSend_StartFrequency name=txtSend_StartFrequency>
	<INPUT type="hidden" id=txtSend_StartPeriod name=txtSend_StartPeriod>
	<INPUT type="hidden" id=txtSend_CustomStart name=txtSend_CustomStart>
	<INPUT type="hidden" id=txtSend_EndType name=txtSend_EndType>
	<INPUT type="hidden" id=txtSend_FixedEnd name=txtSend_FixedEnd>
	<INPUT type="hidden" id=txtSend_EndFrequency name=txtSend_EndFrequency>
	<INPUT type="hidden" id=txtSend_EndPeriod name=txtSend_EndPeriod>
	<INPUT type="hidden" id=txtSend_CustomEnd name=txtSend_CustomEnd>
	
	<INPUT type="hidden" id=txtSend_IncludeBHols name=txtSend_IncludeBHols>
	<INPUT type="hidden" id=txtSend_IncludeWorkingDaysOnly name=txtSend_IncludeWorkingDaysOnly>
	<INPUT type="hidden" id=txtSend_ShadeBHols name=txtSend_ShadeBHols>
	<INPUT type="hidden" id=txtSend_Captions name=txtSend_Captions>
	<INPUT type="hidden" id=txtSend_ShadeWeekends name=txtSend_ShadeWeekends>
	<INPUT type="hidden" id=txtSend_StartOnCurrentMonth name=txtSend_StartOnCurrentMonth>

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
	
	<INPUT type="hidden" id=txtSend_columns name=txtSend_Events>
	<INPUT type="hidden" id=txtSend_columns2 name=txtSend_Events2>

	<INPUT type="hidden" id=txtSend_OrderString name=txtSend_OrderString>
	
	<INPUT type="hidden" id=txtSend_reaction name=txtSend_reaction>

	<INPUT type="hidden" id=txtSend_jobsToHide name=txtSend_jobsToHide>
	<INPUT type="hidden" id=txtSend_jobsToHideGroups name=txtSend_jobsToHideGroups>
</FORM>

<FORM id=frmEventDetails name=frmEventDetails target="eventselection" action="util_def_calendarreportdates_main" method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=eventAction name=eventAction>
	<INPUT type="hidden" id=eventName name=eventName>
	<INPUT type="hidden" id=eventID name=eventID>
	<INPUT type="hidden" id=eventTableID name=eventTableID>
	<INPUT type="hidden" id=eventTable name=eventTable>
	<INPUT type="hidden" id=eventFilterID name=eventFilterID>
	<INPUT type="hidden" id=eventFilter name=eventFilter>
	<INPUT type="hidden" id=eventFilterHidden name=eventFilterHidden>
	
	<INPUT type="hidden" id=eventStartDateID name=eventStartDateID>
	<INPUT type="hidden" id=eventStartDate name=eventStartDate>
	<INPUT type="hidden" id=eventStartSessionID name=eventStartSessionID>
	<INPUT type="hidden" id=eventStartSession name=eventStartSession>
	
	<INPUT type="hidden" id=eventEndDateID name=eventEndDateID>
	<INPUT type="hidden" id=eventEndDate name=eventEndDate>
	<INPUT type="hidden" id=eventEndSessionID name=eventEndSessionID>
	<INPUT type="hidden" id=eventEndSession name=eventEndSession>
	
	<INPUT type="hidden" id=eventDurationID name=eventDurationID>
	<INPUT type="hidden" id=eventDuration name=eventDuration>
	
	<INPUT type="hidden" id=eventLookupType name=eventLookupType>
	<INPUT type="hidden" id=eventKeyCharacter name=eventKeyCharacter>
	<INPUT type="hidden" id=eventLookupTableID name=eventLookupTableID>
	<INPUT type="hidden" id=eventLookupColumnID name=eventLookupColumnID>
	<INPUT type="hidden" id=eventLookupCodeID name=eventLookupCodeID>
	<INPUT type="hidden" id=eventTypeColumnID name=eventTypeColumnID>
	
	<INPUT type="hidden" id=eventDesc1ID name=eventDesc1ID>
	<INPUT type="hidden" id=eventDesc1 name=eventDesc1>
	<INPUT type="hidden" id=eventDesc2ID name=eventDesc2ID>
	<INPUT type="hidden" id=eventDesc2 name=eventDesc2>
	
	<INPUT type="hidden" id=relationNames name=relationNames>
	
</FORM>

<FORM id=frmRecordSelection name=frmRecordSelection target="recordSelection" action="util_recordSelection" method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=recSelType name=recSelType>
	<INPUT type="hidden" id=recSelTableID name=recSelTableID>
	<INPUT type="hidden" id=recSelCurrentID name=recSelCurrentID>
	<INPUT type="hidden" id=recSelTable name=recSelTable>
	<INPUT type="hidden" id=recSelDefOwner name=recSelDefOwner>
</FORM>

<FORM id=frmCalcSelection name=frmCalcSelection target="calcSelection" action="util_calcSelection" method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=calcSelRecInd name=calcSelRecInd>
	<INPUT type="hidden" id=calcSelType name=calcSelType>
	<INPUT type="hidden" id=calcSelTableID name=calcSelTableID>
	<INPUT type="hidden" id=calcSelCurrentID name=calcSelCurrentID>
	<INPUT type="hidden" id=Hidden1 name=recSelDefOwner>
</FORM>

<FORM id=frmEmailSelection name=frmEmailSelection target="emailSelection" action="util_emailSelection" method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=EmailSelCurrentID name=EmailSelCurrentID>
</FORM>

<FORM id=frmSortOrder name=frmSortOrder action="util_sortorderselection" target="sortorderselection" method=post style="visibility:hidden;display:none">
	<INPUT type=hidden id=txtSortInclude name=txtSortInclude>
	<INPUT type=hidden id=txtSortExclude name=txtSortExclude>
	<INPUT type=hidden id=txtSortEditing name=txtSortEditing>
	<INPUT type=hidden id=txtSortColumnID name=txtSortColumnID>
	<INPUT type=hidden id=txtSortColumnName name=txtSortColumnName>
	<INPUT type=hidden id=txtSortOrder name=txtSortOrder>	
</FORM>

<FORM id=frmSelectionAccess name=frmSelectionAccess style="visibility:hidden;display:none">
	<INPUT type="hidden" id=forcedHidden name=forcedHidden value="N">
	<INPUT type="hidden" id=baseHidden name=baseHidden value="N">
	<INPUT type="hidden" id=eventHidden name=eventHidden value=0>
	<INPUT type="hidden" id=descHidden name=descHidden value="N">
	<INPUT type="hidden" id=calcStartDateHidden name=calcStartDateHidden value="N">
	<INPUT type="hidden" id=calcEndDateHidden name=calcEndDateHidden value="N">
</FORM>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

</div>

<script type="text/javascript">
	util_def_calendarreport_window_onload();
	//util_def_mailmerge_addActiveXHandlers();
</script>