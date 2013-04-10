<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">
	function util_def_calendarreportdates_window_onload() {
		frmPopup.txtLoading.value = 1;
		frmPopup.txtFirstLoad_Event.value = 1;
		frmPopup.txtFirstLoad_Lookup.value = 1;
		frmPopup.txtHaveSetLookupValues.value = 0;

		button_disable(frmPopup.cmdCancel, true);
		button_disable(frmPopup.cmdOK, true);

		populateEventTableCombo();

		frmPopup.cboLegendTable.selectedIndex = -1;

		var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");
		var frmDef = OpenHR.getForm("workframe", "frmDefinition");

		if (frmEvent.eventAction.value.toUpperCase() == "NEW") {
			frmPopup.optNoEnd.checked = true;
			frmPopup.optCharacter.checked = true;
		} else {
			frmPopup.rowID.value = frmDef.grdEvents.AddItemRowIndex(frmDef.grdEvents.Bookmark);

			frmPopup.txtEventName.value = frmEvent.eventName.value;
			setEventTable(frmEvent.eventTableID.value);
			frmPopup.txtEventFilter.value = frmEvent.eventFilter.value;
			frmPopup.txtEventFilterID.value = frmEvent.eventFilterID.value;
			frmSelectionAccess.baseHidden.value = frmEvent.eventFilterHidden.value;
		}

		disabledAll();

		frmPopup.txtEventColumnsLoaded.value = 0;
		frmPopup.txtLookupColumnsLoaded.value = 0;

		populateEventColumns();

		frmPopup.txtLoading.value = 0;
	}

<%--<script FOR=window EVENT=onfocus LANGUAGE=JavaScript>--%>
	var frmEvent = window.dialogarguments.OpenHR.getForm("workframe", "frmEventDetails");
</script>

<!-- Have added the object to this page so the message box has focus infront of the child window. -->
<script type="text/javascript" id="scptGeneralFunctions">
	function disabledAll() {
		with (frmPopup) {
			/*Event Frame*/
			text_disable(txtEventName, true);
			combo_disable(cboEventTable, true);
			text_disable(txtEventFilter, true);
			button_disable(cmdEventFilter, true);

			/*Event Start Frame*/
			combo_disable(cboStartDate, true);
			combo_disable(cboStartSession, true);

			/*Event End Frame*/
			radio_disable(optNoEnd, true);
			radio_disable(optEndDate, true);
			combo_disable(cboEndDate, true);
			combo_disable(cboEndSession, true);
			radio_disable(optDuration, true);
			combo_disable(cboDuration, true);

			/*Key Frame*/
			text_disable(txtCharacter, true);
			combo_disable(cboEventType, true);
			combo_disable(cboLegendTable, true);
			combo_disable(cboLegendColumn, true);
			combo_disable(cboLegendCode, true);

			/*Event Description Frame*/
			combo_disable(cboEventDesc1, true);
			combo_disable(cboEventDesc2, true);
		}
	}

	function populateEventColumns() {
		// Get the columns/calcs for the current table selection.
		var frmGetDataForm = OpenHR.getForm("calendardataframe", "frmGetCalendarData");
		var frmDef = OpenHR.getForm("workframe", "frmDefinition");

		frmGetDataForm.txtCalendarAction.value = "LOADCALENDAREVENTDETAILSCOLUMNS";
		frmGetDataForm.txtCalendarBaseTableID.value = frmDef.cboBaseTable.options[frmDef.cboBaseTable.selectedIndex].value;
		frmGetDataForm.txtCalendarEventTableID.value = frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].value;

		//window.parent.frames("calendardataframe").refreshData();
		OpenHR.getFrame("calendardataframe").RefreshData();
	}

	function loadAvailableEventColumns() {
		var i;
		var sSelectedIDs;
		var sTemp;
		var iIndex;
		var sType;
		var sID;
		var iDummy;
		var frmRefresh;

		with (frmPopup) {
			window.cboStartDate.length = 0;
			window.cboStartSession.length = 0;
			window.cboEndDate.length = 0;
			window.cboEndSession.length = 0;
			window.cboDuration.length = 0;
			window.cboEventType.length = 0;
			window.cboEventDesc1.length = 0;
			window.cboEventDesc2.length = 0;
			window.cboEventType.length = 0;

			var oOption = document.createElement("OPTION");
			cboStartSession.options.add(oOption);
			oOption.innerText = "<None>";
			oOption.value = 0;

			oOption = document.createElement("OPTION");
			cboEndDate.options.add(oOption);
			oOption.innerText = "<None>";
			oOption.value = 0;

			oOption = document.createElement("OPTION");
			cboEndSession.options.add(oOption);
			oOption.innerText = "<None>";
			oOption.value = 0;

			oOption = document.createElement("OPTION");
			cboDuration.options.add(oOption);
			oOption.innerText = "<None>";
			oOption.value = 0;

			oOption = document.createElement("OPTION");
			cboEventDesc1.options.add(oOption);
			oOption.innerText = "<None>";
			oOption.value = 0;

			oOption = document.createElement("OPTION");
			cboEventDesc2.options.add(oOption);
			oOption.innerText = "<None>";
			oOption.value = 0;

			combo_disable(window.cboStartDate, true);
			combo_disable(cboStartSession, true);
			combo_disable(cboEndDate, true);
			combo_disable(cboEndSession, true);
			combo_disable(cboDuration, true);
			combo_disable(cboEventDesc1, true);
			combo_disable(cboEventDesc2, true);
			combo_disable(cboEventType, true);

			var frmUtilDefForm = OpenHR.getFrame("calendardataframe", "frmCalendarData");
			var dataCollection = frmUtilDefForm.elements;

			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {

					sControlName = dataCollection.item(i).name;

					window.cboStartDate.selectedIndex = -1;
					cboStartSession.selectedIndex = -1;
					cboEndDate.selectedIndex = -1;
					cboEndSession.selectedIndex = -1;
					cboDuration.selectedIndex = -1;
					cboEventType.selectedIndex = -1;
					cboEventDesc1.selectedIndex = -1;
					cboEventDesc2.selectedIndex = -1;
					cboEventType.selectedIndex = -1;

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
								window.cboStartDate.options.add(oOption);
								oOption.innerText = dataCollection.item(i).value;
								oOption.value = sColumnID;

								oOption = document.createElement("OPTION");
								cboEndDate.options.add(oOption);
								oOption.innerText = dataCollection.item(i).value;
								oOption.value = sColumnID;
							}

							if (iDataTypeControlValue == 12 && iSizeControlValue == 2) {
								oOption = document.createElement("OPTION");
								cboStartSession.options.add(oOption);
								oOption.innerText = dataCollection.item(i).value;
								oOption.value = sColumnID;

								oOption = document.createElement("OPTION");
								cboEndSession.options.add(oOption);
								oOption.innerText = dataCollection.item(i).value;
								oOption.value = sColumnID;
							}

							if (iDataTypeControlValue == 2 || iDataTypeControlValue == 4) {
								oOption = document.createElement("OPTION");
								cboDuration.options.add(oOption);
								oOption.innerText = dataCollection.item(i).value;
								oOption.value = sColumnID;
							}

							if (iTypeControlValue == 1) {
								oOption = document.createElement("OPTION");
								cboEventType.options.add(oOption);
								oOption.innerText = dataCollection.item(i).value;
								oOption.value = sColumnID;
							}
						}

						oOption = document.createElement("OPTION");
						cboEventDesc1.options.add(oOption);
						oOption.innerText = sTableNameControlValue + '.' + dataCollection.item(i).value;
						oOption.value = sColumnID;

						oOption = document.createElement("OPTION");
						cboEventDesc2.options.add(oOption);
						oOption.innerText = sTableNameControlValue + '.' + dataCollection.item(i).value;
						oOption.value = sColumnID;
					}
				}
			}

			if ((window.cboStartDate.selectedIndex < 0)
				&& (window.cboStartDate.length > 0)) {
				window.cboStartDate.selectedIndex = 0;
			}
		}

		var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");

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
		if (frmPopup.txtLookupColumnsLoaded.value == 0) {
			// Get the columns/calcs for the current table selection.
			var frmGetDataForm = OpenHR.getFrame("calendardataframe", "frmGetCalendarData");

			if ((frmPopup.cboLegendTable.options.length > 0) && (frmPopup.cboLegendTable.selectedIndex < 0)) {
				frmPopup.cboLegendTable.selectedIndex = 0;
			}

			frmGetDataForm.txtCalendarAction.value = "LOADCALENDAREVENTKEYLOOKUPCOLUMNS";
			frmGetDataForm.txtCalendarLookupTableID.value = frmPopup.cboLegendTable.options[frmPopup.cboLegendTable.selectedIndex].value;

			OpenHR.getFrame("calendardataframe").refreshData();
		}
		else {
			return;
		}
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
		
		with (frmPopup) {
			cboLegendColumn.length = 0;
			cboLegendCode.length = 0;

			var frmUtilDefForm = OpenHR.getForm("calendardataframe", "frmCalendarData");
			var dataCollection = frmUtilDefForm.elements;

			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;

					if (sControlName.substr(0, 10) == "txtRepCol_") {
						sColumnID = sControlName.substring(10, sControlName.length);

						var sDataTypeControlName = "txtRepColDataType_" + sColumnID;
						var iDataTypeControlValue = frmUtilDefForm.elements(sDataTypeControlName).value;

						if (iDataTypeControlValue == 12) {
							var oOption = document.createElement("OPTION");
							cboLegendColumn.options.add(oOption);
							oOption.innerText = dataCollection.item(i).value;
							oOption.value = sColumnID;

							var oOption = document.createElement("OPTION");
							cboLegendCode.options.add(oOption);
							oOption.innerText = dataCollection.item(i).value;
							oOption.value = sColumnID;
						}
					}
				}
			}

			//document.parentWindow.parent.window.dialogArguments.window.refreshTab3Controls();		  
			OpenHR.refreshTab3Controls();
		}

		var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");

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

	function populateEventTableCombo() {
		var frmTab = OpenHR.getForm("workframe", "frmTables");
		var frmDef = OpenHR.getForm("workframe", "frmDefinition");
		var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");

		var sRelationString = frmEvent.relationNames.value;
		var iRelationID;
		var sTableName;

		var bAdded = false;

		var dataCollection = frmTab.elements;
		var oOption;

		var iIndex = sRelationString.indexOf("	");
		while (iIndex > 0) {
			iRelationID = sRelationString.substr(0, iIndex);

			frmRefresh.submit();

			bAdded = true;

			oOption = document.createElement("OPTION");
			frmPopup.cboEventTable.options.add(oOption);
			oOption.value = iRelationID;

			if (iRelationID == frmDef.cboBaseTable.options[frmDef.cboBaseTable.selectedIndex].value) {
				oOption.selected = true;
			}

			sRelationString = sRelationString.substr(iIndex + 1);
			iIndex = sRelationString.indexOf("	");

			sTableName = sRelationString.substr(0, iIndex);
			oOption.innerText = sTableName;

			if (bAdded) {
				sRelationString = sRelationString.substr(iIndex + 1);
				iIndex = sRelationString.indexOf("	");

				bAdded = false;
			}
			else {
				sRelationString = sRelationString.substr(iIndex + 1);
				iIndex = sRelationString.indexOf("	");

				sRelationString = sRelationString.substr(iIndex + 1);
				iIndex = sRelationString.indexOf("	");

				bAdded = false;
			}
		}
	}

	function setEventTable(piTableID) {
		var i;
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
		for (i = 0; i < frmPopup.cboStartDate.options.length; i++) {
			if (frmPopup.cboStartDate.options(i).value == piColumnID) {
				frmPopup.cboStartDate.selectedIndex = i;
				return;
			}
		}
		frmPopup.cboStartDate.selectedIndex = 0;
	}

	function setStartSession(piColumnID) {
		var i;
		for (i = 0; i < frmPopup.cboStartSession.options.length; i++) {
			if (frmPopup.cboStartSession.options(i).value == piColumnID) {
				frmPopup.cboStartSession.selectedIndex = i;
				return;
			}
		}
		frmPopup.cboStartSession.selectedIndex = 0;
	}

	function setEndDate(piColumnID) {
		var i;
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
		for (i = 0; i < frmPopup.cboLegendCode.options.length; i++) {
			if (frmPopup.cboLegendCode.options(i).value == piColumnID) {
				frmPopup.cboLegendCode.selectedIndex = i;
				return;
			}
		}
		frmPopup.cboLegendCode.selectedIndex = 0;
	}

	function setEventTypeColumn(piColumnID) {
		var i;
		for (i = 0; i < frmPopup.cboEventType.options.length; i++) {
			if (frmPopup.cboEventType.options(i).value == piColumnID) {
				frmPopup.cboEventType.selectedIndex = i;
				return;
			}
		}
		frmPopup.cboEventType.selectedIndex = 0;
	}

	function setDesc1Column(piColumnID) {
		var i;
		for (i = 0; i < frmPopup.cboEventDesc1.options.length; i++) {
			if (frmPopup.cboEventDesc1.options(i).value == piColumnID) {
				frmPopup.cboEventDesc1.selectedIndex = i;
				return;
			}
		}
		frmPopup.cboEventDesc1.selectedIndex = 0;
	}

	function setDesc2Column(piColumnID) {
		var i;
		for (i = 0; i < frmPopup.cboEventDesc2.options.length; i++) {
			if (frmPopup.cboEventDesc2.options(i).value == piColumnID) {
				frmPopup.cboEventDesc2.selectedIndex = i;
				return;
			}
		}
		frmPopup.cboEventDesc2.selectedIndex = 0;
	}

	function setLookupValues() {
		var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");
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
		var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");

		with (frmEvent) {
			setStartDate(eventStartDateID.value);

			if (eventStartSessionID.value > 0) {
				setStartSession(eventStartSessionID.value);
			}

			if (eventEndDateID.value > 0) {
				frmPopup.optEndDate.checked = true;
				setEndDate(eventEndDateID.value);
				if (eventEndSessionID.value > 0) {
					setEndSession(eventEndSessionID.value);
				}
			}
			else if (eventDurationID.value > 0) {
				frmPopup.optDuration.checked = true;
				setDuration(eventDurationID.value);
			}
			else {
				frmPopup.optNoEnd.checked = true;
			}

			if (eventDesc1ID.value > 0) {
				setDesc1Column(eventDesc1ID.value);
			}

			if (eventDesc2ID.value > 0) {
				setDesc2Column(eventDesc2ID.value);
			}

			refreshEventControls();

			if ((eventLookupType.value == 1)) {
				setLookupTable(eventLookupTableID.value);

				populateLookupColumns();
			}
			else {
				setLookupValues();
			}

			refreshLegendControls();
		}
	}

	function changeEventTable() {
		with (frmPopup) {
			txtEventFilterID.value = 0;
			txtEventFilter.value = "";

			cboStartDate.length = 0;
			cboStartSession.length = 0;

			cboEndDate.length = 0;
			cboEndSession.length = 0;
			cboDuration.length = 0;

			cboEventType.length = 0;

			cboEventDesc1.length = 0;
			cboEventDesc2.length = 0;

			txtEventColumnsLoaded.value = 0;
		}

		populateEventColumns();
		refreshEventControls();
		eventChanged();
	}

	function changeLegendTable() {
		with (frmPopup) {
			if (cboLegendTable.selectedIndex < 0) {
				cboLegendTable.selectedIndex = 0;
			}
			cboLegendColumn.length = 0;
			cboLegendCode.length = 0;
			txtLookupColumnsLoaded.value = 0;
		}

		populateLookupColumns();
		refreshLegendControls();
		eventChanged();
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
				sControlName = dataCollection.item(i).name;

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

	function validateEventInfo() {
		var sEventName = new String(trim(frmPopup.txtEventName.value));
		var sMessage = new String("");

		//check a name has been entered
		if (sEventName == '' || sEventName.length < 1) {
			sMessage = "You must give this event a name.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.txtEventName.focus();
			return false;
		}

		//check the name is unique
		if (!checkUniqueEventName(sEventName)) {
			sMessage = "An event called '" + sEventName + "' already exists.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.txtEventName.focus();
			return false;
		}

		//check that a valid event table has been selected
		if (frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].value < 1) {
			sMessage = "A valid event table has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboEventTable.focus();
			return false;
		}

		//check that a valid start date column has been selected
		if (frmPopup.cboStartDate.length < 1) {
			sMessage = "The selected event table has no date columns. Please select an event table that contains date columns.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.txtNoDateColumns.value = 1;
			if (frmPopup.cboStartDate.disabled == false) frmPopup.cboStartDate.focus();
			return false;
		}

		//check that a valid start date column has been selected
		if (frmPopup.cboStartDate.selectedIndex < 0) {
			sMessage = "A valid start date column has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			if (frmPopup.cboStartDate.disabled == false) frmPopup.cboStartDate.focus();
			return false;
		}

		//check that a valid start date column has been selected
		if (frmPopup.cboStartDate.options[frmPopup.cboStartDate.selectedIndex].value < 1) {
			sMessage = "A valid start date column has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			if (frmPopup.cboStartDate.disabled == false) frmPopup.cboStartDate.focus();
			return false;
		}

		//check that either a valid end date or duration column has been selected
		if (frmPopup.optDuration.checked && (frmPopup.cboDuration.options[frmPopup.cboDuration.selectedIndex].value < 1)) {
			sMessage = "A valid duration column has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboDuration.focus();
			return false;
		}
		if (frmPopup.optEndDate.checked && (frmPopup.cboEndDate.options[frmPopup.cboEndDate.selectedIndex].value < 1)) {
			sMessage = "A valid end date column has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboEndDate.focus();
			return false;
		}

		//check that a valid 'set' of lookup tables have been selected
		if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendTable.length < 1)) {
			sMessage = "A valid lookup table has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			return false;
		}
		else if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendTable.options[frmPopup.cboLegendTable.selectedIndex].value < 1)) {
			sMessage = "A valid lookup table has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboLegendTable.focus();
			return false;
		}

		if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendColumn.length < 1)) {
			sMessage = "A valid lookup column has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			return false;
		}
		else if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendColumn.options[frmPopup.cboLegendColumn.selectedIndex].value < 1)) {
			sMessage = "A valid lookup column has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboLegendColumn.focus();
			return false;
		}

		if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendCode.length < 1)) {
			sMessage = "A valid lookup code has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			return false;
		}
		else if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendCode.options[frmPopup.cboLegendCode.selectedIndex].value < 1)) {
			sMessage = "A valid lookup code has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboLegendCode.focus();
			return false;
		}

		if (frmPopup.optLegendLookup.checked && (frmPopup.cboEventType.length < 1)) {
			sMessage = "A valid event type has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			return false;
		}
		else if (frmPopup.optLegendLookup.checked && (frmPopup.cboEventType.options[frmPopup.cboEventType.selectedIndex].value < 1)) {
			sMessage = "A valid event type has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboEventType.focus();
			return false;
		}

		return true;
	}

	function checkUniqueEventName(psEventName) {
		//CODE REQUIRED TO CHECK THAT THE EVENT NAME IS UNIQUE 
		return true;
	}

	function selectRecordOption(psTable, psType) {
		if (psTable == 'event') {
			iTableID = frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].value;
			iCurrentID = frmPopup.txtEventFilterID.value;
		}

		frmRecordSelection.recSelTable.value = psTable;
		frmRecordSelection.recSelType.value = psType;
		frmRecordSelection.recSelTableID.value = iTableID;
		frmRecordSelection.recSelCurrentID.value = iCurrentID;

		var frmDef = OpenHR.getForm("workframe", "frmDefinition");
		var frmUse = OpenHR.getForm("workframe", "frmUseful");

		var strDefOwner = new String(frmDef.txtOwner.value);
		var strCurrentUser = new String(frmUse.txtUserName.value);

		strDefOwner = strDefOwner.toLowerCase();
		strCurrentUser = strCurrentUser.toLowerCase();

		if (strDefOwner == strCurrentUser) {
			frmRecordSelection.recSelDefOwner.value = '1';
		}
		else {
			frmRecordSelection.recSelDefOwner.value = '0';
		}

		sURL = "util_recordSelection" +
			"?recSelType=" + escape(frmRecordSelection.recSelType.value) +
			"&recSelTableID=" + escape(frmRecordSelection.recSelTableID.value) +
			"&recSelCurrentID=" + escape(frmRecordSelection.recSelCurrentID.value) +
			"&recSelTable=" + escape(frmRecordSelection.recSelTable.value) +
			"&recSelDefOwner=" + escape(frmRecordSelection.recSelDefOwner.value);
		openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");

		eventChanged();
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

	function refreshEventControls() {
		var frmUse = OpenHR.getForm("workframe", "frmUseful");
		var fViewing = (frmUse.txtAction.value.toUpperCase() == "VIEW");

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
		var frmUse = OpenHR.getForm("workframe", "frmUseful");
		var fViewing = (frmUse.txtAction.value.toUpperCase() == "VIEW");

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

		var frmUseful = OpenHR.getForm("workframe", "frmUseful");
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

								if (window.dialogArguments.window.isSelectedChildTable(iTableID)) {
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

	function cancelClick() {
		window.parent.close();
		return false;
	}

	function setForm() {
		if (!validateEventInfo()) {
			return false;
		}

		var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");
		var frmDef = OpenHR.getForm("workframe", "frmDefinition");

		var plngRow = frmDef.grdEvents.AddItemRowIndex(frmDef.grdEvents.Bookmark);
		var sADD = new String("");

		//Add the event information to string which will be used to populate the grid.
		with (frmPopup) {
			sADD = sADD + txtEventName.value + '	';
			sADD = sADD + cboEventTable.options[cboEventTable.selectedIndex].value + '	';
			sADD = sADD + cboEventTable.options[cboEventTable.selectedIndex].innerText + '	';
			sADD = sADD + txtEventFilterID.value + '	';
			sADD = sADD + txtEventFilter.value + '	';
			sADD = sADD + cboStartDate.options[cboStartDate.selectedIndex].value + '	';
			sADD = sADD + cboStartDate.options[cboStartDate.selectedIndex].innerText + '	';

			if (cboStartSession.selectedIndex < 0) {
				sADD = sADD + 0 + '	';
				sADD = sADD + '' + '	';
			}
			else {
				sADD = sADD + cboStartSession.options[cboStartSession.selectedIndex].value + '	';
				sADD = sADD + cboStartSession.options[cboStartSession.selectedIndex].innerText + '	';
			}

			if (cboEndDate.selectedIndex < 0) {
				sADD = sADD + 0 + '	';
				sADD = sADD + '' + '	';
			}
			else {
				sADD = sADD + cboEndDate.options[cboEndDate.selectedIndex].value + '	';
				sADD = sADD + cboEndDate.options[cboEndDate.selectedIndex].innerText + '	';
			}

			if (cboEndSession.selectedIndex < 0) {
				sADD = sADD + 0 + '	';
				sADD = sADD + '' + '	';
			}
			else {
				sADD = sADD + cboEndSession.options[cboEndSession.selectedIndex].value + '	';
				sADD = sADD + cboEndSession.options[cboEndSession.selectedIndex].innerText + '	';
			}

			if (cboDuration.selectedIndex < 0) {
				sADD = sADD + 0 + '	';
				sADD = sADD + '' + '	';
			}
			else {
				sADD = sADD + cboDuration.options[cboDuration.selectedIndex].value + '	';
				sADD = sADD + cboDuration.options[cboDuration.selectedIndex].innerText + '	';
			}

			if (optLegendLookup.checked == true) {
				sADD = sADD + '1' + '	';
				sADD = sADD + cboLegendTable.options[cboLegendTable.selectedIndex].innerText + '.' + cboLegendCode.options[cboLegendCode.selectedIndex].innerText + '	';
				sADD = sADD + cboLegendTable.options[cboLegendTable.selectedIndex].value + '	';
				sADD = sADD + cboLegendColumn.options[cboLegendColumn.selectedIndex].value + '	';
				sADD = sADD + cboLegendCode.options[cboLegendCode.selectedIndex].value + '	';
				sADD = sADD + cboEventType.options[cboEventType.selectedIndex].value + '	';
			}
			else {
				sADD = sADD + '0' + '	';
				sADD = sADD + txtCharacter.value + '	';
				sADD = sADD + '0' + '	';
				sADD = sADD + '0' + '	';
				sADD = sADD + '0' + '	';
				sADD = sADD + '0' + '	';
			}

			sADD = sADD + cboEventDesc1.options[cboEventDesc1.selectedIndex].value + '	';
			sADD = sADD + cboEventDesc1.options[cboEventDesc1.selectedIndex].innerText + '	';
			sADD = sADD + cboEventDesc2.options[cboEventDesc2.selectedIndex].value + '	';
			sADD = sADD + cboEventDesc2.options[cboEventDesc2.selectedIndex].innerText + '	';
		}

		sADD = sADD + frmEvent.eventID.value + '	';
		sADD = sADD + frmSelectionAccess.baseHidden.value;

		//Add the event information to the grdEvents in the parent window..
		if (frmEvent.eventAction.value.toUpperCase() == "NEW") {
			frmDef.grdEvents.additem(sADD);
			frmDef.grdEvents.selbookmarks.RemoveAll();
			frmDef.grdEvents.MoveLast();
			frmDef.grdEvents.selbookmarks.Add(frmDef.grdEvents.Bookmark);
		}
		else {
			//' Check if any columns in the report definition are from the table that was
			//' previously selected in the child combo box. If so, prompt user for action.
			var bContinueRemoval;

			//bContinueRemoval = removeChildTable(frmPopup.originalChildID.value);
			bContinueRemoval = true;

			if (bContinueRemoval) {
				frmDef.grdEvents.removeitem(plngRow);
				frmDef.grdEvents.additem(sADD, plngRow);
				frmDef.grdEvents.Bookmark = frmDef.grdEvents.AddItemBookmark(plngRow);
				frmDef.grdEvents.SelBookmarks.RemoveAll();
				frmDef.grdEvents.SelBookmarks.Add(frmDef.grdEvents.AddItemBookmark(plngRow));
			}
		}

		self.close();
		return false;
	}

	function eventChanged() {
		var frmUse = OpenHR.getForm("workframe", "frmUseful");
		var fViewing = (frmUse.txtAction.value.toUpperCase() == "VIEW");
		button_disable(frmPopup.cmdOK, fViewing);
	}

</script>

<div id="bdyMain" name="bdyMain" <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
	<form id="frmPopup" name="frmPopup" onsubmit="return setForm();">
		<table align="center" width="100%" height="100%" class="outline" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<table align="center" width="100%" height="100%" class="invisible" cellpadding="4" cellspacing="0">
						<tr height="5">
							<td align="center" colspan="2" height="10">
								<font size="3"><B>Select Event Information</B></font>
							</td>
						</tr>
						<tr>
							<td valign="top" width="50%">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
									<tr height="10">
										<td height="10" colspan="5" align="left" valign="top">Event :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Name :</td>
													<td width="5"></td>
													<td>
														<input id="txtEventName" name="txtEventName" class="text textdisabled" style="WIDTH: 100%" disabled="disabled"
															onkeypress="eventChanged();"
															onkeydown="eventChanged();"
															onchange="eventChanged();">
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Event Table :</td>
													<td width="5"></td>
													<td>
														<select id="cboEventTable" name="cboEventTable" class="combo combodisabled" disabled="disabled" style="WIDTH: 100%"
															onchange="changeEventTable();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Filter :</td>
													<td width="5"></td>
													<td>
														<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
															<tr>
																<td>
																	<input id="txtEventFilter" name="txtEventFilter" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
																	<input type="hidden" id="txtEventFilterID" name="txtEventFilterID" class="text textdisabled" disabled="disabled" style="WIDTH: 100%"
																		onchange="eventChanged();">
																</td>
																<td width="25">
																	<input id="cmdEventFilter" name="cmdEventFilter" disabled="disabled" class="btn btndisabled" style="WIDTH: 100%" type="button" value="..."
																		onclick="selectRecordOption('event', 'filter')"
																		onmouseover="try{button_onMouseOver(this);}catch(e){}"
																		onmouseout="try{button_onMouseOut(this);}catch(e){}"
																		onfocus="try{button_onFocus(this);}catch(e){}"
																		onblur="try{button_onBlur(this);}catch(e){}" />
																</td>
															</tr>
														</table>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
							<td valign="top" rowspan="2" width="50%">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
									<tr height="10">
										<td height="10" colspan="5" align="left" valign="top">Key :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optKey" id="optCharacter" disabled="disabled"
															onclick="refreshLegendControls(); eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}" />&nbsp;
													</td>
													<td nowrap colspan="1">
														<label
															tabindex="-1"
															for="optCharacter"
															class="radio radiodisabled"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															Character</label>
													</td>
													<td width="5"></td>
													<td nowrap width="100%">
														<input id="txtCharacter" maxlength="2" name="txtCharacter" class="text textdisabled" disabled="disabled" style="WIDTH: 60px"
															onkeypress="eventChanged();"
															onkeydown="eventChanged();"
															onchange="eventChanged();">
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optKey" id="optLegendLookup" disabled="disabled"
															onclick="refreshLegendControls(); eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}" />
													</td>
													<td nowrap width="100" colspan="3">
														<label
															tabindex="-1"
															for="optLegendLookup"
															class="radio radiodisabled"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															Lookup Table</label>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td>&nbsp</td>
													<td width="100" nowrap>Event Type : 
													</td>
													<td width="5"></td>
													<td>
														<select id="cboEventType" name="cboEventType" disabled="disabled" width="100%" style="WIDTH: 100%" class="combo combodisabled"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td colspan="4">
														<hr>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap></td>
													<td width="100" nowrap>Table :
													</td>
													<td width="5"></td>
													<td>
														<select id="cboLegendTable" name="cboLegendTable" disabled="disabled" class="combo combodisabled" style="WIDTH: 100%"
															onchange="changeLegendTable();">
															<%
																Dim sErrorDescription = ""

																' Get the lookup table records.
																Dim cmdLookupTables = CreateObject("ADODB.Command")
																cmdLookupTables.CommandText = "spASRIntGetLookupTables"
																cmdLookupTables.CommandType = 4	' Stored Procedure
																cmdLookupTables.ActiveConnection = Session("databaseConnection")
	
																Err.Number = 0
																Dim rstLookupTablesInfo = cmdLookupTables.Execute
	
																If (Err.Number <> 0) Then
																	sErrorDescription = "The lookup tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
																End If

																If Len(sErrorDescription) = 0 Then
																	Dim iCount = 0
																	Do While Not rstLookupTablesInfo.EOF
																		Response.Write("<OPTION value='" & rstLookupTablesInfo.fields("tableID").value & "'>" & rstLookupTablesInfo.fields("tableName").value & vbCrLf)
																		rstLookupTablesInfo.MoveNext()
																	Loop

																	' Release the ADO recordset object.
																	rstLookupTablesInfo.close()
																	rstLookupTablesInfo = Nothing
																End If
	
																' Release the ADO command object.
																cmdLookupTables = Nothing
															%>
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap></td>
													<td width="100" nowrap>Column :								
													</td>
													<td width="5"></td>
													<td>
														<select id="cboLegendColumn" name="cboLegendColumn" width="100%" style="WIDTH: 100%" class="combo"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td></td>
													<td width="100" nowrap>Code : 
													</td>
													<td width="5"></td>
													<td>
														<select id="cboLegendCode" name="cboLegendCode" width="100%" style="WIDTH: 100%" class="combo"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td valign="top" width="50%">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
									<tr height="5">
										<td height="5" colspan="5" align="left" valign="top">Event Start :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Start Date :</td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboStartDate" name="cboStartDate" class="combo combodisabled" style="WIDTH: 100%"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Start Session :</td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboStartSession" name="cboStartSession" class="combo combodisabled" style="WIDTH: 100%"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td valign="top" width="50%" rowspan="2">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
									<tr height="10">
										<td height="10" colspan="5" rowspan="2" align="left" valign="top">Event End :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optEnd" id="optNoEnd"
															onclick="refreshEventControls(); eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}" />&nbsp;
													</td>
													<td nowrap colspan="3">
														<label
															tabindex="-1"
															for="optNoEnd"
															class="radio"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															None</label>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optEnd" id="optEndDate"
															onclick="refreshEventControls(); eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}" />
													</td>
													<td nowrap colspan="3">
														<label tabindex="-1"
															for="optEndDate"
															class="radio"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															End</label>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td width="5"></td>
													<td nowrap width="65">Date : </td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboEndDate" name="cboEndDate" style="WIDTH: 100%" class="combo combodisabled"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td width="5"></td>
													<td nowrap width="65">Session : </td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboEndSession" name="cboEndSession" style="WIDTH: 100%" class="combo combodisabled"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optEnd" id="optDuration"
															onclick="refreshEventControls();"
															onchange="eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}"></td>
													<td nowrap width="65">
														<label
															tabindex="-1"
															for="optDuration"
															class="radio"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															Duration</label>
													</td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboDuration" name="cboDuration" style="WIDTH: 100%" class="combo combodisabled"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
							<td valign="top">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%">
									<tr height="10">
										<td height="10" colspan="5" align="left">Event Description :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="10">
													<td width="5"></td>
													<td nowrap width="100">Description 1 : </td>
													<td width="5">&nbsp</td>
													<td>
														<select disabled="disabled" id="cboEventDesc1" name="cboEventDesc1" width="100%" class="combo combodisabled" style="WIDTH: 100%"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="5"></td>
												</tr>
												<tr height="10">
													<td width="5"></td>
													<td nowrap width="100">Description 2 : </td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboEventDesc2" name="cboEventDesc2" width="100%" class="combo combodisabled" style="WIDTH: 100%"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td valign="bottom" align="right">
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td>&nbsp;</td>
										<td width="10">
											<input id="cmdOK" type="button" value="OK" name="cmdOK" class="btn" style="WIDTH: 80px" width="80"
												onclick="setForm()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="10">&nbsp;</td>
										<td width="10">
											<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel" class="btn" style="WIDTH: 80px" width="80"
												onclick="cancelClick();"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr height="5">
							<td colspan="5"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>

		<input type="hidden" id="txtLookupColumnsLoaded" name="txtLookupColumnsLoaded">
		<input type="hidden" id="txtEventColumnsLoaded" name="txtEventColumnsLoaded">
		<input type="hidden" id="txtFirstLoad_Lookup" name="txtFirstLoad_Lookup">
		<input type="hidden" id="txtFirstLoad_Event" name="txtFirstLoad_Event">
		<input type="hidden" id="txtHaveSetLookupValues" name="txtHaveSetLookupValues">
		<input type="hidden" id="txtLoading" name="txtLoading">
		<input type="hidden" id="rowID" name="rowID" value="-1">
		<input type="hidden" id="txtNoDateColumns" name="txtNoDateColumns" value="0">
	</form>

	<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="recSelType" name="recSelType">
		<input type="hidden" id="recSelTableID" name="recSelTableID">
		<input type="hidden" id="recSelCurrentID" name="recSelCurrentID">
		<input type="hidden" id="recSelTable" name="recSelTable">
		<input type="hidden" id="recSelDefOwner" name="recSelDefOwner">
	</form>

	<form id="frmSelectionAccess" name="frmSelectionAccess" style="visibility: hidden; display: none">
		<input type="hidden" id="baseHidden" name="baseHidden" value='N'>
	</form>

</div>

<script type="text/javascript">
	util_def_calendarreportdates_window_onload();
</script>
