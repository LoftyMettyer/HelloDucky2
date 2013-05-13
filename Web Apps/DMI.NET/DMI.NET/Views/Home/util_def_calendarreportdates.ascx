<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">
	function util_def_calendarreportdates_window_onload() {
		
		var frmPopup = document.getElementById("frmPopup");
		var frmSelectionAccess = document.getElementById("frmSelectionAccess");
		frmPopup.txtLoading.value = 1;
		frmPopup.txtFirstLoad_Event.value = 1;
		frmPopup.txtFirstLoad_Lookup.value = 1;
		frmPopup.txtHaveSetLookupValues.value = 0;

		button_disable(frmPopup.cmdCancel, true);
		button_disable(frmPopup.cmdOK, true);

		populateEventTableCombo();

		frmPopup.cboLegendTable.selectedIndex = -1;

		var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");
		var frmDef = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");

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


	function validateEventInfo() {
		var frmPopup = document.getElementById("frmPopup");
		var sEventName = new String(trim(frmPopup.txtEventName.value));
		var sMessage = new String("");

		//check a name has been entered
		if (sEventName == '' || sEventName.length < 1) {
			sMessage = "You must give this event a name.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.txtEventName.focus();
			return (false);
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
		} else if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendTable.options[frmPopup.cboLegendTable.selectedIndex].value < 1)) {
			sMessage = "A valid lookup table has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboLegendTable.focus();
			return false;
		}

		if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendColumn.length < 1)) {
			sMessage = "A valid lookup column has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			return false;
		} else if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendColumn.options[frmPopup.cboLegendColumn.selectedIndex].value < 1)) {
			sMessage = "A valid lookup column has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboLegendColumn.focus();
			return false;
		}

		if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendCode.length < 1)) {
			sMessage = "A valid lookup code has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			return false;
		} else if (frmPopup.optLegendLookup.checked && (frmPopup.cboLegendCode.options[frmPopup.cboLegendCode.selectedIndex].value < 1)) {
			sMessage = "A valid lookup code has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboLegendCode.focus();
			return false;
		}

		if (frmPopup.optLegendLookup.checked && (frmPopup.cboEventType.length < 1)) {
			sMessage = "A valid event type has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			return false;
		} else if (frmPopup.optLegendLookup.checked && (frmPopup.cboEventType.options[frmPopup.cboEventType.selectedIndex].value < 1)) {
			sMessage = "A valid event type has not been selected.";
			OpenHR.messageBox(sMessage, 48, "Calendar Reports");
			frmPopup.cboEventType.focus();
			return false;
		}

		return true;
	}

	function setForm() {
		if (!validateEventInfo()) {
			//self.close();
			return false;
		}
		var frmSelectionAccess = document.getElementById("frmSelectionAccess");
		var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");
		var frmDef = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
		var frmPopup = document.getElementById("frmPopup");

		var plngRow = frmDef.grdEvents.AddItemRowIndex(frmDef.grdEvents.Bookmark);
		var sADD = new String("");

		//Add the event information to string which will be used to populate the grid.
		sADD = sADD + frmPopup.txtEventName.value + '	';
		sADD = sADD + frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].value + '	';
		sADD = sADD + frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].innerText + '	';
		sADD = sADD + frmPopup.txtEventFilterID.value + '	';
		sADD = sADD + frmPopup.txtEventFilter.value + '	';
		sADD = sADD + frmPopup.cboStartDate.options[frmPopup.cboStartDate.selectedIndex].value + '	';
		sADD = sADD + frmPopup.cboStartDate.options[frmPopup.cboStartDate.selectedIndex].innerText + '	';

		if (frmPopup.cboStartSession.selectedIndex < 0) {
			sADD = sADD + 0 + '	';
			sADD = sADD + '' + '	';
		} else {
			sADD = sADD + frmPopup.cboStartSession.options[frmPopup.cboStartSession.selectedIndex].value + '	';
			sADD = sADD + frmPopup.cboStartSession.options[frmPopup.cboStartSession.selectedIndex].innerText + '	';
		}

		if (frmPopup.cboEndDate.selectedIndex < 0) {
			sADD = sADD + 0 + '	';
			sADD = sADD + '' + '	';
		} else {
			sADD = sADD + frmPopup.cboEndDate.options[frmPopup.cboEndDate.selectedIndex].value + '	';
			sADD = sADD + frmPopup.cboEndDate.options[frmPopup.cboEndDate.selectedIndex].innerText + '	';
		}

		if (frmPopup.cboEndSession.selectedIndex < 0) {
			sADD = sADD + 0 + '	';
			sADD = sADD + '' + '	';
		} else {
			sADD = sADD + frmPopup.cboEndSession.options[frmPopup.cboEndSession.selectedIndex].value + '	';
			sADD = sADD + frmPopup.cboEndSession.options[frmPopup.cboEndSession.selectedIndex].innerText + '	';
		}

		if (frmPopup.cboDuration.selectedIndex < 0) {
			sADD = sADD + 0 + '	';
			sADD = sADD + '' + '	';
		} else {
			sADD = sADD + frmPopup.cboDuration.options[frmPopup.cboDuration.selectedIndex].value + '	';
			sADD = sADD + frmPopup.cboDuration.options[frmPopup.cboDuration.selectedIndex].innerText + '	';
		}

		if (frmPopup.optLegendLookup.checked == true) {
			sADD = sADD + '1' + '	';
			sADD = sADD + frmPopup.cboLegendTable.options[frmPopup.cboLegendTable.selectedIndex].innerText + '.' + frmPopup.cboLegendCode.options[frmPopup.cboLegendCode.selectedIndex].innerText + '	';
			sADD = sADD + frmPopup.cboLegendTable.options[frmPopup.cboLegendTable.selectedIndex].value + '	';
			sADD = sADD + frmPopup.cboLegendColumn.options[frmPopup.cboLegendColumn.selectedIndex].value + '	';
			sADD = sADD + frmPopup.cboLegendCode.options[frmPopup.cboLegendCode.selectedIndex].value + '	';
			sADD = sADD + frmPopup.cboEventType.options[frmPopup.cboEventType.selectedIndex].value + '	';
		} else {
			sADD = sADD + '0' + '	';
			sADD = sADD + frmPopup.txtCharacter.value + '	';
			sADD = sADD + '0' + '	';
			sADD = sADD + '0' + '	';
			sADD = sADD + '0' + '	';
			sADD = sADD + '0' + '	';
		}

		sADD = sADD + frmPopup.cboEventDesc1.options[frmPopup.cboEventDesc1.selectedIndex].value + '	';
		sADD = sADD + frmPopup.cboEventDesc1.options[frmPopup.cboEventDesc1.selectedIndex].innerText + '	';
		sADD = sADD + frmPopup.cboEventDesc2.options[frmPopup.cboEventDesc2.selectedIndex].value + '	';
		sADD = sADD + frmPopup.cboEventDesc2.options[frmPopup.cboEventDesc2.selectedIndex].innerText + '	';

		sADD = sADD + frmEvent.eventID.value + '	';
		sADD = sADD + frmSelectionAccess.baseHidden.value;

		//Add the event information to the grdEvents in the parent window..
		if (frmEvent.eventAction.value.toUpperCase() == "NEW") {
			frmDef.grdEvents.additem(sADD);
			frmDef.grdEvents.selbookmarks.RemoveAll();
			frmDef.grdEvents.MoveLast();
			frmDef.grdEvents.selbookmarks.Add(frmDef.grdEvents.Bookmark);
		} else {
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
		return true;
	}

	function populateEventTableCombo() {
		//var frmTab = document.parentWindow.parent.window.dialogArguments.parent.frames("workframe").document.forms("frmTables");
		//var frmDef = document.parentWindow.parent.window.dialogArguments.parent.frames("workframe").document.forms("frmDefinition");
		//var frmEvent = document.parentWindow.parent.window.dialogArguments.parent.frames("workframe").document.forms("frmEventDetails");
		
		var frmTab = document.parentWindow.parent.window.dialogArguments.parent.document.getElementById("frmTables");
		var frmDef = document.parentWindow.parent.window.dialogArguments.parent.document.getElementById("frmDefinition");
		var frmEvent = document.parentWindow.parent.window.dialogArguments.parent.document.getElementById("frmEventDetails");

		var frmPopup = document.getElementById("frmPopup");
		var sRelationString = frmEvent.relationNames.value;		

		var iRelationID;
		var sTableName;

		var bAdded = false;

		var dataCollection = frmTab.elements;
		var oOption;

		var frmRefresh = document.parentWindow.parent.window.dialogArguments.parent.document.getElementById("frmHit");

		var iIndex = sRelationString.indexOf("	");
		while (iIndex > 0) {
			iRelationID = sRelationString.substr(0, iIndex);

			//frmRefresh.submit();

			OpenHR.submitForm(frmRefresh);
			//console.log("refreshing...");
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

	function populateEventColumns() {
		// Get the columns/calcs for the current table selection.
		var frmGetDataForm = OpenHR.getForm("calendardataframe", "frmGetCalendarData");
		var frmDef = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
		var frmPopup = document.getElementById("frmPopup");

		frmGetDataForm.txtCalendarAction.value = "LOADCALENDAREVENTDETAILSCOLUMNS";
		frmGetDataForm.txtCalendarBaseTableID.value = frmDef.cboBaseTable.options[frmDef.cboBaseTable.selectedIndex].value;
		frmGetDataForm.txtCalendarEventTableID.value = frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].value;

		//window.parent.frames("calendardataframe").refreshData();
		data_refreshData();
	}

	function selectRecordOptionDates(psTable, psType) {
		var iTableID;
		var iCurrentID;
		var sURL;
		var frmPopup = document.getElementById("frmPopup");
		var frmRecordSelection = document.getElementById("frmRecordSelection");
		var frmGetDataForm = document.getElementById("frmGetCalendarData");

		var frmDef = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
		var frmUse = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");

		if (psTable == 'event') {
			iTableID = frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].value;
			iCurrentID = frmPopup.txtEventFilterID.value;
		}

		frmRecordSelection.recSelTable.value = psTable;
		frmRecordSelection.recSelType.value = psType;
		frmRecordSelection.recSelTableID.value = iTableID;
		frmRecordSelection.recSelCurrentID.value = iCurrentID;

		var strDefOwner = new String(frmDef.txtOwner.value);
		var strCurrentUser = new String(frmUse.txtUserName.value);
		strDefOwner = strDefOwner.toLowerCase();
		strCurrentUser = strCurrentUser.toLowerCase();

		if (strDefOwner == strCurrentUser) {
			frmRecordSelection.recSelDefOwner.value = '1';
		} else {
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

	function changeEventTable() {
		var frmPopup = document.getElementById("frmPopup");
		{
			frmPopup.txtEventFilterID.value = 0;
			frmPopup.txtEventFilter.value = "";
			frmPopup.cboStartDate.length = 0;
			frmPopup.cboStartSession.length = 0;
			frmPopup.cboEndDate.length = 0;
			frmPopup.cboEndSession.length = 0;
			frmPopup.cboDuration.length = 0;
			frmPopup.cboEventType.length = 0;
			frmPopup.cboEventDesc1.length = 0;
			frmPopup.cboEventDesc2.length = 0;
			frmPopup.txtEventColumnsLoaded.value = 0;

			populateEventColumns();
			refreshEventControls();
			eventChanged();
		}
	}

	function changeLegendTable() {
		var frmPopup = document.getElementById("frmPopup");
		if (frmPopup.cboLegendTable.selectedIndex < 0) {
			frmPopup.cboLegendTable.selectedIndex = 0;
		}
		frmPopup.cboLegendColumn.length = 0;
		frmPopup.cboLegendCode.length = 0;
		frmPopup.txtLookupColumnsLoaded.value = 0;
		populateLookupColumns();
		refreshLegendControls();
		eventChanged();
	}

	function eventChanged() {
		var frmUse = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");
		var frmPopup = document.getElementById("frmPopup");
		var fViewing = (frmUse.txtAction.value.toUpperCase() == "VIEW");
		button_disable(frmPopup.cmdOK, fViewing);
	}

	function disabledAll() {
		var frmPopup = document.getElementById("frmPopup");
		/*Event Frame*/
		text_disable(frmPopup.txtEventName, true);
		combo_disable(frmPopup.cboEventTable, true);
		text_disable(frmPopup.txtEventFilter, true);
		button_disable(frmPopup.cmdEventFilter, true);

		/*Event Start Frame*/
		combo_disable(frmPopup.cboStartDate, true);
		combo_disable(frmPopup.cboStartSession, true);

		/*Event End Frame*/
		radio_disable(frmPopup.optNoEnd, true);
		radio_disable(frmPopup.optEndDate, true);
		combo_disable(frmPopup.cboEndDate, true);
		combo_disable(frmPopup.cboEndSession, true);
		radio_disable(frmPopup.optDuration, true);
		combo_disable(frmPopup.cboDuration, true);

		/*Key Frame*/
		text_disable(frmPopup.txtCharacter, true);
		combo_disable(frmPopup.cboEventType, true);
		combo_disable(frmPopup.cboLegendTable, true);
		combo_disable(frmPopup.cboLegendColumn, true);
		combo_disable(frmPopup.cboLegendCode, true);

		/*Event Description Frame*/
		combo_disable(frmPopup.cboEventDesc1, true);
		combo_disable(frmPopup.cboEventDesc2, true);
	}
</script>

<div id="bdyMain" name="bdyMain" <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
	<form id="frmPopup" name="frmPopup" onsubmit=" return setForm(); ">
		<table align="center" width="100%" height="100%" class="outline" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<table align="center" width="100%" height="100%" class="invisible" cellpadding="4" cellspacing="0">
						<tr height="5">
							<td align="center" colspan="2" height="10">
								Select Event Information
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
														       onkeypress=" eventChanged(); "
														       onkeydown=" eventChanged(); "
														       onchange=" eventChanged(); ">
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
														        onchange=" changeEventTable(); ">
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
																	       onchange=" eventChanged(); ">
																</td>
																<td width="25">
																	<input id="cmdEventFilter" name="cmdEventFilter" disabled="disabled" class="btn btndisabled" style="WIDTH: 100%" type="button" value="..."
																	       onclick=" selectRecordOptionDates('event', 'filter') "
																	       onmouseover=" try {button_onMouseOver(this);} catch(e) {} "
																	       onmouseout=" try {button_onMouseOut(this);} catch(e) {} "
																	       onfocus=" try {button_onFocus(this);} catch(e) {} "
																	       onblur=" try {button_onBlur(this);} catch(e) {} " />
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
														       onclick=" refreshLegendControls(); eventChanged(); "
														       onmouseover=" try {radio_onMouseOver(this);} catch(e) {} "
														       onmouseout=" try {radio_onMouseOut(this);} catch(e) {} "
														       onfocus=" try {radio_onFocus(this);} catch(e) {} "
														       onblur=" try {radio_onBlur(this);} catch(e) {} " />&nbsp;
													</td>
													<td nowrap colspan="1">
														<label
															tabindex="-1"
															for="optCharacter"
															class="radio radiodisabled"
															onmouseover=" try {radioLabel_onMouseOver(this);} catch(e) {} "
															onmouseout=" try {radioLabel_onMouseOut(this);} catch(e) {} ">
															Character</label>
													</td>
													<td width="5"></td>
													<td nowrap width="100%">
														<input id="txtCharacter" maxlength="2" name="txtCharacter" class="text textdisabled" disabled="disabled" style="WIDTH: 60px"
														       onkeypress=" eventChanged(); "
														       onkeydown=" eventChanged(); "
														       onchange=" eventChanged(); ">
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
														       onclick=" refreshLegendControls(); eventChanged(); "
														       onmouseover=" try {radio_onMouseOver(this);} catch(e) {} "
														       onmouseout=" try {radio_onMouseOut(this);} catch(e) {} "
														       onfocus=" try {radio_onFocus(this);} catch(e) {} "
														       onblur=" try {radio_onBlur(this);} catch(e) {} " />
													</td>
													<td nowrap width="100" colspan="3">
														<label
															tabindex="-1"
															for="optLegendLookup"
															class="radio radiodisabled"
															onmouseover=" try {radioLabel_onMouseOver(this);} catch(e) {} "
															onmouseout=" try {radioLabel_onMouseOut(this);} catch(e) {} ">
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
														        onchange=" eventChanged(); ">
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
														        onchange=" changeLegendTable(); ">
															
															<%
																Dim sErrorDescription = ""

																' Get the lookup table records.
																Dim cmdLookupTables = CreateObject("ADODB.Command")
																cmdLookupTables.CommandText = "spASRIntGetLookupTables"
																cmdLookupTables.CommandType = 4	' Stored Procedure
																cmdLookupTables.ActiveConnection = Session("databaseConnection")

																Err.Clear()
																Dim rstLookupTablesInfo = cmdLookupTables.Execute

																If (Err.Number <> 0) Then
																	sErrorDescription = "The lookup tables information could not be retrieved." & vbCrLf &
																	                    FormatError(Err.Description)
																End If

																If Len(sErrorDescription) = 0 Then
																	Dim iCount = 0
																	Do While Not rstLookupTablesInfo.EOF
																		Response.Write(
																			"<OPTION value='" & rstLookupTablesInfo.fields("tableID").value & "'>" &
																			rstLookupTablesInfo.fields("tableName").value & vbCrLf)
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
														        onchange=" eventChanged(); ">
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
														        onchange=" eventChanged(); ">
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
														        onchange=" eventChanged(); ">
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
														        onchange=" eventChanged(); ">
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
														       onclick=" refreshEventControls(); eventChanged(); "
														       onmouseover=" try {radio_onMouseOver(this);} catch(e) {} "
														       onmouseout=" try {radio_onMouseOut(this);} catch(e) {} "
														       onfocus=" try {radio_onFocus(this);} catch(e) {} "
														       onblur=" try {radio_onBlur(this);} catch(e) {} " />&nbsp;
													</td>
													<td nowrap colspan="3">
														<label
															tabindex="-1"
															for="optNoEnd"
															class="radio"
															onmouseover=" try {radioLabel_onMouseOver(this);} catch(e) {} "
															onmouseout=" try {radioLabel_onMouseOut(this);} catch(e) {} ">
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
														       onclick=" refreshEventControls(); eventChanged(); "
														       onmouseover=" try {radio_onMouseOver(this);} catch(e) {} "
														       onmouseout=" try {radio_onMouseOut(this);} catch(e) {} "
														       onfocus=" try {radio_onFocus(this);} catch(e) {} "
														       onblur=" try {radio_onBlur(this);} catch(e) {} " />
													</td>
													<td nowrap colspan="3">
														<label tabindex="-1"
														       for="optEndDate"
														       class="radio"
														       onmouseover=" try {radioLabel_onMouseOver(this);} catch(e) {} "
														       onmouseout=" try {radioLabel_onMouseOut(this);} catch(e) {} ">
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
														        onchange=" eventChanged(); ">
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
														        onchange=" eventChanged(); ">
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
														       onclick=" refreshEventControls(); "
														       onchange=" eventChanged(); "
														       onmouseover=" try {radio_onMouseOver(this);} catch(e) {} "
														       onmouseout=" try {radio_onMouseOut(this);} catch(e) {} "
														       onfocus=" try {radio_onFocus(this);} catch(e) {} "
														       onblur=" try {radio_onBlur(this);} catch(e) {} "></td>
													<td nowrap width="65">
														<label
															tabindex="-1"
															for="optDuration"
															class="radio"
															onmouseover=" try {radioLabel_onMouseOver(this);} catch(e) {} "
															onmouseout=" try {radioLabel_onMouseOut(this);} catch(e) {} ">
															Duration</label>
													</td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboDuration" name="cboDuration" style="WIDTH: 100%" class="combo combodisabled"
														        onchange=" eventChanged(); ">
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
														        onchange=" eventChanged(); ">
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
														        onchange=" eventChanged(); ">
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
											       onclick=" setForm() "
											       onmouseover=" try {button_onMouseOver(this);} catch(e) {} "
											       onmouseout=" try {button_onMouseOut(this);} catch(e) {} "
											       onfocus=" try {button_onFocus(this);} catch(e) {} "
											       onblur=" try {button_onBlur(this);} catch(e) {} "/>
										</td>
										<td width="10">&nbsp;</td>
										<td width="10">
											<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel" class="btn" style="WIDTH: 80px" width="80"
												onclick="self.close();"
												onmouseover=" try {button_onMouseOver(this);} catch(e) {} "
												onmouseout=" try {button_onMouseOut(this);} catch(e) {} "
												onfocus=" try {button_onFocus(this);} catch(e) {} "
												onblur=" try {button_onBlur(this);} catch(e) {} " />
											<%--onclick=" cancelClick(); "--%>
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