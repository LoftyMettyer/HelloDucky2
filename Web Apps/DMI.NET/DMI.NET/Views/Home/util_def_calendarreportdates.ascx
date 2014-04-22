<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>

<script type="text/javascript">
	function util_def_calendarreportdates_window_onload() {

		var frmPopup = document.getElementById("frmPopup");
		var frmSelectionAccess = document.getElementById("frmSelectionAccess");
		frmPopup.txtLoading.value = 1;
		frmPopup.txtFirstLoad_Event.value = 1;
		frmPopup.txtFirstLoad_Lookup.value = 1;
		frmPopup.txtHaveSetLookupValues.value = 0;

		$(".button").button();

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
		//$('table').attr("border", "black solid 1px");
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
		openDialog(sURL, (screen.width) / 3 + 40, (screen.height) / 2.1, "no", "no");

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


<!DOCTYPE html>
<html>
<head>
	<title>Event Log Selection - OpenHR</title>
	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<script id="officebarscript" src="<%: Url.LatestContent("~/Scripts/officebar/jquery.officebar.js")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<%--<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />--%>
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />
</head>
<body>
<div id="bdyMain" name="bdyMain" <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
	<form id="frmPopup" name="frmPopup" onsubmit=" return setForm(); ">
		<div style="width: 95%; padding: 20px">
			<table class="outline"
				style="border-spacing: 0; border-collapse: collapse; width: 97%; height: 100%" id="Event End">
				<tr>
					<td style="text-align: center" colspan="6"><h3>Event Information</h3></td>
				</tr>
				<tr style="font-weight: bold">
					<td colspan="6">Event</td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td width="5"></td>
					<td nowrap width="100">Name :</td>
					<td width="5"></td>
					<td style="width: 99.5%">
						<input id="txtEventName" name="txtEventName" class="text textdisabled" style="width: 99%" disabled="disabled"
							onkeypress=" eventChanged(); "
							onkeydown=" eventChanged(); "
							onchange=" eventChanged(); ">
					</td>
					<td width="5"></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td width="5"></td>
					<td nowrap width="100">Event Table :</td>
					<td width="5"></td>
					<td>
						<select id="cboEventTable" name="cboEventTable" class="combo combodisabled" disabled="disabled" style="width: 100%"
							onchange=" changeEventTable(); ">
						</select>
					</td>
					<td width="5"></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td width="5"></td>
					<td nowrap width="100">Filter :</td>
					<td width="5"></td>
					<td>
						<input id="txtEventFilter" name="txtEventFilter" class="text textdisabled" disabled="disabled" style="width: 100%">
						<input type="hidden" id="txtEventFilterID" name="txtEventFilterID" class="text textdisabled" disabled="disabled" style="width: 100%"
							onchange=" eventChanged(); ">
					</td>
					<td width="25">
						<input id="cmdEventFilter" name="cmdEventFilter" disabled="disabled" class="btn " style="width: 100%" type="button" value="..."
							onclick=" selectRecordOptionDates('event', 'filter') " />
					</td>
				</tr>
				<tr style="font-weight: bold; padding-top: 20px">
					<td style="padding-top: 20px" colspan="6">Event Start</td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td style="width: 20px"></td>
					<td style="white-space: nowrap">Date :</td>
					<td width="5"></td>
					<td>
						<select disabled="disabled" id="cboStartDate" name="cboStartDate" class="combo combodisabled"
							style="width: 100%"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td style="width: 20px"></td>
					<td style="white-space: nowrap">Session :</td>
					<td width="5"></td>
					<td>
						<select disabled="disabled" id="cboStartSession" name="cboStartSession" class="combo combodisabled"
							style="width: 100%"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td></td>
				</tr>
				<tr style="font-weight: bold">
					<td colspan="6">Event End</td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td colspan="1">
						<input type="radio" name="optEnd" id="optEndDate"
							onclick=" refreshEventControls(); eventChanged(); " />
					</td>
					<td style="white-space: nowrap; padding-left: 2px">
						<label tabindex="-1"
							for="optEndDate"
							class="radio">
							End</label>
					</td>
					<td width="5" colspan="2"></td>
					<td></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td width="5"></td>
					<td style="white-space: nowrap; padding-left: 2px">Date : </td>
					<td width="5"></td>
					<td>
						<select disabled="disabled" id="cboEndDate" name="cboEndDate" style="width: 100%" class="combo combodisabled"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td width="5"></td>
					<td style="white-space: nowrap; padding-left: 2px">Session : </td>
					<td width="5"></td>
					<td>
						<select disabled="disabled" id="cboEndSession" name="cboEndSession" style="width: 100%" class="combo combodisabled"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td colspan="1">
						<input type="radio" name="optEnd" id="optDuration"
							onclick=" refreshEventControls(); "
							onchange=" eventChanged(); "></td>
					<td style="white-space: nowrap; padding-left: 2px">
						<label
							tabindex="-1"
							for="optDuration"
							class="radio">
							Duration</label>
					</td>
					<td width="5"></td>
					<td>
						<select disabled="disabled" id="cboDuration" name="cboDuration"
							style="width: 100%"
							class="combo combodisabled"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td>
						<input type="radio" name="optEnd" id="optNoEnd"
							onclick=" refreshEventControls(); eventChanged(); " />
					</td>
					<td style="white-space: nowrap; padding-left: 2px">
						<label
							tabindex="-1"
							for="optNoEnd"
							class="radio">
							None</label>
					</td>
					<td width="5" colspan="2"></td>
					<td></td>
				</tr>
				<tr style="font-weight: bold; padding-top: 20px">
					<td style="padding-top: 20px" colspan="6">Key</td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td colspan="1">
						<input type="radio" name="optKey" id="optCharacter" disabled="disabled"
							onclick=" refreshLegendControls(); eventChanged(); " />
					</td>
					<td nowrap colspan="1">
						<label
							tabindex="-1"
							for="optCharacter"
							class="radio radiodisabled">
							Character</label>
					</td>
					<td width="5"></td>
					<td nowrap>
						<input id="txtCharacter" maxlength="2" name="txtCharacter" class="text textdisabled" disabled="disabled" style="width: 30px"
							onkeypress=" eventChanged(); "
							onkeydown=" eventChanged(); "
							onchange=" eventChanged(); ">
					</td>
					<td width="5"></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td colspan="1">
						<input type="radio" name="optKey" id="optLegendLookup" disabled="disabled"
							onclick=" refreshLegendControls(); eventChanged(); " />
					</td>
					<td style="width: 100%; white-space: nowrap" colspan="3">
						<label
							tabindex="-1"
							for="optLegendLookup"
							class="radio radiodisabled">
							Lookup Table</label>
					</td>
					<td width="5"></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td>&nbsp;</td>
					<td width="100" nowrap>Event Type :
					</td>
					<td width="5"></td>
					<td>
						<select id="cboEventType" name="cboEventType" disabled="disabled" width="100%" style="width: 100%" class="combo combodisabled"
							onchange=" eventChanged(); ">
						</select>
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
						<select id="cboLegendTable" name="cboLegendTable" disabled="disabled" class="combo combodisabled" style="width: 100%"
							onchange=" changeLegendTable(); ">

							<%	' Get the lookup table records.
								Dim sErrorDescription = ""

								Try
									Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
							
									Dim rstLookupTablesInfo = objDataAccess.GetFromSP("spASRIntGetLookupTables")
									
									For Each objRow As DataRow In rstLookupTablesInfo.Rows
										Response.Write("<option value='" & objRow("tableID").ToString() & "'>" & objRow("tableName").ToString() & vbCrLf)									
									Next

									
								Catch ex As Exception
									sErrorDescription = "The lookup tables information could not be retrieved." & vbCrLf &
									FormatError(ex.Message)

								End Try
								
							%>
						</select>
					</td>
					<td width="5"></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td nowrap></td>
					<td width="100" nowrap>Column :
					</td>
					<td width="5"></td>
					<td>
						<select id="cboLegendColumn" name="cboLegendColumn" width="100%" style="width: 100%" class="combo"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td width="5"></td>
				</tr>
				<tr height="5">
					<td width="5"></td>
					<td></td>
					<td width="100" nowrap>Code :
					</td>
					<td width="5"></td>
					<td>
						<select id="cboLegendCode" name="cboLegendCode" width="100%" style="width: 100%" class="combo"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td width="5"></td>
				</tr>
				<tr style="font-weight: bold;">
					<td style="padding-top: 20px" colspan="6">Event Description</td>
				</tr>
				<tr height="10">
					<td width="5"></td>
					<td width="5"></td>
					<td nowrap width="100">Description 1 : </td>
					<td width="5">&nbsp;</td>
					<td>
						<select disabled="disabled" id="cboEventDesc1" name="cboEventDesc1" width="100%" class="combo combodisabled" style="width: 100%"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td width="5"></td>
				</tr>
				<tr height="10">
					<td width="5"></td>
					<td width="5"></td>
					<td nowrap width="100">Description 2 : </td>
					<td width="5"></td>
					<td>
						<select disabled="disabled" id="cboEventDesc2" name="cboEventDesc2" width="100%" class="combo combodisabled" style="width: 100%"
							onchange=" eventChanged(); ">
						</select>
					</td>
					<td width="5"></td>
				</tr>
			</table>
		</div>

		<div id="Buttons" class="invisible" style="width: 100%; text-align: center">
			<input id="cmdOK" type="button" value="OK" name="cmdOK"
				class="button"
				style="width: 80px"
				onclick=" setForm() "/>
			<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel"
				class="button"
				style="width: 80px"
				onclick="self.close();" />
		</div>

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
	</body>
</html>
<script type="text/javascript">
	util_def_calendarreportdates_window_onload();
</script>
