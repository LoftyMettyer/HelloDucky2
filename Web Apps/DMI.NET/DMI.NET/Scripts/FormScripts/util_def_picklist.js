
var frmOriginalDefinition = OpenHR.getForm("workframe", "frmOriginalDefinition");
var frmDefinition = OpenHR.getForm("workframe", "frmDefinition");
var frmUseful = OpenHR.getForm("workframe", "frmUseful");

function util_def_picklist_onload() {

	$("#workframe").attr("data-framesource", "UTIL_DEF_PICKLIST");

	setGridFont(frmDefinition.ssOleDBGrid);

	// Expand the work frame and hide the option frame.
	//            window.parent.document.all.item("workframeset").cols = "*, 0";

	if (frmUseful.txtAction.value.toUpperCase() == "NEW") {
		frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
		frmDefinition.txtDescription.value = "";
	} else {
		loadDefinition();
	}

	if (frmUseful.txtAction.value.toUpperCase() != "EDIT") {
		frmUseful.txtUtilID.value = 0;
	}

	if (frmUseful.txtAction.value.toUpperCase() == "COPY") {
		frmUseful.txtChanged.value = 1;
	}

	try {
		frmDefinition.txtName.focus();
	} catch (e) {
	}

	refreshControls();
	frmUseful.txtLoading.value = 'N';
	try {
		frmDefinition.txtName.focus();
	} catch (e) {
	}

	// Get menu.asp to refresh the menu.
	menu_refreshMenu();
	$('#cmdOK').hide();
	$('#cmdCancel').hide();


}

function refreshControls() {

	var frmUseful = OpenHR.getForm("workframe", "frmUseful");

	var fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
	var fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

	radio_disable(frmDefinition.optAccessRW, ((fIsNotOwner) || (fViewing)));
	radio_disable(frmDefinition.optAccessRO, ((fIsNotOwner) || (fViewing)));
	radio_disable(frmDefinition.optAccessHD, ((fIsNotOwner) || (fViewing)));

	var fAddDisabled = fViewing;
	var fAddAllDisabled = fViewing;
	var fRemoveDisabled = ((frmDefinition.ssOleDBGrid.SelBookmarks.Count == 0) || (fViewing == true));
	var fRemoveAllDisabled = ((frmDefinition.ssOleDBGrid.Rows == 0) || (fViewing == true));

	button_disable(frmDefinition.cmdAdd, fAddDisabled);
	button_disable(frmDefinition.cmdAddAll, fAddAllDisabled);
	button_disable(frmDefinition.cmdFilteredAdd, false);
	button_disable(frmDefinition.cmdRemove, fRemoveDisabled);
	button_disable(frmDefinition.cmdRemoveAll, fRemoveAllDisabled);

	button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
			(fViewing == true) ||
			(frmDefinition.ssOleDBGrid.Rows == 0)));
	menu_toolbarEnableItem('mnutoolSaveReport', (!((frmUseful.txtChanged.value == 0) ||
			(fViewing == true) ||
			(frmDefinition.ssOleDBGrid.Rows == 0))));


	// Get menu.asp to refresh the menu.
	menu_refreshMenu();
}

function submitDefinition() {
	
	if (validate() == false) { menu_refreshMenu(); return; }
	if (populateSendForm() == false) { menu_refreshMenu(); return; }

	// first populate the validate fields
	var frmValidate = document.getElementById("frmValidate");
	frmValidate.validatePass.value = 1;
	frmValidate.validateName.value = frmDefinition.txtName.value;
	frmValidate.validateAccess.value = frmSend.txtSend_access.value;

	if (frmUseful.txtAction.value.toUpperCase() == "EDIT") {
		frmValidate.validateTimestamp.value = frmOriginalDefinition.txtDefn_Timestamp.value;
		frmValidate.validateUtilID.value = frmUseful.txtUtilID.value;
	}
	else {
		frmValidate.validateTimestamp.value = 0;
		frmValidate.validateUtilID.value = 0;
	}

	OpenHR.showInReportFrame(frmValidate);

}

function addClick() {
	
	var vBM;

	/* Get the current selected delegate IDs. */
	var sSelectedIDs1 = new String("0");

	frmDefinition.ssOleDBGrid.redraw = false;
	if (frmDefinition.ssOleDBGrid.rows > 0) {
		frmDefinition.ssOleDBGrid.MoveFirst();
	}

	for (var iIndex = 1; iIndex <= frmDefinition.ssOleDBGrid.rows; iIndex++) {
		vBM = frmDefinition.ssOleDBGrid.AddItemBookmark(iIndex);

		var sRecordID = new String(frmDefinition.ssOleDBGrid.Columns("ID").CellValue(vBM));

		sSelectedIDs1 = sSelectedIDs1 + "," + sRecordID;

	}
	frmDefinition.ssOleDBGrid.redraw = true;

	var frmSend = OpenHR.getForm("workframe", "frmPicklistSelection");
	frmSend.selectionAction = "add";
	frmSend.selectionType.value = "ALL";
	frmSend.selectedIDs1.value = sSelectedIDs1;

	$("#workframeset").hide();
	OpenHR.showInReportFrame(frmSend);

}

function openWindow(mypage, myname, w, h, scroll) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
	var winprops;

	if (scroll == 'no') {
		winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resize=no';
	}
	else {
		winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';
	}

	var win = window.open(mypage, myname, winprops);
	if (win.opener == null) win.opener = self;
	if (parseInt(navigator.appVersion) >= 4) win.window.focus();
}

function addAllClick() {
	frmUseful.txtChanged.value = 1;
	picklistdef_makeSelection("ALLRECORDS", 0, "");
}

function filteredAddClick() {

	var vBM;

	/* Get the current selected delegate IDs. */
	var sSelectedIDs1 = new String("0");

	frmDefinition.ssOleDBGrid.redraw = false;
	if (frmDefinition.ssOleDBGrid.rows > 0) {
		frmDefinition.ssOleDBGrid.MoveFirst();
	}

	for (var iIndex = 1; iIndex <= frmDefinition.ssOleDBGrid.rows; iIndex++) {
		vBM = frmDefinition.ssOleDBGrid.AddItemBookmark(iIndex);
		var sRecordID = new String(frmDefinition.ssOleDBGrid.Columns("ID").CellValue(vBM));
		sSelectedIDs1 = sSelectedIDs1 + "," + sRecordID;
	}
	frmDefinition.ssOleDBGrid.redraw = true;

	var frmSend = OpenHR.getForm("workframe", "frmPicklistSelection");
	frmSend.selectionAction = "add";
	frmSend.selectionType.value = "FILTER";
	frmSend.selectedIDs1.value = sSelectedIDs1;

	OpenHR.showInReportFrame(frmSend);

}

function removeClick() {

	var i;

	// Do nothing of the Add button is disabled (read-only mode).
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") return;

	var iCount = frmDefinition.ssOleDBGrid.selbookmarks.Count();
	for (i = iCount - 1; i >= 0; i--) {
		frmDefinition.ssOleDBGrid.bookmark = frmDefinition.ssOleDBGrid.selbookmarks(i);
		var iRowIndex = frmDefinition.ssOleDBGrid.AddItemRowIndex(frmDefinition.ssOleDBGrid.Bookmark);

		if ((frmDefinition.ssOleDBGrid.Rows == 1) && (iRowIndex == 0)) {
			frmDefinition.ssOleDBGrid.RemoveAll();
		}
		else {
			frmDefinition.ssOleDBGrid.RemoveItem(iRowIndex);
		}
	}

	frmUseful.txtChanged.value = 1;

	//Display the number of records
	$('#RecordCountDIV').html(frmDefinition.ssOleDBGrid.Rows.toString() + " Record(s)");

	refreshControls();
}

function removeAllClick() {
	var iAnswer = OpenHR.messageBox("Remove all records from the picklist. \n Are you sure ?", 36, "Confirmation");
	if (iAnswer == 7) {
		// cancel 
		return;
	}

	frmDefinition.ssOleDBGrid.redraw = false;
	frmDefinition.ssOleDBGrid.RemoveAll();
	frmDefinition.ssOleDBGrid.redraw = true;

	frmUseful.txtChanged.value = 1;

	//Display the number of records
	$('#RecordCountDIV').html(frmDefinition.ssOleDBGrid.Rows.toString() + " Record(s)");

	refreshControls();
}

function cancelClick() {
	if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
			(definitionChanged() == false)) {
		menu_loadDefSelPage(10, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, true);
		return (false);
	}

	var answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3);
	if (answer == 7) {
		// No
		menu_loadDefSelPage(10, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, true);
		return (false);
	}
	if (answer == 6) {
		// Yes
		okClick();
	}

	return false;
}

function okClick() {
	menu_refreshMenu();

	frmSend.txtSend_reaction.value = "PICKLISTS";
	submitDefinition();
}

function picklistdef_makeSelection(psType, piID, psPrompts) {

	$(".popup").dialog("close");
	$("#workframeset").show();

	/* Get the current selected delegate IDs. */
	var sSelectedIDs = "0";
	var iIndex;
	var sRecordID;

	if (psType != "ALLRECORDS") {
		frmDefinition.ssOleDBGrid.redraw = false;
		if (frmDefinition.ssOleDBGrid.rows > 0) {
			frmDefinition.ssOleDBGrid.MoveFirst();
		}
		for (iIndex = 1; iIndex <= frmDefinition.ssOleDBGrid.rows; iIndex++) {
			sRecordID = new String(frmDefinition.ssOleDBGrid.Columns("ID").Value);

			sSelectedIDs = sSelectedIDs + "," + sRecordID;

			if (iIndex < frmDefinition.ssOleDBGrid.rows) {
				frmDefinition.ssOleDBGrid.MoveNext();
			}
			else {
				break;
			}
		}
		frmDefinition.ssOleDBGrid.redraw = true;
	}

	if ((psType == "ALL") && (psPrompts.length > 0)) {
		sSelectedIDs = sSelectedIDs + "," + psPrompts;
	}

	// Get the optionData.asp to get the required records.
	var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
	optionDataForm.txtOptionAction.value = "GETPICKLISTSELECTION";
	optionDataForm.txtOptionPageAction.value = psType;
	optionDataForm.txtOptionRecordID.value = piID;
	optionDataForm.txtOptionValue.value = sSelectedIDs;
	optionDataForm.txtOptionPromptSQL.value = psPrompts;
	optionDataForm.txtOptionTableID.value = frmUseful.txtTableID.value;
	optionDataForm.txtOption1000SepCols.value = frmDefinition.txt1000SepCols.value;

	refreshOptionData();
}

function saveChanges() {
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
			// Compare the controls with the original values.
			if (frmDefinition.txtName.value != frmOriginalDefinition.txtDefn_Name.value) {
				return true;
			}

			if (frmDefinition.txtDescription.value != frmOriginalDefinition.txtDefn_Description.value) {
				return true;
			}

			if (frmOriginalDefinition.txtDefn_Access.value == "RW") {
				if (frmDefinition.optAccessRW.checked == false) {
					return true;
				}
			}
			else {
				if (frmOriginalDefinition.txtDefn_Access.value == "RO") {
					if (frmDefinition.optAccessRO.checked == false) {
						return true;
					}
				}
				else {
					if (frmDefinition.optAccessHD.checked == false) {
						return true;
					}
				}
			}
		}
	}

	return false;
}

function openDialog(pDestination, pWidth, pHeight) {
	var dlgwinprops = "center:yes;" +
			"dialogHeight:" + pHeight + "px;" +
			"dialogWidth:" + pWidth + "px;" +
			"help:no;" +
			"resizable:yes;" +
			"scroll:yes;" +
			"status:no;";
	window.showModalDialog(pDestination, self, dlgwinprops);
	//window.open(pDestination);

}

function validate() {
	// Check name has been entered.
	if (frmDefinition.txtName.value == '') {
		OpenHR.messageBox("You must enter a name for this definition.");
		return (false);
	}

	// Check thet picklist list does have some records.      
	if (frmDefinition.ssOleDBGrid.rows == 0) {
		OpenHR.messageBox("Picklists must contain at least one record.");
		return (false);
	}

	return (true);
}

function createNew(pPopup) {
	pPopup.close();

	frmUseful.txtUtilID.value = 0;
	frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
	frmUseful.txtAction.value = "new";

	submitDefinition();
}

function populateSendForm() {
	var i;

	var frmSend = document.getElementById("frmSend");

	// Copy all the header information to frmSend
	frmSend.txtSend_ID.value = frmUseful.txtUtilID.value;
	frmSend.txtSend_name.value = frmDefinition.txtName.value;
	frmSend.txtSend_description.value = frmDefinition.txtDescription.value;
	frmSend.txtSend_userName.value = frmDefinition.txtOwner.value;
	if (frmDefinition.optAccessRW.checked == true) {
		frmSend.txtSend_access.value = "RW";
	}
	if (frmDefinition.optAccessRO.checked == true) {
		frmSend.txtSend_access.value = "RO";
	}
	if (frmDefinition.optAccessHD.checked == true) {
		frmSend.txtSend_access.value = "HD";
	}

	// Now go through the records grid
	var sColumns = '';

	frmDefinition.ssOleDBGrid.Redraw = false;
	frmDefinition.ssOleDBGrid.movefirst();

	for (i = 0; i < frmDefinition.ssOleDBGrid.rows; i++) {
		sColumns = sColumns + frmDefinition.ssOleDBGrid.columns("ID").text + ',';

		frmDefinition.ssOleDBGrid.movenext();
	}
	frmDefinition.ssOleDBGrid.Redraw = true;

	frmSend.txtSend_columns.value = sColumns.substr(0, 8000);
	frmSend.txtSend_columns2.value = sColumns.substr(8000, 8000);

	if (sColumns.length > 16000) {
		OpenHR.messageBox("Too many records selected.");
		return false;
	}
	else {
		return true;
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

	if (frmOriginalDefinition.txtDefn_Access.value == "RW") {
		frmDefinition.optAccessRW.checked = true;
	}
	else {
		if (frmOriginalDefinition.txtDefn_Access.value == "RO") {
			frmDefinition.optAccessRO.checked = true;
		}
		else {
			frmDefinition.optAccessHD.checked = true;
		}
	}

	// Load the selected records into the grid.
	//makeSelection("ALL", 0, frmOriginalDefinition.txtSelectedRecords.value);
	picklistdef_makeSelection("PICKLIST", frmUseful.txtUtilID.value, '');

	frmDefinition.ssOleDBGrid.MoveFirst();
	frmDefinition.ssOleDBGrid.FirstRow = frmDefinition.ssOleDBGrid.Bookmark;

	// If its read only, disable everything.
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
		disableAll();
	}
}

function disableAll() {
	var i;

	var dataCollection = frmDefinition.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			var eElem = frmDefinition.elements[i];

			if (("text" == eElem.type) || ("TEXTAREA" == eElem.tagName)) {
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
				treeView_disable(eElem, true);
			}
		}
	}
}

function locateRecord(psSearchFor) {
	var fFound;
	var iIndex;

	fFound = false;

	frmDefinition.ssOleDBGrid.redraw = false;

	frmDefinition.ssOleDBGrid.MoveLast();
	frmDefinition.ssOleDBGrid.MoveFirst();

	frmDefinition.ssOleDBGrid.SelBookmarks.removeall();

	for (iIndex = 1; iIndex <= frmDefinition.ssOleDBGrid.rows; iIndex++) {
		var sGridValue = new String(frmDefinition.ssOleDBGrid.Columns(0).value);
		sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
		if (sGridValue == psSearchFor.toUpperCase()) {
			frmDefinition.ssOleDBGrid.SelBookmarks.Add(frmDefinition.ssOleDBGrid.Bookmark);
			fFound = true;
			break;
		}

		if (iIndex < frmDefinition.ssOleDBGrid.rows) {
			frmDefinition.ssOleDBGrid.MoveNext();
		}
		else {
			break;
		}
	}

	if ((fFound == false) && (frmDefinition.ssOleDBGrid.rows > 0)) {
		// Select the top row.
		frmDefinition.ssOleDBGrid.MoveFirst();
		frmDefinition.ssOleDBGrid.SelBookmarks.Add(frmDefinition.ssOleDBGrid.Bookmark);
	}

	frmDefinition.ssOleDBGrid.redraw = true;
}

function changeName() {
	frmUseful.txtChanged.value = 1;
	refreshControls();
}

function changeDescription() {
	frmUseful.txtChanged.value = 1;
	refreshControls();
}

function changeAccess() {
	frmUseful.txtChanged.value = 1;
	refreshControls();
}


	function util_def_addhandlers() {
		OpenHR.addActiveXHandler("ssOleDBGrid", "KeyPress", "ssOleDBGrid_KeyPress()");
		OpenHR.addActiveXHandler("ssOleDBGrid", "SelChange", "ssOleDBGrid_SelChange()");
	}

function ssOleDBGrid_rowColChange() {
	refreshControls();
}

function ssOleDBGrid_KeyPress(iKeyAscii) {

	var iLastTick;
	var sFind;
	var txtLastKeyFind = $("#txtLastKeyFind")[0];
	var txtTicker = $("#txtTicker")[0];

	if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
		var dtTicker = new Date();
		var iThisTick = new Number(dtTicker.getTime());
		if (txtLastKeyFind.value.length > 0) {
			iLastTick = new Number(txtTicker.value);
		}
		else {
			iLastTick = new Number("0");
		}

		if (iThisTick > (iLastTick + 1500)) {
			sFind = String.fromCharCode(iKeyAscii);
		}
		else {
			sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
		}

		txtTicker.value = iThisTick;
		txtLastKeyFind.value = sFind;

		locateRecord(sFind);
	}
}

function ssOleDBGrid_SelChange() {
	refreshControls();
}

