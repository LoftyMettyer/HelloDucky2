
var frmOriginalDefinition = OpenHR.getForm("workframe", "frmOriginalDefinition");
var frmDefinition = OpenHR.getForm("workframe", "frmDefinition");
var frmUseful = OpenHR.getForm("workframe", "frmUseful");

function util_def_picklist_onload() {

	$("#workframe").attr("data-framesource", "UTIL_DEF_PICKLIST");

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
	} catch(e) {
	}

	refreshControls();

	frmUseful.txtLoading.value = 'N';
	try {
		frmDefinition.txtName.focus();
	} catch(e) {
	}

	// Get menu.asp to refresh the menu.
	//menu_refreshMenu();
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
	var fRemoveDisabled = (fViewing == true);
	var fRemoveAllDisabled = (fViewing == true);
	
	button_disable(frmDefinition.cmdAdd, fAddDisabled);
	button_disable(frmDefinition.cmdAddAll, fAddAllDisabled);
	//button_disable(frmDefinition.cmdFilteredAdd, false);	
	button_disable(frmDefinition.cmdRemove, fRemoveDisabled);
	button_disable(frmDefinition.cmdRemoveAll, fRemoveAllDisabled);

	// Check if, user has read/write access and the grid has any rows or not. If no rows, then disabled the remove & removeall button
	if (fViewing == false) {
		disableRemoveAndRemoveAllButton();
	}

	menu_toolbarEnableItem('mnutoolSaveReport', (!((frmUseful.txtChanged.value == 0) || (fViewing == true))));

	// Get menu.asp to refresh the menu.
	menu_refreshMenu();
}

function submitDefinition() {
	if (validate() == false) { menu_refreshMenu(); return; }
	if (populateSendForm() == false) { menu_refreshMenu(); return; }

	var sTimeStamp;
	var sUtilID;

	if (frmUseful.txtAction.value.toUpperCase() == "EDIT") {
		sTimeStamp = frmOriginalDefinition.txtDefn_Timestamp.value;
		sUtilID = frmUseful.txtUtilID.value;
		}
	else {
		sTimeStamp = 0;
		sUtilID = 0;
	}

	var postData = {
			validatePass: 1,
			validateName: frmDefinition.txtName.value,
			validateTimestamp: sTimeStamp,
			validateUtilID: sUtilID,
			validateBaseTableID: frmSend.txtSend_tableID.value,
			validateAccess: frmSend.txtSend_access.value
			};

	OpenHR.submitForm(null, "reportframe", null, postData, "util_validate_picklist");

}

function addClick() {
	/* Get the current selected delegate IDs. */
	picklistdef_moveFirst();

	$("#workframeset").show();

	var postData = {
		TableID: $("#txtTableID").val(),
		Action: "add",
		Type: "ALL",
		IDs1: $('#ssOleDBGrid').getDataIDs().join(","),
		__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
	}
	OpenHR.submitForm(null, "reportframe", null, postData, "picklistSelectionMain");

}

function addAllClick() {
	$("#ssOleDBGrid").jqGrid('GridUnload'); //Clear the grid
	$('#RecordCountDIV').html(''); //Clear the "Records" label

	frmUseful.txtChanged.value = 1;
	picklistdef_makeSelection("ALLRECORDS", 0, "");
}

function filteredAddClick() {	
	/* Get the current selected delegate IDs. */
	picklistdef_moveFirst();

	var postData = {
		TableID: $("#txtTableID").val(),
		Action: "add",
		Type: "FILTER",
		IDs1: $('#ssOleDBGrid').getDataIDs().join(","),
		__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
	}
	OpenHR.submitForm(null, "reportframe", null, postData, "picklistSelectionMain");

}

function removeClick() {
	if ($("#ssOleDBGrid").getGridParam('reccount') == 0) {
		return;
	}

	// Do nothing if the Add button is disabled (read-only mode).
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") return;

	var grid = $("#ssOleDBGrid");
	var myDelOptions = {
		// because I use "local" data I don't want to send the changes
		// to the server so I use "processing:true" setting and delete
		// the row manually in onclickSubmit
		onclickSubmit: function (options) {
			var grid_id = $.jgrid.jqID(grid[0].id),
					grid_p = grid[0].p,
					newPage = grid_p.page,
					rowids = grid_p.multiselect ? grid_p.selarrrow : [grid_p.selrow];

			// reset the value of processing option which could be modified
			options.processing = true;

			// delete the row
			$.each(rowids, function () {
				grid.delRowData(this);
			});
			$.jgrid.hideModal("#delmod" + grid_id,
												{
													gb: "#gbox_" + grid_id,
													jqm: options.jqModal, onClose: options.onClose
												});

			if (grid_p.lastpage > 1) {// on the multipage grid reload the grid
				if (grid_p.reccount === 0 && newPage === grid_p.lastpage) {
					// if after deliting there are no rows on the current page
					// which is the last page of the grid
					newPage--; // go to the previous page
				}
				// reload grid to make the row from the next page visable.
				grid.trigger("reloadGrid", [{ page: newPage }]);
			}
			else
			{
				//Reload the grid data again, after selected row is deleted and set the row selection to -1.
				grid.trigger("reloadGrid", [{ current: true }]);
				if (grid.getGridParam("reccount") > 0)
				{
					grid.jqGrid('setSelection', -1);
				}
			}

			//Display the number of records
			$('#RecordCountDIV').html($("#ssOleDBGrid").getGridParam('reccount') + " Record(s)");

			return true;
		},
		processing: true
	};

	grid.jqGrid('delGridRow', grid.jqGrid('getGridParam', 'selarrrow'), myDelOptions);

	$("#dData").click(); //To remove the "delete confirmation" dialog

	frmUseful.txtChanged.value = 1;

	refreshControls();

	button_disable(frmDefinition.cmdRemove, true); //Disable the "Remove" button
	if ($("#ssOleDBGrid").getGridParam('reccount') == 0) { //If the grid is empty, disable the "Remove All" button
		button_disable(frmDefinition.cmdRemoveAll, true);
	}
}

function removeAllClick() {
	if ($("#ssOleDBGrid").getGridParam('reccount') == 0) {
		return;
	}

	OpenHR.modalPrompt("Remove all records from the picklist. \n Are you sure ?", 4, "Confirmation", removeAllClickFollowOn);

}


function removeAllClickFollowOn(iAnswer) {
	if (iAnswer == 7) {
		// cancel 
		return;
	}

	$("#ssOleDBGrid").jqGrid('clearGridData');

	frmUseful.txtChanged.value = 1;

	//Display the number of records
	$('#RecordCountDIV').html($("#ssOleDBGrid").getGridParam('reccount') + " Record(s)");

	refreshControls();

	button_disable(frmDefinition.cmdRemove, true); //Disable the "Remove" button
	button_disable(frmDefinition.cmdRemoveAll, true); //Disable the "Remove All" button
}


function cancelClick() {
	if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") || (definitionChanged() == false)) {
		menu_loadDefSelPage(10, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, true);
	}
	else {
		OpenHR.modalPrompt("You have made changes. Click 'OK' to discard your changes, or 'Cancel' to continue editing.", 1, "Confirm").then(function (answer) {
			if (answer == 1) {  // OK
				menu_loadDefSelPage(10, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, true);
			}
		});
	}
	return (false);
}

function okClick() {

	menu_refreshMenu();

	frmSend.txtSend_reaction.value = "PICKLISTS";
	submitDefinition();
}

function picklistdef_makeSelection(psType, piID, psPrompts) {

	$("#workframeset").show();

	/* Get the current selected delegate IDs. */
	var sSelectedIDs = "0";

	sSelectedIDs = $('#ssOleDBGrid').getDataIDs().join(",");

	if ((psType == "ALL") && (psPrompts.length > 0)) {
		if (sSelectedIDs == "") {
			sSelectedIDs = psPrompts;
		} else {
			sSelectedIDs = sSelectedIDs + "," + psPrompts;
		}
	}

	//Close the popup now that we have read the selected IDs.
	if ($(".popup").dialog("isOpen")) $(".popup").dialog("close");

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
		return 6; // No changes made. Continue navigation
	} else {
		return 0;
	}
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

function validate() {
	// Check name has been entered.
	if (frmDefinition.txtName.value == '') {
		OpenHR.modalMessage("You must enter a name for this definition.");
		return (false);
	}

	// Check the picklist list does have some records.      
	if (($("#ssOleDBGrid").getGridParam('reccount') == 0) ||($("#ssOleDBGrid").getGridParam('reccount') == undefined)) {
		OpenHR.modalMessage("Picklists must contain at least one record.");
		return (false);
	}

	return (true);
}

function udp_createNew() {

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
	var sColumns = $("#ssOleDBGrid").getDataIDs().join(",");

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

	picklistdef_moveFirst();

	// If its read only, disable everything.
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
		disableAll();
	}
}

function picklistdef_moveFirst() {
	$('#ssOleDBGrid').jqGrid('setSelection', 1);
	menu_refreshMenu();
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

function disableRemoveAndRemoveAllButton() {	
	if ($("#ssOleDBGrid").getGridParam('reccount') == 0)//If the grid is empty, disable the "Remove" and "Remove All" button
	{		
		button_disable(frmDefinition.cmdRemove, true);
		button_disable(frmDefinition.cmdRemoveAll, true);
	}
	else if (($("#ssOleDBGrid").getGridParam('selrow') == null) && ($("#ssOleDBGrid").getGridParam('reccount') > 0))
	{
		//Remove button should not enables when no rows are selected.
		button_disable(frmDefinition.cmdRemove, true);
		button_disable(frmDefinition.cmdRemoveAll, false);
	}
}

function BindDefaultGridOnNewDefinition() {
	//Load empty grid when we clik new button
	if (frmUseful.txtAction.value.toUpperCase() == "NEW") {
		dataCollection = frmDefinition.elements; // Add the grid records.
		fRecordAdded = false;
		iCount = 0;

		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGrid").jqGrid('GridUnload');

		// new bit for colmodel
		colMode = [];
		colNames = [];

		if (dataCollection != null) {
			for (i = 0; i < dataCollection.length; i++) {
				sControlName = dataCollection.item(i).name;
				sControlName = sControlName.substr(0, 16);

				if (sControlName == "txtOptionColDef_") {
					// Get the column name and type from the control.
					sColDef = dataCollection.item(i).value;

					iIndex = sColDef.indexOf("	");
					if (iIndex >= 0) {
						sColumnName = sColDef.substr(0, iIndex);
						sColumnType = sColDef.substr(iIndex + 1).replace('System.', '').toLowerCase();

						colNames.push(sColumnName);

						if (sColumnName.toUpperCase() == "ID") {
							colMode.push({ name: sColumnName, hidden: true });
						} else {
							switch (sColumnType) {
								case "boolean": // "11":
									colMode.push({ name: sColumnName, edittype: "checkbox", formatter: 'checkbox', formatoptions: { disabled: true }, align: 'center', width: 100 });
									break;
								case "decimal":
									colMode.push({ name: sColumnName, edittype: "numeric", sorttype: 'integer', formatter: 'numeric', formatoptions: { disabled: true }, align: 'right', width: 100 });
									break;
								case "datetime": //Date - 135
									colMode.push({ name: sColumnName, edittype: "date", sorttype: 'date', formatter: 'date', formatoptions: { srcformat: dateFormat, newformat: dateFormat, disabled: true }, align: 'left', width: 100 });
									break;
								default:
									colMode.push({ name: sColumnName, width: 100 });
							}
						}
					}
				}
			}
		}

		//Determine if the grid already exists...
		if ($("#ssOleDBGrid").getGridParam("reccount") == undefined) { //It doesn't exist, create it
			var shrinkToFit = false;
			if (colMode.length < 8) shrinkToFit = true;
			var gridWidth = $('#PickListGrid').width();
			$("#ssOleDBGrid").jqGrid({
				multiselect: true,
				datatype: 'local',
				colNames: colNames,
				colModel: colMode,
				rowNum: 1000,
				width: gridWidth,
				shrinkToFit: shrinkToFit
			}).jqGrid('hideCol', 'cb');

			//resize the grid to the height of its container.
			$("#ssOleDBGrid").jqGrid('setGridHeight', $("#PickListGrid").height());

			if ($("#ssOleDBGrid").jqGrid().width() < $("#PickListGrid").width()) {
				$("#ssOleDBGrid").parent().parent().addClass('jqgridHideHorScroll');
			}
		}

		//Display the number of records
		$('#RecordCountDIV').html($("#ssOleDBGrid").getGridParam('reccount') + " Record(s)");
		//Hide Remove and RemoveAll button
		button_disable(frmDefinition.cmdRemove, true); //Disable the "Remove" button
		button_disable(frmDefinition.cmdRemoveAll, true); //Disable the "Remove All" button
	}
		
}
