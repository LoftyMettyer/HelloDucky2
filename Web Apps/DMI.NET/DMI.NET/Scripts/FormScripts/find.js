
//todo remove this function!
//New functionality - get the selected row's record ID from the hidden tag		
function getRecordID(rowID) {
	return $("#findGridTable").find("#" + rowID + " input[type=hidden]").val();
}

function rowCount() {
	return $("#findGridTable tr").length - 1;
}

function bookmarksCount() {
	var selRowIds = $('#findGridTable').jqGrid('getGridParam', 'selarrrow');
	return selRowIds.length;
}

function moveFirst() {
	$("#findGridTable").jqGrid('setSelection', 1);
	menu_refreshMenu();
}

function find_window_onload() {
	var fOk;
	fOk = true;
	$("#workframe").attr("data-framesource", "FIND");
	$("#optionframe").hide();
	$("#workframe").show();

	var frmFindForm = document.getElementById("frmFindForm");
	var sErrMsg = frmFindForm.txtErrorDescription.value;
	if (sErrMsg.length > 0) {
		fOk = false;

		if (menu_isSSIMode()) {
			raiseWarning("Warning", sErrMsg);
			loadPartialView("linksMain", "Home", "workframe", null);
		} else {
			OpenHR.messageBox(sErrMsg);
			menu_loadPage("_default");
		}

	}

	var sFatalErrorMsg = frmFindForm.txtErrorDescription.value;
	if (sFatalErrorMsg.length > 0) {
	} else {
		// Do nothing if the menu controls are not yet instantiated.
		var sCurrentWorkPage = OpenHR.currentWorkPage();
		//To allow option frame to pop out with jQuery dialog control...
		var sOptionWorkPage = $("#optionframe").attr("data-framesource");
		var sErrorMsg;
		var sAction;
		var dataCollection;
		var sControlName;
		var sColumnName;
		var iCount;
		var fRecordAdded;
		var sColumnType;
		var colMode;
		var colNames;
		var sColDef;
		var iIndex;
		var i;
		var colData;
		var colDataArray;
		var obj;
		var iCount2;

		if (sCurrentWorkPage == "FIND") {
			sErrorMsg = frmFindForm.txtErrorDescription.value;
			if (sErrorMsg.length > 0) {
				// We've got an error so don't update the record edit form.
				OpenHR.messageBox(sErrorMsg);
			}
			sAction = frmFindForm.txtGotoAction.value; // Refresh the link find grid with the data if required.

			dataCollection = frmFindForm.elements; // Configure the grid columns.
			colMode = [];
			colNames = [];
			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 14);
					if (sControlName == "txtFindColDef_") {
						// Get the column name and type from the control.
						sColDef = dataCollection.item(i).value;
						iIndex = sColDef.indexOf("	");
						if (iIndex >= 0) {
							sColumnName = sColDef.substr(0, iIndex);
							sColumnType = sColDef.substr(iIndex + 1);
							colNames.push(sColumnName);

							if (sColumnName == "ID") {
								colMode.push({ name: sColumnName, hidden: true });
							} else {
								switch (sColumnType.toLowerCase()) {
									case "boolean": //checkbox - 11
										colMode.push({ name: sColumnName, edittype: "checkbox", formatter: 'checkbox', formatoptions: { disabled: true }, align: 'center', width: 100 });
										break;
									case "decimal": //Numeric - 131
										colMode.push({ name: sColumnName, edittype: "numeric", sorttype: 'integer', formatter: 'numeric', formatoptions: { disabled: true }, align: 'right', width: 100 });
										break;
									case "datetime": //Date - 135
										colMode.push({ name: sColumnName, edittype: "date", sorttype: 'date', formatter: 'date', formatoptions: { srcformat: 'd/m/Y', newformat: 'd/m/Y', disabled: true }, align: 'left', width: 100 });
										break;
									default:
										colMode.push({ name: sColumnName, width: 100 });
								}
							}
						}
					}
				}
			}

			// Add the grid records.
			fRecordAdded = false;
			iCount = 0;
			if (dataCollection != null) {
				colData = [];
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 13);
					if (sControlName == "txtAddString_") {
						colDataArray = dataCollection.item(i).value.split("\t");
						obj = {};
						for (iCount2 = 0; iCount2 < colNames.length; iCount2++) {
							//loop through columns and add each one to the 'obj' object
							obj[colNames[iCount2]] = colDataArray[iCount2];
						}
						//add the 'obj' object to the 'colData' array
						colData.push(obj);

						fRecordAdded = true;
						iCount = iCount + 1;
					}
				}

				var shrinkToFit = false;
				var wfSetWidth = $('#workframeset').width();
				if (colMode.length < (wfSetWidth / 100)) shrinkToFit = true;
				//var gridWidth = menu_isSSIMode() ? 'auto' : wfSetWidth - 100;
				var gridWidth = wfSetWidth - 100;
				//var rowNum = (Number($('#txtFindRecords').val()) > 100) ? 100 : Number($('#txtFindRecords').val());

				//create the column layout:
				$("#findGridTable").jqGrid({
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colMode,
					rowNum: 50,
					width: gridWidth,
					pager: $('#pager-coldata'),
					ignoreCase: true,
					shrinkToFit: shrinkToFit,
					ondblClickRow: function () {
						menu_editRecord();
					},
					loadComplete: function () {
						moveFirst();
					},
					afterSearch: function() {
						moveFirst();
					}
				});

				$("#findGridTable").jqGrid('bindKeys', {
					"onEnter": function () {
						menu_editRecord();
					}
				});

				//search options.
				$("#findGridTable").jqGrid('navGrid', '#pager-coldata', { del: false, add: false, edit: false, search: false });

				$("#findGridTable").jqGrid('navButtonAdd', "#pager-coldata", {
					caption: '',
					buttonicon: 'ui-icon-search',
					onClickButton: function () {
						$("#findGridTable").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });
					},
					position: 'first',
					title: '',
					cursor: 'pointer'
				});


				//resize the grid to the height of its container.
				var gridRowHeight = $("#findGridRow").height();
				var gridHeaderHeight = $('#findGridRow .ui-jqgrid-hdiv').height();
				var gridFooterHeight = $('#findGridRow .ui-jqgrid-pager').height();
				var newHeight = gridRowHeight - gridHeaderHeight - gridFooterHeight;

				$("#findGridTable").jqGrid('setGridHeight', newHeight);
			}

			//NOTE: may come in useful.
			//http://stackoverflow.com/questions/12572780/jqgrids-addrowdata-hangs-for-large-number-of-records

			frmFindForm.txtRecordCount.value = iCount;

			// **************************************************************
			if (fOk == true) {

				//TODO: check the font settings.
				//setGridFont(frmFindForm.ssOleDBGridFindRecords);

				var frmMenuInfo = document.getElementById("frmMenuInfo");
				var isDMISingle = ($("#txtIsDMISingle")[0].value == "True");

				if ($("#workframe").length == 0) { //only check if not in SSI mode.
					if (isDMISingle &&
							(frmMenuInfo.txtPersonnel_EmpTableID.value == frmFindForm.txtCurrentTableID.value) &&
							(frmFindForm.txtRecordCount.value > 1)) {

						$("#findGridTable").focus();
						$("#findGridTable").html = ""; //empty the grid

						// Get menu.asp to refresh the menu.
						menu_refreshMenu();

						/* The user does NOT have permission to create new records. */
						OpenHR.messageBox("Unable to load personnel records.\n\nYou are logged on as a self-service user and can access only single record personnel record sets.");

						/* Go to the default page. */
						menu_loadPage("main?SSIMode=True");
						return;
					}
				}
			}

			if (fOk == true) {
				var sControlPrefix;
				var sColumnId;
				var ctlSummaryControl;
				var sSummaryControlName;
				var sDataType;

				// Need to dim focus on the grid before adding the items.
				$("#findGridTable").focus();

				var controlCollection = frmFindForm.elements;
				if (controlCollection !== null) {
					for (i = 0; i < controlCollection.length; i++) {
						sControlName = controlCollection.item(i).name;
						sControlPrefix = sControlName.substr(0, 15);

						if (sControlPrefix == "txtSummaryData_") {
							sColumnId = sControlName.substr(15);
							sSummaryControlName = "ctlSummary_";
							sSummaryControlName = sSummaryControlName.concat(sColumnId);
							sSummaryControlName = sSummaryControlName.concat("_");

							for (var j = 0; j < controlCollection.length; j++) {
								sControlName = controlCollection.item(j).name;
								sControlPrefix = sControlName.substr(0, sSummaryControlName.length);

								if (sControlPrefix == sSummaryControlName) {
									ctlSummaryControl = controlCollection.item(j);

									if (ctlSummaryControl.type == "checkbox") {
										ctlSummaryControl.checked = (controlCollection.item(i).value.toUpperCase() == "TRUE");
									} else {
										// Check if the control is for a datevalue.
										sDataType = sControlName.substr(sSummaryControlName.length);

										if (sDataType == "11") {
											// Format dates for the locale setting.							
											if (controlCollection.item(i).value == '') {
												ctlSummaryControl.value = '';
											} else {
												ctlSummaryControl.value = controlCollection.item(i).value;
											}
										} else {
											ctlSummaryControl.value = controlCollection.item(i).value;
										}
									}
									break;
								}
							}
						}
					}
				}

				// Select the current record in the grid if its there, else select the top record if there is one.
				if (rowCount() > 0) {
					if ((frmFindForm.txtCurrentRecordID.value > 0) && (frmFindForm.txtGotoAction.value != 'LOCATE')) {
						// Try to select the current record.
						locateRecord(frmFindForm.txtCurrentRecordID.value, true);
					} else {
						moveFirst();
					}
				}
				// Get menu.asp to refresh the menu.	    		
				menu_refreshMenu();

				if ((rowCount() == 0) && (frmFindForm.txtFilterSQL.value.length > 0)) {
					OpenHR.messageBox("No records match the current filter.\nNo filter is applied.");
					menu_clearFilter();
				}
			}
		}
	}

	$("#findGridTable").keydown(function (event) {
		//If keyboard pressed while grid is in focus, check it's not the grid keys, then pass focus to locate box...

		var keyPressed = event.which;
		//up arrow, down arrow, Enter, spacebar, home, end, pgup and pgdn.
		if ((keyPressed != 40) && (keyPressed != 38) && (keyPressed != 13) && (keyPressed != 32) && (keyPressed != 33) && (keyPressed != 34) && (keyPressed != 35) && (keyPressed != 36))
			$('#txtLocateRecordFind').focus();
	});

}

/* Return the ID of the record selected in the find form. */
function selectedRecordID() {
	var iRecordId;

	iRecordId = $("#findGridTable").getGridParam('selrow');
	iRecordId = $("#findGridTable").jqGrid('getCell', iRecordId, 'ID');

	return (iRecordId);
}

/* Sequential search the grid for the required ID. */
function locateRecord(psSearchFor, pfIdMatch) {
	//select the grid row that contains the record with the passed in ID.
	var rowNumber = $("#findGridTable input[value='" + psSearchFor + "']").parent().parent().attr("id");
	if (rowNumber >= 0) {
		$("#findGridTable").jqGrid('setSelection', rowNumber);
	} else {
		$("#findGridTable").jqGrid('setSelection', 1);
	}
}

