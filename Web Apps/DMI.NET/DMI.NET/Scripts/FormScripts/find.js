
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
			OpenHR.messageBox(sErrMsg);
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
		var colModel;
		var colNames;
		var colNamesOriginal;
		var sColDef;
		var iIndex;
		var i;
		var colData;
		var colDataArray;
		var obj;
		var iCount2;
		var thereIsAtLeastOneEditableColumn = false;
		var columnCount = -1;
		
		if (sCurrentWorkPage == "FIND") {
			sErrorMsg = frmFindForm.txtErrorDescription.value;
			if (sErrorMsg.length > 0) {
				// We've got an error so don't update the record edit form.
				OpenHR.messageBox(sErrorMsg);
			}
			sAction = frmFindForm.txtGotoAction.value; // Refresh the link find grid with the data if required.

			var newFormat = OpenHR.getLocaleDateString();
			var srcFormat = newFormat;
			if (newFormat.toLowerCase().indexOf('y.m.d') >= 0) srcFormat = 'd/m/Y';
			
			dataCollection = frmFindForm.elements; // Configure the grid columns.
			colModel = [];
			colNames = [];
			colNamesOriginal = [];

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
							var sColumnDisplayName = dataCollection.item(i).getAttribute("data-colname");
							var iColumnId = dataCollection.item(i).getAttribute("data-columnid");
							var sColumnEditable = dataCollection.item(i).getAttribute("data-editable") == "1" ? true : false;
							var ColumnDataType = dataCollection.item(i).getAttribute("data-datatype");
							var ColumnControlType = dataCollection.item(i).getAttribute("data-controltype");
							var ColumnSize = dataCollection.item(i).getAttribute("data-size");
							var ColumnDecimals = dataCollection.item(i).getAttribute("data-decimals");
							var ColumnLookupTableID = dataCollection.item(i).getAttribute("data-lookuptableid");
							var ColumnLookupColumnID = dataCollection.item(i).getAttribute("data-lookupcolumnid");
							var ColumnLookupFilterColumnID = dataCollection.item(i).getAttribute("data-lookupfiltercolumnid");
							var ColumnLookupFilterValueID = dataCollection.item(i).getAttribute("data-lookupfiltervalueid");
							var ColumnSpinnerMinimum = parseInt(dataCollection.item(i).getAttribute("data-spinnerminimum"));
							var ColumnSpinnerMaximum = parseInt(dataCollection.item(i).getAttribute("data-spinnermaximum"));
							var ColumnSpinnerIncrement = parseInt(dataCollection.item(i).getAttribute("data-spinnerincrement"));

							if (sColumnEditable == true) {
								thereIsAtLeastOneEditableColumn = true;
							}

							colNames.push(sColumnName);
							colNamesOriginal.push(sColumnDisplayName);

							if (sColumnName == "ID") {
								colModel.push({
									name: sColumnName,
									hidden: true,
									editoptions: {
										defaultValue: 0
									}
								});
							} else if (sColumnName == "Timestamp") {
								colModel.push({
									name: sColumnName,
									hidden: true
								});
							} else {
								columnCount += 1;
								//Determine the column type and set the colModel for this column accordingly
								if (ColumnControlType == 1) { //Logic - checkbox
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										edittype: 'checkbox',
										formatter: 'checkbox',
										editable: sColumnEditable,
										formatoptions: {
											disabled: true,
											defaultValue: getDefaultValueForColumn(iColumnId, "checkbox")
										},
										align: 'center',
										width: 100									
									});
								} else if (ColumnDataType == 4) { //Integer	- NOT numerics								
									if (ColumnControlType == 64) {
										// Integer - not a spinner.
										colModel.push({
											name: sColumnName,
											id: iColumnId,
											edittype: 'text',
											sorttype: 'integer',
											formatter: 'numeric',
											editable: sColumnEditable,
											align: 'right',
											width: 100,
											editoptions: {
												defaultValue: getDefaultValueForColumn(iColumnId, "integer"),
												columnSize: ColumnSize,
												columnDecimals: ColumnDecimals,
												dataInit: function (element) {
													var value = "";
													var ColumnSize = $(element).attr('columnSize');
													var ColumnDecimals = $(element).attr('columnDecimals');

													element.setAttribute("data-a-dec", OpenHR.LocaleDecimalSeparator()); //Decimal separator
													element.setAttribute("data-a-sep", ''); //No Thousand separator
													element.setAttribute('data-m-dec', ColumnDecimals); //Decimal places
													$(element).addClass("textalignright");
													//Size of field includes decimals but not the decimal point; For example if Size=6 and Decimals=2 the maximum value to be allowed is 9999.99
													if (ColumnSize == "0") { //No size specified, set a very long limit
														element.setAttribute('data-v-min', '-2147483647'); //This is -Int32.MaxValue
														element.setAttribute('data-v-max', '2147483647'); //This is Int32.MaxValue
													} else {
														//Determine the length we need and "translate" that to use it in the plugin
														var n = Number(ColumnSize) - Number(ColumnDecimals); //Size minus decimal places
														for (var x = n; x--;) value += "9"; //Create a string of the form "999"

														if (ColumnDecimals != "0") { //If decimal places are specified, add a period and an appropriate number of "9"s
															value += ".";
															for (x = Number(ColumnDecimals) ; x--;) value += "9";
														}

														element.setAttribute('data-v-min', '-' + value);
														element.setAttribute('data-v-max', value);
													}

													$(element).autoNumeric('init');
												}
											}
										});
									}
									else if (ColumnControlType == 32768) { //"Colour picker" integer
										colModel.push({
											name: sColumnName,
											id: iColumnId,
											edittype: 'text',
											sorttype: 'integer',
											formatter: 'numeric',
											editable: false,
											align: 'right',
											width: 100,
											editoptions: {
												defaultValue: getDefaultValueForColumn(iColumnId, "integer")
											}
										});
									}
									else { //Spinner integer
										colModel.push({
											name: sColumnName,
											id: iColumnId,
											editable: sColumnEditable,
											type: 'spinner',
											editoptions: {
												size: 10,
												maxlengh: 10,
												min: ColumnSpinnerMinimum,
												max: ColumnSpinnerMaximum,
												step: ColumnSpinnerIncrement,
												dataInit: function (element) {
													$(element).spinner({ });
												},
												defaultValue: getDefaultValueForColumn(iColumnId, "spinner")
											}
										});
									}
								} else if (ColumnDataType == 11) { //Date
									colModel.push({
										name: sColumnName,
										edittype: "text",
										id: iColumnId,
										sorttype: function (cellValue) { //Sort function that deals correctly with empty dates
											if (Date.parse(cellValue)) {
												var d = cellValue.split("/");
												return new Date(d[2].toString() + "-" + d[1].toString() + "-" + d[0].toString());
											} else {
												return new Date("1901-01-01");
											}
										},
										formatter: 'date',
										formatoptions: {
											srcformat: srcFormat,
											newformat: newFormat,
											disabled: true
										},
										align: 'left',
										width: 100,
										editable: sColumnEditable,
										type: "date",
										editoptions: {
											size: 20,
											maxlengh: 10,
											dataInit: function (element) {
												$(element).datepicker({
													constrainInput: true,
													showOn: 'focus'
												});
												$(element).addClass('datepicker');
											},
											defaultValue: getDefaultValueForColumn(iColumnId, "date")
										}
									});
								} else if (ColumnControlType == 64 && ColumnSize > 2000000000) { //Multiline - Textarea
									colModel.push({
										name: sColumnName,
										edittype: "textarea",
										id: iColumnId,
										editable: sColumnEditable,
										type: 'textarea',
										editoptions: {
											dataInit: function (element) { },
											defaultValue: getDefaultValueForColumn(iColumnId, "textarea")
										}
									});
								} else if ((ColumnDataType == 12 || ColumnDataType == 2) && ColumnControlType == 2 && ColumnLookupColumnID != 0) { //Lookup
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										editable: sColumnEditable,
										type: "lookup",
										columnLookupTableID: ColumnLookupTableID,
										columnLookupColumnID: ColumnLookupColumnID,
										columnLookupFilterColumnID: ColumnLookupFilterColumnID,
										columnLookupFilterValueID: ColumnLookupFilterValueID,
										editoptions: {
											dataInit: function (element) {
												//On clicking any cell on the lookup column, popup the lookup dialog
												$(element).on('click', function () { showLookupForColumn(element); });
											},
											defaultValue: getDefaultValueForColumn(iColumnId, "lookup")
										}
									});
								} else if (((ColumnDataType == 12 && ColumnControlType == 2) || (ColumnDataType == 12 && ColumnControlType == 16))
														&& (ColumnLookupColumnID == 0)
													) { //Option Groups or Dropdown Lists

									colModel.push({
										name: sColumnName,
										edittype: "select",
										id: iColumnId,
										editable: sColumnEditable,
										type: "select",
										editoptions: {
											value: getValuesForColumn(iColumnId, (ColumnControlType == 2)), //This populates the <select>
											defaultValue: getDefaultValueForColumn(iColumnId, "select")
										}
									});
								} else if (ColumnDataType == 12 && ColumnControlType == 16384) { //Navigation control, make it a hyperlink
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										editable: false, //Non-editable by design
										type: "navigation",
										formatter: hyperLinkFormatter,
										unformat: hyperLinkDeformatter,
										editoptions: {
											dataInit: function (element) { },
											defaultValue: getDefaultValueForColumn(iColumnId, "navigation")
										}
									});
								} else if (ColumnDataType == -1 && ColumnControlType == 4096) { //Working pattern
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										editable: false, //Hardcoded to false, see Notes on TFS 12732 for reason why
										type: "workingpattern",
										formatter: workingPatternFormatter,
										unformat: workingPatternDeformatter,
										editoptions: {
											defaultValue: getDefaultValueForColumn(iColumnId, "workingpattern")
										}
									});								
								} else if (ColumnDataType == 12 && ColumnControlType == 64) {
									//Character
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										width: 100,
										editable: sColumnEditable,
										type: 'text',
										editoptions: {
											size: ColumnSize,
											maxlength: ColumnSize,
											defaultValue: getDefaultValueForColumn(iColumnId, "text")
										},
										label: sColumnDisplayName
									});
								}
								else if (ColumnDataType == 2 && ColumnControlType == 64) {
									//"Numeric"		
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										width: 100,
										editable: sColumnEditable,
										type: 'other',
										align: 'right',
										sorttype: 'number',
										formatter: 'number',
										formatoptions: {
											decimalSeparator: OpenHR.LocaleDecimalSeparator(),
											thousandsSeparator: useThousandSeparator(columnCount) ? OpenHR.LocaleThousandSeparator() : "",
											decimalPlaces: Number(ColumnDecimals),
											defaultValue: useBlankIfZero(columnCount) ? '' : space("0", ColumnSize, ColumnDecimals)
										},										
										editoptions: {
											columnSize: ColumnSize,
											columnDecimals: Number(ColumnDecimals),
											decimalSeparator: OpenHR.LocaleDecimalSeparator(),
											thousandsSeparator: useThousandSeparator(columnCount) ? OpenHR.LocaleThousandSeparator() : "",
											defaultValue: getDefaultValueForColumn(iColumnId, "other"),											
											dataInit: function (element) {												
												var value = "";
												var ColumnSize = $(element).attr('columnSize');
												var ColumnDecimals = $(element).attr('columnDecimals');
												element.setAttribute("data-a-dec", OpenHR.LocaleDecimalSeparator()); //Decimal separator
												element.setAttribute("data-a-sep", ""); //Thousand separator - no thousand separator when editing!
												element.setAttribute('data-m-dec', ColumnDecimals); //Decimal places
												$(element).addClass("textalignright");
												//Size of field includes decimals but not the decimal point; For example if Size=6 and Decimals=2 the maximum value to be allowed is 9999.99
												if (ColumnSize == "0") { //No size specified, set a very long limit
													element.setAttribute('data-v-min', '-2147483647'); //This is -Int32.MaxValue
													element.setAttribute('data-v-max', '2147483647'); //This is Int32.MaxValue
												} else {
													value = space("9", ColumnSize, ColumnDecimals);

													element.setAttribute('data-v-min', '-' + value);
													element.setAttribute('data-v-max', value);
												}

												$(element).autoNumeric('init');
											},
										},
										label: sColumnDisplayName
									});
								}
								else { //None of the above
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										width: 100,
										editable: sColumnEditable,
										type: 'other',
										editoptions: {
											size: "20",
											maxlength: "30",
											defaultValue: getDefaultValueForColumn(iColumnId, "other")
										},
										label: sColumnDisplayName
									});
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
						for (iCount2 = 0; iCount2 < (colNames.length); iCount2++) {
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
				if (colModel.length < (wfSetWidth / 100)) shrinkToFit = true;
				//var gridWidth = menu_isSSIMode() ? 'auto' : wfSetWidth - 100;
				var gridWidth = wfSetWidth - 100;
				//var rowNum = (Number($('#txtFindRecords').val()) > 100) ? 100 : Number($('#txtFindRecords').val());

				//create the column layout:
				$("#findGridTable").jqGrid({
					data: colData,
					datatype: "local",
					colNames: colNamesOriginal,
					colModel: colModel,
					rowNum: 50,
					width: gridWidth,
					pager: $('#pager-coldata'),
					editurl: 'clientArray',
					ignoreCase: true,
					shrinkToFit: shrinkToFit,
					loadComplete: function () {
						moveFirst();
					},
					afterSearch: function () {
						moveFirst();
					}
				});

				//search options.
				$("#findGridTable").jqGrid('navGrid', '#pager-coldata', { del: false, add: false, edit: false, search: false });

				$("#findGridTable").jqGrid('navButtonAdd', "#pager-coldata", {
					caption: '',
					buttonicon: 'icon-search',
					onClickButton: function () {
						$("#findGridTable").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });
					},
					position: 'first',
					title: 'Search',
					cursor: 'pointer'
				});

				//Enable inline editing if there is at least one editable column
				var editLicenced = ($("#txtEditableGridGranted").val() == 1);
				if (thereIsAtLeastOneEditableColumn && editLicenced) {
					//Make grid editable
					$("#findGridTable").jqGrid('inlineNav', '#pager-coldata', {
						edit: true,
						editicon: 'icon-pencil',
						add: true,
						addicon: 'icon-plus',
						save: true,
						saveicon: 'icon-save',
						cancel: true,
						cancelicon: 'icon-ban-circle',
						editParams: {
							aftersavefunc: function (rowid, response, options) {
								saveInlineRowToDatabase(rowid);																
								updateRowFromDatabase(rowid);
							}
						}
					});

					//Enable inline edit and autosave buttons
					menu_toolbarEnableItem('mnutoolInlineEditRecordFind', true);					

					$("#findGridTable_iladd").show();
					$("#findGridTable_iledit").show();
					$("#findGridTable_ilsave").show();
					$("#findGridTable_ilcancel").show();

				} else {
					//Disable inline edit and autosave buttons
					menu_toolbarEnableItem('mnutoolInlineEditRecordFind', false);
					//Hide the edit icons by default
					$("#findGridTable_iladd").hide();
					$("#findGridTable_iledit").hide();
					$("#findGridTable_ilsave").hide();
					$("#findGridTable_ilcancel").hide();
				}

				$("#pager-coldata .navtable .ui-pg-div>span.ui-icon-refresh").addClass("icon-refresh");
				$("#pager-coldata .navtable .ui-pg-div>span").removeClass("ui-icon");

				var $pager = $("#findGridTable").closest(".ui-jqgrid").find(".ui-pg-table");
				$pager.find(".ui-pg-button>span.ui-icon-seek-first")
						.removeClass("ui-icon ui-icon-seek-first")
						.addClass("icon-step-backward")
						.css('font-size', '20px');
				$pager.find(".ui-pg-button>span.ui-icon-seek-prev")
						.removeClass("ui-icon ui-icon-seek-prev")
						.addClass("icon-backward")
						.css('font-size', '20px');
				$pager.find(".ui-pg-button>span.ui-icon-seek-next")
						.removeClass("ui-icon ui-icon-seek-next")
						.addClass("icon-forward")
						.css('font-size', '20px');
				$pager.find(".ui-pg-button>span.ui-icon-seek-end")
						.removeClass("ui-icon ui-icon-seek-end")
						.addClass("icon-step-forward")
						.css('font-size', '20px');

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

function showLookupForColumn(element) {
	//If we are editing a lookup cell we need to popup a window with its values

	if (!$("#findGridTable_iledit").hasClass('ui-state-disabled')) //If we are not in edit mode then return
		return;

	var el = $(element, $("#findGridTable").rows).closest("td");
	var clickedColumnId = $("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].id;
	var rowId = $("#findGridTable").getGridParam('selrow') - 1;
	var columnLookupTableID = $("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].columnLookupTableID;
	var columnLookupColumnID = $("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].columnLookupColumnID;
	var columnLookupFilterColumnID = $("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].columnLookupFilterColumnID;
	var columnLookupFilterValueID = $("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].columnLookupFilterValueID;
	var filterCellValue = '';
	var colModelContainsRequiredLookupColumn;
	var thisLookupColumnIsNeededByAnother = false;
	var colModel = $("#findGridTable").jqGrid("getGridParam", "colModel");

	for (var j = 0; j <= colModel.length - 1; j++) {
		if (colModel[j].columnLookupFilterValueID == clickedColumnId) {
			thisLookupColumnIsNeededByAnother = true;
		}
	}

	//Determine if the lookup depends on the value of another column
	if (columnLookupFilterColumnID == 0 && columnLookupFilterValueID == 0) { //It doesn't, i.e. it's not filtered
		colModelContainsRequiredLookupColumn = true;
	} else { // It does, i.e. it's filtered
		columnLookupTableID = $("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].id;
		var colModel = $("#findGridTable").jqGrid("getGridParam", "colModel");
		colModelContainsRequiredLookupColumn = false;
		for (var i = 0; i <= colModel.length - 1; i++) {
			if (colModel[i].id == columnLookupFilterValueID) {
				filterCellValue = $("#findGridTable").jqGrid("getGridParam", "data")[rowId][colModel[i].name];
				colModelContainsRequiredLookupColumn = true;
				break;
			}
		}
	}

	if (!colModelContainsRequiredLookupColumn) {
		OpenHR.modalMessage("Unable to display the filtered lookup records.<br/><br/>The lookup filter value is not present in this view.");
		return false;
	}

	var lookupUrl = window.ROOT;
	var lookupParameters = '';

	if (eval('isLookupTable_' + clickedColumnId) == true) {
		lookupUrl += 'generic/GetLookupFindRecords';
		lookupParameters = { piLookupColumnID: columnLookupColumnID, psFilterValue: filterCellValue, piCallingColumnID: columnLookupTableID, piFirstRecPos: 0 };
	} else {
		lookupUrl += 'generic/GetLookupFindRecords2';
		//tableId and orderId below are defined in Find.ascx so they are be available here
		lookupParameters = { piTableID: tableId, piOrderID: orderId, piLookupColumnID: columnLookupColumnID, psFilterValue: filterCellValue, piCallingColumnID: columnLookupTableID, piFirstRecPos: 0 };
	}

	$.ajax({
		url: lookupUrl,
		data: lookupParameters,
		dataType: 'json',
		type: 'GET',
		success: function (jsondata) {
			var colModelLookup = [];
			var colNamesLookup = [];

			var lookupColumnGridPosition = eval('LookupColumnGridPosition_' + clickedColumnId);

			for (var k in jsondata.rows[0]) {
				colModelLookup.push({ name: k });
				colNamesLookup.push(k.replace("_", " "));
			}

			$("#LookupForEditableGrid_Table").jqGrid('GridUnload'); //Unload previous grid (if any)

			//jqGrid it
			$("#LookupForEditableGrid_Table").jqGrid({
				data: jsondata.rows,
				datatype: "local",
				colModel: colModelLookup,
				colNames: colNamesLookup,
				rowNum: 10000,
				ignoreCase: true,
				multiselect: false,
				shrinkToFit: (colModelLookup.length < 8)
			});

			//Set the dialog's title and open it (the dialog, not the title)
			$("#LookupForEditableGrid_Title").html($("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].name);
			$("#LookupForEditableGrid_Div").dialog("open");

			//Resize the grid
			$("#LookupForEditableGrid_Table").jqGrid("setGridHeight", $("#LookupForEditableGrid_Div").height() - 90);
			$("#LookupForEditableGrid_Table").jqGrid("setGridWidth", $("#LookupForEditableGrid_Div").width() - 10);

			//Set overflow-x to hidden 
			if (colModelLookup.length < 8)
				$("#LookupForEditableGrid_Table").parent().parent().css("overflow-x", "hidden");

			//Search for the value that is currently selected in the find grid
			var rowId = null;
			for (i = 0; i <= jsondata.rows.length - 1; i++) {
				if (jsondata.rows[i][colModelLookup[lookupColumnGridPosition].name] == $(element).val()) {
					rowId = i;
					break;
				}
			}

			//If text found, select the row
			if (rowId != null) {
				$("#LookupForEditableGrid_Table").jqGrid('setSelection', rowId + 1, false);
			}

			//Assign a function call to the onclick event of the "Select" button
			$('#LookupForEditableGridSelect').attr('onclick', 'selectValue("Select", "' + lookupColumnGridPosition + '","' + element.id + '",' + thisLookupColumnIsNeededByAnother + ')');
			//Assign a function call to the onclick event of the "Clear" button
			$('#LookupForEditableGridClear').attr('onclick', 'selectValue("Clear", "' + lookupColumnGridPosition + '","' + element.id + '",' + thisLookupColumnIsNeededByAnother + ')');
		},
		error: function (req, status, errorObj) {
			//debugger;
		}
	});
}

function selectValue(action, lookupColumnGridPosition, elementId, thisLookupColumnIsNeededByAnother) {
	// Get the value selected by the user and update the corresponding value in the find grid
	
	var rowId = $("#LookupForEditableGrid_Table").getGridParam('selrow');

	if (rowId == null && action == "Select") { //No row selected, show a message and return
		OpenHR.modalMessage('Please select a value');
		return;
	}

	var columnName = $("#LookupForEditableGrid_Table").getGridParam('colModel')[lookupColumnGridPosition].name;
	var cellValue = ''; //Default for action="Clear"
	
	if (action == "Select") {
		cellValue = $("#LookupForEditableGrid_Table").getRowData(rowId)[columnName];
	}

	document.getElementById(elementId).value = cellValue;
	$('#LookupForEditableGrid_Div').dialog('close');
	$("#LookupForEditableGrid_Table").jqGrid('GridUnload');

	if (thisLookupColumnIsNeededByAnother) {
		//Save the row to the grid (not the database), restore it and then set the row back into edit mode;
		//this is necessary so any lookup column filtered by another column will pickup the correct value to filter on
		var findGridRowId = $("#findGridTable").getGridParam('selrow');
		$('#findGridTable').saveRow(findGridRowId);
		$('#findGridTable').editRow(findGridRowId);
	}
}

function getValuesForColumn(iColumnId, isDropdown) {	
	//Get the values for this column and return them as a json object that jqGrid will use to create a dropdown
	try {
		var data = eval('colOptionGroupOrDropDownData_' + iColumnId);
	} catch (e) {
		return false;
	}

	var values = {};

	if(isDropdown) values[""] = "";	//add empty first option for dropdown lists (not option groups)

	for (var i = 0; i <= data.length - 1; i++) {
		values[data[i][0]] = data[i][0];
	}

	return values;
}

function hyperLinkFormatter(cellValue, options, rowdata, action) {	
	//Format as hyperlink
	return "<a href='" + encodeURI(cellValue) + "' target='_blank'>Navigation</a>";
}

function hyperLinkDeformatter(cellvalue, options, cell) {	
	//Remove the HTML anchor part	
	var cleanUri = cell.innerHTML.replace('<a href="', '').replace("<a href='", "").replace('" target="_blank">Navigation</a>', '').replace("' target='_blank'>Navigation</a>", "");
	return decodeURI(cleanUri);
}

function workingPatternFormatter(cellValue, options, rowdata, action) {
	if (cellValue == undefined)
		return "";

	return cellValue.replace(/ /g, "&nbsp;"); //Replace all spaces with &nbsp; so the working patterns in the column are neatly aligned
}

function workingPatternDeformatter(cellvalue, options, cell) {
	return cell.innerHTML.replace(/&nbsp;/g, " "); //Replace all the &nbsp; with a space so the user can edit the working pattern as a string of text
}

function getDefaultValueForColumn(columnId, columnType) {
	if (columnsDefaultValues[columnId] == "") {
		return "";
	}

	//Some controls need a bit more logic applied to their default values
	switch (columnType) {
		case "checkbox":
			return columnsDefaultValues[columnId].toString().toLowerCase() == "true" ? "yes" : "no";
			break;
		case "date":
			var d = columnsDefaultValues[columnId].toString().split("/");
			return new Date(d[2].toString() + "-" + d[0].toString() + "-" + d[1].toString());
			break;
	}

	return columnsDefaultValues[columnId];
}

function space(character, columnSize, columnDecimals) {
	columnSize = columnSize - columnDecimals;
	try {
		//Determine the length we need and "translate" that to use it in the plugin
		var n = Number(columnSize) - Number(columnDecimals);
		var value = '';
		for (var x = n; x--;) value += character; //Create a string of the form "999"

		if (columnDecimals != "0") { //If decimal places are specified, add a period and an appropriate number of "9"s
			value += ".";
			for (x = Number(columnDecimals); x--;) value += character;
		}
	} catch (e) {
		return '';
	}

	return value;
}

function useThousandSeparator(columnNumber) {
	try {
		return ($('#txtThousandColumns').val().substr(columnNumber, 1) == '1');
	} catch (e) {
		return false;
	}
}

function useBlankIfZero(columnNumber) {
	try {
		return ($('#txtBlankIfZeroColumns').val().substr(columnNumber, 1) == '1');
	} catch (e) {
		return false;
	}
}
