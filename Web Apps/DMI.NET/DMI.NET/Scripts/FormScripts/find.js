var rowIsEditedOrNew = "";
var thereIsAtLeastOneEditableColumn = false;
var lastRowEdited = "0";
var followOnRow = 0;
var gridDefaultHeight;
var addparameters;

function rowCount() {
	return $("#findGridTable tr").length - 1;
}

function moveFirst() {
	try {
		var firstRecordID = $("#findGridTable").jqGrid('getDataIDs')[0];
		$("#findGridTable").jqGrid('setSelection', firstRecordID);
		refreshInlineNavIcons();
	} catch (e) { }

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
							var sReadOnly = dataCollection.item(i).getAttribute("data-editable") == "1" ? false : true;
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
							var ColumnMask = dataCollection.item(i).getAttribute("data-Mask");
							var iDefaultValueExprID = dataCollection.item(i).getAttribute("data-DefaultValueExprID");
							var BlankIfZero = dataCollection.item(i).getAttribute("data-BlankIfZero");

							if (sReadOnly == false) {
								thereIsAtLeastOneEditableColumn = true;
							}

							colNames.push(sColumnName);
							colNamesOriginal.push(sColumnDisplayName);

							if (sColumnName == "ID") {
								colModel.push({
									name: sColumnName,
									hidden: true,
									key: true,
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
										editable: true,
										editoptions: {
											readonly: sReadOnly,
											dataColumnId: iColumnId,
											dataDefaultCalcExprID: iDefaultValueExprID,
											value: "1:0",
											dataInit: function (element) {
												$(element).on('click', function () { indicateThatRowWasModified(); });
											}
										},
										formatoptions: {
											disabled: true,
											defaultValue: getDefaultValueForColumn(iColumnId, "checkbox")
										},
										align: 'center',
										width: 100
									});
								} else if (ColumnDataType == 4 && ColumnControlType != 2) { //Integer - NOT numerics; the "ColumnControlType != 2" condition is so this if is NOT true for Integer lookups (they are covered below)
									if (ColumnControlType == 64) { // Integer - not a spinner.
										colModel.push({
											name: sColumnName,
											id: iColumnId,
											edittype: 'text',
											sorttype: 'integer',
											formatter: 'numeric',
											editable: true,
											align: 'right',
											width: 100,
											editoptions: {
												readonly: sReadOnly,
												defaultValue: BlankIfZero == '1' ? '' : '0',
												columnSize: ColumnSize,
												columnDecimals: ColumnDecimals,
												dataColumnId: iColumnId,
												dataDefaultCalcExprID: iDefaultValueExprID,
												dataInit: function (element) {
													var value = "";
													var ColumnSize = $(element).attr('columnSize');
													var ColumnDecimals = $(element).attr('columnDecimals');

													$(element).on('keydown', function () { indicateThatRowWasModified(); });

													$(element).on('blur', function (sender) {
														if ($(this).val() == 0) {
															$(this).val(BlankIfZero == '1' ? '' : '0');
														}
													});

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
											editable: true,
											align: 'right',
											width: 100,
											editoptions: {
												readonly: true,
												defaultValue: getDefaultValueForColumn(iColumnId, "integer")
											}
										});
									}
									else if (ColumnControlType == 32) { //Spinner integer
										colModel.push({
											name: sColumnName,
											id: iColumnId,
											editable: true,
											type: 'spinner',
											editoptions: {
												readonly: sReadOnly,
												size: 10,
												maxlength: 10,
												editrules: { integer: true },
												min: ColumnSpinnerMinimum,
												max: ColumnSpinnerMaximum,
												step: ColumnSpinnerIncrement,
												dataColumnId: iColumnId,
												dataDefaultCalcExprID: iDefaultValueExprID,
												dataInit: function (element) {
													var valueBeforeChange = $(element).val();
													$(element).spinner({
														spin: function (event, ui) { indicateThatRowWasModified(); }
													});
													$(element).on('keydown', function () { indicateThatRowWasModified(); });
													$(element).on('change', function () {
														indicateThatRowWasModified();
													});
													$(element).on('blur', function (sender) {
														if ((isNaN(sender.target.value) === true) || (sender.target.value.indexOf(".") >= 0)) {
															OpenHR.modalMessage("Invalid integer value entered: " + encodeURIComponent(sender.target.value));
															sender.target.value = valueBeforeChange;
														}
													});
												},
												defaultValue: getDefaultValueForColumn(iColumnId, "spinner")
											}
										});
									}
									else { } //Integer control not being catered for
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
										editable: true,
										type: "date",
										editoptions: {
											readonly: sReadOnly,
											dataColumnId: iColumnId,
											dataDefaultCalcExprID: iDefaultValueExprID,
											size: 20,
											maxlength: 10,
											dataInit: function (element) {
												var valueBeforeChange = $(element).val();
												$(element).datepicker({
													constrainInput: true,
													showOn: 'focus'
												});
												$(element).addClass('datepicker');
												$(element).on('keydown', function () { indicateThatRowWasModified(); });
												$(element).on('change', function () { indicateThatRowWasModified(); });

												$(element).on('blur', function (sender) {
													if (OpenHR.IsValidDate(sender.target.value) == false && sender.target.value != "") {
														OpenHR.modalMessage("Invalid date value entered: " + encodeURIComponent(sender.target.value));
														sender.target.value = valueBeforeChange;
														$(sender.target.id).focus();
													}
												});
											},
											defaultValue: getDefaultValueForColumn(iColumnId, "date")
										}
									});
								} else if (ColumnControlType == 64 && ColumnSize >= 2147483646) { //Multiline - Textarea
									colModel.push({
										name: sColumnName,
										edittype: "textarea",
										id: iColumnId,
										editable: true,
										type: 'textarea',
										editoptions: {
											readonly: sReadOnly,
											dataColumnId: iColumnId,
											dataDefaultCalcExprID: iDefaultValueExprID,
											dataInit: function (element) {
												$(element).on('keydown', function () { indicateThatRowWasModified(); });
												$(element).attr('onpaste', 'indicateThatRowWasModified();');
											},
											defaultValue: getDefaultValueForColumn(iColumnId, "textarea")
										}
									});
								} else if ((ColumnDataType == 12 || ColumnDataType == 2 || ColumnDataType == 4) && ColumnControlType == 2 && ColumnLookupColumnID != 0) { //Lookup
									var sAlignment = 'left';
									if (ColumnDataType == 2) sAlignment = 'right';

									colModel.push({
										name: sColumnName,
										id: iColumnId,
										editable: true,
										align: sAlignment,
										type: "lookup",
										columnLookupTableID: ColumnLookupTableID,
										columnLookupColumnID: ColumnLookupColumnID,
										columnLookupFilterColumnID: ColumnLookupFilterColumnID,
										columnLookupFilterValueID: ColumnLookupFilterValueID,
										editoptions: {
											readonly: sReadOnly,
											align: sAlignment,
											dataColumnId: iColumnId,
											dataDefaultCalcExprID: iDefaultValueExprID,
											dataInit: function (element) {
												$(element).on('keydown', function () { return false; }); //Prevent the user from typing in lookups
												$(element).attr('onpaste', 'return false;'); //Prevent the user from pasting into lookups
												$(element).addClass('msClear'); //Remove the "x" that IE shows on the right side of input boxes
												var sAlignment = $(element).attr('align');
												$(element).css('text-align', sAlignment);

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
										editable: true,
										type: "select",
										editoptions: {
											readonly: sReadOnly,
											dataColumnId: iColumnId,
											dataDefaultCalcExprID: iDefaultValueExprID,
											dataInit: function (element) {
												$(element).on('change', function () { indicateThatRowWasModified(); });
											},
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
										editable: true, //Hardcoded to false, see Notes on TFS 12732 for reason why
										type: "workingpattern",
										formatter: workingPatternFormatter,
										unformat: workingPatternDeformatter,
										editoptions: {
											readonly: true,
											defaultValue: getDefaultValueForColumn(iColumnId, "workingpattern")
										}
									});
								} else if (ColumnDataType == 12 && ColumnControlType == 64) { //Character
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										width: 100,
										editable: true,
										type: 'text',
										editoptions: {
											readonly: sReadOnly,
											dataColumnId: iColumnId,
											dataDefaultCalcExprID: iDefaultValueExprID,
											size: ColumnSize,
											maxlength: ColumnSize,
											mask: ColumnMask,
											defaultValue: getDefaultValueForColumn(iColumnId, "text"),
											dataInit: function (element) {
												$(element).on('keydown', function () { indicateThatRowWasModified(); });
												$(element).attr('onpaste', 'indicateThatRowWasModified();');
												var ColumnMask = $(element).attr('mask');
												if (ColumnMask == null) return false;
												if (ColumnMask == "") return false;
												$(element).mask(ColumnMask);
												var title = $(element).val() + ' (Input Mask: ' + ColumnMask + ')';
												$(element).attr('title', title);
											}
										},
										label: sColumnDisplayName
									});
								}
								else if ((ColumnDataType == 2 && ColumnControlType == 64) || (ColumnDataType == 2 && ColumnControlType == 2)) {
									//"Numeric"		
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										width: 100,
										editable: true,
										type: 'other',
										edittype: 'custom',
										align: 'right',
										sorttype: 'number',
										formatter: 'number',
										formatoptions: {
											decimalSeparator: OpenHR.LocaleDecimalSeparator(),
											thousandsSeparator: useThousandSeparator(columnCount) ? OpenHR.LocaleThousandSeparator() : "",
											decimalPlaces: Number(ColumnDecimals),
											defaultValue: BlankIfZero == '1' ? '' : space('0', ColumnSize, ColumnDecimals)
										},
										editoptions: {
											readonly: sReadOnly,
											custom_element: ABSNumber,
											custom_value: ABSNumberValue,
											columnSize: ColumnSize,
											columnDecimals: Number(ColumnDecimals),
											decimalSeparator: OpenHR.LocaleDecimalSeparator(),
											thousandsSeparator: useThousandSeparator(columnCount) ? OpenHR.LocaleThousandSeparator() : "",
											defaultValue: getDefaultValueForColumn(iColumnId, "other"),
											dataColumnId: iColumnId,
											dataDefaultCalcExprID: iDefaultValueExprID
										},
										label: sColumnDisplayName
									});
								}
								else { //None of the above
									colModel.push({
										name: sColumnName,
										id: iColumnId,
										width: 100,
										editable: true,
										type: 'other',
										editoptions: {
											readonly: sReadOnly,
											dataColumnId: iColumnId,
											dataDefaultCalcExprID: iDefaultValueExprID,
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
						for (iCount2 = 0; iCount2 < (colNames.length) ; iCount2++) {
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
				var gridWidth = wfSetWidth - 100;
				var rowNum = 50;

				//create the column layout:
				$("#findGridTable").jqGrid({
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colModel,
					rowNum: rowNum,
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

				// Navbar options = i.e. search, edit, save etc 
				$("#findGridTable").jqGrid('navGrid', '#pager-coldata', { del: false, add: false, edit: false, search: false, refresh: false }); // setup the buttons we want
				$("#findGridTable").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });  //instantiate toolbar so we can use toggle.
				$("#findGridTable")[0].toggleToolbar();  // Toggle it off at start up.

				$("#findGridTable").jqGrid('navButtonAdd', "#pager-coldata", {
					caption: '',
					buttonicon: 'icon-search',
					onClickButton: function () {
						$("#findGridTable")[0].toggleToolbar(); // Toggle toolbar on & off when Search button is pressed.
						$("#findGridTable")[0].clearToolbar(); // clear menu

						var isSearching = $('#frmFindForm .ui-search-toolbar').is(':visible');

						$("#findGridTable_iledit").toggleClass('ui-state-disabled', isSearching);
						$("#findGridTable_iladd").toggleClass('ui-state-disabled', isSearching);

						if (isSearching) {
							var currentHeight = $('#findGridRow div.ui-jqgrid-bdiv').outerHeight();
							$("#findGridTable").jqGrid('setGridHeight', currentHeight - 31);
						} else {
							$("#findGridTable").jqGrid('setGridHeight', gridDefaultHeight);
						}
					},
					position: 'first',
					title: 'Search',
					cursor: 'pointer',
					id: 'findGridTable_searchButton'
				});


				addparameters = {
					rowID: "0", //Default ID for New Record
					useDefValues: true,
					position: "last",
					addRowParams: {
						keys: true,
						aftersavefunc: function (rowid, response, options) {
							window.onbeforeunload = null;
							return afterSaveFindGridRow(rowid);
						},						
						oneditfunc: function (rowid) {
							//build a comma separated list of columns that have expression ID's on them.
							var arrCalcColumnsString = [];
							rowWasModified = true;
							lastRowEdited = "0";
							window.onbeforeunload = warning;
							$('#' + rowid).find(':input[datadefaultcalcexprid]').each(function () {
								if (Number(this.attributes['datadefaultcalcexprid'].value) > "0") {
									arrCalcColumnsString.push(this.attributes['dataColumnId'].value);
								}
							});

							var calcColumnsString = arrCalcColumnsString.join(",");

							//Pass list to stored proc.
							$.ajax({
								url: "GetDefaultCalcValueForColumn",
								data: { defaultCalcColumns: calcColumnsString },
								async: false,
								cache: false,
								dataType: 'json',
								type: 'GET'
							}).done(function (jsondata) {
								for (var rowCount = 0; rowCount <= jsondata.length - 1; rowCount++) {
									var key = Object.keys(jsondata[rowCount])[0];
									var value = jsondata[rowCount][key];

									//Some controls need a bit more logic applied to their default values
									if ($('#' + rowid + ' *[datacolumnid="' + key + '"]').hasClass('datepicker')) {
										//Date control.
										$('#' + rowid + ' *[datacolumnid="' + key + '"]').val(OpenHR.ConvertSQLDateToLocale(value));
									} else if ($('#' + rowid + ' *[datacolumnid="' + key + '"]').is(':checkbox')) {
										//checkbox												
										$('#' + rowid + ' *[datacolumnid="' + key + '"]').prop('checked', value.toString().toLowerCase() == "true" ? true : false);
									} else {
										//covers textboxes, dropdowns and option groups
										$('#' + rowid + ' *[datacolumnid="' + key + '"]').val(value);
									}
								}
							});

							return addFindGridRow(rowid);

						}
					}
				};


				//Enable inline editing if there is at least one editable column
				var editLicenced = ($("#txtEditableGridGranted").val() == 1);
				if (editLicenced && linktype != 'multifind') { //The "linktype" variable is defined in Find.ascx
					//Make grid editable
					$("#findGridTable").jqGrid('inlineNav', '#pager-coldata', {
						edit: true, //Set it to always true, but the logic to show or hide the edit icon is now below as well as in menu.js
						editicon: 'icon-pencil',
						add: insertGranted && thereIsAtLeastOneEditableColumn, //Add row should only be enabled if insert is granted AND there is at least one editable column (The insertGranted variable is defined in Find.ascx)
						addicon: 'icon-plus',
						save: true,
						saveicon: 'icon-save',
						cancel: true,
						cancelicon: 'icon-ban-circle',
						editParams: {
							oneditfunc: function (rowid) {
								lastRowEdited = rowid;
								window.onbeforeunload = warning;
								return editFindGridRow(rowid);
							},
							aftersavefunc: function (rowid, response, options) {	//save button clicked in edit mode. NB: row has been 'saved' locally by this time.
								window.onbeforeunload = null;
								return afterSaveFindGridRow(rowid);
							},
							afterrestorefunc: function (rowid) {	//Cancel button clicked in edit mode.
								window.onbeforeunload = null;
								return cancelFindGridRow(rowid);
							}
						},
						addParams: addparameters
					});


					$("#findGridTable_iladd").show();

					//assign click to pager buttons - these fire first and will be rejected if we're editing.
					$('#last_pager-coldata>span, #next_pager-coldata>span, #prev_pager-coldata>span, #first_pager-coldata>span').on('click', function (event) {
						if (rowIsEditedOrNew == "edited") return false;
					});

					//assign click to add button (this will fire before the addrow function)

					//Ensure nothing fires if the button is disabled.
					$('#findGridTable_ilsave div.ui-pg-div, #findGridTable_ilcancel div.ui-pg-div, #findGridTable_iledit div.ui-pg-div').off('click').on('click', function (event) {
						if ($(this).parent().is(":disabled")) {
							return false;
						}
					});

					//Move to last page before adding new row.
					$('#findGridTable_iladd div.ui-pg-div').off('click').on('click', function (event) {

						if (rowIsEditedOrNew == "") {
							//Not editing, no need to save, just scroll to end of grid before adding new row.
							//New row is added by jqGrid's default action of clicking the add button.
							var lastPage = $("#findGridTable").jqGrid('getGridParam', 'lastpage');
							$("#findGridTable").trigger("reloadGrid", [{ page: lastPage }]);
						} else {
							//we need to save the current row before moving on...
							//we'll also manually add the new row (i.e. prevent default addrow click)
							rowIsEditedOrNew = "new";

							//disable aftersavefunc being called by 'saverow'. We save manually.
							var saveparameters = {
								"successfunc": null,
								"url": null,
								"extraparam": {},
								"aftersavefunc": null,
								"errorfunc": null,
								"afterrestorefunc": null,
								"restoreAfterError": true,
								"mtype": "POST"
						}

							$('#findGridTable').saveRow(lastRowEdited, saveparameters);

							saveRowToDatabase(lastRowEdited);

							//Actual adding of row done in submitForm (after server-side validation)
							return false;
						}
					});

					//continuing with window.onload function now....
					var recCountInGrid = $("#findGridTable").getGridParam("reccount");
					if (thereIsAtLeastOneEditableColumn) {
						$("#findGridTable_iledit").show();
						$("#findGridTable_ilsave").show();
						$("#findGridTable_ilcancel").show();
					} else {
						$("#findGridTable_iledit").hide();
						$("#findGridTable_ilsave").hide();
						$("#findGridTable_ilcancel").hide();
					}


				} else {
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
				var gridRowHeight = $('#workframeset').height();
				var pageTitleHeight = $('#row1').outerHeight();
				var gridHeaderHeight = $('#findGridRow .ui-jqgrid-hdiv').outerHeight();
				var gridFooterHeight = $('#findGridRow .ui-jqgrid-pager').outerHeight();
				var footerMargin = 30;
				var summaryRowHeight = 0;

				if (menu_isSSIMode()) pageTitleHeight += 40; //bottom margin for SSI.

				try {
					summaryRowHeight = $('#row3').outerHeight();
					if (summaryRowHeight > 0) summaryRowHeight += 30;
					if (summaryRowHeight > (gridRowHeight * 0.35)) summaryRowHeight = (gridRowHeight * 0.35);
				} catch (e) {

				}
				gridDefaultHeight = gridRowHeight - pageTitleHeight - gridHeaderHeight - gridFooterHeight - footerMargin - summaryRowHeight;

				$("#findGridTable").jqGrid('setGridHeight', gridDefaultHeight);
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
												ctlSummaryControl.value = OpenHR.ConvertSQLDateToLocale(controlCollection.item(i).value);
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

	$('#findGridTable_iledit').addClass('ui-state-disabled');

}


function saveRowToDatabase(rowid) {
	if (saveThisRowToDatabase)
		if (rowWasModified) saveInlineRowToDatabase(rowid);
		else if (rowIsEditedOrNew.substr(0, 9) == 'quickedit') editNextRow();
		else if (rowIsEditedOrNew == "new") addNextRow();
}


/* Return the ID of the record selected in the find form. */
function selectedRecordID() {
	return $("#findGridTable").getGridParam('selrow');
}


/* Sequential search the grid for the required ID. */
function locateRecord(psSearchFor, pfIdMatch) {

	var firstRecordID = $("#findGridTable").jqGrid('getDataIDs')[0];

	//default to top row
	if (Number(firstRecordID) > 0)
		$("#findGridTable").jqGrid('setSelection', firstRecordID);

	if (Number(psSearchFor) > 0)
		$("#findGridTable").jqGrid('setSelection', psSearchFor);

}

function showLookupForColumn(element) {
	//If we are editing a lookup cell we need to popup a window with its values

	if (!$("#findGridTable_iledit").hasClass('ui-state-disabled')) //If we are not in edit mode then return
		return false;

	var el = $(element, $("#findGridTable").rows).closest("td");
	var clickedColumnId = $("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].id;
	var rowId = $("#findGridTable").getGridParam('selrow');
	var rowNumber = $("#findGridTable").jqGrid("getDataIDs").indexOf(rowId);
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
		//columnLookupTableID = $("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].id;
		var colModel = $("#findGridTable").jqGrid("getGridParam", "colModel");
		colModelContainsRequiredLookupColumn = false;
		for (var i = 0; i <= colModel.length - 1; i++) {
			if (colModel[i].id == columnLookupFilterValueID) {
				if ((isNaN(rowId)) || (rowId == 0)) { //If this is a new row get the filterCellValue from the last row added (i.e. the new one)
					filterCellValue = $("#findGridTable").jqGrid("getGridParam", "data")[$("#findGridTable").jqGrid("getGridParam", "data").length - 1][colModel[i].name];
					if (typeof filterCellValue == "undefined") {
						filterCellValue = '';
					}
				} else {//Get the filterCellValue from the current row
					filterCellValue = $("#findGridTable").jqGrid("getGridParam", "data").filter(function (rownum) { return rownum.ID == rowId })[0][colModel[i].name];
				}

				colModelContainsRequiredLookupColumn = true;
				break;
			}
		}
	}

	if (!colModelContainsRequiredLookupColumn) {
		OpenHR.modalMessage("Unable to display the filtered lookup records.<br/><br/>The lookup filter value is not present in this find window.");
		return false;
	}

	var lookupUrl = window.ROOT;
	var lookupParameters = '';

	if (eval('isLookupTable_' + clickedColumnId) == true) {
		lookupUrl += 'generic/GetLookupFindRecords';
		lookupParameters = { piLookupColumnID: columnLookupColumnID, psFilterValue: filterCellValue, piCallingColumnID: clickedColumnId, piFirstRecPos: 0 };
	} else {
		lookupUrl += 'generic/GetLookupFindRecords2';
		//tableId and orderId below are defined in Find.ascx so they are be available here
		lookupParameters = { piTableID: tableId, piOrderID: orderId, piLookupColumnID: columnLookupColumnID, psFilterValue: filterCellValue, piCallingColumnID: clickedColumnId, piFirstRecPos: 0 };
	}

	$.ajax({
		url: lookupUrl,
		data: lookupParameters,
		dataType: 'json',
		type: 'GET',
		cache: false,
		success: function (jsondata) {
			var lookupColumnGridPosition = eval('LookupColumnGridPosition_' + clickedColumnId);

			$("#LookupForEditableGrid_Table").jqGrid('GridUnload'); //Unload previous grid (if any)

			//jqGrid it
			$("#LookupForEditableGrid_Table").jqGrid({
				data: jsondata.rows,
				datatype: "local",
				colModel: jsondata.colmodel,
				rowNum: 10000,
				ignoreCase: true,
				multiselect: false,
				shrinkToFit: (jsondata.colmodel.length < 8)
			});

			//Set the dialog's title and open it (the dialog, not the title)
			$("#LookupForEditableGrid_Title").html($("#findGridTable").jqGrid("getGridParam", "colModel")[$(el).index()].name);
			$("#LookupForEditableGrid_Div").dialog("open");

			//Resize the grid
			$("#LookupForEditableGrid_Table").jqGrid("setGridHeight", $("#LookupForEditableGrid_Div").height() - 90);
			$("#LookupForEditableGrid_Table").jqGrid("setGridWidth", $("#LookupForEditableGrid_Div").width() - 10);

			//Search for the value that is currently selected in the find grid
			rowId = null;
			for (i = 0; i <= jsondata.rows.length - 1; i++) {
				if (jsondata.rows[i][jsondata.colmodel[0].name] == $(element).val()) {
					rowId = i;
					break;
				}
			}

			//If text found, select the row
			if (rowId != null) {
				$("#LookupForEditableGrid_Table").jqGrid('setSelection', rowId + 1, false);
			}

			//If we don't have records in the grid, disable Select button
			if ($("#LookupForEditableGrid_Table").getGridParam('reccount') == 0) {
				$('#LookupForEditableGridSelect').attr('disabled', 'disabled');
				$('#LookupForEditableGridSelect').addClass('disabled');
				$('#LookupForEditableGridSelect').addClass('ui-state-disabled');
			} else { //Enable Select button
				$('#LookupForEditableGridSelect').removeAttr('disabled');
				$('#LookupForEditableGridSelect').removeClass('disabled');
				$('#LookupForEditableGridSelect').removeClass('ui-state-disabled');
				//Assign a function call to the onclick event of the "Select" button
				$('#LookupForEditableGridSelect').attr('onclick', 'selectValue("Select", "' + lookupColumnGridPosition + '","' + element.id + '",' + thisLookupColumnIsNeededByAnother + ')');
			}
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

	var cellValue = ''; //Default for action="Clear"

	if (action == "Select") {
		var columnName = $("#LookupForEditableGrid_Table").getGridParam('colModel')[lookupColumnGridPosition].name;
		cellValue = $("#LookupForEditableGrid_Table").getRowData(rowId)[columnName];
	}

	document.getElementById(elementId).value = cellValue;
	$('#LookupForEditableGrid_Div').dialog('close');
	$("#LookupForEditableGrid_Table").jqGrid('GridUnload');

	indicateThatRowWasModified();

	if (thisLookupColumnIsNeededByAnother) {
		//Save the row to the grid (not the database), restore it and then set the row back into edit mode;
		//this is necessary so any lookup column filtered by another column will pickup the correct value to filter on
		var findGridRowId = $("#findGridTable").getGridParam('selrow');
		//The .saveRow line below triggers the aftersavefunc event which saves the row to the database;
		//when setting a lookup value on a cell we don't want the value to be saved to the database, so...
		saveThisRowToDatabase = false; // ...don't save to the database
		$('#findGridTable').saveRow(findGridRowId);
		saveThisRowToDatabase = true; // ...save to the database again (this is the normal behaviour)
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

	if (isDropdown) values[""] = "";	//add empty first option for dropdown lists (not option groups)

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

function space(character, columnSize, columnDecimals, decimalCharacter) {
	try {
		//Determine the length we need and "translate" that to use it in the plugin
		var n = Number(columnSize) - Number(columnDecimals);
		var value = '';
		for (var x = n; x--;) value += character; //Create a string of the form "999"

		if (columnDecimals != "0") { //If decimal places are specified, add a period and an appropriate number of "9"s			
			value += (OpenHR.nullsafeString(decimalCharacter).length > 0) ? OpenHR.nullsafeString(decimalCharacter) : OpenHR.LocaleDecimalSeparator();
			for (x = Number(columnDecimals) ; x--;) value += character;
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

function indicateThatRowWasModified() {
	rowWasModified = true; //The 'rowWasModified' variable is defined as global in Find.ascx
	$("#findGridTable_ilsave").removeClass('ui-state-disabled'); //Enable the Save button because we edited something
}

function warning() {
	return "You will lose your changes if you do not save before leaving this page.\n\nWhat do you want to do?";
}

function ABSNumber(value, options) {

	var el = document.createElement("input");
	el.type = "text";
	el.value = value.replace(".", OpenHR.LocaleDecimalSeparator());

	$(el).on('keydown', function () { indicateThatRowWasModified(); });
	$(el).attr('onpaste', 'indicateThatRowWasModified();');

	el.setAttribute("data-a-dec", OpenHR.LocaleDecimalSeparator()); //Decimal separator
	el.setAttribute("data-a-sep", OpenHR.LocaleThousandSeparator()); //Thousand separator - no thousand separator when editing!
	$(el).addClass("textalignright");
	el.style.width = '98%';

	el.setAttribute("defaultValue", options.dataColumnId);
	el.setAttribute("dataColumnId", options.dataColumnId);
	el.setAttribute("dataDefaultCalcExprID", options.dataDefaultCalcExprID);
	if (options.readonly) el.setAttribute("readonly", "readonly");

	//Size of field includes decimals but not the decimal point; For example if Size=6 and Decimals=2 the maximum value to be allowed is 9999.99
	if (options.columnSize == "0") { //No size specified, set a very long limit
		el.setAttribute('data-v-min', '-2147483647'); //This is -Int32.MaxValue
		el.setAttribute('data-v-max', '2147483647'); //This is Int32.MaxValue
	} else {
		value = space("9", options.columnSize, options.columnDecimals, ".");

		el.setAttribute('data-v-min', '-' + value);
		el.setAttribute('data-v-max', value);
	}

	$(el).autoNumeric('init');
	return el;
}

function ABSNumberValue(elem, operation, value) {
	if (operation === 'get') {
		var returnVal = OpenHR.replaceAll($(elem).val(), OpenHR.LocaleThousandSeparator(), "");
		returnVal = OpenHR.replaceAll(returnVal, OpenHR.LocaleDecimalSeparator(), ".");
		return returnVal;
	} else if (operation === 'set') {
		$('input', elem).val(value);
	}
}

function saveInlineRowToDatabase(rowId) {

	var sUpdateOrInsert = "";
	var gridData = $("#findGridTable").getRowData(rowId);
	var gridColumns = $("#findGridTable").jqGrid('getGridParam', 'colNames');
	var gridModel = $("#findGridTable").jqGrid('getGridParam', 'colModel');
	var columnValue = "";
	
	for (var i = 0; i <= gridColumns.length - 1; i++) {
		if (gridColumns[i] != '' && gridColumns[i] != 'ID' && gridColumns[i] != 'Timestamp' && gridModel[i].editoptions.readonly == false && gridModel[i].editable == true) {
			columnValue = gridData[gridModel[i].name];

			//If the formatter is undefined then we treat the value as text
			switch (gridModel[i].formatter) {
				case "checkbox":
					if (columnValue == "0" || columnValue == null)
						columnValue = "0";
					else
						columnValue = "1";
					break;
				case "date":
					columnValue = OpenHR.convertLocaleDateToSQL(columnValue);
					break;
			}

			//default empty lookup values to null. Bug 13879
			if ((gridModel[i].type == "lookup") && (columnValue.length == 0)) columnValue = "null";

			sUpdateOrInsert += gridModel[i].id + "\t" + columnValue + "\t";
		}
	}

	var frmDataArea = OpenHR.getForm("dataframe", "frmGetData");
	frmDataArea.txtAction.value = "SAVE";

	//get record id. if it's zero, get new.
	if (selectedRecordID() == "0") {
		frmDataArea.txtReaction.value = "REFRESHFINDAFTERINSERT";
	}

	frmDataArea.txtCurrentViewID.value = $("#txtCurrentViewID").val();;
	frmDataArea.txtCurrentTableID.value = $("#txtCurrentTableID").val();
	frmDataArea.txtRealSource.value = $("#txtRealSource").val();
	if (gridData.ID == "0") { //New record
		frmDataArea.txtRecordID.value = "0";
		var realSource = $('#frmFindForm #txtRealSource').val();
		sUpdateOrInsert = realSource + "\t" + "0\t\t" + sUpdateOrInsert;
	} else { //Update record
		frmDataArea.txtRecordID.value = gridData.ID;
	}

	//	See if we are a history screen and if we are save away the id of the parent also
	if (Number($("#frmFindForm #txtCurrentParentTableID").val()) > 0) {
		sUpdateOrInsert += "ID_" + $.trim($("#frmFindForm #txtCurrentParentTableID").val());
		sUpdateOrInsert += "\t" + $.trim($("#frmFindForm #txtCurrentParentRecordID").val()) + "\t";
	}

	frmDataArea.txtDefaultCalcCols.value = "";	
	frmDataArea.txtInsertUpdateDef.value = sUpdateOrInsert;
	frmDataArea.txtTimestamp.value = gridData.Timestamp;
	frmDataArea.txtOriginalRecordID.value = 0; //This is NOT a copied record

	window.savedRow = rowId;

	//NB: submitform frmData will set a new ID for a new record. 
	OpenHR.submitForm(frmDataArea, null, true, null, null, submitFollowOn);	//leave as async=true to enable the spinner.
	
}

function submitFollowOn() {

	var rowId = window.savedRow; //$("#findGridTable").getGridParam('selrow');	
	if ($('#frmData #txtErrorMessage').val() !== "") { //There was an error while saving (AKA server side validation fail)		
		indicateThatRowWasModified();		

		//After a brief timeout, enable "Add" and "Edit" and disable "Save" and "Cancel"
		setTimeout(function () {
			$("#findGridTable").jqGrid('setSelection', rowId, true);
			$("#findGridTable").editRow(rowId); //Edit the row

			$("#findGridTable_ilsave").removeClass('ui-state-disabled'); //Enable the Save button because we edited something
			$("#findGridTable_ilcancel").removeClass('ui-state-disabled'); //Enable the Cancel button because we edited something
			$("#findGridTable_iledit").addClass('ui-state-disabled'); //Enable the Cancel button because we edited something

			$("#findGridTable_searchButton").addClass("ui-state-disabled");
		}, 100);
	
		
		//Disable navigation buttons on the jqgrid toolbar
		$('#pager-coldata_center input').prop('disabled', true); //Make Page textbox read only
		$("#findGridTable").jqGrid("setGridParam", { ondblClickRow: function (rowID) { return false; } }); //Disable double click on any row
	} else {
		//Mark row as changed if we've successfully saved the record.
		try {
			updateRowFromDatabase(rowId); //Get the row data from the database (show calculated values etc)
			$("#findGridTable #" + rowId + ">td:first").css('border-left', '4px solid green');
			rowWasModified = false; //The 'rowWasModified' variable is defined as global in Find.ascx
			$("#findGridTable_ilsave").addClass('ui-state-disabled'); //Disable the Save button.
			window.onbeforeunload = null;
			$("#findGridTable").jqGrid("setGridParam", { ondblClickRow: function (rowID) { menu_editRecord(); } }); //Enable double click on any row

			//Reevaluate the conditions for the grid's editability
			var recCountInGrid = $("#findGridTable").getGridParam("reccount");
			if (thereIsAtLeastOneEditableColumn && recCountInGrid > 0) {
				$("#findGridTable_iledit").show();
			} else {
				$("#findGridTable_iledit").hide();
			}

			refreshRecordCount();

		} catch (e) {
			OpenHR.modalMessage("Failed to reload data for this row.", "");
		}

		editNextRow();

	}
}

function updateRowFromDatabase(rowid) {
	var recordID = $("#findGridTable").jqGrid('getCell', rowid, 'ID');

	if (Number(recordID) === 0) alert('There was an error reloading the grid.');

	//Get the row from the server
	$.ajax({
		url: "getfindrecordbyid",
		type: "GET",
		cache: false,
		async: false,
		data: { recordid: recordID },
		success: function (jsonstring) {
			var jsondata = JSON.parse(jsonstring);
			var currentRowId = rowid; //The row we need to update (or remove from the view/table)

			//If no data is returned then that means that the row is no longer part of the table/view
			if (jsondata.length == 0) {
				alert('The record saved is no longer in the current view.');
				$('#findGridTable').jqGrid('delRowData', currentRowId);

				refreshRecordCount();

				return false;
			}

			var colModel = $("#findGridTable").jqGrid("getGridParam", "colModel");

			//Loop over the colModel and update the current row
			for (var i = 0; i <= colModel.length - 1; i++) {
				var colNameForColmodel = colModel[i].name.replace(/ /g, '_'); //Replace space by '_' to match the column name in colModel
				var colNameInternalData = colModel[i].name; //For the internal local data
				var cellValue = jsondata[0][colNameForColmodel];

				//Some datatypes need fettling, as always
				switch (colModel[i].type) {
					case "date":
						if (!cellValue == "") { //If the value is not empty then format it as the current date locale
							var d = new Date(cellValue);
							cellValue = d.toString(OpenHR.LocaleDateFormat());
						}
						break;
				}

				//Change each cell in the visible part of the row
				$("#findGridTable").jqGrid('setCell', currentRowId, colNameInternalData, cellValue);

				//Change the internal local data
				$("#findGridTable").jqGrid('getLocalRow', currentRowId)[colNameInternalData] = cellValue;
			}
			//For 'NEW' records assign new ID to the row.
			if (currentRowId == "0") {
				$("#findGridTable #0").attr("ID", recordID);
				lastRowEdited = recordID;

				var frmDataArea = OpenHR.getForm("dataframe", "frmGetData");
				frmDataArea.txtReaction.value = "";

				locateRecord(recordID);
			}

			//refresh menu!
			menu_refreshMenu();

		},
		error: function (e) {
			alert('error updating row from database.');
		}
	});


	addNextRow();

}

function editFindGridRow(rowid) {

	$('#findGridTable').jqGrid('setGridParam', {
		beforeSelectRow: function (newRowid) {
			return beforeSelectFindGridRow(newRowid);
		}
	});


	$('#findGridTable_searchButton').addClass('ui-state-disabled'); //Disable search
	//Disable navigation buttons on the jqgrid toolbar
	$('#pager-coldata_center input').prop('disabled', true); //Make Page textbox read only
	$("#findGridTable").jqGrid("setGridParam", { ondblClickRow: function (rowID) { return false; } }); //Disable double click on any row

	if (Number(rowid) == 0) {
		//we're re-editing a newly created row where the save failed
		rowIsEditedOrNew = "new";
	} else {
		rowIsEditedOrNew = "edited";
		//re-enable add button and highlight new row.
		setTimeout(function () {
			$('#findGridTable_iladd').removeClass('ui-state-disabled');
		}, 100);

	}
}


function addFindGridRow(rowid) {

	$('#findGridTable').jqGrid('setGridParam', {
		beforeSelectRow: function (newRowid) {
			return beforeSelectFindGridRow(newRowid);
		}
	});

	$('#findGridTable_searchButton').addClass('ui-state-disabled'); //Disable search
	//Disable navigation buttons on the jqgrid toolbar
	$('#pager-coldata_center input').prop('disabled', 'true'); //Make Page textbox read only
	$("#findGridTable").jqGrid("setGridParam", { ondblClickRow: function (rowID) { return false; } }); //Disable double click on any row
	rowIsEditedOrNew = "new";

	//re-enable add button and highlight new row.
	setTimeout(function () {
		$('#findGridTable_iladd').removeClass('ui-state-disabled');
		$("#findGridTable").jqGrid('setSelection', "0", true);
	}, 100);

	return true;

}

function cancelFindGridRow(rowid) {
	
	if (rowIsEditedOrNew != "new") { // Not in new record mode.
		updateRowFromDatabase(rowid); //Get the row data from the database
	}

	rowWasModified = false; //The 'rowWasModified' variable is defined as global in Find.ascx
	$("#findGridTable_ilsave").addClass('ui-state-disabled'); //Disable the Save button.
	window.onbeforeunload = null;

	$('#findGridTable_searchButton').removeClass('ui-state-disabled'); //Enable search
	//Enable navigation buttons on the jqgrid toolbar
	$('#pager-coldata_center input').prop('disabled', false); //Remove read only attribute from Page textbox
	$("#findGridTable").jqGrid("setGridParam", { ondblClickRow: function (rowID) { menu_editRecord(); } }); //Enable double click on any row

	rowIsEditedOrNew = "";

	if (rowid == "0") {
		//set selection to last row in grid as the 'new' record has now been removed.
		var recCount = $("#findGridTable").getGridParam("reccount") - 1;
		if (recCount > 0) {
			var lastRowID = $("#findGridTable").jqGrid('getDataIDs')[recCount - 1];
			setTimeout(function() { $("#findGridTable").jqGrid('setSelection', lastRowID, true); }, 200);
		}
	} else {
		//set selection to current row.
		$("#findGridTable").jqGrid('setSelection', rowid, true);
	}

	refreshInlineNavIcons();
}

function beforeSelectFindGridRow(newRowid) {

	if (lastRowEdited == newRowid) return true; //click in same row: allowed.
	if (rowIsEditedOrNew == "") return true;	// not in edit mode: allowed.

	//All checks done, ready to move into Quick Edit mode.
	//Save previous row, then move on to newly clicked row.

	//disable aftersavefunc being called by 'saverow'. We save manually.
	var saveparameters = {
		"successfunc": null,
		"url": null,
		"extraparam": {},
		"aftersavefunc": null,
		"errorfunc": null,
		"afterrestorefunc": null,
		"restoreAfterError": true,
		"mtype": "POST"
	}

	$('#findGridTable').saveRow(lastRowEdited, saveparameters);
	rowIsEditedOrNew = "quickedit_" + newRowid;
	saveRowToDatabase(lastRowEdited);	

	return true;	//always allow row change.
}

function afterSaveFindGridRow(rowid) {
	menu_ShowWait("Saving record...");
	saveRowToDatabase(rowid);
	rowIsEditedOrNew = "";

	return true;
}


function editNextRow() {
	//set the newly selected row to 'edit' mode.
	if (rowIsEditedOrNew.substr(0, 9) == 'quickedit') {
		try {
			var newRowId = rowIsEditedOrNew.substr(10);
			$("#findGridTable").jqGrid('editRow', newRowId);
			lastRowEdited = newRowId;
		} catch (e) {
			alert('couldnt do it!');
		}
	}
}

function addNextRow() {
	if (rowIsEditedOrNew == "new") {
		//quick-add mode
		try {
			var lastPage = $("#findGridTable").jqGrid('getGridParam', 'lastpage');
			$("#findGridTable").trigger("reloadGrid", [{ page: lastPage }]);

			$("#findGridTable").jqGrid('addRow', addparameters);
			lastRowEdited = "0";

			//show editing buttons.
			setTimeout(function () {
				$("#findGridTable_ilsave").removeClass('ui-state-disabled'); //Enable the Save button because we edited something
				$("#findGridTable_ilcancel").removeClass('ui-state-disabled'); //Enable the Cancel button because we edited something
				$("#findGridTable_iledit").addClass('ui-state-disabled'); //Enable the Cancel button because we edited something
				$("#findGridTable_searchButton").addClass("ui-state-disabled");
			}, 100);

		} catch (e) {
		}
	}
}


function refreshInlineNavIcons() {
	//needs the delay; jqGrid may be slow to load.
	setTimeout(function () {
		var selectionMade = (Number(selectedRecordID()) > 0);
		var isSearching = $('#frmFindForm .ui-search-toolbar').is(':visible');
		$("#findGridTable_iledit").toggleClass('ui-state-disabled', (isSearching || !selectionMade));
	}, 100);
}

function refreshRecordCount() {
	
	//Update the record count caption
	var recCount = $("#findGridTable").getGridParam("reccount");
	$('#txtTotalRecordCount').val(recCount);
	if (recCount == 0) {
		menu_SetmnutoolRecordPositionCaption("No Records");
	} else {
		menu_SetmnutoolRecordPositionCaption("Record(s) : " + recCount);
	}
}