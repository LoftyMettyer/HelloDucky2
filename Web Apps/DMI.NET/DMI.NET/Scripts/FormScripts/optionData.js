﻿
function optiondata_onload() {

	var frmOptionData = document.getElementById("frmOptionData");

	var sFatalErrorMsg = frmOptionData.txtErrorDescription.value;
	if (sFatalErrorMsg.length > 0) {
		//window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sFatalErrorMsg);
		//window.parent.location.replace("login.asp");
		$('#FindGridRow').css('border', '1px solid silver');
	} else {
		// Do nothing if the menu controls are not yet instantiated.
		var sCurrentWorkPage = OpenHR.currentWorkPage();
		
		//To allow option frame to pop out with jQuery dialog control...
		var sOptionWorkPage = $("#optionframe").attr("data-framesource");
		if (sCurrentWorkPage == "RECORDEDIT") {
			switch (sOptionWorkPage) {
			case 'LOOKUPFIND':
				sCurrentWorkPage = "LOOKUPFIND";
				break;
			case 'LINKFIND':
				sCurrentWorkPage = "LINKFIND";
				break;
			default:
			}
		}
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

		var dateFormat = OpenHR.getLocaleDateString();

		if (sCurrentWorkPage == "LINKFIND") {

			sErrorMsg = frmOptionData.txtErrorMessage.value;
			if (sErrorMsg.length > 0) {
				// We've got an error so don't update the record edit form.

				// Get menu.asp to refresh the menu.
				//menu_refreshMenu();
				OpenHR.messageBox(sErrorMsg);
			}
			sAction = frmOptionData.txtOptionAction.value; // Refresh the link find grid with the data if required.
			//var grdLinkFind = OpenHR.getForm("optionframe", "frmLinkFindForm").ssOleDBGridLinkRecords;

			linkFind_removeAll('ssOleDBGridLinkRecords');	// Clear the grid.

			dataCollection = frmOptionData.elements; // Configure the grid columns.
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

							if (sColumnName == "ID") {
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

			// Add the grid records.
			fRecordAdded = false;
			iCount = 0;
			if (dataCollection != null) {
				colData = [];
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 14);
					if (sControlName == "txtOptionData_") {
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

				//create the column layout:
				var shrinkToFit = false;
				if (colMode.length < 8) shrinkToFit = true;

				$("#ssOleDBGridLinkRecords").jqGrid({
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colMode,
					rowNum: 1000,
					autowidth: true,
					shrinkToFit: shrinkToFit,
					onSelectRow: function () {
						linkFind_refreshControls();
					},
					ondblClickRow: function () {
						SelectLink();
					}
				});
				
				$("#ssOleDBGridLinkRecords").jqGrid('bindKeys', {
					"onEnter": function () {
						SelectLink();
					}
				});

				//resize the grid to the height of its container.
				$("#ssOleDBGridLinkRecords").jqGrid('setGridHeight', $("#linkFindGridRow").height());
				//var y = $("#gbox_ssOleDBGrid").height();
				//var z = $('#gbox_ssOleDBGrid .ui-jqgrid-bdiv').height();

			}

			//NOTE: may come in useful.
			//http://stackoverflow.com/questions/12572780/jqgrids-addrowdata-hangs-for-large-number-of-records

			frmOptionData.txtRecordCount.value = iCount;

			if (fRecordAdded == true) {
				locateRecord(OpenHR.getForm("optionframe", "frmLinkFindForm").txtOptionLinkRecordID.value, true); //should be in scope!
			}

			linkFind_refreshControls();  ///should be in scope - from lookupFind.ascx


			// Get menu.asp to refresh the menu.
			menu_refreshMenu();
		}
		var grdFind;
		if (sCurrentWorkPage == "LOOKUPFIND") {
			sErrorMsg = frmOptionData.txtErrorMessage.value;
			if (sErrorMsg.length > 0) {
				// We've got an error so don't update the record edit form.

				// Get menu.asp to refresh the menu.
				//disabled as we pop out the grid now, so no toolbar...
				//menu_refreshMenu();
				OpenHR.messageBox(sErrorMsg);
			}

			if (frmOptionData.txtFilterOverride.value == "True")
				// No access to the lookup filter column?
			{
				OpenHR.messageBox("You do not have 'read' permission on the lookup filter value column. No filter will be applied.");
			}
			sAction = frmOptionData.txtOptionAction.value;
			OpenHR.getForm("optionframe", "frmLookupFindForm").txtLookupColumnGridPosition.value = frmOptionData.txtLookupColumnGridPosition.value;

			// Refresh the link find grid with the data if required.
			grdFind = OpenHR.getForm("optionframe", "frmLookupFindForm").ssOleDBGrid; 

			lookupFind_removeAll('ssOleDBGrid');	// Clear the grid.

			dataCollection = frmOptionData.elements; // Configure the grid columns.
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

							if (sColumnName == "ID") {
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
										break;
								}
							}

						}
					}
				}
			}

			// Add the grid records.
			fRecordAdded = false;
			iCount = 0;		        //used to store record count later...
			if (dataCollection != null) {
				colData = [];
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 14);
					if (sControlName == "txtOptionData_") {
						//original line
						//grdFind.addItem(dataCollection.item(i).value);
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

				//create the column layout:
				shrinkToFit = false;
				if (colMode.length < 8) shrinkToFit = true;
				
				$("#ssOleDBGrid").jqGrid({
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colMode,
					rowNum: 30,
					//autowidth: true,
					width: 700,
					ignoreCase: true,
					pager: $('#ssOLEDBPager'),
					shrinkToFit: shrinkToFit,
					onSelectRow: function () {
						lookupFind_refreshControls();
					},
					ondblClickRow: function () {
						SelectLookup();
					}
				});

				//search options.
				$("#ssOleDBGrid").jqGrid('navGrid', '#ssOLEDBPager', { del: false, add: false, edit: false, search: false });
				
				$("#ssOleDBGrid").jqGrid('navButtonAdd', "#ssOLEDBPager", {
					caption: '',
					buttonicon: 'ui-icon-search',
					onClickButton: function () {
						$("#ssOleDBGrid").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });
					},
					position: 'first',
					title: '',
					cursor: 'pointer'
				});

				//$("#ssOleDBGrid").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });

				$("#ssOleDBGrid").jqGrid('bindKeys', {
					"onEnter": function () {
						SelectLookup();
					}
				});

				//resize the grid to the height of its container.
				$("#ssOleDBGrid").jqGrid('setGridHeight', $("#lookupFindGridRow").height());
				//var y = $("#gbox_ssOleDBGrid").height();
				//var z = $('#gbox_ssOleDBGrid .ui-jqgrid-bdiv').height();

			}

			//NOTE: may come in useful.
			//http://stackoverflow.com/questions/12572780/jqgrids-addrowdata-hangs-for-large-number-of-records

			frmOptionData.txtRecordCount.value = iCount;

			if (fRecordAdded == true) {
				locateRecord(OpenHR.getForm("optionframe", "frmLookupFindForm").txtOptionLookupValue.value, true);
			}

			lookupFind_refreshControls();

			// Get menu.asp to refresh the menu.
			//no longer required - we pop out the menu now...
			//menu_refreshMenu();

			//select top row.
			if (lookupFind_rowCount() > 0) {
					// Select the top row.
					lookupFind_moveFirst();
			}

		}


		if ((sCurrentWorkPage == "TBTRANSFERCOURSEFIND") ||
				(sCurrentWorkPage == "TBBOOKCOURSEFIND") ||
				(sCurrentWorkPage == "TBADDFROMWAITINGLISTFIND") ||
				(sCurrentWorkPage == "TBTRANSFERBOOKINGFIND")) {
			sErrorMsg = frmOptionData.txtErrorMessage.value;
			if (sErrorMsg.length > 0) {
				// We've got an error.
				// Get menu.asp to refresh the menu.
				menu_refreshMenu();
				OpenHR.messageBox(sErrorMsg);
			}

			if ((sCurrentWorkPage == "TBTRANSFERBOOKINGFIND") ||
					(sCurrentWorkPage == "TBADDFROMWAITINGLISTFIND")) {
				sErrorMsg = frmOptionData.txtErrorMessage2.value;
				if (sErrorMsg.length > 0) {
					// We've got an error.
					$('#cmdCancel').click(); //should be in scope!
					//window.parent.frames("menuframe").ASRIntranetFunctions.ClosePopup();
					OpenHR.messageBox(sErrorMsg);
					return false;
				}
			}
			sAction = frmOptionData.txtOptionAction.value; // Refresh the link find grid with the data if required.

			//need this as this grid won't accept live changes :/		
			$("#ssOleDBGridRecords").jqGrid('GridUnload');

			//jqGrid_removeAll('ssOleDBGridRecords');	// clear the grid
			
			dataCollection = frmOptionData.elements; // Configure the grid columns.
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

							if (sColumnName == "ID") {
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

			// Add the grid records.
			fRecordAdded = false;
			iCount = 0;
			if (dataCollection != null) {

				colData = [];
				for (i = 0; i < dataCollection.length; i++) {

					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 14);
					if (sControlName == "txtOptionData_") {
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


				//create the column layout:
				var shrinkToFit = false;
				if (colMode.length < 8) shrinkToFit = true;

				$("#ssOleDBGridRecords").jqGrid({
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colMode,
					rowNum: 1000,
					autowidth: true,
					shrinkToFit: shrinkToFit,
					onSelectRow: function () {
						tbrefreshControls();
					},
					ondblClickRow: function () {
						$('#cmdSelect').click();
					}
				});


				$("#ssOleDBGridRecords").jqGrid('bindKeys', {
					"onEnter": function () {
						$('#cmdSelect').click();
					}
				});

				//resize the grid to the height of its container.
				$("#ssOleDBGridRecords").jqGrid('setGridHeight', $("#FindGridRow").height());

			}
		
			frmOptionData.txtRecordCount.value = iCount;

			if (fRecordAdded == true) {
				locateRecord($('#RecordID').val(), true); //should be in scope!
			}

			tbrefreshControls();

			// Get menu.asp to refresh the menu.
			menu_refreshMenu();
		}

		if (sCurrentWorkPage == "BULKBOOKING") {

			frmOptionData = document.getElementById("frmOptionData");

			sAction = frmOptionData.txtOptionAction.value;

			//need this as this grid won't accept live changes :/		
			$("#ssOleDBGridFindRecords").jqGrid('GridUnload');

			dataCollection = frmOptionData.elements;

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
			//

			// Add the grid records.
			fRecordAdded = false;
			iCount = 0;

			if (dataCollection != null) {

				colData = [];
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 14);

					if (sControlName == "txtOptionData_") {
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

				//create the column layout:
				var shrinkToFit = false;
				if (colMode.length < 8) shrinkToFit = true;				

				$("#ssOleDBGridFindRecords").jqGrid({
					multiselect: true,
					data: colData,
					datatype: 'local',
					colNames: colNames,
					colModel: colMode,
					rowNum: 1000,
					autowidth: true,
					shrinkToFit: shrinkToFit,
					onSelectRow: function () {
						tbrefreshControls();
					},
					editurl: 'clientArray',
					afterShowForm: function ($form) {
						$("#dData", $form.parent()).click();
					},
					beforeSelectRow: handleMultiSelect // handle multi select
				}).jqGrid('hideCol', 'cb');

				//resize the grid to the height of its container.
				$("#ssOleDBGridFindRecords").jqGrid('setGridHeight', $("#FindGridRow").height());

			}			


			// Select the top record.
			if (fRecordAdded == true) {
				moveFirst();
			}

			tbrefreshControls();

			// Get menu.asp to refresh the menu.
			menu_refreshMenu();
		}

		if (sCurrentWorkPage == "UTIL_DEF_PICKLIST") {
			sAction = frmOptionData.txtOptionAction.value; // Refresh the link find grid with the data if required.

			dataCollection = frmOptionData.elements; // Add the grid records.
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

			if (dataCollection != null) {
				colData = [];
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 14);

					if (sControlName == "txtOptionData_") {
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
			}

			//Determine if the grid already exists...
			if ($("#ssOleDBGrid").getGridParam("reccount") == undefined) { //It doesn't exist, create it
				var shrinkToFit = false;
				if (colMode.length < 8) shrinkToFit = true;

				$("#ssOleDBGrid").jqGrid({
					multiselect: true,
					data: colData,
					datatype: 'local',
					colNames: colNames,
					colModel: colMode,
					rowNum: 1000,
					autowidth: true,
					shrinkToFit: shrinkToFit,
					onSelectRow: function() {
						button_disable(frmDefinition.cmdRemove, false); //Enable the "Remove" button
						button_disable(frmDefinition.cmdRemoveAll, false); //Enable the "Remove All" button
					},
					editurl: 'clientArray',
					afterShowForm: function($form) {
						$("#dData", $form.parent()).click();
					},
					beforeSelectRow: handleMultiSelect // handle multi select
				}).jqGrid('hideCol', 'cb');

				//resize the grid to the height of its container.
				$("#ssOleDBGrid").jqGrid('setGridHeight', $("#PickListGrid").height());

	 			if ($("#ssOleDBGrid").jqGrid().width() < $("#PickListGrid").width()) {
					 $("#ssOleDBGrid").parent().parent().addClass('jqgridHideHorScroll');
				 }

				// Select the top record.
				if (fRecordAdded == true) {
					$("#ssOleDBGrid").jqGrid('setSelection', colData[0].id);
				}
			} else { // The grid exists, add rows to it
				for (var j = 0; j <= colData.length - 1; j++) {
					$("#ssOleDBGrid").addRowData(colData[j].id, colData[j], 'last');
				}
			}

			//Display the number of records
			$('#RecordCountDIV').html($("#ssOleDBGrid").getGridParam('reccount') + " Record(s)");

			if (frmOptionData.txtExpectedCount.value > iCount) {
				if (iCount == 0) {
					OpenHR.messageBox("You do not have 'read' permission on any of the records in the selected picklist.\nUnable to display records.");
					OpenHR.getForm("workframe", "frmUseful").txtAction.value = "VIEW";
					OpenHR.getFrame("workframe").cancelClick();
				} else {
					if (OpenHR.getForm("workframe", "frmUseful").txtAction.value.toUpperCase() == "COPY") {
						OpenHR.messageBox("You do not have 'read' permission on all of the records in the selected picklist.\nOnly permitted records will be shown.");
					} else {
						OpenHR.messageBox("You do not have 'read' permission on all of the records in the selected picklist.\nOnly permitted records will be shown and the definition will be read only.");
						OpenHR.getForm("workframe", "frmUseful").txtAction.value = "VIEW";
						OpenHR.getFrame("workframe").disableAll();
					}
				}
			}

			refreshControls();

			// Get menu.asp to refresh the menu.
			menu_refreshMenu();
		}

		if (sCurrentWorkPage == "UTIL_DEF_EXPRCOMPONENT") {
			sAction = frmOptionData.txtOptionAction.value;
			if ((sAction == "LOADEXPRFIELDCOLUMNS") ||
						(sAction == "LOADEXPRLOOKUPCOLUMNS")) {
				dataCollection = frmOptionData.elements;
				if (dataCollection != null) {
					for (i = 0; i < dataCollection.length; i++) {
						sControlName = dataCollection.item(i).name;
						sControlName = sControlName.substr(0, 10);
						if (sControlName == "txtColumn_") {
							component_addColumn(dataCollection.item(i).value);
						}
					}
				}

				component_setColumn(frmOptionData.txtOptionColumnID.value);
			}

			if (sAction == "LOADEXPRLOOKUPVALUES") {
				dataCollection = frmOptionData.elements;
				if (dataCollection != null) {
					for (i = 0; i < dataCollection.length; i++) {
						sControlName = dataCollection.item(i).name;
						sControlName = sControlName.substr(0, 9);
						if (sControlName == "txtValue_") {
							component_addValue(dataCollection.item(i).value);
						}
					}
				}

				component_setValue(frmOptionData.txtOptionLocateValue.value);
			}

			// Get menu.asp to refresh the menu.
			menu_refreshMenu();
		}

		if (sCurrentWorkPage == "FIND") {
			sAction = frmOptionData.txtOptionAction.value;
			if ((sAction == "BOOKCOURSEERROR") ||
						(sAction == "TRANSFERBOOKINGERROR") ||
						(sAction == "ADDFROMWAITINGLISTERROR") ||
						(sAction == "BULKBOOKINGERROR")) {
				OpenHR.messageBox(frmOptionData.txtNonFatalErrorDescription.value);
			}
			if ((sAction == "BOOKCOURSESUCCESS") ||
					(sAction == "TRANSFERBOOKINGSUCCESS") ||
					(sAction == "ADDFROMWAITINGLISTSUCCESS") ||
					(sAction == "BULKBOOKINGSUCCESS")) {
				// Reload the find records.
				OpenHR.messageBox("Booking(s) made successfully.");
				menu_reloadFindPage("MOVEFIRST", "");
			}
		}

	}
}

function refreshOptionData() {
	var frmGetOptionData = document.getElementById("frmGetOptionData");
	OpenHR.submitForm(frmGetOptionData);
}


function jqGrid_removeAll(jqGridID) {
	//remove all rows from the jqGrid.
	$('#' + jqGridID).jqGrid('clearGridData');
}


// handle jqGrid multiselect => thanks to solution from Byron Cobb on http://goo.gl/UvGku
var handleMultiSelect = function (rowid, e) {
	var grid = $(this);
	if (!e.ctrlKey && !e.shiftKey) {
		grid.jqGrid('resetSelection');
	}
	else if (e.shiftKey) {
		var initialRowSelect = grid.jqGrid('getGridParam', 'selrow');		
		grid.jqGrid('resetSelection');

		var CurrentSelectIndex = grid.jqGrid('getInd', rowid);
		var InitialSelectIndex = grid.jqGrid('getInd', initialRowSelect);
		var startID = "";
		var endID = "";
		if (CurrentSelectIndex > InitialSelectIndex) {
			startID = initialRowSelect;
			endID = rowid;
		}
		else {
			startID = rowid;
			endID = initialRowSelect;
		}
		var shouldSelectRow = false;

		$.each(grid.getDataIDs(), function (_, id) {
			if ((shouldSelectRow = id == startID || shouldSelectRow)) {
				grid.jqGrid('setSelection', id, false);
			}
			return id != endID;

		});

		//last selected row too
		grid.jqGrid('setSelection', endID, false);

	}
	return true;
};
