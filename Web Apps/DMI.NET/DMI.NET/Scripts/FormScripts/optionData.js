
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
									case "int32":
										colMode.push({ name: sColumnName, edittype: "integer", sorttype: 'integer', formatter: 'integer', formatoptions: { disabled: true }, align: 'right', width: 100 });
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

				var formHeight = $('#optionframe').outerHeight();
				var navButtonsHeight = $('#navButtons').outerHeight();
				var buttonHeight = $('#divLinkFindButtons').outerHeight();
				var pageTitleHeight = $('.pageTitleDiv').outerHeight();
				var gridHeight = formHeight - navButtonsHeight - buttonHeight - pageTitleHeight - 120;
				var gridWidth = $('#optionframe').outerWidth() - 50;

				$("#ssOleDBGridLinkRecords").jqGrid({
					autoencode: true,
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colMode,
					rowNum: 1000,
					pager: $('#pager-coldata-optiondata'),
					width: gridWidth,
					height: gridHeight,
					shrinkToFit: shrinkToFit,
					onSelectRow: function () {
						linkFind_refreshControls();
					},
					ondblClickRow: function () {
						SelectLink();
					},
					loadComplete: function () {						
						$('#optionframe').dialog({ position: { my: "center", at: "center", of: window } });
					}
				});
				
				$("#ssOleDBGridLinkRecords").jqGrid('bindKeys', {
					"onEnter": function () {
						SelectLink();
					}
				});

				$("#ssOleDBGridLinkRecords").jqGrid('navGrid', '#pager-coldata-optiondata', { del: false, add: false, edit: false, search: false, refresh: false }); // setup the buttons we want

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
					autoencode: true,
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colMode,
					rowNum: 30,
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

				// Navbar options = i.e. search, edit, save etc 
				$("#ssOleDBGrid").jqGrid('navGrid', '#ssOLEDBPager', { del: false, add: false, edit: false, search: false, refresh: false }); // setup the buttons we want
				$("#ssOleDBGrid").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });  //instantiate toolbar so we can use toggle.
				$("#ssOleDBGrid")[0].toggleToolbar();  // Toggle it off at start up.

				$("#ssOleDBGrid").jqGrid('navButtonAdd', "#ssOLEDBPager", {
					caption: '',
					buttonicon: 'ui-icon-search',
					onClickButton: function () {
						$("#ssOleDBGrid")[0].toggleToolbar(); // Toggle toolbar on & off when Search button is pressed.
						$("#ssOleDBGrid")[0].clearToolbar();  // clear menu
						var isSearching = $('#frmLookupFindForm .ui-search-toolbar').is(':visible');
						$("#ssOleDBGrid_iledit").toggleClass('ui-state-disabled', isSearching);
						$("#ssOleDBGrid_iladd").toggleClass('ui-state-disabled', isSearching);
					},
					position: 'first',
					title: '',
					cursor: 'pointer'
				});

				$("#ssOleDBGrid").jqGrid('bindKeys', {
					"onEnter": function () {
						SelectLookup();
					}
				});

				//resize the grid to the height of its container.
				$("#ssOleDBGrid").jqGrid('setGridHeight', $("#lookupFindGridRow").height());
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

			//need this as this grid won't accept live changes :/		
			$("#ssOleDBGridRecords").jqGrid('GridUnload');
			
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
					autoencode: true,
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colMode,
					pager: $('#pager-coldata-optiondata'),
					ignorecase: true,
					rowNum: 100,
					autowidth: true,
					shrinkToFit: shrinkToFit,
					onSelectRow: function () {
						tbrefreshControls();
					},
					ondblClickRow: function () {
						$('#cmdSelect').click();
					},
					loadComplete: function () {
						$("#ssOleDBGridRecords").jqGrid('setSelection', 1);
						tbrefreshControls();
					},
					afterSearch: function () {
						$("#ssOleDBGridRecords").jqGrid('setSelection', 1);
						tbrefreshControls();
					}
				});


				$("#ssOleDBGridRecords").jqGrid('bindKeys', {
					"onEnter": function () {
						$('#cmdSelect').click();
					}
				});

				// Navbar options = i.e. search, edit, save etc 
				$("#ssOleDBGridRecords").jqGrid('navGrid', '#pager-coldata-optiondata', { del: false, add: false, edit: false, search: false, refresh: false }); // setup the buttons we want
				$("#ssOleDBGridRecords").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });  //instantiate toolbar so we can use toggle.
				$("#ssOleDBGridRecords")[0].toggleToolbar();  // Toggle it off at start up.

				$("#ssOleDBGridRecords").jqGrid('navButtonAdd', "#pager-coldata-optiondata", {
					caption: '',
					buttonicon: 'ui-icon-search',
					onClickButton: function () {
						$("#ssOleDBGridRecords").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });
						$("#ssOleDBGridRecords")[0].toggleToolbar(); // Toggle toolbar on & off when Search button is pressed.
						$("#ssOleDBGridRecords")[0].clearToolbar();  // clear menu
						var isSearching = $('#frmLookupFindForm .ui-search-toolbar').is(':visible');
						$("#ssOleDBGridRecords_iledit").toggleClass('ui-state-disabled', isSearching);
						$("#ssOleDBGridRecords_iladd").toggleClass('ui-state-disabled', isSearching);
					},
					position: 'first',
					title: '',
					cursor: 'pointer'
				});

				//resize the grid to the height of its container.
				$("#ssOleDBGridRecords").jqGrid('setGridHeight', $("#FindGridRow").height() - 40);

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

				if (colMode.length > 0) {
					//create the column layout:
					var shrinkToFit = false;
					if (colMode.length < 8) shrinkToFit = true;
					var gridWidth = $('#FindGridRow').width();

					$("#ssOleDBGridFindRecords").jqGrid({
						autoencode: true,
						multiselect: true,
						data: colData,
						datatype: 'local',
						colNames: colNames,
						colModel: colMode,
						rowNum: 1000,
						width: gridWidth,
						//autowidth: true,
						shrinkToFit: shrinkToFit,
						onSelectRow: function () {
							tbrefreshControls();
						},
						editurl: 'clientArray',
						afterShowForm: function ($form) {
							$("#dData", $form.parent()).click();
						},
						loadComplete: function () {
							grid_HideCheckboxes('ssOleDBGridFindRecords');
						},
						beforeSelectRow: handleMultiSelect // handle multi select
					}).jqGrid('hideCol', 'cb');

					//resize the grid to the height of its container.
					$("#ssOleDBGridFindRecords").jqGrid('setGridHeight', $("#FindGridRow").height());
				} 


				// Select the top record.
				if ((fRecordAdded == true) && (colMode.length > 0)) {
					moveFirst();
				}

			}
			tbrefreshControls();

			// Get menu.asp to refresh the menu.
			menu_refreshMenu();
		}

		if (sCurrentWorkPage == "UTIL_DEF_PICKLIST") {
			
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
				var gridWidth = $('#PickListGrid').width();
				$("#ssOleDBGrid").jqGrid({
					autoencode: true,
					multiselect: true,
					data: colData,
					datatype: 'local',
					colNames: colNames,
					colModel: colMode,
					rowNum: 1000,
					width: gridWidth,
					shrinkToFit: shrinkToFit,
					onSelectRow: function() {
						refreshControls();
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

			if (fRecordAdded) frmUseful.txtChanged.value = 1;

			refreshControls();

			// Get menu.asp to refresh the menu.
			menu_refreshMenu();
		}

		if (sCurrentWorkPage == "UTIL_DEF_EXPRCOMPONENT") {
			
			sAction = parseInt(frmOptionData.txtOptionAction.value);
			if (sAction === optionActionType.LOADEXPRFIELDCOLUMNS ||
					sAction === optionActionType.LOADEXPRLOOKUPCOLUMNS) {

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

			if (sAction === optionActionType.LOADEXPRLOOKUPVALUES) {
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

			sAction = parseInt(frmOptionData.txtOptionAction.value);
			if ((sAction === optionActionType.BOOKCOURSEERROR) ||
						(sAction === optionActionType.TRANSFERBOOKINGERROR) ||
						(sAction === optionActionType.ADDFROMWAITINGLISTERROR) ||
						(sAction === optionActionType.BULKBOOKINGERROR)) {
				OpenHR.messageBox(frmOptionData.txtNonFatalErrorDescription.value);
			}

			if ((sAction === optionActionType.BOOKCOURSESUCCESS) ||
					(sAction === optionActionType.TRANSFERBOOKINGSUCCESS) ||
					(sAction === optionActionType.ADDFROMWAITINGLISTSUCCESS) ||
					(sAction === optionActionType.BULKBOOKINGSUCCESS)) {
				// Reload the find records.
				OpenHR.messageBox("Booking(s) made successfully.");
				menu_reloadFindPage("MOVEFIRST", "");
			}
		}

	}
}


function grid_HideCheckboxes(gridID) {
	//Hide the checkboxes.
	if ($('#' + gridID)) {
		$('#' + gridID + '_cb').css('visibility', 'hidden');
		$('#' + gridID + ' .cbox').css('visibility', 'hidden');
	}
}

function refreshOptionData() {
	var frmGetOptionData = document.getElementById("frmGetOptionData");
	OpenHR.submitForm(frmGetOptionData, "optiondataframe");
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
