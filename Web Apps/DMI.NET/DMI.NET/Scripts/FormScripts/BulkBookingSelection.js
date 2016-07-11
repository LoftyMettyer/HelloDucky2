"use strict";

$(document).ready(function () {

	//jQuery styling
	$("input[type=submit], input[type=button], button").button();
	$("input").addClass("ui-widget ui-corner-all");
	$("input").removeClass("text");


	$("select").addClass("ui-widget ui-corner-tl ui-corner-bl");
	$("select").removeClass("text");
	$("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
	$("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");

	$('#selectOrder, #selectView').change(function () { refreshData(); });

	refreshData();

	menu_refreshMenu();

	$('#tbBBS_OK').button('disable');


});

function resizeGrid() {
	//resize grid			
	var workPageHeight = $('#BulkBookingSelection.optiondatagridpage').height();
	var pageTitleHeight = $('#BulkBookingSelection.optiondatagridpage>.pageTitleDiv').height();
	var dropdownHeight = $('#BulkBookingSelection .floatleft').height();

	var gridMarginBottom = 60;

	var newGridHeight = workPageHeight - pageTitleHeight - dropdownHeight - gridMarginBottom;

	$("#BulkBookingSelection #ssOleDBGridSelRecords").jqGrid('setGridHeight', newGridHeight);
}

function refreshData() {

	var tableID = $("#txtTableID").val();
	var viewID = $("#selectView").val();
	var orderID = $("#selectOrder").val();
	var pageAction = $('#txtPageAction').val();

	OpenHR.ResetSession(); //Reset the session so it doesn't timeout	

	$.getJSON("BulkBookingSelectionData", { tableID: tableID, viewID: viewID, orderID: orderID, pageAction: pageAction })
		.done(function (jsonData) {
			refreshGrid(jsonData);
		})
		.fail(function (jqxhr, textStatus, error) {
			var err = textStatus + ", " + error;
			console.log("Request Failed: " + err);
		});

	tbBBS_RefreshControls();

}

function refreshGrid(jsonData) {
	//need this as this grid won't accept live changes :/		
	$("#ssOleDBGridSelRecords").jqGrid('GridUnload');

	var colMode = [];
	var colNames = [];

	// Configure the grid columns.
	if (jsonData != null) {
		var gridColDef = $.parseJSON(JSON.stringify(jsonData)).colDef;
		var dateFormat = OpenHR.getLocaleDateString();
		
		for (var sColumnName in gridColDef) {
			if (gridColDef.hasOwnProperty(sColumnName)) {

				var aColumnType = gridColDef[sColumnName].split('\t');
				var sColumnType = aColumnType[0];

				//Add column Name to grid
				colNames.push(OpenHR.replaceAll(sColumnName, "_", " "));

				if (sColumnName == "ID") {
					colMode.push({ name: sColumnName, hidden: true });
				}
				else {
					switch (sColumnType) {
						case "Boolean":
							colMode.push({ name: sColumnName, edittype: "checkbox", formatter: 'checkbox', formatoptions: { disabled: true }, align: 'center', width: 100 });
							break;
						case "Int32":
							colMode.push({ name: sColumnName, edittype: "numeric", sorttype: 'integer', formatter: 'number', formatoptions: { disabled: true }, align: 'right', width: 100 });
							break;
						case "Decimal":
							var numDecimals = Number(aColumnType[1]);
							var sThousandSeparator = (aColumnType[2] === 'true') ? OpenHR.LocaleThousandSeparator() : "";
							colMode.push({ name: sColumnName, edittype: "numeric", sorttype: 'integer', formatter: 'number', formatoptions: {defaultValue: "", thousandsSeparator: sThousandSeparator, decimalSeparator: OpenHR.LocaleDecimalSeparator(), decimalPlaces: numDecimals, disabled: true }, align: 'right', width: 100 });
							break;
						case "DateTime":
							colMode.push({ name: sColumnName, edittype: "date", sorttype: 'date', formatter: 'date', formatoptions: { srcformat: dateFormat, newformat: dateFormat, disabled: true }, align: 'left', width: 100 });
							break;
						default:	//text
							colMode.push({ name: sColumnName, width: 100 });
					}
				}
			}
		}


		// Add the grid records.
		var colData = $.parseJSON(JSON.stringify(jsonData)).rows;


		//create the column layout:
		var shrinkToFit = false;

		var dialogWidth = ((screen.width) / 2) - 59;	//40 margin, 19 scroll bar.
		if (((colMode.length - 1) * 100) < dialogWidth) shrinkToFit = true;

		//If no columns available then set shrinkToFit to false. This will show the grid with full width having no columns.
		if (colData.length == 0) {
			shrinkToFit = false;
		}

		$("#ssOleDBGridSelRecords").jqGrid({
			autoencode: true,
			data: colData,
			datatype: "local",
			colNames: colNames,
			colModel: colMode,
			rowNum: 1000,
			multiselect: true,
			autowidth: true,
			shrinkToFit: shrinkToFit,
			beforeSelectRow: handleMultiSelect, // handle multi select
			onSelectRow: function () {
				tbBBS_RefreshControls();
			},
			loadComplete: function () { resizeGrid(); },
			pager: $('#ssOLEDBPager'),
			ondblClickRow: function () {
				tbBBS_OKClick();
			},
			onSortCol: function () {
				resetSelection();
			}
		}).jqGrid('hideCol', 'cb');
		

		$("#ssOleDBGridSelRecords").jqGrid('bindKeys', {
			"onEnter": function () {
				tbBBS_OKClick();
			}
		});

		resetSelection();
		$("#BulkBookingSelection #ssOleDBGridSelRecords").jqGrid('setGridWidth', $('#BulkBookingSelection.optiondatagridpage>.pageTitleDiv').width());
	}

}

function resetSelection() {
	$('#ssOleDBGridSelRecords').jqGrid('resetSelection');
	button_disable($('#tbBBS_OK'), true);
}


function tbBBS_OKClick() {
	var sSelectedRows = $('#ssOleDBGridSelRecords').jqGrid('getGridParam', 'selarrrow');
	var sSelectedIDs = "";
	for (var iIndex = 0; iIndex < sSelectedRows.length; iIndex++) {
		var sRecordID = $("#ssOleDBGridSelRecords").jqGrid('getCell', sSelectedRows[iIndex], 'ID');
		if (sSelectedIDs.length > 0) {
			sSelectedIDs = sSelectedIDs + ",";
		}
		sSelectedIDs = sSelectedIDs + sRecordID;
	}
	bulkbooking_makeSelection("ALL", 0, sSelectedIDs);

	$('#BulkBookingSelect').dialog("close");
}


function tbBBS_RefreshControls() {

	var selRowId = $("#ssOleDBGridSelRecords").jqGrid('getGridParam', 'selrow');
	var fNoneSelected = (selRowId == null || selRowId == 'undefined');

	button_disable($('#tbBBS_OK'), fNoneSelected);
}
