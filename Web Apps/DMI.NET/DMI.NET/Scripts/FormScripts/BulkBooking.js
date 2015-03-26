"use strict";

$(function () {
	$("#optionframe").attr("data-framesource", "BULKBOOKING");
	$("#workframe").hide();
	$("#optionframe").show();

	resizeGridToFit();

	//Configure Buttons
	$('#cmd_tbBBSelect').click(function () { add(); });
	$('#cmd_tbBBPicklistAdd').click(function () { getRecordGroup('PICKLIST'); });
	$('#cmd_tbBBFilteredAdd').click(function () { getRecordGroup('FILTER'); });
	$('#cmd_tbBBCancel').click(function () { cancel(); });
	$('#cmd_tbBBOK').click(function () { ok(); });
	$('#cmd_tbBBRemove').click(function () { remove(); });
	$('#cmd_tbBBRemoveAll').click(function () { removeAll(); });

	tbrefreshControls();

});


function add() {
	$("#BulkBookingSelect").dialog('open');
}


function cancel() {
	$("#optionframe").hide();
	$("#workframe").show();

	menu_refreshMenu();

	var postData = {
		Action: optionActionType.CANCEL,
		__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
	};
	OpenHR.submitForm(null, "optionframe", null, postData, "BulkBooking_Submit");

}

function getRecordGroup(groupType) {
	var tableID = $('#TableID').val();
	var currentID = "";
	var newHeight = (screen.height) / 2;
	var newWidth = (screen.width) / 2;
	var returnFilterResults = false;
	if (groupType == "FILTER") returnFilterResults = true;
	
	OpenHR.modalExpressionSelect(groupType, tableID, currentID, function (id) {
		bulkbooking_makeSelection(groupType, id, '');
	}, newWidth - 40, newHeight - 160, returnFilterResults);
	$('#ExpressionSelectNone').hide();
}

function bulkbooking_makeSelection(psType, piID, psPrompts) {

	/* Get the current selected delegate IDs. */
	var sSelectedIDs;

	sSelectedIDs = $('#ssOleDBGridFindRecords').getDataIDs().join(",");

	if ((psType == "ALL") && (psPrompts.length > 0)) {
		if (sSelectedIDs.length > 0) {
			sSelectedIDs = sSelectedIDs + ",";
		}
		sSelectedIDs = sSelectedIDs + psPrompts;
	}

	if ($(".popup").dialog("isOpen")) $(".popup").dialog("close");


	// Get the optionData.asp to get the required records.
	var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
	optionDataForm.txtOptionAction.value = optionActionType.GETBULKBOOKINGSELECTION;
	optionDataForm.txtOptionPageAction.value = psType;
	optionDataForm.txtOptionRecordID.value = piID;
	optionDataForm.txtOptionValue.value = sSelectedIDs;
	optionDataForm.txtOptionPromptSQL.value = psPrompts;
	optionDataForm.txtOption1000SepCols.value = $('#txt1000SepCols').val();
	refreshOptionData(); //should be in scope.		
}


function ok() {

	var bookingStatus = "B";
	var sSelectedIDs = $('#ssOleDBGridFindRecords').getDataIDs().join(",");

	if ($('#TbStatusPExists').val() == 'True') {
		bookingStatus = $('#selStatus').val();
	} 

	var postData = {
		Action: optionActionType.SELECTBULKBOOKINGS,
		BookingStatus: bookingStatus,
		CourseID: $("#txtOptionRecordID").val(),
		EmployeeIDs: sSelectedIDs,
		__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
	};
	OpenHR.submitForm(null, "optionframe", null, postData, "BulkBooking_Submit");

}

function remove() {

	var grid = $("#ssOleDBGridFindRecords");
	var myDelOptions = {
		// because I use "local" data I don't want to send the changes
		// to the server so I use "processing:true" setting and delete
		// the row manually in onclickSubmit
		onclickSubmit: function (options) {
			var gridID = $.jgrid.jqID(grid[0].id),
					gridP = grid[0].p,
					newPage = gridP.page,
					rowids = gridP.multiselect ? gridP.selarrrow : [gridP.selrow];			

			// reset the value of processing option which could be modified
			options.processing = true;

			// delete the row
			//BUG: Convert the rowids variable to a 'static' new variable.
			var iDsToDelete = rowids.toString().split(',');

			$.each(iDsToDelete, function () {
				grid.delRowData(this);
			});
			$.jgrid.hideModal("#delmod" + gridID,
												{
													gb: "#gbox_" + gridID,
													jqm: options.jqModal, onClose: options.onClose
												});

			if (gridP.lastpage > 1) {// on the multipage grid reload the grid
				if (gridP.reccount === 0 && newPage === gridP.lastpage) {
					// if after deliting there are no rows on the current page
					// which is the last page of the grid
					newPage--; // go to the previous page
				}
				// reload grid to make the row from the next page visable.
				grid.trigger("reloadGrid", [{ page: newPage }]);
			}

			return true;
		},
		processing: true
	};


	//Now select the previous remaining row...
	var firstSelectedRowNumber = 1;
	try {
		//get first row and calculate previous row ID so we can select it after removal
		var firstSelectedRowID = $('#ssOleDBGridFindRecords').jqGrid('getGridParam', 'selarrrow')[0];

		//get row number from rowID
		firstSelectedRowNumber = $('#ssOleDBGridFindRecords #' + firstSelectedRowID)[0].rowIndex;
	} catch (e) { }

	grid.jqGrid('delGridRow', grid.jqGrid('getGridParam', 'selarrrow'), myDelOptions);

	$("#dData").click(); //To remove the "delete confirmation" dialog		
	tbMoveSpecificOrFirst(firstSelectedRowNumber - 1); //deduct one to select previous row.
	tbrefreshControls();
}

function removeAll() {
	$("#ssOleDBGridFindRecords").jqGrid('GridUnload');
	tbrefreshControls();
}


function resizeGridToFit() {

	//resize grid	
	var workPageHeight = $('.optiondatagridpage').height();
	var pageTitleHeight = $('.optiondatagridpage>.pageTitleDiv').height();
	var dropdownHeight = $('.floatleft').height();
	var gridMarginBottom = 70;

	var newGridHeight = workPageHeight - pageTitleHeight - dropdownHeight - gridMarginBottom;

	$("#ssOleDBGridRecords").jqGrid('setGridHeight', newGridHeight);
	$('#FindGridRow').height(newGridHeight);
}

function tbMoveSpecificOrFirst(rowNumber) {
	if ($("#ssOleDBGridFindRecords").getGridParam("reccount") > 0) {
		var specificRowID = $("#ssOleDBGridFindRecords").getDataIDs()[0];	//default to top row.

		//Get previous row by number, or first if selected row is at the top.
		rowNumber = (rowNumber <= 0 ? 0 : rowNumber - 1);

		try {
			specificRowID = $("#ssOleDBGridFindRecords").getDataIDs()[rowNumber];
		} catch (e) { }
		finally {
			$('#ssOleDBGridFindRecords').jqGrid('resetSelection');
			$('#ssOleDBGridFindRecords').jqGrid('setSelection', specificRowID);
		}
	} else {
		//grid is empty.
		$("#ssOleDBGridFindRecords").jqGrid('GridUnload');
	}

	menu_refreshMenu();
}


function tbrefreshControls() {
	var fNoneSelected;
// ReSharper disable once Html.IdNotResolved
	var frmBulkBooking = document.getElementById('frmBulkBooking');
	var fGridHasRows = ($("#ssOleDBGridFindRecords").getGridParam("reccount") > 0);

	var selRowId = $("#ssOleDBGridFindRecords").jqGrid('getGridParam', 'selrow');
	fNoneSelected = (selRowId == null || selRowId == 'undefined');

	button_disable(frmBulkBooking.cmd_tbBBRemove, fNoneSelected);
	button_disable(frmBulkBooking.cmd_tbBBRemoveAll, !fGridHasRows);
	button_disable(frmBulkBooking.cmd_tbBBOK, !fGridHasRows);

	$('#FindGridRow').toggleClass('silverborder', !fGridHasRows);

	menu_refreshMenu();
}
