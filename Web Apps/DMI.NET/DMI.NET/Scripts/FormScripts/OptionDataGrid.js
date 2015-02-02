$(function () {
	$("#optionframe").attr("data-framesource", $('#DataFrameSource').val());
	$("#workframe").hide();
	$("#optionframe").show();

	//resize grid	
	var workPageHeight = $('.optiondatagridpage').height();
	var pageTitleHeight = $('.optiondatagridpage>.pageTitleDiv').height();
	var dropdownHeight = $('.floatleft').height();
	var navbuttonheight = $('.optiondatagridpage>footer').height();
	var gridMarginBottom = $('#FindGridRow').css('marginBottom').replace("px","");
	
	var newGridHeight = workPageHeight - pageTitleHeight - dropdownHeight - navbuttonheight - gridMarginBottom;
	
	$("#ssOleDBGridRecords").jqGrid('setGridHeight', newGridHeight);
	$('#FindGridRow').height(newGridHeight);

	refreshData();

	tbrefreshControls();

});

function refreshData() {
	// Get the optionData.asp to get the link find records.
	var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
	optionDataForm.txtOptionAction.value = $('#OptionAction').val();
	optionDataForm.txtOptionTableID.value = $('#TableID').val();
	optionDataForm.txtOptionViewID.value = $('#selectView').val();
	optionDataForm.txtOptionOrderID.value = $('#selectOrder').val();
	optionDataForm.txtOptionCourseTitle.value = $('#CourseTitle').val();
	optionDataForm.txtOptionRecordID.value = $('#RecordID').val();
	optionDataForm.txtOptionFirstRecPos.value = "1";
	optionDataForm.txtOptionCurrentRecCount.value = "0";
	optionDataForm.txtOptionPageAction.value = "LOAD";

	refreshOptionData();	//should be in scope
}

$('#selectOrder, #selectView').change(function () { refreshData(); });

$('#cmdCancel').click(function () {
	$("#optionframe").hide();
	$("#workframe").show();

	var frmGotoOption = document.getElementById("frmGotoOption");

	frmGotoOption.txtGotoOptionAction.value = $('#GotoOptionActionCancel').val();
	frmGotoOption.txtGotoOptionLinkRecordID.value = 0;
	frmGotoOption.txtGotoOptionPage.value = "emptyoption";
	OpenHR.submitForm(frmGotoOption);
});

$('#cmdSelect').click(function () {
	var frmGotoOption = document.getElementById("frmGotoOption");
	
	frmGotoOption.txtGotoOptionAction.value = $('#GotoOptionActionSelect').val();
	frmGotoOption.txtGotoOptionRecordID.value = $('#RecordID').val();
	var selRowId = $("#ssOleDBGridRecords").jqGrid('getGridParam', 'selrow');
	var recordID = $("#ssOleDBGridRecords").jqGrid('getCell', selRowId, 'ID');
	frmGotoOption.txtGotoOptionLinkRecordID.value = recordID;
	frmGotoOption.txtGotoOptionPage.value = "emptyoption";

	var optionDataForm = OpenHR.getForm("optiondataframe", "frmOptionData");
	frmGotoOption.txtGotoOptionLookupValue.value = $('#selStatus').val();

	OpenHR.submitForm(frmGotoOption);
});

function tbrefreshControls() {	
	var selRowId = $("#ssOleDBGridRecords").jqGrid('getGridParam', 'selrow');
	
	button_disable($('#cmdSelect'), (!selRowId > 0));

	if ($('#selectOrder').children().length <= 1) {
		combo_disable($('#selectOrder'), true);
		button_disable($('#btnGoOrder'), true);
	}

	if ($('#selectView').children().length <= 1) {
		combo_disable($('#selectView'), true);
		button_disable($('#btnGoView'), true);
	}

}


