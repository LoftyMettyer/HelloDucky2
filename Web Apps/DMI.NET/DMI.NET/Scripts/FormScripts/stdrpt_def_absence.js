
function stdrpt_def_absence_window_onload() {

	var frmPostDefinition = $('#frmPostDefinition')[0];

	if (frmPostDefinition.txtRecSelCurrentID.value > 0) {
		$("#optionframe").attr("data-framesource", "STDRPT_DEF_ABSENCE");

		$("#workframe").hide();

		$("#toolbarUtilities").show();
		$("#toolbarUtilities").click();

	} else {
		loadEmptyOption();
		$("#workframe").attr("data-framesource", "STDRPT_DEF_ABSENCE");		
	}

		showDefaultRibbon();

		menu_refreshMenu();

		SetReportDefaults();
		display_Absence_Page(1);
		absenceBreakdownRefreshTab3Controls();
		// Disable the menu
		//menu_disableMenu();
		if (frmPostDefinition.txtRecSelCurrentID.value > 0) {
			$("#workframe").hide();
			$("#optionframe").show();
		}
		else {
			$("#optionframe").hide();
			$("#workframe").show();
		}
	}

function formatAbsenceClick(index) {

	var frmAbsenceDefinition = $('#frmAbsenceDefinition')[0];
	var fViewing = (frmAbsenceDefinition.txtAction.value.toUpperCase() == "VIEW");

	checkbox_disable(frmAbsenceDefinition.chkPreview, ((index == 0) || (fViewing == true)));
	frmAbsenceDefinition.chkPreview.checked = (index != 0);

	frmAbsenceDefinition.chkDestination2.checked = false;
	frmAbsenceDefinition.chkDestination3.checked = false;

	if (index == 1) {
		frmAbsenceDefinition.chkDestination2.checked = true;
		frmAbsenceDefinition.txtFilename.value = '';
	}
	absenceBreakdownRefreshTab3Controls();
}

function selectAbsenceEmailGroup() {
	var currentID = $("#txtAbsenceEmailGroupID").val();
		OpenHR.modalExpressionSelect("EMAIL", 0, currentID, function (id, name) {
		$("#txtAbsenceEmailGroupID").val(id);
		$("#txtAbsenceEmailGroup").val(name);
	}, 400, 400);
}

function validateNumeric(pobjNumericControl)
{
		var sValue = pobjNumericControl.value;

		if (sValue.length == 0) 
		{            
				OpenHR.messageBox("Invalid numeric value entered.");
				pobjNumericControl.focus();
				return false;
		}
		else 
		{
				if (isNaN(sValue) == true)
				{
						OpenHR.messageBox("Invalid numeric value entered.");
						pobjNumericControl.focus();
						return false;
				}
				else 
				{
						return true;
				}
		}	
}

function validateAbsenceTab3() {
	return (true);
}

function absence_returnToRecEdit() {

	var frmPostDefinition = document.getElementById('frmPostDefinition');
	
	if (frmPostDefinition.txtRecSelCurrentID.value == 0) {
		//window.location.href = "default";
		$("#optionframe").hide();
		$("#workframe").show();
		$("#toolbarRecord").click();

	} else {
		$("#optionframe").hide();
		$("#workframe").show();
		$("#toolbarRecord").click();
		
		refreshData(); //workframe

		loadEmptyOption();

	}
}

function loadEmptyOption() {
	$.ajax({
		url: 'emptyoption',
		type: "POST",
		dataType: 'html',
		async: true,
		success: function (html) {
			try {
				$('#optionframe').html('');
				$('#optionframe').html(html);
			} catch (e) { }
		}
	});
}

function absence_okClick() {

	var fOK = true;
	var frmAbsenceDefinition = $('#frmAbsenceDefinition')[0];
	var frmPostDefinition = $('#frmPostDefinition')[0];
	var dataCollection = frmAbsenceDefinition.elements;

	var sValue = frmAbsenceDefinition.txtDateFrom.value;
	if (sValue.length == 0) {
		fOK = false;
	}
	else {
		sValue = absencedef_convertLocaleDateToSQL(sValue);
		if (sValue.length == 0) {
			fOK = false;
		}
	}

	if (fOK == false) {
		OpenHR.messageBox("Invalid start date value entered.");
		display_Absence_Page(1);
		frmAbsenceDefinition.txtDateFrom.focus();
		return;
	}

	sValue = frmAbsenceDefinition.txtDateTo.value;
	if (sValue.length == 0) {
		fOK = false;
	}
	else {
		sValue = absencedef_convertLocaleDateToSQL(sValue);
		if (sValue.length == 0) {
			fOK = false;
		}
	}

	if (fOK == false) {
		OpenHR.messageBox("Invalid end date value entered.");
		display_Absence_Page(1);
		frmAbsenceDefinition.txtDateTo.focus();
		return;
	}

		//Check if report end date is before the report start date
		var dStartDate = new Date(frmAbsenceDefinition.txtDateFrom.value);
		var dEndDate = new Date(frmAbsenceDefinition.txtDateTo.value);

		if (dEndDate < dStartDate) {
				OpenHR.messageBox("The report end date is before the report start date.");
				display_Absence_Page(1);
				frmAbsenceDefinition.txtDateFrom.focus();
				return;
		}

	frmPostDefinition.txtFromDate.value = absencedef_convertLocaleDateToSQL(frmAbsenceDefinition.txtDateFrom.value);
	frmPostDefinition.txtToDate.value = absencedef_convertLocaleDateToSQL(frmAbsenceDefinition.txtDateTo.value);
		frmPostDefinition.txtAbsenceTypes.value = "";

		if (dataCollection!=null) 
		{
				for (iIndex=0; iIndex<dataCollection.length; iIndex++)  
				{
					try {
						sControlName = dataCollection.item(iIndex).name;

						if (sControlName.substr(0, 15) == "chkAbsenceType_") {
							if (dataCollection.item(iIndex).checked == true) {
								//Who hardcoded the "7"???? - frmPostDefinition.txtAbsenceTypes.value = frmPostDefinition.txtAbsenceTypes.value + dataCollection.item(iIndex).attributes[7].nodeValue + ",";
								frmPostDefinition.txtAbsenceTypes.value = frmPostDefinition.txtAbsenceTypes.value + $(dataCollection.item(iIndex)).attr('tagname') + ",";
							}
						}
					}
					catch(e) {}
				}
		}


		if (frmPostDefinition.txtAbsenceTypes.value == "") {
			OpenHR.messageBox("You must have at least 1 absence type selected.");
			display_Absence_Page(1);
			fOK = false;
		}


	frmPostDefinition.utilid.value = "0";
	frmPostDefinition.txtPicklistName.value = "";
	frmPostDefinition.txtFilterName.value = "";
	if (frmAbsenceDefinition.optPickList.checked == true) {
		frmPostDefinition.utilid.value = frmPostDefinition.txtBasePicklistID.value;
		if ($("#RecordSelection").css("visibility") != "hidden") {
			frmPostDefinition.txtPicklistName.value = frmAbsenceDefinition.txtBasePicklist.value;
		}
	}
	if (frmAbsenceDefinition.optFilter.checked == true) {
		frmPostDefinition.utilid.value = frmPostDefinition.txtBaseFilterID.value;
		if ($("#RecordSelection").css("visibility") != "hidden") {
			frmPostDefinition.txtFilterName.value = frmAbsenceDefinition.txtBaseFilter.value;
		}
	}
	if ((frmAbsenceDefinition.optPickList.checked == true) && (frmPostDefinition.txtBasePicklistID.value == "0")) 
		{
				OpenHR.messageBox("You must have a picklist selected.");
				display_Absence_Page(1);
				fOK = false;
		}
		
		if ((frmAbsenceDefinition.optFilter.checked == true) && (frmPostDefinition.txtBaseFilterID.value == "0"))
		{
				OpenHR.messageBox("You must have a filter selected.");
				display_Absence_Page(1);		
				fOK = false;
		}

		frmPostDefinition.txtPrintFPinReportHeader.value = frmAbsenceDefinition.chkPrintInReportHeader.checked;

		// Bradford Specific data
		frmPostDefinition.txtSRV.value = frmAbsenceDefinition.chkSRV.checked;
		frmPostDefinition.txtShowDurations.value = frmAbsenceDefinition.chkShowDurations.checked;
		frmPostDefinition.txtShowInstances.value = frmAbsenceDefinition.chkShowInstances.checked;
		frmPostDefinition.txtShowFormula.value = frmAbsenceDefinition.chkShowFormula.checked;
		frmPostDefinition.txtOmitBeforeStart.value = frmAbsenceDefinition.chkOmitBeforeStart.checked;
		frmPostDefinition.txtOmitAfterEnd.value = frmAbsenceDefinition.chkOmitAfterEnd.checked;
		frmPostDefinition.txtOrderBy1.value = frmAbsenceDefinition.cboOrderBy1.options[frmAbsenceDefinition.cboOrderBy1.selectedIndex].text;
		frmPostDefinition.txtOrderBy1ID.value = frmAbsenceDefinition.cboOrderBy1.options[frmAbsenceDefinition.cboOrderBy1.selectedIndex].value;
		frmPostDefinition.txtOrderBy1Asc.value = frmAbsenceDefinition.chkOrderBy1Asc.checked;
		frmPostDefinition.txtOrderBy2.value = frmAbsenceDefinition.cboOrderBy2.options[frmAbsenceDefinition.cboOrderBy2.selectedIndex].text;
		frmPostDefinition.txtOrderBy2ID.value = frmAbsenceDefinition.cboOrderBy2.options[frmAbsenceDefinition.cboOrderBy2.selectedIndex].value;
		frmPostDefinition.txtOrderBy2Asc.value = frmAbsenceDefinition.chkOrderBy2Asc.checked;
		frmPostDefinition.txtMinimumBradfordFactor.value = frmAbsenceDefinition.chkMinimumBradfordFactor.checked;
		frmPostDefinition.txtMinimumBradfordFactorAmount.value = frmAbsenceDefinition.txtMinimumBradfordFactor.value;
		frmPostDefinition.txtDisplayBradfordDetail.value = frmAbsenceDefinition.chkShowAbsenceDetails.checked;
    
	// Validate the output options
		if (fOK == true) {
			if (validateAbsenceTab3() == false) {
				return;
			}
		}


		var sAttachmentName = new String(frmAbsenceDefinition.txtEmailAttachAs.value);
		if ((sAttachmentName.indexOf("/") != -1) ||
        (sAttachmentName.indexOf("?") != -1) ||
        (sAttachmentName.indexOf(String.fromCharCode(34)) != -1) ||
        (sAttachmentName.indexOf("<") != -1) ||
        (sAttachmentName.indexOf(">") != -1) ||
        (sAttachmentName.indexOf("|") != -1) ||
        (sAttachmentName.indexOf("@") != -1) ||
        (sAttachmentName.indexOf("~") != -1) ||
        (sAttachmentName.indexOf("}") != -1) ||
        (sAttachmentName.indexOf("{") != -1) ||
        (sAttachmentName.indexOf("[") != -1) ||
        (sAttachmentName.indexOf("]") != -1) ||
        (sAttachmentName.indexOf("#") != -1) ||
        (sAttachmentName.indexOf(";") != -1) ||
        (sAttachmentName.indexOf("+") != -1) ||
	    (sAttachmentName.indexOf("'") != -1) ||
	    (sAttachmentName.indexOf("*") != -1)) {
		    OpenHR.messageBox("The email attachment file name can not contain any of the following characters:\n/ ? " + String.fromCharCode(34) + " < > | * @ ~ [] {} # ' + ¬", 48, "Output Options");
		    return;
		}

		sAttachmentName = new String(frmAbsenceDefinition.txtFilename.value);
		if ((sAttachmentName.indexOf("/") != -1) ||
        (sAttachmentName.indexOf("?") != -1) ||
        (sAttachmentName.indexOf(String.fromCharCode(34)) != -1) ||
        (sAttachmentName.indexOf("<") != -1) ||
        (sAttachmentName.indexOf(">") != -1) ||
        (sAttachmentName.indexOf("|") != -1) ||
        (sAttachmentName.indexOf("@") != -1) ||
        (sAttachmentName.indexOf("~") != -1) ||
        (sAttachmentName.indexOf("}") != -1) ||
        (sAttachmentName.indexOf("{") != -1) ||
        (sAttachmentName.indexOf("[") != -1) ||
        (sAttachmentName.indexOf("]") != -1) ||
        (sAttachmentName.indexOf("#") != -1) ||
        (sAttachmentName.indexOf(";") != -1) ||
        (sAttachmentName.indexOf("+") != -1) ||
	    (sAttachmentName.indexOf("'") != -1) ||
	    (sAttachmentName.indexOf("*") != -1)) {
		    OpenHR.messageBox("The Save To file name can not contain any of the following characters:\n/ ? " + String.fromCharCode(34) + " < > | * @ ~ [] {} # ' + ¬", 48, "Output Options");
		    return;
		}

		if (frmAbsenceDefinition.chkPreview.checked == true) {
			frmPostDefinition.txtSend_OutputPreview.value = 1;
		}
		else {
			frmPostDefinition.txtSend_OutputPreview.value = 0;
		}

		frmPostDefinition.txtSend_OutputFormat.value = 0;
		if (frmAbsenceDefinition.optDefOutputFormat1.checked) frmPostDefinition.txtSend_OutputFormat.value = 1;
		if (frmAbsenceDefinition.optDefOutputFormat2.checked) frmPostDefinition.txtSend_OutputFormat.value = 2;
		if (frmAbsenceDefinition.optDefOutputFormat3.checked) frmPostDefinition.txtSend_OutputFormat.value = 3;
		if (frmAbsenceDefinition.optDefOutputFormat4.checked) frmPostDefinition.txtSend_OutputFormat.value = 4;
		if (frmAbsenceDefinition.optDefOutputFormat5.checked) frmPostDefinition.txtSend_OutputFormat.value = 5;
		if (frmAbsenceDefinition.optDefOutputFormat6.checked) frmPostDefinition.txtSend_OutputFormat.value = 6;

		if (frmAbsenceDefinition.chkDestination2.checked == true) {
			frmPostDefinition.txtSend_OutputSave.value = 1;
			//frmPostDefinition.txtSend_OutputSaveExisting.value = frmAbsenceDefinition.cboSaveExisting.options[frmAbsenceDefinition.cboSaveExisting.selectedIndex].value;
		}
		else {
			frmPostDefinition.txtSend_OutputSave.value = 0;
			frmPostDefinition.txtSend_OutputSaveExisting.value = 0;
		}

		if (frmAbsenceDefinition.chkDestination3.checked == true) {
			frmPostDefinition.txtSend_OutputEmail.value = 1;
			frmPostDefinition.txtSend_OutputEmailAddr.value = frmAbsenceDefinition.txtAbsenceEmailGroupID.value;
			frmPostDefinition.txtSend_OutputEmailSubject.value = frmAbsenceDefinition.txtEmailSubject.value;
			frmPostDefinition.txtSend_OutputEmailAttachAs.value = frmAbsenceDefinition.txtEmailAttachAs.value;
		}
		else {
			frmPostDefinition.txtSend_OutputEmail.value = 0;
			frmPostDefinition.txtSend_OutputEmailAddr.value = 0;
			frmPostDefinition.txtSend_OutputEmailSubject.value = '';
			frmPostDefinition.txtSend_OutputEmailAttachAs.value = '';
		}

		frmPostDefinition.txtSend_OutputFilename.value = frmAbsenceDefinition.txtFilename.value;

		if (fOK == true) {
			OpenHR.submitForm(frmPostDefinition, "reportframe", null, null, "util_run_standardreport_promptedvalues");
		}

		return;
}


function selectAbsencePicklist() {

	var tableID = $("#recSelTableID").val();
	var currentID = $("#txtBasePicklistID").val();

	OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name) {
		$("#txtBasePicklistID").val(id);
		$("#txtBasePicklist").val(name);
	}, getPopupWidth(), getPopupHeight());

}

function selectAbsenceFilter() {

	var tableID = $("#recSelTableID").val();
	var currentID = $("#txtBaseFilterID").val();

	OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name) {
		$("#txtBaseFilterID").val(id);
		$("#txtBaseFilter").val(name);
	}, getPopupWidth(), getPopupHeight());

}

function changeRecordOptions(psType) {

	var frmPostDefinition = $('#frmPostDefinition')[0];
	var frmAbsenceDefinition = $('#frmAbsenceDefinition')[0];
	
	if (psType == 'PICKLIST') {
		button_disable(frmAbsenceDefinition.cmdBasePicklist, false);
		button_disable(frmAbsenceDefinition.cmdBaseFilter, true);

		frmAbsenceDefinition.optAllRecords.checked = false;
		frmAbsenceDefinition.optFilter.checked = false;
		frmAbsenceDefinition.txtBaseFilter.value = "";
		frmPostDefinition.txtBaseFilter.value = "";
		frmPostDefinition.txtBaseFilterID.value = 0;
	}

	if (psType == 'FILTER') {
		button_disable(frmAbsenceDefinition.cmdBasePicklist, true);
		button_disable(frmAbsenceDefinition.cmdBaseFilter, false);

		frmAbsenceDefinition.optAllRecords.checked = false;
		frmAbsenceDefinition.optPickList.checked = false;
		frmAbsenceDefinition.txtBasePicklist.value = "";
		frmPostDefinition.txtBasePicklist.value = "";
		frmPostDefinition.txtBasePicklistID.value = 0;
	}

	if (psType == 'ALL') {
		button_disable(frmAbsenceDefinition.cmdBasePicklist, true);
		button_disable(frmAbsenceDefinition.cmdBaseFilter, true);

		frmAbsenceDefinition.optPickList.checked = false;
		frmAbsenceDefinition.optFilter.checked = false;

		frmAbsenceDefinition.txtBasePicklist.value = "";
		frmPostDefinition.txtBasePicklist.value = "";
		frmPostDefinition.txtBasePicklistID.value = 0;

		frmAbsenceDefinition.txtBaseFilter.value = "";
		frmPostDefinition.txtBaseFilter.value = "";
		frmPostDefinition.txtBaseFilterID.value = 0;
	}

	//refreshTab1Controls();
	absenceBreakdownRefreshTab1Controls();
}

function display_Absence_Page(piPageNumber) {

	var div1 = document.getElementById("div1");
	var div2 = document.getElementById("div2");
	var div3 = document.getElementById("div3");
	var frmAbsenceDefinition = $('#frmAbsenceDefinition')[0];

	if (piPageNumber == 1) {
		div1.style.visibility = "visible";
		div1.style.display = "block";
		div2.style.visibility = "hidden";
		div2.style.display = "none";
		div3.style.visibility = "hidden";
		div3.style.display = "none";
		button_disable(frmAbsenceDefinition.btnTab1, true);
		button_disable(frmAbsenceDefinition.btnTab2, false);

		if (frmAbsenceDefinition.txtUtilID.value == 16) {
			button_disable(frmAbsenceDefinition.btnTab2, false);
		}
		button_disable(frmAbsenceDefinition.btnTab3, false);
		absenceBreakdownRefreshTab1Controls();
	}

	if (piPageNumber == 2) {
		div1.style.visibility = "hidden";
		div1.style.display = "none";
		div2.style.visibility = "visible";
		div2.style.display = "block";
		div3.style.visibility = "hidden";
		div3.style.display = "none";
		button_disable(frmAbsenceDefinition.btnTab1, false);
		button_disable(frmAbsenceDefinition.btnTab2, true);
		button_disable(frmAbsenceDefinition.btnTab3, false);
		absenceBreakdownRefreshTab2Controls();
	}

	if (piPageNumber == 3) {
		div1.style.visibility = "hidden";
		div1.style.display = "none";
		div2.style.visibility = "hidden";
		div2.style.display = "none";
		div3.style.visibility = "visible";
		div3.style.display = "block";
		button_disable(frmAbsenceDefinition.btnTab1, false);
		button_disable(frmAbsenceDefinition.btnTab2, false);
		if (frmAbsenceDefinition.txtUtilID.value == 16) {
			button_disable(frmAbsenceDefinition.btnTab2, false);
		}
		button_disable(frmAbsenceDefinition.btnTab3, true);
		absenceBreakdownRefreshTab3Controls();
	}
}

function absenceBreakdownRefreshTab1Controls() {
	var frmAbsenceDefinition = $('#frmAbsenceDefinition')[0];

	if (frmAbsenceDefinition.optAllRecords.checked == true) {
		checkbox_disable(frmAbsenceDefinition.chkPrintInReportHeader, true);
		frmAbsenceDefinition.chkPrintInReportHeader.checked = false;
	}
	else {
		checkbox_disable(frmAbsenceDefinition.chkPrintInReportHeader, false);
	}
}

function absenceBreakdownRefreshTab2Controls() {

	var frmAbsenceDefinition = $('#frmAbsenceDefinition')[0];
	var frmPostDefinition = $('#frmPostDefinition')[0];

	if (frmPostDefinition.txtRecSelCurrentID.value > 0) {
		combo_disable(frmAbsenceDefinition.cboOrderBy1, true);
		combo_disable(frmAbsenceDefinition.cboOrderBy2, true);

		checkbox_disable(frmAbsenceDefinition.chkOrderBy1Asc, true);
		checkbox_disable(frmAbsenceDefinition.chkOrderBy2Asc, true);

		checkbox_disable(frmAbsenceDefinition.chkMinimumBradfordFactor, true);
		frmAbsenceDefinition.chkMinimumBradfordFactor.checked = false;
		text_disable(frmAbsenceDefinition.txtMinimumBradfordFactor, true);
		frmAbsenceDefinition.txtMinimumBradfordFactor.value = 0;
	}
	else {
		combo_disable(frmAbsenceDefinition.cboOrderBy1, false);
		combo_disable(frmAbsenceDefinition.cboOrderBy2, false);

		if (!frmAbsenceDefinition.chkMinimumBradfordFactor.checked) {
			frmAbsenceDefinition.txtMinimumBradfordFactor.value = 0;
			text_disable(frmAbsenceDefinition.txtMinimumBradfordFactor, true);
		}
		else {
			text_disable(frmAbsenceDefinition.txtMinimumBradfordFactor, false);
		}

		if (frmAbsenceDefinition.cboOrderBy1.options[frmAbsenceDefinition.cboOrderBy1.selectedIndex].value > 0) {
			checkbox_disable(frmAbsenceDefinition.chkOrderBy1Asc, false);
		}
		else {
			frmAbsenceDefinition.chkOrderBy1Asc.checked = false;
			checkbox_disable(frmAbsenceDefinition.chkOrderBy1Asc, true);
		}

		if (frmAbsenceDefinition.cboOrderBy2.options[frmAbsenceDefinition.cboOrderBy2.selectedIndex].value > 0) {
			checkbox_disable(frmAbsenceDefinition.chkOrderBy2Asc, false);
		}
		else {
			frmAbsenceDefinition.chkOrderBy2Asc.checked = false;
			checkbox_disable(frmAbsenceDefinition.chkOrderBy2Asc, true);
		}

	}

	if (!frmAbsenceDefinition.chkShowAbsenceDetails.checked) {
		frmAbsenceDefinition.chkSRV.checked = false;
		checkbox_disable(frmAbsenceDefinition.chkSRV, true);
	}
	else {
		checkbox_disable(frmAbsenceDefinition.chkSRV, false);
	}
}

function absenceBreakdownRefreshTab3Controls() {

	var frmAbsenceDefinition = $('#frmAbsenceDefinition')[0];   
	var fViewing = (frmAbsenceDefinition.txtAction.value.toUpperCase() == "VIEW");

		with (frmAbsenceDefinition)
		{
			text_disable(txtAbsenceEmailGroup, true);

				if (optDefOutputFormat0.checked == true)		//Data Only
				{
						//disable preview opitons
						chkPreview.checked = false;
						checkbox_disable(chkPreview, true);
								
						//disable save options
					chkDestination2.checked = false;
						checkbox_disable(chkDestination2, true);
						combo_disable(cboSaveExisting, true);
						txtFilename.value = '';
						text_disable(txtFilename, true);
						button_disable(cmdFilename, true);
			
						//disable email options
						chkDestination3.checked = false;
						checkbox_disable(chkDestination3, true);
						txtAbsenceEmailGroup.value = '';
						txtAbsenceEmailGroupID.value = 0;
						button_disable(cmdEmailGroup, true);
						text_disable(txtEmailSubject, true);
						text_disable(txtEmailAttachAs, true);

				}
				else if (optDefOutputFormat1.checked == true)   //CSV File
				{
						//enable preview opitons
						checkbox_disable(chkPreview, (fViewing == true));
									
						//enable-disable save options
						checkbox_disable(chkDestination2, false);
						if (chkDestination2.checked == true)
						{
								combo_disable(cboSaveExisting, false);
								text_disable(txtFilename, false);
								button_disable(cmdFilename, false);
						}	
						else
						{
								combo_disable(cboSaveExisting, true);
								text_disable(txtFilename, true);
								txtFilename.value = '';
								button_disable(cmdFilename, true);
						}
			
						//enable-disable email options
						checkbox_disable(chkDestination3, false);
						if (chkDestination3.checked == true)
						{
								text_disable(txtEmailSubject, false);
								button_disable(cmdEmailGroup, false);
								text_disable(txtEmailAttachAs, false);
						}
						else
						{
								txtAbsenceEmailGroup.value = '';
								txtAbsenceEmailGroupID.value = 0;
								button_disable(cmdEmailGroup, true);
								text_disable(txtEmailSubject, true);
								text_disable(txtEmailAttachAs, true);
						}
				}
				else if (optDefOutputFormat2.checked == true)		//HTML Document
				{
						//enable preview opitons
						checkbox_disable(chkPreview, (fViewing == true));
								
						//enable-disable save options
						checkbox_disable(chkDestination2, false);
						if (chkDestination2.checked == true)
						{
								combo_disable(cboSaveExisting, false);
								text_disable(txtFilename, false);
								button_disable(cmdFilename, false);
						}	
						else
						{
								combo_disable(cboSaveExisting, true);
								text_disable(txtFilename, true);
								txtFilename.value = '';
								button_disable(cmdFilename, true);
						}

						//enable-disable email options
						checkbox_disable(chkDestination3, false);
						if (chkDestination3.checked == true)
						{
								text_disable(txtEmailSubject, false);
								button_disable(cmdEmailGroup, false);
								text_disable(txtEmailAttachAs, false);
						}
						else
						{
								txtAbsenceEmailGroup.value = '';
								txtAbsenceEmailGroupID.value = 0;
								button_disable(cmdEmailGroup, true);
								text_disable(txtEmailSubject, true);
								text_disable(txtEmailAttachAs, true);
						}
				}
				else if (optDefOutputFormat3.checked == true)		//Word Document
				{
						//enable preview opitons
						checkbox_disable(chkPreview, (fViewing == true));
					
						//enable-disable save options
						checkbox_disable(chkDestination2, false);
						if (chkDestination2.checked == true)
						{
								combo_disable(cboSaveExisting,  false);
								text_disable(txtFilename, false);
								button_disable(cmdFilename, false);
						}	
						else
						{
								combo_disable(cboSaveExisting,  true);
								text_disable(txtFilename, true);
								txtFilename.value = '';
								button_disable(cmdFilename, true);
						}

						//enable-disable email options
						checkbox_disable(chkDestination3, false);
						if (chkDestination3.checked == true)
						{
								text_disable(txtEmailSubject, false);
								button_disable(cmdEmailGroup, false);
								text_disable(txtEmailAttachAs, false);
						}
						else
						{
								txtAbsenceEmailGroup.value = '';
								txtAbsenceEmailGroupID.value = 0;
								button_disable(cmdEmailGroup, true);
								text_disable(txtEmailSubject, true);
								text_disable(txtEmailAttachAs, true);
						}
				}
				else if ((optDefOutputFormat4.checked == true) ||		//Excel Worksheet
								 (optDefOutputFormat5.checked == true) ||
								 (optDefOutputFormat6.checked == true))
				{
						//enable preview opitons
						checkbox_disable(chkPreview, (fViewing == true));
						
						//enable-disable save options
						checkbox_disable(chkDestination2, false);
						if (chkDestination2.checked == true)
						{
								combo_disable(cboSaveExisting, false);
								text_disable(txtFilename, false);
								button_disable(cmdFilename, false);
						}	
						else
						{
								combo_disable(cboSaveExisting, true);
								text_disable(txtFilename, true);
								txtFilename.value = '';
								button_disable(cmdFilename, true);
						}

						//enable-disable email options
						checkbox_disable(chkDestination3, false);
						if (chkDestination3.checked == true)
						{
								text_disable(txtEmailSubject, false);
								button_disable(cmdEmailGroup, false);
								text_disable(txtEmailAttachAs, false);
						}
						else
						{
								txtAbsenceEmailGroup.value = '';
								txtAbsenceEmailGroupID.value = 0;
								button_disable(cmdEmailGroup, true);
								text_disable(txtEmailSubject, true);
								text_disable(txtEmailAttachAs, true);
						}
				}
						/*else if (optDefOutputFormat5.checked == true)		//Excel Chart
								{
								}
						else if (optDefOutputFormat6.checked == true)		//Excel Pivot Table
								{
								}*/
				else
				{
						optDefOutputFormat0.checked = true;
						absenceBreakdownRefreshTab3Controls();
				}
		
				if (txtEmailSubject.disabled)
				{
						txtEmailSubject.value = '';
				}

				if (txtEmailAttachAs.disabled)
				{
						txtEmailAttachAs.value = '';
				}
				else
				{
						if (txtEmailAttachAs.value == '') {
								if (txtFilename.value != '') {
										sAttachmentName = new String(txtFilename.value);
										txtEmailAttachAs.value = sAttachmentName.substr(sAttachmentName.lastIndexOf("\\")+1);
								}
						}
				}

				if (cmdFilename.disabled == true) {
						txtFilename.value = "";
				}

		}

}

function populateAbsenceFileName() {


	var sFileName;
	var dialog = document.getElementById("cmdGetFilename");

	if (frmAbsenceDefinition.optDefOutputFormat1.checked == true) {
		//CSV
		dialog.accept = "test/csv";
	}
	else if (frmAbsenceDefinition.optDefOutputFormat2.checked == true) {
		//HTML
		dialog.accept = "text/html";
	}

	else if (frmAbsenceDefinition.optDefOutputFormat3.checked == true) {
		//WORD
		dialog.accept = "application/msword, application/vnd.openxmlformats-officedocument.wordprocessingml.document";
	}

	else {
		//EXCEL
		dialog.accept = "application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	}

	if (frmAbsenceDefinition.txtFilename.value.length != 0) {
		dialog.value = frmAbsenceDefinition.txtFilename.value;
	}


	try {
		dialog.click();
		sFileName = dialog.value;

		if (sFileName.length > 256) {
			OpenHR.messageBox("Path and file name must not exceed 256 characters in length");
			return;
		}

		if (sFileName.length > 0) {
			frmAbsenceDefinition.txtFilename.value = sFileName;
		}

	}
	catch (e) {
	}
	
}

function absencedef_convertLocaleDateToSQL(psDateString)
{ 
		/* Convert the given date string (in locale format) into 
		SQL format (mm/dd/yyyy). */
		var sDateFormat;
		var iDays;
		var iMonths;
		var iYears;
		var sDays;
		var sMonths;
		var sYears;
		var iValuePos;
		var sTempValue;
		var sValue;
		var iLoop;

	sDateFormat = window.LocaleDateFormat;

		sDays="";
		sMonths="";
		sYears="";
		iValuePos = 0;

		// Trim leading spaces.
		sTempValue = psDateString.substr(iValuePos,1);
		while (sTempValue.charAt(0) == " ") 
		{
				iValuePos = iValuePos + 1;		
				sTempValue = psDateString.substr(iValuePos,1);
		}

		for (iLoop=0; iLoop<sDateFormat.length; iLoop++)  {
				if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'D') && (sDays.length==0)){
						sDays = psDateString.substr(iValuePos,1);
						iValuePos = iValuePos + 1;
						sTempValue = psDateString.substr(iValuePos,1);

						if (isNaN(sTempValue) == false) {
								sDays = sDays.concat(sTempValue);			
						}
						iValuePos = iValuePos + 1;		
				}

				if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'M') && (sMonths.length==0)){
						sMonths = psDateString.substr(iValuePos,1);
						iValuePos = iValuePos + 1;
						sTempValue = psDateString.substr(iValuePos,1);

						if (isNaN(sTempValue) == false) {
								sMonths = sMonths.concat(sTempValue);			
						}
						iValuePos = iValuePos + 1;
				}

				if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'Y') && (sYears.length==0)){
						sYears = psDateString.substr(iValuePos,1);
						iValuePos = iValuePos + 1;
						sTempValue = psDateString.substr(iValuePos,1);

						if (isNaN(sTempValue) == false) {
								sYears = sYears.concat(sTempValue);			
						}
						iValuePos = iValuePos + 1;
						sTempValue = psDateString.substr(iValuePos,1);

						if (isNaN(sTempValue) == false) {
								sYears = sYears.concat(sTempValue);			
						}
						iValuePos = iValuePos + 1;
						sTempValue = psDateString.substr(iValuePos,1);

						if (isNaN(sTempValue) == false) {
								sYears = sYears.concat(sTempValue);			
						}
						iValuePos = iValuePos + 1;
				}

				// Skip non-numerics
				sTempValue = psDateString.substr(iValuePos,1);
				while (isNaN(sTempValue) == true) {
						iValuePos = iValuePos + 1;		
						sTempValue = psDateString.substr(iValuePos,1);
				}
		}

		while (sDays.length < 2) {
				sTempValue = "0";
				sDays = sTempValue.concat(sDays);
		}

		while (sMonths.length < 2) {
				sTempValue = "0";
				sMonths = sTempValue.concat(sMonths);
		}

		while (sYears.length < 2) {
				sTempValue = "0";
				sYears = sTempValue.concat(sYears);
		}

		if (sYears.length == 2) {
				var iValue = parseInt(sYears);
				if (iValue < 30) {
						sTempValue = "20";
				}
				else {
						sTempValue = "19";
				}
		
				sYears = sTempValue.concat(sYears);
		}

		while (sYears.length < 4) {
				sTempValue = "0";
				sYears = sTempValue.concat(sYears);
		}

		sTempValue = sMonths.concat("/");
		sTempValue = sTempValue.concat(sDays);
		sTempValue = sTempValue.concat("/");
		sTempValue = sTempValue.concat(sYears);
	
		sValue = OpenHR.ConvertSQLDateToLocale(sTempValue);

		iYears = parseInt(sYears);
	
		while (sMonths.substr(0, 1) == "0") {
				sMonths = sMonths.substr(1);
		}
		iMonths = parseInt(sMonths);
	
		while (sDays.substr(0, 1) == "0") {
				sDays = sDays.substr(1);
		}
		iDays = parseInt(sDays);

		var newDateObj = new Date(iYears, iMonths - 1, iDays);
		if ((newDateObj.getDate() != iDays) || 
				(newDateObj.getMonth() + 1 != iMonths) || 
				(newDateObj.getFullYear() != iYears)) {
				return "";
		}
		else {
				return sTempValue;
		}
}

function setcancel() {
	// If we arrived here from RecEdit, switch back to RecEdit
	if (frmPostDefinition.txtRecSelCurrentID.value > 0) { 
		$("#workframe").attr("data-framesource", "RECORDEDIT");
	}
	refreshData();

	menu_disableMenu();

	$("#optionframe").hide();
	$("#workframe").show();
	$("#toolbarRecord").show();
	$("#toolbarRecord").click();

	menu_refreshMenu();

}
