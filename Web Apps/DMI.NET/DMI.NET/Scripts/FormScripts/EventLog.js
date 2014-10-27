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
	
	$('#evlEmailSelect_OK').button('disable');


});

function emailEvent() {
	var sTo = getEmails(4);
	var sCC = getEmails(5);
	var sBCC = getEmails(6);
	var sSubject = getSubject();
	var sBody = getBody();

	$.ajax({
		type: "POST",
		url: "SendEmail",
		data: { 'to': sTo, 'cc': sCC, 'bcc': sBCC, 'subject': sSubject, 'body': sBody },
		dataType: "text",
		success: function (a, b, c) {
			alert(OpenHR.replaceAll(c.statusText, '<br/>', '\n'));
			self.close();
		},
		error: function (req, status, errorObj) {
			if (!(errorObj == "" || req.responseText == "")) {
				alert(OpenHR.replaceAll(errorObj, '<br/>', '\n'));
			}
		}
	});

	return true;
}

function closeEmail() {
	try {
		//$("#EventLogEmailSelect").dialog("close");
		$(this).dialog("close");
		
		//$("#EventLogEmailSelect").dialog().dialog("close");
	}
	catch (e) { }
}

//function okClick() {

//	var frmOpenerPurge = window.dialogArguments.OpenHR.getForm("workframe", "frmPurge");
//	var frmMainLog = window.dialogArguments.OpenHR.getForm("workframe", "frmLog");


//	if ((frmEventPurge.cboPeriod.selectedIndex == 3) && (frmEventPurge.txtPeriod.value > 200)) {
//		OpenHR.messageBox("You cannot select a purge period of greater than 200 years.", 48, "Event Log");
//	}
//	else {
//		if (frmEventPurge.cboPeriod.selectedIndex == 0) {
//			frmOpenerPurge.txtPurgePeriod.value = 'dd';
//		}
//		else if (frmEventPurge.cboPeriod.selectedIndex == 1) {
//			frmOpenerPurge.txtPurgePeriod.value = 'wk';
//		}
//		else if (frmEventPurge.cboPeriod.selectedIndex == 2) {
//			frmOpenerPurge.txtPurgePeriod.value = 'mm';
//		}
//		else if (frmEventPurge.cboPeriod.selectedIndex == 3) {
//			frmOpenerPurge.txtPurgePeriod.value = 'yy';
//		}

//		frmOpenerPurge.txtPurgeFrequency.value = frmEventPurge.txtPeriod.value;
//		if (frmEventPurge.optPurge.checked == true) {
//			frmOpenerPurge.txtDoesPurge.value = 1;
//		}
//		else {
//			frmOpenerPurge.txtDoesPurge.value = 0;
//		}

//		frmOpenerPurge.txtCurrentUsername.value = frmMainLog.cboUsername.options[frmMainLog.cboUsername.selectedIndex].value;
//		frmOpenerPurge.txtCurrentType.value = frmMainLog.cboType.options[frmMainLog.cboType.selectedIndex].value;
//		frmOpenerPurge.txtCurrentMode.value = frmMainLog.cboMode.options[frmMainLog.cboMode.selectedIndex].value;
//		frmOpenerPurge.txtCurrentStatus.value = frmMainLog.cboStatus.options[frmMainLog.cboStatus.selectedIndex].value;

//		window.dialogArguments.OpenHR.submitForm(frmOpenerPurge);
//		self.close();
//	}
//}

function emailDetailEvent() {
	var sBatchInfo = "";
	var sURL;

	if (frmEventDetails.txtEventBatch.value == 1) {
		frmEmail.txtBatchy.value = 1;
		frmEmail.txtSelectedEventIDs.value = frmEventDetails.cboOtherJobs.options[frmEventDetails.cboOtherJobs.selectedIndex].value;

		sBatchInfo = sBatchInfo + "<%:DetailsLabel1%> :	" + document.getElementById('tdBatchJobName').innerText + String.fromCharCode(13) + String.fromCharCode(13);

		sBatchInfo = sBatchInfo + "<%:DetailsLabel2%> :	" + String.fromCharCode(13) + String.fromCharCode(13);

		for (var iCount = 0; iCount < frmEventDetails.cboOtherJobs.options.length; iCount++) {
			sBatchInfo = sBatchInfo + String(frmEventDetails.cboOtherJobs.options[iCount].text) + String.fromCharCode(13) + String.fromCharCode(13);
		}
	}
	else {
		frmEmail.txtBatchy.value = 0;
		frmEmail.txtSelectedEventIDs.value = frmEventDetails.txtOriginalEventID.value;
	}

	frmEmail.txtBatchInfo.value = sBatchInfo;

	sURL = "emailSelection" +
		"?txtSelectedEventIDs=" + frmEmail.txtSelectedEventIDs.value +
		"&txtEmailOrderColumn=" +
		"&txtEmailOrderOrder=" +
		"&txtFromMain=" + frmEmail.txtFromMain.value +
		"&txtBatchInfo=" + escape(frmEmail.txtBatchInfo.value) +
		"&txtBatchy=" + frmEmail.txtBatchy.value;


	var sURLString = sURL;
	$('#EventLogEmailSelect').data('sURLData', sURLString);
	$('#EventLogEmailSelect').dialog("open");

	//OpenHR.windowOpen(sURL, (screen.width) / 3, (screen.height) / 2, 'no', 'no');
}