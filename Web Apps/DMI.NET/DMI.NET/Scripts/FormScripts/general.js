// Enums
var actionType = {
	New: 0, Edit: 1, View: 2, Copy: 3, Run: 4, Delete: 5,
	MoveFirst: 6, MovePrevious: 7, MoveNext: 8, MoveLast: 9, Locate: 10, LocateId: 11,
	Save: 12
};

var utilityType = {
	CrossTab: 1,
	CustomReport: 2,
	MailMerge: 9,
	Picklist: 10,
	Filter: 11,
	Calculation: 12,
	AbsenceBreakdown: 15,
	BradfordFactor: 16,
	CalendarReport: 17,
	Workflow: 25,
	NineBoxGrid: 35
};

var optionActionType = {
	Empty : 0,
	ABSENCEBREAKDOWNALL : 5,
	ABSENCEBREAKDOWNREC : 6,
	ADDEXPRCOMPONENT : 7,
	ADDFROMWAITINGLIST : 8,
	ADDFROMWAITINGLISTERROR : 9,
	ADDFROMWAITINGLISTSUCCESS : 10,
	ALL : 11,
	ALLRECORDS : 12,
	BOOKCOURSE : 13,
	BOOKCOURSEERROR : 14,
	BOOKCOURSESUCCESS : 15,
	BRADFORDFACTORALL : 16,
	BRADFORDFACTORREC : 17,
	BULKBOOKINGERROR : 18,
	BULKBOOKINGSUCCESS : 19,
	CALCULATIONS : 20,
	CALENDARREPORTS : 21,
	CALENDARREPORTSREC : 22,
	CANCEL : 23,
	CANCELBOOKING : 24,
	CANCELBOOKING_1 : 25,
	CANCELCOURSE : 26,
	CANCELCOURSE_1 : 27,
	CANCELCOURSE_2 : 28,
	CLEAR : 29,
	CLEARFILTER : 30,
	COPY : 31,
	CROSSTABS : 32,
	NINEBOXGRID : 33,
	CUSTOMREPORTS : 34,
	DEFAULT : 35,
	DELETE : 36,
	EDITEXPRCOMPONENT : 37,
	EVENTLOG : 38,
	EXIT : 39,
	FILTER : 40,
	FILTERS : 41,
	FIND : 42,
	GETBULKBOOKINGSELECTION : 43,
	GETEXPRESSIONRETURNTYPES : 44,
	GETPICKLISTSELECTION : 45,
	INSERTEXPRCOMPONENT : 46,
	LINK : 47,
	LINKOLE : 48,
	LOAD : 49,
	LOADCALENDARREPORTCOLUMNS : 50,
	LOADEVENTLOG : 51,
	LOADEVENTLOGUSERS : 52,
	LOADEXPRFIELDCOLUMNS : 53,
	LOADEXPRLOOKUPCOLUMNS : 54,
	LOADEXPRLOOKUPVALUES : 55,
	LOADFIND : 56,
	LOADLOOKUPFIND : 57,
	LOADTRANSFERCOURSE : 58,
	LOADBOOKCOURSE : 59,
	LOADADDFROMWAITINGLIST : 60,
	LOADTRANSFERBOOKING : 61,
	LOADREPORTCOLUMNS : 62,
	LOCATE : 63,
	LOCATEID : 64,
	LOGOFF : 65,
	LOOKUP : 66,
	MAILMERGE : 67,
	MOVEFIRST : 68,
	MOVENEXT : 69,
	MOVELAST : 70,
	MOVEPREVIOUS : 71,
	NEW : 72,
	PARENT : 73,
	PICKLIST : 74,
	QUICKFIND : 75,
	REFRESHFINDAFTERDELETE : 76,
	REFRESHFINDAFTERINSERT : 77,
	RELOAD : 78,
	SAVE : 79,
	SAVEERROR : 80,
	SELECTADDFROMWAITINGLIST_1 : 81,
	SELECTADDFROMWAITINGLIST_2 : 82,
	SELECTADDFROMWAITINGLIST_3 : 83,
	SELECTBOOKCOURSE_1 : 84,
	SELECTBOOKCOURSE_2 : 85,
	SELECTBOOKCOURSE_3 : 86,
	SELECTBULKBOOKINGS : 87,
	SELECTBULKBOOKINGS_2 : 88,
	SELECTCOMPONENT : 89,
	SELECTFILTER : 90,
	SELECTIMAGE : 91,
	SELECTLINK : 92,
	SELECTLOOKUP : 93,
	SELECTOLE : 94,
	SELECTORDER : 95,
	SELECTTRANSFERBOOKING_1 : 96,
	SELECTTRANSFERBOOKING_2 : 97,
	SELECTTRANSFERCOURSE : 98,
	STDRPT_ABSENCECALENDAR : 99,
	STDREPORT_DATEPROMPT : 100,
	TRANSFERBOOKING : 101,
	TRANSFERBOOKINGERROR : 102,
	TRANSFERBOOKINGSUCCESS : 103,
	TRANSFERCOURSE : 104,
	VIEW : 105,
	WORKFLOW : 106,
	WORKFLOWOUTOFOFFICE : 107,
	WORKFLOWPENDINGSTEPS : 108
}


function closeclick() {
	try {
		$('.popup').dialog("option", "buttons", {});
		$(".month-year-input").remove();
		$(".popup").dialog("close");
	}
	catch (e) { }
}

function closepromptedclick() {
	try {
		$(".popup").dialog("close");
		$('.popup').dialog("option", "buttons", {});
		if (menu_isSSIMode()) {
			window.loadPartialView("linksMain", "Home", "workframe", null);
		}
	}
	catch (e) { }
}

function disableAll() {
	var i;

	var dataCollection = frmDefinition.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			var eElem = frmDefinition.elements[i];

			if ("text" == eElem.type) {
				text_disable(eElem, true);
			} else if ("TEXTAREA" == eElem.tagName) {
				textarea_disable(eElem, true);
			} else if ("checkbox" == eElem.type) {
				checkbox_disable(eElem, true);
			} else if ("radio" == eElem.type) {
				radio_disable(eElem, true);
			} else if ("button" == eElem.type) {
				if (eElem.value != "Cancel") {
					button_disable(eElem, true);
				}
			} else if ("SELECT" == eElem.tagName) {
				combo_disable(eElem, true);
			} else {
				grid_disable(eElem, true);
			}
		}
	}
}

function populateFileName(frmBase) {

	var sFileName;
	var dialog = document.getElementById("cmdGetFilename");

	if (frmBase.optOutputFormat1.checked == true) {
		//CSV
		dialog.accept = "test/csv";
	}
	else if (frmBase.optOutputFormat2.checked == true) {
		//HTML
		dialog.accept = "text/html";
	}

	else if (frmBase.optOutputFormat3.checked == true) {
		//WORD
		dialog.accept = "application/msword, application/vnd.openxmlformats-officedocument.wordprocessingml.document";
	}

	else {
		//EXCEL
		dialog.accept = "application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	}

	if (frmBase.txtFilename.value.length != 0) {
		dialog.value = frmBase.txtFilename.value;
	}


	try {
		dialog.click();		
		sFileName = dialog.value;

		if (sFileName.length > 256) {
			OpenHR.messageBox("Path and file name must not exceed 256 characters in length");
			return;
		}

		if (sFileName.length > 0) {
			frmBase.txtFilename.value = sFileName;
		}
		
	}
	catch (e) {
	}

}
