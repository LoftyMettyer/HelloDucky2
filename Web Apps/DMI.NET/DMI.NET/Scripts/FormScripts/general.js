
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
