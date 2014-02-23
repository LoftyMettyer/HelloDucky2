
function closeclick() {
	try {
		$(".month-year-input").remove();
		$(".popup").dialog("close");
	}
	catch (e) { }
}

function closepromptedclick() {
	try {
		$(".popup").dialog("close");

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
