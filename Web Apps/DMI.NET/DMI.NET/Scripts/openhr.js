

(function (window, $) {
	"use strict";

	var version = "1.0.0",
	    mbStyle = { vbExclamation: 48, vbQuestion: 32, vbYesNo: 4 },
	    mbResult = { vbYes: 6, vbNo: 7 },

	    messageBox = function (prompt, buttons, title) {
	        

	    	switch (buttons) {
	    		case mbStyle.vbExclamation:
	    			//48
	    			alert(prompt);
	    		case mbStyle.vbQuestion + mbStyle.vbYesNo:
	    			//36
	    		    return confirm(prompt) ? mbResult.vbYes : mbResult.vbNo;
	    		
	    	    default:
	    	        alert(prompt);
	    			//throw Error("OpenHR.messageBox buttons not coded for.");
	    	}
	    },
	    showPopup = function (prompt) {

	    },
	    getForm = function (frameId, formId) {
	    	return document.forms[formId];
	    },
	    submitForm = function (form) {	       
	    	var $form = $(form),
	    	    $frame = $form.closest("div[data-framesource]").first(),
	    	    url = $form.attr("action"),
	    	    target = $form.attr("target"),
	    	    data = $form.serialize();
	        
	 //       $("#post_form").attr("target", target);

	    	$.ajax({
	    		url: url,
	    		type: "POST",
	    		data: data,
	    		async: false,
	    		success: function (html) {
	    		    $("#workframeset").show();
	    			$frame.html(html);
	    		},
	    		error: function (req, status, errorObj) {
	    			alert("OpenHR.submitForm ajax call to '" + url + "' failed with '" + errorObj + "'.");
	    		}
	    	});
	    },
	    addActiveXHandler = function (controlId, eventName, func) {
	    	var ctl = document.getElementById(controlId);
	    	ctl.attachEvent(eventName, func);
	    },
	    refreshMenu = function () {
	    	//TODO
	    },
	    disableMenu = function () {
	    	//TODO
	    },
	    locateDateFormat = "DD/MM/YYYY",
	    localeDecimalSeparator = ".",
	    localeThousandSeparator = ",",
	    convertSqlDateToLocale = function(z) {
	        // TODO
	    },
	    printerCount = function () {
	    	//TODO
	    },
	    printerName = function (iLoop) {
	    	//TODO
	    },
	    getRegistrySetting = function (x, y, z) {
	    	//TODO
	    },
	    saveRegistrySetting = function(w, x, y, z) {
	        //TODO
	    },
	    validateDir = function(x, y) {
	        //TODO	        
	    },
	    SendMail = function(sTo, sSubject, sBody, sCC, sBCC) {
	        //TODO
	    },
    currentWorkPage = function () {
	        var sCurrentPage;
	        if (!($("#workframe").css('display') == 'none')) {
	            //Work frame is in view.
	            sCurrentPage = $("#workframe").attr("data-framesource").replace(".asp", "");
	        } else {
	            //Option frame is in view.
	            sCurrentPage = $("#optionframe").attr("data-framesource").replace(".asp", "");
	        }

	        sCurrentPage = sCurrentPage.toUpperCase();
	        return sCurrentPage;
	    };

	window.OpenHR = {
		version: version,
		messageBox: messageBox,
		showPopup: showPopup,
		getForm: getForm,
		submitForm: submitForm,
		addActiveXHandler: addActiveXHandler,
		refreshMenu: refreshMenu,
		disableMenu: disableMenu,
	    LocaleDateFormat: locateDateFormat,
		LocaleDecimalSeparator: localeDecimalSeparator,
		LocaleThousandSeparator: localeThousandSeparator,
		ConvertSQLDateToLocale : convertSqlDateToLocale,
		PrinterCount: printerCount,
		PrinterName: printerName,
		GetRegistrySetting: getRegistrySetting,
	    SaveRegistrySetting: saveRegistrySetting,
	    ValidateDir: validateDir,
	    sendMail: SendMail,
        currentWorkPage: currentWorkPage
	};

})(window, jQuery);