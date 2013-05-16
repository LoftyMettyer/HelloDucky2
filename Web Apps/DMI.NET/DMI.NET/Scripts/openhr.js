﻿

(function(window, $) {
    "use strict";

    function handleAjaxError(html) {
        //handle error
        messageBox(html.ErrorMessage.replace("<p>", "\n\n"), 48, html.ErrorTitle);

        //redirect if specified...
        if (html.Redirect.length > 0) {
          //alert("redirecting to " + html.Redirect);
          window.location(html.Redirect);
        }
        
    }


	var version = "1.0.0",
	    mbStyle = { vbExclamation: 48, vbQuestion: 32, vbYesNo: 4, vbYesNoCancel: 3 },
	    mbResult = { vbYes: 6, vbNo: 7, vbCancel: 2 },

	    messageBox = function (prompt, buttons, title) {

	    	switch (buttons) {
	    		case mbStyle.vbExclamation:
	    			//48
	    			alert(prompt);
	    			break;
	    		case mbStyle.vbYesNoCancel:

	    			//TODO - Need to find a popup that can handle multiple buttons
	    			return confirm(prompt) ? mbResult.vbYes : mbResult.vbNo;

	    		case mbStyle.vbQuestion + mbStyle.vbYesNo:
	    			//36
	    			return confirm(prompt) ? mbResult.vbYes : mbResult.vbNo;

	    			break;
	    		case mbStyle.vbQuestion + mbStyle.vbYesNoCancel:
	    			//35
	    			return confirm(prompt) ? mbResult.vbYes : mbResult.vbNo;
	    			break;

	    		default:
	    			alert(prompt);
	    			//throw Error("OpenHR.messageBox buttons not coded for.");
	    			break;
	    	}

            

	    },

        showInReportFrame = function (form) {
            var $form = $(form),
   	    	    $frame = $("#reportframe"),
	    	    url = $form.attr("action"),
	    	    data = $form.serialize();

            $.ajax({
                url: url,
                type: "POST",
                data: data,
                async: false,
                success: function (html) {
                    try {
                        if (html.ErrorMessage.length > 0) {
                            //A handled error was returned. Display error message, then redirect accordingly...
                            handleAjaxError(html);
                            return false;
                        }
                    } catch (e) { }

                    //clear the frame...
                    $frame.html('');

                    $( ".popup" ).dialog( "open" );

                    //OK
                    $frame.html(html);
                    
                },
                error: function (req, status, errorObj) {
                    alert("OpenHR.showInReportFrame ajax call to '" + url + "' failed with '" + errorObj + "'.");
                }
            });
        },

	    showPopup = function (prompt) {

	    },
        getFrame = function(frameId) {
            return document.frames[frameId];
        },
	    getForm = function (frameId, formId) {
	    	return document.forms[formId];
	    },
	    submitForm = function(form, targetWin, asyncFlag) {		    
		    var $form = $(form),
			    $frame = $form.closest("div[data-framesource]").first(),
			    url = $form.attr("action"),
			    target = $form.attr("target"),
			    data = $form.serialize();

		    if ((asyncFlag == undefined) || (asyncFlag.length == 0) || (asyncFlag == true)) {
			    asyncFlag = true;
		    } else {
			    asyncFlag = false;
		    }
		    
	    	$.ajax({
	    		url: url,
	    		type: "POST",
	    		data: data,
	    		async: asyncFlag,
	    		success: function (html) {
	    		    
	    		    //console.log(html);

	    			try {	    				
	    			    if ((html.ErrorMessage != null) && (html.ErrorMessage != undefined) && (html.ErrorMessage != "undefined")) {
	    		            if (html.ErrorMessage.length > 0) {
	    		                //A handled error was returned. Display error message, then redirect accordingly...
	    		                handleAjaxError(html);
	    		                return false;
	    		            }
	    		        }
	    		    } catch (e) {		    		    
	    		    	//alert("OpenHR.submitForm ajax call to '" + url + "' failed with '" + e.toString() + "'.");
	    		    	$("#errorDialogTitle").text(e.toString);
	    		    	$("#errorDialogContentText").text(e.responseText);
	    		    	$("#errorDialog").dialog("open");
	    		    }	    			
	    		    //clear the frame...
	    		    
	    		    //OK
	    		    
	    		    if (targetWin != null) {
	    		        $frame = $form.closest("div[" + targetWin + "]").first();
	    		        //	    		        $(targetWin.document.body).html(html);	    		        
	    		        $frame.html(html);
	    		        
	    		    } else {	    		        
	    		        $frame.html('');
	    		        $frame.html(html);
	    		    }
	    			
	    			//jqwuery stylin	    			
	    		    $(function () {
		    		    $("input[type=submit], input[type=button], button")
			    		    .button();
		    		    $("input").addClass("ui-widget ui-widget-content ui-corner-all");
		    		    $("input").removeClass("text");
		    		    
		    		    $("textarea").addClass("ui-widget ui-widget-content ui-corner-all");
		    		    $("textarea").removeClass("text");
		    		    
		    		    $("select").addClass("ui-widget ui-widget-content ui-corner-all");
		    		    $("select").removeClass("text");
		    		    $("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
		    		    $("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");

	    		    });

	    		},
	    		error: function (req, status, errorObj) {	    			
	    			//alert("OpenHR.submitForm ajax call to '" + url + "' failed with '" + errorObj + "'.");
	    			$("#errorDialogTitle").text(errorObj);
	    			$("#errorDialogContentText").text(req.responseText);
	    			$("#errorDialog").dialog("open");

	    		}
	    	});
	    },
	    addActiveXHandler = function (controlId, eventName, func) {
	        var ctl = document.getElementById(controlId);
	        
            if (ctl != null) {
                ctl.attachEvent(eventName, func);
            }
	    },
	    refreshMenu = function () {
	    	//TODO
	    },
	    disableMenu = function () {
	    	//TODO
	    },
	    locateDateFormat = function () {
	        //TODO
	        return "dd/MM/yyyy";
	    },
	    localeDecimalSeparator = function () {
	        //TODO
	        return ".";
	    },
	    localeThousandSeparator = function () {
	        //TODO
	        return ",";
	    },
	    localeDateSeparator = function () {
	        //TODO
	        return "/";
	    },
	    convertSqlDateToLocale = function (z) {
	        var convertDate = Date.parseExact(z, "MM/dd/yyyy");
	        return convertDate.format(OpenHR.LocaleDateFormat());
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
	        return true;
	    },
			validateFilePath = function(sPath) {
				//TODO
				return true;
			},
			sendMail = function(sTo, sSubject, sBody, sCC, sBCC) {
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
				},
				mmwordCreateTemplateFile = function (psTemplatePath) {
					//TODO
					return true;
				};

	window.OpenHR = {
		version: version,
		messageBox: messageBox,
		showPopup: showPopup,
		getFrame: getFrame,
		getForm: getForm,
		submitForm: submitForm,
		showInReportFrame: showInReportFrame,
		addActiveXHandler: addActiveXHandler,
		refreshMenu: refreshMenu,
		disableMenu: disableMenu,
		LocaleDateFormat: locateDateFormat,
		LocaleDateSeparator: localeDateSeparator,
		LocaleDecimalSeparator: localeDecimalSeparator,
		LocaleThousandSeparator: localeThousandSeparator,
		ConvertSQLDateToLocale: convertSqlDateToLocale,
		PrinterCount: printerCount,
		PrinterName: printerName,
		GetRegistrySetting: getRegistrySetting,
		SaveRegistrySetting: saveRegistrySetting,
		ValidateDir: validateDir,
		ValidateFilePath: validateFilePath,
		sendMail: sendMail,
		currentWorkPage: currentWorkPage,
		MM_WORD_CreateTemplateFile: mmwordCreateTemplateFile
	};

})(window, jQuery);