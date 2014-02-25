

(function (window, $) {
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
		mbStyle = { vbExclamation: 48, vbQuestion: 32, vbYesNo: 4, vbYesNoCancel: 3, vbOKCancel: 1 },
		mbResult = { vbYes: 6, vbNo: 7, vbCancel: 2, vbOK: 1},
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


		displayModalDialog = function (prompt, dialogButtons, title) {

			// Default parameters
			if (!title || title.length == 0) title = 'OpenHR Web';
			
			$('#dialog-confirm').dialog('option', 'buttons', dialogButtons);
			$('#dialog-confirm').dialog('option', 'title', title);
			$('#dialog-confirm p').text(prompt);
			$('#dialog-confirm').dialog('open');
      // If Any ActiveX controls are in the workframeset, move the dailog to the very top of the screen to avoid it being hidden behind the ActiveX
			if ($('#workframeset object').length > 0) {
				$('#dialog-confirm').dialog('option', 'position', 'top');
			} else {
				$('#dialog-confirm').dialog('option', 'position', 'center');
			}
				
},

		modalMessage = function (message, title) {
			var dialogButtons = {
				"OK": function () {
					$(this).dialog("close");
				}
			};

			displayModalDialog(message, dialogButtons, title);

		},

		modalPrompt = function (prompt, buttons, title, followOnFunctionName) {
			var defer = $.Deferred();
			switch (buttons) {
			case 1:
				var dialogButtons = {
					"OK": function() {
						defer.resolve(1);
						$(this).dialog("close");
						if (followOnFunctionName) followOnFunctionName(1);
					},
					"Cancel": function() {
						defer.resolve(2);
						$(this).dialog("close");
						if (followOnFunctionName) followOnFunctionName(2);
					}
				};
				break;
			case 3:
				var dialogButtons = {
					"Yes": function() {
						defer.resolve(6);
						$(this).dialog("close");
						if (followOnFunctionName) followOnFunctionName(6);
					},
					"No": function() {
						defer.resolve(7);
						$(this).dialog("close");
						if (followOnFunctionName) followOnFunctionName(7);
					},
					"Cancel": function() {
						defer.resolve(2);
						$(this).dialog("close");
						if (followOnFunctionName) followOnFunctionName(2);
					}
				};
				break;
			default:
				var dialogButtons = {
					"OK": function() {
						defer.resolve(1);
						$(this).dialog("close");
						if (followOnFunctionName) followOnFunctionName(1);
					}
				};
			}
			displayModalDialog(prompt, dialogButtons, title);
			return defer.promise();
		},
		
		showInReportFrame = function (form, asyncFlag) {

			var $form = $(form),
				$frame = $("#reportframe"),
				url = $form.attr("action"),
				data = $form.serialize();

			if ((asyncFlag == undefined) || (asyncFlag.length == 0) || (asyncFlag == true)) {
				asyncFlag = true;
			} else {
				asyncFlag = false;
			}

			$.ajax({
				url: url,
				dataType: 'html',
				type: "POST",
				data: data,
				async: asyncFlag,
				
				success: function (html) {
					try {

						if ((html.ErrorMessage != null) && (html.ErrorMessage != undefined) && (html.ErrorMessage != "undefined")) {
							if (html.ErrorMessage.length > 0) {
								//A handled error was returned. Display error message, then redirect accordingly...
								handleAjaxError(html);
								return false;
							}
						}
						
					} catch (e) {
					}

					//clear the frame...
					$frame.html('');

					if (asyncFlag == true) {
						$(".popup").dialog("open");
					}

					//OK
					$frame.html(html);


					//jQuery styling
					$(function () {
						$("input[type=submit], input[type=button], button").button();
						$("input").addClass("ui-widget ui-corner-all");
						$("input").removeClass("text");

						$("textarea").addClass("ui-widget ui-corner-tl ui-corner-bl");
						$("textarea").removeClass("text");

						$("select").addClass("ui-widget ui-corner-tl ui-corner-bl");
						$("select").removeClass("text");
						$("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
						$("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");

					});

				},
				error: function (req, status, errorObj) {
					$("#errorDialogTitle").text(errorObj);
					$("#errorDialogContentText").html(req.responseText);
					$("#errorDialog").dialog("open");
				}
			});
		},
		showPopup = function (prompt) {

		},
		getFrame = function (frameId) {
			return document.frames[frameId];
		},
		getForm = function (frameId, formId) {
			//return document.forms[formId];

			return document.querySelector('#' + frameId + ' #' + formId);

		},
		submitForm = function (form, targetWin, asyncFlag) {
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
				dataType: 'html',
				data: data,
				async: asyncFlag,
				success: function (html) {
					try {
						var jsonResponse = $.parseJSON(html);
						if (jsonResponse.ErrorMessage.length > 0) {
							handleAjaxError(jsonResponse);
							return false;
						}
					} catch (e) {
					}
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
						$("#errorDialogContentText").html(e.responseText);
						$("#errorDialog").dialog("open");
					}
					//clear the frame...

					//OK

					if (targetWin != null) {

						//$frame = $form.closest("div[" + targetWin + "]").first();
						$frame = $("#" + targetWin);
						$frame.html(html);
						$frame.show();

					} else {
						$frame.html('');
						$frame.html(html);
					}

					//jQuery styling
					$(function () {
						$("input[type=submit], input[type=button], button").button();
						$("input").addClass("ui-widget ui-corner-all");
						$("input").removeClass("text");

						$("textarea").addClass("ui-widget ui-corner-tl ui-corner-bl");
						$("textarea").removeClass("text");

						$("select").addClass("ui-widget ui-corner-tl ui-corner-bl");
						$("select").removeClass("text");
						$("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
						$("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");						
					});

				},
				error: function (req, status, errorObj) {
					//alert("OpenHR.submitForm ajax call to '" + url + "' failed with '" + errorObj + "'.");

					//Sometimes (when?) an error is thrown with both errorObj and/or req.Response being empty; in this case don't show the empty error window
					if (!(errorObj == "" || req.responseText == "")) {
						//alert("OpenHR.submitForm ajax call to '" + url + "' failed with '" + errorObj + "'.");
						$("#errorDialogTitle").text(errorObj);
						$("#errorDialogContentText").html(req.responseText);
						$("#errorDialog").dialog("open");
					}
				}
			});
		},
		addActiveXHandler = function (controlId, eventName, func) {
			var ctl = document.getElementById(controlId);
			var handler;
			
			if (ctl != null) {
				if (eventName == "mouseUp") {
					handler = document.createElement("script");
					handler.setAttribute("for", controlId);
					handler.event = eventName + "(param1, param2, param3, param4)";
					handler.appendChild(document.createTextNode("javascript: " + func + ";"));
					document.body.appendChild(handler);
				} else {
					handler = document.createElement("script");
					handler.setAttribute("for", controlId);
					handler.event = eventName + "(param1, param2)";
					handler.appendChild(document.createTextNode("javascript: " + func + ";"));
					document.body.appendChild(handler);
				}
			}
		},
		refreshMenu = function () {
			//TODO
		},
		disableMenu = function () {
			//TODO
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
		printerCount = function () {
			//TODO
		},
		printerName = function (iLoop) {
			//TODO
		},
		getRegistrySetting = function (x, y, z) {
			return getCookie(z);
		},
		saveRegistrySetting = function (w, x, y, z) {
			setCookie(y, z, 365);
		},
		validateDir = function (x, y) {
			//TODO
			return true;
		},
		validateFilePath = function (sPath) {
			//TODO
			return true;
		},
		sendMail = function (sTo, sSubject, sBody, sCC, sBCC) {
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
		},
		isValidDate = function (d) {

			//TODO - Get the proper regional settings
			if (Date.parseExact(d, "d/M/yyyy") == null) {
				return false;
			}
			return true;
		},
		localeDateFormat = function () {
			//TODO - Get the proper regional settings
				return "dd/MM/yyyy";
		},
				convertSqlDateToLocale = function (z) {
			//TODO - Get the proper regional settings
			var convertDate = Date.parseExact(z, "M/d/yyyy");
			if (convertDate != null) {
						return convertDate.toString(OpenHR.LocaleDateFormat());
			} else {
				return "";
			}
				},
		convertLocaleDateToSQL = function (psDateString) {
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

			if (!isValidDate(psDateString)) return "";

			sDateFormat = OpenHR.LocaleDateFormat();

			sDays = "";
			sMonths = "";
			sYears = "";
			iValuePos = 0;

			// Trim leading spaces.
			sTempValue = psDateString.substr(iValuePos, 1);
			while (sTempValue.charAt(0) == " ") {
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);
			}

			for (iLoop = 0; iLoop < sDateFormat.length; iLoop++) {
				if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'D') && (sDays.length == 0)) {
					sDays = psDateString.substr(iValuePos, 1);
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos, 1);

					if (isNaN(sTempValue) == false) {
						sDays = sDays.concat(sTempValue);
					}
					iValuePos = iValuePos + 1;
				}

				if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'M') && (sMonths.length == 0)) {
					sMonths = psDateString.substr(iValuePos, 1);
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos, 1);

					if (isNaN(sTempValue) == false) {
						sMonths = sMonths.concat(sTempValue);
					}
					iValuePos = iValuePos + 1;
				}

				if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'Y') && (sYears.length == 0)) {
					sYears = psDateString.substr(iValuePos, 1);
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos, 1);

					if (isNaN(sTempValue) == false) {
						sYears = sYears.concat(sTempValue);
					}
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos, 1);

					if (isNaN(sTempValue) == false) {
						sYears = sYears.concat(sTempValue);
					}
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos, 1);

					if (isNaN(sTempValue) == false) {
						sYears = sYears.concat(sTempValue);
					}
					iValuePos = iValuePos + 1;
				}

				// Skip non-numerics
				sTempValue = psDateString.substr(iValuePos, 1);
				while (isNaN(sTempValue) == true) {
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos, 1);
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
				} else {
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
			} else {
				return sTempValue;
			}
		},
		getFileNameOnly = function (pstrFilePath) {
			//Extracts just the filename from a path
			var astrPath = pstrFilePath.split("\\");
			return astrPath[astrPath.length - 1];
		},
		ConvertToUNC = function (pstrFileName) {
			//TODO 
			return pstrFileName;
		},
		GetPathOnly = function (pstrFilePath, pbStripDriveLetter) {
			var L = pstrFilePath.length;

			while (L > 0) {
				var tempchar = pstrFilePath.substr(L, 1);
				if (tempchar == "\\") {
					var strPath = pstrFilePath.substr(0, L - 1);

					//Strip off drive letter
					if ((pbStripDriveLetter) && (strPath.substr(2, 1) == ":")) {
						strPath = strPath.substring(3, strPath.length);
					}

					return strPath;
				}
				L -= 1;
			}
		},
		CheckOLEFileNameLength = function (strFilename) {
			var bOK = true;

			//defined maximum filename length of 50
			if (getFileNameOnly(strFilename).length > 50) {
				return 'File name is too long.\nMaximum file length is 50 characters.';
			}

			if (GetPathOnly(strFilename, true).length > 100) {
				return 'Directory structure is too long.\nMaximum length is 50 characters.';
			}

			if ($.trim(ConvertToUNC(strFilename)).length > 50) {
				return 'Network path is too long.\nMaximum length is 50 characters.';
			}

			return '';
		},
	getCookie = function (c_name) {
		var i, x, y, ARRcookies = document.cookie.split(";");
		for (i = 0; i < ARRcookies.length; i++) {
			x = ARRcookies[i].substr(0, ARRcookies[i].indexOf("="));
			y = ARRcookies[i].substr(ARRcookies[i].indexOf("=") + 1);
			x = x.replace(/^\s+|\s+$/g, "");
			if (x == c_name) {
				if (y === undefined || y === "undefined" || y.length <= 0) y = "";
				return unescape(y);
			}
		}
		return '';
		},
		getFileExtension = function(strFilename) {
			return strFilename.substr(strFilename.lastIndexOf('.') + 1);

		};

	window.OpenHR = {
		version: version,
		messageBox: messageBox,
		modalPrompt: modalPrompt,
		modalMessage: modalMessage,
		showPopup: showPopup,
		getFrame: getFrame,
		getForm: getForm,
		submitForm: submitForm,
		showInReportFrame: showInReportFrame,
		addActiveXHandler: addActiveXHandler,
		refreshMenu: refreshMenu,
		disableMenu: disableMenu,
		LocaleDateFormat: localeDateFormat,
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
		MM_WORD_CreateTemplateFile: mmwordCreateTemplateFile,
		convertLocaleDateToSQL: convertLocaleDateToSQL,
		getFileNameOnly: getFileNameOnly,
		ConvertToUNC: ConvertToUNC,
		GetPathOnly: GetPathOnly,		
		getCookie: getCookie,
		CheckOLEFileNameLength: CheckOLEFileNameLength,
		GetFileExtension: getFileExtension
	};

})(window, jQuery);