

(function (window, $) {
	"use strict";

	function setDatepickerLanguage() {

		var language = window.navigator.userLanguage || window.navigator.language;

		if ($($.datepicker.regional[language]).length > 0) {
			//language found - use it.
			$.datepicker.setDefaults($.datepicker.regional[language]);
		} else {
			if ($($.datepicker.regional[language.substr(0, 2)]).length > 0) {
				//language found using code only - use it.
				$.datepicker.setDefaults($.datepicker.regional[language.substr(0, 2)]);
			} else {
				//english.
				$.datepicker.setDefaults($.datepicker.regional["en-GB"]);
			}
		}
	}

	function checkForMessages() {
		
		var frmMessage = OpenHR.getForm("divPollMessage", "frmPollMessage");

		try {
			$('#txtIsSessionTimeout').val('false');
			OpenHR.submitForm(frmMessage, "divPollMessage");
		} catch(e) {
		}
	}

	function handleAjaxError(html) {
		//handle error
		messageBox(html.ErrorMessage.replace("<p>", "\n\n"), 48, html.ErrorTitle);

		//redirect if specified...
		if (html.Redirect.length > 0) {
			//alert("redirecting to " + html.Redirect);
			window.location.href = html.Redirect;
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
			$('#dialog-confirm p').html(prompt);
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
			var dialogButtons;
			switch (buttons) {
			case 1:
				dialogButtons = {
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
				dialogButtons = {
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
				case 4:
					dialogButtons = {
						"Yes": function() {
							defer.resolve(6);
							$(this).dialog("close");
							if (followOnFunctionName) followOnFunctionName(6);
						},
						"No": function() {
							defer.resolve(7);
							$(this).dialog("close");
							if (followOnFunctionName) followOnFunctionName(7);
						}
					};
					break;
			default:
				dialogButtons = {
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
						$(".popup").dialog('option', 'title', '');
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
			
			$.ajax({
				type: "POST",
				url: "SendEmail",
				data: { 'to': sTo, 'cc': sCC, 'bcc': sBCC, 'subject': sSubject, 'body': sBody },
				dataType: "text",
				success: function (html) {
					alert("Email sent successfully");
				},
				error: function (req, status, errorObj) {
					if (!(errorObj == "" || req.responseText == "")) {
						$("#errorDialogTitle").text(errorObj);
						$("#errorDialogContentText").html(req.responseText);
						$("#errorDialog").dialog("open");
					}
				}

			});


		},
		currentWorkPage = function () {
			var sCurrentPage;
			if (!($("#workframe").css('display') == 'none')) {
				//Work frame is in view.
				sCurrentPage = $("#workframe").attr("data-framesource");
			} else {
				//Option frame is in view.
				sCurrentPage = $("#optionframe").attr("data-framesource");
			}

			try {
				sCurrentPage = sCurrentPage.toUpperCase();
			} catch(e) {}

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

			if (!isValidDate(psDateString)) return "null";

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

		},
	printDiv = function (divID) {
		//Creates a new window, copies the specified div contents to it and sends to printer.
		var divToPrint = document.getElementById(divID);
		var newWin = window.open("", "_blank", 'toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes, width=1, height=1, visible=none', "");
		newWin.document.write('<sty');
		newWin.document.write('le>');
		newWin.document.write('</sty');
		newWin.document.write('le>');
		newWin.document.write(divToPrint.innerHTML);
		newWin.document.write('<scri');
		newWin.document.write('pt type="text/javascript">');
		newWin.document.write('</scri');
		newWin.document.write('pt>');
		newWin.document.close();
		newWin.focus();
		newWin.print();
		newWin.close();
	},
	nullsafeString = function(arg) {
		var returnvalue = "";
		if ((arg == undefined) || (arg == "") || arg.length <= 0) {
			return returnvalue;
		} else {			
			try {
				returnvalue = String(arg);
			} catch(e) {
				return returnvalue;
			}
		}
		return returnvalue;
	},
	sessionTimeout = function() {
		var frmMessage = OpenHR.getForm("divPollMessage", "frmPollMessage");
		try {
			$('#txtIsSessionTimeout').val('true');
			frmMessage.action = "TimedOut";
			OpenHR.submitForm(frmMessage, "divPollMessage");
		} catch (e) {
		}
	},
	replaceAll = function (string, searchValue, newValue) {
		if ((searchValue.length == 0) || (string.length == 0)) return string;
		return string.split(searchValue).join(newValue);
	},
getLocaleDateString = function () {

	var formats = {
		"af-ZA": "Y/m/d",
		"am-ET": "d/m/Y",
		"ar-AE": "d/m/Y",
		"ar-BH": "d/m/Y",
		"ar-DZ": "d-m-Y",
		"ar-EG": "d/m/Y",
		"ar-IQ": "d/m/Y",
		"ar-JO": "d/m/Y",
		"ar-KW": "d/m/Y",
		"ar-LB": "d/m/Y",
		"ar-LY": "d/m/Y",
		"ar-MA": "d-m-Y",
		"arn-CL": "d-m-Y",
		"ar-OM": "d/m/Y",
		"ar-QA": "d/m/Y",
		"ar-SA": "d/m/Y",
		"ar-SY": "d/m/Y",
		"ar-TN": "d-m-Y",
		"ar-YE": "d/m/Y",
		"as-IN": "d-m-Y",
		"az-Cyrl-AZ": "d.m.Y",
		"az-Latn-AZ": "d.m.Y",
		"ba-RU": "d.m.Y",
		"be-BY": "d.m.Y",
		"bg-BG": "d.m.Y",
		"bn-BD": "d-m-Y",
		"bn-IN": "d-m-Y",
		"bo-CN": "Y/m/d",
		"br-FR": "d/m/Y",
		"bs-Cyrl-BA": "d.m.Y",
		"bs-Latn-BA": "d.m.Y",
		"ca-ES": "d/m/Y",
		"co-FR": "d/m/Y",
		"cs-CZ": "d.m.Y",
		"cy-GB": "d/m/Y",
		"da-DK": "d-m-Y",
		"de-AT": "d.m.Y",
		"de-CH": "d.m.Y",
		"de-DE": "d.m.Y",
		"de-LI": "d.m.Y",
		"de-LU": "d.m.Y",
		"dsb-DE": "d. m. Y",
		"dv-MV": "d/m/Y",
		"el-GR": "d/m/Y",
		"en-029": "m/d/Y",
		"en-AU": "d/m/Y",
		"en-BZ": "d/m/Y",
		"en-CA": "d/m/Y",
		"en-GB": "d/m/Y",
		"en-IE": "d/m/Y",
		"en-IN": "d-m-Y",
		"en-JM": "d/m/Y",
		"en-MY": "d/m/Y",
		"en-NZ": "d/m/Y",
		"en-PH": "m/d/Y",
		"en-SG": "d/m/Y",
		"en-TT": "d/m/Y",
		"en-US": "m/d/Y",
		"en-ZA": "Y/m/d",
		"en-ZW": "m/d/Y",
		"es-AR": "d/m/Y",
		"es-BO": "d/m/Y",
		"es-CL": "d-m-Y",
		"es-CO": "d/m/Y",
		"es-CR": "d/m/Y",
		"es-DO": "d/m/Y",
		"es-EC": "d/m/Y",
		"es-ES": "d/m/Y",
		"es-GT": "d/m/Y",
		"es-HN": "d/m/Y",
		"es-MX": "d/m/Y",
		"es-NI": "d/m/Y",
		"es-PA": "m/d/Y",
		"es-PE": "d/m/Y",
		"es-PR": "d/m/Y",
		"es-PY": "d/m/Y",
		"es-SV": "d/m/Y",
		"es-US": "m/d/Y",
		"es-UY": "d/m/Y",
		"es-VE": "d/m/Y",
		"et-EE": "d.m.Y",
		"eu-ES": "Y/m/d",
		"fa-IR": "m/d/Y",
		"fi-FI": "d.m.Y",
		"fil-PH": "m/d/Y",
		"fo-FO": "d-m-Y",
		"fr-BE": "d/m/Y",
		"fr-CA": "Y-m-d",
		"fr-CH": "d.m.Y",
		"fr-FR": "d/m/Y",
		"fr-LU": "d/m/Y",
		"fr-MC": "d/m/Y",
		"fy-NL": "d-m-Y",
		"ga-IE": "d/m/Y",
		"gd-GB": "d/m/Y",
		"gl-ES": "d/m/Y",
		"gsw-FR": "d/m/Y",
		"gu-IN": "d-m-Y",
		"ha-Latn-NG": "d/m/Y",
		"he-IL": "d/m/Y",
		"hi-IN": "d-m-Y",
		"hr-BA": "d.m.Y.",
		"hr-HR": "d.m.Y",
		"hsb-DE": "d. m. Y",
		"hu-HU": "Y. m. d.",
		"hy-AM": "d.m.Y",
		"id-ID": "d/m/Y",
		"ig-NG": "d/m/Y",
		"ii-CN": "Y/m/d",
		"is-IS": "d.m.Y",
		"it-CH": "d.m.Y",
		"it-IT": "d/m/Y",
		"iu-Cans-CA": "d/m/Y",
		"iu-Latn-CA": "d/m/Y",
		"ja-JP": "Y/m/d",
		"ka-GE": "d.m.Y",
		"kk-KZ": "d.m.Y",
		"kl-GL": "d-m-Y",
		"km-KH": "Y-m-d",
		"kn-IN": "d-m-Y",
		"kok-IN": "d-m-Y",
		"ko-KR": "Y-m-d",
		"ky-KG": "d.m.Y",
		"lb-LU": "d/m/Y",
		"lo-LA": "d/m/Y",
		"lt-LT": "Y.m.d",
		"lv-LV": "Y.m.d.",
		"mi-NZ": "d/m/Y",
		"mk-MK": "d.m.Y",
		"ml-IN": "d-m-Y",
		"mn-MN": "Y.m.d",
		"mn-Mong-CN": "Y/m/d",
		"moh-CA": "m/d/Y",
		"mr-IN": "d-m-Y",
		"ms-BN": "d/m/Y",
		"ms-MY": "d/m/Y",
		"mt-MT": "d/m/Y",
		"nb-NO": "d.m.Y",
		"ne-NP": "m/d/Y",
		"nl-BE": "d/m/Y",
		"nl-NL": "d-m-Y",
		"nn-NO": "d.m.Y",
		"nso-ZA": "Y/m/d",
		"oc-FR": "d/m/Y",
		"or-IN": "d-m-Y",
		"pa-IN": "d-m-Y",
		"pl-PL": "Y-m-d",
		"prs-AF": "d/m/Y",
		"ps-AF": "d/m/Y",
		"pt-BR": "d/m/Y",
		"pt-PT": "d-m-Y",
		"qut-GT": "d/m/Y",
		"quz-BO": "d/m/Y",
		"quz-EC": "d/m/Y",
		"quz-PE": "d/m/Y",
		"rm-CH": "d/m/Y",
		"ro-RO": "d.m.Y",
		"ru-RU": "d.m.Y",
		"rw-RW": "m/d/Y",
		"sah-RU": "m.d.Y",
		"sa-IN": "d-m-Y",
		"se-FI": "d.m.Y",
		"se-NO": "d.m.Y",
		"se-SE": "Y-m-d",
		"si-LK": "Y-m-d",
		"sk-SK": "d. m. Y",
		"sl-SI": "d.m.Y",
		"sma-NO": "d.m.Y",
		"sma-SE": "Y-m-d",
		"smj-NO": "d.m.Y",
		"smj-SE": "Y-m-d",
		"smn-FI": "d.m.Y",
		"sms-FI": "d.m.Y",
		"sq-AL": "Y-m-d",
		"sr-Cyrl-BA": "d.m.Y",
		"sr-Cyrl-CS": "d.m.Y",
		"sr-Cyrl-ME": "d.m.Y",
		"sr-Cyrl-RS": "d.m.Y",
		"sr-Latn-BA": "d.m.Y",
		"sr-Latn-CS": "d.m.Y",
		"sr-Latn-ME": "d.m.Y",
		"sr-Latn-RS": "d.m.Y",
		"sv-FI": "d.m.Y",
		"sv-SE": "Y-m-d",
		"sw-KE": "m/d/Y",
		"syr-SY": "d/m/Y",
		"ta-IN": "d-m-Y",
		"te-IN": "d-m-Y",
		"tg-Cyrl-TJ": "d.m.Y",
		"th-TH": "d/m/Y",
		"tk-TM": "d.m.Y",
		"tn-ZA": "Y/m/d",
		"tr-TR": "d.m.Y",
		"tt-RU": "d.m.Y",
		"tzm-Latn-DZ": "d-m-Y",
		"ug-CN": "Y-m-d",
		"uk-UA": "d.m.Y",
		"ur-PK": "d/m/Y",
		"uz-Cyrl-UZ": "d.m.Y",
		"uz-Latn-UZ": "d/m Y",
		"vi-VN": "d/m/Y",
		"wo-SN": "d/m/Y",
		"xh-ZA": "Y/m/d",
		"yo-NG": "d/m/Y",
		"zh-CN": "Y/m/d",
		"zh-HK": "d/m/Y",
		"zh-MO": "d/m/Y",
		"zh-SG": "d/m/Y",
		"zh-TW": "Y/m/d",
		"zu-ZA": "Y/m/d"
	};

	//return formats[navigator.language] || 'dd/MM/yyyy';
	return formats[window.UserLocale] || 'dd/MM/yyyy';

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
		GetFileExtension: getFileExtension,
		CheckForMessages: checkForMessages,
		SessionTimeout: sessionTimeout,
		printDiv: printDiv,
		nullsafeString: nullsafeString,
		replaceAll: replaceAll,
		getLocaleDateString: getLocaleDateString,
		setDatepickerLanguage: setDatepickerLanguage,
	};

})(window, jQuery);