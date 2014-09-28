

(function (window, $) {
	"use strict";

	function setDatepickerLanguage() {

		var language = window.UserLocale || window.dialogArguments.window.UserLocale;

		// No regional setting for US - assumed as the default.
		if (language == "en-US") {
			$.datepicker.setDefaults($.datepicker.regional[""]);
		}
		else if ($($.datepicker.regional[language]).length > 0) {
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

		modalExpressionSelect = function (type, tableId, currentID, followOnFunctionName, screenwidth, screenheight) {

			var frame = $("#divExpressionSelection");
			var capitalizedText = capitalizeMe(type);

			if (type == 'CALC') {
				capitalizedText = 'Calculations';
			}
			
			$("#ExpressionsAvailable").jqGrid('GridUnload');
			$("#ExpressionsAvailable").jqGrid({
				url: 'Reports/GetExpressionsForTable?TableID=' + tableId + '&&selectionType=' + type,
				datatype: 'json',
				mtype: 'GET',
				jsonReader: {
					root: "rows", //array containing actual data
					page: "page", //current page
					total: "total", //total pages for the query
					records: "records", //total number of records
					repeatitems: false,
					id: "ID"
				},
				colNames: ['ID', 'Name', 'Description', 'Access'],
				colModel: [
					{ name: 'ID', index: 'ID', hidden: true },
					{ name: 'Name', index: 'Name', width: 40, sortable: false },
					{ name: 'Description', index: 'Description', hidden: true },
					{ name: 'Access', index: 'Access', hidden: true }],
				viewrecords: true,
				width: screenwidth,
				height: screenheight,
				sortname: 'Name',
				sortorder: "desc",
				rowNum: 10000,
				scrollrows: true,
				onSelectRow: function () {
					button_disable($('#ExpressionSelectOK'), false);
				},
				ondblClickRow: function (rowid) {
					var gridData = $(this).getRowData(rowid);
					followOnFunctionName(gridData.ID, gridData.Name, gridData.Access);
					frame.dialog("close");
				},
				loadComplete: function(json) {

					button_disable($('#ExpressionSelectOK'), true);

					$("#ExpressionSelectOK").off('click').on('click', function() {
						var rowid = $('#ExpressionsAvailable').jqGrid('getGridParam', 'selrow');
						var gridData = $("#ExpressionsAvailable").getRowData(rowid);
						followOnFunctionName(gridData.ID, gridData.Name, gridData.Access);
						frame.dialog("close");
					});

					$("#ExpressionSelectCancel").off('click').on('click', function () {					
						frame.dialog("close");
					});

					$("#ExpressionSelectNone").off('click').on('click', function () {					
						followOnFunctionName(0, "None", "RW");
						frame.dialog("close");
					});

					$("#ExpressionsAvailable").jqGrid("setSelection", currentID);
					$("#ExpressionSelection_PageTitle").text(capitalizedText);					
				}
			});

			//$frame.html(html);
			frame.show();
			frame.dialog("open");

			function capitalizeMe(val) {
				return val.charAt(0).toUpperCase() + val.substr(1).toLowerCase();
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


		modalPrompt = function (prompt, buttons, title, followOnFunctionName) {
			var defer = $.Deferred();
			var dialogButtons;
			switch (buttons) {
				case 0:
					dialogButtons = {
						"OK": function() {
							defer.resolve(1);
							$(this).dialog("close");
							if (followOnFunctionName) followOnFunctionName(1);
						}
					};
					break;

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
						//$(".popup").dialog('option', 'title', '');
						$(".popup").dialog({ dialogClass: 'no-close' });
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

		getFrame = function (frameId) {
			return document.frames[frameId];
		},
		getForm = function (frameId, formId) {
			//return document.forms[formId];

			return document.querySelector('#' + frameId + ' #' + formId);

		},

		postData = function (url, jsonData, followOnFunctionName) {

			$.ajax({
				url: url,
				type: "POST",
				async: true,
				data: jsonData,

				success: function (html) {

					try {
						var jsonResponse = $.parseJSON(html);
						if (jsonResponse.ErrorMessage.length > 0) {
							handleAjaxError(jsonResponse);
							return false;
						}
					} catch (e) {
					}

					if (followOnFunctionName) followOnFunctionName(html);

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

				},
				error: function (req, status, errorObj) {
					if (!(errorObj == "" || req.responseText == "")) {
						//alert("OpenHR.submitForm ajax call to '" + url + "' failed with '" + errorObj + "'.");
						$("#errorDialogTitle").text(errorObj);
						$("#errorDialogContentText").html(req.responseText);
						$("#errorDialog").dialog("open");
					}
				}
			});


		},

	openDialog = function (url, targetWin, jsonData, dialogWidth) { //dialogWidth should be passed as a string, not a number: i.e 'auto' or '900px'

		var $frame;

		$.ajax({
			url: url,
			type: "POST",
			data: JSON.stringify(jsonData),
			contentType: "application/json;charset=utf-8",
			async: true,
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
					//alert("OpenHR.submitForm ajax call to '" + url + "' failed with '" + e.toString() + "'.");
					$("#errorDialogTitle").text(e.toString);
					$("#errorDialogContentText").html(e.responseText);
					$("#errorDialog").dialog("open");
				}

				$frame = $("#" + targetWin);
				$frame.html(html);
				$frame.dialog('option', 'width', dialogWidth);

				$frame.show();
				$frame.dialog({ position: { 'my': 'top', 'at': 'center center', 'of': 'header' } });
				$frame.dialog("open");

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


	submitForm = function (form, targetWin, asyncFlag, jsonData, action) {

		var $form = $(form),
			$frame = $form.closest("div[data-framesource]").first(),
			target = $form.attr("target"),
			method = $form.attr("method");

		var data;
		var url;

		if (action == undefined) {
			url = $form.attr("action");
		}	else {
			url = action;
		}


		if (jsonData == undefined) {
			data = $form.serialize();
		}	else {
			data = jsonData;
		}
	

			if ((asyncFlag == undefined) || (asyncFlag.length == 0) || (asyncFlag == true)) {
				asyncFlag = true;
			} else {
				asyncFlag = false;
			}

			$.ajax({
				url: url,
				type: method,
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
			return window.LocaleDecimalSeparator;
		},
		localeThousandSeparator = function () {
			return window.LocaleThousandSeparator;
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
		currentWorkPage = function () {
			var sCurrentPage;
			
			if (!($("#workframe").css('display') == 'none')) {
				//Work frame is in view.
				sCurrentPage = $("#workframe").attr("data-framesource");
			} else {
				//Option frame is in view.
				sCurrentPage = $("#optionframe").attr("data-framesource");
			}

			//Popout optionframe check
			try {
				if ($("#optionframe").dialog("isOpen") == true) {
					sCurrentPage = $("#optionframe").attr("data-framesource");
				}
			} catch (e) {}

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
			
			var localeFormat = window.LocaleDateFormat || window.dialogArguments.window.LocaleDateFormat;

			if (Date.parseExact(d, localeFormat) == null) {
				return false;
			}
			return true;
		},

	localeDateFormat = function () {

		var formats = {
			"ar-SA" : "dd/MM/yy",
			"bg-BG" : "dd.M.yyyy",
			"ca-ES" : "dd/MM/yyyy",
			"zh-TW" : "yyyy/M/d",
			"cs-CZ" : "d.M.yyyy",
			"da-DK" : "dd-MM-yyyy",
			"de-DE" : "dd.MM.yyyy",
			"el-GR" : "d/M/yyyy",
			"en-US" : "M/d/yyyy",
			"fi-FI" : "d.M.yyyy",
			"fr-FR" : "dd/MM/yyyy",
			"he-IL" : "dd/MM/yyyy",
			"hu-HU" : "yyyy. MM. dd.",
			"is-IS" : "d.M.yyyy",
			"it-IT" : "dd/MM/yyyy",
			"ja-JP" : "yyyy/MM/dd",
			"ko-KR" : "yyyy-MM-dd",
			"nl-NL" : "d-M-yyyy",
			"nb-NO" : "dd.MM.yyyy",
			"pl-PL" : "yyyy-MM-dd",
			"pt-BR" : "d/M/yyyy",
			"ro-RO" : "dd.MM.yyyy",
			"ru-RU" : "dd.MM.yyyy",
			"hr-HR" : "d.M.yyyy",
			"sk-SK" : "d. M. yyyy",
			"sq-AL" : "yyyy-MM-dd",
			"sv-SE" : "yyyy-MM-dd",
			"th-TH" : "d/M/yyyy",
			"tr-TR" : "dd.MM.yyyy",
			"ur-PK" : "dd/MM/yyyy",
			"id-ID" : "dd/MM/yyyy",
			"uk-UA" : "dd.MM.yyyy",
			"be-BY" : "dd.MM.yyyy",
			"sl-SI" : "d.M.yyyy",
			"et-EE" : "d.MM.yyyy",
			"lv-LV" : "yyyy.MM.dd.",
			"lt-LT" : "yyyy.MM.dd",
			"fa-IR" : "MM/dd/yyyy",
			"vi-VN" : "dd/MM/yyyy",
			"hy-AM" : "dd.MM.yyyy",
			"az-Latn-AZ" : "dd.MM.yyyy",
			"eu-ES" : "yyyy/MM/dd",
			"mk-MK" : "dd.MM.yyyy",
			"af-ZA" : "yyyy/MM/dd",
			"ka-GE" : "dd.MM.yyyy",
			"fo-FO" : "dd-MM-yyyy",
			"hi-IN" : "dd-MM-yyyy",
			"ms-MY" : "dd/MM/yyyy",
			"kk-KZ" : "dd.MM.yyyy",
			"ky-KG" : "dd.MM.yy",
			"sw-KE" : "M/d/yyyy",
			"uz-Latn-UZ" : "dd/MM yyyy",
			"tt-RU" : "dd.MM.yyyy",
			"pa-IN" : "dd-MM-yy",
			"gu-IN" : "dd-MM-yy",
			"ta-IN" : "dd-MM-yyyy",
			"te-IN" : "dd-MM-yy",
			"kn-IN" : "dd-MM-yy",
			"mr-IN" : "dd-MM-yyyy",
			"sa-IN" : "dd-MM-yyyy",
			"mn-MN" : "yy.MM.dd",
			"gl-ES" : "dd/MM/yy",
			"kok-IN" : "dd-MM-yyyy",
			"syr-SY" : "dd/MM/yyyy",
			"dv-MV" : "dd/MM/yy",
			"ar-IQ" : "dd/MM/yyyy",
			"zh-CN" : "yyyy/M/d",
			"de-CH" : "dd.MM.yyyy",
			"en-GB" : "dd/MM/yyyy",
			"es-MX" : "dd/MM/yyyy",
			"fr-BE" : "d/MM/yyyy",
			"it-CH" : "dd.MM.yyyy",
			"nl-BE" : "d/MM/yyyy",
			"nn-NO" : "dd.MM.yyyy",
			"pt-PT" : "dd-MM-yyyy",
			"sr-Latn-CS" : "d.M.yyyy",
			"sv-FI" : "d.M.yyyy",
			"az-Cyrl-AZ" : "dd.MM.yyyy",
			"ms-BN" : "dd/MM/yyyy",
			"uz-Cyrl-UZ" : "dd.MM.yyyy",
			"ar-EG" : "dd/MM/yyyy",
			"zh-HK" : "d/M/yyyy",
			"de-AT" : "dd.MM.yyyy",
			"en-AU" : "d/MM/yyyy",
			"es-ES" : "dd/MM/yyyy",
			"fr-CA" : "yyyy-MM-dd",
			"sr-Cyrl-CS" : "d.M.yyyy",
			"ar-LY" : "dd/MM/yyyy",
			"zh-SG" : "d/M/yyyy",
			"de-LU" : "dd.MM.yyyy",
			"en-CA" : "dd/MM/yyyy",
			"es-GT" : "dd/MM/yyyy",
			"fr-CH" : "dd.MM.yyyy",
			"ar-DZ" : "dd-MM-yyyy",
			"zh-MO" : "d/M/yyyy",
			"de-LI" : "dd.MM.yyyy",
			"en-NZ" : "d/MM/yyyy",
			"es-CR" : "dd/MM/yyyy",
			"fr-LU" : "dd/MM/yyyy",
			"ar-MA" : "dd-MM-yyyy",
			"en-IE" : "dd/MM/yyyy",
			"es-PA" : "MM/dd/yyyy",
			"fr-MC" : "dd/MM/yyyy",
			"ar-TN" : "dd-MM-yyyy",
			"en-ZA" : "yyyy/MM/dd",
			"es-DO" : "dd/MM/yyyy",
			"ar-OM" : "dd/MM/yyyy",
			"en-JM" : "dd/MM/yyyy",
			"es-VE" : "dd/MM/yyyy",
			"ar-YE" : "dd/MM/yyyy",
			"en-029" : "MM/dd/yyyy",
			"es-CO" : "dd/MM/yyyy",
			"ar-SY" : "dd/MM/yyyy",
			"en-BZ" : "dd/MM/yyyy",
			"es-PE" : "dd/MM/yyyy",
			"ar-JO" : "dd/MM/yyyy",
			"en-TT" : "dd/MM/yyyy",
			"es-AR" : "dd/MM/yyyy",
			"ar-LB" : "dd/MM/yyyy",
			"en-ZW" : "M/d/yyyy",
			"es-EC" : "dd/MM/yyyy",
			"ar-KW" : "dd/MM/yyyy",
			"en-PH" : "M/d/yyyy",
			"es-CL" : "dd-MM-yyyy",
			"ar-AE" : "dd/MM/yyyy",
			"es-UY" : "dd/MM/yyyy",
			"ar-BH" : "dd/MM/yyyy",
			"es-PY" : "dd/MM/yyyy",
			"ar-QA" : "dd/MM/yyyy",
			"es-BO" : "dd/MM/yyyy",
			"es-SV" : "dd/MM/yyyy",
			"es-HN" : "dd/MM/yyyy",
			"es-NI" : "dd/MM/yyyy",
			"es-PR" : "dd/MM/yyyy",
			"am-ET" : "d/M/yyyy",
			"tzm-Latn-DZ" : "dd-MM-yyyy",
			"iu-Latn-CA" : "d/MM/yyyy",
			"sma-NO" : "dd.MM.yyyy",
			"mn-Mong-CN" : "yyyy/M/d",
			"gd-GB" : "dd/MM/yyyy",
			"en-MY" : "d/M/yyyy",
			"prs-AF" : "dd/MM/yy",
			"bn-BD" : "dd-MM-yy",
			"wo-SN" : "dd/MM/yyyy",
			"rw-RW" : "M/d/yyyy",
			"qut-GT" : "dd/MM/yyyy",
			"sah-RU" : "MM.dd.yyyy",
			"gsw-FR" : "dd/MM/yyyy",
			"co-FR" : "dd/MM/yyyy",
			"oc-FR" : "dd/MM/yyyy",
			"mi-NZ" : "dd/MM/yyyy",
			"ga-IE" : "dd/MM/yyyy",
			"se-SE" : "yyyy-MM-dd",
			"br-FR" : "dd/MM/yyyy",
			"smn-FI" : "d.M.yyyy",
			"moh-CA" : "M/d/yyyy",
			"arn-CL" : "dd-MM-yyyy",
			"ii-CN" : "yyyy/M/d",
			"dsb-DE" : "d. M. yyyy",
			"ig-NG" : "d/M/yyyy",
			"kl-GL" : "dd-MM-yyyy",
			"lb-LU" : "dd/MM/yyyy",
			"ba-RU" : "dd.MM.yy",
			"nso-ZA" : "yyyy/MM/dd",
			"quz-BO" : "dd/MM/yyyy",
			"yo-NG" : "d/M/yyyy",
			"ha-Latn-NG" : "d/M/yyyy",
			"fil-PH" : "M/d/yyyy",
			"ps-AF" : "dd/MM/yy",
			"fy-NL" : "d-M-yyyy",
			"ne-NP" : "M/d/yyyy",
			"se-NO" : "dd.MM.yyyy",
			"iu-Cans-CA" : "d/M/yyyy",
			"sr-Latn-RS" : "d.M.yyyy",
			"si-LK" : "yyyy-MM-dd",
			"sr-Cyrl-RS" : "d.M.yyyy",
			"lo-LA" : "dd/MM/yyyy",
			"km-KH" : "yyyy-MM-dd",
			"cy-GB" : "dd/MM/yyyy",
			"bo-CN" : "yyyy/M/d",
			"sms-FI" : "d.M.yyyy",
			"as-IN" : "dd-MM-yyyy",
			"ml-IN" : "dd-MM-yy",
			"en-IN" : "dd-MM-yyyy",
			"or-IN" : "dd-MM-yy",
			"bn-IN" : "dd-MM-yy",
			"tk-TM" : "dd.MM.yy",
			"bs-Latn-BA" : "d.M.yyyy",
			"mt-MT" : "dd/MM/yyyy",
			"sr-Cyrl-ME" : "d.M.yyyy",
			"se-FI" : "d.M.yyyy",
			"zu-ZA" : "yyyy/MM/dd",
			"xh-ZA" : "yyyy/MM/dd",
			"tn-ZA" : "yyyy/MM/dd",
			"hsb-DE" : "d. M. yyyy",
			"bs-Cyrl-BA" : "d.M.yyyy",
			"tg-Cyrl-TJ" : "dd.MM.yy",
			"sr-Latn-BA" : "d.M.yyyy",
			"smj-NO" : "dd.MM.yyyy",
			"rm-CH" : "dd/MM/yyyy",
			"smj-SE" : "yyyy-MM-dd",
			"quz-EC" : "dd/MM/yyyy",
			"quz-PE" : "dd/MM/yyyy",
			"hr-BA" : "d.M.yyyy.",
			"sr-Latn-ME" : "d.M.yyyy",
			"sma-SE" : "yyyy-MM-dd",
			"en-SG" : "d/M/yyyy",
			"ug-CN" : "yyyy-M-d",
			"sr-Cyrl-BA" : "d.M.yyyy",
			"es-US" : "M/d/yyyy"
		};

		var language = window.UserLocale || window.dialogArguments.window.UserLocale;
		return formats[language] || 'dd/MM/yyyy';

	} ,

		convertSqlDateToLocale = function (z) {

			var convertDate = Date.parseExact(z, "M/d/yyyy");
			if (convertDate != null) {
				return convertDate.toString(window.LocaleDateFormat);
			} else {
				return "";
			}
		},

		convertLocaleDateToSQL = function(psDateString) {

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

			//TODO - This is not good, as it will mean the server will return a "Conversion failed when converting date and/or time from character string" if the user puts in
			// garbage data. Our problem is that at present there is no validation of the record before its is sent. Checking validtity of dates
			// is something that the old ActivbeX control used to do. This is just to get things running.
			if (psDateString == "") return "null";

			if (!isValidDate(psDateString)) {
			return psDateString;
		}
	
			sDateFormat = window.LocaleDateFormat || window.dialogArguments.window.LocaleDateFormat;

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
	printDiv = function (divID, cssObj) {
		//Creates a new window, copies the specified div contents to it and sends to printer.
		var divToPrint = document.getElementById(divID);
		var newWin = window.open("", "_blank", 'toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes, width=1, height=1, visible=none', "");
		newWin.document.write('<sty');
		newWin.document.write('le>');

		if (cssObj) {
			for (var i = 0; i < cssObj.length; i++) {
				newWin.document.write(cssObj[i].toString());
			}
		}

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

	moveItemInGrid = function (grid, direction) {

		if (grid.getGridParam('selrow')) {

			var ids = grid.getDataIDs();
			var currRow = grid.getGridParam('selrow');
			var index = grid.getInd(currRow) - 1;
			var rowData = grid.getRowData(ids[index]);

			var recordCount = grid.getGridParam("records");

			if (direction === 'up' && index > 0) {
				var rowAbove = grid.getRowData(ids[index - 1]);
				rowAbove.Sequence = parseInt(rowAbove.Sequence) + 1;
				grid.setRowData(rowAbove.ID, rowAbove);

				rowData.Sequence = parseInt(rowData.Sequence) - 1;
				grid.delRowData(rowData.ID);
				grid.addRowData(rowData.ID, rowData, 'before', rowAbove.ID);
			}
			if (direction === 'down' && index < recordCount) {
				var rowBelow = grid.getRowData(ids[index + 1]);
				rowBelow.Sequence = parseInt(rowBelow.Sequence) - 1;
				grid.setRowData(rowBelow.ID, rowBelow);

				rowData.Sequence = parseInt(rowData.Sequence) + 1;
				grid.delRowData(rowData.ID);
				grid.addRowData(rowData.ID, rowData, 'after', rowBelow.ID);
			}

			grid.jqGrid("setSelection", rowData.ID);
		}
	},
		
	getLocaleDateString = function () {

		var res = window.LocaleDateFormat.replace("dd", "d").replace("MM", "m").replace("M", "m").replace("yyyy", "Y");
		return res;
	},

	parentExists = function ()
	{
		//function to detect if this form is displayed in a dialog or a jquery modal div.
		//true = window.dialog.
		var opener = window.dialogArguments;
		return (opener == null)?false:true;
	},

	windowOpen = function (destination, width, height) {
		// From https://developer.mozilla.org/en-US/docs/Web/API/Window.open:
		// "only list the features to be enabled or rendered; the others (except titlebar and close) will be disabled or removed."
		var windowProperties = "centerscreen,chrome," + "height=" + height + "," + "width=" + width;

		return window.open(destination, "_blank", windowProperties);
	};

	window.OpenHR = {
		version: version,
		messageBox: messageBox,
		modalPrompt: modalPrompt,
		modalMessage: modalMessage,
		getFrame: getFrame,
		getForm: getForm,
		postData: postData,
		submitForm: submitForm,
		showInReportFrame: showInReportFrame,
		addActiveXHandler: addActiveXHandler,
		refreshMenu: refreshMenu,
		disableMenu: disableMenu,
		LocaleDateFormat: localeDateFormat,
		LocaleDecimalSeparator: localeDecimalSeparator,
		LocaleThousandSeparator: localeThousandSeparator,
		ConvertSQLDateToLocale: convertSqlDateToLocale,
		PrinterCount: printerCount,
		PrinterName: printerName,
		GetRegistrySetting: getRegistrySetting,
		SaveRegistrySetting: saveRegistrySetting,
		ValidateDir: validateDir,
		ValidateFilePath: validateFilePath,
		currentWorkPage: currentWorkPage,
		MM_WORD_CreateTemplateFile: mmwordCreateTemplateFile,
		convertLocaleDateToSQL: convertLocaleDateToSQL,
		getFileNameOnly: getFileNameOnly,
		ConvertToUNC: ConvertToUNC,
		GetPathOnly: GetPathOnly,		
		getCookie: getCookie,
		CheckOLEFileNameLength: CheckOLEFileNameLength,
		GetFileExtension: getFileExtension,
		SessionTimeout: sessionTimeout,
		printDiv: printDiv,
		nullsafeString: nullsafeString,
		replaceAll: replaceAll,
		getLocaleDateString: getLocaleDateString,
		setDatepickerLanguage: setDatepickerLanguage,
		IsValidDate: isValidDate,
		MoveItemInGrid: moveItemInGrid,
		OpenDialog: openDialog,
		modalExpressionSelect: modalExpressionSelect,
		parentExists: parentExists,
		windowOpen: windowOpen
	};

})(window, jQuery);