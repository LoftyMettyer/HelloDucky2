

(function (window, $) {
	"use strict";

	function setDatepickerLanguage() {

		var language = window.top.window.UserLocale || window.opener.window.UserLocale;

		// No regional setting for US - assumed as the default.
		if ((language.toUpperCase() == "EN-US") || (language.toUpperCase() == "EN")) {
			$.datepicker.setDefaults($.datepicker.regional[""]);
		}
		else {
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
	}

	function handleAjaxError(html) {
		//handle error
		messageBox(html.ErrorMessage.replace("<p>", "\n\n"), 48, html.ErrorTitle);

		//redirect if specified...
		if (html.Redirect.length > 0) {
			if (html.Route === "" || html.Route === null) html.Route = "home";
			window.top.window.location.href = window.top.window.ROOT + html.Route + "/" + html.Redirect;
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

			window.top.$('#dialog-confirm').dialog('option', 'buttons', dialogButtons);
			//$('#dialog-confirm').dialog('option', 'title', title);
			window.top.$('#dialog-confirm p').html($.parseHTML(prompt));
			window.top.$('#dialog-confirm').dialog('open');
			// If Any ActiveX controls are in the workframeset, move the dailog to the very top of the screen to avoid it being hidden behind the ActiveX
			if (window.top.$('#workframeset object').length > 0) {
				window.top.$('#dialog-confirm').dialog('option', 'position', 'top');
			} else {
				window.top.$('#dialog-confirm').dialog('option', 'position', 'center');
			}

		},

		modalMessage = function (message, title) {
			var dialogButtons = {
				"OK": function () {
					window.top.$(this).dialog("close");
				}
			};

			displayModalDialog(message, dialogButtons, title);

		},

		modalExpressionSelect = function (type, tableId, currentID, followOnFunctionName, screenwidth, screenheight, returnResults, followOnFunctionToEnableSaveButton) {

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
				autoencode: true,
				onSelectRow: function () {
					button_disable($('#ExpressionSelectOK'), false);
				},
				ondblClickRow: function (rowid) {
					var gridData = $(this).getRowData(rowid);

					if (returnResults) {
						//launch promptedvalues to return filter result set.
						if (followOnFunctionToEnableSaveButton != undefined) { followOnFunctionToEnableSaveButton();}
						returnFilterResults(gridData);
						frame.dialog("close");
					} else {
						followOnFunctionName(gridData.ID, gridData.Name, gridData.Access);
						frame.dialog("close");
					}
				},
				loadComplete: function(json) {

					button_disable($('#ExpressionSelectOK'), true);

					$("#ExpressionSelectOK").off('click').on('click', function() {
						var rowid = $('#ExpressionsAvailable').jqGrid('getGridParam', 'selrow');
						var gridData = $("#ExpressionsAvailable").getRowData(rowid);

						if (returnResults) {
							//launch promptedvalues to return filter result set.
							if (followOnFunctionToEnableSaveButton != undefined) { followOnFunctionToEnableSaveButton(); }
							returnFilterResults(gridData);
							frame.dialog("close");
						} else {
							//Just return the filter name
							followOnFunctionName(gridData.ID, gridData.Name, gridData.Access);
							frame.dialog("close");
						}
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

			var dialogwidth = $("#ExpressionsAvailable").width();
			frame.dialog({
				width: dialogwidth + 50,
				modal: true
			});

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


			function returnFilterResults(gridData) {
				//launch promptedvalues to return filter result set.
				//OpenHR.clearTmpDialog();
				//$('body').append('<div id="tmpDialog"></div>');
				//$('#tmpDialog').dialog({
				//	width: 'auto',
				//	height: 'auto',
				//	modal: true
				//});

				var postData = {
					ID: gridData.ID,
					UtilType: utilityType.Filter,
					FilteredAdd: true,
					__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
				};
				OpenHR.submitForm(null, "reportframe", null, postData, "util_run_promptedValues");


				//$.ajax({
				//	url: "util_run_promptedValues",
				//	type: "POST",
				//	async: true,
				//	data: { ID: gridData.ID, UtilType: utilityType.Filter, __RequestVerificationToken: $('[name="__RequestVerificationToken"]').val() },
				//	success: function (html) {

				//		$('#tmpDialog').html('').html(html);

				//		//jQuery styling
				//		$(function () {
				//			$("input[type=submit], input[type=button], button").button();
				//			$("input").addClass("ui-widget ui-corner-all");
				//			$("input").removeClass("text");

				//			$("textarea").addClass("ui-widget ui-corner-tl ui-corner-bl");
				//			$("textarea").removeClass("text");

				//			$("select").addClass("ui-widget ui-corner-tl ui-corner-bl");
				//			$("select").removeClass("text");
				//			$("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
				//			$("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");

				//		});

				//		$('#tmpDialog').dialog("option", "position", ['center', 'center']);

				//	},
				//	error: function () { alert('error!!!!!'); }
				//});				
			}

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

		getFrame = function (frameId) {
			return document.frames[frameId];
		},
		getForm = function (frameId, formId) {
			//calculate floatingWindow
			var floatingWindow = false;
			if (!menu_isSSIMode()) {
				if (formId == "frmFind" || formId == "frmFindForm") floatingWindow = true;
				if (formId == "frmRecordEdit" || formId == "frmRecordEditForm") floatingWindow = true;
			}
			if (!floatingWindow) return window.top.document.querySelector('#' + frameId + ' #' + formId);

			//floating window -work out where we are.
			var activeIframeId = 'iframe_' + OpenHR.activeWindowID().replace('mwid_', '');
			if (activeIframeId == "iframe_") return null;	// no valid form found.

			return window.top.document.getElementById(activeIframeId).contentWindow.document.getElementById(formId);
		},

		postData = function (url, jsonData, followOnFunctionName) {
			OpenHR.ResetSession(); //Reset the session so it doesn't timeout

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
		OpenHR.ResetSession(); //Reset the session so it doesn't timeout

		var $frame;
		$.ajax({
			url: url,
			type: "POST",
			data: jsonData,
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
				$frame.dialog({ position: 'center' });
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


	submitForm = function (form, targetWin, asyncFlag, jsonData, action, followOnFunctionName, floatWindow, reuseWindow) {
		OpenHR.ResetSession(); //Reset the session so it doesn't timeout

		var $form = $(form),
			$frame = $form.closest("div[data-framesource]").first(),
			target = $form.attr("target"),
			method = $form.attr("method");

		var data;
		var url;

		if (action == undefined) {
			url = $form.attr("action");
		} else {
			url = action;
		}


		if (jsonData == undefined) {
			data = $form.serialize();
		} else {
			data = jsonData;
			method = "POST";  // bit trigger happy this, maybe some of the the controller actions should be gets???

			//	var globalToken = $('#__AjaxAntiForgeryForm');
			//		var token = $('input[name="__RequestVerificationToken"]', globalToken).val();
			//		data.push(token);
		}


		if ((asyncFlag == undefined) || (asyncFlag.length == 0) || (asyncFlag == true)) {
			asyncFlag = true;
		} else {
			asyncFlag = false;
		}

		OpenHR.ResetSession(); //Reset the session so it doesn't timeout

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
				if (!floatWindow || floatWindow === false) {
					if (targetWin != null) {

						//$frame = $form.closest("div[" + targetWin + "]").first();
						if (targetWin.toLowerCase() == "workframe") $frame = $("#" + targetWin);
						else $frame = window.top.$("#" + targetWin);

						if ($("#" + targetWin).hasClass("reportoutput") === true && asyncFlag === true) {
							$frame.html('');
							$(".popup").dialog("open");
							$(".popup").dialog({ dialogClass: 'no-close' });
						}

					}

					$frame.html('');
					$frame.html(html);

					//jQuery styling
					$(function () {
						$frame.find("input[type=submit], input[type=button], button").button();
						$frame.find("input").addClass("ui-widget ui-corner-all");
						$frame.find("input").removeClass("text");

						$frame.find("textarea").addClass("ui-widget ui-corner-tl ui-corner-bl");
						$frame.find("textarea").removeClass("text");

						$frame.find("select").addClass("ui-widget ui-corner-tl ui-corner-bl");
						$frame.find("select").removeClass("text");
						$frame.find("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
						$frame.find("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");
					});
				} else {
					var multiwindowid;
					if (reuseWindow) {
						//reuse the active window (reloading after data change)
						multiwindowid = Number(OpenHR.activeWindowID().replace("mwid_", ""));
					} else {
						if (window.top.window.multiwindowid) {
							multiwindowid = window.top.window.multiwindowid + 1;
						} else {
							multiwindowid = 1;
						}
						window.top.window.multiwindowid = multiwindowid;
						window.top.$('#workframeset').append('<div id="mwid_' + multiwindowid + '"><iframe frameborder="0" style="width: 100%; height: 100%;padding: 0;" id="iframe_' + multiwindowid + '"/></div>');
					}

					//set parent form id. 					
					var parentFormId = OpenHR.activeWindowID();
					//if this parent has a parent use it.
					var parentsParentFormId = $('#' + parentFormId).attr('data-parentformid');
					if (parentsParentFormId) {
						if (parentsParentFormId !== "") parentFormId = parentsParentFormId;
					}

					var thisMwId = 'mwid_' + multiwindowid;

					var newHeight = screen.height / 2;
					var newWidth = screen.width / 2;
					var position = {};
					
					//Calculate size of new dialog
					//The first find window is full size, all other windows are half screen size.
					if (OpenHR.listOpenWindows().length === 0) {
						var frmWorkArea = OpenHR.getForm("workframeset", "frmWorkAreaRefresh");
						if (frmWorkArea.txtGotoPage.value.toLowerCase() === "find") {
							newHeight = $('#workframeset').outerHeight() - 10;
							newWidth = $('#workframeset').width() - 10;
							position = { my: "left top", at: "left top", of: "#workframeset" };
						}
					}

					if (!reuseWindow) {
						window.top.$('#mwid_' + multiwindowid).dialog({
							dialogClass: "mwid_" + multiwindowid,
							width: newWidth,
							height: newHeight,
							position: position,
							beforeClose: function (event, ui) {
								var hasChanged = menu_saveChanges('', true, false);
								if (hasChanged === 0) {
									// Prompt for navigation and redirect as required
									saveChangesPrompt("floatingWindow", 'OpenHR.closeDialog("' + thisMwId + '")');
									return false;
								} else {
									//close all dependent dialogs
									$('[id^="mwid_"]').each(function () {
										var parentFormId = $(this).attr('data-parentformid');
										if (parentFormId === thisMwId) OpenHR.closeDialog($(this).attr('id'));
									});
									//close this dialog.
									OpenHR.closeDialog(thisMwId);
								}
							},
							open: function() {
								//set all existing dialogs to 'inactive'
								window.top.$('[id^="mwid_"]').each(function() {
									$(this).siblings(".ui-dialog-titlebar").addClass('ui-state-disabled');
								});
								//set this new screen to 'active'
								window.top.$('#' + thisMwId).siblings(".ui-dialog-titlebar").removeClass('ui-state-disabled');
							},
							close: function () {								
								try {
									$(this).dialog('destroy').remove();
								}
								catch (e) { }
							}
						});


						//Prevent floating window from leaving the workarea
						var container = $('#workframeset'),
							dialog = $('#mwid_' + multiwindowid).parent();
						dialog.draggable({ "containment": container});

					}

					//todo: using document.write isn't good. It also kills the ajax spinner. May need to find a jquery way of populating the iframe...					
					//todo: is this populating the iframe twice when reusewindow = true?
					var iframe = window.top.document.getElementById('iframe_' + multiwindowid);
					iframe = (iframe.contentWindow) ? iframe.contentWindow : (iframe.contentDocument.document) ? iframe.contentDocument.document : iframe.contentDocument;
					iframe.document.open();
					//Populate the iframe with the AJAX response
					iframe.document.write("<div id='workframe' name='mwid_" + multiwindowid + "' class='absolutefull'>" +
						html +
						"<script>window.top.OpenHR.setupMwIframe(" + multiwindowid + ", $('#workframe').attr('data-framesource'), $('.pageTitle').text());<\/script>" +
						"</div>");
					iframe.document.close();

					window.top.$('#iframe_' + multiwindowid).load(function () {						
						//NPG: I thought this was fired when the dialogs are completely loaded. But for recedit at least it isn't. The search continues...

						//get screen type and assign parent id - parentFormId
						if (OpenHR.getScreenType('mwid_' + multiwindowid) === "history") {
							$(this).parent().attr('data-parentFormID', parentFormId);
						} else {
							$(this).parent().attr('data-parentFormID', "");
						}
						
						OpenHR.updateDialogPageTitle();

						//remove 'please wait' spinner
						window.top.$('body').removeClass('loading');
					});
				}

				if (typeof followOnFunctionName !== "undefined" && followOnFunctionName !== null) {
					followOnFunctionName();
				}

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
			return window.top.window.LocaleDecimalSeparator;
		},
		localeThousandSeparator = function () {
			return window.top.window.LocaleThousandSeparator;
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

			if (OpenHR.activeWindowID() !== "") {
				//Multi window display active
				var iFrame = OpenHR.activeWindowID().replace('mwid', 'iframe');
				sCurrentPage = $("#" + iFrame).contents().find('#workframe').attr("data-framesource");
			}

			if (!($("#workframe").css('display') == 'none')) {
				//Work frame is in view.
				sCurrentPage = $("#workframe").attr("data-framesource");
			}
			else if (!($("#optionframe").css('display') == 'none')) {
				//Option frame is in view.
				sCurrentPage = $("#optionframe").attr("data-framesource");
			}

				//ToolsFrame frame is in view.
			else if (!($("#ToolsFrame").css('display') == 'none')) {
				sCurrentPage = $("#ToolsFrame").attr("data-framesource");
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

			var localeFormat = OpenHR.LocaleDateFormat();

			if (Date.parseExact(d, localeFormat) == null) {
				return false;
			}
			return true;
		},

	localeDateFormat = function () {
		var formats = {
			"AF-ZA": "yyyy/MM/dd",
			"AM-ET": "d/M/yyyy",
			"AR-AE": "dd/MM/yyyy",
			"AR-BH": "dd/MM/yyyy",
			"AR-DZ": "dd-MM-yyyy",
			"AR-EG": "dd/MM/yyyy",
			"AR-IQ": "dd/MM/yyyy",
			"AR-JO": "dd/MM/yyyy",
			"AR-KW": "dd/MM/yyyy",
			"AR-LB": "dd/MM/yyyy",
			"AR-LY": "dd/MM/yyyy",
			"AR-MA": "dd-MM-yyyy",
			"ARN-CL": "dd-MM-yyyy",
			"AR-OM": "dd/MM/yyyy",
			"AR-QA": "dd/MM/yyyy",
			"AR-SA": "dd/MM/yy",
			"AR-SY": "dd/MM/yyyy",
			"AR-TN": "dd-MM-yyyy",
			"AR-YE": "dd/MM/yyyy",
			"AS-IN": "dd-MM-yyyy",
			"AZ-CYRL-AZ": "dd.MM.yyyy",
			"AZ-LATN-AZ": "dd.MM.yyyy",
			"BA-RU": "dd.MM.yy",
			"BE-BY": "dd.MM.yyyy",
			"BG-BG": "dd.M.yyyy",
			"BN-BD": "dd-MM-yy",
			"BN-IN": "dd-MM-yy",
			"BO-CN": "yyyy/M/d",
			"BR-FR": "dd/MM/yyyy",
			"BS-CYRL-BA": "d.M.yyyy",
			"BS-LATN-BA": "d.M.yyyy",
			"CA-ES": "dd/MM/yyyy",
			"CO-FR": "dd/MM/yyyy",
			"CS-CZ": "d.M.yyyy",
			"CY-GB": "dd/MM/yyyy",
			"DA-DK": "dd-MM-yyyy",
			"DE-AT": "dd.MM.yyyy",
			"DE-CH": "dd.MM.yyyy",
			"DE-DE": "dd.MM.yyyy",
			"DE-LI": "dd.MM.yyyy",
			"DE-LU": "dd.MM.yyyy",
			"DSB-DE": "d. M. yyyy",
			"DV-MV": "dd/MM/yy",
			"EL-GR": "d/M/yyyy",
			"EN": "MM/dd/yyyy",
			"EN-029": "MM/dd/yyyy",
			"EN-AU": "d/MM/yyyy",
			"EN-BZ": "dd/MM/yyyy",
			"EN-CA": "dd/MM/yyyy",
			"EN-GB": "dd/MM/yyyy",
			"EN-IE": "dd/MM/yyyy",
			"EN-IN": "dd-MM-yyyy",
			"EN-JM": "dd/MM/yyyy",
			"EN-MY": "d/M/yyyy",
			"EN-NZ": "d/MM/yyyy",
			"EN-PH": "M/d/yyyy",
			"EN-SG": "d/M/yyyy",
			"EN-TT": "dd/MM/yyyy",
			"EN-US": "MM/dd/yyyy",
			"EN-ZA": "yyyy/MM/dd",
			"EN-ZW": "M/d/yyyy",
			"ES-AR": "dd/MM/yyyy",
			"ES-BO": "dd/MM/yyyy",
			"ES-CL": "dd-MM-yyyy",
			"ES-CO": "dd/MM/yyyy",
			"ES-CR": "dd/MM/yyyy",
			"ES-DO": "dd/MM/yyyy",
			"ES-EC": "dd/MM/yyyy",
			"ES-ES": "dd/MM/yyyy",
			"ES-GT": "dd/MM/yyyy",
			"ES-HN": "dd/MM/yyyy",
			"ES-MX": "dd/MM/yyyy",
			"ES-NI": "dd/MM/yyyy",
			"ES-PA": "MM/dd/yyyy",
			"ES-PE": "dd/MM/yyyy",
			"ES-PR": "dd/MM/yyyy",
			"ES-PY": "dd/MM/yyyy",
			"ES-SV": "dd/MM/yyyy",
			"ES-US": "M/d/yyyy",
			"ES-UY": "dd/MM/yyyy",
			"ES-VE": "dd/MM/yyyy",
			"ET-EE": "d.MM.yyyy",
			"EU-ES": "yyyy/MM/dd",
			"FA-IR": "MM/dd/yyyy",
			"FI-FI": "d.M.yyyy",
			"FIL-PH": "M/d/yyyy",
			"FO-FO": "dd-MM-yyyy",
			"FR-BE": "d/MM/yyyy",
			"FR-CA": "yyyy-MM-dd",
			"FR-CH": "dd.MM.yyyy",
			"FR-FR": "dd/MM/yyyy",
			"FR-LU": "dd/MM/yyyy",
			"FR-MC": "dd/MM/yyyy",
			"FY-NL": "d-M-yyyy",
			"GA-IE": "dd/MM/yyyy",
			"GD-GB": "dd/MM/yyyy",
			"GL-ES": "dd/MM/yy",
			"GSW-FR": "dd/MM/yyyy",
			"GU-IN": "dd-MM-yy",
			"HA-LATN-NG": "d/M/yyyy",
			"HE-IL": "dd/MM/yyyy",
			"HI-IN": "dd-MM-yyyy",
			"HR-BA": "d.M.yyyy.",
			"HR-HR": "d.M.yyyy",
			"HSB-DE": "d. M. yyyy",
			"HU-HU": "yyyy. MM. dd.",
			"HY-AM": "dd.MM.yyyy",
			"ID-ID": "dd/MM/yyyy",
			"IG-NG": "d/M/yyyy",
			"II-CN": "yyyy/M/d",
			"IS-IS": "d.M.yyyy",
			"IT-CH": "dd.MM.yyyy",
			"IT-IT": "dd/MM/yyyy",
			"IU-CANS-CA": "d/M/yyyy",
			"IU-LATN-CA": "d/MM/yyyy",
			"JA-JP": "yyyy/MM/dd",
			"KA-GE": "dd.MM.yyyy",
			"KK-KZ": "dd.MM.yyyy",
			"KL-GL": "dd-MM-yyyy",
			"KM-KH": "yyyy-MM-dd",
			"KN-IN": "dd-MM-yy",
			"KOK-IN": "dd-MM-yyyy",
			"KO-KR": "yyyy-MM-dd",
			"KY-KG": "dd.MM.yy",
			"LB-LU": "dd/MM/yyyy",
			"LO-LA": "dd/MM/yyyy",
			"LT-LT": "yyyy.MM.dd",
			"LV-LV": "yyyy.MM.dd.",
			"MI-NZ": "dd/MM/yyyy",
			"MK-MK": "dd.MM.yyyy",
			"ML-IN": "dd-MM-yy",
			"MN-MN": "yy.MM.dd",
			"MN-MONG-CN": "yyyy/M/d",
			"MOH-CA": "M/d/yyyy",
			"MR-IN": "dd-MM-yyyy",
			"MS-BN": "dd/MM/yyyy",
			"MS-MY": "dd/MM/yyyy",
			"MT-MT": "dd/MM/yyyy",
			"NB-NO": "dd.MM.yyyy",
			"NE-NP": "M/d/yyyy",
			"NL-BE": "d/MM/yyyy",
			"NL-NL": "d-M-yyyy",
			"NN-NO": "dd.MM.yyyy",
			"NSO-ZA": "yyyy/MM/dd",
			"OC-FR": "dd/MM/yyyy",
			"OR-IN": "dd-MM-yy",
			"PA-IN": "dd-MM-yy",
			"PL-PL": "yyyy-MM-dd",
			"PRS-AF": "dd/MM/yy",
			"PS-AF": "dd/MM/yy",
			"PT-BR": "d/M/yyyy",
			"PT-PT": "dd-MM-yyyy",
			"QUT-GT": "dd/MM/yyyy",
			"QUZ-BO": "dd/MM/yyyy",
			"QUZ-EC": "dd/MM/yyyy",
			"QUZ-PE": "dd/MM/yyyy",
			"RM-CH": "dd/MM/yyyy",
			"RO-RO": "dd.MM.yyyy",
			"RU-RU": "dd.MM.yyyy",
			"RW-RW": "M/d/yyyy",
			"SAH-RU": "MM.dd.yyyy",
			"SA-IN": "dd-MM-yyyy",
			"SE-FI": "d.M.yyyy",
			"SE-NO": "dd.MM.yyyy",
			"SE-SE": "yyyy-MM-dd",
			"SI-LK": "yyyy-MM-dd",
			"SK-SK": "d. M. yyyy",
			"SL-SI": "d.M.yyyy",
			"SMA-NO": "dd.MM.yyyy",
			"SMA-SE": "yyyy-MM-dd",
			"SMJ-NO": "dd.MM.yyyy",
			"SMJ-SE": "yyyy-MM-dd",
			"SMN-FI": "d.M.yyyy",
			"SMS-FI": "d.M.yyyy",
			"SQ-AL": "yyyy-MM-dd",
			"SR-CYRL-BA": "d.M.yyyy",
			"SR-CYRL-CS": "d.M.yyyy",
			"SR-CYRL-ME": "d.M.yyyy",
			"SR-CYRL-RS": "d.M.yyyy",
			"SR-LATN-BA": "d.M.yyyy",
			"SR-LATN-CS": "d.M.yyyy",
			"SR-Latn-ME": "d.M.yyyy",
			"SR-LATN-RS": "d.M.yyyy",
			"SV-FI": "d.M.yyyy",
			"SV-SE": "yyyy-MM-dd",
			"SW-KE": "M/d/yyyy",
			"SYR-SY": "dd/MM/yyyy",
			"TA-IN": "dd-MM-yyyy",
			"TE-IN": "dd-MM-yy",
			"TG-CYRL-TJ": "dd.MM.yy",
			"TH-TH": "d/M/yyyy",
			"TK-TM": "dd.MM.yy",
			"TN-ZA": "yyyy/MM/dd",
			"TR-TR": "dd.MM.yyyy",
			"TT-RU": "dd.MM.yyyy",
			"TZM-LATN-DZ": "dd-MM-yyyy",
			"UG-CN": "yyyy-M-d",
			"UK-UA": "dd.MM.yyyy",
			"UR-PK": "dd/MM/yyyy",
			"UZ-CYRL-UZ": "dd.MM.yyyy",
			"UZ-LATN-UZ": "dd/MM yyyy",
			"VI-VN": "dd/MM/yyyy",
			"WO-SN": "dd/MM/yyyy",
			"XH-ZA": "yyyy/MM/dd",
			"YO-NG": "d/M/yyyy",
			"ZH-CN": "yyyy/M/d",
			"ZH-HK": "d/M/yyyy",
			"ZH-MO": "d/M/yyyy",
			"ZH-SG": "d/M/yyyy",
			"ZH-TW": "yyyy/M/d",
			"ZU-ZA": "yyyy/MM/dd"
		};

		var language = window.top.window.UserLocale || window.opener.window.UserLocale;
		return formats[language.toUpperCase()] || 'dd/MM/yyyy';

	},

		convertSqlDateToLocale = function (z) {

			var convertDate = Date.parseExact(z, "M/d/yyyy");
			if (convertDate != null) {
				return convertDate.toString(window.top.window.LocaleDateFormat);
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
			if (psDateString.toString().trim() == "") return "null";

			if (!isValidDate(psDateString)) {
				return psDateString;
			}

			sDateFormat = window.top.window.LocaleDateFormat || window.dialogArguments.window.LocaleDateFormat;

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
	nullsafeInteger = function (arg) {
		var returnvalue = 0; //default to 0
		if ((arg == undefined) || (arg == 0) || arg.length <= 0) {
			return returnvalue;
		} else {
			try {
				returnvalue = Number(arg);
			} catch (e) {
				return returnvalue;
			}
		}
		return returnvalue;
	},
	sessionTimeout = function () {

		$("#SignalRDialogTitle").html("You are about to be logged out");
		$("#SignalRDialogContentText").html("Your browser has been inactive for a while, so for your security<BR/>you will be automatically logged off your OpenHR session.");

		$("#divSignalRMessage").dialog({
			dialogClass: 'alert',
			closeText: 'Log Out',
			close: function () {
				window.top.onbeforeunload = null; try { window.location.href = "Main"; } catch (e) { } return false;
			}
		});

		$("#divSignalRMessage").dialog('open');

		$("#SignalRDialogClick").off('click').on('click', function () {
			$("#divSignalRMessage").dialog('close');
		});

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
		if (inIframe) return window.top.window.LocaleDateFormat.replace("dd", "d").replace("MM", "m").replace("M", "m").replace("yyyy", "Y");
		return window.top.window.LocaleDateFormat.replace("dd", "d").replace("MM", "m").replace("M", "m").replace("yyyy", "Y");
	},

	parentExists = function ()
	{
		//function to detect if this form is displayed in a dialog or a jquery modal div.
		//true = window.dialog.		
		var opener = window.dialogArguments || window.opener;
		return (opener == null)?false:true;
	},

	windowOpen = function (destination, width, height) {
		// From https://developer.mozilla.org/en-US/docs/Web/API/Window.open:
		// "only list the features to be enabled or rendered; the others (except titlebar and close) will be disabled or removed."
		var windowProperties = "centerscreen,chrome," + "height=" + height + "," + "width=" + width;

		return window.open(destination, "_blank", windowProperties);
	},

	isChrome = function() {
		// please note, that IE11 now returns undefined again for window.chrome
		var isChromium = window.chrome,
				vendorName = window.navigator.vendor;
		if (isChromium !== null && isChromium !== undefined && vendorName === "Google Inc.") {
			// is Google chrome 
			return true;
		} else {
			// not Google chrome 
			return false;
		}
	},
	clearTmpDialog = function () {
		try {
			if ($('#tmpDialog').dialog('isOpen') == true) {
				$('#tmpDialog').dialog('close');
				$('#tmpDialog').dialog('destroy');
				$('#tmpDialog').remove();
			}
		}
		catch (e) { }
	},
	gridSelectTopRow = function (grid) {
		var ids = grid.getDataIDs();
		grid.jqGrid("setSelection", ids[0], true);
	},
	gridSelectLastRow = function (grid) {
		grid.jqGrid('resetSelection');
		var rowCount = grid.getGridParam("reccount");
		var ids = grid.getDataIDs();
		grid.jqGrid("setSelection", ids[rowCount - 1], true);
	},
	gridPageDown = function (grid) {
		//skips 18 rows at a time
		var rowid = grid.jqGrid('getGridParam', 'selrow');
		var rowNumber = grid.jqGrid('getInd', rowid) - 1; // zero based 
		var rowCount = grid.getGridParam("reccount");
		var ids = grid.getDataIDs();
		if ((rowNumber + 18) >= rowCount) { gridSelectLastRow(grid); }
		else { grid.jqGrid("setSelection", ids[rowNumber + 18], true); }
	},
	gridPageUp = function (grid) {
		//skips 18 rows at a time
		var rowid = grid.jqGrid('getGridParam', 'selrow');
		var rowNumber = grid.jqGrid('getInd', rowid) - 1; // zero based 
		var ids = grid.getDataIDs();
		if ((rowNumber - 18) < 0) { gridSelectTopRow(grid); }
		else { grid.jqGrid("setSelection", ids[rowNumber - 18], true); }
	},
	gridKeyboardEvent = function(keyPressed, grid) {

		if ((keyPressed != 40) && (keyPressed != 38) && (keyPressed != 13) && (keyPressed != 32) && (keyPressed != 33) && (keyPressed != 34) && (keyPressed != 35) && (keyPressed != 36)) {
			//Character search
			try {
				var gridID = $(grid).attr('id');
				var id = $('#' + gridID + ' td:visible').filter(function () {
					return $(this).text().substring(0, 1).toLowerCase() == String.fromCharCode(keyPressed).toLowerCase();
				}).first().closest('tr').attr('id');
				if (Number(id) > 0)
					grid.jqGrid('resetSelection');
				grid.jqGrid('setSelection', id);
			}
			catch (e) { }
		}
		else {

			//Get the current rowId
			if ((grid.getGridParam("records") > 0) && (grid.jqGrid('getGridParam', 'selrow') != null)) {

				var rowid = grid.jqGrid('getGridParam', 'selrow');

				//Get the current row number
				var rowNumber = grid.jqGrid('getInd', rowid) - 1; // zero based 
				var rowCount = grid.getGridParam("reccount");
				var ids = grid.getDataIDs();
				grid.jqGrid('resetSelection');

				//up arrow, down arrow, Enter, spacebar, home, end, pgup and pgdn.
				switch (keyPressed) {
					case 40:
						//Down arrow
						if ((rowNumber + 1) == rowCount) { OpenHR.gridSelectLastRow(grid); }
						else { grid.jqGrid("setSelection", ids[rowNumber + 1], true); }
						break;
					case 38:
						//Up arrow
						if (rowNumber == 0) { OpenHR.gridSelectTopRow(grid); }
						else { grid.jqGrid("setSelection", ids[rowNumber - 1], true); }
						break;
					case 33:
						//Page Up
						OpenHR.gridPageUp(grid);
						break;
					case 34:
						//Page Down
						OpenHR.gridPageDown(grid);
						break;
					case 35:
						//End
						OpenHR.gridSelectLastRow(grid);
						break;
					case 36:
						//Home
						OpenHR.gridSelectTopRow(grid);
						break;

				}
			}
			else { alert('nothing selected'); }

		}

	},

	// Check invalid characters 
	checkInvalidCharacters = function (input) {
		return !openhrBlackListValidator.IsBlackListPattern(input, new RegExp($ESAPI.properties.openHRValidationBlackList.AllInvalidCharacters));
	},

	// Validate integer value 
	validateInteger = function (input) {
		return openhrBlackListValidator.IsValidIntegerValue(input);
	},

	// Validate numeric value 
	validateNumeric = function (input) {
		return openhrBlackListValidator.IsValidNumericValue(input);
	},

	displayServerMessage = function (messageFrom, message, forceLogout, loggedInUsersOnly) {

		var isLoggedIn = ($("#frmLoginForm").length === 0);
		if (loggedInUsersOnly && isLoggedIn) {

			$("#SignalRDialogClick").val("Close");
			$("#SignalRDialogTitle").html(messageFrom);
			$("#SignalRDialogContentText").html(message);

			$("#divSignalRMessage").dialog({
				closeOnEscape: !forceLogout
			});

			$("#divSignalRMessage").dialog("open");

			if (forceLogout === true) {
				$("#SignalRDialogClick").val("Log Out");
			}

			$("#SignalRDialogClick").off("click").on("click", function () {
				$("#divSignalRMessage").dialog("close");

				if (forceLogout === true) {
					menu_logoffIntranet();
				}

			});

		}
	},

	showAboutPopup = function () {
		var aboutUrl = window.top.window.ROOT + "/account/about";
		if (window.top.window.ROOT.slice(-1) == "/") aboutUrl = window.top.window.ROOT + "account/about";

		OpenHR.ResetSession(); //Reset the session so it doesn't timeout

		$.ajax({
			url: aboutUrl,
			dataType: 'html',
			cache: false,
			success: function (html) {
				$('#About').html(html);

				$("#About input[type=submit], input[type=button], button").button();
				$("#About input").addClass("ui-widget ui-corner-all");
				$("#About input").removeClass("text");

				$("#About textarea").addClass("ui-widget ui-corner-tl ui-corner-bl");
				$("#About textarea").removeClass("text");

				$("#About select").addClass("ui-widget ui-corner-tl ui-corner-bl");
				$("#About select").removeClass("text");
				$("#About input[type=submit], input[type=button], button").removeClass("ui-corner-all");
				$("#About input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");

				$("#About").dialog("open");
			},
			error: function () {

			}
		});
	},

	isValidFileExtension = function (filename) {
		var validFileExtensions = window.top.$("#globals>#ValidFileExtensions").val().toLowerCase().split(",");

		var fileExtension = OpenHR.GetFileExtension(filename).toLocaleLowerCase();
		if (validFileExtensions.indexOf(fileExtension) == -1) return false;
		return true;

	},

	inIframe = function () {
		try {
			return window.self !== window.top;
		} catch (e) {
			return false;
		}
	},

	setupMwIframe = function (iFrameId, framesource, pageTitle) {		
		//move the page title to the modal dialog title
		$('#mwid_' + iFrameId).dialog('option', 'title', pageTitle);
		$('#iframe_' + iFrameId).contents().find('.pageTitleDiv').hide();

		//show the close button for these dialogs
		$('#mwid_' + iFrameId).siblings(".ui-widget-header").find("button.ui-dialog-titlebar-close").show();

		//set the parent workframe data-framesource to match the dialog workframe
		$('#workframe').attr('data-framesource', framesource);	//todo: is this a good idea??
		$('#workframe').attr('data-mwid', 'mwid_' + iFrameId);

		//clear the parent workframe
		$('#workframe').html('');

		var $frame = OpenHR.activeFrame().find("#workframe");
		//jQuery styling
		$(function () {
			$frame.find("input[type=submit], input[type=button], button").button();
			$frame.find("input").addClass("ui-widget ui-corner-all");
			$frame.find("input").removeClass("text");

			$frame.find("textarea").addClass("ui-widget ui-corner-tl ui-corner-bl");
			$frame.find("textarea").removeClass("text");

			$frame.find("select").addClass("ui-widget ui-corner-tl ui-corner-bl");
			$frame.find("select").removeClass("text");
			$frame.find("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
			$frame.find("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");
		});


		//Add a click event to the dialog		
		$('#mwid_' + iFrameId).parent('.ui-dialog, .ui-dialog-content').on('mousedown', function (events) {
			window.top.window.isDragging = false;
		})
		.on('mousemove', function () {
			window.top.window.isDragging = true;
		})
		.on('mouseup', function () {
			var wasDragging = window.top.window.isDragging;
			window.top.window.isDragging = false;
			if (!wasDragging) {
				//check for changes!
				var activeWindowNumber = Number(OpenHR.activeWindowID().replace('mwid_', ''));
				var selectedWindowNumber = Number(iFrameId);
				
				if (activeWindowNumber !== selectedWindowNumber) {
					//is dialog clicked not the active one?				
					var hasChanged = menu_saveChanges('', true, false);
					if (hasChanged === 0) {
						// Prompt for navigation and redirect as required
						saveChangesPrompt("floatingWindow", 'OpenHR.activateDialog(' + iFrameId + '")'); //todo: bring window to top if navigate cancelled.
					} else {
						//activate clicked dialog.
						OpenHR.activateDialog(iFrameId);
					}
				}

			}
		});
	
	},

	activateDialog = function (iFrameId) {
		$('#mwid_' + iFrameId).dialog('moveToTop');

		//make this dialog 'active' and any others 'inactive'
		$('[id^="mwid_"]').each(function () {
			$(this).siblings(".ui-dialog-titlebar").addClass('ui-state-disabled');
		});

		$('#mwid_' + iFrameId).siblings(".ui-dialog-titlebar").removeClass('ui-state-disabled');

		//Now the selected frame is active.
		var framesource = OpenHR.activeFrame().find('#workframe').attr('data-framesource');
		$('#workframe').attr('data-framesource', framesource);
		$('#workframe').attr('data-mwid', 'mwid_' + iFrameId);

		//if this is recedit, clone the frmdata back
		if (OpenHR.currentWorkPage() === "RECORDEDIT" && !menu_isSSIMode()) {
			OpenHR.activeFrame().find("#frmData").children().each(function(element) {
				window.top.$('#frmData').find('#' + this.id).val(this.value);
			});
		}

		OpenHR.setWorkFrameDialogsVisible(true);

		menu_refreshMenu();		

	},


	listOpenWindows = function () {
		var idList = [];

		window.top.$('div[role="dialog"][class*="mwid_"]').each(function () {
			var active = (!$(this).find('.ui-dialog-titlebar').hasClass('ui-state-disabled'));
			var classList = $(this).attr('class').split(' ');
			var screenId = 0;
			var viewId = 0;						

			try {
				screenId = $(this).find('iframe').contents().find('#workframe').find('#txtCurrentScreenID').val();
				viewId = $(this).find('iframe').contents().find('#workframe').find('#txtCurrentViewID').val();	//todo: should this be tableID?				
			}
			catch (e) { }

			var idvar;
			var windowNumber;
			$.each(classList, function (i, v) {
				if (v.match(/^mwid_/)) {
					idvar = v;
					windowNumber = Number(v.replace('mwid_', ''));
				}
			});

			var screenType = OpenHR.getScreenType('mwid_' + windowNumber);

			idList.push({ windowNumber: windowNumber, id: idvar, active: active, screenid: screenId, viewid: viewId, screentype: screenType });

		});

		return idList;
	},

	activeWindowID = function () {
		var windowList = listOpenWindows();
		var result = windowList.filter(function (item) {
			return (item.active === true);
		});

		if (result) {
			if (result.length > 0) return result[0].id;
		}
		return '';
	},

	activeFrame = function () {
		var activeIframeId = OpenHR.activeWindowID().replace('mwid', 'iframe');
		if (activeIframeId !== "") return window.top.$('#' + activeIframeId).contents();
		return $('body');
	},

	activeDialog = function () {
		var activeDialogId = OpenHR.activeWindowID();
		if (activeDialogId !== "") return window.top.$('#' + activeDialogId);
		return $('body');
	},

	closeDialog = function (dialogId) {
		window.top.$('#workframe').removeAttr('data-mwid');
		window.top.$('#workframe').removeAttr('data-framesource');
		window.top.$('#' + dialogId.replace("mwid", "iframe")).remove();	//clear contents before destroying dialog
		showDefaultRibbon();
		try {
			$("#" + dialogId).dialog('destroy').remove();
		}
		catch (e) { }
		//display the top window in the list.
		var windowList = OpenHR.listOpenWindows();
		if (windowList.length > 0) {
			var windowNumber = OpenHR.listOpenWindows()[0].windowNumber;
			if (windowNumber) OpenHR.activateDialog(windowNumber);
		}
	},

	populateSwitchWindows = function () {
		var openWindows = listOpenWindows();

		window.top.$('.buttonsplitmenu>ul').empty();

		openWindows.forEach(function (item) {
			var iFrameNumber = item.id.replace('mwid_', '');
			var pageTitle = OpenHR.getIframePageTitle(iFrameNumber);

			window.top.$('.buttonsplitmenu>ul').append('<li data-frameid="' +
				iFrameNumber + '"><a onclick="OpenHR.activateDialog(' +
				iFrameNumber + ');"><span><img style="width: 16px; height: 16px; display: ' + (item.active ? 'inline-block;' : 'none;') + '" src="' +
				window.top.window.ROOT + '/Content/images/greenLightIcon.png" alt=""/>' +
				pageTitle + '</span></a></li>');

		});

	},

	getIframePageTitle = function (iFrameNumber) {
		return window.top.$('#iframe_' + iFrameNumber).contents().find('.pageTitle').text();
	},

	setWorkFrameDialogsVisible = function(visibility) {
		window.top.$('[id^="mwid_"]').parent().toggle(visibility);
	},

	activeIFrameID = function () {
		var windowList = listOpenWindows();
		var result = windowList.filter(function (item) {
			return (item.active === true);
		});

		if (result) {
			if (result.length > 0) return result[0].id.replace('mwid_', 'iframe_');
		}
		return '';
	},

	getScreenType = function (mwid) {
		var screenType = "";
		try {
			var parentTableId = Number($('#' + mwid).find('iframe').contents().find('#workframe').find('#txtCurrentParentTableID').val());
			if (parentTableId === 0) screenType = "parent";
			else if (parentTableId > 0) screenType = "history";
		}
		catch (e) { return "" }

		return screenType;
	},

	updateDialogPageTitle = function () {
		
		var caption;
		if (!menu_isSSIMode()) {

			var parentFormId = OpenHR.activeDialog().attr("data-parentformid");
			var originalPageTitle = OpenHR.activeFrame().find('#txtOriginalPageTitle').val();

			if (OpenHR.activeFrame().find('#frmFindForm').length > 0) {
				//find window
				if (parentFormId === "") {
					OpenHR.activeFrame().find('.pageTitle').html("Find - " + originalPageTitle);
					OpenHR.activeDialog().dialog("option", "title", "Find - " + originalPageTitle + (caption ? ' - ' + caption : ""));
				} else {
					var parentPageTitle = window.top.$('#' + parentFormId).find('iframe').contents().find('.pageTitle').text();
					OpenHR.activeFrame().find('.pageTitle').html(originalPageTitle + ' (' + parentPageTitle + ')');
					OpenHR.activeDialog().dialog("option", "title", originalPageTitle + ' (' + parentPageTitle + ')');
				}
			} else {
				//recedit
				var frmData = OpenHR.getForm("dataframe", "frmData");
				var frmRecEdit = OpenHR.getForm("workframe", "frmRecordEditForm");

				if (frmRecEdit.txtCurrentRecordID.value > 0) {
					if (frmData.txtRecordDescription.value.length > 0) {
						caption = frmData.txtRecordDescription.value;
					}
				} else {
					caption = "New Record";
				}

				if (parentFormId === "") {
					OpenHR.activeFrame().find('.pageTitle').html(originalPageTitle + (caption ? ' - ' + caption : ""));
					OpenHR.activeDialog().dialog("option", "title", originalPageTitle + (caption ? ' - ' + caption : ""));
				} else {
					var parentPageTitle = window.top.$('#' + parentFormId).find('iframe').contents().find('.pageTitle').text();
					OpenHR.activeFrame().find('.pageTitle').html(originalPageTitle + ' (' + parentPageTitle + ')');
					OpenHR.activeDialog().dialog("option", "title", originalPageTitle + ' (' + parentPageTitle + ')');
				}
			}


		}
	},

	resetSession = function () { //Some variables used in this function are global and declared in Site.Master
		$.post('RefreshSession', { __RequestVerificationToken: window.top.$('[name="__RequestVerificationToken"]').val() }, function () { });

		window.top.timeoutSecondsLeft = window.top.originalTimeoutSeconds;

		if (!window.top.decrementFunctionIsSetup) {
			var countdownTimer = setInterval(function() {
					window.top.currentMinutes = Math.floor(window.top.timeoutSecondsLeft / 60);
					window.top.currentSeconds = window.top.timeoutSecondsLeft % 60;
					if (window.top.currentSeconds <= 9) window.top.currentSeconds = "0" + window.top.currentSeconds;
					window.top.timeoutSecondsLeft--;

					try {
						if (window.top.timeoutSecondsLeft < 300)
							$("#sessionWarning").show(); //show countdown for the last 5 minutes.
						else
							$("#sessionWarning").hide();

						if (window.top.timeoutSecondsLeft === 0) {
							clearInterval(countdownTimer);
							document.getElementById("timerText").innerHTML = "0:00";
						}

						document.getElementById("timerText").innerHTML = window.top.currentMinutes + ":" + window.top.currentSeconds; //Set the element id you need the time put into.
					} catch (e) {
						//do nothing if this fails - we've probably navigated away and the elements no longer exist. That's the trouble with using 1 second delays.
					}
				}, 1000
			);

			window.top.decrementFunctionIsSetup = true;
		}
	}

	window.OpenHR = {
		version: version,
		messageBox: messageBox,
		modalPrompt: modalPrompt,
		modalMessage: modalMessage,
		getFrame: getFrame,
		getForm: getForm,
		postData: postData,
		submitForm: submitForm,
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
		nullsafeInteger: nullsafeInteger,
		replaceAll: replaceAll,
		getLocaleDateString: getLocaleDateString,
		setDatepickerLanguage: setDatepickerLanguage,
		IsValidDate: isValidDate,
		MoveItemInGrid: moveItemInGrid,
		OpenDialog: openDialog,
		modalExpressionSelect: modalExpressionSelect,
		parentExists: parentExists,
		windowOpen: windowOpen,
		isChrome: isChrome,
		clearTmpDialog: clearTmpDialog,
		gridSelectTopRow: gridSelectTopRow,
		gridSelectLastRow: gridSelectLastRow,
		gridPageDown: gridPageDown,
		gridPageUp: gridPageUp,
		gridKeyboardEvent: gridKeyboardEvent,
		showAboutPopup: showAboutPopup,
		checkInvalidCharacters: checkInvalidCharacters,
		validateInteger: validateInteger,
		validateNumeric: validateNumeric,
		displayServerMessage: displayServerMessage,
		isValidFileExtension: isValidFileExtension,
		inIframe: inIframe,
		setupMwIframe: setupMwIframe,
		listOpenWindows: listOpenWindows,
		activeWindowID: activeWindowID,
		activeFrame: activeFrame,
		activateDialog: activateDialog,
		closeDialog: closeDialog,
		populateSwitchWindows: populateSwitchWindows,
		getIframePageTitle: getIframePageTitle,
		setWorkFrameDialogsVisible: setWorkFrameDialogsVisible,
		activeIFrameID: activeIFrameID,
		getScreenType: getScreenType,
		activeDialog: activeDialog,
		updateDialogPageTitle: updateDialogPageTitle,
		ResetSession: resetSession
	};

})(window, jQuery);