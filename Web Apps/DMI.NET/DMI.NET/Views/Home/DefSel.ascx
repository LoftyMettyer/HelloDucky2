<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Extensions" %>

<%="" %>

<%
	Dim SelectedTableID As String = Request.Form("SelectedTableID")
	Dim fGotId As Boolean
	Dim iBaseTableID As Integer

	Dim iDefSelType = CType(Session("defseltype"), UtilityType)

	Dim objSession As SessionInfo = CType(Session("sessionContext"), SessionInfo)
	
	Session("objCalendar" & Session("UtilID")) = Nothing

	If Session("fromMenu") = 0 Then
		If Session("singleRecordID") < 1 Then
			If Not String.IsNullOrEmpty(Request.Form("txtTableID")) Then
				iBaseTableID = Request.Form("txtTableID")
			Else

				If Len(Session("tableID")) > 0 Then
					If CLng(Session("tableID")) > 0 Then
						iBaseTableID = Session("tableID")
						fGotId = True
					End If
				End If

				If fGotId = False Then
					If (Session("singleRecordID") > 0) Then
						iBaseTableID = SettingsConfig.Personnel_EmpTableID
					End If
				End If
			End If
		Else
			If Len(Session("tableID")) > 0 Then
				iBaseTableID = Session("tableID")
			End If
		End If
	End If
	
	If Session("singleRecordID") = 0 Then
		If Session("optionTableID") > 0 Then
			iBaseTableID = Session("optionTableID")
		End If
		Session("tableID") = Session("utilTableID")
	End If
	
	Session("optionTableID") = 0
	
	If iDefSelType = UtilityType.utlPicklist Or iDefSelType = UtilityType.utlFilter Or iDefSelType = UtilityType.utlCalculation Then
		iBaseTableID = CInt(Session("utilTableID"))
	End If
%>

<script type="text/javascript">

	var defSelType = "<%:iDefSelType%>";
	var menuSection = "Report";
	if ((defSelType === "utlMailMerge") || (defSelType === "utlWorkflow")) menuSection = "Utilities";
	if ((defSelType === "utlPicklist") || (defSelType === "utlFilter") || (defSelType === "utlCalculation")) menuSection = "Tools";


		function ssOleDBGridDefSelRecords_dblClick() {

				var frmDefSel = document.getElementById("frmDefSel");

				if ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12)) {
					menu_MenuClick('mnutoolEditToolsFind');

				}
				else {
						// DblClick triggers Run after prompting for confirmation. 
						if (frmDefSel.cmdRun.disabled == true) {
								return (false);
						}

						var answer = 0;
					var sCaption = "";

					if (frmDefSel.utiltype.value == 1) {
						sCaption = "Cross Tab";
						}

						if (frmDefSel.utiltype.value == 2) {
							sCaption = "Custom Report";
						}
						if (frmDefSel.utiltype.value == 9) {
							sCaption = "Mail Merge";
						}
						if (frmDefSel.utiltype.value == 17) {
							sCaption = "Calendar Report";
						}
						if (frmDefSel.utiltype.value == 25) {
							sCaption = " Workflow";
						}
						if (frmDefSel.utiltype.value == 35) {
							sCaption = "9-Box Grid Report";
						}

					var sMessage = "Are you sure you want to run " + $("#utilname").text() + " ?";
					OpenHR.modalPrompt(sMessage, 4, sCaption, setrun);

				}
				return false;
		}

		function ssOleDBGridDefSelRecords_rowcolchange() {			
			var rowId = $("#DefSelRecords").getGridParam("selrow");
			var gridData = $("#DefSelRecords").getRowData(rowId);
			var username = $("#frmDefSel #txtusername").val();

			$("#txtDescription").val(gridData.description);

			// Populate the hidden fields with the selected utils information       
			$("#UtilID").val($("#DefSelRecords").getGridParam("selrow"));
			$("#UtilName").val(gridData.Name);

			if (gridData.Username !== username) {
				if (!allowEdit()) {
					menu_SetmnutoolButtonCaption("mnutoolEdit" + menuSection + "Find", "View");
					menu_toolbarEnableItem("mnutoolDelete" + menuSection + "Find", false);
				} else {
					menu_SetmnutoolButtonCaption("mnutoolEdit" + menuSection + "Find", "Edit");

					if ($("#DeleteGranted").val() === "True") {
						menu_toolbarEnableItem("mnutoolDelete" + menuSection + "Find", true);
					} else {
						menu_toolbarEnableItem("mnutoolDelete" + menuSection + "Find", false);
					}
				}
			} else {
				menu_SetmnutoolButtonCaption("mnutoolEdit" + menuSection + "Find", "Edit");

				if ($("#DeleteGranted").val() === "True") {
					menu_toolbarEnableItem("mnutoolDelete" + menuSection + "Find", true);
				} else {
					menu_toolbarEnableItem("mnutoolDelete" + menuSection + "Find", false);
				}
			}
			refreshControls();

		}
	
		function defsel_window_onload() {

				var frmDefSel = document.getElementById('frmDefSel');

				// Expand the option frame and hide the work frame.
				if (parseInt($("#txtSingleRecordID").val()) > 0) {
						$("#optionframe").attr("data-framesource", "DEFSEL");
						$("#workframe").hide();
						$("#optionframe").show();
				} else {
						$("#workframe").attr("data-framesource", "DEFSEL");
						$("#optionframe").hide();
						$("#workframe").show();
				}


				$("#DefSelRecords").jqGrid('bindKeys', {
					"onEnter": function (rowid) {
						ssOleDBGridDefSelRecords_dblClick();
					}
				});

				refreshControls();

			// Navbar options = i.e. search, edit, save etc 
			$("#DefSelRecords").jqGrid('navGrid', '#pager-coldata', { del: false, add: false, edit: false, search: false, refresh: false }); // setup the buttons we want
			$("#DefSelRecords").jqGrid('filterToolbar', {stringResult: true, searchOnEnter: false	});  //instantiate toolbar so we can use toggle.
			
			if ($('#pager-coldata :has(".ui-icon-search")').length == 0) {
				$("#DefSelRecords").jqGrid('navButtonAdd', "#pager-coldata", {
					caption: '',
					buttonicon: 'ui-icon-search',
					position: 'first',
					onClickButton: function () {
						this.clearToolbar();
						this.toggleToolbar();
						if ($('.ui-search-toolbar', this.grid.hDiv).is(':visible'))
						{
							$('.ui-search-toolbar', this.grid.fhDiv).show();
						} else {
							$('.ui-search-toolbar', this.grid.fhDiv).hide();
						}},
					title: 'Search',
					cursor: 'pointer'
				});
				$('.ui-search-toolbar').hide(); // Hide it on setting up the grid - NB Remove this line to have it open on setup
			}

				$("#findGridRow").height("60%");
				$(window).bind('resize', function () {
					$("#DefSelRecords").setGridWidth($('#findGridRow').width(), true);
				}).trigger('resize');

				$("#DefSelRecords").closest('.ui-jqgrid-bdiv').width($("#DefSelRecords").closest('.ui-jqgrid-bdiv').width() + 1);

				frmDefSel.cmdCancel.focus();

				if (rowCount() > 0) {

					var isSingleRecord = (parseInt($("#txtSingleRecordID").val()) <= 0);
						var gotoID;

						if (isSingleRecord === true) {
								gotoID = $("#lastSelectedID")[0].value;
								if (Number(gotoID) == 0) gotoID = $("#DefSelRecords").getDataIDs()[0];
						} else {
								gotoID = $("#DefSelRecords").getDataIDs()[0];
						}
						$("#DefSelRecords").jqGrid("setSelection", gotoID);

						// If no row is selected then select first row
						if ($("#DefSelRecords").getGridParam('selrow') == null) {
								$("#DefSelRecords").jqGrid("setSelection", $("#DefSelRecords").getDataIDs()[0]);
						}

				} else {
					 //If the table is empty disable Copy, Edit, Delete and Properties buttons
						menu_toolbarEnableItem("mnutoolCopyToolsFind", false);
						menu_toolbarEnableItem("mnutoolEditToolsFind", false);
						menu_toolbarEnableItem("mnutoolDeleteToolsFind", false);
						menu_toolbarEnableItem("mnutoolPropertiesToolsFind", false);
				}
		}

		function rowCount() {
			return $("#DefSelRecords").jqGrid('getGridParam', 'records');
		}

		function disableNonDefselTabs() {
			$("#toolbarRecordFind").parent().hide();
			$("#toolbarRecord").parent().hide();
			$("#toolbarRecordAbsence").parent().hide();
			$("#toolbarRecordQuickFind").parent().hide();
			$("#toolbarRecordSortOrder").parent().hide();
			$("#toolbarRecordFilter").parent().hide();
			$("#toolbarRecordMailMerge").parent().hide();
			$("#toolbarReportNewEditCopy").parent().hide();
			$("#toolbarReportRun").parent().hide();
			$("#toolbarUtilitiesNewEditCopy").parent().hide();
			$("#toolbarEventLogView").parent().hide();
			$("#toolbarAdminConfig").parent().hide();
		}

		function refreshControls() {
			
			//show the Defsel-Find menu block.

			disableNonDefselTabs();

			var fFromMenu = (parseInt($("#txtSingleRecordID").val()) <= 0);
			var fHasRows = (rowCount() > 0);
			var isNewPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "NEW").ToString.ToLower%>");
			var isEditPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "EDIT").ToString.ToLower%>");
			var isViewPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "VIEW").ToString.ToLower%>");
			var isDeletePermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "DELETE").ToString.ToLower%>");
			var isRunPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "RUN").ToString.ToLower%>");

			var isWorkflow = (defSelType === "utlWorkflow");

			$("#toolbarUtilitiesFind").parent().hide();
			$("#toolbarToolsFind").parent().hide();
			$("#toolbarEventLogFind").parent().hide();
			$("#toolbarWFPendingStepsFind").parent().hide();
			$("#toolbarReportFind").parent().hide();

			menu_setVisibleMenuItem("mnutoolNew" + menuSection + "Find", !isWorkflow);
			menu_setVisibleMenuItem("mnutoolCopy" + menuSection + "Find", !isWorkflow);
			menu_setVisibleMenuItem("mnutoolEdit" + menuSection + "Find", !isWorkflow);
			menu_setVisibleMenuItem("mnutoolDelete" + menuSection + "Find", !isWorkflow);
			menu_setVisibleMenuItem("mnutoolProperties" + menuSection + "Find", !isWorkflow);
			menu_setVisibleMenuItem("mnutoolRun" + menuSection + "Find", (menuSection !== "Tools"));
			menu_setVisibleMenuItem("mnutoolClose" + menuSection + "Find", !fFromMenu);

			if (!isWorkflow) {
				menu_toolbarEnableItem("mnutoolNew" + menuSection + "Find", isNewPermitted && fFromMenu);
				menu_toolbarEnableItem("mnutoolCopy" + menuSection + "Find", fHasRows && isNewPermitted && fFromMenu);
				menu_toolbarEnableItem("mnutoolEdit" + menuSection + "Find", fHasRows && (isEditPermitted || isViewPermitted) && fFromMenu);
				menu_toolbarEnableItem("mnutoolProperties" + menuSection + "Find", fHasRows && fFromMenu);
				if (menuSection !== "Tools") menu_toolbarEnableItem("mnutoolRun" + menuSection + "Find", fHasRows && isRunPermitted && fFromMenu);
				if (menuSection !== "Tools") menu_toolbarEnableItem("mnutoolClose" + menuSection + "Find", !fFromMenu);
				if (defSelType === "17") menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && isRunPermitted); //Calendar Reports
			} else {
				menu_toolbarEnableItem("mnutoolRunUtilitiesFind", fFromMenu);
			}

			// Finally show and select the tab
			$("#toolbar" + menuSection + "Find").parent().show();
			$("#toolbar" + menuSection + "Find").click();

			// If delete permission is given for the report but the 'Read Only' permission has been given in Group Access then disable the delete button
			if (fHasRows && isDeletePermitted && fFromMenu) {
				DisableDeleteButtonIfDefinitionHasReadOnlyAccess("mnutoolDeleteReportFind");
			}

		}


	// If the selected record has Read Only permission given in the Group Access then disable the delete button
	function DisableDeleteButtonIfDefinitionHasReadOnlyAccess(menuItem) {
		var rowId = $("#DefSelRecords").getGridParam('selrow');
		if (rowId != null) {
			var gridData = $("#DefSelRecords").getRowData(rowId);
			if (gridData.Access == 'ro') {
				menu_toolbarEnableItem(menuItem, false);
			}
		}
	}

	function showproperties() {

			if (!$("#mnutoolPropertiesUtil").hasClass("disabled")) {
				var id = $("#DefSelRecords").getGridParam('selrow');
				var type = $("#utiltype").val();
				var name = $("#utilname").text();
				OpenHR.OpenDialog("DefinitionProperties", "divPopupReportDefinition", { ID: id, Type: type, Name: name }, '900px');

			}
		}

		function pausecomp(millis) {
				var date = new Date();
				var curDate;

				do {
						curDate = new Date();
				} while (curDate - date < millis);
		}

		function NewWindow(mypage, myname, w, h, scroll) {
				var winl = (screen.width - w) / 2;
				var wint = (screen.height - h) / 2;
				var winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';
				var win = window.open(mypage, myname, winprops);

				if (parseInt(navigator.appVersion) >= 4) {
						// Delay fixes a problem with IE7 and Vista (don't know why though!)
						pausecomp(300);
						win.window.focus();
				}
		}

		function ReturnNewWindow(mypage, myname, w, h, scroll) {
				var winl = (screen.width - w) / 2;
				var wint = (screen.height - h) / 2;
				var winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';
				var win = window.open(mypage, myname, winprops);

				if (parseInt(navigator.appVersion) >= 4) {
						// Delay fixes a problem with IE7 and Vista (don't know why though!)
						pausecomp(300);
						win.window.focus();
				}

				return win;

		}

		function ToggleCheck() {

			var piTableID = 0;
			var frmDefSel = document.getElementById('frmDefSel');

			if ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12)) {
				piTableID = frmDefSel.selectTable.options[frmDefSel.selectTable.selectedIndex].value;
			}

			// Load the required definition selection screen
			var displayDiv = (parseInt($("#txtSingleRecordID").val()) === 0 ? "workframe" : "optionframe");
			var postData = {
				txtTableID: piTableID,
				utiltype: frmDefSel.utiltype.value,
				OnlyMine: $("#OnlyMine").prop('checked'),
				RecordID: parseInt($("#txtSingleRecordID").val()),
				__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
			};

			OpenHR.submitForm(null, displayDiv, null, postData, "DefSel");

		}

		function setdelete() {
				if (!$("#mnutoolDeleteUtil").hasClass("disabled")) {
					var frmDefSel = document.getElementById('frmDefSel');

					OpenHR.modalPrompt("Delete '" + $("#utilname").text() + "'. Are you sure ?", 4, "Confirm").then(function (answer) {
						if (answer === 6) { 

							var postData = {
								Action: "delete",
								utiltype: frmDefSel.utiltype.value,
								utilID: frmDefSel.utilid.value,
								utilName: $("#utilname").text(),
								__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
							}

									OpenHR.submitForm(null, "divPopupReportDefinition", null, postData, "defsel_submit");
							//debugger;
							//OpenHR.OpenDialog("defsel_submit", "divPopupReportDefinition", postData, '900px');


						}
					});

				}
		}

		function setrun(answer) {

			if (answer === 6) {

				var postData;

				if (!$("#mnutillRunUtil").hasClass("disabled")) {
					var frmDefSel = document.getElementById('frmDefSel');

					if (frmDefSel.utiltype.value == 25) {
						// Workflow
						postData = {
							ID: frmDefSel.utilid.value,
							Name: $("#utilname").text(),
							__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
						}

						OpenHR.submitForm(null, "optionframe", null, postData, "util_run_workflow");

					} else {

						postData = {
							UtilType: frmDefSel.utiltype.value,
							ID: frmDefSel.utilid.value,
							Name: $("#utilname").text(),
							__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
						}
						OpenHR.submitForm(null, "reportframe", false, postData, "util_run_promptedValues");

					}
				}
			}
		}

		function setnew() {
			if (!$("#mnutoolNewUtil").hasClass("disabled")) {
				var frmDefSel = document.getElementById('frmDefSel');
				frmDefSel.action.value = "new";
				OpenHR.submitForm(frmDefSel);
			}
		}

		function setcopy() {
				if (!$("#mnutoolCopyUtil").hasClass("disabled")) {
						var frmDefSel = document.getElementById('frmDefSel');

						frmDefSel.action.value = "copy";
						OpenHR.submitForm(frmDefSel);
				}
		}

		function setedit() {

				if (!$("#mnutoolEditUtil").hasClass("disabled")) {
						var frmDefSel = document.getElementById('frmDefSel');

						if (frmDefSel.cmdEdit.value == "Edit") {
								frmDefSel.action.value = "edit";
								OpenHR.submitForm(frmDefSel);
						} else {
								frmDefSel.action.value = "view";
								OpenHR.submitForm(frmDefSel);
						}
				}
		}

		function setcancel() {

			if (parseInt($("#txtSingleRecordID").val()) > 0) {
				menu_disableMenu();

				$("#optionframe").hide();
				$("#workframe").show();
				$("#toolbarRecord").show();
				$("#toolbarRecord").click();

				menu_refreshMenu();
			}
		}


		function loadEmptyOption() {
				$.ajax({
						url: 'emptyoption',
						type: "POST",
						dataType: 'html',
						async: true,
						success: function (html) {
								try {
										$('#optionframe').html('');
										$('#optionframe').html(html);
								} catch (e) { }
						}
				});
		}



		function defsel_currentWorkFramePage() {
			var sCurrentPage = $("#workframe").attr("data-framesource");
			try
			{
				sCurrentPage = sCurrentPage.toUpperCase();
			} catch (e) { }

			return sCurrentPage;
		}

	 
</script>

<div id="defsel" data-framesource="defsel" style="display: block; height:100%; width: 99.9%">

		<form name="frmDefSel" class="absolutefull" action="defsel_submit" method="post" id="frmDefSel">
<div id="findGridRow" style="height: 70%; margin-right: 20px; margin-left: 20px;">

										<table width="100%" height="100%" class="invisible">
												<tr>
														<td colspan="5" height="10">
																<span class="pageTitle">
																		<%
																			If iDefSelType = UtilityType.utlCrossTab Then
																				Response.Write("Cross Tabs")
																			ElseIf iDefSelType = UtilityType.utlCustomReport Then
																				Response.Write("Custom Reports")
																			ElseIf iDefSelType = UtilityType.utlMailMerge Then
																				Response.Write("Mail Merge")
																			ElseIf iDefSelType = UtilityType.utlPicklist Then
																				Response.Write("Picklists")
																			ElseIf iDefSelType = UtilityType.utlFilter Then
																				Response.Write("Filters")
																			ElseIf iDefSelType = UtilityType.utlCalculation Then
																				Response.Write("Calculations")
																			ElseIf iDefSelType = UtilityType.utlCalendarReport Then
																				Response.Write("Calendar Reports")
																			ElseIf iDefSelType = UtilityType.utlWorkflow Then
																				Response.Write("Workflow")
																			ElseIf iDefSelType = UtilityType.utlNineBoxGrid Then
																				Response.Write("9-Box Grid Reports")
																			End If
																		%>
																</span>
														</td>
												</tr>

												<% 
														Dim sErrorDescription = ""
	
													If iDefSelType = UtilityType.utlPicklist Or iDefSelType = UtilityType.utlFilter Or iDefSelType = UtilityType.utlCalculation Then
												%>
												<tr height="10">

														<td height="10" colspan="3">
																<table width="100%" class="invisible">
																		<tr>
																				<td style="width: 44px;">Table :
																				</td>
																				<td width="10">&nbsp;
																				</td>
																				<td width="175">
																						<select id="selectTable" name="selectTable" class="combo" style="height: 22px; width: 200px" >
																								<%
	
																									Try

																										For Each objTable In objSession.Tables.OrderBy(Function(t) t.Name) 'Order by table name
																												
																											Response.Write("						<option value=" & objTable.ID)
																											If SelectedTableID Is Nothing Or SelectedTableID = "" Then
																												If objTable.ID = iBaseTableID Then
																													Response.Write(" SELECTED")
																												End If
																											Else
																												If objTable.ID = CLng(SelectedTableID) Then
																													Response.Write(" SELECTED")
																												End If
																											End If

																											Response.Write(">" & Replace(objTable.Name, "_", " ") & "</option>" & vbCrLf)

																										Next
				
																									Catch ex As Exception
																										sErrorDescription = "The table records could not be retrieved." & vbCrLf & ex.Message

																									End Try
																								%>
																						</select>
																				</td>

																				<td>&nbsp;
																				</td>
																		</tr>
																</table>
														</td>
												</tr>
												<tr>
														<td colspan="5" height="10"></td>
												</tr>
												<%
												End If
												%>

												<tr>

														<td width="100%">
																<table height="100%" width="100%">
																		<tr>
																				<td width="100%">																				
																						<table id="DefSelRecords"></table>
																						<div id='pager-coldata'></div>
																				</td>
																		</tr>

																		<tr height="10">
																				<td></td>
																		</tr>

																		<tr>
																				<td height="70">
																						<textarea cols="20" class="disabled" style="WIDTH: 100%;" name="txtDescription" rows="4"  tabindex="-1" disabled="disabled" >
									</textarea>
																				</td>
																		</tr>
																</table>
														</td>

														<td width="80" style="display: none;">
																<table height="100%" class="invisible">
																		<tr>
																				<td>
																						<input type="button" id="cmdNew" class="btn" name="cmdNew" value="New" style="width: 80px"
																								<% 
																							If (Session("singleRecordID") > 0) Or iDefSelType = UtilityType.utlWorkflow Then
																								Response.Write(" style=""visibility:hidden""")
																							End If
%>
																								onclick="setnew();" />
																				</td>
																		</tr>
																		<tr height="10">
																				<td></td>
																		</tr>
																		<tr>
																				<td>
																						<input type="button" name="cmdEdit" class="btn" value="Edit" style="width: 80px"
																								<% 
																							If (Session("singleRecordID") > 0) Or iDefSelType = UtilityType.utlWorkflow Then
																								Response.Write(" style=""visibility:hidden""")
																							End If
%>
																								onclick="setedit();" />
																				</td>
																		</tr>
																		<tr height="10">
																				<td></td>
																		</tr>
																		<tr>
																				<td>
																						<input type="button" name="cmdCopy" class="btn" id="cmdCopy" value="Copy" style="width: 80px"
																								<% 
																							If (Session("singleRecordID") > 0) Or iDefSelType = UtilityType.utlWorkflow Then
																								Response.Write(" style=""visibility:hidden""")
																							End If
%>
																								onclick="setcopy();" />
																				</td>
																		</tr>
																		<tr height="10">
																				<td></td>
																		</tr>
																		<tr>
																				<td>
																						<input type="button" name="cmdDelete" class="btn" value="Delete" style="width: 80px"
																								<% 
																							If (Session("singleRecordID") > 0) Or iDefSelType = UtilityType.utlWorkflow Then
																								Response.Write(" style=""visibility:hidden""")
																							End If
%>
																								onclick="setdelete();" />
																				</td>
																		</tr>
																		<tr height="10">
																				<td></td>
																		</tr>
																		<tr>
																				<td>
																						<input type="button" name="cmdPrint" class="btn btndisabled" value="Print" style="width: 80px" disabled
																								<% 

																							If (Session("singleRecordID") > 0) Or iDefSelType = UtilityType.utlWorkflow Then
																								Response.Write(" style=""visibility:hidden""")
																							End If
%> />
																				</td>
																		</tr>
																		<tr height="10">
																				<td></td>
																		</tr>
																		<tr>
																				<td>
																						<input type="button" name="cmdProperties" class="btn" value="Properties" style="width: 80px"
																								<% 
																							If (Session("singleRecordID") > 0) Or iDefSelType = UtilityType.utlWorkflow Then
																								Response.Write(" style=""visibility:hidden""")
																							End If
%>
																								onclick="showproperties();" />
																				</td>
																		</tr>
																		<tr height="10">
																				<td></td>
																		</tr>
																		<tr height="100%">
																				<td></td>
																		</tr>
																		<tr>
																				<td>
																						<input type="button" name="cmdRun" class="btn" value="Run" style="width: 80px" id="cmdRun"
																								<% 																						
																							If iDefSelType = UtilityType.utlPicklist Or iDefSelType = UtilityType.utlFilter Or iDefSelType = UtilityType.utlCalculation Then
																								Response.Write(" style=""visibility:hidden""")
																							End If
%>
																								onclick="setrun(6);" />
																				</td>
																		</tr>
																		<tr height="10">
																				<td></td>
																		</tr>
																	<tr>
																		<td>
																			<input type="button" name="cmdCancel" class="btn" value='<% 
																				If iDefSelType = UtilityType.utlPicklist Or iDefSelType = UtilityType.utlFilter Or iDefSelType = UtilityType.utlCalculation Then
																					Response.Write("""OK""")
																				Else
																					Response.Write("""Cancel""")
																				End If
%>'
																				style="width: 80px"
																				onclick="setcancel()" />
																		</td>
																	</tr>
																</table>
														</td>
														<td width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>
												</tr>

											<tr>
												<td colspan="5" height="10"
													<%
													If iDefSelType = UtilityType.utlWorkflow Then
														Response.Write(" style=""visibility:hidden""")
															 End If%>>
													<input type='hidden' id="txtusername" name="txtusername" value="<%=lcase(session("Username"))%>">
												</td>
											</tr>

												<tr>
														<td colspan="4" height="10"
																<%
															If iDefSelType = UtilityType.utlWorkflow Then
																Response.Write(" style=""visibility:hidden""")
															End If
%>>
																<input  <% If Session("OnlyMine") Then Response.Write("checked")%>  type="checkbox" tabindex="0" id="OnlyMine" onclick="ToggleCheck();" />
																<label for="OnlyMine" class="checkbox" tabindex="-1" onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}">
																		Only show definitions where owner is '<%:Session("Username")%>'
																</label>
														</td>
												</tr>
										</table>

				<input type="hidden" id="utiltype" name="utiltype" value="<%:CInt(iDefSelType)%>">
				<input type="hidden" id="utilid" name="utilid" value='<%:Session("utilid")%>'>
				<input type="hidden" id="utilname" name="utilname">
				<input type="hidden" id="action" name="action">
				<input type="hidden" id="txtTableID" name="txtTableID" value='<%=iBaseTableID%>'>

</div>
			<%=Html.AntiForgeryToken()%>
		</form>


	<input type="hidden" id="txtSingleRecordID" name="txtSingleRecordID" value='<%:session("singleRecordID")%>'>
	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">
	
	<input type="hidden" id="lastSelectedID" name="lastSelectedID" value='<%=Session("utilid")%>'>

</div>


<script>
	$("#DefSelRecords").keydown(function (event) {		
		//Add first letter search to the grid...
		try {
			var id = $('#DefSelRecords td:visible').filter(function () {
				return $(this).text().substring(0, 1).toLowerCase() == String.fromCharCode(event.which).toLowerCase();
			}).first().closest('tr').attr('id');
			if (Number(id) > 0)
				$("#DefSelRecords").jqGrid('setSelection', id);
		}
		catch(e) {}
	});

	function attachDefSelGrid() {
		var onlyMine = $("#OnlyMine").prop('checked');
		
		//resize grid		
		var gridWidth = $("#findGridRow").width();
		var gridHeight = $("#workframeset").height() * 0.6;	//findGridRow is hardcoded to 60% of workframeset.

		$("#DefSelRecords").jqGrid({
			url: 'GetDefinitionsForType?UtilityType=' + <%:CInt(iDefSelType)%> + '&&TableID=' + <%=iBaseTableID%> + '&&OnlyMine=' + onlyMine,
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
			colNames: ['ID', 'Name', 'description', 'Username', 'Access' ],
			colModel: [
				{ name: 'ID', index: 'ID', hidden: true },
				{ name: 'Name', index: 'Name', width: 40, sortable: false },
				{ name: 'description', index: 'description', hidden: true },
				{ name: 'Username', index: 'Username', hidden: true },
				{ name: 'Access', index: 'Access', hidden: true }],
			viewrecords: false,
			width: gridWidth,
			height: gridHeight,
			sortname: 'Name',
			sortorder: "asc",
			rowNum: 10000,
			cmTemplate: { sortable: false },
			ignoreCase: true,
			onSelectRow: function (rowID) {
				ssOleDBGridDefSelRecords_rowcolchange();
			},
			ondblClickRow: function (rowID) {
				ssOleDBGridDefSelRecords_dblClick();
			},
			loadComplete: function(json) {		
				defsel_window_onload();
			},
			rowTotal: 50,
			rowList: [],
			pager: $('#pager-coldata'),
			pgbuttons: false,
			pgtext: null,
			loadonce: true,
			autoencode: true,
			loadui: "disable"
		});

	}

	$(function () {
		attachDefSelGrid();

		$("#selectTable").change(function () {
			$('#SelectedTableID').val(($('#selectTable').val()));
			ToggleCheck();
		});
	});

	function allowEdit() {
		var rowId = $("#DefSelRecords").getGridParam("selrow");
		var gridData = $("#DefSelRecords").getRowData(rowId);
		if (gridData.Access === "ro") {
			return false;
		}
		return true;
	}
</script>

