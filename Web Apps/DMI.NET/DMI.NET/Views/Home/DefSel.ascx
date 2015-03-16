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
		If CStr(Session("optionTableID")) <> "" Then
			If Session("optionTableID") > 0 Then
				iBaseTableID = Session("optionTableID")
			End If
		End If
		Session("tableID") = Session("utilTableID")
	End If
	
	'Session("optionDefSelType") = ""
	Session("optionTableID") = ""
	
	If iDefSelType = UtilityType.utlPicklist Or iDefSelType = UtilityType.utlFilter Or iDefSelType = UtilityType.utlCalculation Then
		iBaseTableID = CInt(Session("utilTableID"))
	End If
%>

<script type="text/javascript">

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

						if (frmDefSel.utiltype.value == 1) {
								answer = OpenHR.messageBox("Are you sure you want to run the '" + $.trim(frmDefSel.utilname.value) + "' Cross Tab ?", 36, "Confirmation...");
						}

						if (frmDefSel.utiltype.value == 2) {
								answer = OpenHR.messageBox("Are you sure you want to run the '" + $.trim(frmDefSel.utilname.value) + "' Custom Report ?", 36, "Confirmation...");
						}
						if (frmDefSel.utiltype.value == 9) {
								answer = OpenHR.messageBox("Are you sure you want to run the '" + $.trim(frmDefSel.utilname.value) + "' Mail Merge ?", 36, "Confirmation...");
						}
						if (frmDefSel.utiltype.value == 17) {
								answer = OpenHR.messageBox("Are you sure you want to run the '" + $.trim(frmDefSel.utilname.value) + "' Calendar Report ?", 36, "Confirmation...");
						}
						if (frmDefSel.utiltype.value == 25) {
								answer = OpenHR.messageBox("Are you sure you want to run the '" + $.trim(frmDefSel.utilname.value) + "' Workflow ?", 36, "Confirmation...");
						}
						if (frmDefSel.utiltype.value == 35) {
							answer = OpenHR.messageBox("Are you sure you want to run the '" + $.trim(frmDefSel.utilname.value) + "' 9-Box Grid Report ?", 36, "Confirmation...");
						}

						if (answer == 6) {
								setrun();
						}
				}
				return false;
		}

		function ssOleDBGridDefSelRecords_rowcolchange() {
			var frmDefSel = document.getElementById("frmDefSel");

			var rowId = $("#DefSelRecords").getGridParam('selrow');
			var gridData = $("#DefSelRecords").getRowData(rowId);

			frmDefSel.txtDescription.value = gridData.description;

			// Populate the hidden fields with the selected utils information       
			frmDefSel.utilid.value = $("#DefSelRecords").getGridParam('selrow');
			frmDefSel.utilname.value = gridData.Name;

			var IsNewPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "NEW").ToString.ToLower%>");
			var IsEditPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "EDIT").ToString.ToLower%>");
			var IsRunPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "RUN").ToString.ToLower%>");
			var IsDeletePermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "DELETE").ToString.ToLower%>");

			button_disable(frmDefSel.cmdRun, !IsRunPermitted);
			button_disable(frmDefSel.cmdNew, !IsNewPermitted);
			button_disable(frmDefSel.cmdCopy, !IsNewPermitted);
			button_disable(frmDefSel.cmdEdit, !IsEditPermitted);

			if (gridData.Username != frmDefSel.txtusername.value) {
				if (gridData.Access == 'ro') {
					frmDefSel.cmdEdit.value = 'View';
					menu_SetmnutoolButtonCaption("mnutoolEditToolsFind", "View");
					button_disable(frmDefSel.cmdDelete, true);
					menu_toolbarEnableItem("mnutoolDeleteToolsFind", false);
				} else {
					frmDefSel.cmdEdit.value = 'Edit';
					menu_SetmnutoolButtonCaption("mnutoolEditToolsFind", "Edit");

					if (IsDeletePermitted) {
						button_disable(frmDefSel.cmdDelete, false);
						menu_toolbarEnableItem("mnutoolDeleteToolsFind", true);
					} else {
						button_disable(frmDefSel.cmdDelete, true);
						menu_toolbarEnableItem("mnutoolDeleteToolsFind", false);
					}
				}
			} else {
				frmDefSel.cmdEdit.value = 'Edit';
				menu_SetmnutoolButtonCaption("mnutoolEditToolsFind", "Edit");

				if (IsDeletePermitted) {
					button_disable(frmDefSel.cmdDelete, false);
					menu_toolbarEnableItem("mnutoolDeleteToolsFind", true);
				} else {
					button_disable(frmDefSel.cmdDelete, true);
					menu_toolbarEnableItem("mnutoolDeleteToolsFind", false);
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
					$("#DefSelRecords").setGridHeight($("#findGridRow").height(), true);
				}).trigger('resize');

				$("#DefSelRecords").setGridHeight($("#findGridRow").height());
				$("#DefSelRecords").setGridWidth($("#findGridRow").width());

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
				//$("#toolbarReportFind").hide();
				$("#toolbarReportNewEditCopy").parent().hide();
				$("#toolbarReportRun").parent().hide();
				//$("#toolbarUtilitiesFind").hide();
				$("#toolbarUtilitiesNewEditCopy").parent().hide();
				//$("#toolbarToolsFind").hide();
				//$("#toolbarEventLogFind").hide();
				$("#toolbarEventLogView").parent().hide();
				//$("#toolbarWFPendingStepsFind").hide();
				$("#toolbarAdminConfig").parent().hide();
		}

		function refreshControls() {

				//show the Defsel-Find menu block.
				//$("#mnuSectionUtilities").show();
				frmDefSel = document.getElementById('frmDefSel');			

				disableNonDefselTabs();

				//reset utilities tab
				menu_setVisibleMenuItem("mnutoolNewUtilitiesFind", true);
				menu_setVisibleMenuItem("mnutoolCopyUtilitiesFind", true);
				menu_setVisibleMenuItem("mnutoolEditUtilitiesFind", true);
				menu_setVisibleMenuItem("mnutoolDeleteUtilitiesFind", true);
				menu_setVisibleMenuItem("mnutoolPropertiesUtilitiesFind", true);
				menu_setVisibleMenuItem("mnutoolRunUtilitiesFind", true);
				var isSingleRecord = (parseInt($("#txtSingleRecordID").val()) <= 0);
				var fHasRows = (rowCount() > 0);

			var IsNewPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "NEW").ToString.ToLower%>");
			var IsEditPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "EDIT").ToString.ToLower%>");
			var IsViewPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "VIEW").ToString.ToLower%>");
			var IsDeletePermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "DELETE").ToString.ToLower%>");
			var IsRunPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "RUN").ToString.ToLower%>");

			switch ('<%:CInt(Session("defseltype"))%>') {
				case '0':  // "BatchJobs"
						break;
				case '1':  // "CrossTabs"
						// Hide the remaining tabs
						$("#toolbarUtilitiesFind").parent().hide();
						$("#toolbarToolsFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_toolbarEnableItem("mnutoolNewReportFind", IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolNewReportFind", true);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows && IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows && (IsEditPermitted || IsViewPermitted) && isSingleRecord);
						menu_SetmnutoolButtonCaption("mnutoolEditReportFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditReportFind", true);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows && IsDeletePermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);
						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && IsRunPermitted && isSingleRecord);
						// Only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseReportFind', !isSingleRecord);
						menu_toolbarEnableItem('mnutoolCloseReportFind', !isSingleRecord);
						// Show and select the tab
						$("#toolbarReportFind").parent().show();
						$("#toolbarReportFind").click();
						break;
				case '2':  // "CustomReports"

						// Hide the remaining tabs
						$("#toolbarUtilitiesFind").parent().hide();
						$("#toolbarToolsFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_toolbarEnableItem("mnutoolNewReportFind", IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolNewReportFind", true);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows && IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows && (IsEditPermitted || IsViewPermitted) && isSingleRecord);
						menu_SetmnutoolButtonCaption("mnutoolEditReportFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditReportFind", true);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows && IsDeletePermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);
						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && IsRunPermitted && isSingleRecord);
						// Only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseReportFind', !isSingleRecord);
						menu_toolbarEnableItem('mnutoolCloseReportFind', !isSingleRecord);
						// Show and select the tab
						$("#toolbarReportFind").parent().show();
						$("#toolbarReportFind").click();
						break;
				case '3':  //sTemp = sTemp & "DataTransfer"
						break;
				case '4':  //sTemp = sTemp & "Export"
						break;
				case '5':  //sTemp = sTemp & "GlobalAdd"
						break;
				case '6':  //sTemp = sTemp & "GlobalDelete"
						break;
				case '7':  //sTemp = sTemp & "GlobalUpdate"
						break;
				case '8':  //sTemp = sTemp & "Import"
						break;
				case '9':  // "MailMerge"
						// Hide the remaining tabs
						$("#toolbarToolsFind").parent().hide();
						$("#toolbarReportFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();

						// Enable the buttons
						
						menu_toolbarEnableItem("mnutoolNewUtilitiesFind", IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolNewUtilitiesFind", true);
						menu_toolbarEnableItem("mnutoolCopyUtilitiesFind", fHasRows && IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolCopyUtilitiesFind", true);
						menu_toolbarEnableItem("mnutoolEditUtilitiesFind", fHasRows && (IsEditPermitted || IsViewPermitted) && isSingleRecord);
						menu_SetmnutoolButtonCaption("mnutoolEditUtilitiesFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditUtilitiesFind", true);
						menu_toolbarEnableItem("mnutoolDeleteUtilitiesFind", fHasRows && IsDeletePermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolDeleteUtilitiesFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesUtilitiesFind", fHasRows && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolPropertiesUtilitiesFind", true);

						menu_toolbarEnableItem("mnutoolRunUtilitiesFind", fHasRows && IsRunPermitted);
						//only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseUtilitiesFind', !isSingleRecord);
						menu_toolbarEnableItem('mnutoolCloseUtilitiesFind', !isSingleRecord);

						// Show and select the tab
						$("#toolbarUtilitiesFind").parent().show();
						$("#toolbarUtilitiesFind").click();
						break;

				case '10': // "Picklists"
						// Hide the remaining tabs
						$("#toolbarUtilitiesFind").parent().hide();
						$("#toolbarReportFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_toolbarEnableItem("mnutoolNewToolsFind", IsNewPermitted && isSingleRecord);
						menu_toolbarEnableItem("mnutoolCopyToolsFind", fHasRows && IsNewPermitted && isSingleRecord);
						menu_toolbarEnableItem("mnutoolEditToolsFind", fHasRows && (IsEditPermitted || IsViewPermitted) && isSingleRecord);
						menu_toolbarEnableItem("mnutoolPropertiesToolsFind", fHasRows && isSingleRecord);
						menu_toolbarEnableItem("mnutoolRunToolsFind", false);
						menu_setVisibleMenuItem('mnutoolRunToolsFind', false);
						// Show and select the tab
						$("#toolbarToolsFind").parent().show();
						$("#toolbarToolsFind").click();
						break;
				case '11': // "Filters"
						// Hide the remaining tabs
						$("#toolbarUtilitiesFind").parent().hide();
						$("#toolbarReportFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_toolbarEnableItem("mnutoolNewToolsFind", IsNewPermitted && isSingleRecord);
						menu_toolbarEnableItem("mnutoolCopyToolsFind", fHasRows && IsNewPermitted && isSingleRecord);
						menu_toolbarEnableItem("mnutoolEditToolsFind", fHasRows && (IsEditPermitted || IsViewPermitted) && isSingleRecord);
						menu_toolbarEnableItem("mnutoolPropertiesToolsFind", fHasRows && isSingleRecord);
						menu_toolbarEnableItem("mnutoolRunToolsFind", false);
						menu_setVisibleMenuItem('mnutoolRunToolsFind', false);
						// Show and select the tab
						$("#toolbarToolsFind").parent().show();
						$("#toolbarToolsFind").click();
						break;
				case '12': // "Calculations"
						// Hide the remaining tabs
						$("#toolbarUtilitiesFind").parent().hide();
						$("#toolbarReportFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_toolbarEnableItem("mnutoolNewToolsFind", IsNewPermitted && isSingleRecord);
						menu_toolbarEnableItem("mnutoolCopyToolsFind", fHasRows && IsNewPermitted && isSingleRecord);
						menu_toolbarEnableItem("mnutoolEditToolsFind", fHasRows && (IsEditPermitted || IsViewPermitted) && isSingleRecord);
						menu_toolbarEnableItem("mnutoolPropertiesToolsFind", fHasRows && isSingleRecord);
						menu_toolbarEnableItem("mnutoolRunToolsFind", false);
						menu_setVisibleMenuItem('mnutoolRunToolsFind', false);
						// Show and select the tab
						$("#toolbarToolsFind").parent().show();
						$("#toolbarToolsFind").click();
						break;
				case '17': // "CalendarReports"
						// Hide the remaining tabs
						$("#toolbarUtilitiesFind").parent().hide();
						$("#toolbarToolsFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_toolbarEnableItem("mnutoolNewReportFind", IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolNewReportFind", true);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows && IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows && (IsEditPermitted || IsViewPermitted) && isSingleRecord);
						menu_SetmnutoolButtonCaption("mnutoolEditReportFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditReportFind", true);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows && IsDeletePermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);

						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && IsRunPermitted);
						//only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseReportFind', !isSingleRecord);
						menu_toolbarEnableItem('mnutoolCloseReportFind', !isSingleRecord);

						// Show and select the tab
						$("#toolbarReportFind").parent().show();
						$("#toolbarReportFind").click();
						break;
				case '25': // "Workflow"
						// Hide the remaining tabs
						$("#toolbarToolsFind").parent().hide();
						$("#toolbarReportFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_setVisibleMenuItem("mnutoolNewUtilitiesFind", IsNewPermitted);
						menu_setVisibleMenuItem("mnutoolCopyUtilitiesFind", false);
						menu_setVisibleMenuItem("mnutoolEditUtilitiesFind", false);
						menu_setVisibleMenuItem("mnutoolDeleteUtilitiesFind", false);
						menu_setVisibleMenuItem("mnutoolPropertiesUtilitiesFind", false);
						menu_toolbarEnableItem("mnutoolRunUtilitiesFind", isSingleRecord);
						//only display the 'close' button for defsel when called from rec edit...
						if (isSingleRecord === true) {
								menu_setVisibleMenuItem('mnutoolCloseUtilitiesFind', true);
								menu_toolbarEnableItem('mnutoolCloseUtilitiesFind', true);
						}
						else {
								menu_setVisibleMenuItem('mnutoolCloseUtilitiesFind', false);
						}
						// Show and select the tab
						$("#toolbarUtilitiesFind").parent().show();
						$("#toolbarUtilitiesFind").click();
						break;
					case '35':  // "NineBoxGrid"
						// Hide the remaining tabs
						$("#toolbarUtilitiesFind").parent().hide();
						$("#toolbarToolsFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_toolbarEnableItem("mnutoolNewReportFind", IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolNewReportFind", true);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows && IsNewPermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows && (IsEditPermitted || IsViewPermitted) && isSingleRecord);
						menu_SetmnutoolButtonCaption("mnutoolEditReportFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditReportFind", true);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows && IsDeletePermitted && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows && isSingleRecord);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);
						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && IsRunPermitted && isSingleRecord);
						// Only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseReportFind', !isSingleRecord);
						menu_toolbarEnableItem('mnutoolCloseReportFind', !isSingleRecord);
						// Show and select the tab
						$("#toolbarReportFind").parent().show();
						$("#toolbarReportFind").click();
						break;
		}

		var fNoneSelected;
		var frmDefSel = document.getElementById('frmDefSel');

			//TODO - Check if anything selected
			//fNoneSelected = (frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Count == 0);

		button_disable(frmDefSel.cmdEdit, (fNoneSelected ||
				(!IsEditPermitted && !IsNewPermitted)));
		button_disable(frmDefSel.cmdNew, !IsNewPermitted);
		button_disable(frmDefSel.cmdCopy, (fNoneSelected || !IsNewPermitted));
		button_disable(frmDefSel.cmdDelete, (fNoneSelected ||
				!IsDeletePermitted ||
				(frmDefSel.cmdEdit.value.toUpperCase() == "VIEW")));

		if ((IsEditPermitted &&
				!IsViewPermitted) ||
				(frmDefSel.cmdEdit.value.toUpperCase() == "VIEW")) {
				frmDefSel.cmdEdit.value = "View";
				$('#mnutoolEditReportFind h6').text('View');
				$('#mnutoolEditReportFind a').attr('title', 'View');
		}
		else {
				frmDefSel.cmdEdit.value = "Edit";
				$('#mnutoolEditReportFind h6').text('Edit');
				$('#mnutoolEditReportFind a').attr('title', 'Edit');
		}

		button_disable(frmDefSel.cmdProperties, (fNoneSelected ||
				(!IsNewPermitted &&
				!IsEditPermitted &&
				!IsViewPermitted &&
				!IsDeletePermitted &&
				!IsDeletePermitted)));
		button_disable(frmDefSel.cmdRun, (fNoneSelected || !IsRunPermitted));

			// If delete permission is given for the report but the 'Read Only' permission has been given in Group Access then disable the delete button
		if (fHasRows && IsDeletePermitted && isSingleRecord) {
			DisableDeleteButtonIfDefinationHasReadOnlyAccess('mnutoolDeleteReportFind');
		}
	}

	// If the selected record has Read Only permission given in the Group Access then disable the delete button
	function DisableDeleteButtonIfDefinationHasReadOnlyAccess(menuItem) {
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
				var name = $("#utilname").val();
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
						var answer = OpenHR.messageBox("Delete this definition. Are you sure ?", 36, "Confirmation");

						if (answer == 6) {
								frmDefSel.action.value = "delete";
								OpenHR.submitForm(frmDefSel);
						}
				}
		}

		function setrun() {
				if (!$("#mnutillRunUtil").hasClass("disabled")) {
						var frmDefSel = document.getElementById('frmDefSel');

						frmDefSel.action.value = "run";

						var sUtilId;

						if (frmDefSel.utiltype.value == 25) {
								// Workflow
								var postData = {
									utiltype: frmDefSel.utiltype.value,
									ID: frmDefSel.utilid.value,
									Name: frmDefSel.utilname.value,
									__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
								}

								OpenHR.submitForm(null, "optionframe", null, postData, "util_run_workflow");

						} else {

								var frmPrompt = document.getElementById('frmPrompt');

								frmPrompt.utilid.value = frmDefSel.utilid.value;
								frmPrompt.utilname.value = frmDefSel.utilname.value;
								frmPrompt.action.value = frmDefSel.action.value;

								OpenHR.showInReportFrame(frmPrompt, false);

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
																								onclick="setrun();" />
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


		<form name="frmPrompt" method="post" action="util_run_promptedValues" id="frmPrompt" style="visibility: hidden; display: none">
				<input type="hidden" id="utiltype" name="utiltype" value="<%:CInt(iDefSelType)%>">
				<input type="hidden" id="utilid" name="utilid" value='<%=Session("utilid")%>'>
				<input type="hidden" id="utilname" name="utilname">
				<input type="hidden" id="action" name="action">
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
			width: 600,
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
			autoencode: true
		});

	}

	$(function () {
		attachDefSelGrid();

		$("#selectTable").change(function () {
			$('#SelectedTableID').val(($('#selectTable').val()));
			ToggleCheck();
		});
	});
		
</script>

