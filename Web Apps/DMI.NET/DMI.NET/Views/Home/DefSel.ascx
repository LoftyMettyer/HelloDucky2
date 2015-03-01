<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="DMI.NET" %>

<%="" %>

<%
	Dim SelectedTableID As String = Request.Form("SelectedTableID")
	Dim fGotId As Boolean
	Dim sTemp As String
	Dim iBaseTableID As Integer

	Dim objSession As SessionInfo = CType(Session("sessionContext"), SessionInfo)
	
	Session("objCalendar" & Session("UtilID")) = Nothing
	
	If Not String.IsNullOrEmpty(Request.Form("OnlyMine")) Then
		Session("OnlyMine") = ValidateIntegerValue(Request.Form("OnlyMine"))
	Else
		If Session("fromMenu") = 1 Then
			' Read the defSel 'only mine' setting from the database.
			sTemp = "onlymine "
			Select Case Session("defseltype")
				Case 0
					sTemp = sTemp & "BatchJobs"
				Case 1
					sTemp = sTemp & "CrossTabs"
				Case 2
					sTemp = sTemp & "CustomReports"
				Case 3
					sTemp = sTemp & "DataTransfer"
				Case 4
					sTemp = sTemp & "Export"
				Case 5
					sTemp = sTemp & "GlobalAdd"
				Case 6
					sTemp = sTemp & "GlobalDelete"
				Case 7
					sTemp = sTemp & "GlobalUpdate"
				Case 8
					sTemp = sTemp & "Import"
				Case 9
					sTemp = sTemp & "MailMerge"
				Case 10
					sTemp = sTemp & "Picklists"
				Case 11
					sTemp = sTemp & "Filters"
				Case 12
					sTemp = sTemp & "Calculations"
				Case 17
					sTemp = sTemp & "CalendarReports"
				Case 25
					sTemp = sTemp & "Workflow"
				Case 35
					sTemp = sTemp & "NineBoxGrid"
			End Select

			Session("OnlyMine") = (CLng(objSession.GetUserSetting("defsel", sTemp, "0")) = 1)
	
		Else
			If CStr(Session("OnlyMine")) = "" Then Session("OnlyMine") = False
		End If
	End If

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
	
	If CStr(Session("optionDefSelType")) <> "" Then
		Session("defseltype") = Session("optionDefSelType")
	End If
	
	If Session("singleRecordID") = 0 Then
		If CStr(Session("optionTableID")) <> "" Then
			If Session("optionTableID") > 0 Then
				iBaseTableID = Session("optionTableID")
			End If
		End If
		Session("tableID") = Session("utilTableID")
	End If
	
	Session("optionDefSelType") = ""
	Session("optionTableID") = ""
	
	If (Session("defseltype") = UtilityType.utlPicklist) Or (Session("defseltype") = UtilityType.utlFilter) Or (Session("defseltype") = UtilityType.utlCalculation) Then
		iBaseTableID = CInt(Session("utilTableID"))
	End If
%>

<script type="text/javascript">

		//var fFromMenu = false;
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
			var frmpermissions = document.getElementById("frmpermissions");

			var rowId = $("#DefSelRecords").getGridParam('selrow');
			var gridData = $("#DefSelRecords").getRowData(rowId);

			frmDefSel.txtDescription.value = gridData.description;

			// Populate the hidden fields with the selected utils information       
			frmDefSel.utilid.value = $("#DefSelRecords").getGridParam('selrow');
			frmDefSel.utilname.value = gridData.Name;

			button_disable(frmDefSel.cmdRun, (frmpermissions.grantrun.value == 0));
			button_disable(frmDefSel.cmdNew, (frmpermissions.grantnew.value == 0));
			button_disable(frmDefSel.cmdCopy, (frmpermissions.grantnew.value == 0));
			button_disable(frmDefSel.cmdEdit, (frmpermissions.grantedit.value == 0));

			if (gridData.Username != frmDefSel.txtusername.value) {
				if (gridData.Access == 'ro') {
					frmDefSel.cmdEdit.value = 'View';
					menu_SetmnutoolButtonCaption("mnutoolEditToolsFind", "View");
					button_disable(frmDefSel.cmdDelete, true);
					menu_toolbarEnableItem("mnutoolDeleteToolsFind", false);
				} else {
					frmDefSel.cmdEdit.value = 'Edit';
					menu_SetmnutoolButtonCaption("mnutoolEditToolsFind", "Edit");

					if (frmpermissions.grantdelete.value == 1) {
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

				if (frmpermissions.grantdelete.value == 1) {
					button_disable(frmDefSel.cmdDelete, false);
					menu_toolbarEnableItem("mnutoolDeleteToolsFind", true);
				} else {
					button_disable(frmDefSel.cmdDelete, true);
					menu_toolbarEnableItem("mnutoolDeleteToolsFind", false);
				}
			}
			//fFromMenu = true;
			refreshControls();
		}
	
		function defsel_window_onload() {
			
				var frmDefSel = document.getElementById('frmDefSel');
			
				// Expand the option frame and hide the work frame.
				if (frmDefSel.txtSingleRecordID.value > 0) {
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

				refreshControls();

				if (rowCount() > 0) {

						var fFromMenu = (Number(frmDefSel.txtSingleRecordID.value) <= 0);
						var gotoID;

						if (fFromMenu == true) {
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
				var fFromMenu = (Number(frmDefSel.txtSingleRecordID.value) <= 0);
				var fHasRows = (rowCount() > 0);


				var IsNewPermitted = ($("#grantnew")[0].value > 0);
				var IsEditPermitted = ($("#grantedit")[0].value > 0);
				var IsViewPermitted = ($("#grantview")[0].value > 0);
				var IsDeletePermitted = ($("#grantdelete")[0].value > 0);
				var IsRunPermitted = ($("#grantrun")[0].value > 0);

				switch ('<%=Session("defseltype")%>') {
				case '0':  // "BatchJobs"
						break;
				case '1':  // "CrossTabs"
						// Hide the remaining tabs
						$("#toolbarUtilitiesFind").parent().hide();
						$("#toolbarToolsFind").parent().hide();
						$("#toolbarEventLogFind").parent().hide();
						$("#toolbarWFPendingStepsFind").parent().hide();
						// Enable the buttons
						menu_toolbarEnableItem("mnutoolNewReportFind", IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolNewReportFind", true);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows && IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows && (IsEditPermitted || IsViewPermitted) && fFromMenu);
						menu_SetmnutoolButtonCaption("mnutoolEditReportFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditReportFind", true);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows && IsDeletePermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows && fFromMenu);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);
						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && IsRunPermitted && fFromMenu);
						// Only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseReportFind', !fFromMenu);
						menu_toolbarEnableItem('mnutoolCloseReportFind', !fFromMenu);
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
						menu_toolbarEnableItem("mnutoolNewReportFind", IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolNewReportFind", true);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows && IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows && (IsEditPermitted || IsViewPermitted) && fFromMenu);
						menu_SetmnutoolButtonCaption("mnutoolEditReportFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditReportFind", true);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows && IsDeletePermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows && fFromMenu);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);
						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && IsRunPermitted && fFromMenu);
						// Only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseReportFind', !fFromMenu);
						menu_toolbarEnableItem('mnutoolCloseReportFind', !fFromMenu);
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
						
						menu_toolbarEnableItem("mnutoolNewUtilitiesFind", IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolNewUtilitiesFind", true);
						menu_toolbarEnableItem("mnutoolCopyUtilitiesFind", fHasRows && IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolCopyUtilitiesFind", true);
						menu_toolbarEnableItem("mnutoolEditUtilitiesFind", fHasRows && (IsEditPermitted || IsViewPermitted) && fFromMenu);
						menu_SetmnutoolButtonCaption("mnutoolEditUtilitiesFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditUtilitiesFind", true);
						menu_toolbarEnableItem("mnutoolDeleteUtilitiesFind", fHasRows && IsDeletePermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolDeleteUtilitiesFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesUtilitiesFind", fHasRows && fFromMenu);
						menu_setVisibleMenuItem("mnutoolPropertiesUtilitiesFind", true);

						menu_toolbarEnableItem("mnutoolRunUtilitiesFind", fHasRows && IsRunPermitted);
						//only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseUtilitiesFind', !fFromMenu);
						menu_toolbarEnableItem('mnutoolCloseUtilitiesFind', !fFromMenu);

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
						menu_toolbarEnableItem("mnutoolNewToolsFind", IsNewPermitted && fFromMenu);
						menu_toolbarEnableItem("mnutoolCopyToolsFind", fHasRows && IsNewPermitted && fFromMenu);
						menu_toolbarEnableItem("mnutoolEditToolsFind", fHasRows && (IsEditPermitted || IsViewPermitted) && fFromMenu);
						menu_toolbarEnableItem("mnutoolPropertiesToolsFind", fHasRows && fFromMenu);
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
						menu_toolbarEnableItem("mnutoolNewToolsFind", IsNewPermitted && fFromMenu);
						menu_toolbarEnableItem("mnutoolCopyToolsFind", fHasRows && IsNewPermitted && fFromMenu);
						menu_toolbarEnableItem("mnutoolEditToolsFind", fHasRows && (IsEditPermitted || IsViewPermitted) && fFromMenu);
						menu_toolbarEnableItem("mnutoolPropertiesToolsFind", fHasRows && fFromMenu);
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
						menu_toolbarEnableItem("mnutoolNewToolsFind", IsNewPermitted && fFromMenu);
						menu_toolbarEnableItem("mnutoolCopyToolsFind", fHasRows && IsNewPermitted && fFromMenu);
						menu_toolbarEnableItem("mnutoolEditToolsFind", fHasRows && (IsEditPermitted || IsViewPermitted) && fFromMenu);
						menu_toolbarEnableItem("mnutoolPropertiesToolsFind", fHasRows && fFromMenu);
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
						//fFromMenu = (Number(frmDefSel.txtSingleRecordID.value) <= 0);
						menu_toolbarEnableItem("mnutoolNewReportFind", IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolNewReportFind", true);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows && IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows && (IsEditPermitted || IsViewPermitted) && fFromMenu);
						menu_SetmnutoolButtonCaption("mnutoolEditReportFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditReportFind", true);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows && IsDeletePermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows && fFromMenu);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);

						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && IsRunPermitted);
						//only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseReportFind', !fFromMenu);
						menu_toolbarEnableItem('mnutoolCloseReportFind', !fFromMenu);

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
						menu_toolbarEnableItem("mnutoolRunUtilitiesFind", fFromMenu);
						//only display the 'close' button for defsel when called from rec edit...
						if (Number(frmDefSel.txtSingleRecordID.value) > 0) {
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
						menu_toolbarEnableItem("mnutoolNewReportFind", IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolNewReportFind", true);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows && IsNewPermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows && (IsEditPermitted || IsViewPermitted) && fFromMenu);
						menu_SetmnutoolButtonCaption("mnutoolEditReportFind", (IsEditPermitted == false ? 'View' : 'Edit'));
						menu_setVisibleMenuItem("mnutoolEditReportFind", true);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows && IsDeletePermitted && fFromMenu);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows && fFromMenu);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);
						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && IsRunPermitted && fFromMenu);
						// Only display the 'close' button for defsel when called from rec edit...
						menu_setVisibleMenuItem('mnutoolCloseReportFind', !fFromMenu);
						menu_toolbarEnableItem('mnutoolCloseReportFind', !fFromMenu);
						// Show and select the tab
						$("#toolbarReportFind").parent().show();
						$("#toolbarReportFind").click();
						break;
		}
			//menu_toolbarEnableItem("mnutoolNewReportFind", true);
			//menu_toolbarEnableItem("mnutoolCopyReportFind", true);
			//menu_toolbarEnableItem("mnutoolEditReportFind", true);
			//menu_toolbarEnableItem("mnutoolDeleteReportFind", true);
			//menu_toolbarEnableItem("mnutoolPropertiesReportFind", true);
			//menu_toolbarEnableItem("mnutoolRunReportFind", true);
			////only display the 'close' button for defsel when called from rec edit...
			//if (Number(frmDefSel.txtSingleRecordID.value) > 0) {
			//    menu_setVisibleMenuItem('mnutoolCloseReportFind', true);
			//    menu_toolbarEnableItem('mnutoolCloseReportFind', true);
			//}
			//$("#toolbarReportFind").click();

		var fNoneSelected;
		var frmpermissions = document.getElementById('frmpermissions');
		var frmDefSel = document.getElementById('frmDefSel');

			//TODO - Check if anything selected
			//fNoneSelected = (frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Count == 0);

		button_disable(frmDefSel.cmdEdit, (fNoneSelected ||
				((frmpermissions.grantedit.value == 0) && (frmpermissions.grantview.value == 0))));
		button_disable(frmDefSel.cmdNew, (frmpermissions.grantnew.value == 0));
		button_disable(frmDefSel.cmdCopy, (fNoneSelected || (frmpermissions.grantnew.value == 0)));
		button_disable(frmDefSel.cmdDelete, (fNoneSelected ||
				(frmpermissions.grantdelete.value == 0) ||
				(frmDefSel.cmdEdit.value.toUpperCase() == "VIEW")));

		if (((frmpermissions.grantedit.value == 0) &&
				(frmpermissions.grantview.value == 1)) ||
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
				((frmpermissions.grantnew.value == 0) &&
				(frmpermissions.grantedit.value == 0) &&
				(frmpermissions.grantview.value == 0) &&
				(frmpermissions.grantdelete.value == 0) &&
				(frmpermissions.grantrun.value == 0))));
		button_disable(frmDefSel.cmdRun, (fNoneSelected || (frmpermissions.grantrun.value == 0)));

			// If delete permission is given for the report but the 'Read Only' permission has been given in Group Access then disable the delete button
			if (fHasRows && IsDeletePermitted && fFromMenu) {
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
				var frmOnlyMine = document.getElementById('frmOnlyMine');
				var frmDefSel = document.getElementById('frmDefSel');

				if ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12)) {
						frmOnlyMine.txtTableID.value = frmDefSel.selectTable.options[frmDefSel.selectTable.selectedIndex].value;
						frmDefSel.txtTableID.value = frmOnlyMine.txtTableID.value;
				}

				frmOnlyMine.OnlyMine.value = frmDefSel.checkbox.checked;

				OpenHR.submitForm(frmOnlyMine);
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
								var frmWorkflow = document.getElementById('frmWorkflow');
								frmWorkflow.utiltype.value = frmDefSel.utiltype.value;
								frmWorkflow.utilid.value = frmDefSel.utilid.value;
								frmWorkflow.utilname.value = frmDefSel.utilname.value;
								frmWorkflow.action.value = frmDefSel.action.value;
								sUtilId = new String(frmDefSel.utilid.value);

								frmWorkflow.target = sUtilId;
								//NewWindow('', sUtilId, '500', '200', 'yes');
								OpenHR.submitForm(frmWorkflow, 'optionframe', false);
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
				var frmDefSel = document.getElementById('frmDefSel');
				if (frmDefSel.txtSingleRecordID.value > 0) {
						var sWorkPage = defsel_currentWorkFramePage();
						if (sWorkPage == "RECORDEDIT") {
								refreshData(); //workframe
						}

						loadEmptyOption();

						menu_disableMenu();

						$("#optionframe").hide();
						$("#workframe").show();
						$("#toolbarRecord").show();
						$("#toolbarRecord").click();
					
						menu_refreshMenu();
				}
				else {
						window.location.href = "_default";
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

<form id=frmpermissions name=frmpermissions style="visibility:hidden;display:none">
<%
	
	Dim fNewGranted As Boolean
	Dim fEditGranted As Boolean
	Dim fDeleteGranted As Boolean
	Dim fRunGranted As Boolean
	Dim fViewGranted As Boolean
	
	Dim strKeyPrefix As String = ""
	
	Select Case Session("defseltype")
		Case "1"
			strKeyPrefix = "CROSSTABS"

		Case "2"
			strKeyPrefix = "CUSTOMREPORTS"
				
		Case "9"
			strKeyPrefix = "MAILMERGE"
				
		Case "10"
			strKeyPrefix = "PICKLISTS"
				
		Case "11"
			strKeyPrefix = "FILTERS"

		Case "12"
			strKeyPrefix = "CALCULATIONS"
				
		Case "17"
			strKeyPrefix = "CALENDARREPORTS"

		Case "25"
			strKeyPrefix = "WORKFLOW"
		
		Case "35"
			strKeyPrefix = "NINEBOXGRID"
	End Select
	
	fNewGranted = objSession.IsPermissionGranted(strKeyPrefix, "NEW")
	fEditGranted = objSession.IsPermissionGranted(strKeyPrefix, "EDIT")
	fDeleteGranted = objSession.IsPermissionGranted(strKeyPrefix, "DELETE")
	fRunGranted = objSession.IsPermissionGranted(strKeyPrefix, "RUN")
	fViewGranted = objSession.IsPermissionGranted(strKeyPrefix, "VIEW")
	
	Response.Write("<input type=hidden id=""grantnew"" name=""grantnew"" value = " & IIf(fNewGranted, 1, 0) & ">" & vbCrLf)
	Response.Write("<input type=hidden id=""grantedit"" name=""grantedit"" value = " & IIf(fEditGranted, 1, 0) & ">" & vbCrLf)
	Response.Write("<input type=hidden id=""grantdelete"" name=""grantdelete"" value = " & IIf(fDeleteGranted, 1, 0) & ">" & vbCrLf)
	Response.Write("<input type=hidden id=""grantrun"" name=""grantrun"" value = " & IIf(fRunGranted, 1, 0) & ">" & vbCrLf)
	Response.Write("<input type=hidden id=""grantview"" name=""grantview"" value = " & IIf(fViewGranted, 1, 0) & ">" & vbCrLf)
%>
</form>

		<form name="frmDefSel" class="absolutefull" action="defsel_submit" method="post" id="frmDefSel">
<div id="findGridRow" style="height: 70%; margin-right: 20px; margin-left: 20px;">

										<table width="100%" height="100%" class="invisible">
												<tr>
														<td colspan="5" height="10">
																<span class="pageTitle">
																		<%
																			If Session("defseltype") = UtilityType.utlBatchJob Then
																				Response.Write("Batch Jobs")
																			ElseIf Session("defseltype") = UtilityType.utlCrossTab Then
																				Response.Write("Cross Tabs")
																			ElseIf Session("defseltype") = UtilityType.utlCustomReport Then
																				Response.Write("Custom Reports")
																			ElseIf Session("defseltype") = UtilityType.utlDataTransfer Then
																				Response.Write("Data Transfer")
																			ElseIf Session("defseltype") = UtilityType.utlExport Then
																				Response.Write("Export")
																			ElseIf Session("defseltype") = UtilityType.UtlGlobalAdd Then
																				Response.Write("Global Add")
																			ElseIf Session("defseltype") = UtilityType.utlGlobalUpdate Then
																				Response.Write("Global Update")
																			ElseIf Session("defseltype") = UtilityType.utlGlobalDelete Then
																				Response.Write("Global Delete")
																			ElseIf Session("defseltype") = UtilityType.utlImport Then
																				Response.Write("Import")
																			ElseIf Session("defseltype") = UtilityType.utlMailMerge Then
																				Response.Write("Mail Merge")
																			ElseIf Session("defseltype") = UtilityType.utlPicklist Then
																				Response.Write("Picklists")
																			ElseIf Session("defseltype") = UtilityType.utlFilter Then
																				Response.Write("Filters")
																			ElseIf Session("defseltype") = UtilityType.utlCalculation Then
																				Response.Write("Calculations")
																			ElseIf Session("defseltype") = UtilityType.utlCalendarReport Then
																				Response.Write("Calendar Reports")
																			ElseIf Session("defseltype") = UtilityType.utlWorkflow Then
																				Response.Write("Workflow")
																			ElseIf Session("defseltype") = UtilityType.utlNineBoxGrid Then
																				Response.Write("9-Box Grid Reports")
																			End If
																		%>
																</span>
														</td>
												</tr>

												<% 
														Dim sErrorDescription = ""
	
													If Session("defseltype") = UtilityType.utlPicklist Or Session("defseltype") = UtilityType.utlFilter Or Session("defseltype") = UtilityType.utlCalculation Then
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
																							If (Session("singleRecordID") > 0) Or Session("defseltype") = UtilityType.utlWorkflow Then
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
																							If (Session("singleRecordID") > 0) Or Session("defseltype") = UtilityType.utlWorkflow Then
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
																							If (Session("singleRecordID") > 0) Or Session("defseltype") = UtilityType.utlWorkflow Then
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
																							If (Session("singleRecordID") > 0) Or Session("defseltype") = UtilityType.utlWorkflow Then
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

																							If (Session("singleRecordID") > 0) Or Session("defseltype") = UtilityType.utlWorkflow Then
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
																							If (Session("singleRecordID") > 0) Or Session("defseltype") = UtilityType.utlWorkflow Then
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
																							If Session("defseltype") = UtilityType.utlPicklist Or Session("defseltype") = UtilityType.utlFilter Or Session("defseltype") = UtilityType.utlCalculation Then
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
																				If Session("defseltype") = UtilityType.utlPicklist Or Session("defseltype") = UtilityType.utlFilter Or Session("defseltype") = UtilityType.utlCalculation Then
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
													If Session("defseltype") = UtilityType.utlWorkflow Then
														Response.Write(" style=""visibility:hidden""")
															End If%>>
													<input type='hidden' id="txtusername" name="txtusername" value="<%=lcase(session("Username"))%>">
												</td>
											</tr>

												<tr>
														<td colspan="4" height="10"
																<%
															If Session("defseltype") = UtilityType.utlWorkflow Then
																Response.Write(" style=""visibility:hidden""")
															End If
%>>
																<input <% If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" tabindex="0" id="checkbox" name="checkbox" value="checkbox"
																		onclick="ToggleCheck();" />
																<label for="checkbox" class="checkbox" tabindex="-1" onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}">
																		Only show definitions where owner is '<%=Session("Username")%>'
																</label>
														</td>
												</tr>
										</table>

				<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("defseltype")%>">
				<input type="hidden" id="utilid" name="utilid" value='<%=Session("utilid")%>'>
				<input type="hidden" id="utilname" name="utilname">
				<input type="hidden" id="action" name="action">
				<input type="hidden" id="txtTableID" name="txtTableID" value='<%=iBaseTableID%>'>
				<input type="hidden" id="txtSingleRecordID" name="txtSingleRecordID" value='<%=session("singleRecordID")%>'>
</div>
			<%=Html.AntiForgeryToken()%>
		</form>


		<form name="frmPrompt" method="post" action="util_run_promptedValues" id="frmPrompt" style="visibility: hidden; display: none">
				<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("defseltype")%>">
				<input type="hidden" id="utilid" name="utilid" value='<%=Session("utilid")%>'>
				<input type="hidden" id="utilname" name="utilname">
				<input type="hidden" id="action" name="action">
		</form>

		<form name="frmWorkflow" method="post" action="util_run_workflow" id="frmWorkflow" style="visibility: hidden; display: none">
				<input type="hidden" id="utiltype" name="utiltype">
				<input type="hidden" id="utilid" name="utilid">
				<input type="hidden" id="utilname" name="utilname">
				<input type="hidden" id="action" name="action">
		</form>

		<form action="defsel" method="post" id="frmOnlyMine" name="frmOnlyMine" style="visibility: hidden; display: none">
				<input type="hidden" id="OnlyMine" name="OnlyMine" value='<%=Session("OnlyMine")%>'>
				<input type="hidden" id="txtTableID" name="txtTableID" value='<%=iBaseTableID%>'>
				<input type="hidden" id="SelectedTableID" name="SelectedTableID">
				<%=Html.AntiForgeryToken()%>
		</form>


	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">
	
	<input type="hidden" id="lastSelectedID" name="lastSelectedID" value='<%=Session("utilid")%>'>

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
		<%=Html.AntiForgeryToken()%>
	</form>

	<form action="emptyoption_Submit" method="post" id="frmGotoOption" name="frmGotoOption" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
		<%=Html.AntiForgeryToken()%>
	</form>

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

		var onlyMine = frmOnlyMine.OnlyMine.value;

	//	$("#DefSelRecords").jqGrid('GridUnload');

		$("#DefSelRecords").jqGrid({
			url: 'GetDefinitionsForType?UtilityType=' + <%=Session("defseltype")%> + '&&TableID=' + <%=iBaseTableID%> + '&&OnlyMine=' + onlyMine,
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
			loadonce: true
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

