<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<%="" %>

<%
	Dim fGotId As Boolean
	Dim sTemp As String
	
	Session("objCalendar" & Session("UtilID")) = Nothing
	
	If Not String.IsNullOrEmpty(Request.Form("OnlyMine")) Then
		Session("OnlyMine") = Request.Form("OnlyMine")
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
			End Select
					
			Dim cmdDefSelOnlyMine As New Command
			cmdDefSelOnlyMine.CommandText = "sp_ASRIntGetSetting"
			cmdDefSelOnlyMine.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdDefSelOnlyMine.ActiveConnection = Session("databaseConnection")

			Dim prmSection = cmdDefSelOnlyMine.CreateParameter("section", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefSelOnlyMine.Parameters.Append(prmSection)
			prmSection.value = "defsel"

			Dim prmKey = cmdDefSelOnlyMine.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
			cmdDefSelOnlyMine.Parameters.Append(prmKey)
			prmKey.value = sTemp

			Dim prmDefault = cmdDefSelOnlyMine.CreateParameter("default", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefSelOnlyMine.Parameters.Append(prmDefault)
			prmDefault.value = "0"

			Dim prmUserSetting = cmdDefSelOnlyMine.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
			cmdDefSelOnlyMine.Parameters.Append(prmUserSetting)
			prmUserSetting.value = 1

			Dim prmResult = cmdDefSelOnlyMine.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
			cmdDefSelOnlyMine.Parameters.Append(prmResult)

			Err.Clear()
			cmdDefSelOnlyMine.Execute()
			Session("OnlyMine") = (CLng(cmdDefSelOnlyMine.Parameters("result").Value) = 1)

			cmdDefSelOnlyMine = Nothing
		Else
			If CStr(Session("OnlyMine")) = "" Then Session("OnlyMine") = False
		End If
	End If
	Session("fromMenu") = 0

	If CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined" Then
		If Not String.IsNullOrEmpty(Request.Form("txtTableID")) Then
			Session("utilTableID") = Request.Form("txtTableID")
		Else
			If Len(Session("tableID")) > 0 Then
				If CLng(Session("tableID")) > 0 Then
					Session("utilTableID") = Session("tableID")
					fGotId = True
				End If
			End If

			If fGotId = False Then
				If (CStr(Session("optionDefSelRecordID")) <> "") Then
					If (Session("optionDefSelRecordID") > 0) Then
						Session("utilTableID") = Session("Personnel_EmpTableID")
					Else
						If (Session("Personnel_EmpTableID") > 0) Then
							Session("utilTableID") = Session("Personnel_EmpTableID")
						Else
							Session("utilTableID") = 0
						End If
					End If
				Else
					If (Session("Personnel_EmpTableID") > 0) Then
						Session("utilTableID") = Session("Personnel_EmpTableID")
					Else
						Session("utilTableID") = 0
					End If
				End If
			End If
		End If
	End If
	
	If CStr(Session("optionDefSelType")) <> "" Then
		Session("defseltype") = Session("optionDefSelType")
	End If

	Session("tableID") = Session("utilTableID")
	
	If CStr(Session("singleRecordID")) <> "" Then
		If CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined" Then
			If CStr(Session("optionDefSelRecordID")) <> "" Then
				If Session("optionDefSelRecordID") > 0 Then
					Session("singleRecordID") = Session("optionDefSelRecordID")
				End If
			End If
		Else
			If CStr(Session("optionTableID")) <> "" Then
				If Session("optionTableID") > 0 Then
					Session("utilTableID") = Session("optionTableID")
				End If
			End If
			Session("tableID") = Session("utilTableID")
		End If
	Else
		Session("singleRecordID") = 0
	End If
	
	Session("optionDefSelType") = ""
	Session("optionTableID") = ""
	Session("optionDefSelRecordID") = ""
	
	If CStr(Session("utilTableID")) = "" Then
		Session("utilTableID") = 0
	End If

	If (Session("defseltype") <> 10) And (Session("defseltype") <> 11) And (Session("defseltype") <> 12) Then
		If CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined" Then
			Session("utilTableID") = 0
		End If
	Else 'defseltype=10 or 11 or 12 (picklist, filter or calculation)
		Session("utilTableID") = Session("Personnel_EmpTableID")
	End If
%>

<script type="text/javascript">
		
		function ssOleDBGridDefSelRecords_dblClick() {

				var frmDefSel = document.getElementById("frmDefSel");

				if ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12)) {
						// DblClick triggers Edit.
						setedit();
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

						if (answer == 6) {
								setrun();
						}
				}
				return false;
		}

		function ssOleDBGridDefSelRecords_rowcolchange() {

				var frmDefSel = document.getElementById("frmDefSel");
				var frmpermissions = document.getElementById("frmpermissions");

			// Populate the textbox with the definitions description
				frmDefSel.txtDescription.value = selectedRecordDetails("description");

				// Populate the hidden fields with the selected utils information       
				frmDefSel.utilid.value = $("#DefSelRecords").getGridParam('selrow');
				frmDefSel.utilname.value = selectedRecordDetails("Name");

				// Check for RO access and set EDIT/VIEW caption as appropriate
				var username = selectedRecordDetails("Username");
				var access = selectedRecordDetails("Access");

				button_disable(frmDefSel.cmdRun, (frmpermissions.grantrun.value == 0));
				button_disable(frmDefSel.cmdNew, (frmpermissions.grantnew.value == 0));
				button_disable(frmDefSel.cmdCopy, (frmpermissions.grantnew.value == 0));
				button_disable(frmDefSel.cmdEdit, (frmpermissions.grantedit.value == 0));

				if (username != frmDefSel.txtusername.value) {
						if (access == 'ro') {
								frmDefSel.cmdEdit.value = 'View';
								button_disable(frmDefSel.cmdDelete, true);
						} else {
								frmDefSel.cmdEdit.value = 'Edit';

								if (frmpermissions.grantdelete.value == 1) {
										button_disable(frmDefSel.cmdDelete, false);
								} else {
										button_disable(frmDefSel.cmdDelete, true);
								}
						}
				} else {
						frmDefSel.cmdEdit.value = 'Edit';

						if (frmpermissions.grantdelete.value == 1) {
								button_disable(frmDefSel.cmdDelete, false);
						} else {
								button_disable(frmDefSel.cmdDelete, true);
						}
				}

				refreshControls();
		}

		/* Return the value of a column selected in the find form. */
		function selectedRecordDetails(columnName) {

				var iRecordId;
				var rowId;

				rowId = $("#DefSelRecords").getGridParam('selrow');
				iRecordId = $("#DefSelRecords").find("#" + rowId + " #" + columnName).val();

				return (iRecordId);
		}

		function defsel_window_onload() {
		
				var frmDefSel = document.getElementById('frmDefSel');
				//if (frmDefSel.txtSingleRecordID.value > 0) {
				//	// Expand the option frame and hide the work frame.
				//	menu_disableMenu();
				//} else {
				//	refreshControls();
				//}
			
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

				tableToGrid("#DefSelRecords", {
						onSelectRow: function (rowID) {
								ssOleDBGridDefSelRecords_rowcolchange();
						},
						ondblClickRow: function (rowID) {
								ssOleDBGridDefSelRecords_dblClick();
						},
						rowNum: 1000    //TODO set this to blocksize...
				});

				$("#DefSelRecords").jqGrid('bindKeys', {
						"onEnter": function (rowid) {
								ssOleDBGridDefSelRecords_dblClick();
						}
				});

				$("#findGridRow").height("60%");
				$(window).bind('resize', function () {
					$("#DefSelRecords").setGridWidth($('#findGridRow').width(), true);
					$("#DefSelRecords").setGridHeight($("#findGridRow").height(), true);
				}).trigger('resize');

				$('#DefSelRecords').hideCol("description");
				$('#DefSelRecords').hideCol("Username");
				$('#DefSelRecords').hideCol("Access");
				$('#DefSelRecords').hideCol("ID");

				$("#DefSelRecords").setGridHeight($("#findGridRow").height());
				$("#DefSelRecords").setGridWidth($("#findGridRow").width());

				$("#DefSelRecords").closest('.ui-jqgrid-bdiv').width($("#DefSelRecords").closest('.ui-jqgrid-bdiv').width()+1);
		
				frmDefSel.cmdCancel.focus();

				refreshControls();

				if (rowCount() > 0) {

					<% if Session("utilid").ToString() > 0 then %>
						$("#DefSelRecords").setSelection(<% =Session("utilid").ToString()%>);					
					<% else %>
						var firstid = $("#DefSelRecords").getDataIDs()[0];
						$("#DefSelRecords").setSelection(firstid);
					<% end If %>				
					
				}

		}


		function rowCount() {
			return $("#DefSelRecords tr").length - 1;
		}


		//Case 0
		//sTemp = sTemp & "BatchJobs"
		//Case 1
		//sTemp = sTemp & "CrossTabs"
		//Case 2
		//sTemp = sTemp & "CustomReports"
		//Case 3
		//sTemp = sTemp & "DataTransfer"
		//Case 4
		//sTemp = sTemp & "Export"
		//Case 5
		//sTemp = sTemp & "GlobalAdd"
		//Case 6
		//sTemp = sTemp & "GlobalDelete"
		//Case 7
		//sTemp = sTemp & "GlobalUpdate"
		//Case 8
		//sTemp = sTemp & "Import"
		//Case 9
		//sTemp = sTemp & "MailMerge"
		//Case 10
		//sTemp = sTemp & "Picklists"
		//Case 11
		//sTemp = sTemp & "Filters"
		//Case 12
		//sTemp = sTemp & "Calculations"
		//Case 17
		//sTemp = sTemp & "CalendarReports"
		//Case 25
		//sTemp = sTemp & "Workflow"

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
			var fFromMenu;
			var fHasRows = (rowCount() > 0);
			
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
							menu_toolbarEnableItem("mnutoolNewReportFind", true);
							menu_setVisibleMenuItem("mnutoolNewReportFind", true);
							menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows);
							menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
							menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows);
							menu_setVisibleMenuItem("mnutoolEditReportFind", true);
							menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows);
							menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
							menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows);
							menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);
							menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows);
							//only display the 'close' button for defsel when called from rec edit...

							if (Number(frmDefSel.txtSingleRecordID.value) > 0) {
									menu_setVisibleMenuItem('mnutoolCloseReportFind', true);
									menu_toolbarEnableItem('mnutoolCloseReportFind', true);
							}
							else {
									menu_setVisibleMenuItem('mnutoolCloseReportFind', false);
							}
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
							menu_toolbarEnableItem("mnutoolNewReportFind", true);
							menu_setVisibleMenuItem("mnutoolNewReportFind", true);
							menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows);
							menu_setVisibleMenuItem("mnutoolCopyReportFind", true);
							menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows);
							menu_setVisibleMenuItem("mnutoolEditReportFind", true);
							menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows);
							menu_setVisibleMenuItem("mnutoolDeleteReportFind", true);
							menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows);
							menu_setVisibleMenuItem("mnutoolPropertiesReportFind", true);
							menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows);
							//only display the 'close' button for defsel when called from rec edit...
							if (Number(frmDefSel.txtSingleRecordID.value) > 0) {
									menu_setVisibleMenuItem('mnutoolCloseReportFind', true);
									menu_toolbarEnableItem('mnutoolCloseReportFind', true);
							}
							else {
									menu_setVisibleMenuItem('mnutoolCloseReportFind', false);
							}
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
						fFromMenu = (Number(frmDefSel.txtSingleRecordID.value) <= 0);
						menu_toolbarEnableItem("mnutoolNewUtilitiesFind", true);
						menu_setVisibleMenuItem("mnutoolNewUtilitiesFind", fFromMenu);
						menu_toolbarEnableItem("mnutoolCopyUtilitiesFind", fHasRows);
						menu_setVisibleMenuItem("mnutoolCopyUtilitiesFind", fFromMenu);
						menu_toolbarEnableItem("mnutoolEditUtilitiesFind", fHasRows);
						menu_setVisibleMenuItem("mnutoolEditUtilitiesFind", fFromMenu);
						menu_toolbarEnableItem("mnutoolDeleteUtilitiesFind", fHasRows);
						menu_setVisibleMenuItem("mnutoolDeleteUtilitiesFind", fFromMenu);
						menu_toolbarEnableItem("mnutoolPropertiesUtilitiesFind", fHasRows);
						menu_setVisibleMenuItem("mnutoolPropertiesUtilitiesFind", fFromMenu);

						menu_toolbarEnableItem("mnutoolRunUtilitiesFind", fHasRows);
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
							menu_toolbarEnableItem("mnutoolNewToolsFind", true);
							menu_toolbarEnableItem("mnutoolCopyToolsFind", true);
							menu_toolbarEnableItem("mnutoolEditToolsFind", true);
							menu_toolbarEnableItem("mnutoolDeleteToolsFind", true);
							menu_toolbarEnableItem("mnutoolPropertiesToolsFind", true);
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
							menu_toolbarEnableItem("mnutoolNewToolsFind", true);
							menu_toolbarEnableItem("mnutoolCopyToolsFind", true);
							menu_toolbarEnableItem("mnutoolEditToolsFind", true);
							menu_toolbarEnableItem("mnutoolDeleteToolsFind", true);
							menu_toolbarEnableItem("mnutoolPropertiesToolsFind", true);
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
							menu_toolbarEnableItem("mnutoolNewToolsFind", true);
							menu_toolbarEnableItem("mnutoolCopyToolsFind", true);
							menu_toolbarEnableItem("mnutoolEditToolsFind", true);
							menu_toolbarEnableItem("mnutoolDeleteToolsFind", true);
							menu_toolbarEnableItem("mnutoolPropertiesToolsFind", true);
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
						fFromMenu = (Number(frmDefSel.txtSingleRecordID.value) <= 0);
						menu_toolbarEnableItem("mnutoolNewReportFind", true);
						menu_setVisibleMenuItem("mnutoolNewReportFind", fFromMenu);
						menu_toolbarEnableItem("mnutoolCopyReportFind", fHasRows);
						menu_setVisibleMenuItem("mnutoolCopyReportFind", fFromMenu);
						menu_toolbarEnableItem("mnutoolEditReportFind", fHasRows);
						menu_setVisibleMenuItem("mnutoolEditReportFind", fFromMenu);
						menu_toolbarEnableItem("mnutoolDeleteReportFind", fHasRows);
						menu_setVisibleMenuItem("mnutoolDeleteReportFind", fFromMenu);
						menu_toolbarEnableItem("mnutoolPropertiesReportFind", fHasRows);
						menu_setVisibleMenuItem("mnutoolPropertiesReportFind", fFromMenu);

						menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows);
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
							menu_setVisibleMenuItem("mnutoolNewUtilitiesFind", false);
							menu_setVisibleMenuItem("mnutoolCopyUtilitiesFind", false);
							menu_setVisibleMenuItem("mnutoolEditUtilitiesFind", false);
							menu_setVisibleMenuItem("mnutoolDeleteUtilitiesFind", false);
							menu_setVisibleMenuItem("mnutoolPropertiesUtilitiesFind", false);
							menu_toolbarEnableItem("mnutoolRunUtilitiesFind", true);
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
			}
			else {
					frmDefSel.cmdEdit.value = "Edit";
			}

			button_disable(frmDefSel.cmdProperties, (fNoneSelected ||
					((frmpermissions.grantnew.value == 0) &&
							(frmpermissions.grantedit.value == 0) &&
							(frmpermissions.grantview.value == 0) &&
							(frmpermissions.grantdelete.value == 0) &&
							(frmpermissions.grantrun.value == 0))));
			button_disable(frmDefSel.cmdRun, (fNoneSelected || (frmpermissions.grantrun.value == 0)));
		}

		function showproperties() {
			if (!$("#mnutoolPropertiesUtil").hasClass("disabled")) {
				var sUrl;
				var frmDefSel = document.getElementById('frmDefSel');

				var frmProp = document.getElementById('frmProp');
				frmProp.prop_id.value = $("#DefSelRecords").getGridParam('selrow');
				frmProp.prop_name.value = selectedRecordDetails("Name");
				frmProp.utiltype.value = frmDefSel.utiltype.value;

				sUrl = "defselproperties" +
					"?prop_name=" + escape(frmProp.prop_name.value) +
					"&prop_id=" + frmProp.prop_id.value +
					"&utiltype=" + frmProp.utiltype.value;
				openDialog(sUrl, 600, 350);
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

		function openDialog(pDestination, pWidth, pHeight) {
				var dlgwinprops = "center:yes;" +
						"dialogHeight:" + pHeight + "px;" +
						"dialogWidth:" + pWidth + "px;" +
						"help:no;" +
						"resizable:yes;" +
						"scroll:yes;" +
						"status:no;";
				window.showModalDialog(pDestination, self, dlgwinprops);
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
					document.frmDefSel.action.value = "delete";
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
				OpenHR.showPopup("Loading form. Please wait...");
				document.frmDefSel.action.value = "new";
				OpenHR.submitForm(document.frmDefSel);
			}
		}

		function setcopy() {
			if (!$("#mnutoolCopyUtil").hasClass("disabled")) {
				var frmDefSel = document.getElementById('frmDefSel');

				OpenHR.showPopup("Copying definition. Please wait...");
				frmDefSel.action.value = "copy";
				OpenHR.submitForm(frmDefSel);
			}
		}

		function setedit() {

			if (!$("#mnutoolEditUtil").hasClass("disabled")) {
				var frmDefSel = document.getElementById('frmDefSel');

				OpenHR.showPopup("Loading definition. Please wait...");

				if (frmDefSel.cmdEdit.value == "Edit") {
					document.frmDefSel.action.value = "edit";
					OpenHR.submitForm(document.frmDefSel);
				} else {
					document.frmDefSel.action.value = "view";
					OpenHR.submitForm(document.frmDefSel);
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

			var sCurrentPage = $("#workframe").attr("data-framesource").replace(".asp", "");

			return (sCurrentPage.toUpperCase());
		}

</script>

<div id="defsel" data-framesource="defsel" style="display: block; height:100%">

<form id=frmpermissions name=frmpermissions style="visibility:hidden;display:none">
<%
	Dim cmdDefSelAccess As New Command
	cmdDefSelAccess.CommandText = "sp_ASRIntGetSystemPermissions"
	cmdDefSelAccess.CommandType = CommandTypeEnum.adCmdStoredProc
	cmdDefSelAccess.ActiveConnection = Session("databaseConnection")

	Err.Clear()
	Dim rstDefSelAccess = cmdDefSelAccess.Execute
	
	Dim fNewGranted = 0
	Dim fEditGranted = 0
	Dim fDeleteGranted = 0
	Dim fRunGranted = 0
	Dim fViewGranted = 0

	'MH20011008 We should probably change this so that we pass
	'over the category key and get back a smaller recordset

	do until rstdefselaccess.eof
		if session("defseltype") = 1 then
			if rstdefselaccess.fields(0).value = "CROSSTABS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CROSSTABS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CROSSTABS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CROSSTABS_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CROSSTABS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		if session("defseltype") = 2 then
			if rstdefselaccess.fields(0).value = "CUSTOMREPORTS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CUSTOMREPORTS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CUSTOMREPORTS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CUSTOMREPORTS_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CUSTOMREPORTS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		if session("defseltype") = 9 then
			if rstdefselaccess.fields(0).value = "MAILMERGE_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "MAILMERGE_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "MAILMERGE_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "MAILMERGE_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "MAILMERGE_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		if session("defseltype") = 10 then
			if rstdefselaccess.fields(0).value = "PICKLISTS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "PICKLISTS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "PICKLISTS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "PICKLISTS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		if session("defseltype") = 11 then
			if rstdefselaccess.fields(0).value = "FILTERS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "FILTERS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "FILTERS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "FILTERS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if
	
		if session("defseltype") = 12 then
			if rstdefselaccess.fields(0).value = "CALCULATIONS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALCULATIONS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALCULATIONS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALCULATIONS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if
		
		if session("defseltype") = 17 then
			if rstdefselaccess.fields(0).value = "CALENDARREPORTS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALENDARREPORTS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALENDARREPORTS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALENDARREPORTS_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALENDARREPORTS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if
		
		if session("defseltype") = 25 then
			fNewGranted = 0
			fEditGranted = 0
			fDeleteGranted = 0
			fViewGranted = 0

			if rstdefselaccess.fields(0).value = "WORKFLOW_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		rstdefselaccess.movenext
	loop
	
	rstDefSelAccess = Nothing
	cmdDefSelAccess = Nothing
	
	Response.Write("<INPUT type=hidden id=grantnew name=grantnew value = " & fNewGranted & ">" & vbCrLf)
	Response.Write("<INPUT type=hidden id=grantedit name=grantedit value = " & fEditGranted & ">" & vbCrLf)
	Response.Write("<INPUT type=hidden id=grantdelete name=grantdelete value = " & fDeleteGranted & ">" & vbCrLf)
	Response.Write("<INPUT type=hidden id=grantrun name=grantrun value = " & fRunGranted & ">" & vbCrLf)
	Response.Write("<INPUT type=hidden id=grantview name=grantview value = " & fViewGranted & ">" & vbCrLf)
%>
</form>

		<form name="frmDefSel" class="absolutefull" action="defsel_submit" method="post" id="frmDefSel">
<div id="findGridRow" style="height: 70%; margin-right: 20px; margin-left: 20px;">

										<table width="100%" height="100%" class="invisible">
												<tr>
														<td colspan="5" height="10">
																<span class="pageTitle">
																		<%
																				If Session("defseltype") = 0 Then           'BATCH JOB
																						Response.Write("Batch Jobs")
																				ElseIf Session("defseltype") = 1 Then       'CROSS TAB
																						Response.Write("Cross Tabs")
																				ElseIf Session("defseltype") = 2 Then       'CUSTOM REPORTS
																						Response.Write("Custom Reports")
																				ElseIf Session("defseltype") = 3 Then       'DATA TRANSFER
																						Response.Write("Data Transfer")
																				ElseIf Session("defseltype") = 4 Then       'EXPORT
																						Response.Write("Export")
																				ElseIf Session("defseltype") = 5 Then       'GLOBAL ADD
																						Response.Write("Global Add")
																				ElseIf Session("defseltype") = 6 Then       'GLOBAL UPDATE
																						Response.Write("Global Update")
																				ElseIf Session("defseltype") = 7 Then       'GLOBAL DELETE
																						Response.Write("Global Delete")
																				ElseIf Session("defseltype") = 8 Then       'IMPORT
																						Response.Write("Import")
																				ElseIf Session("defseltype") = 9 Then       'MAIL MERGE
																						Response.Write("Mail Merge")
																				ElseIf Session("defseltype") = 10 Then      'PICKLIST
																						Response.Write("Picklists")
																				ElseIf Session("defseltype") = 11 Then      'FILTERS
																						Response.Write("Filters")
																				ElseIf Session("defseltype") = 12 Then      'CALCULATIONS
																						Response.Write("Calculations")
																				ElseIf Session("defseltype") = 17 Then      'CALENDAR REPORTS
																						Response.Write("Calendar Reports")
																				ElseIf Session("defseltype") = 25 Then      'WORKFLOW
																						Response.Write("Workflow")
																				End If
																		%>
																</span>
														</td>
												</tr>

												<% 
														Dim sErrorDescription = ""
	
														If Session("defseltype") = 10 Or Session("defseltype") = 11 Or Session("defseltype") = 12 Then
												%>
												<tr height="10">

														<td height="10" colspan="3">
																<table width="100%" class="invisible">
																		<tr>
																				<td width="40">Table :
																				</td>
																				<td width="10">&nbsp;
																				</td>
																				<td width="175">
																						<select id="selectTable" name="selectTable" class="combo" style="HEIGHT: 22px; WIDTH: 200px">
																								<%
																										On Error Resume Next
	
																										If (Len(sErrorDescription) = 0) Then
																												' Get the view records.
																										Dim cmdTableRecords As New Command
																												cmdTableRecords.CommandText = "sp_ASRIntGetTables"
																										cmdTableRecords.CommandType = CommandTypeEnum.adCmdStoredProc
																												cmdTableRecords.ActiveConnection = Session("databaseConnection")

																												Err.Clear()
																												Dim rstTableRecords = cmdTableRecords.Execute

																												If (Err.Number <> 0) Then
																														sErrorDescription = "The table records could not be retrieved." & vbCrLf & FormatError(Err.Description)
																												End If

																												If (Len(sErrorDescription) = 0) Then
																														Do While Not rstTableRecords.EOF
																																Response.Write("						<OPTION value=" & rstTableRecords.Fields(0).Value)
																																If rstTableRecords.Fields(0).Value = CLng(Session("utilTableID")) Then
																																		Response.Write(" SELECTED")
																																End If

																																Response.Write(">" & Replace(CStr(rstTableRecords.Fields(1).Value), "_", " ") & "</OPTION>" & vbCrLf)

																																rstTableRecords.MoveNext()
																														Loop
					
																														' Release the ADO recordset object.
																														rstTableRecords.close()
																														rstTableRecords = Nothing
																												End If

																												' Release the ADO command object.
																												cmdTableRecords = Nothing
																										End If
																								%>
																						</select>
																				</td>
																				<td width="10">
																						<input type="button" value="Go" id="btnGoTable" name="btnGoTable" class="btn" onclick="ToggleCheck();" />
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
																						<%
																								If Len(sErrorDescription) = 0 Then
																										' Get the records.
																								Dim cmdDefSelRecords As New Command
																										cmdDefSelRecords.CommandText = "sp_ASRIntPopulateDefSel"
																								cmdDefSelRecords.CommandType = CommandTypeEnum.adCmdStoredProc
																										cmdDefSelRecords.ActiveConnection = Session("databaseConnection")

																										Dim prmType = cmdDefSelRecords.CreateParameter("type", 3, 1)
																										cmdDefSelRecords.Parameters.Append(prmType)
																										prmType.value = CleanNumeric(Session("defseltype"))

																										Dim prmOnlyMine = cmdDefSelRecords.CreateParameter("onlymine", 11, 1)
																										cmdDefSelRecords.Parameters.Append(prmOnlyMine)
																										prmOnlyMine.value = CleanBoolean(Session("OnlyMine")) ' 0 '1

																										Dim prmTableId = cmdDefSelRecords.CreateParameter("tableID", 3, 1)
																										cmdDefSelRecords.Parameters.Append(prmTableId)
																										prmTableId.value = CleanNumeric(Session("utilTableID"))

																										Err.Clear()
																										Dim rstDefSelRecords = cmdDefSelRecords.Execute

																										If (Err.Number <> 0) Then
																												sErrorDescription = "The Defsel records could not be retrieved." & vbCrLf & FormatError(Err.Description)
																										End If
																									 
																										If Len(sErrorDescription) = 0 Then
																												' Instantiate and initialise the grid. 
																									Response.Write("<table class='outline' style='width : 100%; ' id='DefSelRecords'>" & vbCrLf)																									
																									Response.Write("<tr class='header'>" & vbCrLf)																								
																									Response.Write("<th style='display: none;'>ID</th>")
																									
																									For iLoop = 0 To (rstDefSelRecords.Fields.Count - 1)
								
																										Dim headerStyle As New StringBuilder
																										Dim headerCaption As String
								
																										If Not rstDefSelRecords.Fields(iLoop).Name = "ID" Then
																											headerStyle.Append("width: 373px; ")
								
																											If rstDefSelRecords.Fields(iLoop).Name <> "Name" Then
																												headerStyle.Append("display: none; ")
																											End If

																											headerCaption = Replace(rstDefSelRecords.Fields(iLoop).Name.ToString(), "_", " ")
																											headerStyle.Append("text-align: left; ")
						
																											Response.Write("<th style='" & headerStyle.ToString() & "'>" & headerCaption & "</th>")
																										End If
																									Next

																									Response.Write("</tr>")
						
																												Dim lngRowCount = 0
																												Do While Not rstDefSelRecords.EOF
																														Dim sAddString = ""
																														Dim iLoop As Integer = 0

																										Dim IDRowNumber As Long = rstDefSelRecords.Fields("ID").Value
								

																										Response.Write("<tr disabled='disabled' id='" & IDRowNumber & "'>")
																										Response.Write("<td><input type='radio' id='sel' value='" & IDRowNumber & "'></td>")
																										
																										For iLoop = 0 To (rstDefSelRecords.Fields.Count - 1)
																											If Not rstDefSelRecords.Fields(iLoop).Name = "ID" Then																											
																												sAddString = CleanStringForHTML(rstDefSelRecords.Fields(iLoop).Value.ToString())
																												Response.Write("<td class='findGridCell' id='col_" & iLoop.ToString() & "'>" & sAddString & "<input id='" & rstDefSelRecords.Fields(iLoop).Name & "' type='hidden' value='" & sAddString & "'></td>")
																											End If
																										Next

																														Response.Write("</tr>")
																										'																										Response.Write("<input type='hidden' id=txtAddString_" & lngRowCount & " name=txtAddString_" & lngRowCount & " value=""" & sAddString & """>" & vbCrLf)

																														lngRowCount = lngRowCount + 1
																														rstDefSelRecords.MoveNext()
																
																												Loop

																												Response.Write("</table>")
						
																												' Release the ADO recordset object.
																												rstDefSelRecords.close()

																										End If
							
																								End If
																						%>
	 
																				</td>
																		</tr>

																		<tr height="10">
																				<td></td>
																		</tr>

																		<tr>
																				<td height="70">
																						<textarea cols="20" class="disabled" style="WIDTH: 100%;" name="txtDescription" rows="4" readonly="readonly" tabindex="-1">
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
																							If Not (CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined") Then
																								If (Session("singleRecordID") > 0) Or Session("defseltype") = 25 Then
																									Response.Write(" style=""visibility:hidden""")
																								End If
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
																							If Not (CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined") Then
																								If (Session("singleRecordID") > 0) Or Session("defseltype") = 25 Then
																									Response.Write(" style=""visibility:hidden""")
																								End If
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
																							If Not (CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined") Then
																								If (Session("singleRecordID") > 0) Or Session("defseltype") = 25 Then
																									Response.Write(" style=""visibility:hidden""")
																								End If
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
																							If Not (CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined") Then
																								If (Session("singleRecordID") > 0) Or Session("defseltype") = 25 Then
																									Response.Write(" style=""visibility:hidden""")
																								End If
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
																							If Not (CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined") Then
																								If (Session("singleRecordID") > 0) Or Session("defseltype") = 25 Then
																									Response.Write(" style=""visibility:hidden""")
																								End If
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
																							If Not (CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined") Then
																								If (Session("singleRecordID") > 0) Or Session("defseltype") = 25 Then
																									Response.Write(" style=""visibility:hidden""")
																								End If
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
																								If Session("defseltype") = 10 Or Session("defseltype") = 11 Or Session("defseltype") = 12 Then
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
																				If Session("defseltype") = 10 Or Session("defseltype") = 11 Or Session("defseltype") = 12 Then
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
													If Session("defseltype") = 25 Then
														Response.Write(" style=""visibility:hidden""")
															End If%>>
													<input type='hidden' id="txtusername" name="txtusername" value="<%=lcase(session("Username"))%>">
												</td>
											</tr>

												<tr>
														<td colspan="4" height="10"
																<%
																If Session("defseltype") = 25 Then
																		Response.Write(" style=""visibility:hidden""")
																End If
%>>
																<input <% If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" tabindex="-1" id="checkbox" name="checkbox" value="checkbox"
																		onclick="ToggleCheck();" />
																<label for="checkbox" class="checkbox" tabindex="0" onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}">
																		Only show definitions where owner is '<%=Session("Username")%>'
																</label>
														</td>
												</tr>
										</table>

				<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("defseltype")%>">
				<input type="hidden" id="utilid" name="utilid" value='<%=Session("utilid")%>'>
				<input type="hidden" id="utilname" name="utilname">
				<input type="hidden" id="action" name="action">
				<input type="hidden" id="txtTableID" name="txtTableID" value='<%=session("utilTableID")%>'>
				<input type="hidden" id="txtSingleRecordID" name="txtSingleRecordID" value='<%=session("singleRecordID")%>'>
</div>
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
				<input type="hidden" id="txtTableID" name="txtTableID" value='<%=Session("utilTableID")%>'>
		</form>

		<form target="properties" action="defselproperties" method="post" id="frmProp" name="frmProp" style="visibility: hidden; display: none">
				<input type="hidden" id="prop_name" name="prop_name">
				<input type="hidden" id="prop_id" name="prop_id">
				<input type="hidden" id="utiltype" name="utiltype">
		</form>

	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>

</div>

<script type="text/javascript">
		defsel_window_onload();
</script>
