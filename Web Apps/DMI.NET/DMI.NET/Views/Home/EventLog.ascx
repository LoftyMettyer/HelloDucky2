<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="DMI.NET.Helpers" %>

<%

	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

	'This section of script is used for saving the new purge criteria.
	Dim sDoesPurge As String = Trim(Request.Form("txtDoesPurge"))
	Dim sPeriod As String = Request.Form("txtPurgePeriod")
	Dim iFrequency As Integer = CInt(Request.Form("txtPurgeFrequency"))
	
	If sDoesPurge <> vbNullString Then
		
		' Delete old purge information to the database
		objDataAccess.ExecuteSP("spASRIntClearEventLogPurge")
	
		If sDoesPurge = "1" Then

			' Insert the new purge criteria	
			objDataAccess.ExecuteSP("spASRIntSetEventLogPurge" _
						, New SqlParameter("psPeriod", SqlDbType.VarChar, 2) With {.Value = sPeriod} _
						, New SqlParameter("piFrequency", SqlDbType.Int) With {.Value = iFrequency})
					
			Session("showPurgeMessage") = 1
		Else
			Session("showPurgeMessage") = 0
		End If
	End If
	
	'This section of script is used for deleting Event Log records according to the selection on the Delete screen. 
	Dim sDeleteSelection As String = Request.Form("txtDeleteSel")
	Dim sSelectedEventIDs As String = Request.Form("txtSelectedIDs")
	Dim sHasViewAllPermission As Boolean = CBool(CleanBoolean(Request.Form("txtViewAllPerm")))
	
	If Not sDeleteSelection Is Nothing Then
			
		objDataAccess.ExecuteSP("spASRIntDeleteEventLogRecords" _
			, New SqlParameter("piDeleteType", SqlDbType.Int) With {.Value = CInt(sDeleteSelection)} _
			, New SqlParameter("psSelectedEventIDs", SqlDbType.VarChar, -1) With {.Value = sSelectedEventIDs} _
			, New SqlParameter("pfCanViewAll", SqlDbType.Bit) With {.Value = sHasViewAllPermission})
		
	End If

%>

<script type="text/javascript">
	var frmEventUseful = OpenHR.getForm("workframe", "frmEventUseful");
	var frmPurge = OpenHR.getForm("workframe", "frmPurge");
	
	var frmDetails = OpenHR.getForm("workframe", "frmDetails");
	var frmRefresh = OpenHR.getForm("workframe", "frmRefresh");

	function eventLog_window_onload() {
		$("#workframe").attr("data-framesource", "EVENTLOG");
		var frmLog = OpenHR.getForm("workframe", "frmLog");
		
		var sErrMsg = $('#txtErrorDescription').val();
		var fOK;
		fOK = true;

		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		// Hide the on-screen buttons
		$('#cmdView').hide();
		$('#cmdDelete').hide();
		$('#cmdPurge').hide();
		$('#cmdEmail').hide();

		if (fOK == true) {
			// Get menu to refresh the menu.
			menu_refreshMenu();

		}

		frmLog.txtELDeletePermission.value = menu_GetItemValue("txtSysPerm_EVENTLOG_DELETE");
		frmLog.txtELViewAllPermission.value = menu_GetItemValue("txtSysPerm_EVENTLOG_VIEWALL");
		frmLog.txtELPurgePermission.value = menu_GetItemValue("txtSysPerm_EVENTLOG_PURGE");
		frmLog.txtELEmailPermission.value = menu_GetItemValue("txtSysPerm_EVENTLOG_EMAIL");

		// Buttons are enabled/disabled in EventLog_refreshButtons();
		// Visible buttons and tabs set in menu.js

		refreshUsers();

		$("#optionframe").hide();
		$("#workframe").show();
	}

	function EventLog_moveRecord(psMovement) {
		var frmGetData = OpenHR.getForm("dataframe", "frmGetData");
		var frmData = OpenHR.getForm("dataframe", "frmData");

		frmGetData.txtELAction.value = psMovement;
		frmGetData.txtELCurrRecCount.value = frmData.txtELCurrentRecCount.value;
		frmGetData.txtEL1stRecPos.value = frmData.txtELFirstRecPos.value;

		refreshGrid();

		return;
	}
	
	function loadEventLog() {
		var i;
		var frmLog = OpenHR.getForm("workframe", "frmLog");

		//Clear the log table
		$("#LogEvents").jqGrid('GridUnload');
		
		var colNames = ['ID', 'Start Time', 'End Time', 'Duration', 'Type', 'Name', 'Status', 'Mode', 'User name', 'Batch Run ID'];
		var colData = [];
		var obj;

		//Get the data from the hidden inputs
		$("input[id^='txtAddString_']").each(function () {
			obj = {};
			var splitValue = this.value.split("\t");
			//We can't use a forEach loop becasue we want to ignore some of the values, so we'll use a 'traditional' loop
			for (i = 0; i <= 8; i++) {
				obj[colNames[i]] = splitValue[i];
			};

			obj["Batch Run ID"] = splitValue[10]; //We also need this value

			colData.push(obj);
		});

		$("#LogEvents").jqGrid({
			colNames: colNames,
			datatype: 'local',
			data: colData,
			colModel: [
				{ name: 'ID', hidden: true },
				{ name: 'Start Time', width: 120 },
				{ name: 'End Time', width: 120 },
				{ name: 'Duration', width: 85 },
				{ name: 'Type' },
				{ name: 'Name', width: 355 },
				{ name: 'Status', width: 120 },
				{ name: 'Mode', width: 100 },
				{ name: 'User name' },
				{ name: 'Batch Run ID', hidden: true }
			],
			multiselect: true,
			loadComplete: function () {
				moveFirst();
			},
			onSelectRow: function () { //Enable ribbon
			},
			ondblClickRow: function () {
				EventLog_viewEvent();
			},
			cmTemplate: { sortable: true },
			pager: $('#pager-coldata'),
			ignoreCase: true,
			autoencode: true,
			shrinkToFit: false,
			rowNum: 500,
			beforeSelectRow: function (rowid, e) { // handle jqGrid multiselect => thanks to solution from Byron Cobb on http://goo.gl/UvGku
				if (!e.ctrlKey && !e.shiftKey) {
					$("#LogEvents").jqGrid('resetSelection');
				} else if (e.shiftKey) {
					var initialRowSelect = $("#LogEvents").jqGrid('getGridParam', 'selrow');
					$("#LogEvents").jqGrid('resetSelection');

					var CurrentSelectIndex = $("#LogEvents").jqGrid('getInd', rowid);
					var InitialSelectIndex = $("#LogEvents").jqGrid('getInd', initialRowSelect);
					var startID;
					var endID;
					if (CurrentSelectIndex > InitialSelectIndex) {
						startID = initialRowSelect;
						endID = rowid; 
					}
					else {
						startID = rowid;
						endID = initialRowSelect;
					}

					var shouldSelectRow = false;
					$.each($("#LogEvents").getDataIDs(), function (_, id) {
						if ((shouldSelectRow = id == startID || shouldSelectRow)) {
							$("#LogEvents").jqGrid('setSelection', id, false);
						}
						return id != endID;
					});
				}
				return true;
			}
		}).jqGrid('hideCol', 'cb');

		//search options.
		$("#LogEvents").jqGrid('navGrid', '#pager-coldata', { del: false, add: false, edit: false, search: false });

		$("#LogEvents").jqGrid('navButtonAdd', "#pager-coldata", {
			caption: '',
			buttonicon: 'ui-icon-search',
			onClickButton: function () {
				$("#LogEvents").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });
			},
			position: 'first',
			title: '',
			cursor: 'pointer'
		});

		$("#LogEvents").jqGrid('setGridHeight', $("#gridContainer").height());
		$("#LogEvents").jqGrid('setGridWidth', $("#gridContainer").width());

		frmLog.cboUsername.style.color = 'black';
		frmLog.cboType.style.color = 'black';
		frmLog.cboMode.style.color = 'black';
		frmLog.cboStatus.style.color = 'black';

		//Set the event log loaded flag, used in the menu
		frmLog.txtELLoaded.value = 1;

		// Get menu to refresh the menu.
		menu_refreshMenu();

		EventLog_refreshButtons();

		refreshStatusBar();

		if ($('#txtShowPurgeMSG').val() == 1) {
			OpenHR.modalPrompt("Purge Event Log completed.",0,"Event Log");
			$('#EventLogPurge').dialog('close');
			$('#txtShowPurgeMSG').val(0);
		}
	}

	function moveFirst() {
		$("#LogEvents").jqGrid('setSelection', 1);
	}

	function filterSQL() {
		var SSql = new String("");
		var frmLog = OpenHR.getForm("workframe", "frmLog");

		if (frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value != -1) {
			var sUsername = new String(frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value);
			SSql = SSql + " LOWER(Username) = '" + sUsername.toLowerCase() + "' ";
		}

		if (frmLog.cboType.options[frmLog.cboType.selectedIndex].value != -1) {
			if (SSql.length > 0) {
				SSql = SSql + " AND ";
			}
			SSql = SSql + " Type = " + frmLog.cboType.options[frmLog.cboType.selectedIndex].value + " ";
		}

		if (frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value != -1) {
			if (SSql.length > 0) {
				SSql = SSql + " AND ";
			}
			SSql = SSql + "Status = " + frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value + " ";
		}

		if (frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value != -1) {
			if (SSql.length > 0) {
				SSql = SSql + " AND ";
			}
			SSql = SSql + " Mode = " + frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value + " ";
		}

		return SSql;
	}

	function refreshGrid() {
		var frmLog = OpenHR.getForm("workframe", "frmLog");
		var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
		frmGetDataForm.txtAction.value = "LOADEVENTLOG";
		
		frmGetDataForm.txtELFilterUser.value = frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value;
		frmGetDataForm.txtELFilterType.value = frmLog.cboType.options[frmLog.cboType.selectedIndex].value;
		frmGetDataForm.txtELFilterStatus.value = frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value;
		frmGetDataForm.txtELFilterMode.value = frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value;
		frmGetDataForm.txtELOrderColumn.value = frmLog.txtELOrderColumn.value;
		frmGetDataForm.txtELOrderOrder.value = frmLog.txtELOrderOrder.value;

		EventLog_refreshButtons();
		OpenHR.submitForm(frmGetDataForm);
	}

	function EventLog_viewEvent() {
		var sURL;
		var rowID = $("#LogEvents").jqGrid('getGridParam', 'selrow');
		var frmLog = OpenHR.getForm("workframe", "frmLog");

		if (rowID == null) { //No row selected
			return;
		}
		//Get the row data
		var rowData = $("#LogEvents").getRowData(rowID);

		frmDetails.txtEventID.value = rowData["ID"];
		frmDetails.txtEventName.value = rowData["Name"];
		frmDetails.txtEventMode.value = rowData["Mode"];
		frmDetails.txtEventStartTime.value = rowData["Start Time"];
		frmDetails.txtEventEndTime.value = rowData["End Time"];
		frmDetails.txtEventDuration.value = rowData["Duration"];
		frmDetails.txtEventType.value = rowData["Type"];
		frmDetails.txtEventStatus.value = rowData["Status"];
		frmDetails.txtEventUser.value = rowData["User name"];
		frmDetails.txtEventBatchRunID.value = rowData["Batch Run ID"];
		frmDetails.txtEmailPermission.value = frmLog.txtELEmailPermission.value;

		var postData = {
			ID: rowData["ID"],
			Mode: rowData["Mode"],
			BatchRunID: rowData["Batch Run ID"],
			<%:Html.AntiForgeryTokenForAjaxPost() %> };

		$('#EventLogViewDetails').dialog("open");
		OpenHR.submitForm(null, "EventLogViewDetails", null, postData, "EventLogDetails");

	}

	function EventLog_deleteEvent() {
		var sURL = "eventLogSelection";
		
		$('#EventLogDelete').data('sURLData', sURL);
		$('#EventLogDelete').dialog("open");

		EventLog_refreshButtons();
	}

	function EventLog_purgeEvent() {
		var sURL = "EventLogPurge";
		$('#EventLogPurge').data('sURLData', sURL);
		$('#EventLogPurge').dialog("open");
	}

	function EventLog_emailEvent() {
		var eventID;
		var sEventList = new String("");
		var sURL;
		var selectedRows = $("#LogEvents").jqGrid('getGridParam', 'selarrrow');

		//populate the txtSelectedIDs list
		for (var i = 0; i <= selectedRows.length - 1; i++) {
			var rowData = $("#LogEvents").getRowData(selectedRows[i]);
			eventID = rowData["ID"];
			sEventList = sEventList + eventID + ",";
		}

		var postData = {
			SelectedEventIDs: sEventList.substr(0, sEventList.length - 1),
			IsFromMain: 1,
			EmailOrderColumn: frmLog.txtELOrderColumn.value,
			EmailOrderOrder: frmLog.txtELOrderOrder.value,
			<%:Html.AntiForgeryTokenForAjaxPost() %>
		};

		$('#EventLogEmailSelect').dialog("open");
		OpenHR.submitForm(null, "EventLogEmailSelect", null, postData, "EmailSelection");
	}
	
	function EventLog_refreshButtons() {
		var frmLog = OpenHR.getForm("workframe", "frmLog");
		var logEventRowCount = $("#LogEvents").getGridParam('reccount') == undefined ? 0 : $("#LogEvents").getGridParam('reccount');
		var logEventSelectedRows = $('#LogEvents').jqGrid('getGridParam', 'selarrrow') == undefined ? 0 : $('#LogEvents').jqGrid('getGridParam', 'selarrrow').length;

		menu_toolbarEnableItem("mnutoolViewEventLogFind", (logEventRowCount > 0));
		menu_toolbarEnableItem("mnutoolPurgeEventLogFind", (frmLog.txtELPurgePermission.value == "1"));
		menu_toolbarEnableItem("mnutoolDeleteEventLogFind", ((logEventRowCount > 0) && (frmLog.txtELDeletePermission.value == "1")));
		menu_toolbarEnableItem("mnutoolEmailEventLogFind", ((logEventSelectedRows > 0) && (frmLog.txtELEmailPermission.value == "1")));
	}

	function refreshUsers() {
		// Get the columns/calcs for the current table selection.
		var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
		frmGetDataForm.txtAction.value = "LOADEVENTLOGUSERS";
		//    data_refreshData();
		OpenHR.submitForm(frmGetDataForm);

	}

	function loadEventLogUsers(pbViewAll, psCurrentFilterUser) {
		var i;
		var bFoundUser = false;
		var frmLog = OpenHR.getForm("workframe", "frmLog");

		if (pbViewAll == 1) {
			var OOptionAll = document.createElement("OPTION");
			OOptionAll.innerHTML = '&lt;All&gt;';
			OOptionAll.value = -1;
			frmLog.cboUsername.options.add(OOptionAll);

			var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
			var dataCollection = frmUtilDefForm.elements;

			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {
					var sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 16);
					if (sControlName == "txtEventLogUser_") {
						oOption = document.createElement("OPTION");
						
						oOption.innerHTML = dataCollection.item(i).value;
						oOption.value = dataCollection.item(i).value;
						frmLog.cboUsername.options.add(oOption);
						combo_disable(frmLog.cboUsername, false);

						if (psCurrentFilterUser == dataCollection.item(i).value) {
							bFoundUser = true;
							oOption.selected = true;
						}
					}
				}
			}

			if (psCurrentFilterUser == '-1' || psCurrentFilterUser == '' || !bFoundUser) {
				OOptionAll.selected = true;
			}

			EventLog_refreshButtons();

			// Get menu to refresh the menu.
			menu_refreshMenu();


		}
		else {
			combo_disable(frmLog.cboUsername, true);
			var oOption = document.createElement("OPTION");
			oOption.innerHTML = frmEventUseful.txtUserName.value;
			oOption.value = frmEventUseful.txtUserName.value;
			oOption.selected = true;
			frmLog.cboUsername.options.add(oOption);
		}

		refreshGrid();
	}
	
	function refreshStatusBar() {
		var sRecords;
		var sCaption;
		var frmData = OpenHR.getForm("dataframe", "frmData");
		var frmLog = OpenHR.getForm("workframe", "frmLog");

		sRecords = frmData.txtELTotalRecordCount.value;

		var iStartPosition = parseInt(frmData.txtELFirstRecPos.value);
		var iEndPosition = iStartPosition - 1 + parseInt(frmData.txtELCurrentRecCount.value);

		if (sRecords > 0) {
			sCaption = "Record(s): " + sRecords;
		} else {
			sCaption = "No Records";
		}

		if (frmLog.txtELViewAllPermission.value == 0) {
			sCaption = sCaption + "     [Viewing own entries only]";
		}

		//TODO: We don't have a record position indicator yet on the ribbon for this form

		menu_SetmnutoolEventLogRecordPositionCaption(sCaption);
		//Enable/disable navigation controls based on certain conditions

		var elFindRecords = Number('<%:Session("findRecords")%>');

		if (elFindRecords <= Number(sRecords)) {
			if (iStartPosition == 1) { //Disable first and previous
				menu_toolbarEnableItem("mnutoolFirstEventLogFind", false);
				menu_toolbarEnableItem("mnutoolPreviousEventLogFind", false);
				menu_toolbarEnableItem("mnutoolNextEventLogFind", true);
				menu_toolbarEnableItem("mnutoolLastEventLogFind", true);
			} else if (iEndPosition == sRecords) { //Disable next and last
				menu_toolbarEnableItem("mnutoolFirstEventLogFind", true);
				menu_toolbarEnableItem("mnutoolPreviousEventLogFind", true);
				menu_toolbarEnableItem("mnutoolNextEventLogFind", false);
				menu_toolbarEnableItem("mnutoolLastEventLogFind", false);
			} else { //Enable all
				menu_toolbarEnableItem("mnutoolFirstEventLogFind", true);
				menu_toolbarEnableItem("mnutoolPreviousEventLogFind", true);
				menu_toolbarEnableItem("mnutoolNextEventLogFind", true);
				menu_toolbarEnableItem("mnutoolLastEventLogFind", true);
			}
		} else { //Disable all
			menu_toolbarEnableItem("mnutoolFirstEventLogFind", false);
			menu_toolbarEnableItem("mnutoolPreviousEventLogFind", false);
			menu_toolbarEnableItem("mnutoolNextEventLogFind", false);
			menu_toolbarEnableItem("mnutoolLastEventLogFind", false);
		}

		return true;
	}

</script>

<div class="absolutefull">
	<div class="pageTitleDiv" style="margin-bottom: 15px; margin-top: 10px">
		<span class="pageTitle" id="EventLog_PageTitle">Event Log</span>
	</div>

	<fieldset style="border:0">
		<legend class="fontsmalltitle">Filter :</legend>
		<form id="frmLog">

			<table style="height: 100%; width: 100%; padding: 0 10px 10px 0" class="invisible">
				<tr style="height:10px">
					<td style="width: 7%">User name : 
					</td>
					<td>
						<select class="width100" id="cboUsername" name="cboUsername" onchange="refreshGrid();">
						</select>
					</td>
					<td style="width: 5%">Type :</td>
					<td>
						<select class="width100" id="cboType" name="cboType" onchange="refreshGrid();">
							<%
								If Session("CurrentType") = "-1" Then
									Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
								Else
									Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltCalandarReport.ToString Then
									Response.Write("											<option value=17 selected>Calendar Report" & vbCrLf)
								Else
									Response.Write("											<option value=17>Calendar Report" & vbCrLf)
								End If

								If Session("CurrentType") = EventLog_Type.eltCareerProgression.ToString Then
									Response.Write("											<option value=22 selected>Career Progression" & vbCrLf)
								Else
									Response.Write("											<option value=22>Career Progression" & vbCrLf)
								End If

								If Session("CurrentType") = EventLog_Type.eltCrossTab.ToString Then
									Response.Write("											<option value=1 selected>Cross Tab" & vbCrLf)
								Else
									Response.Write("											<option value=1>Cross Tab" & vbCrLf)
								End If

								If Session("CurrentType") = EventLog_Type.elt9GridBox.ToString Then
									Response.Write("											<option value=35 selected>9-Box Grid Report" & vbCrLf)
								Else
									Response.Write("											<option value=35>9-Box Grid Report" & vbCrLf)
								End If
								
								If Session("CurrentType") = EventLog_Type.eltCustomReport.ToString Then
									Response.Write("											<option value=2 selected>Custom Report" & vbCrLf)
								Else
									Response.Write("											<option value=2>Custom Report" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltDataTransfer.ToString Then
									Response.Write("											<option value=3 selected>Data Transfer" & vbCrLf)
								Else
									Response.Write("											<option value=3>Data Transfer" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltDiaryRebuild.ToString Then
									Response.Write("											<option value=11 selected>Diary Rebuild" & vbCrLf)
								Else
									Response.Write("											<option value=11>Diary Rebuild" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltEmailRebuild.ToString Then
									Response.Write("											<option value=12 selected>Email Rebuild" & vbCrLf)
								Else
									Response.Write("											<option value=12>Email Rebuild" & vbCrLf)
								End If

								If Session("CurrentType") = EventLog_Type.eltLabel.ToString Then
									Response.Write("											<option value=18 selected>Envelopes & Labels" & vbCrLf)
								Else
									Response.Write("											<option value=18>Envelopes & Labels" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltExport.ToString Then
									Response.Write("											<option value=4 selected>Export" & vbCrLf)
								Else
									Response.Write("											<option value=4>Export" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltGlobalAdd.ToString Then
									Response.Write("											<option value=5 selected>Global Add" & vbCrLf)
								Else
									Response.Write("											<option value=5>Global Add" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltGlobalDelete.ToString Then
									Response.Write("											<option value=6 selected>Global Delete" & vbCrLf)
								Else
									Response.Write("											<option value=6>Global Delete" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltGlobalUpdate.ToString Then
									Response.Write("											<option value=7 selected>Global Update" & vbCrLf)
								Else
									Response.Write("											<option value=7>Global Update" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltImport.ToString Then
									Response.Write("											<option value=8 selected>Import" & vbCrLf)
								Else
									Response.Write("											<option value=8>Import" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltMailMerge.ToString Then
									Response.Write("											<option value=9 selected>Mail Merge" & vbCrLf)
								Else
									Response.Write("											<option value=9>Mail Merge" & vbCrLf)
								End If

								If Session("CurrentType") = EventLog_Type.eltMatchReport.ToString Then
									Response.Write("											<option value=16 selected>Match Report" & vbCrLf)
								Else
									Response.Write("											<option value=16>Match Report" & vbCrLf)
								End If
	
								If Session("CurrentType") = EventLog_Type.eltRecordProfile.ToString Then
									Response.Write("											<option value=20 selected>Record Profile" & vbCrLf)
								Else
									Response.Write("											<option value=20>Record Profile" & vbCrLf)
								End If

								If Session("CurrentType") = EventLog_Type.eltStandardReport.ToString Then
									Response.Write("											<option value=13 selected>Standard Report" & vbCrLf)
								Else
									Response.Write("											<option value=13>Standard Report" & vbCrLf)
								End If

								If Session("CurrentType") = EventLog_Type.eltSuccessionPlanning.ToString Then
									Response.Write("											<option value=21 selected>Succession Planning" & vbCrLf)
								Else
									Response.Write("											<option value=21>Succession Planning" & vbCrLf)
								End If

								If Session("CurrentType") = EventLog_Type.eltSystemError.ToString Then
									Response.Write("											<option value=15 selected>System Error" & vbCrLf)
								Else
									Response.Write("											<option value=15>System Error" & vbCrLf)
								End If

								If Session("WF_Enabled") Then
									If Session("CurrentType") = EventLog_Type.eltWorkflowRebuild.ToString Then
										Response.Write("											<option value=25 selected>Workflow Rebuild" & vbCrLf)
									Else
										Response.Write("											<option value=25>Workflow Rebuild" & vbCrLf)
									End If
								End If
		
							%>
						</select>
					</td>
					<td style="width: 5%">Mode : 
					</td>
					<td>
						<select class="width100" id="cboMode" name="cboMode" onchange="refreshGrid();">
							<%
								If Session("CurrentMode") = "-1" Then
									Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
								Else
									Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
								End If
	
								If Session("CurrentMode") = EventLogMode.elsBatch.ToString Then
									Response.Write("											<option value=1 selected>Batch" & vbCrLf)
								Else
									Response.Write("											<option value=1>Batch" & vbCrLf)
								End If
	
								If Session("CurrentMode") = EventLogMode.elsManual.ToString Then
									Response.Write("											<option value=0 selected>Manual" & vbCrLf)
								Else
									Response.Write("											<option value=0>Manual" & vbCrLf)
								End If
								
								If Session("CurrentMode") = EventLogMode.elsPack.ToString Then
									Response.Write("											<option value=2 selected>Pack" & vbCrLf)
								Else
									Response.Write("											<option value=2>Pack" & vbCrLf)
								End If
							%>
						</select>
					</td>
					<td style="width: 5%">Status : 
					</td>
					<td>
						<select class="width100" id="cboStatus" name="cboStatus" onchange="refreshGrid();">
							<%	
								If Session("CurrentStatus") = "-1" Then
									Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
								Else
									Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
								End If
		
								If Session("CurrentStatus") = EventLog_Status.elsCancelled.ToString Then
									Response.Write("											<option value=1 selected>Cancelled" & vbCrLf)
								Else
									Response.Write("											<option value=1>Cancelled" & vbCrLf)
								End If
	
								If Session("CurrentStatus") = EventLog_Status.elsError.ToString Then
									Response.Write("											<option value=5 selected>Error" & vbCrLf)
								Else
									Response.Write("											<option value=5>Error" & vbCrLf)
								End If
	
								If Session("CurrentStatus") = EventLog_Status.elsFailed.ToString Then
									Response.Write("											<option value=2 selected>Failed" & vbCrLf)
								Else
									Response.Write("											<option value=2>Failed" & vbCrLf)
								End If
	
								If Session("CurrentStatus") = EventLog_Status.elsPending.ToString Then
									Response.Write("											<option value=0 selected>Pending" & vbCrLf)
								Else
									Response.Write("											<option value=0>Pending" & vbCrLf)
								End If
	
								If Session("CurrentStatus") = EventLog_Status.elsSkipped.ToString Then
									Response.Write("											<option value=4 selected>Skipped" & vbCrLf)
								Else
									Response.Write("											<option value=4>Skipped" & vbCrLf)
								End If
	
								If Session("CurrentStatus") = EventLog_Status.elsSuccessful.ToString  Then
									Response.Write("											<option value=3 selected>Successful" & vbCrLf)
								Else
									Response.Write("											<option value=3>Successful" & vbCrLf)
								End If
							%>
						</select>
					</td>
				</tr>
			</table>

			<input id="cmdView" class="btn" type="button" value="View..." name="cmdView">
			<input id="cmdDelete" class="btn" type="button" value="Delete..." name="cmdDelete">
			<input id="cmdPurge" class="btn" type="button" value="Purge..." name="cmdPurge">
			<input id="cmdEmail" class="button" type="button" value="Email..." name="cmdEmail">

			<input type='hidden' id="txtELDeletePermission" name="txtELDeletePermission">
			<input type='hidden' id="txtELViewAllPermission" name="txtELViewAllPermission">
			<input type='hidden' id="txtELPurgePermission" name="txtELPurgePermission">
			<input type='hidden' id="txtELEmailPermission" name="txtELEmailPermission">
			<input type='hidden' id="txtELOrderColumn" name="txtELOrderColumn" value='DateTime'>
			<input type='hidden' id="txtELOrderOrder" name="txtELOrderOrder" value='DESC'>
			<input type='hidden' id="txtELSortColumnIndex" name="txtELSortColumnIndex" value="1">
			<input type='hidden' id="txtELLoaded" name="txtELLoaded" value="0">
			<input type="hidden" id="txtCurrUserFilter" name="txtCurrUserFilter" value='<%=Session("CurrentUsername")%>'>
			<%=Html.AntiForgeryToken()%>
		</form>

		<div id="gridContainer" style="height: 450px">
			<table id='LogEvents'></table>
			<div id='pager-coldata'></div>
		</div>
	</fieldset>
</div>

<form id="frmDetails" name="frmDetails" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="txtEventID" name="txtEventID">

	<input type="hidden" id="txtEventName" name="txtEventName">
	<input type="hidden" id="txtEventMode" name="txtEventMode">

	<input type="hidden" id="txtEventStartTime" name="txtEventStartTime">
	<input type="hidden" id="txtEventEndTime" name="txtEventEndTime">
	<input type="hidden" id="txtEventDuration" name="txtEventDuration">

	<input type="hidden" id="txtEventType" name="txtEventType">
	<input type="hidden" id="txtEventStatus" name="txtEventStatus">
	<input type="hidden" id="txtEventUser" name="txtEventUser">

	<input type="hidden" id="txtEventSuccessCount" name="txtEventSuccessCount">
	<input type="hidden" id="txtEventFailCount" name="txtEventFailCount">

	<input type="hidden" id="txtEventBatchName" name="txtEventBatchName">
	<input type="hidden" id="txtEventBatchJobID" name="txtEventBatchJobID">
	<input type="hidden" id="txtEventBatchRunID" name="txtEventBatchRunID">

	<input type="hidden" id="txtEmailPermission" name="txtEmailPermission">
	<%=Html.AntiForgeryToken()%>
</form>

<form id="frmPurge" name="frmPurge" method="post" style="visibility: hidden; display: none" action="eventLog">
	<input type="hidden" id="txtDoesPurge" name="txtDoesPurge">
	<input type="hidden" id="txtPurgePeriod" name="txtPurgePeriod">
	<input type="hidden" id="txtPurgeFrequency" name="txtPurgeFrequency">
	<input type="hidden" id="txtShowPurgeMSG" name="txtShowPurgeMSG" value='<%=Session("showPurgeMessage")%>'>
	<input type="hidden" id="txtCurrentUsername" name="txtCurrentUsername">
	<input type="hidden" id="txtCurrentType" name="txtCurrentType">
	<input type="hidden" id="txtCurrentMode" name="txtCurrentMode">
	<input type="hidden" id="txtCurrentStatus" name="txtCurrentStatus">	
	<%=Html.AntiForgeryToken()%>
</form>

<form id="frmDelete" name="frmDelete" method="post" style="visibility: hidden; display: none" action="eventLog">
	<input type="hidden" id="txtDeleteSel" name="txtDeleteSel">
	<input type="hidden" id="txtSelectedIDs" name="txtSelectedIDs">
	<input type="hidden" id="txtViewAllPerm" name="txtViewAllPerm">
	<input type="hidden" id="txtCurrentUsername" name="txtCurrentUsername">
	<input type="hidden" id="txtCurrentType" name="txtCurrentType">
	<input type="hidden" id="txtCurrentMode" name="txtCurrentMode">
	<input type="hidden" id="txtCurrentStatus" name="txtCurrentStatus">
	<%=Html.AntiForgeryToken()%>
</form>

<form id="frmRefresh" name="frmRefresh" method="post" style="visibility: hidden; display: none" action="eventLog">
	<input type="hidden" id="txtEventExisted" name="txtEventExisted">
	<input type="hidden" id="txtCurrentUsername" name="txtCurrentUsername">
	<input type="hidden" id="txtCurrentType" name="txtCurrentType">
	<input type="hidden" id="txtCurrentMode" name="txtCurrentMode">
	<input type="hidden" id="txtCurrentStatus" name="txtCurrentStatus">
	<%=Html.AntiForgeryToken()%>
</form>

<form id="frmEventUseful" name="frmEventUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<%
		Dim sParameterValue As String = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")
		Response.Write("<input type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & sParameterValue & ">" & vbCrLf)
		
		Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value="""">" & vbCrLf)
		Response.Write("<input type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)

		Session("showPurgeMessage") = 0
		Session("CurrentUsername") = ""
		Session("CurrentType") = ""
		Session("CurrentMode") = ""
		Session("CurrentStatus") = ""
	%>	
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<script type="text/javascript">
	eventLog_window_onload();
</script>


