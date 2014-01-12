﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<%
	
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
	Dim SPParameters() As SqlParameter

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

	function eventLog_window_onload() {

		$("#workframe").attr("data-framesource", "EVENTLOG");

		setGridFont(frmLog.ssOleDBGridEventLog);

		var fOK;
		fOK = true;

		var sErrMsg = frmEventUseful.txtErrorDescription.value;
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
		
		$("#toolbarAdminConfig").parent().show();
		$("#toolbarAdminConfig").click();

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

	function refreshStatusBar() {
		
		var sRecords;
		var sCaption;
		var frmData = OpenHR.getForm("dataframe", "frmData");

		sRecords = frmData.txtELTotalRecordCount.value;

		var iStartPosition = parseInt(frmData.txtELFirstRecPos.value);
		var iEndPosition = iStartPosition - 1 + parseInt(frmData.txtELCurrentRecCount.value);
		
		if (sRecords > 0) {
			sCaption = "Records " +
					iStartPosition +
					" to " +
					iEndPosition +
					" of " +
					sRecords;
		}
		else {
			sCaption = "No Records";
		}

		if (frmLog.txtELViewAllPermission.value == 0) {
			sCaption = sCaption + "     [Viewing own entries only]";
		}
		
		//TODO: We don't have a record position indicator yet on the ribbon for this form
		
		menu_SetmnutoolRecordPositionCaption(sCaption);

		//Enable/disable navigation controls based on certain conditions
		if (sRecords <= 1000) { //TODO set this to blocksize...
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
			}
			else { //Enable all
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

	function loadEventLog() {
		var i;
		var iPollCounter;
		var iPollPeriod;
		var sControlName;
		var sControlPrefix;

		iPollPeriod = 100;
		iPollCounter = iPollPeriod;

		var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
		var dataCollection = frmUtilDefForm.elements;

		var frmLog = OpenHR.getForm("workframe", "frmLog");
		
		if (dataCollection != null) {
			frmLog.ssOleDBGridEventLog.focus();
			frmLog.ssOleDBGridEventLog.Redraw = false;
			if (frmLog.ssOleDBGridEventLog.Rows > 0) {
				frmLog.ssOleDBGridEventLog.RemoveAll();
			}

			for (i = 0; i < dataCollection.length; i++) {

				if (i == iPollCounter) {
					//TODO
					//frmRefresh.submit();
					iPollCounter = iPollCounter + iPollPeriod;
				}

				sControlName = dataCollection.item(i).name;
				sControlPrefix = sControlName.substr(0, 13);

				if (sControlPrefix == "txtAddString_") {
					frmLog.ssOleDBGridEventLog.AddItem(dataCollection.item(i).value);
				}
			}

			frmLog.ssOleDBGridEventLog.Redraw = true;
			//TODO
			//frmRefresh.submit();

			if (frmLog.ssOleDBGridEventLog.Rows > 0) {
				frmLog.ssOleDBGridEventLog.SelBookmarks.RemoveAll();
				frmLog.ssOleDBGridEventLog.MoveFirst();
				frmLog.ssOleDBGridEventLog.SelBookmarks.Add(frmLog.ssOleDBGridEventLog.Bookmark);
			}
		}

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

		if (frmPurge.txtShowPurgeMSG.value == 1) {
			OpenHR.messageBox("Purge completed.", 64, "Event Log");
			frmPurge.txtShowPurgeMSG.value = 0;
		}
	}

	function filterSQL() {
		var sSQL = new String("");

		if (frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value != -1) {
			var sUsername = new String(frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value);
			sSQL = sSQL + " LOWER(Username) = '" + sUsername.toLowerCase() + "' ";
		}

		if (frmLog.cboType.options[frmLog.cboType.selectedIndex].value != -1) {
			if (sSQL.length > 0) {
				sSQL = sSQL + " AND ";
			}
			sSQL = sSQL + " Type = " + frmLog.cboType.options[frmLog.cboType.selectedIndex].value + " ";
		}

		if (frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value != -1) {
			if (sSQL.length > 0) {
				sSQL = sSQL + " AND ";
			}
			sSQL = sSQL + "Status = " + frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value + " ";
		}

		if (frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value != -1) {
			if (sSQL.length > 0) {
				sSQL = sSQL + " AND ";
			}
			sSQL = sSQL + " Mode = " + frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value + " ";
		}

		return sSQL;
	}

	function refreshGrid() {

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

		
		if (frmLog.ssOleDBGridEventLog.Rows > 0 && frmLog.ssOleDBGridEventLog.SelBookmarks.Count == 1) {
			frmDetails.txtEventID.value = frmLog.ssOleDBGridEventLog.Columns(0).text;

			frmDetails.txtEventName.value = frmLog.ssOleDBGridEventLog.Columns(5).text;
			frmDetails.txtEventMode.value = frmLog.ssOleDBGridEventLog.Columns(7).text;

			frmDetails.txtEventStartTime.value = frmLog.ssOleDBGridEventLog.Columns(1).text;
			frmDetails.txtEventEndTime.value = frmLog.ssOleDBGridEventLog.Columns(2).text;
			frmDetails.txtEventDuration.value = frmLog.ssOleDBGridEventLog.Columns(3).text;

			frmDetails.txtEventType.value = frmLog.ssOleDBGridEventLog.Columns(4).text;
			frmDetails.txtEventStatus.value = frmLog.ssOleDBGridEventLog.Columns(6).text;
			frmDetails.txtEventUser.value = frmLog.ssOleDBGridEventLog.Columns(8).text;

			frmDetails.txtEventSuccessCount.value = frmLog.ssOleDBGridEventLog.Columns(12).text;
			frmDetails.txtEventFailCount.value = frmLog.ssOleDBGridEventLog.Columns(13).text;

			frmDetails.txtEventBatchName.value = frmLog.ssOleDBGridEventLog.Columns("BatchName").text;
			frmDetails.txtEventBatchJobID.value = frmLog.ssOleDBGridEventLog.Columns("BatchJobID").text;
			frmDetails.txtEventBatchRunID.value = frmLog.ssOleDBGridEventLog.Columns("BatchRunID").text;

			frmDetails.txtEmailPermission.value = frmLog.txtELEmailPermission.value;

			sURL = "eventLogDetails" +
					"?txtEventID=" + frmDetails.txtEventID.value +
					"&txtEventName=" + escape(frmDetails.txtEventName.value) +
					"&txtEventMode=" + escape(frmDetails.txtEventMode.value) +
					"&txtEventStartTime=" + frmDetails.txtEventStartTime.value +
					"&txtEventEndTime=" + frmDetails.txtEventEndTime.value +
					"&txtEventDuration=" + frmDetails.txtEventDuration.value +
					"&txtEventType=" + escape(frmDetails.txtEventType.value) +
					"&txtEventStatus=" + escape(frmDetails.txtEventStatus.value) +
					"&txtEventUser=" + escape(frmDetails.txtEventUser.value) +
					"&txtEventSuccessCount=" + frmDetails.txtEventSuccessCount.value +
					"&txtEventFailCount=" + frmDetails.txtEventFailCount.value +
					"&txtEventBatchName=" + escape(frmDetails.txtEventBatchName.value) +
					"&txtEventBatchJobID=" + frmDetails.txtEventBatchJobID.value +
					"&txtEventBatchRunID=" + frmDetails.txtEventBatchRunID.value +
					"&txtEmailPermission=" + escape(frmDetails.txtEmailPermission.value);

			openDialog(sURL, 900, 770);
		}

		EventLog_refreshButtons();
	}

	function EventLog_deleteEvent() {
		var sURL;

		sURL = "eventLogSelection" +
				"?txtEventID=" + frmDetails.txtEventID.value +
				"&txtEventName=" + escape(frmDetails.txtEventName.value) +
				"&txtEventMode=" + escape(frmDetails.txtEventMode.value) +
				"&txtEventStartTime=" + frmDetails.txtEventStartTime.value +
				"&txtEventEndTime=" + frmDetails.txtEventEndTime.value +
				"&txtEventDuration=" + frmDetails.txtEventDuration.value +
				"&txtEventType=" + escape(frmDetails.txtEventType.value) +
				"&txtEventStatus=" + escape(frmDetails.txtEventStatus.value) +
				"&txtEventUser=" + escape(frmDetails.txtEventUser.value) +
				"&txtEventSuccessCount=" + frmDetails.txtEventSuccessCount.value +
				"&txtEventFailCount=" + frmDetails.txtEventFailCount.value +
				"&txtEventBatchName=" + escape(frmDetails.txtEventBatchName.value) +
				"&txtEventBatchJobID=" + frmDetails.txtEventBatchJobID.value +
				"&txtEventBatchRunID=" + frmDetails.txtEventBatchRunID.value +
				"&txtEmailPermission=" + escape(frmDetails.txtEmailPermission.value);

		openDialog(sURL, 500, 220);
	}

	function EventLog_purgeEvent() {
		var sURL;

		sURL = "EventLogPurge" +
				"?txtEventID=" + frmDetails.txtEventID.value +
				"&txtEventName=" + escape(frmDetails.txtEventName.value) +
				"&txtEventMode=" + escape(frmDetails.txtEventMode.value) +
				"&txtEventStartTime=" + frmDetails.txtEventStartTime.value +
				"&txtEventEndTime=" + frmDetails.txtEventEndTime.value +
				"&txtEventDuration=" + frmDetails.txtEventDuration.value +
				"&txtEventType=" + escape(frmDetails.txtEventType.value) +
				"&txtEventStatus=" + escape(frmDetails.txtEventStatus.value) +
				"&txtEventUser=" + escape(frmDetails.txtEventUser.value) +
				"&txtEventSuccessCount=" + frmDetails.txtEventSuccessCount.value +
				"&txtEventFailCount=" + frmDetails.txtEventFailCount.value +
				"&txtEventBatchName=" + escape(frmDetails.txtEventBatchName.value) +
				"&txtEventBatchJobID=" + frmDetails.txtEventBatchJobID.value +
				"&txtEventBatchRunID=" + frmDetails.txtEventBatchRunID.value +
				"&txtEmailPermission=" + escape(frmDetails.txtEmailPermission.value);

		openDialog(sURL, 600, 280);

	}

	function EventLog_emailEvent() {
		var eventID;
		var sEventList = new String("");
		var sURL;

		//populate the txtSelectedIDs list
		for (var i = 0; i < frmLog.ssOleDBGridEventLog.SelBookmarks.Count; i++) {
			eventID = frmLog.ssOleDBGridEventLog.Columns("ID").CellText(frmLog.ssOleDBGridEventLog.SelBookmarks(i));

			sEventList = sEventList + eventID + ",";
		}

		frmEmail.txtSelectedEventIDs.value = sEventList.substr(0, sEventList.length - 1);

		sURL = "emailSelection" +
				"?txtSelectedEventIDs=" + frmEmail.txtSelectedEventIDs.value +
				"&txtFromMain=" + frmEmail.txtFromMain.value +
				"&txtEmailOrderColumn=" + frmLog.txtELOrderColumn.value +
				"&txtEmailOrderOrder=" + frmLog.txtELOrderOrder.value;

		openDialog(sURL, 500, 400);
	}

	function EventLog_refreshButtons() {

		var frmLog = OpenHR.getForm("workframe", "frmLog");
		menu_toolbarEnableItem("mnutoolViewEventLogFind", (frmLog.ssOleDBGridEventLog.Rows > 0));
		menu_toolbarEnableItem("mnutoolPurgeEventLogFind", (frmLog.txtELPurgePermission.value == "1"));
		menu_toolbarEnableItem("mnutoolDeleteEventLogFind", ((frmLog.ssOleDBGridEventLog.Rows > 0) && (frmLog.txtELDeletePermission.value == "1")));
		menu_toolbarEnableItem("mnutoolEmailEventLogFind", ((frmLog.ssOleDBGridEventLog.SelBookmarks.Count > 0) && (frmLog.txtELEmailPermission.value == "1")));

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

		if (pbViewAll == 1) {
			var oOptionALL = document.createElement("OPTION");
			frmLog.cboUsername.options.add(oOptionALL);
			oOptionALL.innerText = '<All>';
			oOptionALL.value = -1;

			var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
			var dataCollection = frmUtilDefForm.elements;

			if (dataCollection != null) {
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 16);
					if (sControlName == "txtEventLogUser_") {
						var oOption = document.createElement("OPTION");
						frmLog.cboUsername.options.add(oOption);
						oOption.innerText = dataCollection.item(i).value;
						oOption.value = dataCollection.item(i).value;
						combo_disable(frmLog.cboUsername, false);

						if (psCurrentFilterUser == dataCollection.item(i).value) {
							bFoundUser = true;
							oOption.selected = true;
						}
					}
				}
			}

			if (psCurrentFilterUser == '-1' || psCurrentFilterUser == '' || !bFoundUser) {
				oOptionALL.selected = true;
			}

			// Get menu to refresh the menu.
			menu_refreshMenu();


		}
		else {
			combo_disable(frmLog.cboUsername, true);
			var oOption = document.createElement("OPTION");
			frmLog.cboUsername.options.add(oOption);
			oOption.innerText = frmEventUseful.txtUserName.value;
			oOption.value = oOption.innerText;
			oOption.selected = true;
		}

		EventLog_refreshButtons();

		refreshGrid();
	}

	function openDialog(pDestination, pWidth, pHeight) {

		dlgwinprops = "center:yes;" +
				"dialogHeight:" + pHeight + "px;" +
				"dialogWidth:" + pWidth + "px;" +
				"help:no;" +
				"resizable:no;" +
				"scroll:no;" +
				"status:no;";
		window.showModalDialog(pDestination, self, dlgwinprops);
	}

</script>

<object classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
	id="dialog"
	codebase="cabs/comdlg32.cab#Version=1,0,0,0"
	style="LEFT: 0px; TOP: 0px">
	<param name="_ExtentX" value="847">
	<param name="_ExtentY" value="847">
	<param name="_Version" value="393216">
	<param name="CancelError" value="0">
	<param name="Color" value="0">
	<param name="Copies" value="1">
	<param name="DefaultExt" value="">
	<param name="DialogTitle" value="bob">
	<param name="FileName" value="">
	<param name="Filter" value="">
	<param name="FilterIndex" value="0">
	<param name="Flags" value="0">
	<param name="FontBold" value="0">
	<param name="FontItalic" value="0">
	<param name="FontName" value="">
	<param name="FontSize" value="8">
	<param name="FontStrikeThru" value="0">
	<param name="FontUnderLine" value="0">
	<param name="FromPage" value="0">
	<param name="HelpCommand" value="0">
	<param name="HelpContext" value="0">
	<param name="HelpFile" value="">
	<param name="HelpKey" value="">
	<param name="InitDir" value="">
	<param name="Max" value="0">
	<param name="Min" value="0">
	<param name="MaxFileSize" value="260">
	<param name="PrinterDefault" value="1">
	<param name="ToPage" value="0">
	<param name="Orientation" value="1">
</object>

<form id="frmLog">
	<table align="center" cellpadding="5" cellspacing="0" width="100%" height="100%">
		<tr>
			<tr>
				<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
					<tr height="5">
						<td colspan="3"></td>
					</tr>

					<tr>
						<td width="5"></td>
						<td>
							<table width="100%" height="100%" cellspacing="0" cellpadding="5">
								<tr valign="top">
									<td>
										<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="4">
											<tr height="10">
												<td colspan="8">Filters : 
												</td>
											</tr>
											<tr height="10">
												<td width="82" nowrap>User name : 
												</td>
												<td>
													<select id="cboUsername" name="cboUsername" class="combo" style="WIDTH: 100%" onchange="refreshGrid();">
													</select>
												</td>
												<td width="25">Type : 
												</td>
												<td>
													<select id="cboType" name="cboType" class="combo" style="WIDTH: 100%" onchange="refreshGrid();">

														<%
															If Session("CurrentType") = "-1" Then
																Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
															Else
																Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
															End If
	
															If Session("CurrentType") = "17" Then
																Response.Write("											<option value=17 selected>Calendar Report" & vbCrLf)
															Else
																Response.Write("											<option value=17>Calendar Report" & vbCrLf)
															End If

															If Session("CurrentType") = "22" Then
																Response.Write("											<option value=22 selected>Career Progression" & vbCrLf)
															Else
																Response.Write("											<option value=22>Career Progression" & vbCrLf)
															End If

															If Session("CurrentType") = "1" Then
																Response.Write("											<option value=1 selected>Cross Tab" & vbCrLf)
															Else
																Response.Write("											<option value=1>Cross Tab" & vbCrLf)
															End If
	
															If Session("CurrentType") = "2" Then
																Response.Write("											<option value=2 selected>Custom Report" & vbCrLf)
															Else
																Response.Write("											<option value=2>Custom Report" & vbCrLf)
															End If
	
															If Session("CurrentType") = "3" Then
																Response.Write("											<option value=3 selected>Data Transfer" & vbCrLf)
															Else
																Response.Write("											<option value=3>Data Transfer" & vbCrLf)
															End If
	
															If Session("CurrentType") = "11" Then
																Response.Write("											<option value=11 selected>Diary Rebuild" & vbCrLf)
															Else
																Response.Write("											<option value=11>Diary Rebuild" & vbCrLf)
															End If
	
															If Session("CurrentType") = "12" Then
																Response.Write("											<option value=12 selected>Email Rebuild" & vbCrLf)
															Else
																Response.Write("											<option value=12>Email Rebuild" & vbCrLf)
															End If

															If Session("CurrentType") = "18" Then
																Response.Write("											<option value=18 selected>Envelopes & Labels" & vbCrLf)
															Else
																Response.Write("											<option value=18>Envelopes & Labels" & vbCrLf)
															End If
	
															If Session("CurrentType") = "4" Then
																Response.Write("											<option value=4 selected>Export" & vbCrLf)
															Else
																Response.Write("											<option value=4>Export" & vbCrLf)
															End If
	
															If Session("CurrentType") = "5" Then
																Response.Write("											<option value=5 selected>Global Add" & vbCrLf)
															Else
																Response.Write("											<option value=5>Global Add" & vbCrLf)
															End If
	
															If Session("CurrentType") = "6" Then
																Response.Write("											<option value=6 selected>Global Delete" & vbCrLf)
															Else
																Response.Write("											<option value=6>Global Delete" & vbCrLf)
															End If
	
															If Session("CurrentType") = "7" Then
																Response.Write("											<option value=7 selected>Global Update" & vbCrLf)
															Else
																Response.Write("											<option value=7>Global Update" & vbCrLf)
															End If
	
															If Session("CurrentType") = "8" Then
																Response.Write("											<option value=8 selected>Import" & vbCrLf)
															Else
																Response.Write("											<option value=8>Import" & vbCrLf)
															End If

															'if Session("CurrentType") = "19" then
															'	Response.Write "											<option value=19 selected>Label Definition" & vbCrLf
															'else
															'	Response.Write "											<option value=19>Label Definition" & vbCrLf
															'end if
	
															If Session("CurrentType") = "9" Then
																Response.Write("											<option value=9 selected>Mail Merge" & vbCrLf)
															Else
																Response.Write("											<option value=9>Mail Merge" & vbCrLf)
															End If

															If Session("CurrentType") = "16" Then
																Response.Write("											<option value=16 selected>Match Report" & vbCrLf)
															Else
																Response.Write("											<option value=16>Match Report" & vbCrLf)
															End If
	
															If Session("CurrentType") = "20" Then
																Response.Write("											<option value=20 selected>Record Profile" & vbCrLf)
															Else
																Response.Write("											<option value=20>Record Profile" & vbCrLf)
															End If

															If Session("CurrentType") = "13" Then
																Response.Write("											<option value=13 selected>Standard Report" & vbCrLf)
															Else
																Response.Write("											<option value=13>Standard Report" & vbCrLf)
															End If

															If Session("CurrentType") = "21" Then
																Response.Write("											<option value=21 selected>Succession Planning" & vbCrLf)
															Else
																Response.Write("											<option value=21>Succession Planning" & vbCrLf)
															End If

															If Session("CurrentType") = "15" Then
																Response.Write("											<option value=15 selected>System Error" & vbCrLf)
															Else
																Response.Write("											<option value=15>System Error" & vbCrLf)
															End If

															If Session("WF_Enabled") Then
																If Session("CurrentType") = "25" Then
																	Response.Write("											<option value=25 selected>Workflow Rebuild" & vbCrLf)
																Else
																	Response.Write("											<option value=25>Workflow Rebuild" & vbCrLf)
																End If
															End If
		
														%>
													</select>
												</td>
												<td width="25">Mode : 
												</td>
												<td>
													<select id="cboMode" name="cboMode" class="combo" style="WIDTH: 100%" onchange="refreshGrid();">
														<%
															If Session("CurrentMode") = "-1" Then
																Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
															Else
																Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
															End If
	
															If Session("CurrentMode") = "1" Then
																Response.Write("											<option value=1 selected>Batch" & vbCrLf)
															Else
																Response.Write("											<option value=1>Batch" & vbCrLf)
															End If
	
															If Session("CurrentMode") = "0" Then
																Response.Write("											<option value=0 selected>Manual" & vbCrLf)
															Else
																Response.Write("											<option value=0>Manual" & vbCrLf)
															End If
														%>
													</select>
												</td>
												<td width="25">Status : 
												</td>
												<td>
													<select id="cboStatus" name="cboStatus" class="combo" style="width: 100%" onchange="refreshGrid();">
														<%	
															If Session("CurrentStatus") = "-1" Then
																Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
															Else
																Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
															End If
		
															If Session("CurrentStatus") = "1" Then
																Response.Write("											<option value=1 selected>Cancelled" & vbCrLf)
															Else
																Response.Write("											<option value=1>Cancelled" & vbCrLf)
															End If
	
															If Session("CurrentStatus") = "5" Then
																Response.Write("											<option value=5 selected>Error" & vbCrLf)
															Else
																Response.Write("											<option value=5>Error" & vbCrLf)
															End If
	
															If Session("CurrentStatus") = "2" Then
																Response.Write("											<option value=2 selected>Failed" & vbCrLf)
															Else
																Response.Write("											<option value=2>Failed" & vbCrLf)
															End If
	
															If Session("CurrentStatus") = "0" Then
																Response.Write("											<option value=0 selected>Pending" & vbCrLf)
															Else
																Response.Write("											<option value=0>Pending" & vbCrLf)
															End If
	
															If Session("CurrentStatus") = "4" Then
																Response.Write("											<option value=4 selected>Skipped" & vbCrLf)
															Else
																Response.Write("											<option value=4>Skipped" & vbCrLf)
															End If
	
															If Session("CurrentStatus") = "3" Then
																Response.Write("											<option value=3 selected>Successful" & vbCrLf)
															Else
																Response.Write("											<option value=3>Successful" & vbCrLf)
															End If
														%>
													</select>
												</td>
											</tr>
											<tr height="5">
												<td colspan="8"></td>
											</tr>
											<tr>
												<td colspan="8">
													<%

														Dim avColumnDef(13, 4)
	
														avColumnDef(0, 0) = "ID"			 'name
														avColumnDef(0, 1) = "ID"			 'caption
														avColumnDef(0, 2) = "1600"		 'width
														avColumnDef(0, 3) = "0"				 'visible
	
														avColumnDef(1, 0) = "DateTime" 'name
														avColumnDef(1, 1) = "Start Time"	 'caption
														avColumnDef(1, 2) = "3300"				 'width
														avColumnDef(1, 3) = "-1"					 'visible

														avColumnDef(2, 0) = "EndTime"	 'name
														avColumnDef(2, 1) = "End Time" 'caption
														avColumnDef(2, 2) = "3300"		 'width
														avColumnDef(2, 3) = "-1"			 'visible
	
														avColumnDef(3, 0) = "Duration" 'name
														avColumnDef(3, 1) = "Duration" 'caption
														avColumnDef(3, 2) = "1750"		 'width
														avColumnDef(3, 3) = "-1"			 'visible

														avColumnDef(4, 0) = "Type"	 'name
														avColumnDef(4, 1) = "Type"	 'caption
														avColumnDef(4, 2) = "3250"	 'width
														avColumnDef(4, 3) = "-1"		 'visible

														avColumnDef(5, 0) = "Name"	 'name
														avColumnDef(5, 1) = "Name"	 'caption
														avColumnDef(5, 2) = "5500"	 'width
														avColumnDef(5, 3) = "-1"		 'visible

														avColumnDef(6, 0) = "Status"	 'name
														avColumnDef(6, 1) = "Status"	 'caption
														avColumnDef(6, 2) = "2100"		 'width
														avColumnDef(6, 3) = "-1"			 'visible

														avColumnDef(7, 0) = "Mode"		 'name
														avColumnDef(7, 1) = "Mode"		 'caption
														avColumnDef(7, 2) = "1500"		 'width
														avColumnDef(7, 3) = "-1"			 'visible

														avColumnDef(8, 0) = "Username"	 'name
														avColumnDef(8, 1) = "User name"	 'caption
														avColumnDef(8, 2) = "2500"			 'width
														avColumnDef(8, 3) = "-1"				 'visible

														avColumnDef(9, 0) = "BatchJobID" 'name
														avColumnDef(9, 1) = "BatchJobID" 'caption
														avColumnDef(9, 2) = "1800"			 'width
														avColumnDef(9, 3) = "0"					 'visible
	
														avColumnDef(10, 0) = "BatchRunID"	 'name
														avColumnDef(10, 1) = "BatchRunID"	 'caption
														avColumnDef(10, 2) = "1800"				 'width
														avColumnDef(10, 3) = "0"					 'visible

														avColumnDef(11, 0) = "BatchName"	 'name
														avColumnDef(11, 1) = "Batch Name"	 'caption
														avColumnDef(11, 2) = "1800"				 'width
														avColumnDef(11, 3) = "0"					 'visible
	
														avColumnDef(12, 0) = "SuccessCount"	 'name
														avColumnDef(12, 1) = "SuccessCount"	 'caption
														avColumnDef(12, 2) = "1800"					 'width
														avColumnDef(12, 3) = "0"						 'visible
	
														avColumnDef(13, 0) = "FailCount"	 'name
														avColumnDef(13, 1) = "FailCount"	 'caption
														avColumnDef(13, 2) = "1800"				 'width
														avColumnDef(13, 3) = "0"					 'visible
		
														Response.Write("											<OBJECT classid=clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" & vbCrLf)
														Response.Write("													 codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6""" & vbCrLf)
														'Response.Write("													height=""100%""" & vbCrLf)
														Response.Write("													id=ssOleDBGridEventLog" & vbCrLf)
														Response.Write("													name=ssOleDBGridEventLog" & vbCrLf)
														Response.Write("													style=""HEIGHT: 400px; VISIBILITY: visible; WIDTH: 100%""" & vbCrLf)
														'Response.Write("													width=""100%"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""ScrollBars"" VALUE=""3"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""GroupHeaders"" VALUE=""-1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BevelColorFrame"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BevelColorHighlight"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BevelColorShadow"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BevelColorFace"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
														Response.Write("												<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""SelectTypeRow"" VALUE=""3"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""RowNavigation"" VALUE=""2"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
														Response.Write("												<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
														Response.Write("												<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BackColorEven"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BackColorOdd"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
														Response.Write("												<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Columns.Count"" VALUE=""" & (UBound(avColumnDef) + 1) & """>" & vbCrLf)
	
														For i = 0 To UBound(avColumnDef) Step 1
															Response.Write("												<!--" & avColumnDef(i, 0) & "-->  " & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Width"" VALUE=""" & avColumnDef(i, 2) & """>" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Visible"" VALUE=""" & avColumnDef(i, 3) & """>" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Caption"" VALUE=""" & avColumnDef(i, 1) & """>" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Name"" VALUE=""" & avColumnDef(i, 0) & """>" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Alignment"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Bound"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").DataField"" VALUE=""Column " & i & """>" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").DataType"" VALUE=""8"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Level"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").NumberFormat"" VALUE="""">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Case"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").FieldLen"" VALUE=""256"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Locked"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Style"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").RowCount"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").ColCount"" VALUE=""1"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").ForeColor"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").BackColor"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").StyleSet"" VALUE="""">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Nullable"" VALUE=""1"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").Mask"" VALUE="""">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").ClipMode"" VALUE=""0"">" & vbCrLf)
															Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptChar"" VALUE=""95"">" & vbCrLf)
														Next
		
														Response.Write("												<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""_ExtentX"" VALUE=""11298"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""_ExtentY"" VALUE=""3969"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
														Response.Write("												<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""BackColor"" VALUE=""0"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
														Response.Write("												<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

														Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
														Response.Write("											</OBJECT>" & vbCrLf)
													%>											
												</td>
											</tr>
										</table>
									</td>
									<td width="80">
										<table width="100%" class="invisible" cellspacing="0" cellpadding="2">
											<tr height="30">
												<td></td>
											</tr>
											<tr>
												<td width="10">
													<input id="cmdView" class="btn" type="button" value="View..." name="cmdView" style="WIDTH: 80px" width="80"
												</td>
											</tr>
											<tr height="10">
												<td></td>
											</tr>
											<tr>
												<td width="10">
													<input id="cmdDelete" class="btn" type="button" value="Delete..." name="cmdDelete" style="WIDTH: 80px" width="80"
												</td>
											</tr>
											<tr height="10">
												<td></td>
											</tr>
											<tr>
												<td width="10">
													<input id="cmdPurge" class="btn" type="button" value="Purge..." name="cmdPurge" style="WIDTH: 80px" width="80"
												</td>
											</tr>
											<tr height="10">
												<td></td>
											</tr>
											<tr>
												<td width="10">
													<input id="cmdEmail" class="button" type="button" value="Email..." name="cmdEmail" style="WIDTH: 80px" width="80"
												</td>
											</tr>
										</table>
									</td>
								</tr>

							</table>
						</td>
						<td width="5"></td>
					</tr>
					<tr height="8">
						<td width="5"></td>
						<tr colspan="1">
							<table width="100%" class="invisible" cellspacing="0" cellpadding="1">
								<tr>
									<td name="sbEventLog" id="sbEventLog">&nbsp
									</td>
						</tr>
				</table>
			</tr>
			<td width="5"></td>
		</tr>
	</table>
	</tr> 
</TABLE>

		<input type='hidden' id="txtELDeletePermission" name="txtELDeletePermission">
	<input type='hidden' id="txtELViewAllPermission" name="txtELViewAllPermission">
	<input type='hidden' id="txtELPurgePermission" name="txtELPurgePermission">
	<input type='hidden' id="txtELEmailPermission" name="txtELEmailPermission">
	<input type='hidden' id="txtELOrderColumn" name="txtELOrderColumn" value='DateTime'>
	<input type='hidden' id="txtELOrderOrder" name="txtELOrderOrder" value='DESC'>
	<input type='hidden' id="txtELSortColumnIndex" name="txtELSortColumnIndex" value="1">
	<input type='hidden' id="txtELLoaded" name="txtELLoaded" value="0">
	<input type="hidden" id="txtCurrUserFilter" name="txtCurrUserFilter" value='<%=Session("CurrentUsername")%>'>
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

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
</form>

<form id="frmDelete" name="frmDelete" method="post" style="visibility: hidden; display: none" action="eventLog">
	<input type="hidden" id="txtDeleteSel" name="txtDeleteSel">
	<input type="hidden" id="txtSelectedIDs" name="txtSelectedIDs">
	<input type="hidden" id="txtViewAllPerm" name="txtViewAllPerm">
	<input type="hidden" id="txtCurrentUsername" name="txtCurrentUsername">
	<input type="hidden" id="txtCurrentType" name="txtCurrentType">
	<input type="hidden" id="txtCurrentMode" name="txtCurrentMode">
	<input type="hidden" id="txtCurrentStatus" name="txtCurrentStatus">
</form>

<form id="frmEmail" name="frmEmail" method="post" style="visibility: hidden; display: none" action="emailSelection">
	<input type="hidden" id="txtSelectedEventIDs" name="txtSelectedEventIDs">
	<input type="hidden" id="txtFromMain" name="txtFromMain" value="1">
	<input type="hidden" id="txtEmailOrderColumn" name="txtEmailOrderColumn">
	<input type="hidden" id="txtEmailOrderOrder" name="txtEmailOrderOrder">
</form>

<form id="frmRefresh" name="frmRefresh" method="post" style="visibility: hidden; display: none" action="eventLog">
	<input type="hidden" id="txtEventExisted" name="txtEventExisted">
	<input type="hidden" id="txtCurrentUsername" name="txtCurrentUsername">
	<input type="hidden" id="txtCurrentType" name="txtCurrentType">
	<input type="hidden" id="txtCurrentMode" name="txtCurrentMode">
	<input type="hidden" id="txtCurrentStatus" name="txtCurrentStatus">
</form>

<%
	Session("showPurgeMessage") = 0
%>

<form id="frmEventUseful" name="frmEventUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<%
		Dim cmdDefinition As Command
		Dim prmModuleKey As ADODB.Parameter
		Dim prmParameterKey As ADODB.Parameter
		Dim prmParameterValue As ADODB.Parameter
		Dim sErrorDescription As String
		
		cmdDefinition = New Command
		cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
		cmdDefinition.CommandType = CommandTypeEnum.adCmdStoredProc
		cmdDefinition.ActiveConnection = Session("databaseConnection")

		prmModuleKey = cmdDefinition.CreateParameter("moduleKey", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 8000)
		cmdDefinition.Parameters.Append(prmModuleKey)
		prmModuleKey.value = "MODULE_PERSONNEL"

		prmParameterKey = cmdDefinition.CreateParameter("paramKey", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 8000)
		cmdDefinition.Parameters.Append(prmParameterKey)
		prmParameterKey.value = "Param_TablePersonnel"

		prmParameterValue = cmdDefinition.CreateParameter("paramValue", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamOutput, 8000)
		cmdDefinition.Parameters.Append(prmParameterValue)

		Err.Clear()
		cmdDefinition.Execute()

		Response.Write("<input type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").Value & ">" & vbCrLf)
	
		cmdDefinition = Nothing

		Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
		Response.Write("<input type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
	%>
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<%
	Session("CurrentUsername") = ""
	Session("CurrentType") = ""
	Session("CurrentMode") = ""
	Session("CurrentStatus") = ""
%>


<script type="text/javascript">

	function eventlog_addActiveXHandlers() {
		OpenHR.addActiveXHandler("ssOleDBGridEventLog", "DblClick", "ssOleDBGridEventLog_dblclick()");
		OpenHR.addActiveXHandler("ssOleDBGridEventLog", "rowcolchange", "ssOleDBGridEventLog_rowcolchange()");
		OpenHR.addActiveXHandler("ssOleDBGridEventLog", "Click", "ssOleDBGridEventLog_click()");
		OpenHR.addActiveXHandler("ssOleDBGridEventLog", "HeadClick", "ssOleDBGridEventLog_headclick()");
	}

	function ssOleDBGridEventLog_dblclick() {
		if ((frmLog.ssOleDBGridEventLog.Rows > 0) && (frmLog.ssOleDBGridEventLog.SelBookmarks.Count == 1)) {
			EventLog_viewEvent();
		}
	}

	function ssOleDBGridEventLog_rowcolchange() {

		menu_enableMenuItem("mnutoolViewEventLogFind", frmLog.ssOleDBGridEventLog.SelBookmarks.Count == 1);


	}

	function ssOleDBGridEventLog_click() {

		menu_enableMenuItem("mnutoolViewEventLogFind",
												(!(frmLog.ssOleDBGridEventLog.SelBookmarks.Count > 1) || (frmLog.ssOleDBGridEventLog.Rows == 0)));

	}

	function ssOleDBGridEventLog_headclick() {

		var ColIndex = arguments[0];

		//Set the sort criteria depending on the column header clicked and refresh the grid
		if (ColIndex == 1) {
			frmLog.txtELOrderColumn.value = 'DateTime';
		}
		else if (ColIndex == 2) {
			frmLog.txtELOrderColumn.value = 'EndTime';
		}
		else if (ColIndex == 3) {
			frmLog.txtELOrderColumn.value = 'Duration';
		}
		else if (ColIndex == 4) {
			frmLog.txtELOrderColumn.value = 'Type';
		}
		else if (ColIndex == 5) {
			frmLog.txtELOrderColumn.value = 'Name';
		}
		else if (ColIndex == 6) {
			frmLog.txtELOrderColumn.value = 'Status';
		}
		else if (ColIndex == 7) {
			frmLog.txtELOrderColumn.value = 'Mode';
		}
		else if (ColIndex == 8) {
			frmLog.txtELOrderColumn.value = 'Username';
		}
		else {
			frmLog.txtELOrderColumn.value = 'DateTime';
		}

		if (ColIndex == frmLog.txtELSortColumnIndex.value) {
			if (frmLog.txtELOrderOrder.value == 'ASC') {
				frmLog.txtELOrderOrder.value = 'DESC';
			}
			else {
				frmLog.txtELOrderOrder.value = 'ASC';
			}
		}
		else {
			frmLog.txtELOrderOrder.value = 'ASC';
		}

		frmLog.txtELSortColumnIndex.value = ColIndex;

		refreshGrid();
	}


</script>


<script type="text/javascript">
	eventLog_window_onload();
	eventlog_addActiveXHandlers();
</script>
