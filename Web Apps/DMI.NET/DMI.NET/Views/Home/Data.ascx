<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%	
	If Len(Session("recordID")) = 0 Then
		Session("recordID") = 0
	End If
%>

<script type="text/javascript">
	function data_window_onload() {
		var frmData = document.getElementById("frmData");
		var frmGetData = document.getElementById("frmGetData");
		var frmMenuInfo = OpenHR.getForm("menuframe", "frmMenuInfo");
		var frmWorkAreaInfo = OpenHR.getForm("menuframe", "frmWorkAreaInfo");
		var frmOptionArea = OpenHR.getForm("optionframe", "frmGotoOption");
		var frmRecEditArea = OpenHR.getForm("workframe", "frmRecordEditForm");
		var frmFindForm = OpenHR.getForm("workframe", "frmFindForm");
		var recEditForm = OpenHR.getForm("workframe", "frmRecordEditForm");
		var frmLog = OpenHR.getForm("workframe", "frmLog");
		
		var sFatalErrorMsg = frmData.txtErrorDescription.value;
		var sInsertGranted;
		var sErrorMsg, sErrMsg;
		var sCourseTitle;
		var iResult;
		
	if (sFatalErrorMsg.length > 0) {
		OpenHR.messageBox(sFatalErrorMsg);
		//TODO
		window.location = "Login";
	}
	else {
		// Do nothing if the menu controls are not yet instantiated.
		if (frmWorkAreaInfo != null) {
			var sCurrentWorkPage = OpenHR.currentWorkPage();

			if (sCurrentWorkPage == "RECORDEDIT") {
				// Refresh the recEdit controls with the data if required.
				var recEditControl = recEditForm.ctlRecordEdit;
				sErrorMsg = frmData.txtErrorMessage.value;

				if (sErrorMsg.length > 0) {
					if (frmData.txtWarning.value == "True") {
						// We've got a warning.
						sErrorMsg = sErrorMsg + "\nDo you still want to save the record ?";
						menu_refreshMenu();		  
						iResult = OpenHR.messageBox(sErrorMsg, 36); // 36 = yesNo + question
						
						if (iResult == 6) { 
							// Yes ... go on and save it.
							menu_saveChanges(frmData.txtAction.value,false,true);		  
							return;
						}
						else {
							// No ... don't save it.
							menu_refreshMenu();		  
							return;
						}
					}
					else {
						// We've got an error so don't update the record edit form.

						// Get menu to refresh the menu.
						menu_refreshMenu();		  
						OpenHR.messageBox(sErrorMsg);
			
						if (frmData.txtAction.value == "SAVEERROR") {
							return;
						}
					}					
				}
		
				var sAction = frmData.txtAction.value;

				if ((sAction == "LOAD")	
					&& (frmData.txtOriginalRecID.value != frmData.txtNewRecID.value)
					&& (frmData.txtOriginalRecID.value != 0))	{

					menu_refreshMenu();	
					
					if (recEditForm.txtRecEditFilterSQL.value == "") {
						OpenHR.messageBox("The record saved is no longer in the current view");
					}
				}

				if (sAction == "LOGOFF") {
					//TODO
					window.location.href = frmMenuInfo.txtDefaultStartPage.value;
					return;	
				}

				if (sAction == "EXIT") {
					//TODO
					window.close();
				}

				if (sAction == "CROSSTABS") {
					menu_loadDefSelPage(1, 0, 0, true);
				}
				
				if (sAction == "CUSTOMREPORTS") {
					menu_loadDefSelPage(2, 0, 0, true);
				}
				
				if (sAction == "CALENDARREPORTS") {
					menu_loadDefSelPage(17, 0, 0, true);
				}
				
				if (sAction == "MAILMERGE") {
					menu_loadDefSelPage(9, 0, 0, true);
				}
				
				if (sAction == "WORKFLOW") {
					menu_loadDefSelPage(25, 0, 0, true);
				}
				
				if (sAction == "WORKFLOWPENDINGSTEPS") {
					menu_autoLoadPage("workflowPendingSteps", false);
				}
				
				if (sAction == "WORKFLOWOUTOFOFFICE") {
					menu_WorkflowOutOfOffice();
				}
				
				if (sAction == "PICKLISTS") {
					menu_loadDefSelPage(10, 0, 0, true);
				}
				
				if (sAction == "FILTERS") {
					menu_loadDefSelPage(11, 0, 0, true);
				}
				
				if (sAction == "CALCULATIONS") {
					menu_loadDefSelPage(12, 0, 0, true);
				}
				
				if (sAction == "CANCELCOURSE") {
					menu_cancelCourse();
				}
				if (sAction == "ABSENCEBREAKDOWNREC") {
					menu_LoadStandardReportNoSaveCheck("ABSENCEBREAKDOWN","REC");
				}
				if (sAction == "ABSENCEBREAKDOWNALL") {
					menu_LoadStandardReportNoSaveCheck("ABSENCEBREAKDOWN","ALL");
				}
				if (sAction == "BRADFORDFACTORREC") {
					menu_LoadStandardReportNoSaveCheck("BRADFORDFACTOR","REC");
				}
				if (sAction == "BRADFORDFACTORALL") {
					menu_LoadStandardReportNoSaveCheck("BRADFORDFACTOR","ALL");
				}
				if (sAction == "STDRPT_ABSENCECALENDAR") {
					menu_LoadAbsenceCalendarNoSaveCheck();
				}
				if (sAction == "CALENDARREPORTSREC") {
					menu_loadRecordDefSelPageNoSaveCheck(17);
				}

				if (sAction == "EVENTLOG") {
					menu_loadPage("eventLog");
				}

				if (sAction == "QUICKFIND") {
					menu_loadQuickFindNoSaveCheck();
				}

				if (sAction == "PARENT") {
					menu_loadParent();
				}

				if (sAction == "CANCELCOURSE_1") {
					sErrMsg = new String(frmData.txtTBErrorMessage.value);
					if (sErrMsg.length > 0){
						menu_refreshMenu();		  
						OpenHR.messageBox(sErrMsg);
					}
					else {	
						sCourseTitle = new String(frmData.txtTBCourseTitle.value);
						
						if ((frmData.txtNumberOfBookings.value > 0) && (sCourseTitle.length > 0)) {
							/* Ask the user if they want to transfer the bookings. */
							menu_refreshMenu();		  
							iResult = OpenHR.messageBox("Transfer bookings to another course ?", 35); // 35 = yesNoCancel + question

							if (iResult	== 6) {
								// Yes
								// Display the course selection page.
								menu_loadTransferCoursePage(sCourseTitle);
							}	
							if (iResult	== 7) {
								// No.
								menu_transferCourse(0, true);
							}	
							if (iResult	== 2) {
								// Cancel.
							}	
						}
						else {
							menu_transferCourse(0, false);
						}
					}
				}
				
				if (sAction.substring(0, 7) == "mnutool") {
					if (sAction == "mnutoolFind") {
						menu_loadFindPage();
						return;
					}				
					
					menu_loadPage(sAction.substring(7, sAction.length));
					return;
				}
				else {		
					if ((sAction.substring(0, 3) == "PT_") ||
						(sAction.substring(0, 3) == "PV_")) {
						// PT_ = primary table
						// PV_ = primary table view
						if (frmMenuInfo.txtPrimaryStartMode.value == 3) {
							frmData.txtRecordDescription.value = "";
							menu_loadFindPageFirst(sAction);
						}
						else {
							menu_loadRecordEditPage(sAction);
						}
						return;
					}
					
					if (sAction.substring(0, 3) == "TS_") {
						// TS_ = Table screen
						if (frmMenuInfo.txtLookupStartMode.value == 3) {
							frmData.txtRecordDescription.value = "";
							menu_loadFindPageFirst(sAction);
						}
						else {
							menu_loadRecordEditPage(sAction);
						}
						return;
					}

					if (sAction.substring(0, 3) == "QE_") {
						// QE_ = quick entry screen
						if (frmMenuInfo.txtQuickAccessStartMode.value == 3) {
							frmData.txtRecordDescription.value = "";
							menu_loadFindPageFirst(sAction);
						}
						else {
							menu_loadRecordEditPage(sAction);
						}
						return;
					}

					if (sAction.substring(0, 3) == "HT_") {
						// HT_ = history table
						if (frmMenuInfo.txtHistoryStartMode.value == 3) {
							frmData.txtRecordDescription.value = "";
							menu_loadFindPageFirst(sAction);
						}
						else {
							menu_loadRecordEditPage(sAction);
						}
						return;
					}
				}

				if ((frmMenuInfo.txtUserType.value == 1) && 
					(frmMenuInfo.txtPersonnel_EmpTableID.value == frmRecEditArea.txtCurrentTableID.value) && 
					(frmData.txtRecordCount.value > 1)) {
					
					// Get menu to refresh the menu.
					menu_refreshMenu();		  

					/* The user does NOT have permission to create new records. */
					OpenHR.messageBox("Unable to load personnel records.\n\nYou are logged on as a self-service user and can access only single record personnel record sets.");

					/* Go to the default page. */
					menu_loadPage("default");
					return;
				}

				if (sAction == "NEW") {
					applyDefaultValues();					
				}

				var sControlName;
				var sColumnId;
				var dataCollection = frmData.elements;

				var frmRecEditForm = document.getElementById("frmRecordEditForm");

				if (dataCollection!=null) {
					// Need to hide the popup in case setdata causes
					// the intrecedit control to display an error message.
					$("#ctlRecordEdit #changed").val("false");
					menu_refreshMenu();

					for (var i=0; i<dataCollection.length; i++)  {
					  sControlName = dataCollection.item(i).name;
						sControlName = sControlName.substr(0, 8);
						if (sControlName=="txtData_") {
						  sColumnId = dataCollection.item(i).name;
						  sColumnId = sColumnId.substr(8);
						    var x = $("#FI_" + sColumnId);
						    //recEditControl.setData(sColumnId, dataCollection.item(i).value);
						    //$("#FI_" + sColumnId).val(dataCollection.item(i).value);
						    //setData function is in recordEdit.ascx.						    
						    recEdit_setData(sColumnId, dataCollection.item(i).value);						    
						}
					}
				}	

				
				//TODO: recEditControl.ChangedOLEPhoto(0, "NONE");				
			    
				recEdit_setRecordID(frmData.txtRecordID.value); //workframe
				recEdit_setParentTableID(frmData.txtParentTableID.value); //workframe
				recEdit_setParentRecordID(frmData.txtParentRecordID.value); //workframe
				
				/* Check if the record is empty. */
				if ((sAction != "NEW") && (sAction != "COPY") && (frmData.txtRecordCount.value == 0)) {
					// No records. Clear the filter.
					if (recEditForm.txtRecEditFilterSQL.value.length  > 0) {
						OpenHR.messageBox("No records match the current filter. No filter is applied.");

						frmGetData.txtAction.value = "LOAD";
						frmGetData.txtCurrentTableID.value = recEditForm.txtCurrentTableID.value;
						frmGetData.txtCurrentScreenID.value = recEditForm.txtCurrentScreenID.value;
						frmGetData.txtCurrentViewID.value = recEditForm.txtCurrentViewID.value;
						frmGetData.txtSelectSQL.value = recEditForm.txtRecEditSelectSQL.value;
						frmGetData.txtFromDef.value = recEditForm.txtRecEditFromDef.value;
						recEditForm.txtRecEditFilterSQL.value = "";
						recEditForm.txtRecEditFilterDef.value = "";
						frmGetData.txtFilterSQL.value = "";
						frmGetData.txtFilterDef.value = "";
						frmGetData.txtRealSource.value = recEditForm.txtRecEditRealSource.value;
						frmGetData.txtRecordID.value = frmData.txtRecordID.value;
						frmGetData.txtParentTableID.value = frmData.txtParentTableID.value;
						frmGetData.txtParentRecordID.value = frmData.txtParentRecordID.value;
						frmGetData.txtDefaultCalcCols.value = CalculatedDefaultColumns();
						data_refreshData();
					}
					else {
						/* If the recordset is empty and we are not already creating a new record,
						then try to create a new record now (if permitted). */
						sInsertGranted = recEditForm.txtRecEditInsertGranted.value;
						sInsertGranted = sInsertGranted.toUpperCase();

						if (sInsertGranted == "TRUE") {
							/* The user does have permission to create new records. */
							frmGetData.txtAction.value = "NEW";
							frmGetData.txtCurrentTableID.value = recEditForm.txtCurrentTableID.value;
							frmGetData.txtCurrentScreenID.value = recEditForm.txtCurrentScreenID.value;
							frmGetData.txtCurrentViewID.value = recEditForm.txtCurrentViewID.value;
							frmGetData.txtSelectSQL.value = recEditForm.txtRecEditSelectSQL.value;
							frmGetData.txtFromDef.value = recEditForm.txtRecEditFromDef.value;
							frmGetData.txtFilterSQL.value = "";
							frmGetData.txtFilterDef.value = "";
							frmGetData.txtRealSource.value = recEditForm.txtRecEditRealSource.value;
							frmGetData.txtRecordID.value = frmData.txtRecordID.value;
							frmGetData.txtParentTableID.value = frmData.txtParentTableID.value;
							frmGetData.txtParentRecordID.value = frmData.txtParentRecordID.value;
							frmGetData.txtDefaultCalcCols.value = CalculatedDefaultColumns();
							data_refreshData();
						}
						else {
							// Get menu to refresh the menu.
							menu_refreshMenu();		  

							/* The user does NOT have permission to create new records. */
							OpenHR.messageBox("This table is empty and you do not have 'new' permission on it.");

							/* Go to the find page. */
							menu_loadFindPage();
						}
					}
				}
				
				if (sAction == "COPY") {
					recEdit_setRecordID(0); //workframe
					recEdit_setCopiedRecordID(<%= CLng(session("recordID"))%>); //workframe
					frmData.txtRecordPosition.value = frmData.txtRecordCount.value + 1;
					ClearUniqueColumnControls();
					//TODO: recEditControl.ChangedOLEPhoto(0, "ALL");
					$("#ctlRecordEdit #changed").val("true");
				}

				if (sAction == "NEW") {
					$("#ctlRecordEdit #changed").val(allDefaults());
				}

			    // Get menu to refresh the menu.
			    
				menu_refreshMenu();

				if (sAction == "SELECTORDER") {
					frmOptionArea.txtGotoOptionScreenID.value = frmRecEditArea.txtCurrentScreenID.value;
					frmOptionArea.txtGotoOptionTableID.value = frmRecEditArea.txtCurrentTableID.value;
					frmOptionArea.txtGotoOptionViewID.value = frmRecEditArea.txtCurrentViewID.value;
					frmOptionArea.txtGotoOptionOrderID.value = frmRecEditArea.txtCurrentOrderID.value;
					frmOptionArea.txtGotoOptionFilterDef.value = frmRecEditArea.txtRecEditFilterDef.value;
					frmOptionArea.txtGotoOptionPage.value = "orderselect";
					frmOptionArea.submit();
					return;
				}				

				if (sAction == "SELECTFILTER") {
					frmOptionArea.txtGotoOptionScreenID.value = frmRecEditArea.txtCurrentScreenID.value;
					frmOptionArea.txtGotoOptionTableID.value = frmRecEditArea.txtCurrentTableID.value;
					frmOptionArea.txtGotoOptionViewID.value = frmRecEditArea.txtCurrentViewID.value;
					frmOptionArea.txtGotoOptionOrderID.value = frmRecEditArea.txtCurrentOrderID.value;
					frmOptionArea.txtGotoOptionFilterDef.value = frmRecEditArea.txtRecEditFilterDef.value;
					frmOptionArea.txtGotoOptionPage.value = "filterselect";
					//frmOptionArea.submit();
					OpenHR.submitForm(frmOptionArea);
					return;
				}				
				
				if (sAction == "CLEARFILTER") {
					frmRecEditArea.txtRecEditFilterDef.value = "";
					frmRecEditArea.txtRecEditFilterSQL.value = "";
					//TODO should call refreshData in workframe form not local refreshData			
					refreshData(); //workframe
					return;
				}				
			}
			else if (sCurrentWorkPage == "FIND") {

				sErrorMsg = frmData.txtErrorMessage.value;
				if (sErrorMsg.length > 0) {
					// Get menu to refresh the menu.
					menu_refreshMenu();		  

					// We've got an error so don't update the find form.
					OpenHR.messageBox(sErrorMsg);
			
					if (frmData.txtAction.value == "SAVEERROR") {
						return;
					}
				}
				
				if (frmData.txtAction.value == "CANCELBOOKING_1") {
					OpenHR.messageBox("Booking cancelled.", 64);
					menu_reloadFindPage("RELOAD", "");
				}

				// No error deleting 
				if (frmData.txtAction.value == "REFRESHFINDAFTERDELETE") {

					var iRowID = $("#findGridTable").getGridParam('selrow');
					$('#findGridTable').jqGrid('delRowData', iRowID);
					

			  //  	var ctlFindGrid = frmFindForm.ssOleDBGridFindRecords;   
			  //  	var iAbsRowNo = ctlFindGrid.AddItemRowIndex(ctlFindGrid.Bookmark);
			  //  	if(ctlFindGrid.Rows == 1) {
			  //  		ctlFindGrid.removeAll();
			  //  	}
			  //  	else {
			  //  		ctlFindGrid.RemoveItem(iAbsRowNo);
			  //  	}					

		      //if (iAbsRowNo < ctlFindGrid.Rows) {
			  //  		ctlFindGrid.Bookmark = ctlFindGrid.AddItemBookmark(iAbsRowNo);
		      //}
    		  //else if (ctlFindGrid.Rows > 0) {
			  //  		ctlFindGrid.Bookmark = ctlFindGrid.AddItemBookmark(ctlFindGrid.Rows - 1);
		      //}
    		  //ctlFindGrid.SelBookmarks.Add(ctlFindGrid.Bookmark);

					// Update controls in the find form to ensure the displayed
					// record count is correct.
					frmFindForm.txtRecordCount.value = frmFindForm.txtRecordCount.value - 1;
					frmFindForm.txtTotalRecordCount.value = frmFindForm.txtTotalRecordCount.value - 1;
					frmFindForm.txtCurrentRecCount.value = frmFindForm.txtCurrentRecCount.value - 1;
					
					// Get menu to refresh the menu.
					menu_refreshMenu();		  
				}
			}
			else if ((sCurrentWorkPage == "UTIL_DEF_CUSTOMREPORTS") ||
					 (sCurrentWorkPage == "UTIL_DEF_CALENDARREPORT") ||
			         (sCurrentWorkPage == "UTIL_DEF_CROSSTABS") ||
			         (sCurrentWorkPage == "UTIL_DEF_MAILMERGE") ||
							 (sCurrentWorkPage == "EVENTLOG")) {
			    
				if (frmData.txtAction.value == "LOADREPORTCOLUMNS") 
					{
					loadAvailableColumns(); //workframe
					}
				else if (frmData.txtAction.value == "LOADCALENDARREPORTCOLUMNS") 
					{
					loadAvailableColumns(); //workframe
					}
				else if (frmData.txtAction.value == "GETEXPRESSIONRETURNTYPES") 
					{
					loadExpressionTypes(); //workframe
					}
				else if (frmData.txtAction.value == "LOADEMAILDEFINITIONS") 
					{
					loadEmailDefs(); //workframe
					}
				else if (frmData.txtAction.value == "LOADEVENTLOG") 
					{
					loadEventLog(); //workframe
					return;
					}
				else if (frmData.txtAction.value == "LOADEVENTLOGUSERS") 
					{
					var bAllPerms = frmLog.txtELViewAllPermission.value;
					var sCurrentUserFilter = frmLog.txtCurrUserFilter.value;
					
					loadEventLogUsers(bAllPerms,sCurrentUserFilter); //workframe
					
					return;
					}
				}
		}
	}
    }

</script>

<script type="text/javascript">

	function data_refreshData() {		
		var f = document.getElementById("frmGetData");
		OpenHR.submitForm(f);
	}

</script>

<div>

<form action="data_submit" method=post id=frmGetData name=frmGetData data-formname="data.ascx">
	<INPUT type="hidden" id=txtAction name=txtAction>
	<INPUT type="hidden" id=txtReaction name=txtReaction>
	<INPUT type="hidden" id=txtCurrentTableID name=txtCurrentTableID>
	<INPUT type="hidden" id=txtCurrentScreenID name=txtCurrentScreenID>
	<INPUT type="hidden" id=txtCurrentViewID name=txtCurrentViewID>
	<INPUT type="hidden" id=txtSelectSQL name=txtSelectSQL>
	<INPUT type="hidden" id=txtFromDef name=txtFromDef>
	<INPUT type="hidden" id=txtFilterSQL name=txtFilterSQL>
	<INPUT type="hidden" id=txtFilterDef name=txtFilterDef>
	<INPUT type="hidden" id=txtRealSource name=txtRealSource>
	<INPUT type="hidden" id=txtRecordID name=txtRecordID>
	<INPUT type="hidden" id=txtParentTableID name=txtParentTableID>
	<INPUT type="hidden" id=txtParentRecordID name=txtParentRecordID>
	<INPUT type="hidden" id=txtDefaultCalcCols name=txtDefaultCalcCols>
	<INPUT type="hidden" id=txtInsertUpdateDef name=txtInsertUpdateDef>
	<INPUT type="hidden" id=txtTimestamp name=txtTimestamp>
	<INPUT type="hidden" id=txtTBCourseRecordID name=txtTBCourseRecordID>
	<INPUT type="hidden" id=txtTBEmployeeRecordID name=txtTBEmployeeRecordID>
	<INPUT type="hidden" id=txtTBBookingStatusValue name=txtTBBookingStatusValue>
	<INPUT type="hidden" id=txtTBOverride name=txtTBOverride>
	<INPUT type="hidden" id=txtTBCreateWLRecords name=txtTBCreateWLRecords>
	<INPUT type="hidden" id=txtReportBaseTableID name=txtReportBaseTableID>
	<INPUT type="hidden" id=txtReportParent1TableID name=txtReportParent1TableID>
	<INPUT type="hidden" id=txtReportParent2TableID name=txtReportParent2TableID>
	<INPUT type="hidden" id=txtReportChildTableID name=txtReportChildTableID>
	<INPUT type="hidden" id=txtUserChoice name=txtUserChoice>
	<INPUT type="hidden" id=txtParam1 name=txtParam1>
	<INPUT type="hidden" id=txtELFilterUser name=txtELFilterUser>
	<INPUT type="hidden" id=txtELFilterType name=txtELFilterType>
	<INPUT type="hidden" id=txtELFilterStatus name=txtELFilterStatus>
	<INPUT type="hidden" id=txtELFilterMode name=txtELFilterMode>
	<INPUT type="hidden" id=txtELOrderColumn name=txtELOrderColumn>
	<INPUT type="hidden" id=txtELOrderOrder name=txtELOrderOrder>
	<INPUT type="hidden" id=txtELAction name=txtELAction>
	<INPUT type="hidden" id=txtELCurrRecCount name=txtELCurrRecCount value="0">
	<INPUT type="hidden" id=txtEL1stRecPos name=txtEL1stRecPos value ="0">
	
</form>

<form id=frmData name=frmData>
<%
'	on error resume next
	
	Dim lngRecordID As Object
	
	Const DEADLOCK_ERRORNUMBER = -2147467259
	Const DEADLOCK_MESSAGESTART = "YOUR TRANSACTION (PROCESS ID #"
	Const DEADLOCK_MESSAGEEND = ") WAS DEADLOCKED WITH ANOTHER PROCESS AND HAS BEEN CHOSEN AS THE DEADLOCK VICTIM. RERUN YOUR TRANSACTION."
	Const DEADLOCK2_MESSAGESTART = "TRANSACTION (PROCESS ID "
	Const DEADLOCK2_MESSAGEEND = ") WAS DEADLOCKED ON "

	Const iRETRIES = 5
	Dim iRetryCount = 0

	Dim sErrorDescription = ""

	Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtParentTableID name=txtParentTableID value=" & Session("parentTableID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtParentRecordID name=txtParentRecordID value=" & Session("parentRecordID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtWarning name=txtWarning value=" & Session("warningFlag") & ">" & vbCrLf)
	' Clear the error message session variable.
	Session("errorMessage") = ""

	' Get the required record count if we have a query.
	if len(session("selectSQL")) > 0 then
		if session("action") = "NEW" then
			Dim cmdGetRecord = CreateObject("ADODB.Command")
			cmdGetRecord.CommandText = "sp_ASRIntCalcDefaults"
			cmdGetRecord.CommandType = 4 ' Stored procedure
			cmdGetRecord.ActiveConnection = Session("databaseConnection")

			Dim prmRecordCount = cmdGetRecord.CreateParameter("recordCount", 3, 2)
			cmdGetRecord.Parameters.Append(prmRecordCount)

			Dim prmFromDef = cmdGetRecord.CreateParameter("fromDef", 200, 1, 2147483646)
			cmdGetRecord.Parameters.Append(prmFromDef)
			prmFromDef.value = Session("fromDef")

			Dim prmFilterDef = cmdGetRecord.CreateParameter("filterDef", 200, 1, 2147483646)
			cmdGetRecord.Parameters.Append(prmFilterDef)
			prmFilterDef.value = Session("filterDef")

			Dim prmTableId = cmdGetRecord.CreateParameter("tableID", 3, 1)
			cmdGetRecord.Parameters.Append(prmTableId)
			prmTableId.value = CleanNumeric(Session("tableID"))

			Dim prmParentTableId = cmdGetRecord.CreateParameter("parentTableID", 3, 1)
			cmdGetRecord.Parameters.Append(prmParentTableId)
			prmParentTableId.value = CleanNumeric(Session("parentTableID"))

			Dim prmParentRecordId = cmdGetRecord.CreateParameter("parentRecordID", 3, 1)
			cmdGetRecord.Parameters.Append(prmParentRecordId)
			prmParentRecordId.value = CleanNumeric(Session("parentRecordID"))
	
			Dim prmDefaultCalcCols = cmdGetRecord.CreateParameter("defaultCalcCols", 200, 1, 2147483646)
			cmdGetRecord.Parameters.Append(prmDefaultCalcCols)
			prmDefaultCalcCols.value = Session("defaultCalcColumns")
	
			Dim prmDecSeparator = cmdGetRecord.CreateParameter("decSeparator", 200, 1, 255)
			cmdGetRecord.Parameters.Append(prmDecSeparator)
			prmDecSeparator.value = Session("LocaleDecimalSeparator")

			Dim prmDateFormat = cmdGetRecord.CreateParameter("dateFormat", 200, 1, 255)
			cmdGetRecord.Parameters.Append(prmDateFormat)
			prmDateFormat.value = Session("LocaleDateFormat")
		
			Err.Clear()
			Dim rstRecord = cmdGetRecord.Execute
	
			If (Err.Number <> 0) Then
				sErrorDescription = "Default values could not be calculated." & vbCrLf & formatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				If rstRecord.state = 1 Then 'adStateOpen
					If Not (rstRecord.bof And rstRecord.eof) Then
						For iloop = 0 To (rstRecord.fields.count - 1)
							If IsDBNull(rstRecord.fields(iloop).value) Then
								Response.Write("<INPUT type='hidden' id=txtData_" & rstRecord.fields(iloop).name & " name=txtData_" & rstRecord.fields(iloop).name & " value="""">" & vbCrLf)
							Else
								Response.Write("<INPUT type='hidden' id=txtData_" & rstRecord.fields(iloop).name & " name=txtData_" & rstRecord.fields(iloop).name & " value=""" & Replace(rstRecord.fields(iloop).value, """", "&quot;") & """>" & vbCrLf)
							End If
						Next
					End If
	
					' Release the ADO recordset object.
					rstRecord.close()
				End If
			End If
			rstRecord = Nothing

			If Session("parentTableID") > 0 Then
				Dim cmdGetParentValues = CreateObject("ADODB.Command")
				cmdGetParentValues.CommandText = "spASRIntGetParentValues"
				cmdGetParentValues.CommandType = 4 ' Stored procedure
				cmdGetParentValues.ActiveConnection = Session("databaseConnection")

				Dim prmScreenId = cmdGetParentValues.CreateParameter("screenID", 3, 1)
				cmdGetParentValues.Parameters.Append(prmScreenId)
				prmScreenId.value = cleanNumeric(Session("screenID"))

				Dim prmParentTableId2 = cmdGetParentValues.CreateParameter("parentTableID", 3, 1)
				cmdGetParentValues.Parameters.Append(prmParentTableId2)
				prmParentTableId2.value = CleanNumeric(Session("parentTableID"))

				Dim prmParentRecordId2 = cmdGetParentValues.CreateParameter("parentRecordID", 3, 1)
				cmdGetParentValues.Parameters.Append(prmParentRecordId2)
				prmParentRecordId2.value = CleanNumeric(Session("parentRecordID"))
					
				Err.Clear()
				Dim rstParentValues = cmdGetParentValues.Execute
					
				If (Err.Number <> 0) Then
					sErrorDescription = "Parent values could not be determined." & vbCrLf & formatError(Err.Description)
				End If

				If Len(sErrorDescription) = 0 Then
					If rstParentValues.state = 1 Then 'adStateOpen
						If Not (rstParentValues.bof And rstParentValues.eof) Then
							For iloop = 0 To (rstParentValues.fields.count - 1)
								If IsDBNull(rstParentValues.fields(iloop).value) Then
									Response.Write("<INPUT type='hidden' id=txtData_" & rstParentValues.fields(iloop).name & " name=txtData_" & rstParentValues.fields(iloop).name & " value="""">" & vbCrLf)
								Else
									Response.Write("<INPUT type='hidden' id=txtData_" & rstParentValues.fields(iloop).name & " name=txtData_" & rstParentValues.fields(iloop).name & " value=""" & Replace(rstParentValues.fields(iloop).value, """", "&quot;") & """>" & vbCrLf)
								End If
							Next
						End If
					
						' Release the ADO recordset object.
						rstParentValues.close()
					End If
				End If
				rstParentValues = Nothing
			End If

			If Len(sErrorDescription) = 0 Then
				Response.Write("<INPUT type='hidden' id=txtRecordID name=txtRecordID value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=" & cmdGetRecord.Parameters("recordCount").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordPosition name=txtRecordPosition value=" & cmdGetRecord.Parameters("recordCount").Value + 1 & ">" & vbCrLf)
			Else
				Response.Write("<INPUT type='hidden' id=txtRecordID name=txtRecordID value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordPosition name=txtRecordPosition value=0>" & vbCrLf)
			End If
			
			lngRecordID = 0
			Response.Write("<INPUT type='hidden' id=txtOriginalRecID name=txtOriginalRecID value=0>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtNewRecID name=txtNewRecID value=0>" & vbCrLf)
		Else
			Dim cmdGetRecord = CreateObject("ADODB.Command")
			cmdGetRecord.CommandText = "sp_ASRIntGetRecord"
			cmdGetRecord.CommandType = 4 ' Stored procedure
			cmdGetRecord.ActiveConnection = Session("databaseConnection")
			cmdGetRecord.CommandTimeout = 180

			Dim prmRecordId = cmdGetRecord.CreateParameter("recordID", 3, 3) ' 3 = integer, 3 = input & output
			cmdGetRecord.Parameters.Append(prmRecordId)
			prmRecordId.value = CleanNumeric(Session("recordID"))

			Dim prmRecordCount = cmdGetRecord.CreateParameter("recordCount", 3, 2) ' 3 = integer, 2 = output
			cmdGetRecord.Parameters.Append(prmRecordCount)

			Dim prmRecordPosition = cmdGetRecord.CreateParameter("recordPosition", 3, 2) ' 3 = integer, 2 = output
			cmdGetRecord.Parameters.Append(prmRecordPosition)

			Dim prmFilterDef = cmdGetRecord.CreateParameter("filterDef", 201, 1, 2147483646)	' 200 = varchar, 1 = input, 8000 = size
			cmdGetRecord.Parameters.Append(prmFilterDef)
			prmFilterDef.value = session("filterDef")

			Dim prmAction = cmdGetRecord.CreateParameter("action", 200, 1, 100) ' 200 = varchar, 1 = input, 100 = size
			cmdGetRecord.Parameters.Append(prmAction)
			prmAction.value = session("action")

			Dim prmParentTableId = cmdGetRecord.CreateParameter("parentTableID", 3, 1)	' 3 = integer, 1 = input
			cmdGetRecord.Parameters.Append(prmParentTableId)
			prmParentTableId.value = CleanNumeric(Session("parentTableID"))

			Dim prmParentRecordId = cmdGetRecord.CreateParameter("parentRecordID", 3, 1) ' 3 = integer, 1 = input
			cmdGetRecord.Parameters.Append(prmParentRecordId)
			prmParentRecordId.value = CleanNumeric(Session("parentRecordID"))
	
			Dim prmDecSeparator = cmdGetRecord.CreateParameter("decSeparator", 200, 1, 100) ' 200=varchar, 1=input, 8000=size
			cmdGetRecord.Parameters.Append(prmDecSeparator)
			prmDecSeparator.value = session("LocaleDecimalSeparator")

			Dim prmDateFormat = cmdGetRecord.CreateParameter("dateFormat", 200, 1, 100) ' 200=varchar, 1=input, 8000=size
			cmdGetRecord.Parameters.Append(prmDateFormat)
			prmDateFormat.value = session("LocaleDateFormat")

			Dim prmScreenId = cmdGetRecord.CreateParameter("screenID", 3, 1) ' 3=integer, 1=input
			cmdGetRecord.Parameters.Append(prmScreenId)
			prmScreenId.value = CleanNumeric(Session("screenID"))

			Dim prmViewId = cmdGetRecord.CreateParameter("viewID", 3, 1) ' 3=integer, 1=input
			cmdGetRecord.Parameters.Append(prmViewId)
			prmViewId.value = CleanNumeric(Session("viewID"))

			Dim prmOrderId = cmdGetRecord.CreateParameter("orderID", 3, 1)	' 3=integer,  1=input
			cmdGetRecord.Parameters.Append(prmOrderId)
			prmOrderId.value = CleanNumeric(Session("orderID"))

			Dim fOk = True
			Dim fDeadlock = True
			Dim sErrMsg = ""
			Dim oleColumnData As New List(Of Object)

			Do While fDeadlock
				fDeadlock = False

				cmdGetRecord.ActiveConnection.Errors.Clear()

				Dim rstRecord = cmdGetRecord.Execute
		
				If cmdGetRecord.ActiveConnection.Errors.Count > 0 Then
					For iLoop = 1 To cmdGetRecord.ActiveConnection.Errors.Count
						sErrMsg = formatError(cmdGetRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)

						If (cmdGetRecord.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
						 (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
						  (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
						 ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
										 (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then

							' The error is for a deadlock.
							' Sorry about having to use the err.description to trap the error but the err.number
							' is not specific and MSDN suggests using the err.description.
							If (iRetryCount < iRETRIES) And (cmdGetRecord.ActiveConnection.Errors.Count = 1) Then
								iRetryCount = iRetryCount + 1
								fDeadlock = True
							Else
								If Len(sErrorDescription) > 0 Then
									sErrorDescription = sErrorDescription & vbCrLf
								End If
								sErrorDescription = sErrorDescription & "Another user is deadlocking the database. Please try again."
								fOk = False
							End If
						Else
							sErrorDescription = sErrorDescription & vbCrLf & _
							 formatError(cmdGetRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)
							fOk = False
						End If
					Next

					cmdGetRecord.ActiveConnection.Errors.Clear()
												
					If Not fOk Then
						sErrorDescription = "Unable to retrieve the required record." & vbCrLf & sErrorDescription
					End If
				Else
				
					If Not (rstRecord.bof And rstRecord.eof) Then
						For iloop = 0 To (rstRecord.fields.count - 1)
						
							If IsDBNull(rstRecord.fields(iloop).value) Then
								Response.Write("<INPUT type='hidden' id=txtData_" & rstRecord.fields(iloop).name & " name=txtData_" & rstRecord.fields(iloop).name & " value="""">" & vbCrLf)
							Else
								' Is column a embedded/linked OLE
								If VarType(rstRecord.fields(iloop).value) = 8209 Then
									oleColumnData.Add(rstRecord.fields(iloop).name)
								Else
									Response.Write("<INPUT type='hidden' id=txtData_" & rstRecord.fields(iloop).name & " name=txtData_" & rstRecord.fields(iloop).name & " value=""" & Replace(rstRecord.fields(iloop).value, """", "&quot;") & """>" & vbCrLf)
								End If
							End If
						Next
					End If

					'	Release the ADO recordset object.
					rstRecord.close()
					rstRecord = Nothing
				End If
			Loop

			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and Dim it to nothing before using the output parameters. 

			Response.Write("<INPUT type='hidden' id=txtOriginalRecID name=txtOriginalRecID value=" & session("recordID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtNewRecID name=txtNewRecID value=" & cmdGetRecord.Parameters("recordID").Value & ">" & vbCrLf)
			
			' Loop through the OLE columns that have data in them
			Dim objOle = Session("OLEObject")
			objOle.CleanupOLEFiles()
			
			For Each item As Object In oleColumnData
				Dim strDisplayValue = objOle.GetPropertiesFromStream(cmdGetRecord.Parameters("recordID").Value, item, Session("realSource"))
				Response.Write("<INPUT type='hidden' id=txtData_" & item & " name=txtData_" & item & " value=""" & strDisplayValue & """>" & vbCrLf)
			Next
						
			Session("OLEObject") = objOle
			objOle = Nothing


			If Len(sErrorDescription) = 0 Then
				If session("action") = "COPY" Then
					Response.Write("<INPUT type='hidden' id=txtRecordID name=txtRecordID value=0>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=" & cmdGetRecord.Parameters("recordCount").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtRecordPosition name=txtRecordPosition value=" & cmdGetRecord.Parameters("recordCount").Value + 1 & ">" & vbCrLf)
								
					lngRecordID = 0
				Else
					Response.Write("<INPUT type='hidden' id=txtRecordID name=txtRecordID value=" & cmdGetRecord.Parameters("recordID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=" & cmdGetRecord.Parameters("recordCount").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtRecordPosition name=txtRecordPosition value=" & cmdGetRecord.Parameters("recordPosition").Value & ">" & vbCrLf)
				
					lngRecordID = cmdGetRecord.Parameters("recordID").Value
				End If
			Else
				Response.Write("<INPUT type='hidden' id=txtRecordID name=txtRecordID value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordPosition name=txtRecordPosition value=0>" & vbCrLf)
				
				lngRecordID = 0
			End If

			' Release the ADO command object.
			cmdGetRecord = Nothing
		End If
		
		' Get the record description.
		Dim sRecDesc = ""
		If (Len(sErrorDescription) = 0) Then
			Dim cmdGetRecordDesc = CreateObject("ADODB.Command")
			cmdGetRecordDesc.CommandText = "sp_ASRIntGetRecordDescription"
			cmdGetRecordDesc.CommandType = 4	' Stored procedure
			cmdGetRecordDesc.ActiveConnection = Session("databaseConnection")

			Dim prmTableId = cmdGetRecordDesc.CreateParameter("tableID", 3, 1) ' 3 = integer, 1 = input
			cmdGetRecordDesc.Parameters.Append(prmTableId)
			prmTableId.value = CleanNumeric(Session("tableID"))

			Dim prmRecordId = cmdGetRecordDesc.CreateParameter("recordID", 3, 1)	' 3 = integer, 1 = input
			cmdGetRecordDesc.Parameters.Append(prmRecordId)
			prmRecordId.value = CleanNumeric(lngRecordID)

			Dim prmParentTableId = cmdGetRecordDesc.CreateParameter("parentTableID", 3, 1) ' 3 = integer, 1 = input
			cmdGetRecordDesc.Parameters.Append(prmParentTableId)
			prmParentTableId.value = CleanNumeric(Session("parentTableID"))

			Dim prmParentRecordId = cmdGetRecordDesc.CreateParameter("parentRecordID", 3, 1)	' 3=integer, 1=input
			cmdGetRecordDesc.Parameters.Append(prmParentRecordId)
			prmParentRecordId.value = CleanNumeric(Session("parentRecordID"))

			Dim prmRecordDesc = cmdGetRecordDesc.CreateParameter("recordDesc", 200, 2, 8000)	' 200=varchar, 2=output, 8000=size
			cmdGetRecordDesc.Parameters.Append(prmRecordDesc)

			Dim fOk = True
			Dim fDeadlock = True
			Dim sErrMsg As String
			Do While fDeadlock
				fDeadlock = False

				cmdGetRecordDesc.ActiveConnection.Errors.Clear()

				cmdGetRecordDesc.Execute()

				If cmdGetRecordDesc.ActiveConnection.Errors.Count > 0 Then
					For iLoop = 1 To cmdGetRecordDesc.ActiveConnection.Errors.Count
						sErrMsg = formatError(cmdGetRecordDesc.ActiveConnection.Errors.Item(iLoop - 1).Description)

						If (cmdGetRecordDesc.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
						 (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
						  (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
						 ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
									 (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
							' The error is for a deadlock.
							' Sorry about having to use the err.description to trap the error but the err.number
							' is not specific and MSDN suggests using the err.description.
							If (iRetryCount < iRETRIES) And (cmdGetRecordDesc.ActiveConnection.Errors.Count = 1) Then
								iRetryCount = iRetryCount + 1
								fDeadlock = True
							Else
								If Len(sErrorDescription) > 0 Then
									sErrorDescription = sErrorDescription & vbCrLf
								End If
								sErrorDescription = sErrorDescription & "Another user is deadlocking the database. Please try again."
								fOk = False
							End If
						Else
							sErrorDescription = sErrorDescription & vbCrLf & _
							 formatError(cmdGetRecordDesc.ActiveConnection.Errors.Item(iLoop - 1).Description)
							fOk = False
						End If
					Next

					cmdGetRecordDesc.ActiveConnection.Errors.Clear()
												
					If Not fOk Then
						sErrorDescription = "Unable to get the record description." & vbCrLf & sErrorDescription
					End If
				End If
			Loop
				
			If Len(sErrorDescription) = 0 Then
				sRecDesc = cmdGetRecordDesc.Parameters("recordDesc").Value
			End If
					
			cmdGetRecordDesc = Nothing
		End If
		
		Response.Write("<INPUT type='hidden' id=txtRecordDescription name=txtRecordDescription value=""" & Replace(sRecDesc, """", "&quot;") & """>" & vbCrLf)
	Else
		Response.Write("<INPUT type='hidden' id=txtOriginalRecID name=txtOriginalRecID value=0>" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtNewRecID name=txtNewRecID value=0>" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtRecordDescription name=txtRecordDescription value="""">" & vbCrLf)
	End If

	If Session("action") = "LOADREPORTCOLUMNS" Then
		Dim cmdReportsCols = CreateObject("ADODB.Command")
		cmdReportsCols.CommandText = "sp_ASRIntGetReportColumns"
		cmdReportsCols.CommandType = 4 ' Stored procedure
		cmdReportsCols.ActiveConnection = Session("databaseConnection")
								
		Dim prmBaseTableId = cmdReportsCols.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
		cmdReportsCols.Parameters.Append(prmBaseTableId)
		prmBaseTableId.value = CleanNumeric(Session("ReportBaseTableID"))

		Dim prmParent1TableId = cmdReportsCols.CreateParameter("parent1TableID", 3, 1) ' 3=integer, 1=input
		cmdReportsCols.Parameters.Append(prmParent1TableId)
		prmParent1TableId.value = CleanNumeric(Session("ReportParent1TableID"))

		Dim prmParent2TableId = cmdReportsCols.CreateParameter("parent2TableID", 3, 1) ' 3=integer, 1=input
		cmdReportsCols.Parameters.Append(prmParent2TableId)
		prmParent2TableId.value = CleanNumeric(Session("ReportParent2TableID"))

		Dim prmChildTableId = cmdReportsCols.CreateParameter("childTableID", 200, 1, 8000) ' 200=varchar 1=input
		cmdReportsCols.Parameters.Append(prmChildTableId)
		prmChildTableId.value = Session("ReportChildTableID")

		Err.Clear()
		Dim rstReportColumns = cmdReportsCols.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "Error getting the report columns." & vbCrLf & formatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			Dim iLoop = 1
			Do While Not rstReportColumns.EOF
				Response.Write("<INPUT type='hidden' id=txtRepCol_" & iLoop & " name=txtRepCol_" & iLoop & " value=""" & Replace(rstReportColumns.Fields("columnDefn").Value, """", "&quot;") & """>" & vbCrLf)
				rstReportColumns.MoveNext()
				iLoop = iLoop + 1
			Loop

			' Release the ADO recordset object.
			rstReportColumns.close()
		End If
				
		rstReportColumns = Nothing
		cmdReportsCols = Nothing
	
	ElseIf Session("action") = "LOADCALENDARREPORTCOLUMNS" Then
		Dim cmdReportsCols = CreateObject("ADODB.Command")
		cmdReportsCols.CommandText = "spASRIntGetCalendarReportColumns"
		cmdReportsCols.CommandType = 4 ' Stored procedure
		cmdReportsCols.ActiveConnection = Session("databaseConnection")
								
		Dim prmBaseTableId = cmdReportsCols.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
		cmdReportsCols.Parameters.Append(prmBaseTableId)
		prmBaseTableId.value = CleanNumeric(Session("ReportBaseTableID"))
		
		Dim prmEventTableId = cmdReportsCols.CreateParameter("eventTableID", 3, 1)	' 3=integer, 1=input
		cmdReportsCols.Parameters.Append(prmEventTableId)
		prmEventTableId.value = CleanNumeric(Session("ReportBaseTableID"))
		
		Err.Clear()
		Dim rstReportColumns = cmdReportsCols.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "Error getting the calendar report columns." & vbCrLf & formatError(Err.Description)
		End If
		
		If Len(sErrorDescription) = 0 Then
			Dim iLoop = 1
			Do While Not rstReportColumns.EOF
				Response.Write("<INPUT type='hidden' id=txtRepCol_" & rstReportColumns.Fields("columnid").Value & " name=txtRepCol_" & rstReportColumns.Fields("columnid").Value & " value=""" & Replace(rstReportColumns.Fields("columnName").Value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRepColDataType_" & rstReportColumns.Fields("columnid").Value & " name=txtRepColDataType_" & rstReportColumns.Fields("columnid").Value & " value='" & Replace(rstReportColumns.Fields("datatype").Value, """", "&quot;") & "'>" & vbCrLf)
				rstReportColumns.MoveNext()
				iLoop = iLoop + 1
			Loop

			' Release the ADO recordset object.
			rstReportColumns.close()
		End If
				
		rstReportColumns = Nothing
		cmdReportsCols = Nothing

	ElseIf Session("action") = "LOADEMAILDEFINITIONS" Then
		Dim cmdReportsCols = CreateObject("ADODB.Command")
		cmdReportsCols.CommandText = "sp_ASRIntGetEmailAddresses"
		cmdReportsCols.CommandType = 4 ' Stored procedure
		cmdReportsCols.ActiveConnection = Session("databaseConnection")
								
		Dim prmBaseTableId = cmdReportsCols.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
		cmdReportsCols.Parameters.Append(prmBaseTableId)
		prmBaseTableId.value = CleanNumeric(Session("ReportBaseTableID"))

		Err.Clear()
		Dim rstReportColumns = cmdReportsCols.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "Error getting the report columns." & vbCrLf & formatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			Dim iLoop = 1
			Do While Not rstReportColumns.EOF
				Response.Write("<INPUT type='hidden' id=txtEmail_" & iLoop & " name=txtEmail_" & iLoop & " value=""" & Replace(rstReportColumns.Fields("columnDefn").Value, """", "&quot;") & """>" & vbCrLf)
				rstReportColumns.MoveNext()
				iLoop = iLoop + 1
			Loop

			' Release the ADO recordset object.
			rstReportColumns.close()
		End If
				
		rstReportColumns = Nothing
		cmdReportsCols = Nothing

	ElseIf Session("action") = "GETEXPRESSIONRETURNTYPES" Then
		Dim sParam1 As String = CStr(Session("Param1"))
		Dim iCharIndex As Integer
		
		' Get the server DLL to test the expression definition
        Dim objExpression = New HR.Intranet.Server.Expression
        

		' Pass required info to the DLL
        objExpression.Username = Session("username").ToString()
        CallByName(objExpression, "Connection", CallType.Let, Session("databaseConnection"))
        
		Do While Len(sParam1) > 0
			iCharIndex = InStr(sParam1, ",")

			If iCharIndex >= 0 Then
				Dim sExprId As String = Left(sParam1, iCharIndex - 1)
				sParam1 = Mid(sParam1, iCharIndex + 1)

				Dim iReturnType = objExpression.ExistingExpressionReturnType(CLng(sExprId))

				Response.Write("<INPUT type='hidden' id=txtExprType_" & sExprId & " name=txtExprType_" & sExprId & " value=" & iReturnType & ">" & vbCrLf)
			End If
		Loop

		objExpression = Nothing
	
		'**********************************************************************************************
	ElseIf Session("action") = "LOADEVENTLOG" Then
				
		Dim objUtilities = Session("UtilitiesObject")
		
		Dim cmdEventLogRecords = CreateObject("ADODB.Command")
		cmdEventLogRecords.CommandText = "spASRIntGetEventLogRecords"
		cmdEventLogRecords.CommandType = 4 ' Stored procedure.
		cmdEventLogRecords.ActiveConnection = Session("databaseConnection")

		Dim prmError = cmdEventLogRecords.CreateParameter("error", 11, 2)	' 11=bit, 2=output
		cmdEventLogRecords.Parameters.Append(prmError)

		Dim prmUser = cmdEventLogRecords.CreateParameter("user", 200, 1, 8000)
		cmdEventLogRecords.Parameters.Append(prmUser)
		prmUser.value = Session("ELFilterUser")

		Dim prmType = cmdEventLogRecords.CreateParameter("type", 3, 1)
		cmdEventLogRecords.Parameters.Append(prmType)
		prmType.value = Session("ELFilterType")

		Dim prmStatus = cmdEventLogRecords.CreateParameter("status", 3, 1)
		cmdEventLogRecords.Parameters.Append(prmStatus)
		prmStatus.value = Session("ELFilterStatus")

		Dim prmMode = cmdEventLogRecords.CreateParameter("mode", 3, 1)
		cmdEventLogRecords.Parameters.Append(prmMode)
		prmMode.value = Session("ELFilterMode")

		Dim prmOrderColumn = cmdEventLogRecords.CreateParameter("orderColumn", 200, 1, 8000)
		cmdEventLogRecords.Parameters.Append(prmOrderColumn)
		prmOrderColumn.value = Session("ELOrderColumn")

		Dim prmOrderOrder = cmdEventLogRecords.CreateParameter("orderOrder", 200, 1, 8000)
		cmdEventLogRecords.Parameters.Append(prmOrderOrder)
		prmOrderOrder.value = Session("ELOrderOrder")

		Dim prmReqRecs = cmdEventLogRecords.CreateParameter("reqRecs", 3, 1)
		cmdEventLogRecords.Parameters.Append(prmReqRecs)
		prmReqRecs.value = cleanNumeric(Session("findRecords"))

		Dim prmIsFirstPage = cmdEventLogRecords.CreateParameter("isFirstPage", 11, 2)	' 11=bit, 2=output
		cmdEventLogRecords.Parameters.Append(prmIsFirstPage)

		Dim prmIsLastPage = cmdEventLogRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
		cmdEventLogRecords.Parameters.Append(prmIsLastPage)

		Dim prmAction = cmdEventLogRecords.CreateParameter("action", 200, 1, 8000)
		cmdEventLogRecords.Parameters.Append(prmAction)
		prmAction.value = Session("ELAction")

		Dim prmTotalRecCount = cmdEventLogRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
		cmdEventLogRecords.Parameters.Append(prmTotalRecCount)

		Dim prmFirstRecPos = cmdEventLogRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
		cmdEventLogRecords.Parameters.Append(prmFirstRecPos)
		prmFirstRecPos.value = cleanNumeric(Session("ELFirstRecPos"))

		Dim prmCurrentRecCount = cmdEventLogRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
		cmdEventLogRecords.Parameters.Append(prmCurrentRecCount)
		prmCurrentRecCount.value = cleanNumeric(Session("ELCurrentRecCount"))
	
		Err.Clear()
		Dim rsEventLogRecords = cmdEventLogRecords.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "Error getting the event log records." & vbCrLf & formatError(Err.Description)
		End If

		Dim lngRowCount = 0
		If Len(sErrorDescription) = 0 Then
			
			Dim sAddString = vbNullString
			
			If Not (rsEventLogRecords.BOF And rsEventLogRecords.EOF) Then
				Do Until rsEventLogRecords.EOF
					sAddString = vbNullString
					
					sAddString = sAddString & rsEventLogRecords.Fields("ID").Value & vbTab
					
					sAddString = sAddString & convertSQLDateToLocale(rsEventLogRecords.Fields("DateTime").Value) & " " & convertSQLDateToTime(rsEventLogRecords.Fields("DateTime").Value) & vbTab
					
					If IsDBNull(rsEventLogRecords.Fields("EndTime").Value) Then
						sAddString = sAddString & "" & vbTab
					Else
						sAddString = sAddString & ConvertSqlDateToLocale(rsEventLogRecords.Fields("EndTime").Value) & " " & ConvertSqlDateToTime(rsEventLogRecords.Fields("EndTime").Value) & vbTab
					End If
						
                    sAddString = sAddString & objUtilities.FormatEventDuration(CLng(rsEventLogRecords.Fields("Duration").Value)) & vbTab
					
					sAddString = sAddString & Replace(rsEventLogRecords.Fields("EventInfo").Value, """", "&quot;")
					
					Response.Write("<INPUT type='hidden' id=txtAddString_" & lngRowCount & " name=txtAddString_" & lngRowCount & " value=""" & sAddString & """>" & vbCrLf)

					lngRowCount = lngRowCount + 1
					rsEventLogRecords.MoveNext()
				Loop
			End If
			
		End If
		
		rsEventLogRecords.close()
		rsEventLogRecords = Nothing
		
		Response.Write("<INPUT type='hidden' id=txtELIsFirstPage name=txtELIsFirstPage value=" & cmdEventLogRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtELIsLastPage name=txtELIsLastPage value=" & cmdEventLogRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtELRecordCount name=txtELRecordCount value=" & lngRowCount & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtELTotalRecordCount name=txtELTotalRecordCount value=" & cmdEventLogRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtELFindRecords name=txtELFindRecords value=" & Session("findRecords") & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtELFirstRecPos name=txtELFirstRecPos value=" & cmdEventLogRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtELCurrentRecCount name=txtELCurrentRecCount value=" & lngRowCount & ">" & vbCrLf)

		cmdEventLogRecords = Nothing
		objUtilities = Nothing
		
	ElseIf Session("action") = "LOADEVENTLOGUSERS" Then
		'Purge the event log.
		Dim cmdPurgeCommand = CreateObject("ADODB.Command")
		cmdPurgeCommand.CommandText = "sp_AsrEventLogPurge"
		cmdPurgeCommand.CommandType = 4 ' Stored procedure
		cmdPurgeCommand.ActiveConnection = Session("databaseConnection")
		Err.Clear()
		cmdPurgeCommand.Execute()
		cmdPurgeCommand = Nothing
		
		'Get the list of users
		Dim cmdEventLogUsers = CreateObject("ADODB.Command")
		cmdEventLogUsers.CommandText = "spASRIntGetEventLogUsers"
		cmdEventLogUsers.CommandType = 4	' Stored procedure
		cmdEventLogUsers.ActiveConnection = Session("databaseConnection")
		
		Err.Clear()
		Dim rstEventLogUsers = cmdEventLogUsers.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "Error getting the event log users." & vbCrLf & formatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			Dim iLoop = 1
			Do While Not rstEventLogUsers.EOF
				Response.Write("<INPUT type='hidden' id=txtEventLogUser_" & iLoop & " name=txtEventLogUser_" & iLoop & " value=""" & Replace(rstEventLogUsers.Fields("Username").Value, """", "&quot;") & """>" & vbCrLf)
				rstEventLogUsers.MoveNext()
				iLoop = iLoop + 1
			Loop

			' Release the ADO recordset object.
			rstEventLogUsers.close()
		End If
				
		rstEventLogUsers = Nothing
		cmdEventLogUsers = Nothing
			
		'**********************************************************************************************
	End If

	Response.Write("<INPUT type='hidden' id=txtNumberOfBookings name=txtNumberOfBookings value=" & Session("numberOfBookings") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtTBErrorMessage name=txtTBErrorMessage value=""" & Session("tbErrorMessage") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtTBCourseTitle name=txtTBCourseTitle value=""" & Session("tbCourseTitle") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)

%>

</form>

</div>

<script type="text/javascript"> data_window_onload();</script>
