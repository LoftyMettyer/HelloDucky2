<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET.Code" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server.Expressions" %>
<%@ Import Namespace="HR.Intranet.Server.Structures" %>

<%		
	If Len(Session("recordID")) = 0 Then
		Session("recordID") = 0
	End If
%>

<script type="text/javascript">
	function data_window_onload() {		
		var frmData = document.getElementById("frmData");
		var frmGetData = document.getElementById("frmGetData");
		var frmMenuInfo = $("#frmMenuInfo")[0].children;
		var frmOptionArea = OpenHR.getForm("optionframeset", "frmGotoOption");
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
		window.location = "Login";
	}
	else {
		// Do nothing if the menu controls are not yet instantiated.
		if (frmMenuInfo != null) {
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
				
				//No errors and recordedit navigation. Reset warning 'Are you sure you want to leave this page?'.
				window.onbeforeunload = null;

				var sAction = frmData.txtAction.value;
				
				if ((sAction == "LOAD")	
					&& (frmData.txtOriginalRecordID.value != frmData.txtNewRecID.value)
					&& (frmData.txtOriginalRecordID.value != 0)) {
					
					menu_refreshMenu();	
					
					if (recEditForm.txtRecEditFilterSQL.value == "") {
						OpenHR.messageBox("The record saved is no longer in the current view");
					}
				}

				if (sAction == "LOGOFF") {
					window.location.href = frmMenuInfo.txtDefaultStartPage.value;
					return;	
				}

				if (sAction == "EXIT") {
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
							OpenHR.modalPrompt("Transfer bookings to another course ?", 4, '').then(function (answer) {
								var frmData1 = document.getElementById("frmData");
								var sCourseTitle1 = new String(frmData1.txtTBCourseTitle.value);

								if (answer == 6) { // Yes
									// Display the course selection page.
									menu_loadTransferCoursePage(sCourseTitle1);
								}
								if (answer == 7) { // No.
									menu_transferCourse(0, true);
								}
							});

							//We've now opened the modalPrompt, so execution stops here. modalPrompt will redirect to the applicable function: cancelBookingResponse

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
						frmGetData.txtOriginalRecordID.value = frmData.txtOriginalRecordID.value;
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
							frmGetData.txtOriginalRecordID.value = frmData.txtOriginalRecordID.value;
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

					frmOptionArea.action = "emptyoption_Submit";
					OpenHR.submitForm(frmOptionArea, "optionframe");
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
					// We've got an error so don't update the find form.
					OpenHR.messageBox(sErrorMsg);
			
					if (frmData.txtAction.value == "SAVEERROR") {
						return false;
					}

					// Get menu to refresh the menu.
					menu_refreshMenu();
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


				if (frmData.txtAction.value == "REFRESHFINDAFTERINSERT") {
					
					//we're reloading after creating an inline new record							
					var newID = '<%:Session("recordID")%>';
					if (Number(newID) > 0) {										
						$("#findGridTable").jqGrid('setCell', '0', 'ID', newID);
					}
				}


			}
			else if ((sCurrentWorkPage == "UTIL_DEF_CUSTOMREPORTS") ||
					 (sCurrentWorkPage == "UTIL_DEF_CALENDARREPORT") ||
							 (sCurrentWorkPage == "UTIL_DEF_CROSSTABS") ||
							 (sCurrentWorkPage == "UTIL_DEF_9BOXGRID") ||
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


	function data_refreshData() {		
		var f = document.getElementById("frmGetData");
		OpenHR.submitForm(f);
	}

</script>

<div>

<form action="data_submit" method="post" id="frmGetData" name="frmGetData" data-formname="data.ascx">
		<input type="hidden" id="txtAction" name="txtAction">
		<input type="hidden" id="txtReaction" name="txtReaction">
		<input type="hidden" id="txtCurrentTableID" name="txtCurrentTableID">
		<input type="hidden" id="txtCurrentScreenID" name="txtCurrentScreenID">
		<input type="hidden" id="txtCurrentViewID" name="txtCurrentViewID">
		<input type="hidden" id="txtSelectSQL" name="txtSelectSQL">
		<input type="hidden" id="txtFromDef" name="txtFromDef">
		<input type="hidden" id="txtFilterSQL" name="txtFilterSQL">
		<input type="hidden" id="txtFilterDef" name="txtFilterDef">
		<input type="hidden" id="txtRealSource" name="txtRealSource">
		<input type="hidden" id="txtOriginalRecordID" name="txtOriginalRecordID">
		<input type="hidden" id="txtRecordID" name="txtRecordID">
		<input type="hidden" id="txtParentTableID" name="txtParentTableID">
		<input type="hidden" id="txtParentRecordID" name="txtParentRecordID">
		<input type="hidden" id="txtDefaultCalcCols" name="txtDefaultCalcCols">
		<input type="hidden" id="txtInsertUpdateDef" name="txtInsertUpdateDef">
		<input type="hidden" id="txtTimestamp" name="txtTimestamp">
		<input type="hidden" id="txtTBCourseRecordID" name="txtTBCourseRecordID">
		<input type="hidden" id="txtTBEmployeeRecordID" name="txtTBEmployeeRecordID">
		<input type="hidden" id="txtTBBookingStatusValue" name="txtTBBookingStatusValue">
		<input type="hidden" id="txtTBOverride" name="txtTBOverride">
		<input type="hidden" id="txtTBCreateWLRecords" name="txtTBCreateWLRecords">
		<input type="hidden" id="txtReportBaseTableID" name="txtReportBaseTableID">
		<input type="hidden" id="txtReportParent1TableID" name="txtReportParent1TableID">
		<input type="hidden" id="txtReportParent2TableID" name="txtReportParent2TableID">
		<input type="hidden" id="txtReportChildTableID" name="txtReportChildTableID">
		<input type="hidden" id="txtUserChoice" name="txtUserChoice">
		<input type="hidden" id="txtParam1" name="txtParam1">
		<input type="hidden" id="txtELFilterUser" name="txtELFilterUser">
		<input type="hidden" id="txtELFilterType" name="txtELFilterType">
		<input type="hidden" id="txtELFilterStatus" name="txtELFilterStatus">
		<input type="hidden" id="txtELFilterMode" name="txtELFilterMode">
		<input type="hidden" id="txtELOrderColumn" name="txtELOrderColumn">
		<input type="hidden" id="txtELOrderOrder" name="txtELOrderOrder">
		<input type="hidden" id="txtELAction" name="txtELAction">
		<input type="hidden" id="txtELCurrRecCount" name="txtELCurrRecCount" value="0">
		<input type="hidden" id="txtEL1stRecPos" name="txtEL1stRecPos" value="0">
		<%=Html.AntiForgeryToken()%>
</form>

<form id="frmData" name="frmData">
<%
	
	Dim lngRecordID As Long
	
	Dim sErrorDescription = ""
	Dim SPParameters() As SqlParameter
	
	Dim objSessionInfo As SessionInfo = CType(Session("SessionContext"), SessionInfo)
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
	
	Response.Write("<input type='hidden' id='txtAction' name='txtAction' value='" & Session("action") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtParentTableID' name='txtParentTableID' value='" & Session("parentTableID") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtParentRecordID' name='txtParentRecordID' value='" & Session("parentRecordID") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtErrorMessage' name='txtErrorMessage' value=""" & Replace(Session("errorMessage"), """", "'") & """>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtWarning' name='txtWarning' value='" & Session("warningFlag") & "'>" & vbCrLf)
	' Clear the error message session variable.
	Session("errorMessage") = ""

	' Get the required record count if we have a query.
	If Len(Session("selectSQL")) > 0 Then
		If Session("action") = "NEW" Then

			Dim prmRecordCount As New SqlParameter("piRecordCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		
			Try
							
				Dim rstRecord = objDataAccess.GetDataTable("sp_ASRIntCalcDefaults", CommandType.StoredProcedure _
						, prmRecordCount _
						, New SqlParameter("psFromDef", SqlDbType.VarChar, -1) With {.Value = Session("fromDef")} _
						, New SqlParameter("psFilterDef", SqlDbType.VarChar, -1) With {.Value = Session("filterDef")} _
						, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))} _
						, New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))} _
						, New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))} _
						, New SqlParameter("psDefaultCalcColumns", SqlDbType.VarChar, -1) With {.Value = Session("defaultCalcColumns")} _
						, New SqlParameter("psDecimalSeparator", SqlDbType.VarChar, 255) With {.Value = Session("LocaleDecimalSeparator")} _
						, New SqlParameter("psLocaleDateFormat", SqlDbType.VarChar, 255) With {.Value = Platform.LocaleDateFormatForSQL()})
		
				If Not rstRecord Is Nothing Then
					If rstRecord.Rows.Count > 0 Then
						For iloop = 0 To (rstRecord.Columns.Count - 1)
							If IsDBNull(rstRecord(iloop)) Then
								Response.Write("<input type='hidden' id='txtData_" & rstRecord.Columns(iloop).ColumnName & "' name='txtData_" & rstRecord.Columns(iloop).ColumnName & "' value=''>" & vbCrLf)
							Else
								Response.Write("<input type='hidden' id='txtData_" & rstRecord.Columns(iloop).ColumnName & "' name='txtData_" & rstRecord.Columns(iloop).ColumnName & "' value='" & Replace(rstRecord(0)(iloop).ToString(), """", "&quot;") & "'>" & vbCrLf)
							End If
						Next
					End If
				End If
				
			Catch ex As Exception
				sErrorDescription = "Default values could not be calculated." & vbCrLf & FormatError(ex.Message)

			End Try

		
			If Session("parentTableID") > 0 Then
				
				Try
					
					Dim rstParentValues = objDataAccess.GetDataTable("spASRIntGetParentValues", CommandType.StoredProcedure _
						, New SqlParameter("piScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))} _
						, New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))} _
						, New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))})

					If Not rstParentValues Is Nothing Then
						If rstParentValues.Rows.Count > 0 Then
							For iloop = 0 To (rstParentValues.Columns.Count - 1)
								If IsDBNull(rstParentValues(iloop)) Then
									Response.Write("<input type='hidden' id='txtData_" & rstParentValues.Columns(iloop).ColumnName & "' name='txtData_" & rstParentValues.Columns(iloop).ColumnName & "' value=''>" & vbCrLf)
								Else
									Response.Write("<input type='hidden' id='txtData_" & rstParentValues.Columns(iloop).ColumnName & "' name='txtData_" & rstParentValues.Columns(iloop).ColumnName & "' value='" & Replace(rstParentValues(0)(iloop).ToString(), """", "&quot;") & "'>" & vbCrLf)
								End If
							Next
						End If
					End If
										
				Catch ex As Exception
					sErrorDescription = "Parent values could not be determined." & vbCrLf & FormatError(ex.Message)

				End Try


					
			End If

			If Len(sErrorDescription) = 0 Then
				Response.Write("<input type='hidden' id='txtRecordID' name='txtRecordID' value='0'>" & vbCrLf)
				Response.Write("<input type='hidden' id='txtRecordCount' name='txtRecordCount' value='" & prmRecordCount.Value & "'>" & vbCrLf)
				Response.Write("<input type='hidden' id='txtRecordPosition' name='txtRecordPosition' value='" & prmRecordCount.Value + 1 & "'>" & vbCrLf)
			Else
				Response.Write("<input type='hidden' id='txtRecordID' name='txtRecordID' value='0'>" & vbCrLf)
				Response.Write("<input type='hidden' id='txtRecordCount' name='txtRecordCount' value='0'>" & vbCrLf)
				Response.Write("<input type='hidden' id='txtRecordPosition' name='txtRecordPosition' value='0'>" & vbCrLf)
			End If
			
			lngRecordID = 0
			Response.Write("<input type='hidden' id='txtOriginalRecordID' name='txtOriginalRecordID' value='0'>" & vbCrLf)
			Response.Write("<input type='hidden' id='txtNewRecID' name='txtNewRecID' value='0'>" & vbCrLf)
		Else
			
			
			Dim prmRecordId = New SqlParameter("piRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("recordID"))}
			Dim prmRecordCount = New SqlParameter("piRecordCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmRecordPosition = New SqlParameter("piRecordPosition", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFilterDef = New SqlParameter("psFilterDef", SqlDbType.VarChar, -1) With {.Value = Session("filterDef")}
			Dim prmAction = New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("action")}
			Dim prmParentTableId = New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))}
			Dim prmParentRecordId = New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))}
			Dim prmDecSeparator = New SqlParameter("psDecimalSeparator", SqlDbType.VarChar, 100) With {.Value = Session("LocaleDecimalSeparator")}
			Dim prmDateFormat = New SqlParameter("psLocaleDateFormat", SqlDbType.VarChar, 100) With {.Value = Platform.LocaleDateFormatForSQL()}
			Dim prmScreenId = New SqlParameter("piScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))}
			Dim prmViewId = New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))}
			Dim prmOrderId = New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("orderID"))}


			Dim oleColumnData As New List(Of OLEValue)

			Try

				SPParameters = New SqlParameter() {prmRecordId, prmRecordCount, prmRecordPosition, prmFilterDef, _
						prmAction, prmParentTableId, prmParentRecordId, prmDecSeparator, prmDateFormat, prmScreenId, prmViewId, prmOrderId}

				Dim rstRecord = objDataAccess.GetFromSP("sp_ASRIntGetRecord", SPParameters)
					
				For Each objRow As DataRow In rstRecord.Rows
					For iloop = 0 To (rstRecord.Columns.Count - 1)
						
						If IsDBNull(objRow(iloop)) Then
							Response.Write("<input type='hidden' id='txtData_" & rstRecord.Columns(iloop).ColumnName & "' name='txtData_" & rstRecord.Columns(iloop).ColumnName & "' value=''>" & vbCrLf)
						Else

							' Is column a embedded/linked OLE								
							If rstRecord.Columns(iloop).DataType.ToString().ToLower = "system.byte[]" Then
								
								Dim bytValue() As Byte = CType(objRow(iloop), Byte())
								Dim oleVersion As String = Encoding.UTF8.GetString(bytValue, 0, 8)
								
								If oleVersion = "<<V002>>" Then
									Dim objColumnOLE As New OLEValue With {
										.ColumnID = CInt(rstRecord.Columns(iloop).ColumnName), _
										.Value = bytValue}
									oleColumnData.Add(objColumnOLE)
								Else
									' Incorrect header version for column. Treat as empty.
									Response.Write("<input type='hidden' id='txtData_" & rstRecord.Columns(iloop).ColumnName & "' name='txtData_" & rstRecord.Columns(iloop).ColumnName & "' value=''>" & vbCrLf)
								End If
								
							Else
								Response.Write("<input type='hidden' id='txtData_" & rstRecord.Columns(iloop).ColumnName & "' name='txtData_" & rstRecord.Columns(iloop).ColumnName & "' value='" & Html.Encode(objRow(iloop).ToString()) & "'>" & vbCrLf)
							End If
						End If
					Next
				Next


				Response.Write("<input type='hidden' id='txtOriginalRecordID' name='txtOriginalRecordID' value='" & Session("recordID") & "'>" & vbCrLf)
				Response.Write("<input type='hidden' id='txtNewRecID' name='txtNewRecID' value='" & prmRecordId.Value.ToString() & "'>" & vbCrLf)
			
				' Loop through each of the OLE (Documents and photos and render)
				For Each item As OLEValue In oleColumnData
				
					If item.Value Is Nothing Then
						Response.Write("<input type='hidden' id='txtData_" & item.ColumnID & "' name='txtData_" & item.ColumnID & "' value=''>" & vbCrLf)
					Else

						item.ExtractProperties()
				
						If objSessionInfo.IsPhotoDataType(item.ColumnID) Then
					
							Response.Write("<input type='hidden' id='txtData_" & item.ColumnID & "' data-Img='" & item.ConvertPhotoToBase64() & "' name='txtData_" & item.ColumnID & "' value='" & item.FileName & "'>" & vbCrLf)
						Else
							Response.Write("<input type='hidden' id='txtData_" & item.ColumnID & "' name='txtData_" & item.ColumnID & "' data-filesize='" & item.DocumentSize & "' data-filecreatedate='" & item.FileCreateDate & "' data-filemodifydate='" & item.FileModifyDate & "' value='" & item.FileName & "'>" & vbCrLf)
						End If
					End If
				
				Next

			Catch ex As Exception
				sErrorDescription = "Unable to retrieve the required record." & vbCrLf & ex.Message

			End Try


			If Len(sErrorDescription) = 0 Then
				If Session("action") = "COPY" Then
					Response.Write("<input type='hidden' id='txtRecordID' name='txtRecordID' value='0'>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtRecordCount' name='txtRecordCount' value='" & prmRecordCount.Value.ToString() & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtRecordPosition' name='txtRecordPosition' value='" & prmRecordCount.Value.ToString() + 1 & "'>" & vbCrLf)
								
					lngRecordID = 0
				Else
					Response.Write("<input type='hidden' id='txtRecordID' name='txtRecordID' value='" & prmRecordId.Value.ToString() & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtRecordCount' name='txtRecordCount' value='" & prmRecordCount.Value.ToString() & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtRecordPosition' name='txtRecordPosition' value='" & prmRecordPosition.Value.ToString() & "'>" & vbCrLf)
				
					lngRecordID = CInt(prmRecordId.Value)
					Session("PreviousRecordID") = lngRecordID
				End If
			Else
				Response.Write("<input type='hidden' id='txtRecordID' name='txtRecordID' value='0'>" & vbCrLf)
				Response.Write("<input type='hidden' id='txtRecordCount' name='txtRecordCount' value='0'>" & vbCrLf)
				Response.Write("<input type='hidden' id='txtRecordPosition' name='txtRecordPosition' value='0'>" & vbCrLf)
				
				lngRecordID = 0
			End If

		End If
		
		' Get the record description.
		Dim sRecDesc As String = ""
		If (Len(sErrorDescription) = 0) Then
								
			Try

				Dim prmRecordDesc As New SqlParameter("psRecDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			
				objDataAccess.ExecuteSP("sp_ASRIntGetRecordDescription", _
						New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))}, _
						New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = lngRecordID}, _
						New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))}, _
						New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))}, _
						prmRecordDesc)

				sRecDesc = prmRecordDesc.Value.ToString()

			Catch ex As Exception
				sRecDesc = ""
				
			End Try
			
		End If
				
		Response.Write("<input type='hidden' id='txtRecordDescription' name='txtRecordDescription' value='" & Html.Encode(sRecDesc) & "'>" & vbCrLf)
	Else
		Response.Write("<input type='hidden' id='txtOriginalRecordID' name='txtOriginalRecordID' value='0'>" & vbCrLf)
		Response.Write("<input type='hidden' id='txtNewRecID' name='txtNewRecID' value='0'>" & vbCrLf)
		Response.Write("<input type='hidden' id='txtRecordDescription' name='txtRecordDescription' value=''>" & vbCrLf)
	End If
	
	If Session("action") = "GETEXPRESSIONRETURNTYPES" Then
		Dim sParam1 As String = CStr(Session("Param1"))
		Dim iCharIndex As Integer
		
		' Get the server DLL to test the expression definition
		Dim objExpression = New Expression(objSessionInfo)
				
		' Pass required info to the DLL			
		Do While Len(sParam1) > 0
			iCharIndex = InStr(sParam1, ",")

			If iCharIndex >= 0 Then
				Dim sExprId As String = Left(sParam1, iCharIndex - 1)
				sParam1 = Mid(sParam1, iCharIndex + 1)

				Dim iReturnType As Short = objExpression.ExistingExpressionReturnType(CInt(sExprId))

				Response.Write("<input type='hidden' id='txtExprType_" & sExprId & "' name='txtExprType_" & sExprId & "' value='" & iReturnType & "'>" & vbCrLf)
			End If
		Loop

		objExpression = Nothing
	
		'**********************************************************************************************
	ElseIf Session("action") = "LOADEVENTLOG" Then
				
		Try
			
			Dim prmError As New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmIsFirstPage As New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmIsLastPage As New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmTotalRecCount As New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFirstRecPos As New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("ELFirstRecPos"))}
				
			Dim rsEventLogRecords = objDataAccess.GetDataTable("spASRIntGetEventLogRecords", CommandType.StoredProcedure _
				, prmError _
				, New SqlParameter("psFilterUser", SqlDbType.VarChar, -1) With {.Value = Session("ELFilterUser")} _
				, New SqlParameter("piFilterType", SqlDbType.Int) With {.Value = Session("ELFilterType")} _
				, New SqlParameter("piFilterStatus", SqlDbType.Int) With {.Value = Session("ELFilterStatus")} _
				, New SqlParameter("piFilterMode", SqlDbType.Int) With {.Value = Session("ELFilterMode")} _
				, New SqlParameter("psOrderColumn", SqlDbType.VarChar, -1) With {.Value = Session("ELOrderColumn")} _
				, New SqlParameter("psOrderOrder", SqlDbType.VarChar, -1) With {.Value = Session("ELOrderOrder")} _
				, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = 100000} _
				, prmIsFirstPage _
				, prmIsLastPage _
				, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("ELAction")} _
				, prmTotalRecCount _
				, prmFirstRecPos _
				, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("ELCurrentRecCount"))})
		
			Dim lngRowCount = 0
			If Len(sErrorDescription) = 0 Then
			
				Dim sAddString As String = vbNullString
			
				For Each objRow As DataRow In rsEventLogRecords.Rows

					sAddString = vbNullString
					sAddString = sAddString & objRow("ID").ToString() & vbTab
					sAddString = sAddString & ConvertSQLDateToLocale(objRow("DateTime")) & " " & ConvertSqlDateToTime(objRow("DateTime")) & vbTab
					
					If IsDBNull(objRow("EndTime")) Then
						sAddString = sAddString & "" & vbTab
					Else
						sAddString = sAddString & ConvertSQLDateToLocale(objRow("EndTime")) & " " & ConvertSqlDateToTime(objRow("EndTime")) & vbTab
					End If
						
					sAddString = sAddString & FormatEventDuration(CInt(objRow("Duration"))) & vbTab
					
					sAddString = sAddString & Replace(objRow("EventInfo").ToString(), """", "&quot;")
					
					Response.Write("<input type='hidden' id='txtAddString_" & lngRowCount & "' name='txtAddString_" & lngRowCount & "' value='" & sAddString.Replace("'", "&#39") & "'>" & vbCrLf)

					lngRowCount += 1

				Next
			End If
						
			Response.Write("<input type='hidden' id='txtELIsFirstPage' name='txtELIsFirstPage' value='" & prmIsFirstPage.Value & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id='txtELIsLastPage' name='txtELIsLastPage' value='" & prmIsLastPage.Value & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id='txtELRecordCount' name='txtELRecordCount' value='" & lngRowCount & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id='txtELTotalRecordCount' name='txtELTotalRecordCount' value='" & prmTotalRecCount.Value & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id='txtELFindRecords' name='txtELFindRecords' value='" & Session("findRecords") & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id='txtELFirstRecPos' name='txtELFirstRecPos' value='" & prmFirstRecPos.Value & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id='txtELCurrentRecCount' name='txtELCurrentRecCount' value='" & lngRowCount & "'>" & vbCrLf)

		
		Catch ex As Exception
			sErrorDescription = "Error getting the event log records." & vbCrLf & FormatError(ex.Message)
	
		End Try

			
	ElseIf Session("action") = "LOADEVENTLOGUSERS" Then

		Try
	
			'Purge the event log.
			objDataAccess.ExecuteSP("sp_AsrEventLogPurge")
	
			Dim rstEventLogUsers = objDataAccess.GetDataTable("spASRIntGetEventLogUsers", CommandType.StoredProcedure)

			Dim iLoop = 1
			For Each objRow As DataRow In rstEventLogUsers.Rows
				Response.Write("<input type='hidden' id='txtEventLogUser_" & iLoop & "' name='txtEventLogUser_" & iLoop & "' value='" & Replace(objRow("Username").ToString(), """", "&quot;") & "'>" & vbCrLf)
				iLoop += 1
			Next

			
		Catch ex As Exception
			sErrorDescription = "Error getting the event log users." & vbCrLf & ex.Message.RemoveSensitive

		End Try

			
		'**********************************************************************************************
	End If

	Response.Write("<input type='hidden' id='txtNumberOfBookings' name='txtNumberOfBookings' value='" & Session("numberOfBookings") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtTBErrorMessage' name='txtTBErrorMessage' value='" & Session("tbErrorMessage") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtTBCourseTitle' name='txtTBCourseTitle' value=""" & Session("tbCourseTitle") & """>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtErrorDescription' name='txtErrorDescription' value='" & sErrorDescription & "'>" & vbCrLf)

%>

</form>

</div>

<script type="text/javascript">data_window_onload();</script>
