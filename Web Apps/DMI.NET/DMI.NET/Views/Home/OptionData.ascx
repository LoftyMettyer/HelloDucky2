<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>



<script type="text/javascript">
<!--
    function optiondata_onload() {

        var sFatalErrorMsg = frmOptionData.txtErrorDescription.value
        if (sFatalErrorMsg.length > 0) {
            //window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sFatalErrorMsg);
            //window.parent.location.replace("login.asp");
        } else {
            // Do nothing if the menu controls are not yet instantiated.
            if (2 != 3) {
                var sCurrentWorkPage = OpenHR.currentWorkPage();

                if (sCurrentWorkPage == "LINKFIND") {
                    var sErrorMsg = frmOptionData.txtErrorMessage.value;
                    if (sErrorMsg.length > 0) {
                        // We've got an error so don't update the record edit form.

                        // Get menu.asp to refresh the menu.
                        menu_refreshMenu();                        
                        OpenHR.messageBox(sErrorMsg);
                    }

                    var sAction = frmOptionData.txtOptionAction.value;

                    // Refresh the link find grid with the data if required.
                    var grdLinkFind = OpenHR.getForm("optionframe","frmLinkFindForm").ssOleDBGridLinkRecords;

                    grdLinkFind.redraw = false;
                    grdLinkFind.removeAll();
                    grdLinkFind.columns.removeAll();

                    var dataCollection = frmOptionData.elements;
                    var sControlName;
                    var sColumnName;
                    var iColumnType;
                    var iCount;

                    // Configure the grid columns.
                    if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 16);
                            if (sControlName == "txtOptionColDef_") {
                                // Get the column name and type from the control.
                                sColDef = dataCollection.item(i).value;

                                iIndex = sColDef.indexOf("	");
                                if (iIndex >= 0) {
                                    sColumnName = sColDef.substr(0, iIndex);
                                    sColumnType = sColDef.substr(iIndex + 1);

                                    grdLinkFind.columns.add(grdLinkFind.columns.count);
                                    grdLinkFind.columns.item(grdLinkFind.columns.count - 1).name = sColumnName;
                                    grdLinkFind.columns.item(grdLinkFind.columns.count - 1).caption = sColumnName;

                                    if (sColumnName == "ID") {
                                        grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Visible = false;
                                    }

                                    if ((sColumnType == "131") || (sColumnType == "3")) {
                                        grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Alignment = 1;
                                    } else {
                                        grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Alignment = 0;
                                    }
                                    if (sColumnType == 11) {
                                        grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Style = 2;
                                    } else {
                                        grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Style = 0;
                                    }
                                }
                            }
                        }
                    }

                    // Add the grid records.
                    var sAddString;
                    var fRecordAdded;
                    fRecordAdded = false;
                    iCount = 0;
                    if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 14);
                            if (sControlName == "txtOptionData_") {
                                grdLinkFind.addItem(dataCollection.item(i).value);
                                fRecordAdded = true;
                                iCount = iCount + 1
                            }
                        }
                    }
                    grdLinkFind.redraw = true;

                    frmOptionData.txtRecordCount.value = iCount;

                    if (fRecordAdded == true) {
                        OpenHR.getFrame("optionframe").locateRecord(OpenHR.getForm("optionframe","frmLinkFindForm").txtOptionLinkRecordID.value, true);
                    }

                    OpenHR.getFrame("optionframe").refreshControls();
                    

                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();
                }

                if (sCurrentWorkPage == "LOOKUPFIND") {
                    var sErrorMsg = frmOptionData.txtErrorMessage.value;
                    if (sErrorMsg.length > 0) {
                        // We've got an error so don't update the record edit form.

                        // Get menu.asp to refresh the menu.
                        menu_refreshMenu();
                        OpenHR.messageBox(sErrorMsg);
                    }

                    if (frmOptionData.txtFilterOverride.value == "True")
                        // No access to the lookup filter column?
                    {
                        OpenHR.messageBox("You do not have 'read' permission on the lookup filter value column. No filter will be applied.");
                    }

                    var sAction = frmOptionData.txtOptionAction.value;

                    OpenHR.getFrame("optionframe").document.forms("frmLookupFindForm").txtLookupColumnGridPosition.value = frmOptionData.txtLookupColumnGridPosition.value;

                    // Refresh the link find grid with the data if required.
                    var grdFind = OpenHR.getForm("optionframe","frmLookupFindForm").ssOleDBGrid;                    

                    // Clear the grid.
                    grdFind.redraw = false;
                    grdFind.removeAll();
                    grdFind.columns.removeAll();

                    var dataCollection = frmOptionData.elements;
                    var sControlName;
                    var sColumnName;
                    var iColumnType;
                    var iCount;

                    // Configure the grid columns.
                    if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 16);
                            if (sControlName == "txtOptionColDef_") {
                                // Get the column name and type from the control.
                                sColDef = dataCollection.item(i).value;

                                iIndex = sColDef.indexOf("	");
                                if (iIndex >= 0) {
                                    sColumnName = sColDef.substr(0, iIndex);
                                    sColumnType = sColDef.substr(iIndex + 1);

                                    grdFind.columns.add(grdFind.columns.count);
                                    grdFind.columns.item(grdFind.columns.count - 1).name = sColumnName;
                                    grdFind.columns.item(grdFind.columns.count - 1).caption = sColumnName;

                                    if (sColumnName == "ID") {
                                        grdFind.columns.item(grdFind.columns.count - 1).Visible = false;
                                    }

                                    if ((sColumnType == "131") || (sColumnType == "3")) {
                                        grdFind.columns.item(grdFind.columns.count - 1).Alignment = 1;
                                    } else {
                                        grdFind.columns.item(grdFind.columns.count - 1).Alignment = 0;
                                    }
                                    if (sColumnType == 11) {
                                        grdFind.columns.item(grdFind.columns.count - 1).Style = 2;
                                    } else {
                                        grdFind.columns.item(grdFind.columns.count - 1).Style = 0;
                                    }
                                }
                            }
                        }
                    }

                    // Add the grid records.
                    var sAddString;
                    var fRecordAdded;
                    fRecordAdded = false;
                    iCount = 0;
                    if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 14);
                            if (sControlName == "txtOptionData_") {
                                grdFind.addItem(dataCollection.item(i).value);
                                fRecordAdded = true;
                                iCount = iCount + 1
                            }
                        }
                    }

                    grdFind.redraw = true;

                    frmOptionData.txtRecordCount.value = iCount;

                    if (fRecordAdded == true) {
                        OpenHR.getFrame("optionframe").locateRecord(OpenHR.getForm("optionframe","frmLookupFindForm").txtOptionLookupValue.value, true);
                    }

                    OpenHR.getFrame("optionframe").refreshControls();

                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();
                }

                if ((sCurrentWorkPage == "TBTRANSFERCOURSEFIND") ||
                    (sCurrentWorkPage == "TBBOOKCOURSEFIND") ||
                    (sCurrentWorkPage == "TBADDFROMWAITINGLISTFIND") ||
                    (sCurrentWorkPage == "TBTRANSFERBOOKINGFIND")) {
                    var sErrorMsg = frmOptionData.txtErrorMessage.value;
                    if (sErrorMsg.length > 0) {
                        // We've got an error.
                        // Get menu.asp to refresh the menu.
                        menu_refreshMenu();
                        OpenHR.messageBox(sErrorMsg);
                    }

                    if ((sCurrentWorkPage == "TBTRANSFERBOOKINGFIND") ||
                        (sCurrentWorkPage == "TBADDFROMWAITINGLISTFIND")) {
                        var sErrorMsg = frmOptionData.txtErrorMessage2.value;
                        if (sErrorMsg.length > 0) {
                            // We've got an error.
                            OpenHR.getFrame("optionframe").Cancel();
                            //window.parent.frames("menuframe").ASRIntranetFunctions.ClosePopup();
                            OpenHR.messageBox(sErrorMsg);
                            return;
                        }
                    }

                    var sAction = frmOptionData.txtOptionAction.value;

                    // Refresh the link find grid with the data if required.
                    var grdFind = OpenHR.getForm("optionframe","frmFindForm").ssOleDBGridRecords;
                    grdFind.redraw = false;
                    grdFind.removeAll();
                    grdFind.columns.removeAll();

                    var dataCollection = frmOptionData.elements;
                    var sControlName;
                    var sColumnName;
                    var iColumnType;
                    var iCount;

                    // Configure the grid columns.
                    if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 16);
                            if (sControlName == "txtOptionColDef_") {
                                // Get the column name and type from the control.
                                sColDef = dataCollection.item(i).value;

                                iIndex = sColDef.indexOf("	");
                                if (iIndex >= 0) {
                                    sColumnName = sColDef.substr(0, iIndex);
                                    sColumnType = sColDef.substr(iIndex + 1);

                                    grdFind.columns.add(grdFind.columns.count);
                                    grdFind.columns.item(grdFind.columns.count - 1).name = sColumnName;
                                    grdFind.columns.item(grdFind.columns.count - 1).caption = sColumnName;

                                    if (sColumnName == "ID") {
                                        grdFind.columns.item(grdFind.columns.count - 1).Visible = false;
                                    }

                                    if ((sColumnType == "131") || (sColumnType == "3")) {
                                        grdFind.columns.item(grdFind.columns.count - 1).Alignment = 1;
                                    } else {
                                        grdFind.columns.item(grdFind.columns.count - 1).Alignment = 0;
                                    }
                                    if (sColumnType == 11) {
                                        grdFind.columns.item(grdFind.columns.count - 1).Style = 2;
                                    } else {
                                        grdFind.columns.item(grdFind.columns.count - 1).Style = 0;
                                    }
                                }
                            }
                        }
                    }

                    // Add the grid records.
                    var sAddString;
                    var fRecordAdded;
                    fRecordAdded = false;
                    iCount = 0;
                    if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 14);
                            if (sControlName == "txtOptionData_") {
                                grdFind.addItem(dataCollection.item(i).value);
                                fRecordAdded = true;
                                iCount = iCount + 1
                            }
                        }
                    }

                    grdFind.redraw = true;

                    frmOptionData.txtRecordCount.value = iCount;

                    // Select the top record.
                    if (fRecordAdded == true) {
                        grdFind.MoveFirst();
                        grdFind.SelBookmarks.Add(grdFind.Bookmark);
                    }

                    OpenHR.getFrame("optionframe").refreshControls();

                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();
                }

                if (sCurrentWorkPage == "TBBULKBOOKING") {
                    var sAction = frmOptionData.txtOptionAction.value;

                    // Refresh the link find grid with the data if required.
                    var grdFind = OpenHR.getForm("optionframe","frmBulkBooking").ssOleDBGridFindRecords;
                    grdFind.redraw = false;
                    grdFind.removeAll();

                    var dataCollection = frmOptionData.elements;
                    var sControlName;
                    var sColumnName;
                    var iColumnType;
                    var iCount;

                    // Add the grid records.
                    var sAddString;
                    var fRecordAdded;
                    fRecordAdded = false;
                    iCount = 0;

                    if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 14);

                            if (sControlName == "txtOptionData_") {
                                grdFind.addItem(dataCollection.item(i).value);
                                fRecordAdded = true;
                                iCount = iCount + 1
                            }
                        }
                    }

                    grdFind.redraw = true;

                    // Select the top record.
                    if (fRecordAdded == true) {
                        grdFind.MoveFirst();
                        grdFind.SelBookmarks.Add(grdFind.Bookmark);
                    }

                    OpenHR.getFrame("optionframe").refreshControls();

                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();
                }

                if (sCurrentWorkPage == "UTIL_DEF_PICKLIST") {
                    var sAction = frmOptionData.txtOptionAction.value;

                    // Refresh the link find grid with the data if required.
                    var grdFind = OpenHR.getForm("workframe","frmDefinition").ssOleDBGrid;
                    grdFind.redraw = false;
                    grdFind.removeAll();

                    var dataCollection = frmOptionData.elements;
                    var sControlName;
                    var sColumnName;
                    var iColumnType;
                    var iCount;

                    // Add the grid records.
                    var sAddString;
                    var fRecordAdded;
                    fRecordAdded = false;
                    iCount = 0;

                    if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 14);

                            if (sControlName == "txtOptionData_") {
                                grdFind.addItem(dataCollection.item(i).value);
                                fRecordAdded = true;
                                iCount = iCount + 1
                            }
                        }
                    }

                    grdFind.redraw = true;

                    if (frmOptionData.txtExpectedCount.value > iCount) {
                        if (iCount == 0) {
                            OpenHR.messageBox("You do not have 'read' permission on any of the records in the selected picklist.\nUnable to display records.");
                            OpenHR.getForm("workframe","frmUseful").txtAction.value = "VIEW";
                            OpenHR.getFrame("workframe").cancelClick();
                        } else {
                            if (OpenHR.getForm("workframe","frmUseful").txtAction.value.toUpperCase() == "COPY") {
                                OpenHR.messageBox("You do not have 'read' permission on all of the records in the selected picklist.\nOnly permitted records will be shown.");
                            } else {
                                OpenHR.messageBox("You do not have 'read' permission on all of the records in the selected picklist.\nOnly permitted records will be shown and the definition will be read only.");
                                OpenHR.getForm("workframe","frmUseful").txtAction.value = "VIEW";
                                OpenHR.getFrame("workframe").disableAll();
                            }
                        }
                    }

                    // Select the top record.
                    if (fRecordAdded == true) {
                        grdFind.MoveFirst();
                        grdFind.SelBookmarks.Add(grdFind.Bookmark);
                    }

                    OpenHR.getFrame("workframe").refreshControls();

                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();
                }

                if (sCurrentWorkPage == "UTIL_DEF_EXPRCOMPONENT") {
                    var sAction = frmOptionData.txtOptionAction.value;
                    var sControlName;

                    if ((sAction == "LOADEXPRFIELDCOLUMNS") ||
                        (sAction == "LOADEXPRLOOKUPCOLUMNS")) {
                        var dataCollection = frmOptionData.elements;

                        if (dataCollection != null) {
                            for (i = 0; i < dataCollection.length; i++) {
                                sControlName = dataCollection.item(i).name;
                                sControlName = sControlName.substr(0, 10);
                                if (sControlName == "txtColumn_") {
                                    component_addColumn(dataCollection.item(i).value);
                                }
                            }
                        }

                        component_setColumn(frmOptionData.txtOptionColumnID.value);
                    }

                    if (sAction == "LOADEXPRLOOKUPVALUES") {
                        var dataCollection = frmOptionData.elements;

                        if (dataCollection != null) {
                            for (i = 0; i < dataCollection.length; i++) {
                                sControlName = dataCollection.item(i).name;
                                sControlName = sControlName.substr(0, 9);
                                if (sControlName == "txtValue_") {
                                    component_addValue(dataCollection.item(i).value);
                                }
                            }
                        }

                        component_setValue(frmOptionData.txtOptionLocateValue.value);
                    }

                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();
                }

                if (sCurrentWorkPage == "FIND") {
                    var sAction = frmOptionData.txtOptionAction.value;

                    if ((sAction == "BOOKCOURSEERROR") ||
                        (sAction == "TRANSFERBOOKINGERROR") ||
                        (sAction == "ADDFROMWAITINGLISTERROR") ||
                        (sAction == "BULKBOOKINGERROR")) {
                        OpenHR.messageBox(frmOptionData.txtNonFatalErrorDescription.value);
                    }
                    if ((sAction == "BOOKCOURSESUCCESS") ||
                        (sAction == "TRANSFERBOOKINGSUCCESS") ||
                        (sAction == "ADDFROMWAITINGLISTSUCCESS") ||
                        (sAction == "BULKBOOKINGSUCCESS")) {
                        // Reload the find records.
                        OpenHR.messageBox("Booking(s) made successfully.");
                        menu_reloadFindPage("MOVEFIRST", "");
                    }
                }
            }
        }
    }
    -->
</script>

<script type="text/javascript">
<!--
    function refreshOptionData() {
        OpenHR.submitForm(frmGetOptionData);
    }

    -->
</script>

<form action="optionData_Submit" method="post" id="frmGetOptionData" name="frmGetOptionData">
    <input type="hidden" id="txtOptionAction" name="txtOptionAction">
    <input type="hidden" id="txtOptionTableID" name="txtOptionTableID">
    <input type="hidden" id="txtOptionViewID" name="txtOptionViewID">
    <input type="hidden" id="txtOptionOrderID" name="txtOptionOrderID">
    <input type="hidden" id="txtOptionColumnID" name="txtOptionColumnID">
    <input type="hidden" id="txtOptionPageAction" name="txtOptionPageAction">
    <input type="hidden" id="txtOptionFirstRecPos" name="txtOptionFirstRecPos">
    <input type="hidden" id="txtOptionCurrentRecCount" name="txtOptionCurrentRecCount">
    <input type="hidden" id="txtGotoLocateValue" name="txtGotoLocateValue">
    <input type="hidden" id="txtOptionCourseTitle" name="txtOptionCourseTitle">
    <input type="hidden" id="txtOptionRecordID" name="txtOptionRecordID">
    <input type="hidden" id="txtOptionLinkRecordID" name="txtOptionLinkRecordID">
    <input type="hidden" id="txtOptionValue" name="txtOptionValue">
    <input type="hidden" id="txtOptionSQL" name="txtOptionSQL">
    <input type="hidden" id="txtOptionPromptSQL" name="txtOptionPromptSQL">
    <input type="hidden" id="txtOptionOnlyNumerics" name="txtOptionOnlyNumerics">
    <input type="hidden" id="txtOptionLookupColumnID" name="txtOptionLookupColumnID">
    <input type="hidden" id="txtOptionLookupFilterValue" name="txtOptionLookupFilterValue">
    <input type="hidden" id="txtOptionIsLookupTable" name="txtOptionIsLookupTable">
    <input type="hidden" id="txtOptionParentTableID" name="txtOptionParentTableID">
    <input type="hidden" id="txtOptionParentRecordID" name="txtOptionParentRecordID">
    <input type="hidden" id="txtOption1000SepCols" name="txtOption1000SepCols">
</form>

<form id="frmOptionData" name="frmOptionData">
    <%
        On Error Resume Next
		
        Dim aPrompts(1, 0)

        Const adStateOpen = 1
    
        Const DEADLOCK_ERRORNUMBER = -2147467259
        Const DEADLOCK_MESSAGESTART = "YOUR TRANSACTION (PROCESS ID #"
        Const DEADLOCK_MESSAGEEND = ") WAS DEADLOCKED WITH ANOTHER PROCESS AND HAS BEEN CHOSEN AS THE DEADLOCK VICTIM. RERUN YOUR TRANSACTION."
        Const DEADLOCK2_MESSAGESTART = "TRANSACTION (PROCESS ID "
        Const DEADLOCK2_MESSAGEEND = ") WAS DEADLOCKED ON "
        Const SQLMAILNOTSTARTEDMESSAGE = "SQL MAIL SESSION IS NOT STARTED."

        Const iRETRIES = 5
        Dim iRetryCount As Integer = 0
        ' NPG20080904 Fault 13018
        Session("flagOverrideFilter") = False

        Dim objUtilities

        Dim sErrorDescription As String = ""
        Dim sNonFatalErrorDescription As String = ""

        Dim cmdThousandFindColumns
        Dim prmError
        Dim prmTableID
        Dim prmViewID
        Dim prmOrderID
        Dim prmThousandColumns
        Dim cmdGetFindRecords
        Dim sThousandColumns As String
        Dim prmReqRecs

        Dim prmIsFirstPage
        Dim prmIsLastPage
        Dim prmLocateValue
        Dim prmColumnType
        Dim prmAction
        Dim prmTotalRecCount
        Dim prmFirstRecPos
        Dim prmCurrentRecCount
        Dim prmExcludedIDs
        Dim prmColumnSize
        Dim prmColumnDecimals
        Dim rstFindRecords
    
        Dim cmdGetFilterValue
        Dim prmScreenID
        Dim prmColumnID
        Dim prmRecordID
        Dim prmFilterValue
        Dim prmParentTableID
        Dim prmParentRecordID

        Dim prmLookupColumnID
        Dim prmLookupColumnGridPosition
        Dim prmOverrideFilter
    
        Dim prmCourseTitle
        Dim prmCourseRecordID
        Dim prmWLRecordID
        Dim prmEmpRecordID

        Dim cmdTransferCourse
        Dim cmdBookCourse
        Dim prmStatus
        Dim fDeadlock As Boolean
        Dim sErrMsg As String
    
        Dim prmTBRecordID
        Dim prmErrorMessage

        Dim iCount As Integer
        Dim sAddString As String
        Dim sColDef As String
        Dim sTemp As String
    
        Dim j As Integer
        Dim sPrompts As String
        Dim iIndex1 As Integer
        Dim iIndex2 As Integer
    
        Dim cmdBulkBooking
        Dim prmSelectionType
        Dim prmSelectionID
        Dim prmSelectedIDs
        Dim prmPromptSQL
    
        Dim fOK As Boolean

        Dim prmErrMsg
        Dim cmdPicklist
        Dim prmExpectedCount
        Dim cmdBulkBook
        Dim prmEmployeeRecordIDs
        Dim prmOnlyNumerics
        Dim cmdExprColumns
        Dim prmComponentType
        Dim rstExprColumns
        Dim cmdExprValues
        Dim rstExprValues
        Dim prmDataType
    
        Response.Write("<INPUT type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

        ' Get the required record count if we have a query.
        '	if len(session("selectSQL")) > 0 then
        If Session("optionAction") = "LOADFIND" Then
            sThousandColumns = ""
			
            cmdThousandFindColumns = CreateObject("ADODB.Command")
            cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
            cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
            cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
            cmdThousandFindColumns.CommandTimeout = 180
		
            prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdThousandFindColumns.Parameters.Append(prmError)

            prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))

            prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
            cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
            Err.Clear()
            cmdThousandFindColumns.Execute()

            If (Err.Number <> 0) Then
                sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
            End If
	
            ' Release the ADO command object.
            cmdThousandFindColumns = Nothing

            cmdGetFindRecords = CreateObject("ADODB.Command")
            cmdGetFindRecords.CommandText = "sp_ASRIntGetLinkFindRecords"
            cmdGetFindRecords.CommandType = 4 ' Stored procedure
            cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
            cmdGetFindRecords.CommandTimeout = 180
			
            prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))

            prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmError)

        
            prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmReqRecs)
            prmReqRecs.value = CleanNumeric(Session("FindRecords"))

            prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

            prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsLastPage)

            prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 2147483646)
            cmdGetFindRecords.Parameters.Append(prmLocateValue)
            prmLocateValue.value = Session("optionLocateValue")

            prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnType)

            prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 100)
            cmdGetFindRecords.Parameters.Append(prmAction)
            prmAction.value = Session("optionPageAction")

            prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

            prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
            cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
            prmFirstRecPos.value = CleanNumeric(Session("optionFirstRecPos"))

            prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
            cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
            prmCurrentRecCount.value = CleanNumeric(Session("optionCurrentRecCount"))

            prmExcludedIDs = cmdGetFindRecords.CreateParameter("excludedIDs", 200, 1, 2147483646) ' 200=varchar, 1=input, 8000=size
            cmdGetFindRecords.Parameters.Append(prmExcludedIDs)
            prmExcludedIDs.value = ""
		
            prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnSize)

            prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

            Err.Clear()
            rstFindRecords = cmdGetFindRecords.Execute
	
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading the link find records." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                If rstFindRecords.state = adStateOpen Then
                    iCount = 0
                    Do While Not rstFindRecords.EOF
                        sAddString = ""
						
                        For iloop = 0 To (rstFindRecords.fields.count - 1)
                            If iloop > 0 Then
                                sAddString = sAddString & "	"
                            End If
							
                            If iCount = 0 Then
                                sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                                Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
                            End If
							
                            If rstFindRecords.fields(iloop).type = 135 Then
                                ' Field is a date so format as such.
                                sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
                            ElseIf rstFindRecords.fields(iloop).type = 131 Then
                                ' Field is a numeric so format as such.
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, True)
                                    Else
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, False)
                                    End If
                                    sTemp = Replace(sTemp, ".", "x")
                                    sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
                                    sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
                                    sAddString = sAddString & sTemp
                                End If
                            Else
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
                                End If
                            End If
                        Next

                        Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
                        iCount = iCount + 1
                        rstFindRecords.moveNext()
                    Loop
	
                    ' Release the ADO recordset object.
                    rstFindRecords.close()
                End If
            End If
            rstFindRecords = Nothing

            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If cmdGetFindRecords.Parameters("error").Value <> 0 Then
                'Session("ErrorTitle") = "Link Find Page"
                'Session("ErrorText") = "Error reading link records definition."
                'Response.Clear	  
                'Response.Redirect("error.asp")
            End If

            Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

            cmdGetFindRecords = Nothing
			
        ElseIf Session("optionAction") = "LOADLOOKUPFIND" Then
            ' Check if the filter value column is in the current screen.
            ' If not, try and get the filter value from the database.
            If Len(Session("optionFilterValue")) = 0 Then
                cmdGetFilterValue = CreateObject("ADODB.Command")
                cmdGetFilterValue.CommandText = "spASRIntGetLookupFilterValue"
                cmdGetFilterValue.CommandType = 4 ' Stored procedure
                cmdGetFilterValue.ActiveConnection = Session("databaseConnection")

                prmScreenID = cmdGetFilterValue.CreateParameter("screenID", 3, 1)
                cmdGetFilterValue.Parameters.Append(prmScreenID)
                prmScreenID.value = CleanNumeric(Session("screenID"))

                prmColumnID = cmdGetFilterValue.CreateParameter("LookupColumnID", 3, 1)
                cmdGetFilterValue.Parameters.Append(prmColumnID)
                prmColumnID.value = CleanNumeric(Session("optionColumnID"))

                prmTableID = cmdGetFilterValue.CreateParameter("tableID", 3, 1)
                cmdGetFilterValue.Parameters.Append(prmTableID)
                prmTableID.value = CleanNumeric(Session("tableID"))

                prmViewID = cmdGetFilterValue.CreateParameter("viewID", 3, 1)
                cmdGetFilterValue.Parameters.Append(prmViewID)
                prmViewID.value = CleanNumeric(Session("viewID"))
				
                prmRecordID = cmdGetFilterValue.CreateParameter("recordID", 3, 1)
                cmdGetFilterValue.Parameters.Append(prmRecordID)
                prmRecordID.value = CleanNumeric(Session("optionRecordID"))
				
                prmFilterValue = cmdGetFilterValue.CreateParameter("FilterValue", 200, 2, 8000) ' 200=adVarChar, 2=output, 8000=size
                cmdGetFilterValue.Parameters.Append(prmFilterValue)

                prmParentTableID = cmdGetFilterValue.CreateParameter("ParentTableID", 3, 1)
                cmdGetFilterValue.Parameters.Append(prmParentTableID)
                prmParentTableID.value = CleanNumeric(Session("optionParentTableID"))

                prmParentRecordID = cmdGetFilterValue.CreateParameter("ParentRecordID", 3, 1)
                cmdGetFilterValue.Parameters.Append(prmParentRecordID)
                prmParentRecordID.value = CleanNumeric(Session("optionParentRecordID"))

                ' NPG20080904 Fault 13018
                prmError = cmdGetFilterValue.CreateParameter("Error", 11, 2) ' 11=bit, 2=output
                cmdGetFilterValue.Parameters.Append(prmError)


                Err.Clear()
                cmdGetFilterValue.Execute()

                If (Err.Number <> 0) Then
                    sErrorDescription = "Error reading the lookup filter value." & vbCrLf & formatError(Err.Description)
                End If
				
                If Len(sErrorDescription) = 0 Then
                    Session("optionFilterValue") = cmdGetFilterValue.Parameters("FilterValue").Value
                    Session("flagOverrideFilter") = cmdGetFilterValue.Parameters("Error").Value
                    cmdGetFilterValue = Nothing
                End If
            End If
		
            If Len(sErrorDescription) = 0 Then
                sThousandColumns = ""

                If Session("IsLookupTable") = "False" Then
                    cmdThousandFindColumns = CreateObject("ADODB.Command")
                    cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
                    cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
                    cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
                    cmdThousandFindColumns.CommandTimeout = 180
		
                    prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
                    cmdThousandFindColumns.Parameters.Append(prmError)

                    prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
                    cmdThousandFindColumns.Parameters.Append(prmTableID)
                    prmTableID.value = CleanNumeric(Session("optionTableID"))

                    prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
                    cmdThousandFindColumns.Parameters.Append(prmViewID)
                    prmViewID.value = CleanNumeric(Session("optionViewID"))

                    prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
                    cmdThousandFindColumns.Parameters.Append(prmOrderID)
                    prmOrderID.value = CleanNumeric(Session("optionOrderID"))

                    prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
                    cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
                    Err.Clear()
                    cmdThousandFindColumns.Execute()

                    If (Err.Number <> 0) Then
                        sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
                    End If

                    If Len(sErrorDescription) = 0 Then
                        sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
                    End If
	
                    ' Release the ADO command object.
                    cmdThousandFindColumns = Nothing

                    cmdGetFindRecords = CreateObject("ADODB.Command")
                    cmdGetFindRecords.CommandText = "spASRIntGetLookupFindRecords2"
                    cmdGetFindRecords.CommandType = 4 ' Stored procedure
                    cmdGetFindRecords.ActiveConnection = Session("databaseConnection")

                    prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
                    cmdGetFindRecords.Parameters.Append(prmTableID)
                    prmTableID.value = CleanNumeric(Session("optionTableID"))

                    prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
                    cmdGetFindRecords.Parameters.Append(prmViewID)
                    prmViewID.value = CleanNumeric(Session("optionViewID"))

                    prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
                    cmdGetFindRecords.Parameters.Append(prmOrderID)
                    prmOrderID.value = CleanNumeric(Session("optionOrderID"))

                    prmColumnID = cmdGetFindRecords.CreateParameter("LookupColumnID", 3, 1)
                    cmdGetFindRecords.Parameters.Append(prmColumnID)
                    prmColumnID.value = CleanNumeric(Session("optionLookupColumnID"))

                    prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
                    cmdGetFindRecords.Parameters.Append(prmReqRecs)
                    prmReqRecs.value = CleanNumeric(Session("FindRecords"))

                    prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
                    cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

                    prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
                    cmdGetFindRecords.Parameters.Append(prmIsLastPage)

                    prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
                    cmdGetFindRecords.Parameters.Append(prmLocateValue)
                    prmLocateValue.value = Session("optionLocateValue")

                    prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
                    cmdGetFindRecords.Parameters.Append(prmColumnType)

                    prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
                    cmdGetFindRecords.Parameters.Append(prmColumnSize)

                    prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
                    cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

                    prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
                    cmdGetFindRecords.Parameters.Append(prmAction)
                    prmAction.value = Session("optionPageAction")

                    prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
                    cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

                    prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
                    cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
                    prmFirstRecPos.value = CleanNumeric(Session("optionFirstRecPos"))

                    prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
                    cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
                    prmCurrentRecCount.value = CleanNumeric(Session("optionCurrentRecCount"))

                    prmFilterValue = cmdGetFindRecords.CreateParameter("FilterValue", 200, 1, 8000) ' 200=adVarChar, 1=input
                    cmdGetFindRecords.Parameters.Append(prmFilterValue)
                    prmFilterValue.value = Session("optionFilterValue")
                
                    prmLookupColumnID = cmdGetFindRecords.CreateParameter("CallingColumnID", 3, 1) ' 200=adVarChar, 1=input
                    cmdGetFindRecords.Parameters.Append(prmLookupColumnID)
                    prmLookupColumnID.value = CleanNumeric(Session("optionColumnID"))

                    prmLookupColumnGridPosition = cmdGetFindRecords.CreateParameter("LookupColumnGridPosition", 3, 2) ' 200=adVarChar, 2=output
                    cmdGetFindRecords.Parameters.Append(prmLookupColumnGridPosition)
					
                    ' NPG20080904 Fault 13018
                    prmOverrideFilter = cmdGetFindRecords.CreateParameter("pfOverrideFilter", 11, 1) ' 11=bit, 1=input
                    cmdGetFindRecords.Parameters.Append(prmOverrideFilter)
                    prmOverrideFilter.value = Session("flagOverrideFilter")
                Else
                    cmdThousandFindColumns = CreateObject("ADODB.Command")
                    cmdThousandFindColumns.CommandText = "spASRIntGetLookupFindColumnInfo"
                    cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
                    cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
                    cmdThousandFindColumns.CommandTimeout = 180
		
                    prmLookupColumnID = cmdThousandFindColumns.CreateParameter("lookupColumnID", 3, 1)
                    cmdThousandFindColumns.Parameters.Append(prmLookupColumnID)
                    prmLookupColumnID.value = CleanNumeric(Session("optionLookupColumnID"))

                    prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
                    cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
                    Err.Clear()
                    cmdThousandFindColumns.Execute()

                    If (Err.Number <> 0) Then
                        sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
                    End If

                    If Len(sErrorDescription) = 0 Then
                        sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
                    End If
	
                    ' Release the ADO command object.
                    cmdThousandFindColumns = Nothing

                    cmdGetFindRecords = CreateObject("ADODB.Command")
                    cmdGetFindRecords.CommandText = "spASRIntGetLookupFindRecords"
                    cmdGetFindRecords.CommandType = 4 ' Stored procedure
                    cmdGetFindRecords.ActiveConnection = Session("databaseConnection")

                    prmColumnID = cmdGetFindRecords.CreateParameter("LookupColumnID", 3, 1)
                    cmdGetFindRecords.Parameters.Append(prmColumnID)
                    prmColumnID.value = CleanNumeric(Session("optionLookupColumnID"))

                    prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
                    cmdGetFindRecords.Parameters.Append(prmReqRecs)
                    prmReqRecs.value = CleanNumeric(Session("FindRecords"))

                    prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
                    cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

                    prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
                    cmdGetFindRecords.Parameters.Append(prmIsLastPage)

                    prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
                    cmdGetFindRecords.Parameters.Append(prmLocateValue)
                    prmLocateValue.value = Session("optionLocateValue")

                    prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
                    cmdGetFindRecords.Parameters.Append(prmColumnType)

                    prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
                    cmdGetFindRecords.Parameters.Append(prmColumnSize)

                    prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
                    cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

                    prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
                    cmdGetFindRecords.Parameters.Append(prmAction)
                    prmAction.value = Session("optionPageAction")

                    prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
                    cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

                    prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
                    cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
                    prmFirstRecPos.value = CleanNumeric(Session("optionFirstRecPos"))

                    prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
                    cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
                    prmCurrentRecCount.value = CleanNumeric(Session("optionCurrentRecCount"))

                    prmFilterValue = cmdGetFindRecords.CreateParameter("FilterValue", 200, 1, 8000) ' 200=adVarChar, 1=input
                    cmdGetFindRecords.Parameters.Append(prmFilterValue)
                    prmFilterValue.value = Session("optionFilterValue")

                    prmLookupColumnID = cmdGetFindRecords.CreateParameter("@CallingColumnID", 3, 1) ' 200=adVarChar, 1=input
                    cmdGetFindRecords.Parameters.Append(prmLookupColumnID)
                    prmLookupColumnID.value = CleanNumeric(Session("optionColumnID"))
					
                    ' NPG20080904 Fault 13018					
                    prmOverrideFilter = cmdGetFindRecords.CreateParameter("pfOverrideFilter", 11, 1) ' 11=bit, 1=input
                    cmdGetFindRecords.Parameters.Append(prmOverrideFilter)
                    prmOverrideFilter.value = Session("flagOverrideFilter")
                End If
					
                Err.Clear()
                rstFindRecords = cmdGetFindRecords.Execute

                If (Err.Number <> 0) Then
                    sErrorDescription = "Error reading the lookup find records." & vbCrLf & formatError(Err.Description)
                End If

                If Len(sErrorDescription) = 0 Then
                    If rstFindRecords.state = adStateOpen Then
                        iCount = 0
                        Do While Not rstFindRecords.EOF
                            sAddString = ""
							
                            For iloop = 0 To (rstFindRecords.fields.count - 1)
                                If iloop > 0 Then
                                    sAddString = sAddString & "	"
                                End If
								
                                If iCount = 0 Then
                                    sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                                    Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
                                End If
								
                                If rstFindRecords.fields(iloop).type = 135 Then
                                    ' Field is a date so format as such.
                                    sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
                                ElseIf rstFindRecords.fields(iloop).type = 131 Then
                                    ' Field is a numeric so format as such.
                                    If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                        If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
                                            sTemp = ""
                                            sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, True)
                                        Else
                                            sTemp = ""
                                            sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, False)
                                        End If
                                        sTemp = Replace(sTemp, ".", "x")
                                        sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
                                        sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
                                        sAddString = sAddString & sTemp
                                    End If
                                Else
                                    If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                        sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
                                    End If
                                End If
                            Next

                            Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
						
                            iCount = iCount + 1
                            rstFindRecords.moveNext()
                        Loop
	
                        ' Release the ADO recordset object.
                        rstFindRecords.close()
                    End If
                End If
                rstFindRecords = Nothing

                ' NB. IMPORTANT ADO NOTE.
                ' When calling a stored procedure which returns a recordset AND has output parameters
                ' you need to close the recordset and set it to nothing before using the output parameters. 
                Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtFilterOverride name=txtFilterOverride value=" & Session("flagOverrideFilter") & ">" & vbCrLf)

                If Session("IsLookupTable") = "False" Then
                    Response.Write("<INPUT type='hidden' id=txtLookupColumnGridPosition name=txtLookupColumnGridPosition value=" & cmdGetFindRecords.Parameters("LookupColumnGridPosition").Value & ">" & vbCrLf)
                Else
                    Response.Write("<INPUT type='hidden' id=txtLookupColumnGridPosition name=txtLookupColumnGridPosition value=0>" & vbCrLf)
                End If
							
                cmdGetFindRecords = Nothing
            End If
        ElseIf Session("optionAction") = "LOADTRANSFERCOURSE" Then
            sThousandColumns = ""
			
            cmdThousandFindColumns = CreateObject("ADODB.Command")
            cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
            cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
            cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
            cmdThousandFindColumns.CommandTimeout = 180
		
            prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdThousandFindColumns.Parameters.Append(prmError)

            prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))

            prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
            cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
            Err.Clear()
            cmdThousandFindColumns.Execute()

            If (Err.Number <> 0) Then
                sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
            End If
	
            ' Release the ADO command object.
            cmdThousandFindColumns = Nothing

            cmdGetFindRecords = CreateObject("ADODB.Command")
            cmdGetFindRecords.CommandText = "sp_ASRIntGetTransferCourseRecords"
            cmdGetFindRecords.CommandType = 4 ' Stored procedure
            cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
            cmdGetFindRecords.CommandTimeout = 180
			
            prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))
        
            prmCourseTitle = cmdGetFindRecords.CreateParameter("courseTitle", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmCourseTitle)
            prmCourseTitle.value = Session("optionCourseTitle")

            prmCourseRecordID = cmdGetFindRecords.CreateParameter("courseRecordID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmCourseRecordID)
            prmCourseRecordID.value = CleanNumeric(Session("optionRecordID"))

            prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmError)

            prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmReqRecs)
            prmReqRecs.value = CleanNumeric(Session("FindRecords"))

            prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

            prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsLastPage)

            prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmLocateValue)
            prmLocateValue.value = Session("optionLocateValue")

            prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnType)

            prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmAction)
            prmAction.value = Session("optionPageAction")

            prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

            prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
            cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
            prmFirstRecPos.value = CleanNumeric(Session("optionFirstRecPos"))

            prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
            cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
            prmCurrentRecCount.value = CleanNumeric(Session("optionCurrentRecCount"))

            prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnSize)

            prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

            Err.Clear()
            rstFindRecords = cmdGetFindRecords.Execute
	
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                If rstFindRecords.state = adStateOpen Then
                    iCount = 0
                    Do While Not rstFindRecords.EOF
                        sAddString = ""
						
                        For iloop = 0 To (rstFindRecords.fields.count - 1)
                            If iloop > 0 Then
                                sAddString = sAddString & "	"
                            End If
							
                            If iCount = 0 Then
                                sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                                Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
                            End If
							
                            If rstFindRecords.fields(iloop).type = 135 Then
                                ' Field is a date so format as such.
                                sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
                            ElseIf rstFindRecords.fields(iloop).type = 131 Then
                                ' Field is a numeric so format as such.
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, True)
                                    Else
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, False)
                                    End If
                                    sTemp = Replace(sTemp, ".", "x")
                                    sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
                                    sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
                                    sAddString = sAddString & sTemp
                                End If
                            Else
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
                                End If
                            End If
                        Next

                        Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
                        iCount = iCount + 1
                        rstFindRecords.moveNext()
                    Loop
	
                    ' Release the ADO recordset object.
                    rstFindRecords.close()
                End If
            End If
            rstFindRecords = Nothing

            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If cmdGetFindRecords.Parameters("error").Value <> 0 Then
                'Session("ErrorTitle") = "Transfer Course Find Page"
                'Session("ErrorText") = "Error reading records definition."
                'Response.Clear	  
                'Response.Redirect("error.asp")
            End If

            Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

            cmdGetFindRecords = Nothing

        ElseIf Session("optionAction") = "LOADBOOKCOURSE" Then
            sThousandColumns = ""
			
            cmdThousandFindColumns = CreateObject("ADODB.Command")
            cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
            cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
            cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
            cmdThousandFindColumns.CommandTimeout = 180
		
            prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdThousandFindColumns.Parameters.Append(prmError)

            prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))

            prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
            cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
            Err.Clear()
            cmdThousandFindColumns.Execute()

            If (Err.Number <> 0) Then
                sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
            End If
	
            ' Release the ADO command object.
            cmdThousandFindColumns = Nothing

            cmdGetFindRecords = CreateObject("ADODB.Command")
            cmdGetFindRecords.CommandText = "sp_ASRIntGetBookCourseRecords"
            cmdGetFindRecords.CommandType = 4 ' Stored procedure
            cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
            cmdGetFindRecords.CommandTimeout = 180
			
            prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))
        
            prmWLRecordID = cmdGetFindRecords.CreateParameter("WLRecordID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmWLRecordID)
            prmWLRecordID.value = CleanNumeric(Session("optionRecordID"))

            prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmError)

            prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmReqRecs)
            prmReqRecs.value = CleanNumeric(Session("FindRecords"))

            prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

            prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsLastPage)

            prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmLocateValue)
            prmLocateValue.value = Session("optionLocateValue")

            prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnType)

            prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmAction)
            prmAction.value = Session("optionPageAction")

            prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

            prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
            cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
            prmFirstRecPos.value = CleanNumeric(Session("optionFirstRecPos"))

            prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
            cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
            prmCurrentRecCount.value = CleanNumeric(Session("optionCurrentRecCount"))

            prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnSize)

            prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

            Err.Clear()
            rstFindRecords = cmdGetFindRecords.Execute
	
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                If rstFindRecords.state = adStateOpen Then
                    iCount = 0
                    Do While Not rstFindRecords.EOF
                        sAddString = ""
						
                        For iloop = 0 To (rstFindRecords.fields.count - 1)
                            If iloop > 0 Then
                                sAddString = sAddString & "	"
                            End If
							
                            If iCount = 0 Then
                                sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                                Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
                            End If
							
                            If rstFindRecords.fields(iloop).type = 135 Then
                                ' Field is a date so format as such.
                                sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
                            ElseIf rstFindRecords.fields(iloop).type = 131 Then
                                ' Field is a numeric so format as such.
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, True)
                                    Else
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, False)
                                    End If
                                    sTemp = Replace(sTemp, ".", "x")
                                    sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
                                    sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
                                    sAddString = sAddString & sTemp
                                End If
                            Else
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
                                End If
                            End If
                        Next

                        Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
                        iCount = iCount + 1
                        rstFindRecords.moveNext()
                    Loop
	
                    ' Release the ADO recordset object.
                    rstFindRecords.close()
                End If
            End If
            rstFindRecords = Nothing

            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If cmdGetFindRecords.Parameters("error").Value <> 0 Then
                'Session("ErrorTitle") = "Book Course Find Page"
                'Session("ErrorText") = "Error reading records definition."
                'Response.Clear	  
                'Response.Redirect("error.asp")
            End If

            Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

            cmdGetFindRecords = Nothing

        ElseIf Session("optionAction") = "SELECTBOOKCOURSE_3" Then
		        
            cmdBookCourse = CreateObject("ADODB.Command")
            cmdBookCourse.CommandText = "sp_ASRIntBookCourse"
            cmdBookCourse.CommandType = 4 ' Stored procedure
            cmdBookCourse.CommandTimeout = 180
            cmdBookCourse.ActiveConnection = Session("databaseConnection")
					
            prmWLRecordID = cmdBookCourse.CreateParameter("WLRecordID", 3, 1) ' 3=integer, 1=input
            cmdBookCourse.Parameters.Append(prmWLRecordID)
            prmWLRecordID.value = CleanNumeric(Session("optionRecordID"))

            prmCourseRecordID = cmdBookCourse.CreateParameter("CourseRecordID", 3, 1) ' 3=integer, 1=input
            cmdBookCourse.Parameters.Append(prmCourseRecordID)
            prmCourseRecordID.value = CleanNumeric(Session("optionLinkRecordID"))

            prmStatus = cmdBookCourse.CreateParameter("status", 200, 1, 2147483646)
            cmdBookCourse.Parameters.Append(prmStatus)
            prmStatus.value = Session("optionValue")

            fDeadlock = True
            Do While fDeadlock
                fDeadlock = False
									
                cmdBookCourse.ActiveConnection.Errors.Clear()
									
                ' Run the insert stored procedure.
                cmdBookCourse.Execute()

                If cmdBookCourse.ActiveConnection.Errors.Count > 0 Then
                    For iLoop = 1 To cmdBookCourse.ActiveConnection.Errors.Count
                        sErrMsg = formatError(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)

                        If (cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
                            (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                                (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                            ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
                            ' The error is for a deadlock.
                            ' Sorry about having to use the err.description to trap the error but the err.number
                            ' is not specific and MSDN suggests using the err.description.
                            If (iRetryCount < iRETRIES) And (cmdBookCourse.ActiveConnection.Errors.Count = 1) Then
                                iRetryCount = iRetryCount + 1
                                fDeadlock = True
                            Else
                                If Len(sNonFatalErrorDescription) > 0 Then
                                    sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf
                                End If
                                sNonFatalErrorDescription = sNonFatalErrorDescription & "Another user is deadlocking the database."
                                fOK = False
                            End If
                        ElseIf UCase(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
                            '"SQL Mail session is not started."
                            'Ignore this error
                            'ElseIf (cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
                            '	(UCase(Left(cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
                            '"EXECUTE permission denied on object 'xp_sendmail'"
                            'Ignore this error
					
                        Else
                            sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf & _
                                formatError(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)
                            fOK = False
                        End If
                    Next

                    cmdBookCourse.ActiveConnection.Errors.Clear()
												
                    If Not fOK Then
                        sNonFatalErrorDescription = "The booking could not be made." & vbCrLf & sNonFatalErrorDescription
                        Session("optionAction") = "BOOKCOURSEERROR"
                    End If
                Else
                    Session("optionAction") = "BOOKCOURSESUCCESS"
                End If
            Loop
            cmdBookCourse = Nothing

        ElseIf Session("optionAction") = "SELECTADDFROMWAITINGLIST_3" Then
            cmdBookCourse = CreateObject("ADODB.Command")
            cmdBookCourse.CommandText = "sp_ASRIntAddFromWaitingList"
            cmdBookCourse.CommandType = 4 ' Stored procedure
            cmdBookCourse.CommandTimeout = 180
            cmdBookCourse.ActiveConnection = Session("databaseConnection")
					
            prmEmpRecordID = cmdBookCourse.CreateParameter("EmpRecordID", 3, 1) ' 3=integer, 1=input
            cmdBookCourse.Parameters.Append(prmEmpRecordID)
            prmEmpRecordID.value = CleanNumeric(Session("optionLinkRecordID"))

            prmCourseRecordID = cmdBookCourse.CreateParameter("CourseRecordID", 3, 1) ' 3=integer, 1=input
            cmdBookCourse.Parameters.Append(prmCourseRecordID)
            prmCourseRecordID.value = CleanNumeric(Session("optionRecordID"))

            prmStatus = cmdBookCourse.CreateParameter("status", 200, 1, 8000)
            cmdBookCourse.Parameters.Append(prmStatus)
            prmStatus.value = Session("optionValue")

            fDeadlock = True
            Do While fDeadlock
                fDeadlock = False
									
                cmdBookCourse.ActiveConnection.Errors.Clear()
									
                ' Run the insert stored procedure.
                cmdBookCourse.Execute()

                If cmdBookCourse.ActiveConnection.Errors.Count > 0 Then
                    For iLoop = 1 To cmdBookCourse.ActiveConnection.Errors.Count
                        sErrMsg = formatError(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)

                        If (cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
                            (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                                (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                            ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
                            ' The error is for a deadlock.
                            ' Sorry about having to use the err.description to trap the error but the err.number
                            ' is not specific and MSDN suggests using the err.description.
                            If (iRetryCount < iRETRIES) And (cmdBookCourse.ActiveConnection.Errors.Count = 1) Then
                                iRetryCount = iRetryCount + 1
                                fDeadlock = True
                            Else
                                If Len(sNonFatalErrorDescription) > 0 Then
                                    sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf
                                End If
                                sNonFatalErrorDescription = sNonFatalErrorDescription & "Another user is deadlocking the database."
                                fOK = False
                            End If
                        ElseIf UCase(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
                            '"SQL Mail session is not started."
                            'Ignore this error
                            'ElseIf (cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
                            '	(UCase(Left(cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
                            '"EXECUTE permission denied on object 'xp_sendmail'"
                            'Ignore this error
					
                        Else
                            sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf & _
                                formatError(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)
                            fOK = False
                        End If
                    Next

                    cmdBookCourse.ActiveConnection.Errors.Clear()
												
                    If Not fOK Then
                        sNonFatalErrorDescription = "The booking could not be made." & vbCrLf & sNonFatalErrorDescription
                        Session("optionAction") = "ADDFROMWAITINGLISTERROR"
                    End If
                Else
                    Session("optionAction") = "ADDFROMWAITINGLISTSUCCESS"
                End If
            Loop
            cmdBookCourse = Nothing

        ElseIf Session("optionAction") = "LOADTRANSFERBOOKING" Then
            sThousandColumns = ""
			
            cmdThousandFindColumns = CreateObject("ADODB.Command")
            cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
            cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
            cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
            cmdThousandFindColumns.CommandTimeout = 180
		
            prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdThousandFindColumns.Parameters.Append(prmError)

            prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))

            prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
            cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
            Err.Clear()
            cmdThousandFindColumns.Execute()

            If (Err.Number <> 0) Then
                sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
            End If
	
            ' Release the ADO command object.
            cmdThousandFindColumns = Nothing

            cmdGetFindRecords = CreateObject("ADODB.Command")
            cmdGetFindRecords.CommandText = "sp_ASRIntGetTransferBookingRecords"
            cmdGetFindRecords.CommandType = 4 ' Stored procedure
            cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
            cmdGetFindRecords.CommandTimeout = 180
			
            prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))
        
            prmTBRecordID = cmdGetFindRecords.CreateParameter("TBRecordID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmTBRecordID)
            prmTBRecordID.value = CleanNumeric(Session("optionRecordID"))

            prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmError)

            prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmReqRecs)
            prmReqRecs.value = CleanNumeric(Session("FindRecords"))

            prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

            prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsLastPage)

            prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmLocateValue)
            prmLocateValue.value = Session("optionLocateValue")

            prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnType)

            prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmAction)
            prmAction.value = Session("optionPageAction")

            prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

            prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
            cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
            prmFirstRecPos.value = CleanNumeric(Session("optionFirstRecPos"))

            prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
            cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
            prmCurrentRecCount.value = CleanNumeric(Session("optionCurrentRecCount"))

            prmErrorMessage = cmdGetFindRecords.CreateParameter("errorMessage", 200, 2, 8000) ' 200=varchar, 2=output,8000=size
            cmdGetFindRecords.Parameters.Append(prmErrorMessage)

            prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnSize)

            prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

            prmStatus = cmdGetFindRecords.CreateParameter("status", 200, 2, 8000) ' 200=varchar, 2=output,8000=size
            cmdGetFindRecords.Parameters.Append(prmStatus)

            Err.Clear()
            rstFindRecords = cmdGetFindRecords.Execute
	
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
            End If

        
            If Len(sErrorDescription) = 0 Then
                If rstFindRecords.state = adStateOpen Then
                    iCount = 0
                    Do While Not rstFindRecords.EOF
                        sAddString = ""
						
                        For iloop = 0 To (rstFindRecords.fields.count - 1)
                            If iloop > 0 Then
                                sAddString = sAddString & "	"
                            End If
							
                            If iCount = 0 Then
                                sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                                Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
                            End If
							
                            If rstFindRecords.fields(iloop).type = 135 Then
                                ' Field is a date so format as such.
                                sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
                            ElseIf rstFindRecords.fields(iloop).type = 131 Then
                                ' Field is a numeric so format as such.
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, True)
                                    Else
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, False)
                                    End If
                                    sTemp = Replace(sTemp, ".", "x")
                                    sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
                                    sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
                                    sAddString = sAddString & sTemp
                                End If
                            Else
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
                                End If
                            End If
                        Next

                        Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
                        iCount = iCount + 1
                        rstFindRecords.moveNext()
                    Loop
	
                    ' Release the ADO recordset object.
                    rstFindRecords.close()
                End If

            End If
            rstFindRecords = Nothing

            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If cmdGetFindRecords.Parameters("error").Value <> 0 Then
                'Session("ErrorTitle") = "Book Course Find Page"
                'Session("ErrorText") = "Error reading records definition."
                'Response.Clear	  
                'Response.Redirect("error.asp")
            End If

            Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtErrorMessage2 name=txtErrorMessage2 value=""" & Replace(cmdGetFindRecords.Parameters("errorMessage").Value, """", "&quot;") & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtStatus name=txtStatus value=""" & Replace(cmdGetFindRecords.Parameters("status").Value, """", "&quot;") & """>" & vbCrLf)
			
            cmdGetFindRecords = Nothing

        ElseIf Session("optionAction") = "LOADADDFROMWAITINGLIST" Then
            sThousandColumns = ""
			
            cmdThousandFindColumns = CreateObject("ADODB.Command")
            cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
            cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
            cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
            cmdThousandFindColumns.CommandTimeout = 180
		
            prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdThousandFindColumns.Parameters.Append(prmError)

            prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
            cmdThousandFindColumns.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))

            prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
            cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
            Err.Clear()
            cmdThousandFindColumns.Execute()

            If (Err.Number <> 0) Then
                sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
            End If
	
            ' Release the ADO command object.
            cmdThousandFindColumns = Nothing

            cmdGetFindRecords = CreateObject("ADODB.Command")
            cmdGetFindRecords.CommandText = "sp_ASRIntGetAddFromWaitingListRecords"
            cmdGetFindRecords.CommandType = 4 ' Stored procedure
            cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
            cmdGetFindRecords.CommandTimeout = 180
			
            prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("optionViewID"))

            prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmOrderID)
            prmOrderID.value = CleanNumeric(Session("optionOrderID"))

            prmCourseRecordID = cmdGetFindRecords.CreateParameter("CourseRecordID", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmCourseRecordID)
            prmCourseRecordID.value = CleanNumeric(Session("optionRecordID"))

            prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmError)

            prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
            cmdGetFindRecords.Parameters.Append(prmReqRecs)
            prmReqRecs.value = CleanNumeric(Session("FindRecords"))

            prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

            prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
            cmdGetFindRecords.Parameters.Append(prmIsLastPage)

            prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmLocateValue)
            prmLocateValue.value = Session("optionLocateValue")

            prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnType)

            prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
            cmdGetFindRecords.Parameters.Append(prmAction)
            prmAction.value = Session("optionPageAction")

            prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

            prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
            cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
            prmFirstRecPos.value = CleanNumeric(Session("optionFirstRecPos"))

            prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
            cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
            prmCurrentRecCount.value = CleanNumeric(Session("optionCurrentRecCount"))

            prmErrorMessage = cmdGetFindRecords.CreateParameter("errorMessage", 200, 2, 8000) ' 200=varchar, 2=output,8000=size
            cmdGetFindRecords.Parameters.Append(prmErrorMessage)

            prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnSize)

            prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
            cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

            Err.Clear()
            rstFindRecords = cmdGetFindRecords.Execute
	
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                If rstFindRecords.state = adStateOpen Then
                    iCount = 0
                    Do While Not rstFindRecords.EOF
                        sAddString = ""
						
                        For iloop = 0 To (rstFindRecords.fields.count - 1)
                            If iloop > 0 Then
                                sAddString = sAddString & "	"
                            End If
							
                            If iCount = 0 Then
                                sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                                Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
                            End If
							
                            If rstFindRecords.fields(iloop).type = 135 Then
                                ' Field is a date so format as such.
                                sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
                            ElseIf rstFindRecords.fields(iloop).type = 131 Then
                                ' Field is a numeric so format as such.
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, True)
                                    Else
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, False)
                                    End If
                                    sTemp = Replace(sTemp, ".", "x")
                                    sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
                                    sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
                                    sAddString = sAddString & sTemp
                                End If
                            Else
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
                                End If
                            End If
                        Next

                        Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
                        iCount = iCount + 1
                        rstFindRecords.moveNext()
                    Loop
	
                    ' Release the ADO recordset object.
                    rstFindRecords.close()
                End If

            End If
            rstFindRecords = Nothing

            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If cmdGetFindRecords.Parameters("error").Value <> 0 Then
                'Session("ErrorTitle") = "Add From Waiting List Find Page"
                'Session("ErrorText") = "Error reading records definition."
                'Response.Clear	  
                'Response.Redirect("error.asp")
            End If

            Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtErrorMessage2 name=txtErrorMessage2 value=""" & Replace(cmdGetFindRecords.Parameters("errorMessage").Value, """", "&quot;") & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)
			
            cmdGetFindRecords = Nothing

        ElseIf Session("optionAction") = "SELECTTRANSFERBOOKING_2" Then
		
        
        
            cmdTransferCourse = CreateObject("ADODB.Command")
            cmdTransferCourse.CommandText = "sp_ASRIntTransferCourse"
            cmdTransferCourse.CommandType = 4 ' Stored procedure
            cmdTransferCourse.CommandTimeout = 180
            cmdTransferCourse.ActiveConnection = Session("databaseConnection")
					
            prmTBRecordID = cmdTransferCourse.CreateParameter("TBRecordID", 3, 1) ' 3=integer, 1=input
            cmdTransferCourse.Parameters.Append(prmTBRecordID)
            prmTBRecordID.value = CleanNumeric(Session("optionRecordID"))

            prmCourseRecordID = cmdTransferCourse.CreateParameter("CourseRecordID", 3, 1) ' 3=integer, 1=input
            cmdTransferCourse.Parameters.Append(prmCourseRecordID)
            prmCourseRecordID.value = CleanNumeric(Session("optionLinkRecordID"))

            fDeadlock = True
            Do While fDeadlock
                fDeadlock = False
									
                cmdTransferCourse.ActiveConnection.Errors.Clear()
									
                ' Run the insert stored procedure.
                cmdTransferCourse.Execute()

                If cmdTransferCourse.ActiveConnection.Errors.Count > 0 Then
                    For iLoop = 1 To cmdTransferCourse.ActiveConnection.Errors.Count
                        sErrMsg = formatError(cmdTransferCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)

                        If (cmdTransferCourse.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
                            (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                                (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                            ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
                            ' The error is for a deadlock.
                            ' Sorry about having to use the err.description to trap the error but the err.number
                            ' is not specific and MSDN suggests using the err.description.
                            If (iRetryCount < iRETRIES) And (cmdTransferCourse.ActiveConnection.Errors.Count = 1) Then
                                iRetryCount = iRetryCount + 1
                                fDeadlock = True
                            Else
                                If Len(sNonFatalErrorDescription) > 0 Then
                                    sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf
                                End If
                                sNonFatalErrorDescription = sNonFatalErrorDescription & "Another user is deadlocking the database."
                                fOK = False
                            End If
                        ElseIf UCase(cmdTransferCourse.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
                            '"SQL Mail session is not started."
                            'Ignore this error
                            'ElseIf (cmdTransferCourse.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
                            '	(UCase(Left(cmdTransferCourse.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
                            '"EXECUTE permission denied on object 'xp_sendmail'"
                            'Ignore this error
					
                        Else
                            sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf & _
                                formatError(cmdTransferCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)
                            fOK = False
                        End If
                    Next

                    cmdTransferCourse.ActiveConnection.Errors.Clear()
												
                    If Not fOK Then
                        sNonFatalErrorDescription = "The booking could not be transferred." & vbCrLf & sNonFatalErrorDescription
                        Session("optionAction") = "TRANSFERBOOKINGERROR"
                    End If
                Else
                    Session("optionAction") = "TRANSFERBOOKINGSUCCESS"
                End If
            Loop
            cmdTransferCourse = Nothing

        ElseIf Session("optionAction") = "GETBULKBOOKINGSELECTION" Then
            If UCase(Session("optionPageAction")) = "FILTER" Then
                objUtilities = Session("UtilitiesObject")

                objUtilities.Connection = Session("databaseConnection")

                j = 0
                ReDim Preserve aPrompts(1, 0)
                sPrompts = Session("optionPromptSQL")
                If Len(Session("optionPromptSQL")) > 0 Then
                    Do While Len(sPrompts) > 0
                        iIndex1 = InStr(sPrompts, vbTab)
					
                        If iIndex1 > 0 Then
                            iIndex2 = InStr(iIndex1 + 1, sPrompts, vbTab)
					
                            If iIndex2 > 0 Then
                                ReDim Preserve aPrompts(1, j)
								
                                aPrompts(0, j) = Left(sPrompts, iIndex1 - 1)
                                aPrompts(1, j) = Mid(sPrompts, iIndex1 + 1, iIndex2 - iIndex1 - 1)
								
                                sPrompts = Mid(sPrompts, iIndex2 + 1)
								
                                j = j + 1
                            End If
                        End If
                    Loop
                End If
                Session("optionPromptSQL") = objUtilities.GetFilteredIDs(Session("optionRecordID"), aPrompts)

                objUtilities = Nothing
            End If

            cmdBulkBooking = CreateObject("ADODB.Command")
            cmdBulkBooking.CommandText = "sp_ASRIntGetBulkBookingRecords"
            cmdBulkBooking.CommandType = 4 ' Stored procedure
            cmdBulkBooking.CommandTimeout = 180
            cmdBulkBooking.ActiveConnection = Session("databaseConnection")

            prmSelectionType = cmdBulkBooking.CreateParameter("selectionType", 200, 1, 8000) '200=varchar,1=input,8000=size
            cmdBulkBooking.Parameters.Append(prmSelectionType)
            prmSelectionType.value = Session("optionPageAction")

            prmSelectionID = cmdBulkBooking.CreateParameter("selectionID", 3, 1) '3=integer,1=input
            cmdBulkBooking.Parameters.Append(prmSelectionID)
            prmSelectionID.value = CleanNumeric(Session("optionRecordID"))

            prmSelectedIDs = cmdBulkBooking.CreateParameter("selectedIDs", 200, 1, 8000) '200=varchar,1=input,8000=size
            cmdBulkBooking.Parameters.Append(prmSelectedIDs)
            prmSelectedIDs.value = Session("optionValue")

            prmPromptSQL = cmdBulkBooking.CreateParameter("promptSQL", 200, 1, 8000) '200=varchar,1=input,8000=size
            cmdBulkBooking.Parameters.Append(prmPromptSQL)
            If Len(Session("optionPromptSQL")) = 0 Then
                prmPromptSQL.value = ""
            Else
                prmPromptSQL.value = Session("optionPromptSQL")
            End If
						
            prmErrMsg = cmdBulkBooking.CreateParameter("errMsg", 200, 2, 8000) '200=varchar,2=output,8000=size
            cmdBulkBooking.Parameters.Append(prmErrMsg)
						
            objUtilities = Session("UtilitiesObject")

            objUtilities.UDFFunctions(True)
			
            Err.Clear()
            rstFindRecords = cmdBulkBooking.Execute
			
            objUtilities.UDFFunctions(False)
			
            objUtilities = Nothing
			
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
            End If
			
            If Len(sErrorDescription) = 0 Then
                If rstFindRecords.state = adStateOpen Then
                    iCount = 0
                    Do While Not rstFindRecords.EOF
                        sAddString = ""
						
                        For iloop = 0 To (rstFindRecords.fields.count - 1)
                            If iloop > 0 Then
                                sAddString = sAddString & "	"
                            End If
							
                            If iCount = 0 Then
                                sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                                Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
                            End If
							
                            If rstFindRecords.fields(iloop).type = 135 Then
                                ' Field is a date so format as such.
                                sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
                            ElseIf rstFindRecords.fields(iloop).type = 131 Then
                                ' Field is a numeric so format as such.
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    If Mid(Session("option1000SepCols"), iloop + 1, 1) = "1" Then
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, True)
                                    Else
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, False)
                                    End If
                                    sTemp = Replace(sTemp, ".", "x")
                                    sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
                                    sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
                                    sAddString = sAddString & sTemp
                                End If
                            Else
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
                                End If
                            End If
                        Next

                        Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
                        iCount = iCount + 1
                        rstFindRecords.moveNext()
                    Loop
	
                    ' Release the ADO recordset object.
                    rstFindRecords.close()
                End If

            End If
            rstFindRecords = Nothing

            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If Len(cmdGetFindRecords.Parameters("errMsg").Value) > 0 Then
                sErrorDescription = cmdGetFindRecords.Parameters("errMsg").Value
            End If
			
            cmdBulkBooking = Nothing

        ElseIf Session("optionAction") = "GETPICKLISTSELECTION" Then
            If UCase(Session("optionPageAction")) = "FILTER" Then
                objUtilities = Session("UtilitiesObject")

                objUtilities.Connection = Session("databaseConnection")

                j = 0
                ReDim Preserve aPrompts(1, 0)
                sPrompts = Session("optionPromptSQL")
                If Len(Session("optionPromptSQL")) > 0 Then
                    Do While Len(sPrompts) > 0
                        iIndex1 = InStr(sPrompts, vbTab)
					
                        If iIndex1 > 0 Then
                            iIndex2 = InStr(iIndex1 + 1, sPrompts, vbTab)
					
                            If iIndex2 > 0 Then
                                ReDim Preserve aPrompts(1, j)
								
                                aPrompts(0, j) = Left(sPrompts, iIndex1 - 1)
                                aPrompts(1, j) = Mid(sPrompts, iIndex1 + 1, iIndex2 - iIndex1 - 1)
								
                                sPrompts = Mid(sPrompts, iIndex2 + 1)
								
                                j = j + 1
                            End If
                        End If
                    Loop
                End If
                Session("optionPromptSQL") = objUtilities.GetFilteredIDs(Session("optionRecordID"), aPrompts)

                objUtilities = Nothing
            End If
		
            cmdPicklist = CreateObject("ADODB.Command")
            cmdPicklist.CommandText = "sp_ASRIntGetSelectedPicklistRecords"
            cmdPicklist.CommandType = 4 ' Stored procedure
            cmdPicklist.CommandTimeout = 180
            cmdPicklist.ActiveConnection = Session("databaseConnection")

            prmSelectionType = cmdPicklist.CreateParameter("selectionType", 200, 1, 8000) '200=varchar,1=input,8000=size
            cmdPicklist.Parameters.Append(prmSelectionType)
            prmSelectionType.value = Session("optionPageAction")

            prmSelectionID = cmdPicklist.CreateParameter("selectionID", 3, 1) '3=integer,1=input
            cmdPicklist.Parameters.Append(prmSelectionID)
            prmSelectionID.value = CleanNumeric(Session("optionRecordID"))
			
            prmSelectedIDs = cmdPicklist.CreateParameter("selectedIDs", 200, 1, 2147483646) '200=varchar,1=input,8000=size
            cmdPicklist.Parameters.Append(prmSelectedIDs)
            prmSelectedIDs.value = Session("optionValue")

            prmPromptSQL = cmdPicklist.CreateParameter("promptSQL", 200, 1, 2147483646) '200=varchar,1=input,8000=size
            cmdPicklist.Parameters.Append(prmPromptSQL)
            If Len(Session("optionPromptSQL")) = 0 Then
                prmPromptSQL.value = ""
            Else
                prmPromptSQL.value = Session("optionPromptSQL")
            End If
						
            prmTableID = cmdPicklist.CreateParameter("tableID", 3, 1) '3=integer,1=input
            cmdPicklist.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmErrMsg = cmdPicklist.CreateParameter("errMsg", 200, 2, 2147483646) '200=varchar,2=output,8000=size
            cmdPicklist.Parameters.Append(prmErrMsg)

            prmExpectedCount = cmdPicklist.CreateParameter("expectedCount", 3, 2) '3=integer,2=output
            cmdPicklist.Parameters.Append(prmExpectedCount)

            objUtilities = Session("UtilitiesObject")

            objUtilities.UDFFunctions(True)
		
            Err.Clear()
            rstFindRecords = cmdPicklist.Execute

            objUtilities.UDFFunctions(False)
			
            objUtilities = Nothing
	
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading the records." & vbCrLf & formatError(Err.Description)
            End If
			
            If Len(sErrorDescription) = 0 Then
                If rstFindRecords.state = adStateOpen Then
                    iCount = 0
                    Do While Not rstFindRecords.EOF
                        sAddString = ""
						
                        For iloop = 0 To (rstFindRecords.fields.count - 1)
                            If iloop > 0 Then
                                sAddString = sAddString & "	"
                            End If
							
                            If iCount = 0 Then
                                sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                                Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
                            End If
							
                            If rstFindRecords.fields(iloop).type = 135 Then
                                ' Field is a date so format as such.
                                sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
                            ElseIf rstFindRecords.fields(iloop).type = 131 Then
                                ' Field is a numeric so format as such.
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    If Mid(Session("option1000SepCols"), iloop + 1, 1) = "1" Then
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, True)
                                    Else
                                        sTemp = ""
                                        sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).numericScale, True, False, False)
                                    End If
                                    sTemp = Replace(sTemp, ".", "x")
                                    sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
                                    sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
                                    sAddString = sAddString & sTemp
                                End If
                            Else
                                If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
                                    sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
                                End If
                            End If
                        Next

                        Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
                        iCount = iCount + 1
                        rstFindRecords.moveNext()
                    Loop
	
                    ' Release the ADO recordset object.
                    rstFindRecords.close()
                End If

            End If
            rstFindRecords = Nothing

            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If Len(cmdPicklist.Parameters("errMsg").Value) > 0 Then
                sErrorDescription = cmdPicklist.Parameters("errMsg").Value
            End If

            Response.Write("<INPUT type='hidden' id=txtExpectedCount name=txtExpectedCount value=" & cmdPicklist.Parameters("expectedCount").Value & ">" & vbCrLf)
			
            cmdPicklist = Nothing

        ElseIf Session("optionAction") = "SELECTBULKBOOKINGS_2" Then
            cmdBulkBook = CreateObject("ADODB.Command")
            cmdBulkBook.CommandText = "sp_ASRIntMakeBulkBookings"
            cmdBulkBook.CommandType = 4 ' Stored procedure
            cmdBulkBook.CommandTimeout = 180
            cmdBulkBook.ActiveConnection = Session("databaseConnection")
					
            prmCourseRecordID = cmdBulkBook.CreateParameter("CourseRecordID", 3, 1) ' 3=integer, 1=input
            cmdBulkBook.Parameters.Append(prmCourseRecordID)
            prmCourseRecordID.value = CleanNumeric(Session("optionRecordID"))

            prmEmployeeRecordIDs = cmdBulkBook.CreateParameter("EmployeeRecordIDs", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
            cmdBulkBook.Parameters.Append(prmEmployeeRecordIDs)
            prmEmployeeRecordIDs.value = Session("optionLinkRecordID")

            prmStatus = cmdBulkBook.CreateParameter("Status", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
            cmdBulkBook.Parameters.Append(prmStatus)
            prmStatus.value = Session("optionValue")

            fDeadlock = True
            Do While fDeadlock
                fDeadlock = False
									
                cmdBulkBook.ActiveConnection.Errors.Clear()
									
                ' Run the insert stored procedure.
                cmdBulkBook.Execute()

                If cmdBulkBook.ActiveConnection.Errors.Count > 0 Then
                    For iLoop = 1 To cmdBulkBook.ActiveConnection.Errors.Count
                        sErrMsg = formatError(cmdBulkBook.ActiveConnection.Errors.Item(iLoop - 1).Description)

                        If (cmdBulkBook.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
                            (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                                (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                            ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
                            ' The error is for a deadlock.
                            ' Sorry about having to use the err.description to trap the error but the err.number
                            ' is not specific and MSDN suggests using the err.description.
                            If (iRetryCount < iRETRIES) And (cmdBulkBook.ActiveConnection.Errors.Count = 1) Then
                                iRetryCount = iRetryCount + 1
                                fDeadlock = True
                            Else
                                If Len(sNonFatalErrorDescription) > 0 Then
                                    sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf
                                End If
                                sNonFatalErrorDescription = sNonFatalErrorDescription & "Another user is deadlocking the database."
                                fOK = False
                            End If
                        ElseIf UCase(cmdBulkBook.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
                            '"SQL Mail session is not started."
                            'Ignore this error
                            'ElseIf (cmdTransferCourse.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
                            '	(UCase(Left(cmdTransferCourse.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
                            '"EXECUTE permission denied on object 'xp_sendmail'"
                            'Ignore this error
					
                        Else
                            sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf & _
                                formatError(cmdBulkBook.ActiveConnection.Errors.Item(iLoop - 1).Description)
                            fOK = False
                        End If
                    Next

                    cmdBulkBook.ActiveConnection.Errors.Clear()
												
                    If Not fOK Then
                        sNonFatalErrorDescription = "Unable to create booking record." & vbCrLf & sNonFatalErrorDescription
                        Session("optionAction") = "BULKBOOKINGERROR"
                    End If
                Else
                    Session("optionAction") = "BULKBOOKINGSUCCESS"
                End If
            Loop
            cmdBulkBook = Nothing

        ElseIf (Session("optionAction") = "LOADEXPRFIELDCOLUMNS") Or _
            (Session("optionAction") = "LOADEXPRLOOKUPCOLUMNS") Then
		
            cmdExprColumns = CreateObject("ADODB.Command")
            cmdExprColumns.CommandText = "sp_ASRIntGetExprColumns"
            cmdExprColumns.CommandType = 4 ' Stored procedure
            cmdExprColumns.CommandTimeout = 180
            cmdExprColumns.ActiveConnection = Session("databaseConnection")
					
            prmTableID = cmdExprColumns.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
            cmdExprColumns.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("optionTableID"))

            prmComponentType = cmdExprColumns.CreateParameter("componentType", 3, 1) ' 3=integer, 1=input
            cmdExprColumns.Parameters.Append(prmComponentType)
            If Session("optionAction") = "LOADEXPRFIELDCOLUMNS" Then
                prmComponentType.value = 1
            Else
                prmComponentType.value = 0
            End If

            prmOnlyNumerics = cmdExprColumns.CreateParameter("onlyNumerics", 3, 1) ' 3=integer, 1=input
            cmdExprColumns.Parameters.Append(prmOnlyNumerics)
            prmOnlyNumerics.value = CleanNumeric(Session("optionOnlyNumerics"))

            Err.Clear()
            rstExprColumns = cmdExprColumns.Execute
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading component columns." & vbCrLf & formatError(Err.Description)
            Else
                If rstExprColumns.state <> 0 Then
                    ' Read recordset values.
                    iCount = 0
                    Do While Not rstExprColumns.EOF
                        iCount = iCount + 1
                        Response.Write("<INPUT type='hidden' id=txtColumn_" & iCount & " name=txtColumn_" & iCount & " value=""" & rstExprColumns.fields("definitionString").value & """>" & vbCrLf)
                        rstExprColumns.MoveNext()
                    Loop

                    ' Release the ADO recordset object.
                    rstExprColumns.close()
                End If
                rstExprColumns = Nothing
            End If
            cmdExprColumns = Nothing

        ElseIf Session("optionAction") = "LOADEXPRLOOKUPVALUES" Then
		
            cmdExprValues = CreateObject("ADODB.Command")
            cmdExprValues.CommandText = "sp_ASRIntGetExprLookupValues"
            cmdExprValues.CommandType = 4 ' Stored procedure
            cmdExprValues.CommandTimeout = 180
            cmdExprValues.ActiveConnection = Session("databaseConnection")
					
            prmColumnID = cmdExprValues.CreateParameter("columnID", 3, 1) ' 3=integer, 1=input
            cmdExprValues.Parameters.Append(prmColumnID)
            prmColumnID.value = CleanNumeric(Session("optionColumnID"))

            prmDataType = cmdExprValues.CreateParameter("dataType", 3, 2) ' 3=integer, 2=output
            cmdExprValues.Parameters.Append(prmDataType)

            Err.Clear()
            rstExprValues = cmdExprValues.Execute
            If (Err.Number <> 0) Then
                sErrorDescription = "Error reading component values." & vbCrLf & formatError(Err.Description)
            Else
                If rstExprValues.state <> 0 Then
                    ' Read recordset values.
                    iCount = 0
                    Do While Not rstExprValues.EOF
                        iCount = iCount + 1
                        Response.Write("<INPUT type='hidden' id=txtValue_" & iCount & " name=txtValue_" & iCount & " value=""" & rstExprValues.fields("lookupValue").value & """>" & vbCrLf)
                        rstExprValues.MoveNext()
                    Loop

                    ' Release the ADO recordset object.
                    rstExprValues.close()
                End If
                rstExprValues = Nothing

                ' NB. IMPORTANT ADO NOTE.
                ' When calling a stored procedure which returns a recordset AND has output parameters
                ' you need to close the recordset and set it to nothing before using the output parameters. 
                Response.Write("<INPUT type='hidden' id=txtLookupDataType name=txtLookupDataType value=" & cmdExprValues.Parameters("dataType").Value & ">" & vbCrLf)
            End If
            cmdExprValues = Nothing

        End If
        '	end if

        Response.Write("<INPUT type='hidden' id=txtOptionAction name=txtOptionAction value=" & Session("optionAction") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtOptionTableID name=txtOptionTableID value=" & Session("optionTableID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtOptionViewID name=txtOptionViewID value=" & Session("optionViewID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtOptionOrderID name=txtOptionOrderID value=" & Session("optionOrderID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtOptionColumnID name=txtOptionColumnID value=" & Session("optionColumnID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtOptionLocateValue name=txtOptionLocateValue value=""" & Replace(Session("optionLocateValue"), """", "&quot;") & """>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
        Response.Write("<INPUT type='hidden' id=txtNonFatalErrorDescription name=txtNonFatalErrorDescription value=""" & sNonFatalErrorDescription & """>")
    %>
</form>

<script runat="server" language="vb">

    Function formatError(psErrMsg)
        Dim iStart As Integer
        Dim iFound As Integer
  
        iFound = 0
        Do
            iStart = iFound
            iFound = InStr(iStart + 1, psErrMsg, "]")
        Loop While iFound > 0
  
        If (iStart > 0) And (iStart < Len(Trim(psErrMsg))) Then
            formatError = Trim(Mid(psErrMsg, iStart + 1))
        Else
            formatError = psErrMsg
        End If
    End Function

    Function convertSQLDateToLocale(psDate)
        Dim sLocaleFormat As String
        Dim iIndex As Integer
	
        If Len(psDate) > 0 Then
            sLocaleFormat = Session("LocaleDateFormat")
		
            iIndex = InStr(sLocaleFormat, "dd")
            If iIndex > 0 Then
                If Day(psDate) < 10 Then
                    sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                        "0" & Day(psDate) & Mid(sLocaleFormat, iIndex + 2)
                Else
                    sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                        Day(psDate) & Mid(sLocaleFormat, iIndex + 2)
                End If
            End If
		
            iIndex = InStr(sLocaleFormat, "mm")
            If iIndex > 0 Then
                If Month(psDate) < 10 Then
                    sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                        "0" & Month(psDate) & Mid(sLocaleFormat, iIndex + 2)
                Else
                    sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                        Month(psDate) & Mid(sLocaleFormat, iIndex + 2)
                End If
            End If
		
            iIndex = InStr(sLocaleFormat, "yyyy")
            If iIndex > 0 Then
                sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                    Year(psDate) & Mid(sLocaleFormat, iIndex + 4)
            End If

            convertSQLDateToLocale = sLocaleFormat
        Else
            convertSQLDateToLocale = ""
        End If
    End Function

</script>

<script type="text/javascript">
    optiondata_onload()
</script>
