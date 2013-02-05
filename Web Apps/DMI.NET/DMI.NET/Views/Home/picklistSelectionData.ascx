<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>


<div>


<script type="text/javascript">
<!--
    function picklistSelectionData_window_onload() {

        debugger;
        $("#dataframe").attr("data-framesource", "PICKLISTSELECTIONDATA");

        if (frmUseful.txtLoading.value == "True") {
            window.parent.loadAddRecords();
            return;
        }

        var sFatalErrorMsg = frmData.txtErrorDescription.value
        if (sFatalErrorMsg.length > 0) {
            OpenHR.messageBox(sFatalErrorMsg);
            window.parent.close();
        } else {
            // Do nothing if the menu controls are not yet instantiated.
            var sErrorMsg = frmData.txtErrorMessage.value;
            if (sErrorMsg.length > 0) {
                // We've got an error so don't update the record edit form.

                // Get menu.asp to refresh the menu.
                menu_refreshMenu();
                OpenHR.messageBox(sErrorMsg);
            }

            //		var sAction = frmData.txtAction.value;

            // Refresh the link find grid with the data if required.
            var grdLinkFind = document.getElementById("ssOleDBGridSelRecords");

            grdLinkFind.redraw = false;
            grdLinkFind.removeAll();
            grdLinkFind.columns.removeAll();

            var dataCollection = frmData.elements;
            var sControlName;
            var sColumnName;
            var iColumnType;
            var iCount;

            // Configure the grid columns.
            if (dataCollection != null) {
                for (i = 0; i < dataCollection.length; i++) {
                    sControlName = dataCollection.item(i).name;
                    sControlName = sControlName.substr(0, 10);
                    if (sControlName == "txtColDef_") {
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
            iCount = 0;
            if (dataCollection != null) {
                for (i = 0; i < dataCollection.length; i++) {
                    sControlName = dataCollection.item(i).name;
                    sControlName = sControlName.substr(0, 8);
                    if (sControlName == "txtData_") {
                        grdLinkFind.addItem(dataCollection.item(i).value);
                        fRecordAdded = true;
                        iCount = iCount + 1
                    }
                }
            }
            grdLinkFind.redraw = true;

            frmData.txtRecordCount.value = iCount;

            refreshControls();

            // Get menu.asp to refresh the menu.
            menu_refreshMenu();
        }
    }
-->
</script>

<script type="text/javascript">
    function refreshData() {
        OpenHR.submitForm(frmGetData);
    }
</script>

<FORM action="picklistSelectionData_Submit" method=post id=frmGetData name=frmGetData>
	<INPUT type="hidden" id=txtTableID name=txtTableID>
	<INPUT type="hidden" id=txtViewID name=txtViewID>
	<INPUT type="hidden" id=txtOrderID name=txtOrderID>
	<INPUT type="hidden" id=txtPageAction name=txtPageAction>
	<INPUT type="hidden" id=txtFirstRecPos name=txtFirstRecPos>
	<INPUT type="hidden" id=txtCurrentRecCount name=txtCurrentRecCount>
	<INPUT type="hidden" id=txtGotoLocateValue name=txtGotoLocateValue>
</FORM>

<FORM id=frmUseful name=frmUseful>
	<INPUT type='hidden' id=txtLoading name=txtLoading value=<%=session("picklistSelectionDataLoading")%>>
</FORM>

<FORM id=frmData name=frmData>
<%
	on error resume next
		
    Const DEADLOCK_ERRORNUMBER = -2147467259
    Const DEADLOCK_MESSAGESTART = "YOUR TRANSACTION (PROCESS ID #"
    Const DEADLOCK_MESSAGEEND = ") WAS DEADLOCKED WITH ANOTHER PROCESS AND HAS BEEN CHOSEN AS THE DEADLOCK VICTIM. RERUN YOUR TRANSACTION."
    Const DEADLOCK2_MESSAGESTART = "TRANSACTION (PROCESS ID "
    Const DEADLOCK2_MESSAGEEND = ") WAS DEADLOCKED ON "
    Const SQLMAILNOTSTARTEDMESSAGE = "SQL MAIL SESSION IS NOT STARTED."

    Const iRETRIES = 5
    Dim iRetryCount = 0
    Dim sErrorDescription As String = ""
    Dim sThousandColumns As String
    
    Dim cmdThousandFindColumns
    Dim prmError
    Dim prmTableID
    Dim prmViewID
    Dim prmOrderID
    Dim prmThousandColumns
    Dim cmdGetFindRecords
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
    Dim iCount As Integer
    Dim sAddString As String
    Dim sColDef As String
    Dim sTemp As String
    
    
    Response.Write("<INPUT type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

	' Get the required record count if we have a query.
	if session("picklistSelectionDataLoading") = false then

		sThousandColumns = ""
			
        cmdThousandFindColumns = Server.CreateObject("ADODB.Command")
		cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
		cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
        cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
		cmdThousandFindColumns.CommandTimeout = 180
		
        prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
        cmdThousandFindColumns.Parameters.Append(prmError)

        prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
        cmdThousandFindColumns.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("tableID"))

        prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
        cmdThousandFindColumns.Parameters.Append(prmViewID)
		prmViewID.value = cleanNumeric(session("viewID"))

        prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
        cmdThousandFindColumns.Parameters.Append(prmOrderID)
		prmOrderID.value = cleanNumeric(session("orderID"))

        prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
        cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
        Err.Clear()
		cmdThousandFindColumns.Execute

        If (Err.Number <> 0) Then
            sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
        End If

        If Len(sErrorDescription) = 0 Then
            sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
        End If
	
		' Release the ADO command object.
        cmdThousandFindColumns = Nothing

        cmdGetFindRecords = Server.CreateObject("ADODB.Command")
		cmdGetFindRecords.CommandText = "sp_ASRIntGetLinkFindRecords"
		cmdGetFindRecords.CommandType = 4 ' Stored procedure
        cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
		cmdGetFindRecords.CommandTimeout = 180
			
        prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
        cmdGetFindRecords.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("tableID"))

        prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
        cmdGetFindRecords.Parameters.Append(prmViewID)
		prmViewID.value = cleanNumeric(session("viewID"))

        prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
        cmdGetFindRecords.Parameters.Append(prmOrderID)
		prmOrderID.value = cleanNumeric(session("orderID"))

        prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
        cmdGetFindRecords.Parameters.Append(prmError)

        prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
        cmdGetFindRecords.Parameters.Append(prmReqRecs)
		prmReqRecs.value = cleanNumeric(session("FindRecords"))

        prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
        cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

        prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
        cmdGetFindRecords.Parameters.Append(prmIsLastPage)

        prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 2147483646)
        cmdGetFindRecords.Parameters.Append(prmLocateValue)
		prmLocateValue.value = session("locateValue")

        prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
        cmdGetFindRecords.Parameters.Append(prmColumnType)

        prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 100)
        cmdGetFindRecords.Parameters.Append(prmAction)
		prmAction.value = session("pageAction")

        prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
        cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

        prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
        cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
		prmFirstRecPos.value = cleanNumeric(session("firstRecPos"))

        prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1) ' 3=integer, 1=input
        cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
		prmCurrentRecCount.value = cleanNumeric(session("currentRecCount"))

        prmExcludedIDs = cmdGetFindRecords.CreateParameter("excludedIDs", 200, 1, 2147483646) ' 200=varchar, 1=input, 8000=size
        cmdGetFindRecords.Parameters.Append(prmExcludedIDs)
		prmExcludedIDs.value = session("selectedIDs1")
		
        prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
        cmdGetFindRecords.Parameters.Append(prmColumnSize)

        prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
        cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

        rstFindRecords = cmdGetFindRecords.Execute
	
        If (Err.Number <> 0) Then
            sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
        End If

		if len(sErrorDescription) = 0 then
            If rstFindRecords.state = 1 Then
                iCount = 0
                Do While Not rstFindRecords.EOF
                    sAddString = ""
					
                    For iloop = 0 To (rstFindRecords.fields.count - 1)
                        If iloop > 0 Then
                            sAddString = sAddString & "	"
                        End If
							
                        If iCount = 0 Then
                            sColDef = Replace(rstFindRecords.fields(iloop).name, "_", " ") & "	" & rstFindRecords.fields(iloop).type
                            Response.Write("<INPUT type='hidden' id=txtColDef_" & iloop & " name=txtColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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

                    Response.Write("<INPUT type='hidden' id=txtData_" & iCount & " name=txtData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
                    iCount = iCount + 1
                    rstFindRecords.moveNext()
                Loop
	
                ' Release the ADO recordset object.
                rstFindRecords.close()
            End If
		end if
        rstFindRecords = Nothing

		' NB. IMPORTANT ADO NOTE.
		' When calling a stored procedure which returns a recordset AND has output parameters
		' you need to close the recordset and set it to nothing before using the output parameters. 
        '		if cmdGetFindRecords.Parameters("error").Value <> 0 then
        'Session("ErrorTitle") = "Picklist Selection Find Page"
        'Session("ErrorText") = "Error reading records definition."
        'Response.Clear()
        'Response.Redirect("error.asp")
        'End If

        Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

        cmdGetFindRecords = Nothing
			
    End If

    Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
%>
</FORM>

<script runat="server">

    Function formatError(psErrMsg)
        Dim iStart
        Dim iFound
  
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
        Dim sLocaleFormat
        Dim iIndex
	
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
    picklistSelectionData_window_onload()
</script>
    
    </div>
