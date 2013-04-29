r<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import namespace="DMI.NET" %>
<%
	Response.Expires = -1
%>

<script type="text/javascript">

	function tbBulkBookingSelectionData_onload() {
		
		var frmData = document.getElementById("frmData");				
		
		if (document.getElementById("txtLoading").value == "True") {			
			loadAddRecords();		// window.parent.loadAddRecords();
			return;
		}

		var sFatalErrorMsg = document.getElementById("txtErrorDescription").value;
		if (sFatalErrorMsg.length > 0) {
			OpenHR.messageBox(sFatalErrorMsg);
			//window.parent.close();
		} else {
			// Do nothing if the menu controls are not yet instantiated.
			var sErrorMsg = document.getElementById("txtErrorMessage").value;
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
				for (var i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 8);
					if (sControlName == "txtData_") {
						grdLinkFind.addItem(dataCollection.item(i).value);
						fRecordAdded = true;
						iCount = iCount + 1;
					}
				}
			}
			grdLinkFind.redraw = true;

			frmData.txtRecordCount.value = iCount;

			tbrefreshControls();

			// Get menu.asp to refresh the menu.
			tbrefreshMenu();
		}
	}
</script>

<script type="text/javascript">
	function refreshData() {		
		var frmGetData = document.getElementById("frmGetData");
		OpenHR.submitForm(frmGetData);
	}
</script>

<div>

<FORM action="tbBulkBookingSelectionData_Submit" method="post" id="frmGetData" name="frmGetData">
<!--	<INPUT type="hidden" id=txtAction name=txtAction>-->
	<INPUT type="hidden" id=txtTableID name=txtTableID>
	<INPUT type="hidden" id=txtViewID name=txtViewID>
	<INPUT type="hidden" id=txtOrderID name=txtOrderID>
<!--	<INPUT type="hidden" id=txtColumnID name=txtColumnID>-->
	<INPUT type="hidden" id=txtPageAction name=txtPageAction>
	<INPUT type="hidden" id=txtFirstRecPos name=txtFirstRecPos>
	<INPUT type="hidden" id=txtCurrentRecCount name=txtCurrentRecCount>
	<INPUT type="hidden" id=txtGotoLocateValue name=txtGotoLocateValue>
<!--	<INPUT type="hidden" id=txtRecordID name=txtRecordID>
	<INPUT type="hidden" id=txtLinkRecordID name=txtLinkRecordID>
	<INPUT type="hidden" id=txtValue name=txtValue>
	<INPUT type="hidden" id=txtSQL name=txtSQL>
	<INPUT type="hidden" id=txtPromptSQL name=txtPromptSQL>-->
</FORM>

<FORM id="frmUseful" name="frmUseful">
	<INPUT type="hidden" id="txtLoading" name="txtLoading" value="<%=session("tbSelectionDataLoading")%>">
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

	Dim iRETRIES = 5
	Dim iRetryCount = 0
	Dim sErrorDescription = ""

	Response.Write("<INPUT type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

	' Get the required record count if we have a query.
	if session("tbSelectionDataLoading") = false then

		Dim sThousandColumns = ""
			
		Dim cmdThousandFindColumns = CreateObject("ADODB.Command")
		cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
		cmdThousandFindColumns.CommandType = 4 ' Stored Procedure
		cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
		cmdThousandFindColumns.CommandTimeout = 180
		
		Dim prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2)	' 11=bit, 2=output
		cmdThousandFindColumns.Parameters.Append(prmError)

		Dim prmTableID2 = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
		cmdThousandFindColumns.Parameters.Append(prmTableID2)
		prmTableID2.value = CleanNumeric(Session("tableID"))

		Dim prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
		cmdThousandFindColumns.Parameters.Append(prmViewID)
		prmViewID.value = cleanNumeric(session("viewID"))

		Dim prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
		cmdThousandFindColumns.Parameters.Append(prmOrderID)
		prmOrderID.value = cleanNumeric(session("orderID"))

		Dim prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
		cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
		Err.Clear()
		cmdThousandFindColumns.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
		End If

		if len(sErrorDescription) = 0 then
			sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value			
		end if
	
		' Release the ADO command object.
		cmdThousandFindColumns = Nothing

		Dim cmdGetFindRecords = CreateObject("ADODB.Command")
		cmdGetFindRecords.CommandText = "sp_ASRIntGetLinkFindRecords"
		cmdGetFindRecords.CommandType = 4 ' Stored procedure
		cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
		cmdGetFindRecords.CommandTimeout = 180
			
		Dim prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
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

		Dim prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
		cmdGetFindRecords.Parameters.Append(prmReqRecs)
		prmReqRecs.value = cleanNumeric(session("FindRecords"))

		Dim prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
		cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

		Dim prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
		cmdGetFindRecords.Parameters.Append(prmIsLastPage)

		Dim prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 2147483646)
		cmdGetFindRecords.Parameters.Append(prmLocateValue)
		prmLocateValue.value = session("locateValue")

		Dim prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
		cmdGetFindRecords.Parameters.Append(prmColumnType)

		Dim prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 100)
		cmdGetFindRecords.Parameters.Append(prmAction)
		prmAction.value = session("pageAction")

		Dim prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
		cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

		Dim prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
		cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
		prmFirstRecPos.value = cleanNumeric(session("firstRecPos"))

		Dim prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
		cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
		prmCurrentRecCount.value = cleanNumeric(session("currentRecCount"))

		Dim prmExcludedIDs = cmdGetFindRecords.CreateParameter("excludedIDs", 200, 1, 2147483646)	' 200=varchar, 1=input, 8000=size
		cmdGetFindRecords.Parameters.Append(prmExcludedIDs)
		prmExcludedIDs.value = ""

		Dim prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
		cmdGetFindRecords.Parameters.Append(prmColumnSize)

		Dim prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
		cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

		Dim rstFindRecords = cmdGetFindRecords.Execute
	
		If (Err.Number <> 0) Then
			sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
		End If

		if len(sErrorDescription) = 0 then
			If rstFindRecords.state = 1 Then	' adStateOpen = 1.
				Dim iCount = 0
				Dim sColDef = ""
				Dim sTemp = ""
				
				Do While Not rstFindRecords.EOF
					Dim sAddString = ""
					
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
		if cmdGetFindRecords.Parameters("error").Value <> 0 then
		  Session("ErrorTitle") = "Bulk Booking Selection Find Page"
		  Session("ErrorText") = "Error reading employee records definition."
			Response.Clear	  
			Response.Redirect("error")
		end if

		Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

		cmdGetFindRecords = Nothing
			
	end if

'	Response.Write "<INPUT type='hidden' id=txtAction name=txtAction value=" & session("Action") & ">" & vbcrlf
'	Response.Write "<INPUT type='hidden' id=txtTableID name=txtTableID value=" & session("TableID") & ">" & vbcrlf
'	Response.Write "<INPUT type='hidden' id=txtViewID name=txtViewID value=" & session("ViewID") & ">" & vbcrlf
'	Response.Write "<INPUT type='hidden' id=txtOrderID name=txtOrderID value=" & session("OrderID") & ">" & vbcrlf
	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
%>
</FORM>
</div>

<script type="text/javascript"> tbBulkBookingSelectionData_onload();</script>
