<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">

	function picklistSelectionData_window_onload() {

		$("#picklistdataframe").attr("data-framesource", "PICKLISTSELECTIONDATA");
		$("#workframeset").hide();
		$("#reportframe").show();

		if (frmSelectDataUseful.txtLoading.value == "True") {
			loadAddRecords();
			return;
		}

		var sFatalErrorMsg = frmPicklistData.txtErrorDescription.value;
		if (sFatalErrorMsg.length > 0) {
			OpenHR.messageBox(sFatalErrorMsg);
		} else {
			// Do nothing if the menu controls are not yet instantiated.
			var sErrorMsg = frmPicklistData.txtErrorMessage.value;
			if (sErrorMsg.length > 0) {
				// We've got an error so don't update the record edit form.

				// Get menu.asp to refresh the menu.
				menu_refreshMenu();
				OpenHR.messageBox(sErrorMsg);
			}

			// Refresh the link find grid with the data if required.
			var ssOleDBGridSelRecords = document.getElementById("ssOleDBGridSelRecords");
			ssOleDBGridSelRecords.Redraw = false;

			//ssOleDBGridSelRecords.removeAll();
			ssOleDBGridSelRecords.Columns.RemoveAll();

			var dataCollection = frmPicklistData.elements;
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

							ssOleDBGridSelRecords.Columns.Add(ssOleDBGridSelRecords.Columns.Count);
							ssOleDBGridSelRecords.Columns.Item(ssOleDBGridSelRecords.Columns.Count - 1).Name = sColumnName;
							ssOleDBGridSelRecords.Columns.Item(ssOleDBGridSelRecords.Columns.Count - 1).Caption = sColumnName;

							if (sColumnName == "ID") {
								ssOleDBGridSelRecords.Columns.Item(ssOleDBGridSelRecords.Columns.Count - 1).Visible = false;
							}

							if ((sColumnType == "131") || (sColumnType == "3")) {
								ssOleDBGridSelRecords.Columns.Item(ssOleDBGridSelRecords.Columns.Count - 1).Alignment = 1;
							} else {
								ssOleDBGridSelRecords.Columns.Item(ssOleDBGridSelRecords.Columns.Count - 1).Alignment = 0;
							}
							if (sColumnType == 11) {
								ssOleDBGridSelRecords.Columns.Item(ssOleDBGridSelRecords.Columns.Count - 1).Style = 2;
							} else {
								ssOleDBGridSelRecords.Columns.Item(ssOleDBGridSelRecords.Columns.Count - 1).Style = 0;
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
						ssOleDBGridSelRecords.addItem(dataCollection.item(i).value);
						fRecordAdded = true;
						iCount = iCount + 1;
					}
				}
			}
			ssOleDBGridSelRecords.Redraw = true;

			frmPicklistData.txtRecordCount.value = iCount;

			refreshControls();

		}
	}

	function picklist_refreshData() {

		OpenHR.submitForm(frmPicklistGetData);
	}

</script>

<form action="picklistSelectionData_Submit" method="post" id="frmPicklistGetData" name="frmPicklistGetData">
	<input type="hidden" id="txtTableID" name="txtTableID">
	<input type="hidden" id="txtViewID" name="txtViewID">
	<input type="hidden" id="txtOrderID" name="txtOrderID">
	<input type="hidden" id="txtPageAction" name="txtPageAction">
	<input type="hidden" id="txtFirstRecPos" name="txtFirstRecPos">
	<input type="hidden" id="txtCurrentRecCount" name="txtCurrentRecCount">
	<input type="hidden" id="txtGotoLocateValue" name="txtGotoLocateValue">
</form>

<form id="frmSelectDataUseful" name="frmSelectDataUseful">
	<input type='hidden' id="txtLoading" name="txtLoading" value='<%=session("picklistSelectionDataLoading")%>'>
</form>

<form id="frmPicklistData" name="frmPicklistData">
	<%
		On Error Resume Next
		
		Dim iRetryCount = 0
		Dim sErrorDescription As String = ""
		Dim sThousandColumns As String
		
		Dim cmdThousandFindColumns As ADODB.Command
		Dim prmError As ADODB.Parameter
		Dim prmTableID As ADODB.Parameter
		Dim prmViewID As ADODB.Parameter
		Dim prmOrderID As ADODB.Parameter
		Dim prmThousandColumns As ADODB.Parameter
		Dim cmdGetFindRecords As ADODB.Command
		Dim prmReqRecs As ADODB.Parameter
		Dim prmIsFirstPage As ADODB.Parameter
		Dim prmIsLastPage As ADODB.Parameter
		Dim prmLocateValue As ADODB.Parameter
		Dim prmColumnType As ADODB.Parameter
		Dim prmAction As ADODB.Parameter
		Dim prmTotalRecCount As ADODB.Parameter
		Dim prmFirstRecPos As ADODB.Parameter
		Dim prmCurrentRecCount As ADODB.Parameter
		Dim prmExcludedIDs As ADODB.Parameter
		Dim prmColumnSize As ADODB.Parameter
		Dim prmColumnDecimals As ADODB.Parameter
		Dim rstFindRecords As ADODB.Recordset
		Dim iCount As Integer
		Dim sAddString As String
		Dim sColDef As String
		Dim sTemp As String
		
		
		Response.Write("<input type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

		' Get the required record count if we have a query.
		If Session("picklistSelectionDataLoading") = False Then

			sThousandColumns = ""
			
			cmdThousandFindColumns = New ADODB.Command
			cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
			cmdThousandFindColumns.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
			cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
			cmdThousandFindColumns.CommandTimeout = 180
		
			prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2)	' 11=bit, 2=output
			cmdThousandFindColumns.Parameters.Append(prmError)

			prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
			cmdThousandFindColumns.Parameters.Append(prmTableID)
			prmTableID.Value = CleanNumeric(Session("tableID"))

			prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
			cmdThousandFindColumns.Parameters.Append(prmViewID)
			prmViewID.Value = CleanNumeric(Session("viewID"))

			prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
			cmdThousandFindColumns.Parameters.Append(prmOrderID)
			prmOrderID.Value = CleanNumeric(Session("orderID"))

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

			cmdGetFindRecords = New ADODB.Command
			cmdGetFindRecords.CommandText = "sp_ASRIntGetLinkFindRecords"
			cmdGetFindRecords.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
			cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
			cmdGetFindRecords.CommandTimeout = 180
			
			prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
			cmdGetFindRecords.Parameters.Append(prmTableID)
			prmTableID.Value = CleanNumeric(Session("tableID"))

			prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
			cmdGetFindRecords.Parameters.Append(prmViewID)
			prmViewID.Value = CleanNumeric(Session("viewID"))

			prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
			cmdGetFindRecords.Parameters.Append(prmOrderID)
			prmOrderID.Value = CleanNumeric(Session("orderID"))

			prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
			cmdGetFindRecords.Parameters.Append(prmError)

			prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
			cmdGetFindRecords.Parameters.Append(prmReqRecs)
			prmReqRecs.Value = CleanNumeric(Session("FindRecords"))

			prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
			cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

			prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
			cmdGetFindRecords.Parameters.Append(prmIsLastPage)

			prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 2147483646)
			cmdGetFindRecords.Parameters.Append(prmLocateValue)
			prmLocateValue.Value = Session("locateValue")

			prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
			cmdGetFindRecords.Parameters.Append(prmColumnType)

			prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 100)
			cmdGetFindRecords.Parameters.Append(prmAction)
			prmAction.Value = Session("pageAction")

			prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
			cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

			prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
			cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
			prmFirstRecPos.Value = CleanNumeric(Session("firstRecPos"))

			prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
			cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
			prmCurrentRecCount.Value = CleanNumeric(Session("currentRecCount"))

			prmExcludedIDs = cmdGetFindRecords.CreateParameter("excludedIDs", 200, 1, 2147483646)	' 200=varchar, 1=input, 8000=size
			cmdGetFindRecords.Parameters.Append(prmExcludedIDs)
			prmExcludedIDs.Value = Session("selectedIDs1")
		
			prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
			cmdGetFindRecords.Parameters.Append(prmColumnSize)

			prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
			cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

			rstFindRecords = cmdGetFindRecords.Execute
	
			If (Err.Number <> 0) Then
				sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				If rstFindRecords.State = 1 Then
					iCount = 0
					Do While Not rstFindRecords.EOF
						sAddString = ""
					
						For iloop = 0 To (rstFindRecords.Fields.Count - 1)
							If iloop > 0 Then
								sAddString = sAddString & "	"
							End If
							
							If iCount = 0 Then
								sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
								Response.Write("<input type='hidden' id=txtColDef_" & iloop & " name=txtColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
							End If
							
							If rstFindRecords.Fields(iloop).Type = 135 Then
								' Field is a date so format as such.
								sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
							ElseIf rstFindRecords.Fields(iloop).Type = 131 Then
								' Field is a numeric so format as such.
								If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
									If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
										sTemp = ""
										sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).NumericScale, True, False, True)
									Else
										sTemp = ""
										sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).NumericScale, True, False, False)
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

						Response.Write("<input type='hidden' id=txtData_" & iCount & " name=txtData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
						iCount = iCount + 1
						rstFindRecords.MoveNext()
					Loop
	
					' Release the ADO recordset object.
					rstFindRecords.Close()
				End If
			End If
			rstFindRecords = Nothing

			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 

			Response.Write("<input type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
			Response.Write("<input type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

			cmdGetFindRecords = Nothing
			
		End If

		Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
	%>
</form>

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
	picklistSelectionData_window_onload();
</script>

