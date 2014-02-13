<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

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
		Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

		
		Dim sErrorDescription As String = ""
		Dim sThousandColumns As String		
		Dim iCount As Integer
		Dim sAddString As String
		Dim sColDef As String
		Dim sTemp As String
			
		Response.Write("<input type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

		' Get the required record count if we have a query.
		If Session("picklistSelectionDataLoading") = False Then
			
			Try
				sThousandColumns = Get1000SeparatorFindColumns(CleanNumeric(Session("tableID")), CleanNumeric(Session("viewID")), CleanNumeric(Session("orderID")))

				Dim prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("firstRecPos"))}
				Dim prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		
				Dim dsData = objDataAccess.GetDataSet("sp_ASRIntGetLinkFindRecords" _
					, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))} _
					, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))} _
					, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("orderID"))} _
					, prmError _
					, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = 1000000} _
					, prmIsFirstPage _
					, prmIsLastPage _
					, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("locateValue")} _
					, prmColumnType _
					, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("pageAction")} _
					, prmTotalRecCount _
					, prmFirstRecPos _
					, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("currentRecCount"))} _
					, New SqlParameter("psExcludedIDs", SqlDbType.VarChar, -1) With {.Value = Session("selectedIDs1")} _
					, prmColumnSize _
					, prmColumnDecimals)

				If dsData.Tables.Count > 0 Then

					Dim rstFindRecords = dsData.Tables(0)
					For Each objRow As DataRow In rstFindRecords.Rows

						sAddString = ""
					
						For iloop = 0 To (rstFindRecords.Columns.Count - 1)
							If iloop > 0 Then
								sAddString = sAddString & "	"
							End If
							
							If iCount = 0 Then
								sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString()
								Response.Write("<input type='hidden' id=txtColDef_" & iloop & " name=txtColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
							End If
							
							If rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.datetime" Then
								' Field is a date so format as such.
								sAddString = sAddString & ConvertSQLDateToLocale(objRow(iloop))
							ElseIf rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.decimal" Then
								' Field is a numeric so format as such.
								If Not IsDBNull(objRow(iloop)) Then
									If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
										sTemp = FormatNumber(objRow(iloop), , True, False, True)
									Else
										sTemp = FormatNumber(objRow(iloop), , True, False, False)
									End If
									sTemp = Replace(sTemp, ".", "x")
									sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
									sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
									sAddString = sAddString & sTemp
								End If
							Else
								If Not IsDBNull(objRow(iloop)) Then
									sAddString = sAddString & Replace(objRow(iloop).ToString(), """", "&quot;")
								End If
							End If
						Next

						Response.Write("<input type='hidden' id=txtData_" & iCount & " name=txtData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
						iCount += 1
					Next
				End If

				Response.Write("<input type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & prmIsFirstPage.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & prmIsLastPage.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & prmColumnType.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & prmTotalRecCount.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & prmFirstRecPos.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & prmColumnSize.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & prmColumnDecimals.Value & ">" & vbCrLf)
			
			
				
			Catch ex As Exception
				sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
			End Try

		End If

		Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
	%>
</form>

<script type="text/javascript">
	picklistSelectionData_window_onload();
</script>

