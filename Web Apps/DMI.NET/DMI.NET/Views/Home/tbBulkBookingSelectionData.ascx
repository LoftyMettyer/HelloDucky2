<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
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

<FORM id="frmDataUseful" name="frmDataUseful">
	<INPUT type="hidden" id="txtLoading" name="txtLoading" value="<%=session("tbSelectionDataLoading")%>">
</FORM>

<FORM id=frmData name=frmData>
<%

	Dim sErrorDescription = ""

	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
	
	Response.Write("<input type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

	' Get the required record count if we have a query.
	if session("tbSelectionDataLoading") = false then

		Dim sThousandColumns = ""
			
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
		
			Dim dsFindData = objDataAccess.GetDataSet("sp_ASRIntGetLinkFindRecords" _
				, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))} _
				, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))} _
				, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("orderID"))} _
				, prmError _
				, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = CleanNumeric(Session("FindRecords"))} _
				, prmIsFirstPage _
				, prmIsLastPage _
				, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("locateValue")} _
				, prmColumnType _
				, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("pageAction")} _
				, prmTotalRecCount _
				, prmFirstRecPos _
				, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("currentRecCount"))} _
				, New SqlParameter("psExcludedIDs", SqlDbType.VarChar, -1) With {.Value = ""} _
				, prmColumnSize _
				, prmColumnDecimals)

			Dim iCount = 0
			Dim sColDef = ""
			Dim sTemp = ""

			If dsFindData.Tables.Count > 0 Then
			
				Dim rstFindRecords = dsFindData.Tables(0)
			
				For Each objRow As DataRow In rstFindRecords.Rows

					Dim sAddString = ""
					
					For iloop = 0 To (rstFindRecords.Columns.Count - 1)
						If iloop > 0 Then
							sAddString = sAddString & "	"
						End If
							
						If iCount = 0 Then
							sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.Name
							Response.Write("<INPUT type='hidden' id=txtColDef_" & iloop & " name=txtColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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
			
			If prmError.Value <> 0 Then
				Session("ErrorTitle") = "Bulk Booking Selection Find Page"
				Session("ErrorText") = "Error reading employee records definition."
				Response.Clear()
				Response.Redirect("error")
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

	'	Response.Write "<INPUT type='hidden' id=txtAction name=txtAction value=" & session("Action") & ">" & vbcrlf
	'	Response.Write "<INPUT type='hidden' id=txtTableID name=txtTableID value=" & session("TableID") & ">" & vbcrlf
	'	Response.Write "<INPUT type='hidden' id=txtViewID name=txtViewID value=" & session("ViewID") & ">" & vbcrlf
	'	Response.Write "<INPUT type='hidden' id=txtOrderID name=txtOrderID value=" & session("OrderID") & ">" & vbcrlf
	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
%>
</FORM>
</div>

<script type="text/javascript"> tbBulkBookingSelectionData_onload();</script>
