<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">

	function picklistSelectionData_window_onload() {
		$("#picklistdataframe").attr("data-framesource", "PICKLISTSELECTIONDATA");
		
		$(".popup").dialog({
			resizable: false			
		});

		$(".popup").css('height', '540px');
		$(".popup").dialog('option', 'position', 'center');

		$("#reportframe").show();

		if (frmSelectDataUseful.txtLoading.value == "True") {
			loadAddRecords();
			return;
		}		
		var frmPicklistData = window.frmPicklistData.children;
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

			$("#ssOleDBGridSelRecords").jqGrid('GridUnload');			
			var dataCollection = frmPicklistData;
			var sControlName;
			var sColumnName;
			var iCount;
			var colData;
			var dateFormat = OpenHR.getLocaleDateString();

			colMode = [];
			colNames = [];
			
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
							var aColumnType = sColDef.split('\t');
							sColumnName = aColumnType[0];
							var sColumnType = aColumnType[1].replace('System.', "").toLowerCase();

							colNames.push(sColumnName);

							if (sColumnName == "ID") {
								colMode.push({ name: sColumnName, hidden: true });
							} else {
								switch (sColumnType) {
									case "boolean": // "11":
										colMode.push({ name: sColumnName, edittype: "checkbox", formatter: 'checkbox', formatoptions: { disabled: true }, align: 'center', width: 100 });
										break;
									case "decimal":
										var numDecimals = Number(aColumnType[2]);
										var sThousandSeparator = (aColumnType[3] === 'true') ? OpenHR.LocaleThousandSeparator() : "";
										colMode.push({ name: sColumnName, edittype: "numeric", sorttype: 'integer', formatter: 'number', formatoptions: { thousandsSeparator: sThousandSeparator, decimalSeparator: OpenHR.LocaleDecimalSeparator(), decimalPlaces: numDecimals, disabled: true }, align: 'right', width: 100 });
										break;
									case "datetime": //Date - 135
										colMode.push({ name: sColumnName, edittype: "date", sorttype: 'date', formatter: 'date', formatoptions: { srcformat: dateFormat, newformat: dateFormat, disabled: true }, align: 'left', width: 100 });
										break;
									default:
										colMode.push({ name: sColumnName, width: 100 });
								}
							}
						}
					}
				}
			}

			// Add the grid records.
			var sAddString;
			iCount = 0;
			var fRecordAdded;
			if (dataCollection != null) {
				colData = [];
				for (i = 0; i < dataCollection.length; i++) {
					sControlName = dataCollection.item(i).name;
					sControlName = sControlName.substr(0, 8);
					if (sControlName == "txtData_") {
						colDataArray = dataCollection.item(i).value.split("\t");
						obj = {};
						for (iCount2 = 0; iCount2 < colNames.length; iCount2++) {
							//loop through columns and add each one to the 'obj' object
							obj[colNames[iCount2]] = colDataArray[iCount2];
						}
						//add the 'obj' object to the 'colData' array
						colData.push(obj);

						fRecordAdded = true;
						iCount = iCount + 1;
					}
				}
			}

			//create the column layout:
			var shrinkToFit = false;
			if (colMode.length < 8) shrinkToFit = true;

			$("#ssOleDBGridSelRecords").jqGrid({
				multiselect: true,
				data: colData,
				datatype: 'local',
				colNames: colNames,
				colModel: colMode,
				rowNum: 1000,
				autowidth: true,
				shrinkToFit: shrinkToFit,
				onSelectRow: function () {
					$('#cmdSelectFilter').button('enable');
				},
				ondblClickRow: function (rowID) {
					makeSelection();
				},
				editurl: 'clientArray',
				afterShowForm: function ($form) {
					$("#dData", $form.parent()).click();
				},
				beforeSelectRow: handleMultiSelect // handle multi select
			}).jqGrid('hideCol', 'cb');

			//resize the grid to the height of its container.		
			var workPageHeight = $('.optiondatagridpage').outerHeight(true);
			var pageTitleHeight = $('.optiondatagridpage .pageTitle').outerHeight(true);
			var dropdownHeight = $('.optiondatagridpage .nowrap').outerHeight(true);
			var footerheight = $('.optiondatagridpage footer').outerHeight(true);

			var newGridHeight = workPageHeight - pageTitleHeight - dropdownHeight - footerheight;

			$("#ssOleDBGridSelRecords").jqGrid('setGridHeight', newGridHeight);
			$("#ssOleDBGridSelRecords").jqGrid('setGridWidth', $("#ssOleDBGridSelRecordsDiv").width());

			// Select the top record.
			if (fRecordAdded == true) {
			//	$("#ssOleDBGridSelRecords").jqGrid('setSelection', 1);
			}

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
	<%=Html.AntiForgeryToken()%>
</form>

<form id="frmSelectDataUseful" name="frmSelectDataUseful">
	<input type='hidden' id="txtLoading" name="txtLoading" value='<%=session("picklistSelectionDataLoading")%>'>
</form>

<div id="frmPicklistData">
	<%
		Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
		Dim sErrorDescription As String = ""
		Dim sThousandColumns As String = ""
		Dim sBlankIfZeroColumns As String = ""

		Response.Write("<input type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

		' Get the required record count if we have a query.
		If Session("picklistSelectionDataLoading") = False Then
			
			Try
				Get1000SeparatorBlankIfZeroFindColumns(CleanNumeric(Session("tableID")), CleanNumeric(Session("viewID")), CleanNumeric(Session("orderID")), sThousandColumns, sBlankIfZeroColumns)

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
					Dim jqGridColDef = New Dictionary(Of String, String)
					Dim rows As New List(Of Dictionary(Of String, Object))()
					Dim row As Dictionary(Of String, Object)
					Dim iLoop As Integer = 0
					
					Dim rstFindRecords = dsData.Tables(0)
					For Each dr As DataRow In rstFindRecords.Rows
						iLoop += 1
						row = New Dictionary(Of String, Object)()
						For Each col As DataColumn In rstFindRecords.Columns
							If Not jqGridColDef.ContainsKey(col.ColumnName) Then
								Dim sColDef As String = col.DataType.Name

								If sColDef = "Decimal" Then
									Dim numberAsString As String = dr(col).ToString()
									Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(LocaleDecimalSeparator(), System.StringComparison.Ordinal)
									Dim numberOfDecimals As Integer = 0
									If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length

									If Mid(sThousandColumns, iLoop + 1, 1) = "1" Then
										sColDef &= vbTab & numberOfDecimals.ToString() & vbTab & "true"
									Else
										sColDef &= vbTab & numberOfDecimals.ToString() & vbTab & "false"
									End If
								End If

								jqGridColDef.Add(col.ColumnName, sColDef)
							End If

							If col.DataType.Name = "DateTime" And dr(col).ToString().Length > 0 Then
								row.Add(col.ColumnName, dr(col).ToShortDateString())
							Else
								row.Add(col.ColumnName, dr(col))
							End If

						Next
						rows.Add(row)
					Next

					'Now that we have the data, output it to input tags!
					Dim counter As Integer = 0
					Dim addString As String = ""
					'Column definitions for jqGrid's colModel
					For Each key As String In jqGridColDef.Keys
						counter += 1
						Response.Write("<input type='hidden' id='txtColDef_" & counter & "' name='txtColDef_" & counter & "' value=""" & key & vbTab & jqGridColDef(key) & """>" & vbCrLf)
					Next
					'Data for jqGrid's colData
					counter = 0
					For i As Integer = 0 To rows.Count - 1
						addString = ""
						For Each key As String In rows(i).Keys
							addString &= rows(i)(key).ToString & vbTab
						Next
						counter += 1
						Response.Write("<input type='hidden' id='txtData_" & counter & "' name='txtData_" & counter & "' value=""" & addString & """>" & vbCrLf)
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
</div>

<script type="text/javascript">
	picklistSelectionData_window_onload();
</script>

