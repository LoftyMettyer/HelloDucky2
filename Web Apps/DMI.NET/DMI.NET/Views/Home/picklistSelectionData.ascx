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
	
			var dataCollection = frmPicklistData.elements;
			var sControlName;
			var sColumnName;
			var iColumnType;
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
							sColumnName = sColDef.substr(0, iIndex);
							sColumnType = sColDef.substr(iIndex + 1).replace('System.', "").toLowerCase();
							colNames.push(sColumnName);

							if (sColumnName == "ID") {
								colMode.push({ name: sColumnName, hidden: true });
							} else {
								switch (sColumnType) {
									case "boolean": // "11":
										colMode.push({ name: sColumnName, edittype: "checkbox", formatter: 'checkbox', formatoptions: { disabled: true }, align: 'center', width: 100 });
										break;
									case "decimal":
										colMode.push({ name: sColumnName, edittype: "numeric", sorttype: 'integer', formatter: 'numeric', formatoptions: { disabled: true }, align: 'right', width: 100 });
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

