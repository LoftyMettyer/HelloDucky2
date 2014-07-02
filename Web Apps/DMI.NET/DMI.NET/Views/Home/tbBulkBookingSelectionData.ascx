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

			//need this as this grid won't accept live changes :/		
			$(grdLinkFind).jqGrid('GridUnload');

			var dataCollection = frmData.elements;
			var sControlName;
			var sColumnName;
			var iColumnType;
			var iCount;
			var colMode = [];
			var colNames = [];

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

							colNames.push(sColumnName);							

							if (sColumnName == "ID") {
								colMode.push({ name: sColumnName, hidden: true });
							}
							else {
								switch (sColumnType) {
									case "11":	//Boolean
										colMode.push({ name: sColumnName, edittype: "checkbox", formatter: 'checkbox', formatoptions: { disabled: true }, align: 'center', width: 100 });
										break;
									case "131":	//integer
										colMode.push({ name: sColumnName, edittype: "numeric", sorttype: 'integer', formatter: 'numeric', formatoptions: { disabled: true }, align: 'right', width: 100 });
										break;
									case "3":	//numeric
										colMode.push({ name: sColumnName, edittype: "numeric", sorttype: 'integer', formatter: 'numeric', formatoptions: { disabled: true }, align: 'right', width: 100 });
										break;
									case "135": //Date
										colMode.push({ name: sColumnName, edittype: "date", sorttype: 'date', formatter: 'date', formatoptions: { srcformat: dateFormat, newformat: dateFormat, disabled: true }, align: 'left', width: 100 });
										break;
									default:	//text
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
			if (dataCollection != null) {
				var colData = [];

				for (var i = 0; i < dataCollection.length; i++) {
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

				//create the column layout:
				var shrinkToFit = false;
				if (colMode.length < 8) shrinkToFit = true;

				$("#ssOleDBGridSelRecords").jqGrid({
					data: colData,
					datatype: "local",
					colNames: colNames,
					colModel: colMode,
					height: 400,
					rowNum: 1000,
					multiselect: true,
					autowidth: true,
					shrinktofit: shrinkToFit,
					beforeSelectRow: handleMultiSelect, // handle multi select
					onSelectRow: function () {
						ssOleDBGridSelRecords_rowcolchange();
					},
					rowNum: 500,
					pager: $('#ssOLEDBPager'),
					ondblClickRow: function () {
						ssOleDBGridSelRecords_dblClick();
					}
				}).jqGrid('hideCol', 'cb');


				$("#ssOleDBGridSelRecords").jqGrid('bindKeys', {
					"onEnter": function () {
						ssOleDBGridSelRecords_dblClick();
					}
				});

				//resize the grid to the height of its container.
				$("#ssOleDBGridSelRecords").jqGrid('setGridHeight', $("#findGridRow").height());

			}

			frmData.txtRecordCount.value = iCount;

			tbrefreshControls();
		}
	}


	// handle jqGrid multiselect => thanks to solution from Byron Cobb on http://goo.gl/UvGku
	var handleMultiSelect = function (rowid, e) {
		var grid = $(this);
		if (!e.ctrlKey && !e.shiftKey) {
			grid.jqGrid('resetSelection');
		}
		else if (e.shiftKey) {
			var initialRowSelect = grid.jqGrid('getGridParam', 'selrow');
			grid.jqGrid('resetSelection');

			var CurrentSelectIndex = grid.jqGrid('getInd', rowid);
			var InitialSelectIndex = grid.jqGrid('getInd', initialRowSelect);
			var startID = "";
			var endID = "";
			if (CurrentSelectIndex > InitialSelectIndex) {
				startID = initialRowSelect;
				endID = rowid;
			}
			else {
				startID = rowid;
				endID = initialRowSelect;
			}
			var shouldSelectRow = false;

			$.each(grid.getDataIDs(), function (_, id) {
				if ((shouldSelectRow = id == startID || shouldSelectRow)) {
					grid.jqGrid('setSelection', id, false);
				}
				return id != endID;

			});

			//last selected row too
			grid.jqGrid('setSelection', endID, false);

		}
		return true;
	};
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
