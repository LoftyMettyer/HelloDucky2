<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET.Code" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<%
	If Session("SSIMode") = False Then
		
		Response.Write("<script src=""" & Url.LatestContent("~/bundles/jQuery") & """ type=""text/javascript""></script>")
		Response.Write("<script src=""" & Url.LatestContent("~/bundles/jQueryUI7") & """ type=""text/javascript""></script>")
		Response.Write("<script src=""" & Url.LatestContent("~/bundles/OpenHR_General") & """ type=""text/javascript""></script>")
		
	End If	
	%>

<%
	'Data access variables
	Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
	Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
	Dim SPParameters() As SqlParameter
	Dim resultDataSet As DataSet
	Dim rstFindDefinition As DataTable
	Dim rstFindRecords As DataTable
	Dim rstOriginalColumns As DataTable
	Dim resultsDataTable As DataTable
%>

<%--Base stylesheets--%>

<%If Session("SSIMode") = False Then
		Response.Write("<link href=""" & Url.LatestContent("~/Content/font-awesome.min.css") & """ rel=""stylesheet"" type=""text/css"" />")
		Response.Write("<link href=""" & Url.LatestContent("~/Content/Site.css") & """ rel=""stylesheet"" type=""text/css"" />")
		Response.Write("<link href=""" & Url.LatestContent("~/Content/OpenHR.css") & """ rel=""stylesheet"" type=""text/css"" />")
		Response.Write("<link href=""" & Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css") & """ rel=""stylesheet"" type=""text/css"" />")
		Response.Write("<link href=""" & Url.LatestContent("~/Content/ui.jqgrid.css") & """ rel=""stylesheet"" type=""text/css"" />")
		Response.Write("<link href=""" & Url.LatestContent("~/Content/table.css") & """ rel=""stylesheet"" type=""text/css"" />")
		
	End If%>

<script src="<%: Url.LatestContent("~/bundles/recordedit")%>" type="text/javascript"></script>

<script type="text/javascript">
	$(document).ready(function () {
		if ('<%=session("linktype")%>' == 'multifind') {
			
			//for multifind (SSI views) show relevant buttons with applicable functions
			menu_setVisibletoolbarGroupById("mnuSectionRecordFindEdit", false);
			menu_setVisibleMenuItem("mnutoolAccessLinksFind", true);
			menu_setVisibleMenuItem("mnutoolCancelLinksFind", false);

			if (menu_isSSIMode()) menu_setVisibleMenuItem('mnutoolFixedWorkflowOutOfOffice', "<%:ViewData("showOutOfOffice")%>");

			setTimeout('gridBindKeys(true)', 300);


		} else {
			menu_setVisibletoolbarGroupById("mnuSectionRecordFindEdit", true);
			menu_setVisibleMenuItem("mnutoolAccessLinksFind", false);
			menu_setVisibleMenuItem("mnutoolCancelLinksFind", false);

			setTimeout('gridBindKeys(false)', 300);

			//Resize functionality
			window.top.$('#' + OpenHR.activeWindowID()).on("dialogresizestop", function(event, ui) {
				resizeFindGrid();
			});


		}
	});

	function doEdit() {
		var sRecordID = selectedRecordID();

		if ("<%=session("linkType")%>" == "multifind") {
			var sParams = "<%=session("TableID")%>!<%=session("ViewID")%>_";
			sParams = sParams.concat(sRecordID);
			loadPartialView('linksMain', 'Home', 'workframe', sParams);
		}
	}

	function gridBindKeys(multifind) {

		if (multifind) {
			$("#findGridTable").jqGrid("setGridParam", { ondblClickRow: function(rowID) { doEdit(); } });
			$('#findGridTable').jqGrid('bindKeys', { "onEnter": function() { doEdit(); } });
		} else {
			$("#findGridTable").jqGrid("setGridParam", {
				ondblClickRow: function (rowID) {
					if (!IsMultiSelectionModeOn()) {
						var thisWindow = OpenHR.activeWindowID();
						window.top.menu_editRecord();
					}
				}
			});
			$('#findGridTable').jqGrid('bindKeys', {
				"onEnter": function(rowid) {
					//If we are in "Inline-edit" mode and "Enter" was pressed then don't go into "Edit record" mode
					if ($("#findGridTable_iledit").hasClass('ui-state-disabled')) {
						return;
					}

					window.top.menu_editRecord();
				}
			});

			// Refresh find grid
			RefreshFindGrid(!multifind, !thereIsAtLeastOneEditableColumn);
		}
	}

	/******* Begin Changes for the user story 19436: As a user, I want to run reports and utilities from the Find Window  *********/

	//Refresh find grid
	function RefreshFindGrid(isNonMultiFindLinkType, isNonEditableGrid) {

		// Refresh the find window ribbon buttons
		RefreshFindWindowRibbonButtons(isNonMultiFindLinkType, isNonEditableGrid);

		// Bind events for multi select grid
		BindEventsForMultiSelectFindGrid(isNonMultiFindLinkType);

		// Refresh find grid toolbar
		RefreshFindGridToolbar();
	}

	// Binds the events for the multi select find grid
	function BindEventsForMultiSelectFindGrid(isNonMultiFindLinkType) {

		var grid = $("#findGridTable");

		if (IsMultiSelectionModeOn() && isNonMultiFindLinkType && grid.getGridParam("reccount") > 0) {

			//Reset row selection (E.g. after applied filter, or turn multiselect on for editable grid). 
			//If not doing so, the first record will come as selected because of previously bind load complete (movefirst).
			grid.jqGrid('resetSelection');
			$("#mnutoolPositionRecordFind span.selectedRecordsCount").html("Selected : 0");

			// Bind the grid events
			grid.jqGrid("setGridParam", {
				onSelectRow: function (id) {
					var p = this.p, item = p.data[p._index[id]];
					if (typeof (item.cb) === 'undefined') { item.cb = true; } else { item.cb = !item.cb; }
					SetsSelectedRowsCount();
				},
				onSelectAll: function (ids, selected) {
					var p = this.p;
					$.each(ids, function (id, value) { p.data[p._index[value]].cb = selected; });
					SetsSelectedRowsCount();
				},
				loadComplete: function () {

					var p = this.p, data = p.data, item, index = p._index, rowid;
					for (rowid in index) {
						if (index.hasOwnProperty(rowid)) {
							item = data[index[rowid]];
							if (typeof (item.cb) === 'boolean' && item.cb) { $(this).jqGrid('setSelection', rowid, false); }
						}
					}
					SetsSelectedRowsCount();
				}
			});

			// If filter is applied, then sets the previously selected records if they match the filterd criteria
			SelectRecordsWhenFilterApplied();
		}
		else {
			grid.setGridParam({ multiselect: false }).hideCol('cb');
		}
	}

	// Sets count for the selected number of rows when multi selection of grid rows allowed.
	function SetsSelectedRowsCount() {

		var selectedRecords = GetMultiSelectRecordIDs();

		//Sets the selected ids as string. This will be used when user does apply the filter.
		$("#txtSelectedRecordsInFindGrid")[0].value = selectedRecords;
		$("#mnutoolPositionRecordFind span.selectedRecordsCount").html("Selected : " + selectedRecords.length);
	}

	// Sets the previous selection if exist after applying filter
	function SelectRecordsWhenFilterApplied() {

		var count = 0;
		var previouslySelectedRecordIds = "<%=Session("OptionSelectedRecordIds")%>";

		if (IsMultiSelectionModeOn() && previouslySelectedRecordIds != "") {
			var selectedRecords = [];
			var p = $("#findGridTable")[0].p;
			var item;
			$.each(previouslySelectedRecordIds.split(','), function (id, value) {
				item = p.data[p._index[value]];
				if (typeof (item) != typeof (undefined)) {
					item.cb = true;
					selectedRecords.push(item.ID);
					$("#findGridTable").jqGrid('setSelection', value, false);
					count++;
				}
			});

			//Sets the selected ids as string. This will be used when user does apply the filter.
			$("#txtSelectedRecordsInFindGrid")[0].value = selectedRecords;
			$("#mnutoolPositionRecordFind span.selectedRecordsCount").html("Selected : " + count);
		}
	}


	// Provide empty function. Called when clicking close from the report/utility whilst loaded from find window.
	function find_refreshData() { }

	/******* End Changes for the user story 19436: As a user, I want to run reports and utilities from the Find Window  *********/

</script>

<div id="divFindForm" <%=session("BodyTag")%>>
	<form action="" class="absolutefull" method="POST" id="frmFindForm" name="frmFindForm">
		<div class="absolutefull">
			<div id="row1" style="margin-left: 20px; margin-right: 20px">
				<%
					Dim sErrorDescription As String = ""
					Dim sThousSepSummaryFields As String
					
					If ViewBag.pageTitle.ToString().Length = 0 Then
						' DMI View.
						' Display the appropriate page title.
						Dim prm_psTitle As New SqlParameter("@psTitle", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = 500}
						Dim prm_pfQuickEntry As New SqlParameter("@pfQuickEntry ", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						SPParameters = New SqlParameter() { _
								prm_psTitle, _
								prm_pfQuickEntry,
								New SqlParameter("@plngScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))}, _
								New SqlParameter("@plngViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))}
						}

						Try
							objDataAccess.ExecuteSP("sp_ASRIntGetFindWindowInfo", SPParameters)
						Catch ex As Exception
							sErrorDescription = "The page title could not be created." & vbCrLf & FormatError(ex.Message)
						End Try
						
						If Len(sErrorDescription) = 0 Then
							Dim homelinkURL = "javascript:loadPartialView(""linksMain"", ""Home"", ""workframe"", null);"
							Response.Write(String.Format("<div class='pageTitleDiv'><a onclick='{0}' title='Back'><i class='pageTitleIcon icon-circle-arrow-left'></i></a><span class='pageTitle'>" & _
											Replace(prm_psTitle.Value.ToString, "_", " ") & "</span>" & vbCrLf, homelinkURL))
							response.write("<label id='txtRIE' style='float: right;'></label>")
							Response.Write("<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & prm_pfQuickEntry.Value.ToString & "></div>" & vbCrLf)
						End If
					Else
						' SSI View.
						Dim homelinkURL = "javascript:loadPartialView(""linksMain"", ""Home"", ""workframe"", null);"
						Response.Write(String.Format("<div class='pageTitleDiv'><a onclick='{0}' title='Back'><i class='pageTitleIcon icon-circle-arrow-left'></i></a><span class='pageTitle'>" & _
								ViewBag.pageTitle & "</span>" & vbCrLf, homelinkURL))
						Response.Write("<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & ViewBag.pageTitle & "></div>" & vbCrLf)
					End If
				%>
			</div>
			<div id="findGridRow" style="margin-right: 20px; margin-left: 20px;">
				<%
					Dim sTemp As String
					Dim sThousandColumns As String = ""
					Dim sBlankIfZeroColumns As String = ""
					Dim TableOrViewName As String = ""
					Dim sColDef As String
					Dim iCount As Integer
					Dim sAddString As String
								
					Const iNumberOfRecords As Integer = 100000
					
					Dim fCancelDateColumn = True
									
					If (Len(sErrorDescription) = 0) And (Session("TB_CourseTableID") > 0) And Len(NullSafeString(Session("lineage"))) > 0 Then
						
						Dim sSubString As String = Session("lineage").ToString()
						Dim iIndex = InStr(sSubString, "_")
						sSubString = Mid(sSubString, iIndex + 1)
						iIndex = InStr(sSubString, "_")
						sSubString = Mid(sSubString, iIndex + 1)
						iIndex = InStr(sSubString, "_")
						sSubString = Mid(sSubString, iIndex + 1)
						iIndex = InStr(sSubString, "_")
						sSubString = Mid(sSubString, iIndex + 1)
						iIndex = InStr(sSubString, "_")
						Dim lngRecordID = Left(sSubString, iIndex - 1)

						' Get the Course Date
						Dim prmError As New SqlParameter("@pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						Dim prmCancelDateColumn As New SqlParameter("@pfCancelDate", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			
						SPParameters = New SqlParameter() { _
									prmError, _
									New SqlParameter("@piRecID", SqlDbType.Int) With {.Value = CleanNumeric(Session("CleanNumeric(lngRecordID)"))}, _
									prmCancelDateColumn _
						}
						Try
							objDataAccess.ExecuteSP("spASRIntGetCancelCourseDate", SPParameters)
						Catch ex As Exception
							sErrorDescription = "Unable to check for a Cancelled Course Date." & vbCrLf & FormatError(ex.Message)
						End Try

						If Len(sErrorDescription) = 0 Then
							fCancelDateColumn = prmCancelDateColumn.Value
						End If
					End If

					If Len(sErrorDescription) = 0 Then
						' Get the find records.
						Dim prmError As New SqlParameter("@pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						Dim prmSomeSelectable As New SqlParameter("@pfSomeSelectable", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						Dim prmSomeNotSelectable As New SqlParameter("@pfSomeNotSelectable", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						Dim prmRealSource As New SqlParameter("@psRealSource", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmInsertGranted As New SqlParameter("@pfInsertGranted", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						Dim prmDeleteGranted As New SqlParameter("@pfDeleteGranted", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						Dim prmIsFirstPage As New SqlParameter("@pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						Dim prmIsLastPage As New SqlParameter("@pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						Dim prmColumnType As New SqlParameter("@piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						Dim prmColumnSize As New SqlParameter("@piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						Dim prmColumnDecimals As New SqlParameter("@piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						Dim prmTotalRecCount As New SqlParameter("@piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						Dim prmFirstRecPos As New SqlParameter("@piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("firstRecPos"))}
						
						Dim filterDefForCurrentTable As String = IIf(IsNothing(Session("filterDef_" & Session("tableID"))), "", Session("filterDef_" & Session("tableID")))

						SPParameters = New SqlParameter() { _
								prmError, _
								prmSomeSelectable, _
								prmSomeNotSelectable, _
								prmRealSource, _
								prmInsertGranted, _
								prmDeleteGranted, _
								New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))}, _
								New SqlParameter("@piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))}, _
								New SqlParameter("@piOrderID ", SqlDbType.Int) With {.Value = CleanNumeric(Session("orderID"))}, _
								New SqlParameter("@piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))}, _
								New SqlParameter("@piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))}, _
								New SqlParameter("@psFilterDef", SqlDbType.VarChar, -1) With {.Value = filterDefForCurrentTable}, _
								New SqlParameter("@piRecordsRequired", SqlDbType.Int) With {.Value = iNumberOfRecords}, _
								prmIsFirstPage, _
								prmIsLastPage, _
								New SqlParameter("@psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("locateValue")}, _
								prmColumnType, _
								prmColumnSize, _
								prmColumnDecimals, _
								New SqlParameter("@psAction", SqlDbType.VarChar) With {.Value = Session("action"), .Size = 255}, _
								prmTotalRecCount, _
								prmFirstRecPos, _
								New SqlParameter("@piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("currentRecCount"))}, _
								New SqlParameter("@psDecimalSeparator", SqlDbType.VarChar, 255) With {.Value = Session("LocaleDecimalSeparator")}, _
								New SqlParameter("@psLocaleDateFormat", SqlDbType.VarChar, 255) With {.Value = Platform.LocaleDateFormatForSQL()}, _
								New SqlParameter("@RecordID", SqlDbType.Int) With {.Value = -1} _
						}
						'Parameter @RecordID = -1 above means "Return all records"

						Try
							resultDataSet = objDataAccess.GetDataSet("spASRIntGetFindRecords", SPParameters)

							Dim clientArrayData As New ArrayList
							Dim columnsDefaultValues As String = "var columnsDefaultValues = {"	'Save the default values for the columns in an array that we can use client side
							
							If prmSomeSelectable.Value = 0 Then
								sErrorDescription = "You do not have permission to read any of the selected order's find columns."
							Else
								' Get the recordset parameters
								sThousandColumns = resultDataSet.Tables(0).Rows(0)("ThousandColumns").ToString()
								sBlankIfZeroColumns = resultDataSet.Tables(0).Rows(0)("BlankIfZeroColumns").ToString()
								TableOrViewName = String.Concat("var currentTableOrViewName = """, resultDataSet.Tables(0).Rows(0)("TableOrViewName").ToString(), """;", vbCrLf)

								rstFindRecords = resultDataSet.Tables(1) 'Get the actual data
								rstFindDefinition = resultDataSet.Tables(2)	'Get the columns information
								rstOriginalColumns = resultDataSet.Tables(3) 'Get the original columns information
								
								'We need this to fettle the columns information table
								rstFindDefinition.Columns.Add("columnNameOriginal", GetType(String))
								
								For Each r As DataRow In rstFindDefinition.Rows
									r("columnNameOriginal") = r("columnName")
								Next
								
								Dim lastColumnName As String = ""
								Dim ColumnNameCounter As Short = 1
								For Each r As DataRow In rstOriginalColumns.Rows
									If r("columnName") = lastColumnName Then
										lastColumnName = r("columnName")
										r("columnName") = r("columnName") & ColumnNameCounter
										ColumnNameCounter += 1

										'Rename the columns
										For Each r1 In rstFindDefinition.Select("columnID = " & r("ColumnID"))
											r1("columnName") = r("columnName")
										Next
									Else
										ColumnNameCounter = 1
										lastColumnName = r("columnName")
									End If
								Next

								' Instantiate and initialise the grid. 
								Response.Write("<table class='outline' style='width : 100%; ' id='findGridTable'>" & vbCrLf)
								Response.Write("<div id='pager-coldata'></div>" & vbCrLf)

								'Output the grid definition (i.e. columns)
								For iloop = 0 To (rstFindRecords.Columns.Count - 1)
									If (rstFindRecords.Columns(iloop).ColumnName = "ID" OrElse rstFindRecords.Columns(iloop).ColumnName = "Timestamp") Then
										sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString.Replace("System.", "")
										Response.Write(String.Format("<input type='hidden' id='txtFindColDef_{0}' name='txtFindColDef_{0}' value='{1}' data-colname='{2}' data-type='{3}'>" _
												 , iloop, sColDef, rstFindRecords.Columns(iloop).ColumnName, "integer", vbCrLf))
									Else
										Dim objRow = rstFindDefinition.Select("ColumnName='" & rstFindRecords.Columns(iloop).ColumnName & "'")
										sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString.Replace("System.", "")
										Response.Write(String.Format("<input type='hidden' id='txtFindColDef_{0}' name='txtFindColDef_{0}' value='{1}' data-colname='{2}' data-datatype='{3}' data-columnid='{4}' data-editable='{5}' data-controltype='{6}' data-size='{7}' data-decimals='{8}' data-lookuptableid='{9}' data-lookupcolumnid='{10}' data-spinnerminimum='{11}' data-spinnermaximum='{12}' data-spinnerincrement='{13}' data-lookupfiltercolumnid='{14}' data-lookupfiltervalueid='{15}' data-Mask='{16}' data-DefaultValueExprID='{17}' data-BlankIfZero='{18}'>" _
												 , iloop, sColDef, objRow.FirstOrDefault.Item("columnNameOriginal"), _
												 objRow.FirstOrDefault.Item("datatype"), _
												 objRow.FirstOrDefault.Item("columnID"), _
												 objRow.FirstOrDefault.Item("updateGranted"), _
												 objRow.FirstOrDefault.Item("controltype"), _
												 objRow.FirstOrDefault.Item("size"), _
												 objRow.FirstOrDefault.Item("decimals"), _
												 objRow.FirstOrDefault.Item("LookupTableID"), _
												 objRow.FirstOrDefault.Item("LookupColumnID"), _
												 objRow.FirstOrDefault.Item("SpinnerMinimum"), _
												 objRow.FirstOrDefault.Item("SpinnerMaximum"), _
												 objRow.FirstOrDefault.Item("SpinnerIncrement"), _
												 objRow.FirstOrDefault.Item("LookupFilterColumnID"), _
												 objRow.FirstOrDefault.Item("LookupFilterValueID"), _
												 objRow.FirstOrDefault.Item("Mask"), _
												 objRow.FirstOrDefault.Item("DefaultValueExprID"), _
												 IIf(objRow.FirstOrDefault.Item("BlankIfZero"), "1", "0") _
												 , vbCrLf))

										'Save the default value for this column in an array
										columnsDefaultValues = String.Concat(columnsDefaultValues, """", objRow.FirstOrDefault.Item("columnID"), """:""", EncodeStringToJavascriptSpecialCharacters(objRow.FirstOrDefault.Item("DefaultValue")), """,")
												
										'If column is a Lookup, we need to get its associated data
										If (objRow.FirstOrDefault.Item("datatype") = 12 Or objRow.FirstOrDefault.Item("datatype") = 2 Or objRow.FirstOrDefault.Item("datatype") = 4) And objRow.FirstOrDefault.Item("controltype") = 2 And objRow.FirstOrDefault.Item("LookupColumnID") <> 0 Then
											Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
											'Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

											Dim objTable = objDatabase.GetTableFromColumnID(objRow.FirstOrDefault.Item("LookupColumnID"))
											Dim fIsLookupTable = (objTable.TableType = TableTypes.tabLookup)

											Dim _prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
											Dim _prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
											Dim _prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
											Dim _prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
											Dim _prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
											Dim _prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("optionFirstRecPos"))}
											Dim _prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
											Dim _prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
											Dim _prmLookupColumnGridPosition = New SqlParameter("piLookupColumnGridNumber", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
													
											Dim rstLookup As New DataTable
													
											clientArrayData.Add(String.Concat("var isLookupTable_", objRow.FirstOrDefault.Item("columnID"), " = ", fIsLookupTable.ToString.ToLower, ";"))

											If Not fIsLookupTable Then
																										
												Dim iOrderID = objSession.Tables.Where(Function(m) m.ID = objRow.FirstOrDefault.Item("LookupTableID")).FirstOrDefault.DefaultOrderID
														
												rstLookup = objDataAccess.GetFromSP("spASRIntGetLookupFindRecords2" _
														, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = objRow.FirstOrDefault.Item("LookupTableID")} _
														, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = 0} _
														, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = iOrderID} _
														, New SqlParameter("piLookupColumnID", SqlDbType.Int) With {.Value = objRow.FirstOrDefault.Item("LookupColumnID")} _
														, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = 10000} _
														, _prmIsFirstPage _
														, _prmIsLastPage _
														, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = ""} _
														, _prmColumnType _
														, _prmColumnSize _
														, _prmColumnDecimals _
														, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = "LOAD"} _
														, _prmTotalRecCount _
														, _prmFirstRecPos _
														, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = 0} _
														, New SqlParameter("psFilterValue", SqlDbType.VarChar, -1) With {.Value = ""} _
														, New SqlParameter("piCallingColumnID", SqlDbType.Int) With {.Value = objRow.FirstOrDefault.Item("columnID")} _
														, _prmLookupColumnGridPosition _
														, New SqlParameter("pfOverrideFilter", SqlDbType.Bit) With {.Value = "False"})
											Else
												Dim _prmThousandColumns As New SqlParameter("@ps1000SeparatorCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
												Dim _prmBlankIfZeroColumns As New SqlParameter("@psBlanIfZeroCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
												Try
													objDataAccess.ExecuteSP("spASRIntGetLookupFindColumnInfo", _
																			New SqlParameter("@piLookupColumnID", SqlDbType.Int) With {.Value = objRow.FirstOrDefault.Item("LookupColumnID")}, _
																			_prmThousandColumns, _
																			_prmBlankIfZeroColumns
													)
												Catch ex As Exception
													sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
												End Try

												Dim sLookupThousandColumns = _prmThousandColumns.Value.ToString()
														
												rstLookup = objDataAccess.GetFromSP("spASRIntGetLookupFindRecords" _
													, New SqlParameter("piLookupColumnID", SqlDbType.Int) With {.Value = objRow.FirstOrDefault.Item("LookupColumnID")} _
													, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = 10000} _
													, _prmIsFirstPage _
													, _prmIsLastPage _
													, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = ""} _
													, _prmColumnType _
													, _prmColumnSize _
													, _prmColumnDecimals _
													, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = "LOAD"} _
													, _prmTotalRecCount _
													, _prmFirstRecPos _
													, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = 0} _
													, New SqlParameter("psFilterValue", SqlDbType.VarChar, -1) With {.Value = "True"} _
													, New SqlParameter("piCallingColumnID", SqlDbType.Int) With {.Value = objRow.FirstOrDefault.Item("columnID")} _
													, New SqlParameter("pfOverrideFilter", SqlDbType.Bit) With {.Value = "False"})
											End If

											'Place the Lookup Column Grid Position in a Javascript variable
											Dim strLookupColumnGridPosition As String = String.Concat("var LookupColumnGridPosition_", objRow.FirstOrDefault.Item("columnID"), " = ")
											If Not fIsLookupTable Then
												strLookupColumnGridPosition = String.Concat(strLookupColumnGridPosition, _prmLookupColumnGridPosition.Value, ";")
											Else
												strLookupColumnGridPosition = String.Concat(strLookupColumnGridPosition, "0;")
											End If
													
											clientArrayData.Add(strLookupColumnGridPosition)
										End If
												
										'Get the data for Option Groups or Dropdown Lists
										If ( _
													((objRow.FirstOrDefault.Item("datatype") = 12 And objRow.FirstOrDefault.Item("controltype") = 2) Or (objRow.FirstOrDefault.Item("datatype") = 12 And objRow.FirstOrDefault.Item("controltype") = 16)) _
													And (objRow.FirstOrDefault.Item("LookupColumnID") = 0)
												) Then

											Dim _prmColumnIDs = New SqlParameter("ColumnIDs", SqlDbType.NChar, 100) With {.Value = objRow.FirstOrDefault.Item("columnID")}
											Dim rstOptionGroupOrDropDown As DataTable = objDataAccess.GetFromSP("spASRIntGetColumnControlValues", _prmColumnIDs)
											Dim strOptionGroupOrDropDownData As String = String.Concat("var colOptionGroupOrDropDownData_", objRow.FirstOrDefault.Item("columnID"), " = [")
													
											For Each r As DataRow In rstOptionGroupOrDropDown.Rows
												strOptionGroupOrDropDownData = String.Concat(strOptionGroupOrDropDownData, "[")
												For Each c As DataColumn In rstOptionGroupOrDropDown.Columns
													If c.ColumnName.ToLower <> "columnid" Then
														strOptionGroupOrDropDownData = String.Concat(strOptionGroupOrDropDownData, """", EncodeStringToJavascriptSpecialCharacters(r(c).ToString), """,")
													End If
												Next
												strOptionGroupOrDropDownData = String.Concat(strOptionGroupOrDropDownData.TrimEnd(","), "],")
											Next
											strOptionGroupOrDropDownData = String.Concat(strOptionGroupOrDropDownData.TrimEnd(","), "];")
											clientArrayData.Add(strOptionGroupOrDropDownData & vbCrLf)
										End If
									End If
								Next

								'Output the grid data, if any
								iCount = 0
								For Each row As DataRow In rstFindRecords.Rows
									sAddString = ""
						
									For iloop = 0 To (rstFindRecords.Columns.Count - 1)
										If iloop > 0 Then
											sAddString &= "	"
										End If

										If rstFindRecords.Columns(iloop).DataType = GetType(System.DateTime) Then
											' Field is a date so format as such.
											sAddString = sAddString & ConvertSQLDateToLocale(row(iloop))
										ElseIf GeneralUtilities.IsDataColumnDecimal(rstFindRecords.Columns(iloop)) Then
											' Field is a numeric so format as such.
											If Not IsDBNull(row(iloop)) Then
												Dim dec As Decimal = row(iloop)
												
												If Mid(sBlankIfZeroColumns, iloop + 1, 1) = "1" Then
													' blank if zero
													If dec > 0 Then
														sAddString &= dec.ToString(Globalization.CultureInfo.InvariantCulture)
													Else
														sAddString &= ""
													End If
												Else
													sAddString &= dec.ToString(Globalization.CultureInfo.InvariantCulture)
												End If
												
											End If
										Else
											If Not IsDBNull(row(iloop)) Then
												sAddString = sAddString & Replace(row(iloop), """", "&quot;")
											End If
										End If
									Next

									Response.Write("<input type='hidden' id='txtAddString_" & iCount & "' name='txtAddString_" & iCount & "' value=""" & sAddString & """>" & vbCrLf)

									iCount += 1
								Next
							
							End If

							Response.Write("</table>")
				
							columnsDefaultValues = String.Concat(columnsDefaultValues.Trim(","), "};")
							clientArrayData.Add(vbCrLf & TableOrViewName & vbCrLf & columnsDefaultValues)

							'Output the client side array data
							Response.Write("<script type='text/javascript'>" & vbCrLf)
							For Each s As String In clientArrayData
								Response.Write(s & vbCrLf)
							Next

							'Can we add new records to this table/view?
							Response.Write(String.Concat(vbCrLf, "var insertGranted = ", prmInsertGranted.Value.ToString.ToLower, ";", vbCrLf))
							
							'We need TableID and OrderID in the client side
							Response.Write(String.Concat(vbCrLf, "var tableId = ", Session("tableID"), ";", vbCrLf))
							Response.Write(String.Concat(vbCrLf, "var orderId = ", Session("orderID"), ";", vbCrLf))
							Response.Write(String.Concat(vbCrLf, "window.top.rowWasModified = false;", vbCrLf))
							Response.Write(String.Concat(vbCrLf, "var linktype = '", Session("linktype"), "';", vbCr))

							Response.Write("</script>" & vbCrLf)
							
							'/******* Task 19439: Begin Changes for the user story 19436: As a user, I want to run reports and utilities from the Find Window *********/
								
							Dim rstDefSelRecords As DataTable
							Dim baseTableId = CleanNumeric(Session("tableID"))
							
							' Sets the txtCustomReportGrantedForFindWindow value
							Dim prmType = New SqlParameter("intType", SqlDbType.Int) With {.Direction = ParameterDirection.Input, .Value = 2}
							Dim prmOnlyMine = New SqlParameter("blnOnlyMine", SqlDbType.Bit) With {.Direction = ParameterDirection.Input, .Value = False}
							Dim prmTableId = New SqlParameter("intTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Input, .Value = baseTableId}
							
							rstDefSelRecords = objDataAccess.GetDataTable("sp_ASRIntPopulateDefSel", CommandType.StoredProcedure, prmType, prmOnlyMine, prmTableId)
							
							' If atleast one custom report available for the current base table then set to True, False otherwise.
							If (rstDefSelRecords.Rows.Count > 0) Then
								Response.Write("<input type='hidden' id=txtCustomReportGrantedForFindWindow name=txtCustomReportGrantedForFindWindow value=" & True & ">" & vbCrLf)
							Else
								Response.Write("<input type='hidden' id=txtCustomReportGrantedForFindWindow name=txtCustomReportGrantedForFindWindow value=" & False & ">" & vbCrLf)
							End If
							
							' Sets the txtCalendarReportGrantedForFindWindow value
							prmType = New SqlParameter("intType", SqlDbType.Int) With {.Direction = ParameterDirection.Input, .Value = 17}
							prmOnlyMine = New SqlParameter("blnOnlyMine", SqlDbType.Bit) With {.Direction = ParameterDirection.Input, .Value = False}
							prmTableId = New SqlParameter("intTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Input, .Value = baseTableId}
							
							rstDefSelRecords = objDataAccess.GetDataTable("sp_ASRIntPopulateDefSel", CommandType.StoredProcedure, prmType, prmOnlyMine, prmTableId)
							
							' If atleast one calendar report available for the current base table then set to True, False otherwise.
							If (rstDefSelRecords.Rows.Count > 0) Then
								Response.Write("<input type='hidden' id=txtCalendarReportGrantedForFindWindow name=txtCalendarReportGrantedForFindWindow value=" & True & ">" & vbCrLf)
							Else
								Response.Write("<input type='hidden' id=txtCalendarReportGrantedForFindWindow name=txtCalendarReportGrantedForFindWindow value=" & False & ">" & vbCrLf)
							End If
							
							' Sets the txtMailMergeGrantedForFindWindow value
							prmType = New SqlParameter("intType", SqlDbType.Int) With {.Direction = ParameterDirection.Input, .Value = 9}
							prmOnlyMine = New SqlParameter("blnOnlyMine", SqlDbType.Bit) With {.Direction = ParameterDirection.Input, .Value = False}
							prmTableId = New SqlParameter("intTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Input, .Value = baseTableId}
							
							rstDefSelRecords = objDataAccess.GetDataTable("sp_ASRIntPopulateDefSel", CommandType.StoredProcedure, prmType, prmOnlyMine, prmTableId)
							
							' If atleast one mail merge available for the current base table then set to True, False otherwise.
							If (rstDefSelRecords.Rows.Count > 0) Then
								Response.Write("<input type='hidden' id=txtMailMergeGrantedForFindWindow name=txtMailMergeGrantedForFindWindow value=" & True & ">" & vbCrLf)
							Else
								Response.Write("<input type='hidden' id=txtMailMergeGrantedForFindWindow name=txtMailMergeGrantedForFindWindow value=" & False & ">" & vbCrLf)
							End If
							
							'/******* Task 19439: End Changes for the user story 19436: As a user, I want to run reports and utilities from the Find Window  *********/

							Response.Write("<input type='hidden' id=txtInsertGranted name=txtInsertGranted value=" & prmInsertGranted.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtDeleteGranted name=txtDeleteGranted value=" & prmDeleteGranted.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & prmIsFirstPage.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & prmIsLastPage.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & prmColumnType.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtRecordCount name=txtRecordCount value=" & iCount & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & prmTotalRecCount.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtFindRecords name=txtFindRecords value=" & Session("FindRecords") & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & prmFirstRecPos.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtCurrentRecCount name=txtCurrentRecCount value=" & iCount & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & prmColumnSize.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & prmColumnDecimals.Value & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtCancelDateColumn name=txtCancelDateColumn value=" & fCancelDateColumn & ">" & vbCrLf)
							Response.Write("<input type='hidden' id=txtGotoAction name=txtGotoAction value=" & Session("action") & ">" & vbCrLf)
							Response.Write("<input type='hidden' id='txtThousandColumns' name='txtThousandColumns' value='" & sThousandColumns & "'>" & vbCrLf)
							Response.Write("<input type='hidden' id='txtBlankIfZeroColumns' name='txtBlankIfZeroColumns' value='" & sBlankIfZeroColumns & "'>" & vbCrLf)
			
							Session("realSource") = prmRealSource.Value
							
						Catch ex As Exception
							sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
						End Try

					End If
				%>
			</div>
			<%
				If Len(sErrorDescription) = 0 Then
					' Get the summary fields (if required).
					If Not String.IsNullOrEmpty(Session("parentTableID")) AndAlso Session("parentTableID") > 0 Then
						Dim prmCanSelect As New SqlParameter("@pfCanSelect", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
						SPParameters = New SqlParameter() { _
									New SqlParameter("@piHistoryTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))}, _
									New SqlParameter("@piParentTableID ", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))},
									New SqlParameter("@piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))}, _
									prmCanSelect _
						}
						
						Try
							resultsDataTable = objDataAccess.GetDataTable("sp_ASRIntGetSummaryFields", CommandType.StoredProcedure, SPParameters)
						Catch ex As Exception
							sErrorDescription = "The summary field definition could not be retrieved." & vbCrLf & FormatError(ex.Message)
						End Try
												
						Dim aSummaryFields(0, 0) As String
						Dim iTotalCount As Integer
								
						If Len(sErrorDescription) = 0 Then
							sThousSepSummaryFields = ","
							' Read the summary field definitions into an array.
							' We do this as we may be doing a lot of jumping around
							' the definitions and its easy to jump around an array than
							' a recordset.
							ReDim aSummaryFields(9, 0)
							For Each row As DataRow In resultsDataTable.Rows
								iTotalCount = UBound(aSummaryFields, 2) + 1
								ReDim Preserve aSummaryFields(9, iTotalCount)

								aSummaryFields(1, iTotalCount) = row(1).ToString
								aSummaryFields(2, iTotalCount) = row(2).ToString
								aSummaryFields(3, iTotalCount) = row(3).ToString
								aSummaryFields(4, iTotalCount) = row(4).ToString
								aSummaryFields(5, iTotalCount) = row(5).ToString
								aSummaryFields(6, iTotalCount) = row(6).ToString
								aSummaryFields(7, iTotalCount) = row(7).ToString
								aSummaryFields(8, iTotalCount) = row(8).ToString
								aSummaryFields(9, iTotalCount) = row(9).ToString
	
								If row(9).ToString <> "" Then
									sThousSepSummaryFields = sThousSepSummaryFields & row(3).ToString & ","
								End If
							Next

							Dim iRowCount = CLng((iTotalCount + 1) / 2)

							If iTotalCount > 0 Then
								Response.Write("			<div id='row3' style='margin-top: 25px;'>" & vbCrLf)
								Response.Write("<table>" & vbCrLf)
								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("				  <TD colspan=5 height=10></TD>" & vbCrLf)
								Response.Write("				</TR>" & vbCrLf)
								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("				  <TD colspan=5 align=center height=10>" & vbCrLf)
								Response.Write("    				<STRONG>History Summary</STRONG>" & vbCrLf)
								Response.Write("  				</TD>" & vbCrLf)
								Response.Write("				</TR>" & vbCrLf)
								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("				  <TD colspan=5 height=10></TD>" & vbCrLf)
								Response.Write("				</TR>" & vbCrLf)

								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("  				<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)

								Response.Write("  				<TD width=""48%"" height=10>" & vbCrLf)
								Response.Write("      			<TABLE WIDTH=100% class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
							End If

							For iLoop = 1 To iRowCount
								Response.Write("   						<TR>" & vbCrLf)
								Response.Write("								<TD style='width:30%;white-space: nowrap'>" & Replace(aSummaryFields(2, iLoop), "_", " ") & " :</TD>" & vbCrLf)
								Response.Write("								<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
								Response.Write("								<TD style='width:70%'>" & vbCrLf)

								If aSummaryFields(7, iLoop) = 1 Then
									' The summary control is a checkbox.
			%>

			<input type="checkbox" id="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
				name="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
				disabled="disabled">
			<%Else%>
			<%--' The summary control is not a checkbox. Use a textbox for everything else.--%>
			<input type="text" id="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
				name="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
				disabled="disabled" class="text textdisabled width100"
				<%If aSummaryFields(8, iLoop) = 1 Then%>
				style="text-align: right" />
			<%ElseIf aSummaryFields(8, iLoop) = 2 Then%>
					style="text-align: center" />
				<% End If%>
			<%
			End If
			Response.Write("								</TD>" & vbCrLf)
			Response.Write("							</TR>" & vbCrLf)
		Next
		
		If iTotalCount > 0 Then
			Response.Write("      			</TABLE>" & vbCrLf)
			Response.Write("      		</TD>" & vbCrLf)
			Response.Write("  				<TD width=100 height=10>&nbsp;&nbsp;&nbsp;&nbsp;</TD>" & vbCrLf)
			' Do the second column now.
			Response.Write("  				<TD width=""48%"" height=10>" & vbCrLf)
			Response.Write("      			<TABLE WIDTH=100% class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
		End If
				
		Dim iColumn2Index As Integer
				
		For iLoop = 1 To iRowCount
			iColumn2Index = iLoop + iRowCount
						
			If iColumn2Index <= iTotalCount Then
				Response.Write("   						<TR>" & vbCrLf)
				Response.Write("								<TD style='width:30%;white-space: nowrap'>" & Replace(aSummaryFields(2, iColumn2Index), "_", " ") & " :</TD>" & vbCrLf)
				Response.Write("								<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
				Response.Write("								<TD style='width:70%'>" & vbCrLf)

				If aSummaryFields(7, iColumn2Index) = 1 Then%>
			<%--The summary control is a checkbox.--%>
			<input type="checkbox" id="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				name="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				disabled="disabled">
			<%Else%>
			<%--The summary control is not a checkbox. Use a textbox for everything else.--%>
			<input type="text" id="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				name="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				disabled="disabled" class="text textdisabled width100"
				<%If aSummaryFields(8, iColumn2Index) = 1 Then%>
				style="text-align: right" />
			<%ElseIf aSummaryFields(8, iColumn2Index) = 2 Then%>
					style="text-align: center" />
				<%End If%>
			<%	
			End If
		End If

		Response.Write("								</TD>" & vbCrLf)
		Response.Write("							</TR>" & vbCrLf)
	Next

	If iTotalCount > 0 Then
		Response.Write("      			</TABLE>" & vbCrLf)
		Response.Write("      		</TD>" & vbCrLf)
		Response.Write("  				<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
		Response.Write("				</TR>" & vbCrLf)
	End If
End If
Response.Write("</table>" & vbCrLf)
Response.Write("</div>" & vbCrLf)
					
Dim fCanSelect = prmCanSelect.Value

If fCanSelect Then
	SPParameters = New SqlParameter() { _
			New SqlParameter("@piHistoryTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))}, _
			New SqlParameter("@piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))}, _
			New SqlParameter("@piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))} _
	}
	Try
		resultsDataTable = objDataAccess.GetDataTable("spASRIntGetSummaryValues", CommandType.StoredProcedure, SPParameters)
		Dim sTempValue As String
					
		If resultsDataTable.Rows.Count = 0 Then
			  	
			sErrorDescription = "The screen cannot be loaded because you do not have access to its associated parent record"
		End If

		If Len(sErrorDescription) = 0 Then
			For iLoop = 0 To (resultsDataTable.Columns.Count - 1)
				If GeneralUtilities.IsDataColumnDecimal(resultsDataTable.Columns(iLoop)) Then
					sTemp = "," & resultsDataTable.Columns(iLoop).ColumnName & ","
			  	
					If IsDBNull(resultsDataTable.Rows(0)(iLoop)) Then
						sTempValue = "0"
					Else
						sTempValue = resultsDataTable.Rows(0)(iLoop)
					End If

					If InStr(sThousSepSummaryFields, sTemp) > 0 Then
						sTemp = ""
						sTemp = FormatNumber(sTempValue, , True, False, True)
					Else
						sTemp = ""
						sTemp = FormatNumber(sTempValue, , True, False, False)
					End If
					Response.Write("			<INPUT type='hidden' id=txtSummaryData_" & resultsDataTable.Columns(iLoop).ColumnName & " name=txtSummaryData_" & resultsDataTable.Columns(iLoop).ColumnName & " value=""" & sTemp & """>" & vbCrLf)
				Else
					Response.Write("			<INPUT type='hidden' id=txtSummaryData_" & resultsDataTable.Columns(iLoop).ColumnName & " name=txtSummaryData_" & resultsDataTable.Columns(iLoop).ColumnName & " value=""" & resultsDataTable.Rows(0)(iLoop) & """>" & vbCrLf)
				End If
			Next
		End If
	Catch ex As Exception
		sErrorDescription = "The summary field values could not be retrieved." & vbCrLf & FormatError(ex.Message)
	End Try
	
End If
End If
End If
	
If Len(sErrorDescription) = 0 Then
Response.Write("				<input type='hidden' id=txtCurrentTableID name=txtCurrentTableID value=" & Session("tableID") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtCurrentViewID name=txtCurrentViewID value=" & Session("viewID") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtCurrentScreenID name=txtCurrentScreenID value=" & Session("screenID") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtCurrentOrderID name=txtCurrentOrderID value=" & Session("orderID") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtCurrentRecordID name=txtCurrentRecordID value=" & Session("recordID") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtCurrentParentTableID name=txtCurrentParentTableID value=" & Session("parentTableID") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtCurrentParentRecordID name=txtCurrentParentRecordID value=" & Session("parentRecordID") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtRealSource name=txtRealSource value=" & Session("realSource") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtLineage name=txtLineage value=" & Session("lineage") & ">" & vbCrLf)
Response.Write("				<input type='hidden' id=txtFilterDef name=txtFilterDef value=""" & Replace(Session("filterDef_" & Session("tableID")), """", "&quot;") & """>" & vbCrLf)
Response.Write("				<input type='hidden' id=txtFilterSQL name=txtFilterSQL value=""" & Replace(Session("filterSQL_" & Session("tableID")), """", "&quot;") & """>" & vbCrLf)
Response.Write("				<input type='hidden' id='txtThousSepSummary' name='txtThousSepSummary' value='" & sThousSepSummaryFields & "'>" & vbCrLf)
Response.Write("				<input type='hidden' id='txtMaxRequestLength' name='txtMaxRequestLength' value='" & Session("maxRequestLength") & "'>" & vbCrLf)
Response.Write("				<input type='hidden' id=txtSelectedRecordsInFindGrid name=txtSelectedRecordsInFindGrid value=" & Session("OptionSelectedRecordIds") & ">" & vbCrLf)
			
End If

Response.Write("				<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
			%>
		</div>
		
		<input type="hidden" id="txtFindEditRowID" value=""/>
		<input type="hidden" id="txtFindEditLastRowID" value=""/>
		<input type="hidden" id="txtFindEditRowData" value=""/>
		<%=Html.AntiForgeryToken()%>
	</form>


	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">	

	<script type="text/javascript">		
		find_window_onload();

		if (!menu_isSSIMode()) {
			$('div#workframeset').animate({ scrollTop: 0 }, 0);
		}

	</script>
</div>
