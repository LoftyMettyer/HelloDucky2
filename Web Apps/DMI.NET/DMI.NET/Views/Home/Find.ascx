<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<%
	'Data access variables
	Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
	Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
	Dim SPParameters() As SqlParameter
	Dim resultDataSet As DataSet
	Dim rstFindRecords As DataTable
	Dim resultsDataTable As DataTable
%>
<script src="<%: Url.LatestContent("~/bundles/recordedit")%>" type="text/javascript"></script>

<script type="text/javascript">
	$(document).ready(function () {
		if ('<%=session("linktype")%>' == 'multifind') {
			//for multifind (SSI views) show relevant buttons with applicable functions
			menu_setVisibletoolbarGroupById("mnuSectionRecordFindEdit", false);
			menu_setVisibleMenuItem("mnutoolAccessLinksFind", true);
			menu_setVisibleMenuItem("mnutoolCancelLinksFind", false);

			//redo the doubleclick function			
			setTimeout('$("#findGridTable").jqGrid("setGridParam", { ondblClickRow: function (rowID) {doEdit();}});', 200);

		} else {
			menu_setVisibletoolbarGroupById("mnuSectionRecordFindEdit", true);
			menu_setVisibleMenuItem("mnutoolAccessLinksFind", false);
			menu_setVisibleMenuItem("mnutoolCancelLinksFind", false);
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
</script>

<div id="divFindForm" <%=session("BodyTag")%>>
	<form action="" class="absolutefull" method="POST" id="frmFindForm" name="frmFindForm">
		<div class="absolutefull">
			<div id="row1" style="margin-left: 20px; margin-right: 20px">
				<%
					Dim sErrorDescription As String = ""
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
							Response.Write(String.Format("<div class='pageTitleDiv'><a href='{0}' title='Back'><i class='pageTitleIcon icon-circle-arrow-left'></i></a><span class='pageTitle'>Find - " & _
											Replace(prm_psTitle.Value.ToString, "_", " ") & "</span>" & vbCrLf, homelinkURL))
							Response.Write("<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & prm_pfQuickEntry.Value.ToString & "></div>" & vbCrLf)
						End If
					Else
						' SSI View.
						Dim homelinkURL = "javascript:loadPartialView(""linksMain"", ""Home"", ""workframe"", null);"
						Response.Write(String.Format("<div class='pageTitleDiv'><a href='{0}' title='Back'><i class='pageTitleIcon icon-circle-arrow-left'></i></a><span class='pageTitle'>" & _
								ViewBag.pageTitle & "</span>" & vbCrLf, homelinkURL))
						Response.Write("<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & ViewBag.pageTitle & "></div>" & vbCrLf)
					End If
				%>
			</div>
			<div id="findGridRow" style="height: <%If Session("parentTableID") > 0 Then%>65%<%Else%>85%<%End If%>; margin-right: 20px; margin-left: 20px;">
				<%
					Dim sTemp As String
					Dim sThousandColumns As String = ""
					Dim sBlankIfZeroColumns As String
					Dim sColDef As String
					Dim iCount As Integer
					Dim sAddString As String
								
					Const AD_STATE_OPEN = 1
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
							New SqlParameter("@psFilterDef", SqlDbType.VarChar, -1) With {.Value = Session("filterDef")}, _
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
							New SqlParameter("@psLocaleDateFormat", SqlDbType.VarChar, 255) With {.Value = Session("LocaleDateFormat")} _
						}

						Try
							resultDataSet = objDataAccess.GetDataSet("sp_ASRIntGetFindRecords3", SPParameters)

							If prmSomeSelectable.Value = 0 Then								
								sErrorDescription = "You do not have permission to read any of the selected order's find columns."
							Else
													
								' Get the recordset parameters
								sThousandColumns = resultDataSet.Tables(0).Rows(0)("ThousandColumns").ToString()
								sBlankIfZeroColumns = resultDataSet.Tables(0).Rows(0)("BlankIfZeroColumns").ToString()
									
								' Get the actual data
								rstFindRecords = resultDataSet.Tables(1)

								' Instantiate and initialise the grid. 
								Response.Write("<table class='outline' style='width : 100%; ' id='findGridTable'>" & vbCrLf)
								Response.Write("<div id='pager-coldata'></div>" & vbCrLf)
										
								iCount = 0
								For Each row As DataRow In rstFindRecords.Rows
									sAddString = ""
						
									For iloop = 0 To (rstFindRecords.Columns.Count - 1)
										If iloop > 0 Then
											sAddString = sAddString & "	"
										End If
							
										If iCount = 0 Then
											sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString.Replace("System.", "")
											Response.Write("<INPUT type='hidden' id=txtFindColDef_" & iloop & " name=txtFindColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
										End If
							
										If rstFindRecords.Columns(iloop).DataType = GetType(System.DateTime) Then
											' Field is a date so format as such.
											sAddString = sAddString & ConvertSQLDateToLocale(row(iloop))
										ElseIf GeneralUtilities.IsDataColumnDecimal(rstFindRecords.Columns(iloop)) Then
											' Field is a numeric so format as such.
											If Not IsDBNull(row(iloop)) Then
												
												Dim numberAsString As String = row(iloop).ToString()
												Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(".", System.StringComparison.Ordinal)
												Dim numberOfDecimals As Integer = 0
												If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length
												
												If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
													sTemp = FormatNumber(row(iloop), numberOfDecimals, TriState.True, TriState.False, TriState.True)
												Else
													sTemp = FormatNumber(row(iloop), numberOfDecimals, TriState.True, TriState.False, TriState.False)
												End If
												sTemp = Replace(sTemp, ".", "x")
												sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
												sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
												sAddString = sAddString & sTemp
											End If
										Else
											If Not IsDBNull(row(iloop)) Then
												sAddString = sAddString & Replace(row(iloop), """", "&quot;")
											End If
										End If
									Next

									Response.Write("<input type='hidden' id=txtAddString_" & iCount & " name=txtAddString_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
									iCount += 1
								Next
							
							End If

							Response.Write("</table>")
				
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
					If Session("parentTableID") > 0 Then
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
						
						Dim sThousSepSummaryFields As String
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
								Response.Write("   							<TD nowrap=true>" & Replace(aSummaryFields(2, iLoop), "_", " ") & " :</TD>" & vbCrLf)
								Response.Write("								<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
								Response.Write("								<TD width=""100%"">" & vbCrLf)

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
				class="text textdisabled" disabled="disabled"
				<%If aSummaryFields(8, iLoop) = 1 Then%>
				style="width: 100%; text-align: right" />
			<%ElseIf aSummaryFields(8, iLoop) = 2 Then%>
					style="width: 100%;text-align: center" />
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
				Response.Write("								<TD nowrap=true>" & Replace(aSummaryFields(2, iColumn2Index), "_", " ") & " :</TD>" & vbCrLf)
				Response.Write("								<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
				Response.Write("								<TD width=""100%"">" & vbCrLf)

				If aSummaryFields(7, iColumn2Index) = 1 Then%>
			<%--The summary control is a checkbox.--%>
			<input type="checkbox" id="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				name="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				disabled="disabled">
			<%Else%>
			<%--The summary control is not a checkbox. Use a textbox for everything else.--%>
			<input type="text" id="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				name="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				disabled="disabled" class="text textdisabled"
				<%If aSummaryFields(8, iColumn2Index) = 1 Then%>
				style="width: 100%; text-align: right" />
			<%ElseIf aSummaryFields(8, iColumn2Index) = 2 Then%>
					style="width: 100%;text-align: center" />
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
					sTemp = Replace(sTemp, ".", "x")
					sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
					sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
							
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
Response.Write("				<input type='hidden' id=txtFilterDef name=txtFilterDef value=""" & Replace(Session("filterDef"), """", "&quot;") & """>" & vbCrLf)
Response.Write("				<input type='hidden' id=txtFilterSQL name=txtFilterSQL value=""" & Replace(Session("filterSQL"), """", "&quot;") & """>" & vbCrLf)
End If

Response.Write("				<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
			%>
		</div>
	</form>

	<form id="frmTBData" name="frmTBData">
		<%
			If CLng(Session("tableID")) = CLng(Session("TB_TBTableID")) Then
				Response.Write("				<input type='hidden' id=txtTBCancelCourseDate name=txtTBCancelCourseDate value=""" & Session("lineage") & """>")
			End If
		%>
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>

	<script type="text/javascript">
		find_window_onload();
		
		if (menu_isSSIMode()) {
			$('.ViewDescription p').text('My Dashboard');		
		} else {
			$('.ViewDescription p').text('');			
		}

	</script>
</div>
