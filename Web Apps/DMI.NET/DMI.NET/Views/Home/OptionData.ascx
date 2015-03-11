<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
	Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
%>
<script src="<%: Url.LatestContent("~/bundles/recordedit")%>" type="text/javascript"></script>

<form action="optionData_Submit" method="post" id="frmGetOptionData" name="frmGetOptionData">
		<input type="hidden" id="txtOptionAction" name="txtOptionAction">
		<input type="hidden" id="txtOptionTableID" name="txtOptionTableID">
		<input type="hidden" id="txtOptionViewID" name="txtOptionViewID">
		<input type="hidden" id="txtOptionOrderID" name="txtOptionOrderID">
		<input type="hidden" id="txtOptionColumnID" name="txtOptionColumnID">
		<input type="hidden" id="txtOptionPageAction" name="txtOptionPageAction">
		<input type="hidden" id="txtOptionFirstRecPos" name="txtOptionFirstRecPos">
		<input type="hidden" id="txtOptionCurrentRecCount" name="txtOptionCurrentRecCount">
		<input type="hidden" id="txtGotoLocateValue" name="txtGotoLocateValue">
		<input type="hidden" id="txtOptionCourseTitle" name="txtOptionCourseTitle">
		<input type="hidden" id="txtOptionRecordID" name="txtOptionRecordID">
		<input type="hidden" id="txtOptionLinkRecordID" name="txtOptionLinkRecordID">
		<input type="hidden" id="txtOptionValue" name="txtOptionValue">
		<input type="hidden" id="txtOptionSQL" name="txtOptionSQL">
		<input type="hidden" id="txtOptionPromptSQL" name="txtOptionPromptSQL">
		<input type="hidden" id="txtOptionOnlyNumerics" name="txtOptionOnlyNumerics">
		<input type="hidden" id="txtOptionLookupColumnID" name="txtOptionLookupColumnID">
		<input type="hidden" id="txtOptionLookupFilterValue" name="txtOptionLookupFilterValue">
		<input type="hidden" id="txtOptionIsLookupTable" name="txtOptionIsLookupTable">
		<input type="hidden" id="txtOptionParentTableID" name="txtOptionParentTableID">
		<input type="hidden" id="txtOptionParentRecordID" name="txtOptionParentRecordID">
		<input type="hidden" id="txtOption1000SepCols" name="txtOption1000SepCols">
		<%=Html.AntiForgeryToken()%>
</form>

<form id="frmOptionData" name="frmOptionData">
		<%
			Dim aPrompts(1, 0)

			Session("flagOverrideFilter") = False

			Dim objUtilities As Utilities

			Dim sErrorDescription As String = ""
			Dim sNonFatalErrorDescription As String = ""

			Dim prmThousandColumns As SqlParameter
			Dim sThousandColumns As String = ""
			Dim sBlankIfZeroColumns As String = ""
		
			Dim iCount As Integer
			Dim sAddString As String
			Dim sColDef As String
			Dim sTemp As String
		
			Dim j As Integer
			Dim sPrompts As String
			Dim iIndex1 As Integer
			Dim iIndex2 As Integer
				
			Response.Write("<INPUT type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

			' Get the required record count if we have a query.
			'	if len(session("selectSQL")) > 0 then
			If Session("optionAction") = OptionActionType.LOADFIND Then
			
				Try
					Get1000SeparatorBlankIfZeroFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")), sThousandColumns, sBlankIfZeroColumns)

					Dim prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
					Dim prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
					Dim prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
					Dim prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("optionFirstRecPos"))}
					Dim prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		
					Dim dsFindRecords = objDataAccess.GetDataSet("sp_ASRIntGetLinkFindRecords" _
						, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
						, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))} _
						, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionOrderID"))} _
						, prmError _
						, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = CleanNumeric(Session("FindRecords"))} _
						, prmIsFirstPage _
						, prmIsLastPage _
						, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("optionLocateValue")} _
						, prmColumnType _
						, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("optionPageAction")} _
						, prmTotalRecCount _
						, prmFirstRecPos _
						, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("currentRecCount"))} _
						, New SqlParameter("psExcludedIDs", SqlDbType.VarChar, -1) With {.Value = ""} _
						, prmColumnSize _
						, prmColumnDecimals)

					iCount = 0
					If dsFindRecords.Tables.Count > 0 Then
						Dim rstFindRecords = dsFindRecords.Tables(0)
						For Each objRow As DataRow In rstFindRecords.Rows
							sAddString = ""
						
							For iloop = 0 To (rstFindRecords.Columns.Count - 1)
								If iloop > 0 Then
									sAddString = sAddString & "	"
								End If
							
								If iCount = 0 Then
									sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString()
									Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
								End If
							
								If rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.datetime" Then
									' Field is a date so format as such.
									sAddString = sAddString & ConvertSQLDateToLocale(objRow(iloop).ToString())
								ElseIf rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.decimal" Then
									' Field is a numeric so format as such.
									If Not IsDBNull(objRow(iloop)) Then
										
										
										Dim numberAsString As String = objRow(iloop).ToString()
										Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(LocaleDecimalSeparator(), System.StringComparison.Ordinal)
										Dim numberOfDecimals As Integer = 0
										If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length
										
										If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
											sTemp = FormatNumber(objRow(iloop), numberOfDecimals, TriState.True, TriState.False, TriState.True)
										Else
											sTemp = FormatNumber(objRow(iloop), numberOfDecimals, TriState.True, TriState.False, TriState.False)
										End If
										sAddString = sAddString & sTemp
									End If
								Else
									If Not IsDBNull(objRow(iloop)) Then
										sAddString = sAddString & Replace(objRow(iloop).ToString, """", "&quot;")
									End If
								End If
							Next

							Response.Write("<input type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
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

			ElseIf Session("optionAction") = OptionActionType.LOADLOOKUPFIND Then
				' StoredProc defaults to 1000 if no value set.
				Dim iLookupFindRecords As Integer = 10000	' replaces CleanNumeric(Session("FindRecords")
						
				Dim rstFindRecords As DataTable
								
				' Check if the filter value column is in the current screen.
				' If not, try and get the filter value from the database.
				If Len(Session("optionFilterValue")) = 0 Then

					Dim prmFilterValue = New SqlParameter("psFilterValue", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
					Dim prmADOError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}

					Try

						objDataAccess.ExecuteSP("spASRIntGetLookupFilterValue" _
							, New SqlParameter("@piScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))} _
							, New SqlParameter("@piColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionColumnID"))} _
							, New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))} _
							, New SqlParameter("@piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))} _
							, New SqlParameter("@piRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
							, prmFilterValue _
							, New SqlParameter("@piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionParentTableID"))} _
							, New SqlParameter("@piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionParentRecordID"))} _
							, prmADOError)

						Session("optionFilterValue") = prmFilterValue.Value.ToString()
						Session("flagOverrideFilter") = prmADOError.Value
						
					Catch ex As Exception
						sErrorDescription = "Error reading the lookup filter value." & vbCrLf & ex.Message.RemoveSensitive
					End Try
					
				End If

				Dim prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("optionFirstRecPos"))}
				Dim prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmLookupColumnGridPosition = New SqlParameter("piLookupColumnGridNumber", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		
				If Len(sErrorDescription) = 0 Then
					sThousandColumns = ""

					If Session("IsLookupTable") = "False" Then
						Try
							Get1000SeparatorBlankIfZeroFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")), sThousandColumns, sBlankIfZeroColumns)
						Catch ex As Exception
							sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
						End Try
														
						rstFindRecords = objDataAccess.GetFromSP("spASRIntGetLookupFindRecords2" _
							, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
							, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))} _
							, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionOrderID"))} _
							, New SqlParameter("piLookupColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLookupColumnID"))} _
							, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = iLookupFindRecords} _
							, prmIsFirstPage _
							, prmIsLastPage _
							, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("optionLocateValue")} _
							, prmColumnType _
							, prmColumnSize _
							, prmColumnDecimals _
							, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("optionPageAction")} _
							, prmTotalRecCount _
							, prmFirstRecPos _
							, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("currentRecCount"))} _
							, New SqlParameter("psFilterValue", SqlDbType.VarChar, -1) With {.Value = Session("optionFilterValue")} _
							, New SqlParameter("piCallingColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionColumnID"))} _
							, prmLookupColumnGridPosition _
							, New SqlParameter("pfOverrideFilter", SqlDbType.Bit) With {.Value = Session("flagOverrideFilter")})


					Else
						prmThousandColumns = New SqlParameter("@ps1000SeparatorCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmBlankIfZeroColumns As New SqlParameter("@psBlanIfZeroCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Try
							objDataAccess.ExecuteSP("spASRIntGetLookupFindColumnInfo", _
													New SqlParameter("@piLookupColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLookupColumnID"))}, _
													prmThousandColumns, _
													prmBlankIfZeroColumns
							)
						Catch ex As Exception
							sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
						End Try


						sThousandColumns = prmThousandColumns.Value.ToString()
						
						rstFindRecords = objDataAccess.GetFromSP("spASRIntGetLookupFindRecords" _
							, New SqlParameter("piLookupColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLookupColumnID"))} _
							, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = iLookupFindRecords} _
							, prmIsFirstPage _
							, prmIsLastPage _
							, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("optionLocateValue")} _
							, prmColumnType _
							, prmColumnSize _
							, prmColumnDecimals _
							, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("optionPageAction")} _
							, prmTotalRecCount _
							, prmFirstRecPos _
							, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionCurrentRecCount"))} _
							, New SqlParameter("psFilterValue", SqlDbType.VarChar, -1) With {.Value = Session("optionFilterValue")} _
							, New SqlParameter("piCallingColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionColumnID"))} _
							, New SqlParameter("pfOverrideFilter", SqlDbType.Bit) With {.Value = Session("flagOverrideFilter")})
						
					End If
				
					iCount = 0

					For Each objRow As DataRow In rstFindRecords.Rows
						sAddString = ""
							
						For iloop = 0 To (rstFindRecords.Columns.Count - 1)
							If iloop > 0 Then
								sAddString = sAddString & "	"
							End If
								
							If iCount = 0 Then
								sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString()
								Response.Write("<input type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
							End If
								
							If rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.datetime" Then
								' Field is a date so format as such.
								sAddString = sAddString & ConvertSQLDateToLocale(objRow(iloop).ToString())
							ElseIf rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.decimal" Then
								' Field is a numeric so format as such.
								If Not IsDBNull(objRow(iloop)) Then
									
										
									Dim numberAsString As String = objRow(iloop).ToString()
									Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(LocaleDecimalSeparator(), System.StringComparison.Ordinal)
									Dim numberOfDecimals As Integer = 0
									If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length
									
									
									If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
										sTemp = FormatNumber(objRow(iloop), numberOfDecimals, TriState.True, TriState.False, TriState.True)
									Else
										sTemp = FormatNumber(objRow(iloop), numberOfDecimals, TriState.True, TriState.False, TriState.False)
									End If
									sAddString = sAddString & sTemp
								End If
							Else
								If Not IsDBNull(objRow(iloop)) Then
									sAddString = sAddString & Replace(objRow(iloop).ToString(), """", "&quot;")
								End If
							End If
						Next

						Response.Write("<input type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
						
						iCount += 1
					Next
					
					Response.Write("<input type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & prmIsFirstPage.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & prmIsLastPage.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & prmColumnType.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & prmColumnSize.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & prmColumnDecimals.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & prmTotalRecCount.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & prmFirstRecPos.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFilterOverride name=txtFilterOverride value=" & Session("flagOverrideFilter") & ">" & vbCrLf)
			
					
					If Session("IsLookupTable") = "False" Then
						Response.Write("<input type='hidden' id=txtLookupColumnGridPosition name=txtLookupColumnGridPosition value=" & prmLookupColumnGridPosition.Value & ">" & vbCrLf)
					Else
						Response.Write("<input type='hidden' id=txtLookupColumnGridPosition name=txtLookupColumnGridPosition value=0>" & vbCrLf)
					End If
							

				End If
				
				
			ElseIf Session("optionAction") = OptionActionType.LOADTRANSFERCOURSE Then
				sThousandColumns = ""
			
				Try
					Get1000SeparatorBlankIfZeroFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")), sThousandColumns, sBlankIfZeroColumns)
				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
				End Try

				
				
				Dim prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("optionFirstRecPos"))}
				Dim prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		
				Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetTransferCourseRecords" _
					, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
					, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))} _
					, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionOrderID"))} _
					, New SqlParameter("psCourseTitle", SqlDbType.VarChar, -1) With {.Value = Session("optionCourseTitle")} _
					, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
					, prmError _
					, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = CleanNumeric(Session("FindRecords"))} _
					, prmIsFirstPage _
					, prmIsLastPage _
					, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("optionLocateValue")} _
					, prmColumnType _
					, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Session("optionPageAction")} _
					, prmTotalRecCount _
					, prmFirstRecPos _
					, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionCurrentRecCount"))} _
					, prmColumnSize _
					, prmColumnDecimals)

				For Each objRow As DataRow In rstFindRecords.Rows

					sAddString = ""
						
					For iloop = 0 To (rstFindRecords.Columns.Count - 1)
						If iloop > 0 Then
							sAddString = sAddString & "	"
						End If
							
						If iCount = 0 Then
							sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString()
							Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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

					Response.Write("<input type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
					iCount += 1
				Next
	
				Response.Write("<input type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & prmIsFirstPage.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & prmIsLastPage.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & prmColumnType.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & prmTotalRecCount.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & prmFirstRecPos.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & prmColumnSize.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & prmColumnDecimals.Value & ">" & vbCrLf)
			
				
			ElseIf Session("optionAction") = OptionActionType.LOADBOOKCOURSE Then

				Dim prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("firstRecPos"))}
				Dim prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				Try
					Get1000SeparatorBlankIfZeroFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")), sThousandColumns, sBlankIfZeroColumns)
				
					Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetBookCourseRecords" _
						, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
						, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))} _
						, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionOrderID"))} _
						, New SqlParameter("piWLRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
						, prmError _
						, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = CleanNumeric(Session("FindRecords"))} _
						, prmIsFirstPage _
						, prmIsLastPage _
						, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("optionLocateValue")} _
						, prmColumnType _
						, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("optionPageAction")} _
						, prmTotalRecCount _
						, prmFirstRecPos _
						, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionCurrentRecCount"))} _
						, prmColumnSize _
						, prmColumnDecimals)
					

					iCount = 0
					For Each objRow As DataRow In rstFindRecords.Rows

						sAddString = ""
						
						For iloop = 0 To (rstFindRecords.Columns.Count - 1)
							If iloop > 0 Then
								sAddString = sAddString & "	"
							End If
							
							If iCount = 0 Then
								sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString()
								Response.Write("<input type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
							End If
							
							If rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.datetime" Then
								' Field is a date so format as such.
								'sAddString = sAddString & ConvertSQLDateToLocale(objRow(iloop).ToString())
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

						Response.Write("<input type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
						iCount += 1
					Next
	

				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & ex.Message.RemoveSensitive
	
				End Try
						
				Response.Write("<input type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & prmIsFirstPage.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & prmIsLastPage.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & prmColumnType.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & prmTotalRecCount.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & prmFirstRecPos.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & prmColumnSize.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & prmColumnDecimals.Value & ">" & vbCrLf)
			
				
				
			ElseIf Session("optionAction") = OptionActionType.SELECTBOOKCOURSE_3 Then
				
				Try
					objDataAccess.ExecuteSP("sp_ASRIntBookCourse" _
						, New SqlParameter("piWLRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkRecordID"))} _
						, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
						, New SqlParameter("psStatus", SqlDbType.VarChar, -1) With {.Value = Session("optionValue")})
					
					Session("optionAction") = OptionActionType.BOOKCOURSESUCCESS
					
				Catch ex As Exception
					sNonFatalErrorDescription = "The booking could not be made." & vbCrLf & ex.Message.RemoveSensitive
					Session("optionAction") = OptionActionType.BOOKCOURSEERROR
				End Try


			ElseIf Session("optionAction") = OptionActionType.SELECTADDFROMWAITINGLIST_3 Then
				
				Try
					objDataAccess.ExecuteSP("sp_ASRIntAddFromWaitingList" _
						, New SqlParameter("piEmpRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkRecordID"))} _
						, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
						, New SqlParameter("psStatus", SqlDbType.VarChar - 1) With {.Value = Session("optionValue")})
					
					Session("optionAction") = OptionActionType.ADDFROMWAITINGLISTSUCCESS
					
				Catch ex As Exception
					sNonFatalErrorDescription = "The booking could not be made." & vbCrLf & sNonFatalErrorDescription
					Session("optionAction") = OptionActionType.ADDFROMWAITINGLISTERROR
					
				End Try
							

			ElseIf Session("optionAction") = OptionActionType.LOADTRANSFERBOOKING Then
				sThousandColumns = ""

				Dim prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("optionFirstRecPos"))}
				Dim prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmStatus = New SqlParameter("psStatus", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				Try
					Get1000SeparatorBlankIfZeroFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")), sThousandColumns, sBlankIfZeroColumns)
		
					Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetTransferBookingRecords" _
						, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
						, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))} _
						, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionOrderID"))} _
						, New SqlParameter("piTBRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
						, prmError _
						, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = CleanNumeric(Session("FindRecords"))} _
						, prmIsFirstPage _
						, prmIsLastPage _
						, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("optionLocateValue")} _
						, prmColumnType _
						, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("optionPageAction")} _
						, prmTotalRecCount _
						, prmFirstRecPos _
						, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionCurrentRecCount"))} _
						, prmErrorMessage _
						, prmColumnSize _
						, prmColumnDecimals _
						, prmStatus)

					If prmErrorMessage.Value.ToString().Length = 0 Then
						iCount = 0
						For Each objRow As DataRow In rstFindRecords.Rows
							sAddString = ""
						
							For iloop = 0 To (rstFindRecords.Columns.Count - 1)
								If iloop > 0 Then
									sAddString = sAddString & "	"
								End If
							
								If iCount = 0 Then
									sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString()
									Response.Write("<input type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
								End If
							
								If rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.datetime" Then
									' Field is a date so format as such.
									sAddString = sAddString & ConvertSQLDateToLocale(objRow(iloop).ToString())
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

							Response.Write("<input type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
							iCount += 1
						Next
					
					End If
					

				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
				End Try


				Response.Write("<input type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & prmIsFirstPage.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & prmIsLastPage.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & prmColumnType.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & prmTotalRecCount.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & prmFirstRecPos.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtErrorMessage2 name=txtErrorMessage2 value=""" & Replace(prmErrorMessage.Value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & prmColumnSize.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & prmColumnDecimals.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtStatus name=txtStatus value=""" & Replace(prmStatus.Value, """", "&quot;") & """>" & vbCrLf)
			
				
			ElseIf Session("optionAction") = OptionActionType.LOADADDFROMWAITINGLIST Then
				sThousandColumns = ""
			
				Try
					Get1000SeparatorBlankIfZeroFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")), sThousandColumns, sBlankIfZeroColumns)

					Dim prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
					Dim prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
					Dim prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
					Dim prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("optionFirstRecPos"))}
					Dim prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
					
					Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetAddFromWaitingListRecords" _
						, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
						, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))} _
						, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionOrderID"))} _
						, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
						, prmError _
						, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = 100000} _
						, prmIsFirstPage _
						, prmIsLastPage _
						, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("optionLocateValue")} _
						, prmColumnType _
						, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Session("optionPageAction")} _
						, prmTotalRecCount _
						, prmFirstRecPos _
						, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionCurrentRecCount"))} _
						, prmErrorMessage _
						, prmColumnSize _
						, prmColumnDecimals)


					For Each objRow As DataRow In rstFindRecords.Rows

						sAddString = ""
						
						For iloop = 0 To (rstFindRecords.Columns.Count - 1)
							If iloop > 0 Then
								sAddString = sAddString & "	"
							End If
							
							If iCount = 0 Then
								sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString()
								Response.Write("<input type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
							End If
							
							Dim numberAsString As String = objRow(iloop).ToString()
							Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(LocaleDecimalSeparator(), StringComparison.Ordinal)
							Dim numberOfDecimals As Integer = 0
							If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length
							
							
							If rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.datetime" Then
								' Field is a date so format as such.
								sAddString = sAddString & ConvertSQLDateToLocale(objRow(iloop).ToString())
							ElseIf rstFindRecords.Columns(iloop).DataType.ToString().ToLower() = "system.decimal" Then
								' Field is a numeric so format as such.
								If Not IsDBNull(objRow(iloop)) Then
									If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
										sTemp = FormatNumber(objRow(iloop), numberOfDecimals, True, False, True)
									Else
										sTemp = FormatNumber(objRow(iloop), numberOfDecimals, True, False, False)
									End If
									sAddString = sAddString & sTemp
								End If
							Else
								If Not IsDBNull(objRow(iloop)) Then
									sAddString = sAddString & Replace(objRow(iloop).ToString(), """", "&quot;")
								End If
							End If
							
						Next

						Response.Write("<input type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
						iCount += 1
					Next
	

					Response.Write("<input type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & prmIsFirstPage.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & prmIsLastPage.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & prmColumnType.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & prmTotalRecCount.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & prmFirstRecPos.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtErrorMessage2 name=txtErrorMessage2 value=""" & Replace(prmErrorMessage.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & prmColumnSize.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & prmColumnDecimals.Value & ">" & vbCrLf)
			
				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
				End Try


			ElseIf Session("optionAction") = OptionActionType.SELECTTRANSFERBOOKING_2 Then
						
				Try
					objDataAccess.ExecuteSP("sp_ASRIntTransferCourse" _
						, New SqlParameter("piTBRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
						, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkRecordID"))})
					
					Session("optionAction") = OptionActionType.TRANSFERBOOKINGSUCCESS
					
				Catch ex As Exception
					sNonFatalErrorDescription = "The booking could not be transferred." & vbCrLf & sNonFatalErrorDescription
					Session("optionAction") = OptionActionType.TRANSFERBOOKINGERROR
				End Try


			ElseIf Session("optionAction") = OptionActionType.GETBULKBOOKINGSELECTION Then
				If UCase(Session("optionPageAction")) = "FILTER" Then
					objUtilities = Session("UtilitiesObject")

					j = 0
					ReDim Preserve aPrompts(1, 0)
					sPrompts = Session("optionPromptSQL")
					If Len(Session("optionPromptSQL")) > 0 Then
						Do While Len(sPrompts) > 0
							iIndex1 = InStr(sPrompts, vbTab)
					
							If iIndex1 > 0 Then
								iIndex2 = InStr(iIndex1 + 1, sPrompts, vbTab)
					
								If iIndex2 > 0 Then
									ReDim Preserve aPrompts(1, j)
								
									aPrompts(0, j) = Left(sPrompts, iIndex1 - 1)
									aPrompts(1, j) = Mid(sPrompts, iIndex1 + 1, iIndex2 - iIndex1 - 1)
								
									sPrompts = Mid(sPrompts, iIndex2 + 1)
								
									j = j + 1
								End If
							End If
						Loop
					End If
					Session("optionPromptSQL") = objUtilities.GetFilteredIDs(Session("optionRecordID"), aPrompts)

					objUtilities = Nothing
				End If


				Try
					

					Dim prmPromptSQL2 As New SqlParameter("psPromptSQL", SqlDbType.VarChar, -1)
					Dim prmErrorMessage2 As New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
					'				Try

					If Len(Session("optionPromptSQL")) = 0 Then
						prmPromptSQL2.Value = ""
					Else
						prmPromptSQL2.Value = Session("optionPromptSQL")
					End If


					objUtilities = Session("UtilitiesObject")
					objUtilities.UDFFunctions(True)

					Dim rstFindRecords2 = objDataAccess.GetFromSP("sp_ASRIntGetBulkBookingRecords" _
						, New SqlParameter("psSelectionType", SqlDbType.VarChar, -1) With {.Value = Session("optionPageAction")} _
						, New SqlParameter("piSelectionID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
						, New SqlParameter("psSelectedIDs", SqlDbType.VarChar, -1) With {.Value = Session("optionValue")} _
						, prmPromptSQL2 _
						, prmErrorMessage2)

					objUtilities.UDFFunctions(False)
					
					For Each objRow As DataRow In rstFindRecords2.Rows
						sAddString = ""
						
						For iloop = 0 To (rstFindRecords2.Columns.Count - 1)
							If iloop > 0 Then
								sAddString = sAddString & "	"
							End If
							
							If iCount = 0 Then
								sColDef = Replace(rstFindRecords2.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords2.Columns(iloop).DataType.Name
								Response.Write("<input type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
							End If
							
							If rstFindRecords2.Columns(iloop).DataType.Name.ToLower() = "system.datetime" Then
								' Field is a date so format as such.
								sAddString = sAddString & ConvertSQLDateToLocale(objRow(iloop).ToString())
							ElseIf rstFindRecords2.Columns(iloop).DataType.Name.ToLower() = "system.decimal" Then
								' Field is a numeric so format as such.
								If Not IsDBNull(objRow(iloop)) Then
									If Mid(Session("option1000SepCols"), iloop + 1, 1) = "1" Then
										sTemp = FormatNumber(objRow(iloop), objRow(iloop).NumericScale, True, False, True)
									Else
										sTemp = FormatNumber(objRow(iloop), objRow(iloop).NumericScale, True, False, False)
									End If
									sAddString = sAddString & sTemp
								End If
							Else
								If Not IsDBNull(objRow(iloop)) Then
									sAddString = sAddString & Replace(objRow(iloop).ToString(), """", "&quot;")
								End If
							End If
						Next

						Response.Write("<input type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
						iCount += 1
					Next
					
				Catch ex As Exception
					sErrorDescription = "Error reading the find records." & vbCrLf & FormatError(ex.Message)
					
				End Try
			
				
			ElseIf Session("optionAction") = OptionActionType.GETPICKLISTSELECTION Then
				If UCase(Session("optionPageAction")) = "FILTER" Then
					objUtilities = Session("UtilitiesObject")

					j = 0
					ReDim Preserve aPrompts(1, 0)
					sPrompts = Session("optionPromptSQL")
					If Len(Session("optionPromptSQL")) > 0 Then
						Do While Len(sPrompts) > 0
							iIndex1 = InStr(sPrompts, vbTab)
					
							If iIndex1 > 0 Then
								iIndex2 = InStr(iIndex1 + 1, sPrompts, vbTab)
					
								If iIndex2 > 0 Then
									ReDim Preserve aPrompts(1, j)
								
									aPrompts(0, j) = Left(sPrompts, iIndex1 - 1)
									aPrompts(1, j) = Mid(sPrompts, iIndex1 + 1, iIndex2 - iIndex1 - 1)
								
									sPrompts = Mid(sPrompts, iIndex2 + 1)
								
									j = j + 1
								End If
							End If
						Loop
					End If
					Session("optionPromptSQL") = objUtilities.GetFilteredIDs(Session("optionRecordID"), aPrompts)

					objUtilities = Nothing
				End If


				Try

					Dim prmPromptSQL = New SqlParameter("psPromptSQL", SqlDbType.VarChar, -1)
					Dim prmErrMsg = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
					Dim prmExpectedCount = New SqlParameter("piExpectedRecords", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

					If Len(Session("optionPromptSQL")) = 0 Then
						prmPromptSQL.Value = ""
					Else
						prmPromptSQL.Value = Session("optionPromptSQL")
					End If

					objUtilities = CType(Session("UtilitiesObject"), Utilities)
					objUtilities.UDFFunctions(True)
		
					Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetSelectedPicklistRecords" _
									, New SqlParameter("@psSelectionType", SqlDbType.VarChar, 255) With {.Value = Session("optionPageAction")} _
									, New SqlParameter("@piSelectionID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
									, New SqlParameter("@psSelectedIDs", SqlDbType.VarChar, -1) With {.Value = Session("optionValue")} _
									, prmPromptSQL _
									, New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
									, prmErrMsg, prmExpectedCount)

					objUtilities.UDFFunctions(False)
			
					'Output the column definitions; this needs to be done even if the recordset doesn't contain any data so jqGrid ill show and empty grid with column names
					For iloop = 0 To (rstFindRecords.Columns.Count - 1)
						sColDef = Replace(rstFindRecords.Columns(iloop).ColumnName, "_", " ") & "	" & rstFindRecords.Columns(iloop).DataType.ToString.Replace("System.", "")
						Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
					Next
					
					iCount = 0
					For Each objRow As DataRow In rstFindRecords.Rows
						sAddString = ""
						
						For iloop = 0 To (rstFindRecords.Columns.Count - 1)

							If iloop > 0 Then
								sAddString = sAddString & "	"
							End If

							Dim numberAsString As String = objRow(iloop).ToString()
							Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(LocaleDecimalSeparator(), StringComparison.Ordinal)
							Dim numberOfDecimals As Integer = 0
							If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length
							
							If rstFindRecords.Columns(iloop).DataType = GetType(DateTime) Then
								' Field is a date so format as such.
								sAddString = sAddString & ConvertSQLDateToLocale(objRow(iloop))
																
							ElseIf IsDataColumnDecimal(rstFindRecords.Columns(iloop)) Then
								' Field is a numeric so format as such.
								If Not IsDBNull(objRow(iloop)) Then
									If Mid(Session("option1000SepCols"), iloop + 1, 1) = "1" Then
										sTemp = FormatNumber(objRow(iloop), numberOfDecimals, TriState.True, TriState.False, TriState.True)
									Else
										sTemp = FormatNumber(objRow(iloop), numberOfDecimals, TriState.True, TriState.False, TriState.False)
									End If
									
									sAddString = sAddString & sTemp
								End If
							Else
								If Not IsDBNull(objRow(iloop)) Then
									sAddString = sAddString & Replace(objRow(iloop), """", "&quot;")
								End If
							End If
						Next

						Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
						iCount += 1
					Next

					Response.Write("<input type='hidden' id=txtExpectedCount name=txtExpectedCount value=" & prmExpectedCount.Value & ">" & vbCrLf)

				Catch ex As Exception
					sErrorDescription = "Error reading the records." & vbCrLf & FormatError(ex.Message)

				End Try


			ElseIf Session("optionAction") = OptionActionType.SELECTBULKBOOKINGS_2 Then
				
				Try
					objDataAccess.ExecuteSP("sp_ASRIntMakeBulkBookings" _
						, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
						, New SqlParameter("psEmployeeRecordIDs", SqlDbType.VarChar, -1) With {.Value = Session("optionLinkRecordID")} _
						, New SqlParameter("psStatus", SqlDbType.VarChar, -1) With {.Value = Session("optionValue")})
					
					Session("optionAction") = OptionActionType.BULKBOOKINGSUCCESS
					
				Catch ex As Exception
					sNonFatalErrorDescription = "Unable to create booking record." & vbCrLf & sNonFatalErrorDescription
					Session("optionAction") = OptionActionType.BULKBOOKINGERROR
				End Try

				

			ElseIf Session("optionAction") = OptionActionType.LOADEXPRFIELDCOLUMNS Or _
					Session("optionAction") = OptionActionType.LOADEXPRLOOKUPCOLUMNS Then
				
				Try

					Dim prmComponentType = New SqlParameter("piComponentType", SqlDbType.Int)
					If Session("optionAction") = OptionActionType.LOADEXPRFIELDCOLUMNS Then
						prmComponentType.Value = 1
					Else
						prmComponentType.Value = 0
					End If
								
					Dim rstExprColumns = objDataAccess.GetFromSP("sp_ASRIntGetExprColumns" _
							, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
							, prmComponentType _
							, New SqlParameter("piNumericsOnly", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionOnlyNumerics"))})

					iCount = 0
					For Each objRow As DataRow In rstExprColumns.Rows
						iCount += 1
						Response.Write("<input type='hidden' id=txtColumn_" & iCount & " name=txtColumn_" & iCount & " value=""" & objRow("definitionString").ToString() & """>" & vbCrLf)
					Next
					
				Catch ex As Exception
					sErrorDescription = "Error reading component columns." & vbCrLf & FormatError(ex.Message)

				End Try
				

			ElseIf Session("optionAction") = OptionActionType.LOADEXPRLOOKUPVALUES Then
		
				Try
										
					Dim prmDataType = New SqlParameter("piDataType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim rstExprValues = objDataAccess.GetFromSP("sp_ASRIntGetExprLookupValues" _
							, New SqlParameter("piColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionColumnID"))} _
							, prmDataType)

					iCount = 0
					For Each objRow As DataRow In rstExprValues.Rows
						iCount += 1
						Response.Write("<input type='hidden' id=txtValue_" & iCount & " name=txtValue_" & iCount & " value=""" & objRow("lookupValue").ToString & """>" & vbCrLf)
					Next
					
					Response.Write("<input type='hidden' id=txtLookupDataType name=txtLookupDataType value=" & prmDataType.Value.ToString() & ">" & vbCrLf)
					
				Catch ex As Exception
					sErrorDescription = "Error reading component values." & vbCrLf & FormatError(ex.Message)
				End Try
								
			End If

			Response.Write("<input type='hidden' id=txtOptionAction name=txtOptionAction value='" & Session("optionAction") & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionTableID name=txtOptionTableID value='" & CInt(Session("optionTableID")) & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionViewID name=txtOptionViewID value='" & CInt(Session("optionViewID")) & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionOrderID name=txtOptionOrderID value='" & CInt(Session("optionOrderID")) & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionColumnID name=txtOptionColumnID value='" & CInt(Session("optionColumnID")) & "'>" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionLocateValue name=txtOptionLocateValue value=""" & Replace(Session("optionLocateValue"), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
			Response.Write("<input type='hidden' id=txtNonFatalErrorDescription name=txtNonFatalErrorDescription value=""" & sNonFatalErrorDescription & """>")
		%>
</form>


<script type="text/javascript">
	optiondata_onload();
</script>
