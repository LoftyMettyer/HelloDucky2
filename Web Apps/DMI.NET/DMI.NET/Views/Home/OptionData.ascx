<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
	Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
%>
<script src="<%: Url.Content("~/bundles/recordedit")%>" type="text/javascript"></script>

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
</form>

<form id="frmOptionData" name="frmOptionData">
		<%
			Dim aPrompts(1, 0)

			Const adStateOpen = 1

			Const iRETRIES = 5
			Dim iRetryCount As Integer = 0
			' NPG20080904 Fault 13018
			Session("flagOverrideFilter") = False

			Dim objUtilities As HR.Intranet.Server.Utilities

			Dim sErrorDescription As String = ""
			Dim sNonFatalErrorDescription As String = ""

			Dim prmError As ADODB.Parameter
			Dim prmTableID As ADODB.Parameter
			Dim prmViewID As ADODB.Parameter
			Dim prmOrderID As ADODB.Parameter
			Dim prmThousandColumns As SqlParameter
			Dim cmdGetFindRecords As Command
			Dim sThousandColumns As String
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
		
			Dim cmdGetFilterValue As Command
			Dim prmScreenID As ADODB.Parameter
			Dim prmColumnID As ADODB.Parameter
			Dim prmRecordID As ADODB.Parameter
			Dim prmFilterValue As ADODB.Parameter
			Dim prmParentTableID As ADODB.Parameter
			Dim prmParentRecordID As ADODB.Parameter

			Dim prmLookupColumnID As ADODB.Parameter
			Dim prmLookupColumnGridPosition As ADODB.Parameter
			Dim prmOverrideFilter As ADODB.Parameter
		
			Dim prmCourseTitle As ADODB.Parameter
			Dim prmCourseRecordID As ADODB.Parameter
			Dim prmWLRecordID As ADODB.Parameter
			Dim prmEmpRecordID As ADODB.Parameter

			Dim cmdTransferCourse As Command
			Dim cmdBookCourse As Command
			Dim prmStatus As ADODB.Parameter
			Dim fDeadlock As Boolean
			Dim sErrMsg As String
		
			Dim prmTBRecordID As ADODB.Parameter
			Dim prmErrorMessage As ADODB.Parameter

			Dim iCount As Integer
			Dim sAddString As String
			Dim sColDef As String
			Dim sTemp As String
		
			Dim j As Integer
			Dim sPrompts As String
			Dim iIndex1 As Integer
			Dim iIndex2 As Integer
		
			Dim prmSelectionType As ADODB.Parameter
			Dim prmSelectionID As ADODB.Parameter
			Dim prmSelectedIDs As ADODB.Parameter
			Dim prmPromptSQL As ADODB.Parameter
		
			Dim fOK As Boolean

			Dim prmErrMsg As ADODB.Parameter
			Dim cmdPicklist As Command
			Dim prmExpectedCount As ADODB.Parameter
			Dim cmdBulkBook As ADODB.Command
			Dim prmEmployeeRecordIDs As ADODB.Parameter
		
			Response.Write("<INPUT type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)

			' Get the required record count if we have a query.
			'	if len(session("selectSQL")) > 0 then
			If Session("optionAction") = "LOADFIND" Then
				sThousandColumns = ""
			
				Try
					sThousandColumns = Get1000SeparatorFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")))
				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(ex.Message)
				End Try

				cmdGetFindRecords = New Command
				cmdGetFindRecords.CommandText = "sp_ASRIntGetLinkFindRecords"
				cmdGetFindRecords.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
				cmdGetFindRecords.CommandTimeout = 180
			
				prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmTableID)
				prmTableID.Value = CleanNumeric(Session("optionTableID"))

				prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmViewID)
				prmViewID.Value = CleanNumeric(Session("optionViewID"))

				prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmOrderID)
				prmOrderID.Value = CleanNumeric(Session("optionOrderID"))

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
				prmLocateValue.Value = Session("optionLocateValue")

				prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnType)

				prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 100)
				cmdGetFindRecords.Parameters.Append(prmAction)
				prmAction.Value = Session("optionPageAction")

				prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

				prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
				cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
				prmFirstRecPos.Value = CleanNumeric(Session("optionFirstRecPos"))

				prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
				cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
				prmCurrentRecCount.Value = CleanNumeric(Session("optionCurrentRecCount"))

				prmExcludedIDs = cmdGetFindRecords.CreateParameter("excludedIDs", 200, 1, 2147483646)	' 200=varchar, 1=input, 8000=size
				cmdGetFindRecords.Parameters.Append(prmExcludedIDs)
				prmExcludedIDs.Value = ""
		
				prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnSize)

				prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

				Err.Clear()
				rstFindRecords = cmdGetFindRecords.Execute
	
				If (Err.Number <> 0) Then
					sErrorDescription = "Error reading the link find records." & vbCrLf & formatError(Err.Description)
				End If

				If Len(sErrorDescription) = 0 Then
					If rstFindRecords.State = adStateOpen Then
						iCount = 0
						Do While Not rstFindRecords.EOF
							sAddString = ""
						
							For iloop = 0 To (rstFindRecords.Fields.Count - 1)
								If iloop > 0 Then
									sAddString = sAddString & "	"
								End If
							
								If iCount = 0 Then
									sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
									Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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

							Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
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
				If cmdGetFindRecords.Parameters("error").Value <> 0 Then
					'Session("ErrorTitle") = "Link Find Page"
					'Session("ErrorText") = "Error reading link records definition."
					'Response.Clear	  
					'Response.Redirect("error.asp")
				End If

				Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

				cmdGetFindRecords = Nothing
			
			ElseIf Session("optionAction") = "LOADLOOKUPFIND" Then
				' Check if the filter value column is in the current screen.
				' If not, try and get the filter value from the database.
				If Len(Session("optionFilterValue")) = 0 Then
					cmdGetFilterValue = New ADODB.Command
					cmdGetFilterValue.CommandText = "spASRIntGetLookupFilterValue"
					cmdGetFilterValue.CommandType = 4	' Stored procedure
					cmdGetFilterValue.ActiveConnection = Session("databaseConnection")

					prmScreenID = cmdGetFilterValue.CreateParameter("screenID", 3, 1)
					cmdGetFilterValue.Parameters.Append(prmScreenID)
					prmScreenID.Value = CleanNumeric(Session("screenID"))

					prmColumnID = cmdGetFilterValue.CreateParameter("LookupColumnID", 3, 1)
					cmdGetFilterValue.Parameters.Append(prmColumnID)
					prmColumnID.Value = CleanNumeric(Session("optionColumnID"))

					prmTableID = cmdGetFilterValue.CreateParameter("tableID", 3, 1)
					cmdGetFilterValue.Parameters.Append(prmTableID)
					prmTableID.Value = CleanNumeric(Session("tableID"))

					prmViewID = cmdGetFilterValue.CreateParameter("viewID", 3, 1)
					cmdGetFilterValue.Parameters.Append(prmViewID)
					prmViewID.Value = CleanNumeric(Session("viewID"))
				
					prmRecordID = cmdGetFilterValue.CreateParameter("recordID", 3, 1)
					cmdGetFilterValue.Parameters.Append(prmRecordID)
					prmRecordID.Value = CleanNumeric(Session("optionRecordID"))
				
					prmFilterValue = cmdGetFilterValue.CreateParameter("FilterValue", 200, 2, 8000)	' 200=adVarChar, 2=output, 8000=size
					cmdGetFilterValue.Parameters.Append(prmFilterValue)

					prmParentTableID = cmdGetFilterValue.CreateParameter("ParentTableID", 3, 1)
					cmdGetFilterValue.Parameters.Append(prmParentTableID)
					prmParentTableID.Value = CleanNumeric(Session("optionParentTableID"))

					prmParentRecordID = cmdGetFilterValue.CreateParameter("ParentRecordID", 3, 1)
					cmdGetFilterValue.Parameters.Append(prmParentRecordID)
					prmParentRecordID.Value = CleanNumeric(Session("optionParentRecordID"))

					' NPG20080904 Fault 13018
					prmError = cmdGetFilterValue.CreateParameter("Error", 11, 2) ' 11=bit, 2=output
					cmdGetFilterValue.Parameters.Append(prmError)


					Err.Clear()
					cmdGetFilterValue.Execute()

					If (Err.Number <> 0) Then
						sErrorDescription = "Error reading the lookup filter value." & vbCrLf & formatError(Err.Description)
					End If
				
					If Len(sErrorDescription) = 0 Then
						Session("optionFilterValue") = cmdGetFilterValue.Parameters("FilterValue").Value
						Session("flagOverrideFilter") = cmdGetFilterValue.Parameters("Error").Value
						cmdGetFilterValue = Nothing
					End If
				End If
		
				If Len(sErrorDescription) = 0 Then
					sThousandColumns = ""

					If Session("IsLookupTable") = "False" Then
						Try
							sThousandColumns = Get1000SeparatorFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")))
						Catch ex As Exception
							sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(ex.Message)
						End Try

						cmdGetFindRecords = New ADODB.Command
						cmdGetFindRecords.CommandText = "spASRIntGetLookupFindRecords2"
						cmdGetFindRecords.CommandType = 4	' Stored procedure
						cmdGetFindRecords.ActiveConnection = Session("databaseConnection")

						prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
						cmdGetFindRecords.Parameters.Append(prmTableID)
						prmTableID.Value = CleanNumeric(Session("optionTableID"))

						prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
						cmdGetFindRecords.Parameters.Append(prmViewID)
						prmViewID.Value = CleanNumeric(Session("optionViewID"))

						prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
						cmdGetFindRecords.Parameters.Append(prmOrderID)
						prmOrderID.Value = CleanNumeric(Session("optionOrderID"))

						prmColumnID = cmdGetFindRecords.CreateParameter("LookupColumnID", 3, 1)
						cmdGetFindRecords.Parameters.Append(prmColumnID)
						prmColumnID.Value = CleanNumeric(Session("optionLookupColumnID"))

						prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
						cmdGetFindRecords.Parameters.Append(prmReqRecs)
						prmReqRecs.Value = CleanNumeric(Session("FindRecords"))

						prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
						cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

						prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
						cmdGetFindRecords.Parameters.Append(prmIsLastPage)

						prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
						cmdGetFindRecords.Parameters.Append(prmLocateValue)
						prmLocateValue.Value = Session("optionLocateValue")

						prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
						cmdGetFindRecords.Parameters.Append(prmColumnType)

						prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
						cmdGetFindRecords.Parameters.Append(prmColumnSize)

						prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
						cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

						prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
						cmdGetFindRecords.Parameters.Append(prmAction)
						prmAction.Value = Session("optionPageAction")

						prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
						cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

						prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
						cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
						prmFirstRecPos.Value = CleanNumeric(Session("optionFirstRecPos"))

						prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
						cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
						prmCurrentRecCount.Value = CleanNumeric(Session("optionCurrentRecCount"))

						prmFilterValue = cmdGetFindRecords.CreateParameter("FilterValue", 200, 1, 8000)	' 200=adVarChar, 1=input
						cmdGetFindRecords.Parameters.Append(prmFilterValue)
						prmFilterValue.Value = Session("optionFilterValue")
								
						prmLookupColumnID = cmdGetFindRecords.CreateParameter("CallingColumnID", 3, 1) ' 200=adVarChar, 1=input
						cmdGetFindRecords.Parameters.Append(prmLookupColumnID)
						prmLookupColumnID.Value = CleanNumeric(Session("optionColumnID"))

						prmLookupColumnGridPosition = cmdGetFindRecords.CreateParameter("LookupColumnGridPosition", 3, 2)	' 200=adVarChar, 2=output
						cmdGetFindRecords.Parameters.Append(prmLookupColumnGridPosition)
					
						' NPG20080904 Fault 13018
						prmOverrideFilter = cmdGetFindRecords.CreateParameter("pfOverrideFilter", 11, 1) ' 11=bit, 1=input
						cmdGetFindRecords.Parameters.Append(prmOverrideFilter)
						prmOverrideFilter.Value = Session("flagOverrideFilter")
					Else
						prmThousandColumns = New SqlParameter("@ps1000SeparatorCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Try
							objDataAccess.ExecuteSP("spASRIntGetLookupFindColumnInfo", _
													New SqlParameter("@piLookupColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLookupColumnID"))}, _
													prmThousandColumns _
							)
						Catch ex As Exception
							sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(ex.Message)
						End Try

						If Len(sErrorDescription) = 0 Then
							sThousandColumns = prmThousandColumns.Value
						End If
	
						cmdGetFindRecords = New ADODB.Command
						cmdGetFindRecords.CommandText = "spASRIntGetLookupFindRecords"
						cmdGetFindRecords.CommandType = 4	' Stored procedure
						cmdGetFindRecords.ActiveConnection = Session("databaseConnection")

						prmColumnID = cmdGetFindRecords.CreateParameter("LookupColumnID", 3, 1)
						cmdGetFindRecords.Parameters.Append(prmColumnID)
						prmColumnID.Value = CleanNumeric(Session("optionLookupColumnID"))

						prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
						cmdGetFindRecords.Parameters.Append(prmReqRecs)
						prmReqRecs.Value = CleanNumeric(Session("FindRecords"))

						prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
						cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

						prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
						cmdGetFindRecords.Parameters.Append(prmIsLastPage)

						prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
						cmdGetFindRecords.Parameters.Append(prmLocateValue)
						prmLocateValue.Value = Session("optionLocateValue")

						prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
						cmdGetFindRecords.Parameters.Append(prmColumnType)

						prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
						cmdGetFindRecords.Parameters.Append(prmColumnSize)

						prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
						cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

						prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
						cmdGetFindRecords.Parameters.Append(prmAction)
						prmAction.Value = Session("optionPageAction")

						prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
						cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

						prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
						cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
						prmFirstRecPos.Value = CleanNumeric(Session("optionFirstRecPos"))

						prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
						cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
						prmCurrentRecCount.Value = CleanNumeric(Session("optionCurrentRecCount"))

						prmFilterValue = cmdGetFindRecords.CreateParameter("FilterValue", 200, 1, 8000)	' 200=adVarChar, 1=input
						cmdGetFindRecords.Parameters.Append(prmFilterValue)
						prmFilterValue.Value = Session("optionFilterValue")

						prmLookupColumnID = cmdGetFindRecords.CreateParameter("@CallingColumnID", 3, 1)	' 200=adVarChar, 1=input
						cmdGetFindRecords.Parameters.Append(prmLookupColumnID)
						prmLookupColumnID.Value = CleanNumeric(Session("optionColumnID"))
					
						' NPG20080904 Fault 13018					
						prmOverrideFilter = cmdGetFindRecords.CreateParameter("pfOverrideFilter", 11, 1) ' 11=bit, 1=input
						cmdGetFindRecords.Parameters.Append(prmOverrideFilter)
						prmOverrideFilter.Value = Session("flagOverrideFilter")
					End If
					
					Err.Clear()
					rstFindRecords = cmdGetFindRecords.Execute

					If (Err.Number <> 0) Then
						sErrorDescription = "Error reading the lookup find records." & vbCrLf & formatError(Err.Description)
					End If

					If Len(sErrorDescription) = 0 Then
						If rstFindRecords.State = adStateOpen Then
							iCount = 0
							Do While Not rstFindRecords.EOF
								sAddString = ""
							
								For iloop = 0 To (rstFindRecords.Fields.Count - 1)
									If iloop > 0 Then
										sAddString = sAddString & "	"
									End If
								
									If iCount = 0 Then
										sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
										Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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

								Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
						
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
					Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtFilterOverride name=txtFilterOverride value=" & Session("flagOverrideFilter") & ">" & vbCrLf)

					If Session("IsLookupTable") = "False" Then
						Response.Write("<INPUT type='hidden' id=txtLookupColumnGridPosition name=txtLookupColumnGridPosition value=" & cmdGetFindRecords.Parameters("LookupColumnGridPosition").Value & ">" & vbCrLf)
					Else
						Response.Write("<INPUT type='hidden' id=txtLookupColumnGridPosition name=txtLookupColumnGridPosition value=0>" & vbCrLf)
					End If
							
					cmdGetFindRecords = Nothing
				End If
			ElseIf Session("optionAction") = "LOADTRANSFERCOURSE" Then
				sThousandColumns = ""
			
				Try
					sThousandColumns = Get1000SeparatorFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")))
				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(ex.Message)
				End Try

				cmdGetFindRecords = New ADODB.Command
				cmdGetFindRecords.CommandText = "sp_ASRIntGetTransferCourseRecords"
				cmdGetFindRecords.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
				cmdGetFindRecords.CommandTimeout = 180
			
				prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmTableID)
				prmTableID.Value = CleanNumeric(Session("optionTableID"))

				prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmViewID)
				prmViewID.Value = CleanNumeric(Session("optionViewID"))

				prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmOrderID)
				prmOrderID.Value = CleanNumeric(Session("optionOrderID"))
				
				prmCourseTitle = cmdGetFindRecords.CreateParameter("courseTitle", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmCourseTitle)
				prmCourseTitle.Value = Session("optionCourseTitle")

				prmCourseRecordID = cmdGetFindRecords.CreateParameter("courseRecordID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmCourseRecordID)
				prmCourseRecordID.Value = CleanNumeric(Session("optionRecordID"))

				prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmError)

				prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmReqRecs)
				prmReqRecs.Value = CleanNumeric(Session("FindRecords"))

				prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

				prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmIsLastPage)

				prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmLocateValue)
				prmLocateValue.Value = Session("optionLocateValue")

				prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnType)

				prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmAction)
				prmAction.Value = Session("optionPageAction")

				prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

				prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
				cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
				prmFirstRecPos.Value = CleanNumeric(Session("optionFirstRecPos"))

				prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
				cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
				prmCurrentRecCount.Value = CleanNumeric(Session("optionCurrentRecCount"))

				prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnSize)

				prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

				Err.Clear()
				rstFindRecords = cmdGetFindRecords.Execute
	
				If (Err.Number <> 0) Then
					sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
				End If

				If Len(sErrorDescription) = 0 Then
					If rstFindRecords.State = adStateOpen Then
						iCount = 0
						Do While Not rstFindRecords.EOF
							sAddString = ""
						
							For iloop = 0 To (rstFindRecords.Fields.Count - 1)
								If iloop > 0 Then
									sAddString = sAddString & "	"
								End If
							
								If iCount = 0 Then
									sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
									Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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

							Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
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
				If cmdGetFindRecords.Parameters("error").Value <> 0 Then
					'Session("ErrorTitle") = "Transfer Course Find Page"
					'Session("ErrorText") = "Error reading records definition."
					'Response.Clear	  
					'Response.Redirect("error.asp")
				End If

				Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

				cmdGetFindRecords = Nothing

			ElseIf Session("optionAction") = "LOADBOOKCOURSE" Then
				sThousandColumns = ""
			
				Try
					sThousandColumns = Get1000SeparatorFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")))
				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(ex.Message)
				End Try
		
				cmdGetFindRecords = New ADODB.Command
				cmdGetFindRecords.CommandText = "sp_ASRIntGetBookCourseRecords"
				cmdGetFindRecords.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
				cmdGetFindRecords.CommandTimeout = 180
			
				prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmTableID)
				prmTableID.Value = CleanNumeric(Session("optionTableID"))

				prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmViewID)
				prmViewID.Value = CleanNumeric(Session("optionViewID"))

				prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmOrderID)
				prmOrderID.Value = CleanNumeric(Session("optionOrderID"))
				
				prmWLRecordID = cmdGetFindRecords.CreateParameter("WLRecordID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmWLRecordID)
				prmWLRecordID.Value = CleanNumeric(Session("optionRecordID"))

				prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmError)

				prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmReqRecs)
				prmReqRecs.Value = CleanNumeric(Session("FindRecords"))

				prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

				prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmIsLastPage)

				prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmLocateValue)
				prmLocateValue.Value = Session("optionLocateValue")

				prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnType)

				prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmAction)
				prmAction.Value = Session("optionPageAction")

				prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

				prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
				cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
				prmFirstRecPos.Value = CleanNumeric(Session("optionFirstRecPos"))

				prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
				cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
				prmCurrentRecCount.Value = CleanNumeric(Session("optionCurrentRecCount"))

				prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnSize)

				prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

				Err.Clear()
				rstFindRecords = cmdGetFindRecords.Execute
	
				If (Err.Number <> 0) Then
					sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
				End If

				If Len(sErrorDescription) = 0 Then
					If rstFindRecords.State = adStateOpen Then
						iCount = 0
						Do While Not rstFindRecords.EOF
							sAddString = ""
						
							For iloop = 0 To (rstFindRecords.Fields.Count - 1)
								If iloop > 0 Then
									sAddString = sAddString & "	"
								End If
							
								If iCount = 0 Then
									sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
									Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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

							Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
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
				If cmdGetFindRecords.Parameters("error").Value <> 0 Then
					'Session("ErrorTitle") = "Book Course Find Page"
					'Session("ErrorText") = "Error reading records definition."
					'Response.Clear	  
					'Response.Redirect("error.asp")
				End If

				Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)

				cmdGetFindRecords = Nothing

			ElseIf Session("optionAction") = "SELECTBOOKCOURSE_3" Then
						
				cmdBookCourse = New ADODB.Command
				cmdBookCourse.CommandText = "sp_ASRIntBookCourse"
				cmdBookCourse.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdBookCourse.CommandTimeout = 180
				cmdBookCourse.ActiveConnection = Session("databaseConnection")
					
				prmWLRecordID = cmdBookCourse.CreateParameter("WLRecordID", 3, 1)	' 3=integer, 1=input
				cmdBookCourse.Parameters.Append(prmWLRecordID)
				prmWLRecordID.Value = CleanNumeric(Session("optionRecordID"))

				prmCourseRecordID = cmdBookCourse.CreateParameter("CourseRecordID", 3, 1)	' 3=integer, 1=input
				cmdBookCourse.Parameters.Append(prmCourseRecordID)
				prmCourseRecordID.Value = CleanNumeric(Session("optionLinkRecordID"))

				prmStatus = cmdBookCourse.CreateParameter("status", 200, 1, 2147483646)
				cmdBookCourse.Parameters.Append(prmStatus)
				prmStatus.Value = Session("optionValue")

				fDeadlock = True
				Do While fDeadlock
					fDeadlock = False
									
					cmdBookCourse.ActiveConnection.Errors.Clear()
									
					' Run the insert stored procedure.
					cmdBookCourse.Execute()

					If cmdBookCourse.ActiveConnection.Errors.Count > 0 Then
						For iLoop = 1 To cmdBookCourse.ActiveConnection.Errors.Count
							sErrMsg = formatError(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)

							If (cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
									(((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
											(UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
									((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
			(InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
								' The error is for a deadlock.
								' Sorry about having to use the err.description to trap the error but the err.number
								' is not specific and MSDN suggests using the err.description.
								If (iRetryCount < iRETRIES) And (cmdBookCourse.ActiveConnection.Errors.Count = 1) Then
									iRetryCount = iRetryCount + 1
									fDeadlock = True
								Else
									If Len(sNonFatalErrorDescription) > 0 Then
										sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf
									End If
									sNonFatalErrorDescription = sNonFatalErrorDescription & "Another user is deadlocking the database."
									fOK = False
								End If
							ElseIf UCase(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
								'"SQL Mail session is not started."
								'Ignore this error
								'ElseIf (cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
								'	(UCase(Left(cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
								'"EXECUTE permission denied on object 'xp_sendmail'"
								'Ignore this error
					
							Else
								sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf & _
										formatError(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)
								fOK = False
							End If
						Next

						cmdBookCourse.ActiveConnection.Errors.Clear()
												
						If Not fOK Then
							sNonFatalErrorDescription = "The booking could not be made." & vbCrLf & sNonFatalErrorDescription
							Session("optionAction") = "BOOKCOURSEERROR"
						End If
					Else
						Session("optionAction") = "BOOKCOURSESUCCESS"
					End If
				Loop
				cmdBookCourse = Nothing

			ElseIf Session("optionAction") = "SELECTADDFROMWAITINGLIST_3" Then
				cmdBookCourse = New ADODB.Command
				cmdBookCourse.CommandText = "sp_ASRIntAddFromWaitingList"
				cmdBookCourse.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdBookCourse.CommandTimeout = 180
				cmdBookCourse.ActiveConnection = Session("databaseConnection")
					
				prmEmpRecordID = cmdBookCourse.CreateParameter("EmpRecordID", 3, 1)	' 3=integer, 1=input
				cmdBookCourse.Parameters.Append(prmEmpRecordID)
				prmEmpRecordID.Value = CleanNumeric(Session("optionLinkRecordID"))

				prmCourseRecordID = cmdBookCourse.CreateParameter("CourseRecordID", 3, 1)	' 3=integer, 1=input
				cmdBookCourse.Parameters.Append(prmCourseRecordID)
				prmCourseRecordID.Value = CleanNumeric(Session("optionRecordID"))

				prmStatus = cmdBookCourse.CreateParameter("status", 200, 1, 8000)
				cmdBookCourse.Parameters.Append(prmStatus)
				prmStatus.Value = Session("optionValue")

				fDeadlock = True
				Do While fDeadlock
					fDeadlock = False
									
					cmdBookCourse.ActiveConnection.Errors.Clear()
									
					' Run the insert stored procedure.
					cmdBookCourse.Execute()

					If cmdBookCourse.ActiveConnection.Errors.Count > 0 Then
						For iLoop = 1 To cmdBookCourse.ActiveConnection.Errors.Count
							sErrMsg = formatError(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)

							If (cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
									(((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
											(UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
									((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
			(InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
								' The error is for a deadlock.
								' Sorry about having to use the err.description to trap the error but the err.number
								' is not specific and MSDN suggests using the err.description.
								If (iRetryCount < iRETRIES) And (cmdBookCourse.ActiveConnection.Errors.Count = 1) Then
									iRetryCount = iRetryCount + 1
									fDeadlock = True
								Else
									If Len(sNonFatalErrorDescription) > 0 Then
										sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf
									End If
									sNonFatalErrorDescription = sNonFatalErrorDescription & "Another user is deadlocking the database."
									fOK = False
								End If
							ElseIf UCase(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
								'"SQL Mail session is not started."
								'Ignore this error
								'ElseIf (cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
								'	(UCase(Left(cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
								'"EXECUTE permission denied on object 'xp_sendmail'"
								'Ignore this error
					
							Else
								sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf & _
										formatError(cmdBookCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)
								fOK = False
							End If
						Next

						cmdBookCourse.ActiveConnection.Errors.Clear()
												
						If Not fOK Then
							sNonFatalErrorDescription = "The booking could not be made." & vbCrLf & sNonFatalErrorDescription
							Session("optionAction") = "ADDFROMWAITINGLISTERROR"
						End If
					Else
						Session("optionAction") = "ADDFROMWAITINGLISTSUCCESS"
					End If
				Loop
				cmdBookCourse = Nothing

			ElseIf Session("optionAction") = "LOADTRANSFERBOOKING" Then
				sThousandColumns = ""
			
				Try
					sThousandColumns = Get1000SeparatorFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")))
				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(ex.Message)
				End Try

				cmdGetFindRecords = New ADODB.Command
				cmdGetFindRecords.CommandText = "sp_ASRIntGetTransferBookingRecords"
				cmdGetFindRecords.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
				cmdGetFindRecords.CommandTimeout = 180
			
				prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmTableID)
				prmTableID.Value = CleanNumeric(Session("optionTableID"))

				prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmViewID)
				prmViewID.Value = CleanNumeric(Session("optionViewID"))

				prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmOrderID)
				prmOrderID.Value = CleanNumeric(Session("optionOrderID"))
				
				prmTBRecordID = cmdGetFindRecords.CreateParameter("TBRecordID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmTBRecordID)
				prmTBRecordID.Value = CleanNumeric(Session("optionRecordID"))

				prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmError)

				prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmReqRecs)
				prmReqRecs.Value = CleanNumeric(Session("FindRecords"))

				prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

				prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmIsLastPage)

				prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmLocateValue)
				prmLocateValue.Value = Session("optionLocateValue")

				prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnType)

				prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmAction)
				prmAction.Value = Session("optionPageAction")

				prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

				prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
				cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
				prmFirstRecPos.Value = CleanNumeric(Session("optionFirstRecPos"))

				prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
				cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
				prmCurrentRecCount.Value = CleanNumeric(Session("optionCurrentRecCount"))

				prmErrorMessage = cmdGetFindRecords.CreateParameter("errorMessage", 200, 2, 8000)	' 200=varchar, 2=output,8000=size
				cmdGetFindRecords.Parameters.Append(prmErrorMessage)

				prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnSize)

				prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

				prmStatus = cmdGetFindRecords.CreateParameter("status", 200, 2, 8000)	' 200=varchar, 2=output,8000=size
				cmdGetFindRecords.Parameters.Append(prmStatus)

				Err.Clear()
				rstFindRecords = cmdGetFindRecords.Execute
	
				If (Err.Number <> 0) Then
					sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
				End If

				
				If Len(sErrorDescription) = 0 Then
					If rstFindRecords.State = adStateOpen Then
						iCount = 0
						Do While Not rstFindRecords.EOF
							sAddString = ""
						
							For iloop = 0 To (rstFindRecords.Fields.Count - 1)
								If iloop > 0 Then
									sAddString = sAddString & "	"
								End If
							
								If iCount = 0 Then
									sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
									Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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

							Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
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
				If cmdGetFindRecords.Parameters("error").Value <> 0 Then
					'Session("ErrorTitle") = "Book Course Find Page"
					'Session("ErrorText") = "Error reading records definition."
					'Response.Clear	  
					'Response.Redirect("error.asp")
				End If

				Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtErrorMessage2 name=txtErrorMessage2 value=""" & Replace(cmdGetFindRecords.Parameters("errorMessage").Value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtStatus name=txtStatus value=""" & Replace(cmdGetFindRecords.Parameters("status").Value, """", "&quot;") & """>" & vbCrLf)
			
				cmdGetFindRecords = Nothing

			ElseIf Session("optionAction") = "LOADADDFROMWAITINGLIST" Then
				sThousandColumns = ""
			
				Try
					sThousandColumns = Get1000SeparatorFindColumns(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")), CleanNumeric(Session("optionOrderID")))
				Catch ex As Exception
					sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(ex.Message)
				End Try

				cmdGetFindRecords = New ADODB.Command
				cmdGetFindRecords.CommandText = "sp_ASRIntGetAddFromWaitingListRecords"
				cmdGetFindRecords.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdGetFindRecords.ActiveConnection = Session("databaseConnection")
				cmdGetFindRecords.CommandTimeout = 180
			
				prmTableID = cmdGetFindRecords.CreateParameter("tableID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmTableID)
				prmTableID.Value = CleanNumeric(Session("optionTableID"))

				prmViewID = cmdGetFindRecords.CreateParameter("viewID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmViewID)
				prmViewID.Value = CleanNumeric(Session("optionViewID"))

				prmOrderID = cmdGetFindRecords.CreateParameter("orderID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmOrderID)
				prmOrderID.Value = CleanNumeric(Session("optionOrderID"))

				prmCourseRecordID = cmdGetFindRecords.CreateParameter("CourseRecordID", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmCourseRecordID)
				prmCourseRecordID.Value = CleanNumeric(Session("optionRecordID"))

				prmError = cmdGetFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmError)

				prmReqRecs = cmdGetFindRecords.CreateParameter("reqRecs", 3, 1)
				cmdGetFindRecords.Parameters.Append(prmReqRecs)
				prmReqRecs.Value = CleanNumeric(Session("FindRecords"))

				prmIsFirstPage = cmdGetFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmIsFirstPage)

				prmIsLastPage = cmdGetFindRecords.CreateParameter("isLastPage", 11, 2) ' 11=bit, 2=output
				cmdGetFindRecords.Parameters.Append(prmIsLastPage)

				prmLocateValue = cmdGetFindRecords.CreateParameter("locateValue", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmLocateValue)
				prmLocateValue.Value = Session("optionLocateValue")

				prmColumnType = cmdGetFindRecords.CreateParameter("columnType", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnType)

				prmAction = cmdGetFindRecords.CreateParameter("action", 200, 1, 8000)
				cmdGetFindRecords.Parameters.Append(prmAction)
				prmAction.Value = Session("optionPageAction")

				prmTotalRecCount = cmdGetFindRecords.CreateParameter("totalRecCount", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmTotalRecCount)

				prmFirstRecPos = cmdGetFindRecords.CreateParameter("firstRecPos", 3, 3)	' 3=integer, 3=input/output
				cmdGetFindRecords.Parameters.Append(prmFirstRecPos)
				prmFirstRecPos.Value = CleanNumeric(Session("optionFirstRecPos"))

				prmCurrentRecCount = cmdGetFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
				cmdGetFindRecords.Parameters.Append(prmCurrentRecCount)
				prmCurrentRecCount.Value = CleanNumeric(Session("optionCurrentRecCount"))

				prmErrorMessage = cmdGetFindRecords.CreateParameter("errorMessage", 200, 2, 8000)	' 200=varchar, 2=output,8000=size
				cmdGetFindRecords.Parameters.Append(prmErrorMessage)

				prmColumnSize = cmdGetFindRecords.CreateParameter("columnSize", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnSize)

				prmColumnDecimals = cmdGetFindRecords.CreateParameter("columnDecimals", 3, 2)	' 3=integer, 2=output
				cmdGetFindRecords.Parameters.Append(prmColumnDecimals)

				Err.Clear()
				rstFindRecords = cmdGetFindRecords.Execute
	
				If (Err.Number <> 0) Then
					sErrorDescription = "Error reading the find records." & vbCrLf & formatError(Err.Description)
				End If

				If Len(sErrorDescription) = 0 Then
					If rstFindRecords.State = adStateOpen Then
						iCount = 0
						Do While Not rstFindRecords.EOF
							sAddString = ""
						
							For iloop = 0 To (rstFindRecords.Fields.Count - 1)
								If iloop > 0 Then
									sAddString = sAddString & "	"
								End If
							
								If iCount = 0 Then
									sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
									Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
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

							Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
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
				If cmdGetFindRecords.Parameters("error").Value <> 0 Then
					'Session("ErrorTitle") = "Add From Waiting List Find Page"
					'Session("ErrorText") = "Error reading records definition."
					'Response.Clear	  
					'Response.Redirect("error.asp")
				End If

				Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdGetFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdGetFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdGetFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdGetFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdGetFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=0>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtErrorMessage2 name=txtErrorMessage2 value=""" & Replace(cmdGetFindRecords.Parameters("errorMessage").Value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdGetFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdGetFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)
			
				cmdGetFindRecords = Nothing

			ElseIf Session("optionAction") = "SELECTTRANSFERBOOKING_2" Then
		
				
				
				cmdTransferCourse = New ADODB.Command
				cmdTransferCourse.CommandText = "sp_ASRIntTransferCourse"
				cmdTransferCourse.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdTransferCourse.CommandTimeout = 180
				cmdTransferCourse.ActiveConnection = Session("databaseConnection")
					
				prmTBRecordID = cmdTransferCourse.CreateParameter("TBRecordID", 3, 1)	' 3=integer, 1=input
				cmdTransferCourse.Parameters.Append(prmTBRecordID)
				prmTBRecordID.Value = CleanNumeric(Session("optionRecordID"))

				prmCourseRecordID = cmdTransferCourse.CreateParameter("CourseRecordID", 3, 1)	' 3=integer, 1=input
				cmdTransferCourse.Parameters.Append(prmCourseRecordID)
				prmCourseRecordID.Value = CleanNumeric(Session("optionLinkRecordID"))

				fDeadlock = True
				Do While fDeadlock
					fDeadlock = False
									
					cmdTransferCourse.ActiveConnection.Errors.Clear()
									
					' Run the insert stored procedure.
					cmdTransferCourse.Execute()

					If cmdTransferCourse.ActiveConnection.Errors.Count > 0 Then
						For iLoop = 1 To cmdTransferCourse.ActiveConnection.Errors.Count
							sErrMsg = formatError(cmdTransferCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)

							If (cmdTransferCourse.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
									(((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
											(UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
									((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
			(InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
								' The error is for a deadlock.
								' Sorry about having to use the err.description to trap the error but the err.number
								' is not specific and MSDN suggests using the err.description.
								If (iRetryCount < iRETRIES) And (cmdTransferCourse.ActiveConnection.Errors.Count = 1) Then
									iRetryCount = iRetryCount + 1
									fDeadlock = True
								Else
									If Len(sNonFatalErrorDescription) > 0 Then
										sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf
									End If
									sNonFatalErrorDescription = sNonFatalErrorDescription & "Another user is deadlocking the database."
									fOK = False
								End If
							ElseIf UCase(cmdTransferCourse.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
								'"SQL Mail session is not started."
								'Ignore this error
								'ElseIf (cmdTransferCourse.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
								'	(UCase(Left(cmdTransferCourse.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
								'"EXECUTE permission denied on object 'xp_sendmail'"
								'Ignore this error
					
							Else
								sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf & _
										formatError(cmdTransferCourse.ActiveConnection.Errors.Item(iLoop - 1).Description)
								fOK = False
							End If
						Next

						cmdTransferCourse.ActiveConnection.Errors.Clear()
												
						If Not fOK Then
							sNonFatalErrorDescription = "The booking could not be transferred." & vbCrLf & sNonFatalErrorDescription
							Session("optionAction") = "TRANSFERBOOKINGERROR"
						End If
					Else
						Session("optionAction") = "TRANSFERBOOKINGSUCCESS"
					End If
				Loop
				cmdTransferCourse = Nothing

			ElseIf Session("optionAction") = "GETBULKBOOKINGSELECTION" Then
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
								sAddString = sAddString & convertSQLDateToLocale(objRow(iloop))
							ElseIf rstFindRecords2.Columns(iloop).DataType.Name.ToLower() = "system.decimal" Then
								' Field is a numeric so format as such.
								If Not IsDBNull(objRow(iloop)) Then
									If Mid(Session("option1000SepCols"), iloop + 1, 1) = "1" Then
										sTemp = FormatNumber(objRow(iloop), objRow(iloop).NumericScale, True, False, True)
									Else
										sTemp = FormatNumber(objRow(iloop), objRow(iloop).NumericScale, True, False, False)
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

						Response.Write("<input type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
						iCount += 1
					Next
					
				Catch ex As Exception
					sErrorDescription = "Error reading the find records." & vbCrLf & formatError(ex.Message)
					
				End Try
			
				
			ElseIf Session("optionAction") = "GETPICKLISTSELECTION" Then
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
								
				cmdPicklist = New ADODB.Command()
				cmdPicklist.CommandText = "sp_ASRIntGetSelectedPicklistRecords"
				cmdPicklist.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdPicklist.CommandTimeout = 180
				cmdPicklist.ActiveConnection = Session("databaseConnection")

				prmSelectionType = cmdPicklist.CreateParameter("selectionType", 200, 1, 8000)	'200=varchar,1=input,8000=size
				cmdPicklist.Parameters.Append(prmSelectionType)
				prmSelectionType.Value = Session("optionPageAction")

				prmSelectionID = cmdPicklist.CreateParameter("selectionID", 3, 1)	'3=integer,1=input
				cmdPicklist.Parameters.Append(prmSelectionID)
				prmSelectionID.Value = CleanNumeric(Session("optionRecordID"))
			
				prmSelectedIDs = cmdPicklist.CreateParameter("selectedIDs", 200, 1, 2147483646)	'200=varchar,1=input,8000=size
				cmdPicklist.Parameters.Append(prmSelectedIDs)
				prmSelectedIDs.Value = Session("optionValue")

				prmPromptSQL = cmdPicklist.CreateParameter("promptSQL", 200, 1, 2147483646)	'200=varchar,1=input,8000=size
				cmdPicklist.Parameters.Append(prmPromptSQL)
				If Len(Session("optionPromptSQL")) = 0 Then
					prmPromptSQL.Value = ""
				Else
					prmPromptSQL.Value = Session("optionPromptSQL")
				End If
						
				prmTableID = cmdPicklist.CreateParameter("tableID", 3, 1)	'3=integer,1=input
				cmdPicklist.Parameters.Append(prmTableID)
				prmTableID.Value = CleanNumeric(Session("optionTableID"))

				prmErrMsg = cmdPicklist.CreateParameter("errMsg", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamOutput, 2147483646)
				cmdPicklist.Parameters.Append(prmErrMsg)

				prmExpectedCount = cmdPicklist.CreateParameter("expectedCount", 3, 2)	'3=integer,2=output
				cmdPicklist.Parameters.Append(prmExpectedCount)

				objUtilities = Session("UtilitiesObject")

				objUtilities.UDFFunctions(True)
		
				Err.Clear()
				rstFindRecords = cmdPicklist.Execute

				objUtilities.UDFFunctions(False)
			
				objUtilities = Nothing
	
				If (Err.Number <> 0) Then
					sErrorDescription = "Error reading the records." & vbCrLf & formatError(Err.Description)
				End If
			
				If Len(sErrorDescription) = 0 Then
					If rstFindRecords.State = adStateOpen Then
						iCount = 0
						Do While Not rstFindRecords.EOF
							sAddString = ""
						
							For iloop = 0 To (rstFindRecords.Fields.Count - 1)
								If iloop > 0 Then
									sAddString = sAddString & "	"
								End If
							
								If iCount = 0 Then
									sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
									Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iloop & " name=txtOptionColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
								End If
							
								If rstFindRecords.Fields(iloop).Type = 135 Then
									' Field is a date so format as such.
									sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
								ElseIf rstFindRecords.Fields(iloop).Type = 131 Then
									' Field is a numeric so format as such.
									If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
										If Mid(Session("option1000SepCols"), iloop + 1, 1) = "1" Then
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

							Response.Write("<INPUT type='hidden' id=txtOptionData_" & iCount & " name=txtOptionData_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
							iCount = iCount + 1
							rstFindRecords.MoveNext()
						Loop
	
						' Release the ADO recordset object.
						rstFindRecords.Close()
					End If

				End If
				rstFindRecords = Nothing

				Response.Write("<INPUT type='hidden' id=txtExpectedCount name=txtExpectedCount value=" & cmdPicklist.Parameters("expectedCount").Value & ">" & vbCrLf)
			
				cmdPicklist = Nothing

			ElseIf Session("optionAction") = "SELECTBULKBOOKINGS_2" Then
				cmdBulkBook = New ADODB.Command
				cmdBulkBook.CommandText = "sp_ASRIntMakeBulkBookings"
				cmdBulkBook.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdBulkBook.CommandTimeout = 180
				cmdBulkBook.ActiveConnection = Session("databaseConnection")
					
				prmCourseRecordID = cmdBulkBook.CreateParameter("CourseRecordID", 3, 1)	' 3=integer, 1=input
				cmdBulkBook.Parameters.Append(prmCourseRecordID)
				prmCourseRecordID.Value = CleanNumeric(Session("optionRecordID"))

				prmEmployeeRecordIDs = cmdBulkBook.CreateParameter("EmployeeRecordIDs", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdBulkBook.Parameters.Append(prmEmployeeRecordIDs)
				prmEmployeeRecordIDs.Value = Session("optionLinkRecordID")

				prmStatus = cmdBulkBook.CreateParameter("Status", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdBulkBook.Parameters.Append(prmStatus)
				prmStatus.Value = Session("optionValue")

				fDeadlock = True
				Do While fDeadlock
					fDeadlock = False
									
					cmdBulkBook.ActiveConnection.Errors.Clear()
									
					' Run the insert stored procedure.
					cmdBulkBook.Execute()

					If cmdBulkBook.ActiveConnection.Errors.Count > 0 Then
						For iLoop = 1 To cmdBulkBook.ActiveConnection.Errors.Count
							sErrMsg = formatError(cmdBulkBook.ActiveConnection.Errors.Item(iLoop - 1).Description)

							If (cmdBulkBook.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
									(((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
											(UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
									((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
			(InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
								' The error is for a deadlock.
								' Sorry about having to use the err.description to trap the error but the err.number
								' is not specific and MSDN suggests using the err.description.
								If (iRetryCount < iRETRIES) And (cmdBulkBook.ActiveConnection.Errors.Count = 1) Then
									iRetryCount = iRetryCount + 1
									fDeadlock = True
								Else
									If Len(sNonFatalErrorDescription) > 0 Then
										sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf
									End If
									sNonFatalErrorDescription = sNonFatalErrorDescription & "Another user is deadlocking the database."
									fOK = False
								End If
							ElseIf UCase(cmdBulkBook.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
								'"SQL Mail session is not started."
								'Ignore this error
								'ElseIf (cmdTransferCourse.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
								'	(UCase(Left(cmdTransferCourse.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
								'"EXECUTE permission denied on object 'xp_sendmail'"
								'Ignore this error
					
							Else
								sNonFatalErrorDescription = sNonFatalErrorDescription & vbCrLf & _
										formatError(cmdBulkBook.ActiveConnection.Errors.Item(iLoop - 1).Description)
								fOK = False
							End If
						Next

						cmdBulkBook.ActiveConnection.Errors.Clear()
												
						If Not fOK Then
							sNonFatalErrorDescription = "Unable to create booking record." & vbCrLf & sNonFatalErrorDescription
							Session("optionAction") = "BULKBOOKINGERROR"
						End If
					Else
						Session("optionAction") = "BULKBOOKINGSUCCESS"
					End If
				Loop
				cmdBulkBook = Nothing

			ElseIf (Session("optionAction") = "LOADEXPRFIELDCOLUMNS") Or _
					(Session("optionAction") = "LOADEXPRLOOKUPCOLUMNS") Then
				
				Try

					Dim prmComponentType = New SqlParameter("piComponentType", SqlDbType.Int)
					If Session("optionAction") = "LOADEXPRFIELDCOLUMNS" Then
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
					sErrorDescription = "Error reading component columns." & vbCrLf & formatError(ex.Message)

				End Try
				

			ElseIf Session("optionAction") = "LOADEXPRLOOKUPVALUES" Then
		
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
					sErrorDescription = "Error reading component values." & vbCrLf & formatError(ex.Message)
				End Try
								
			End If

			Response.Write("<input type='hidden' id=txtOptionAction name=txtOptionAction value=" & Session("optionAction") & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionTableID name=txtOptionTableID value=" & Session("optionTableID") & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionViewID name=txtOptionViewID value=" & Session("optionViewID") & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionOrderID name=txtOptionOrderID value=" & Session("optionOrderID") & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionColumnID name=txtOptionColumnID value=" & Session("optionColumnID") & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtOptionLocateValue name=txtOptionLocateValue value=""" & Replace(Session("optionLocateValue"), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
			Response.Write("<input type='hidden' id=txtNonFatalErrorDescription name=txtNonFatalErrorDescription value=""" & sNonFatalErrorDescription & """>")
		%>
</form>

<script runat="server" language="vb">

	Function formatError(psErrMsg As String) As String
		Dim iStart As Integer
		Dim iFound As Integer
	
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
				Dim sLocaleFormat As String
				Dim iIndex As Integer
	
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
		optiondata_onload()
</script>
