Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Namespace Expressions
	Friend Class clsExprFunction
		Inherits BaseExpressionComponent

		' Component definition variables.
		Private mlngFunctionID As Integer
		Private msFunctionName As String
		Private miReturnType As ExpressionValueTypes

		' Definition for expanded/unexpanded status of the component
		Private mbExpanded As Boolean

		' Class handling variables.
		Private mobjBaseComponent As clsExprComponent
		Private mcolParameters As Collection
		Private mobjBadComponent As clsExprComponent


		Public Function ContainsExpression(plngExprID As Integer) As Boolean
			' Retrun TRUE if the current expression (or any of its sub expressions)
			' contains the given expression. This ensures no cyclic expressions get created.

			Dim objParameter As clsExprComponent
			Dim objSubExpression As clsExprExpression

			Dim bContainsExpression = False

			Try

				For Each objParameter In mcolParameters
					objSubExpression = objParameter.Component

					bContainsExpression = objSubExpression.ContainsExpression(plngExprID)

					If bContainsExpression Then
						Exit For
					End If
				Next objParameter

			Catch ex As Exception
				Return True

			End Try

			Return bContainsExpression

		End Function

		Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, pfApplyPermissions As Boolean _
																, pfValidating As Boolean, ByRef pavPromptedValues As Object _
																, ByRef psUDFs() As String _
																, Optional plngFixedExprID As Integer = 0, Optional psFixedSQLCode As String = "") As Boolean

			' Return the SQL code for the component.
			Dim fOK As Boolean = True
			Dim fFound As Boolean
			Dim fSrchColumnOK As Boolean
			Dim fRtnColumnOK As Boolean
			Dim iLoop As Integer
			Dim lngSrchTableID As Integer
			Dim lngRtnTableID As Integer
			Dim sCode As String = ""
			Dim sSQL As String
			Dim sRtnColumnCode As String
			Dim sSrchColumnCode As String
			Dim sSrchTableCode As String
			Dim sRealTableSource As String
			Dim sParamCode1 As String = ""
			Dim sParamCode2 As String = ""
			Dim sParamCode3 As String = ""
			Dim sParamCode4 As String = ""
			Dim sSrchColumnName As String = ""
			Dim sRtnColumnName As String = ""
			Dim sSrchTableName As String
			Dim rsInfo As DataTable
			Dim objColumnPrivileges As CColumnPrivileges
			Dim objTableView As TablePrivilege
			Dim asViews(,) As String
			Dim strRemainString As String
			Dim strTempTableName As String
			Dim strTempTableID As String
			Dim objBaseTable As TablePrivilege

			'Currency Conversion Values
			Dim sCConvTable As String
			Dim sCConvExRateCol As String
			Dim sCConvCurrDescCol As String
			Dim sCConvDecCol As String

			Try


				' Get the first parameter's runtime code if required.
				If mcolParameters.Count() >= 1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fOK = mcolParameters.Item(1).Component.RuntimeCode(sParamCode1, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs)
				End If

				' Get the second parameter's runtime code if required.
				If fOK And (mcolParameters.Count() >= 2) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fOK = mcolParameters.Item(2).Component.RuntimeCode(sParamCode2, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs)
				End If

				' Get the third parameter's runtime code if required.
				If fOK And (mcolParameters.Count() >= 3) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fOK = mcolParameters.Item(3).Component.RuntimeCode(sParamCode3, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs)
				End If

				' Get the fourth parameter's runtime code if required.
				If fOK And (mcolParameters.Count() >= 4) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fOK = mcolParameters.Item(4).Component.RuntimeCode(sParamCode4, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs)
				End If

				If fOK Then
					Select Case mlngFunctionID
						Case 1 ' System date
							sCode = "GETDATE()"

						Case 2 ' Convert to uppercase
							sCode = String.Format("UPPER({0})", sParamCode1)

						Case 3 ' Convert numeric to string
							sCode = String.Format("ISNULL(LTRIM(STR({0}, 255, {1})),'')", sParamCode1, sParamCode2)

						Case 4 ' If... Then... Else...
							sCode = String.Format("CASE WHEN ({0} = 1) THEN {1} ELSE {2} END", sParamCode1, sParamCode2, sParamCode3)

						Case 5 ' Remove leading and trailing spaces
							sCode = "ltrim(rtrim(" & sParamCode1 & "))"

						Case 6 ' Extract characters from the left
							sCode = "left(" & sParamCode1 & ", " & sParamCode2 & ")"

						Case 7 ' Length of character field
							sCode = "len(" & sParamCode1 & ")"

						Case 8 ' Convert to lowercase
							sCode = String.Format("LOWER({0})", sParamCode1)

						Case 9 ' Maximum
							sCode = String.Format("CASE WHEN ({0} > {1}) THEN {0} ELSE {1} END", sParamCode1, sParamCode2)

						Case 10	' Minimum
							sCode = String.Format("CASE WHEN ({0} < {1}) THEN {0} ELSE {1} END", sParamCode1, sParamCode2)

						Case 11	' Search for character string.
							sCode = "charindex(" & sParamCode2 & ", " & sParamCode1 & ")"

						Case 12	' Capitalise Initials
							sCode = "(dbo.udf_ASRFn_CapitalizeInitials(" & sParamCode1 & "))"

						Case 13	' Extract characters from the right
							sCode = "right(" & sParamCode1 & ", " & sParamCode2 & ")"

						Case 14	' Extract part of a character string
							sCode = "substring(" & sParamCode1 & ", " & sParamCode2 & ", " & sParamCode3 & ")"

						Case 15	' System Time
							sCode = "convert(varchar(50), getdate(), 8)"

						Case 16	' Is field empty
							sCode = "(CASE WHEN ((" & sParamCode1 & ") IS NULL)"

							' Validate the sub-expression. This is done, not to  validate the expression,
							' but rather to determine the return type of the expression.
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mcolParameters.Item(1).Component.ValidateExpression(False)

							'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters(1).ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Select Case mcolParameters.Item(1).ReturnType
								Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
									sCode = sCode & " OR ((" & sParamCode1 & ") = '')"
								Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
									sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
								Case ExpressionValueTypes.giEXPRVALUE_LOGIC
									sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
							End Select

							sCode = sCode & " THEN 1 ELSE 0 END)"

						Case 17	' Current user
							sCode = "SYSTEM_USER"

						Case 18	' Whole Years Until Current Date
							sCode = String.Format("datediff(year,{0}, getdate()) - case when datepart(month,{0}) > datepart(month, getdate()) then 1 when (datepart(month,{0}) = datepart(month, getdate())) and (datepart(day,{0}) > datepart(day, getdate())) then 1 else 0 end", sParamCode1)

						Case 19	' Remaining Months Since Whole Years
							sCode = String.Format("datepart(month, getdate()) - datepart(month, {0}) - case when datepart(day,{0}) > datepart(day, getdate()) then 1 else 0 end + case when (datepart(month, getdate()) - datepart(month,{0}) - case when datepart(day,{0}) > datepart(day, getdate()) then 1 else 0 end) < 0 then 12 else 0 end", sParamCode1)

						Case 20	' Capitalise Initials
							sCode = "(dbo.udf_ASRFn_InitialsFromForenames(" & sParamCode1 & "))"

						Case 21	' First Name from Forenames
							sCode = "case when charindex(' ', ltrim(" & sParamCode1 & ")) > 0 then substring(ltrim(" & sParamCode1 & "), 1, charindex(' ', ltrim(" & sParamCode1 & "))-1)" & "    else ltrim(" & sParamCode1 & ")" & "end"

						Case 22	' Weekdays From Start and End Dates
							sCode = String.Format(" case when datediff(day, {0}, {1}) < 0 then 0 else datediff(day, {0}, {1} + 1) - (2 * (datediff(day, {0} - (datepart(dw, {0}) - 1), {1} - (datepart(dw, {1}) - 1)) / 7)) - case when datepart(dw, {0}) = 1 then 1 else 0 end - case when datepart(dw, {1}) = 7 then 1 else 0 end end", sParamCode1, sParamCode2)

						Case 23	' Add months to date
							sCode = "dateadd(month, " & sParamCode2 & ", " & sParamCode1 & ")"

						Case 24	' Add years to date
							sCode = "dateadd(year, " & sParamCode2 & ", " & sParamCode1 & ")"

						Case 25	' Convert character to numeric.
							sCode = " case when isnumeric(" & sParamCode1 & ") = 1 then convert(float, convert(money, " & sParamCode1 & "))  else 0 end"

						Case 26	' Whole Months between 2 Dates.
							sCode = " case when " & sParamCode1 & " >= " & sParamCode2 & " then 0 else datediff(month, " & sParamCode1 & ", " & sParamCode2 & ") - case when datepart(day, " & sParamCode2 & ") < datepart(day, " & sParamCode1 & ") then 1 else 0 end end"

						Case 27	' Parentheses
							sCode = sParamCode1

						Case 28	' Day of the week
							sCode = "DATEPART(weekday, " & sParamCode1 & ")"

						Case 29	' Working Days per week
							sCode = "(convert(float, len(replace(left(" & sParamCode1 & ", 14), ' ', ''))) / 2)"

						Case 30	' Absence Duration
							'TM08102003
							If pfValidating Then
								strTempTableName = mobjBaseComponent.ParentExpression.BaseTableName
							Else
								strTempTableName = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName).RealSource
							End If

							If mobjBaseComponent.ParentExpression.BaseTableID = PersonnelModule.glngPersonnelTableID Then
								strTempTableID = "ID"
							Else
								strTempTableID = "ID_" & Trim(Str(PersonnelModule.glngPersonnelTableID))
							End If

							sCode = "(dbo.udf_ASRFn_AbsenceDuration(" & sParamCode1 & "," & sParamCode2 & "," & sParamCode3 & "," & sParamCode4 & "," & strTempTableName & "." & strTempTableID & "))"

						Case 31	' Round down to nearest whole number.
							sCode = String.Format("ROUND({0}, 0, 1)", sParamCode1)

						Case 32	' Year of date.
							sCode = String.Format("DATEPART(year, {0})", sParamCode1)

						Case 33	' Month of date.
							sCode = String.Format("DATEPART(month, {0})", sParamCode1)

						Case 34	' Day of date.
							sCode = String.Format("DATEPART(day, {0})", sParamCode1)

						Case 35	' Nice Date
							sCode = "datename(day, " & sParamCode1 & ") + ' ' + " & "datename(month, " & sParamCode1 & ") + ' ' + " & "datename(year, " & sParamCode1 & ")"

						Case 36	' Nice Time
							sCode = "case when len(ltrim(rtrim(" & sParamCode1 & "))) = 0 then '' else case when isdate(" & sParamCode1 & ") = 0 then '***'" & " else (convert(varchar(2),((datepart(hour,convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)) + 11) % 12) + 1)" & " + ':' + right('00' + datename(minute, convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)),2)" & " + case when datepart(hour, convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)) > 11 then ' pm'" & " else ' am' end) end end"

						Case 37	' Round Date to Start of nearest month
							sCode = " case when datediff(day, (" & sParamCode1 & " - datepart(day, " & sParamCode1 & ") + 1), " & sParamCode1 & ")" & "         <= datediff(day, " & sParamCode1 & ", (dateadd(month, 1, " & sParamCode1 & ") - datepart(day, dateadd(month, 1, " & sParamCode1 & ")) + 1))" & "         then " & sParamCode1 & " - datepart(day, " & sParamCode1 & ") + 1" & "     else dateadd(month, 1, " & sParamCode1 & ")" & "         - datepart(day, dateadd(month, 1, " & sParamCode1 & ")) + 1" & " end"

						Case 38	' Is Between
							sCode = " case when (" & sParamCode1 & " >= " & sParamCode2 & ") AND (" & sParamCode1 & " <= " & sParamCode3 & ") then 1 else 0 end"

						Case 39	' Service Years
							sCode = "   datepart(year, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)" & " - datepart(year, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & " - case" & "       when datepart(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & "          > datepart(month, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end) then 1" & "      when (datepart(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & "          = datepart(month, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end))" & "       and (datepart(day, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & "          > datepart(day, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)) then 1" & "      else 0" & "  end"

						Case 40	' Service Months
							sCode = " (case" & "    when case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end" & "          >= case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end then 0" & "    else datediff(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)" & "        - case" & "              when datepart(day, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end) < datepart(day, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end) then 1" & "              else 0" & "          end" & " end) % 12"

						Case 42	' Get field from database record.
							' Get the column parameter definitions.
							Dim objColumnParameter1 = Columns.GetById(CInt(sParamCode1))
							sSrchColumnName = objColumnParameter1.Name
							sSrchTableName = objColumnParameter1.TableName
							lngSrchTableID = objColumnParameter1.TableID

							Dim objColumnParameter2 = Columns.GetById(CInt(sParamCode3))
							sRtnColumnName = objColumnParameter2.Name
							lngRtnTableID = objColumnParameter2.TableID

							fOK = ((Len(sSrchColumnName) > 0) And (Len(sRtnColumnName) > 0)) And (lngSrchTableID = lngRtnTableID)

							' Construct the select statement to get the required field from the given table,
							' incorporating permissions.
							If fOK Then
								ReDim asViews(2, 0)

								' Check permissions on that return column.
								objColumnPrivileges = GetColumnPrivileges(sSrchTableName)
								sRealTableSource = gcoTablePrivileges.Item(sSrchTableName).RealSource

								fRtnColumnOK = objColumnPrivileges.IsValid(sRtnColumnName)

								If fRtnColumnOK Then
									fRtnColumnOK = objColumnPrivileges.Item(sRtnColumnName).AllowSelect
								End If

								If fRtnColumnOK Then
									' The search column can be read direct from the table.
									sRtnColumnCode = "lookup." & sRtnColumnName
								Else
									' Then column cannot be read direct. If its from a parent, try parent views
									' Loop thru the views on the table, seeing if any have read permis for the column
									' Column 1 = view name
									' Column 2 = "S" if the view is used for the search.
									For Each objTableView In gcoTablePrivileges.Collection
										If (Not objTableView.IsTable) And (objTableView.TableID = lngSrchTableID) And (objTableView.AllowSelect) Then

											' Get the column permission for the view
											objColumnPrivileges = GetColumnPrivileges((objTableView.ViewName))

											' If we can see the column from this view
											If objColumnPrivileges.IsValid(sRtnColumnName) Then
												If objColumnPrivileges.Item(sRtnColumnName).AllowSelect Then

													ReDim Preserve asViews(2, UBound(asViews, 2) + 1)
													asViews(1, UBound(asViews, 2)) = objTableView.ViewName
													asViews(2, UBound(asViews, 2)) = ""
												End If
											End If
										End If
									Next objTableView
									'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
									objTableView = Nothing

									' Does the user have select permission thru ANY views ?
									If UBound(asViews, 2) = 0 Then
										' The search column can be read neither from the table nor any views.
										sRtnColumnCode = "NULL"
									Else
										For iLoop = 1 To UBound(asViews, 2)
											If iLoop = 1 Then
												sRtnColumnCode = "CASE"
											End If

											sRtnColumnCode = sRtnColumnCode & " WHEN NOT lookup_" & iLoop & "." & sRtnColumnName & " IS NULL THEN lookup_" & iLoop & "." & sRtnColumnName
										Next iLoop

										sRtnColumnCode = sRtnColumnCode & " ELSE NULL END"
									End If
								End If


								objColumnPrivileges = GetColumnPrivileges(sSrchTableName)

								fSrchColumnOK = objColumnPrivileges.IsValid(sSrchColumnName)

								If fSrchColumnOK Then
									fSrchColumnOK = objColumnPrivileges.Item(sSrchColumnName).AllowSelect
								End If

								If fSrchColumnOK Then
									' The search column can be read direct from the table.
									sSrchColumnCode = "lookup" & "." & sSrchColumnName & " = " & sParamCode2
								Else
									' Then column cannot be read direct. If its from a parent, try parent views
									' Loop thru the views on the table, seeing if any have read permis for the column
									' Column 1 = view name
									' Column 2 = "S" if the view is used for the search.
									For Each objTableView In gcoTablePrivileges.Collection
										If (Not objTableView.IsTable) And (objTableView.TableID = lngSrchTableID) And (objTableView.AllowSelect) Then

											' Get the column permission for the view
											objColumnPrivileges = GetColumnPrivileges((objTableView.ViewName))

											' If we can see the column from this view
											If objColumnPrivileges.IsValid(sSrchColumnName) Then
												If objColumnPrivileges.Item(sSrchColumnName).AllowSelect Then

													fFound = False
													For iLoop = 1 To UBound(asViews, 2)
														If asViews(1, iLoop) = objTableView.ViewName Then
															fFound = True
															asViews(2, iLoop) = "S"
															Exit For
														End If
													Next iLoop

													If Not fFound Then
														ReDim Preserve asViews(2, UBound(asViews, 2) + 1)
														asViews(1, UBound(asViews, 2)) = objTableView.ViewName
														asViews(2, UBound(asViews, 2)) = "S"
													End If
												End If
											End If
										End If
									Next objTableView
									'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
									objTableView = Nothing

									sSrchColumnCode = ""
									For iLoop = 1 To UBound(asViews, 2)
										If asViews(2, iLoop) = "S" Then
											fSrchColumnOK = True

											If Len(sSrchColumnCode) > 0 Then
												sSrchColumnCode = sSrchColumnCode & " OR "
											End If

											sSrchColumnCode = sSrchColumnCode & sRealTableSource & ".id IN (SELECT id FROM " & asViews(1, iLoop) & " lookup_" & iLoop & " WHERE " & sSrchColumnName & " = " & sParamCode2 & ")"
										End If
									Next iLoop
								End If

								sSrchTableCode = ""
								For iLoop = 1 To UBound(asViews, 2)
									sSrchTableCode = sSrchTableCode & " LEFT OUTER JOIN " & asViews(1, iLoop) & " lookup_" & iLoop & " ON " & sRealTableSource & ".id = " & " lookup_" & iLoop & ".id"
								Next iLoop

								If fSrchColumnOK Then
									If UBound(asViews, 2) = 0 Then
										sCode = "(SELECT TOP 1 " & sRtnColumnCode & " FROM " & sRealTableSource & " lookup" & " " & sSrchTableCode & " WHERE (" & sSrchColumnCode & "))"
									Else
										sCode = "(SELECT TOP 1 " & sRtnColumnCode & " FROM " & sRealTableSource & " " & sSrchTableCode & " WHERE (" & sSrchColumnCode & "))"
									End If
								Else
									sCode = "null"
								End If
							End If

						Case 44	' Add days to date.
							sCode = "dateadd(day, " & sParamCode2 & ", " & sParamCode1 & ")"

						Case 45	' Days Between 2 Dates
							sCode = "datediff(dd, " & sParamCode1 & ", " & sParamCode2 & ")+1"

						Case 46	'Working days between two dates

							If pfValidating Then
								strTempTableName = mobjBaseComponent.ParentExpression.BaseTableName
							Else
								strTempTableName = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName).RealSource
							End If

							If mobjBaseComponent.ParentExpression.BaseTableID = PersonnelModule.glngPersonnelTableID Then
								strTempTableID = "ID"
							Else
								strTempTableID = "ID_" & Trim(Str(PersonnelModule.glngPersonnelTableID))
							End If

							sCode = "(dbo.udf_ASRFn_WorkingDaysBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 & "," & strTempTableName & "." & strTempTableID & "))"

						Case 47	' Absence between two dates

							If pfValidating Then
								strTempTableName = mobjBaseComponent.ParentExpression.BaseTableName
							Else
								strTempTableName = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName).RealSource
							End If

							If mobjBaseComponent.ParentExpression.BaseTableID = PersonnelModule.glngPersonnelTableID Then
								strTempTableID = "ID"
							Else
								strTempTableID = "ID_" & Trim(Str(PersonnelModule.glngPersonnelTableID))
							End If

							sCode = "(dbo.udf_ASRFn_AbsenceBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 & "," & sParamCode3 & "," & strTempTableName & "." & strTempTableID & "," & "convert(datetime,'" & VB6.Format(Now, "MM/dd/yyyy") & "')" & "))"

						Case 48	' Round Up to nearest whole number.
							sCode = "CASE WHEN (" & sParamCode1 & ") < 0 THEN floor(" & sParamCode1 & ")" & " ELSE ceiling(" & sParamCode1 & ") END"

						Case 49	' Round to nearest number.
							strRemainString = "(" & sParamCode1 & ") - ((floor(" & sParamCode1 & "/" & sParamCode2 & "))*" & sParamCode2 & ")"
							sCode = "CASE WHEN (((" & sParamCode1 & ")<0) AND ((" & strRemainString & ")<=((" & sParamCode2 & ")/2.0)))" & " OR (((" & sParamCode1 & ")>=0) AND ((" & strRemainString & ")<((" & sParamCode2 & ")/2.0)))" & " THEN (" & sParamCode1 & ")-(" & strRemainString & ")" & " ELSE (" & sParamCode1 & ")+(" & sParamCode2 & ")-(" & strRemainString & ") END"
							sCode = "CASE WHEN (" & sParamCode2 & " > 0) THEN " & sCode & " ELSE 0 END"

							'TM20011022 Currency Implementation
						Case 51
							'*********** runtime code to go here *************

							' Get the column parameter definitions.
							sSQL = "SELECT ASRSysModuleSetup.*, ASRSysColumns.ColumnName, ASRSysTables.TableName FROM ASRSysModuleSetup INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID WHERE ASRSysModuleSetup.ModuleKey = 'MODULE_CURRENCY'"
							rsInfo = DB.GetDataTable(sSQL)
							sSQL = vbNullString

							With rsInfo
								If .Rows.Count > 0 Then

									For Each objRow As DataRow In .Rows
										sCConvTable = objRow("TableName").ToString
										Select Case objRow("ParameterKey").ToString()
											Case "Param_CurrencyNameColumn" : sCConvCurrDescCol = objRow("ColumnName").ToString()
											Case "Param_ConversionValueColumn" : sCConvExRateCol = objRow("ColumnName").ToString
											Case "Param_DecimalColumn" : sCConvDecCol = objRow("ColumnName").ToString()
										End Select
									Next

									If (Len(sCConvTable) > 0) And (Len(sCConvCurrDescCol) > 0) And (Len(sCConvExRateCol) > 0) And (Len(sCConvDecCol) > 0) Then
										sCode = vbNullString
										sCode = sCode & "ROUND(ISNULL((" & sParamCode1 & " / NULLIF((SELECT " & sCConvTable & "." & sCConvExRateCol
										sCode = sCode & "                                     FROM " & sCConvTable
										sCode = sCode & "                                     WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode2 & "), 0))"
										sCode = sCode & " * "
										sCode = sCode & "                                   (SELECT " & sCConvTable & "." & sCConvExRateCol
										sCode = sCode & "                                    FROM " & sCConvTable
										sCode = sCode & "                                    WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & "), 0)"
										sCode = sCode & "        , "
										sCode = sCode & "         ISNULL((SELECT " & sCConvTable & "." & sCConvDecCol
										sCode = sCode & "                 FROM " & sCConvTable
										sCode = sCode & "                 WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & "), 0))"
									Else
										sCode = "null"
									End If
								Else
									sCode = "null"
								End If

							End With
							'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							rsInfo = Nothing

							' Last field change date
						Case 52
							objTableView = gcoTablePrivileges.FindTableID(Columns.GetById(CInt(sParamCode1)).TableID)
							sCode = "(SELECT Top 1 DateTimeStamp FROM ASRSysAuditTrail WHERE ColumnID = " & sParamCode1
							sCode = sCode & " And " & IIf(Not pfValidating, objTableView.RealSource & ".", "").ToString() & "ID = ASRSysAuditTrail.RecordID ORDER BY DateTimeStamp DESC)"

							' Field changed between two dates
						Case 53
							objTableView = gcoTablePrivileges.FindTableID(Columns.GetById(CInt(sParamCode1)).TableID)
							sCode = " case when " & " Exists(Select DateTimeStamp From ASRSysAuditTrail Where ColumnID = " & sParamCode1 & " And " & IIf(Not pfValidating, objTableView.RealSource & ".", "").ToString() & "ID = ASRSysAuditTrail.RecordID" & " And DateTimeStamp >= " & sParamCode2 & " And DateTimeStamp <= " & sParamCode3 & " + 1)" & " then 1 else 0 end"

							'Whole years between two dates
						Case 54
							sCode = " case " & " when " & sParamCode1 & " >= " & sParamCode2 & " then 0 " & " else " & "   datediff(year, " & sParamCode1 & ", " & sParamCode2 & ") " & "   - " & "   case " & "   when DatePart(Month, " & sParamCode2 & ") < DatePart(Month, " & sParamCode1 & ") " & "   then 1 " & "   else " & "     case " & "     when DatePart(Month, " & sParamCode2 & ") = DatePart(Month, " & sParamCode1 & ") " & "     then " & "       case " & "       when DatePart(Day, " & sParamCode2 & ") < DatePart(Day, " & sParamCode1 & ") " & "       then 1 " & "       else 0 " & "       end " & "     else " & "       0 " & "     end " & "   end " & " end "

							' JPD20021121 Fault 3177
						Case 55	' First Day of Month - VERSION 2 FUNCTION
							sCode = "dateadd(dd, 1 - datepart(dd, " & sParamCode1 & "), " & sParamCode1 & ")"

							' JPD20021121 Fault 3177
						Case 56	' Last Day of Month - VERSION 2 FUNCTION
							sCode = "dateadd(dd, -1, dateadd(mm, 1, dateadd(dd, 1 - datepart(dd, " & sParamCode1 & "), " & sParamCode1 & ")))"

							' JPD20021121 Fault 3177
						Case 57	' First Day of Year - VERSION 2 FUNCTION
							sCode = "dateadd(dd, 1 - datepart(dy, " & sParamCode1 & "), " & sParamCode1 & ")"

							' JPD20021121 Fault 3177
						Case 58	' Last Day of Year - VERSION 2 FUNCTION
							sCode = "dateadd(dd, -1, dateadd(yy, 1, dateadd(dd, 1 - datepart(dy, " & sParamCode1 & "), " & sParamCode1 & ")))"

							' JPD20021129 Fault 4337
						Case 59	' Name of Month. - VERSION 2 FUNCTION
							sCode = "datename(month, " & sParamCode1 & ")"

							' JPD20021129 Fault 4337
						Case 60	' Name of Day. - VERSION 2 FUNCTION
							sCode = "datename(weekday, " & sParamCode1 & ")"

							' JPD20021129 Fault 3606
						Case 61	' Is field populated
							sCode = "(CASE WHEN ((" & sParamCode1 & ") IS NULL)"

							' Validate the sub-expression. This is done, not to  validate the expression,
							' but rather to determine the return type of the expression.
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mcolParameters.Item(1).Component.ValidateExpression(False)

							'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters(1).ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Select Case mcolParameters.Item(1).ReturnType
								Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
									sCode = sCode & " OR ((" & sParamCode1 & ") = '')"
								Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
									sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
								Case ExpressionValueTypes.giEXPRVALUE_LOGIC
									sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
							End Select

							sCode = sCode & " THEN 0 ELSE 1 END)"

							'Case 62 'PARENTAL LEAVE ENTITLEMENT
							'Case 63 'PARENTAL LEAVE TAKEN
							'Case 64 'MATERNITY EXPECTED RETURN DATE

						Case 65	' Is Post Subordinate Of
							'If (Len(PersonnelModule.gsHierarchyTableName) > 0) Then
							'  Set objBaseTable = gcoTablePrivileges.FindTableID(PersonnelModule.glngHierarchyTableID)
							'  sCode = "CASE WHEN dbo.udf_ASRFn_IsPostSubordinateOf(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id) = 1 THEN 1 ELSE 0 END"
							'  Set objBaseTable = Nothing
							'Else
							'  sCode = "0"
							'End If

						Case 66	' Is Post Subordinate Of User
							If (Len(PersonnelModule.gsHierarchyTableName) > 0) Then
								objBaseTable = gcoTablePrivileges.FindTableID(PersonnelModule.glngHierarchyTableID)
								sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_IsPostSubordinateOfUser()) THEN 1 ELSE 0 END"
								'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objBaseTable = Nothing
							Else
								sCode = "0"
							End If

						Case 67	' Is Personnel Subordinate Of
							'If (Len(PersonnelModule.gsPersonnelTableName) > 0) Then
							'  Set objBaseTable = gcoTablePrivileges.FindTableID(PersonnelModule.glngPersonnelTableID)
							'  sCode = "CASE WHEN dbo.udf_ASRFn_IsPersonnelSubordinateOf(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id," & sParamCode2 & ") = 1 THEN 1 ELSE 0 END"
							'  Set objBaseTable = Nothing
							'Else
							'  sCode = "0"
							'End If

						Case 68	' Is Personnel Subordinate Of User
							If (Len(PersonnelModule.gsPersonnelTableName) > 0) Then
								objBaseTable = gcoTablePrivileges.FindTableID(PersonnelModule.glngPersonnelTableID)
								sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_IsPersonnelSubordinateOfUser()) THEN 1 ELSE 0 END"
								'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objBaseTable = Nothing
							Else
								sCode = "0"
							End If

						Case 69	'Has Post Subordinate
							'If (Len(PersonnelModule.gsHierarchyTableName) > 0) Then
							'  Set objBaseTable = gcoTablePrivileges.FindTableID(PersonnelModule.glngHierarchyTableID)
							'  sCode = "CASE WHEN dbo.udf_ASRFn_HasPostSubordinate(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id) = 1 THEN 1 ELSE 0 END"
							'  Set objBaseTable = Nothing
							'Else
							'  sCode = "0"
							'End If

						Case 70	'Has Post Subordinate User
							If (Len(PersonnelModule.gsHierarchyTableName) > 0) Then
								objBaseTable = gcoTablePrivileges.FindTableID(PersonnelModule.glngHierarchyTableID)
								sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_HasPostSubordinateUser()) THEN 1 ELSE 0 END"
								'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objBaseTable = Nothing
							Else
								sCode = "0"
							End If

						Case 71	'Has Personnel Subordinate
							'If (Len(PersonnelModule.gsPersonnelTableName) > 0) Then
							'  Set objBaseTable = gcoTablePrivileges.FindTableID(PersonnelModule.glngPersonnelTableID)
							'  sCode = "CASE WHEN dbo.udf_ASRFn_HasPersonnelSubordinate(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id," & sParamCode2 & ") = 1 THEN 1 ELSE 0 END"
							'  Set objBaseTable = Nothing
							'Else
							'  sCode = "0"
							'End If

						Case 72	'Has Personnel Subordinate User
							If (Len(PersonnelModule.gsPersonnelTableName) > 0) Then
								objBaseTable = gcoTablePrivileges.FindTableID(PersonnelModule.glngPersonnelTableID)
								sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_HasPersonnelSubordinateUser()) THEN 1 ELSE 0 END"
								'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objBaseTable = Nothing
							Else
								sCode = "0"
							End If

						Case 73	'Bradford Factor
							If pfValidating Then
								strTempTableName = mobjBaseComponent.ParentExpression.BaseTableName
							Else
								strTempTableName = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName).RealSource
							End If

							If mobjBaseComponent.ParentExpression.BaseTableID = PersonnelModule.glngPersonnelTableID Then
								strTempTableID = "ID"
							Else
								strTempTableID = "ID_" & Trim(Str(PersonnelModule.glngPersonnelTableID))
							End If

							sCode = "(dbo.udf_ASRFn_BradfordFactor(" & sParamCode1 & "," & sParamCode2 & "," & sParamCode3 & "," & strTempTableName & "." & strTempTableID & "))"

						Case 77	' Replace characters within a String
							sCode = "REPLACE(" & sParamCode1 & ", " & sParamCode2 & ", " & sParamCode3 & ")"

						Case 78	' Last run date time of current utility/report
							sCode = String.Format("'{0}-{1}-{2} {3}:{4}:{5}'", RunDate.Year, RunDate.Month, RunDate.Day, RunDate.Hour, RunDate.Minute, RunDate.Second)

						Case 79	' Current ID of the base record
							sCode = "ID"

						Case 80	' Convert date to string
							sCode = "dbo.udfASRCreateDate(" & sParamCode1 & ", " & sParamCode2 & ", " & sParamCode3 & ")"

						Case Else
							fOK = False

					End Select
				End If

				If fOK Then
					' We need to convert date values to varchars in the format 'mm/dd/yyyy'.
					If miReturnType = ExpressionValueTypes.giEXPRVALUE_DATE Then
						sCode = String.Format("CONVERT(datetime, convert(varchar(20),{0}, 101))", sCode)
					End If
				End If

			Catch ex As Exception
				fOK = False

			Finally
				If fOK Then
					psRuntimeCode = String.Format("({0})", sCode)
				Else
					psRuntimeCode = ""
				End If

			End Try

			Return fOK

		End Function

		Public Function WriteComponent() As Boolean

			Dim fOK As Boolean
			Dim objParameter As clsExprComponent
			Dim objSubExpression As clsExprExpression

			Try
				DB.ExecuteSP("spASRIntSaveComponent", _
						New SqlParameter("componentID", SqlDbType.Int) With {.Value = mobjBaseComponent.ComponentID}, _
						New SqlParameter("expressionID", SqlDbType.Int) With {.Value = mobjBaseComponent.ParentExpression.ExpressionID}, _
						New SqlParameter("type", SqlDbType.TinyInt) With {.Value = ExpressionComponentTypes.giCOMPONENT_FUNCTION}, _
						New SqlParameter("calculationID", SqlDbType.Int), _
						New SqlParameter("filterID", SqlDbType.Int), _
						New SqlParameter("functionID", SqlDbType.Int) With {.Value = mlngFunctionID}, _
						New SqlParameter("operatorID", SqlDbType.Int), _
						New SqlParameter("valueType", SqlDbType.TinyInt), _
						New SqlParameter("valueCharacter", SqlDbType.VarChar, 255), _
						New SqlParameter("valueNumeric", SqlDbType.Float), _
						New SqlParameter("valueLogic", SqlDbType.Bit), _
						New SqlParameter("valueDate", SqlDbType.DateTime), _
						New SqlParameter("LookupTableID", SqlDbType.Int), _
						New SqlParameter("LookupColumnID", SqlDbType.Int), _
						New SqlParameter("fieldTableID", SqlDbType.Int), _
						New SqlParameter("fieldColumnID", SqlDbType.Int), _
						New SqlParameter("fieldPassBy", SqlDbType.TinyInt), _
						New SqlParameter("fieldSelectionRecord", SqlDbType.TinyInt), _
						New SqlParameter("fieldSelectionLine", SqlDbType.Int), _
						New SqlParameter("fieldSelectionOrderID", SqlDbType.Int), _
						New SqlParameter("fieldSelectionFilter", SqlDbType.Int), _
						New SqlParameter("promptDescription", SqlDbType.VarChar, 255), _
						New SqlParameter("promptSize", SqlDbType.SmallInt), _
						New SqlParameter("promptDecimals", SqlDbType.SmallInt), _
						New SqlParameter("promptMask", SqlDbType.VarChar, 255), _
						New SqlParameter("promptDateType", SqlDbType.Int))

				' Write the function parameter expressions.
				For Each objParameter In mcolParameters
					objSubExpression = objParameter.Component
					objSubExpression.ParentComponentID = mobjBaseComponent.ComponentID
					objSubExpression.ExpressionID = 0
					fOK = objSubExpression.WriteExpression

					If Not fOK Then
						Exit For
					End If
				Next objParameter

				Return True

			Catch ex As Exception
				Return False

			End Try

		End Function

		Public ReadOnly Property BadComponent() As clsExprComponent
			Get
				' Return the component last caused the function to fail its validity check.
				BadComponent = mobjBadComponent

			End Get
		End Property

		Public ReadOnly Property ReturnType() As ExpressionValueTypes
			Get
				Return miReturnType
			End Get
		End Property

		Public ReadOnly Property ComponentType() As ExpressionComponentTypes
			Get
				Return ExpressionComponentTypes.giCOMPONENT_FUNCTION
			End Get
		End Property

		Public Property BaseComponent() As clsExprComponent
			Get
				' Return the component's base component object.
				BaseComponent = mobjBaseComponent

			End Get
			Set(ByVal Value As clsExprComponent)
				' Set the component's base component object property.
				mobjBaseComponent = Value

			End Set
		End Property



		Public Property FunctionID() As Integer
			Get
				' Return the function ID property.
				FunctionID = mlngFunctionID

			End Get
			Set(ByVal Value As Integer)
				' Set the function ID property.
				mlngFunctionID = Value

				' Read the function definition from the database.
				ReadFunction()

			End Set
		End Property
		Public ReadOnly Property ComponentDescription() As String
			Get
				' Return the Function description.
				ComponentDescription = msFunctionName

			End Get
		End Property


		Public Property Parameters() As Collection
			Get
				' Return the parameter collection.
				Parameters = mcolParameters

			End Get
			Set(ByVal Value As Collection)
				' Set the parameter collection.
				mcolParameters = Value

			End Set
		End Property


		Public Property ExpandedNode() As Boolean
			Get
				'Return whether this node is expanded or not
				ExpandedNode = mbExpanded

			End Get
			Set(ByVal Value As Boolean)
				'Set whether this component node is expanded or not
				mbExpanded = Value

			End Set
		End Property

		Public Function CopyComponent() As Object
			' Copies the selected component.
			' When editting a component we actually copy the component first
			' and edit the copy. If the changes are confirmed then the copy
			' replaces the original. If the changes are cancelled then the
			' copy is discarded.
			Dim objFunctionCopy As New clsExprFunction(SessionInfo)

			' Copy the component's basic properties.
			' ie. the function id, not its parameters, etc.
			With objFunctionCopy
				.FunctionID = mlngFunctionID
			End With

			' JDM - 06/02/01 - Now copies it's children so that cut'n paste works
			' Copy all the child components
			Dim iCount As Integer
			Dim objParameter As clsExprComponent
			For iCount = 1 To mcolParameters.Count()
				'        Set objParameter = New clsExprComponent
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters().CopyComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				objFunctionCopy.Parameters.Add((mcolParameters.Item(iCount).CopyComponent))

				'        objFunctionCopy.Parameters.Add (mcolParameters.Item(iCount))
			Next iCount
			'UPGRADE_NOTE: Object objParameter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objParameter = Nothing


			CopyComponent = objFunctionCopy

			' Disassociate object variables.
			'UPGRADE_NOTE: Object objFunctionCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objFunctionCopy = Nothing

		End Function

		Public Function ValidateFunction() As ExprValidationCodes
			' Validate the function. Return a code describing the validity.
			Dim iLoop As Integer = 0
			Dim iValidationCode As ExprValidationCodes = ExprValidationCodes.giEXPRVALIDATION_NOERRORS
			Dim iFunctionReturnType As ExpressionValueTypes
			Dim aiDummyValues(6) As Integer
			Dim objSubExpression As clsExprExpression
			Dim objParameter As clsExprComponent

			Try

				'UPGRADE_NOTE: Object mobjBadComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				mobjBadComponent = Nothing

				' Validate the function parameter expressions.
				For Each objParameter In mcolParameters
					iLoop += 1
					objSubExpression = objParameter.Component
					With objSubExpression
						' Validate the parameter expression.
						' NB. Reset the sub-expression's return type to that defined by the parameter definition
						' as it may be changeable. The evaluated return type will be determined when the
						' sub-expression is validated.
						objSubExpression.ReturnType = Functions.GetById(mlngFunctionID).Parameters.GetByIndex(iLoop - 1).ParameterType
						objSubExpression.SessionInfo = SessionInfo

						iValidationCode = .ValidateExpression(False)

						If iValidationCode <> ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
							' Interpret the parameter sub-expression validation code to reflect
							' the fact that a function parameter was invalid.
							Select Case iValidationCode
								Case ExprValidationCodes.giEXPRVALIDATION_NOCOMPONENTS
									iValidationCode = ExprValidationCodes.giEXPRVALIDATION_PARAMETERNOCOMPONENTS
								Case ExprValidationCodes.giEXPRVALIDATION_SYNTAXERROR
									iValidationCode = ExprValidationCodes.giEXPRVALIDATION_PARAMETERSYNTAXERROR
								Case ExprValidationCodes.giEXPRVALIDATION_EXPRTYPEMISMATCH
									iValidationCode = ExprValidationCodes.giEXPRVALIDATION_PARAMETERTYPEMISMATCH
							End Select

							If .BadComponent Is Nothing Then
								mobjBadComponent = objParameter
							Else
								mobjBadComponent = .BadComponent
							End If
							Exit For
						End If

						' Write the given return type into the array.
						aiDummyValues(iLoop) = .ReturnType
					End With
				Next objParameter
				'UPGRADE_NOTE: Object objParameter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objParameter = Nothing

				If iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
					If Not ValidateFunctionParameters(mlngFunctionID, iFunctionReturnType, aiDummyValues(1), aiDummyValues(2), aiDummyValues(3), aiDummyValues(4), aiDummyValues(5), aiDummyValues(6)) Then

						iValidationCode = ExprValidationCodes.giEXPRVALIDATION_PARAMETERTYPEMISMATCH
					Else
						miReturnType = iFunctionReturnType
					End If
				End If

			Catch ex As Exception
				iValidationCode = ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR

			End Try

			Return iValidationCode

		End Function

		' Read the function name, return type, etc.
		Private Sub ReadFunction()

			Dim fOK As Boolean = True
			Dim iIndex As Integer
			Dim sSQL As String

			Dim objNewParameter As clsExprComponent

			Try

				Dim objFunction = Functions.GetById(mlngFunctionID)
				msFunctionName = objFunction.Name
				miReturnType = objFunction.ReturnType

				' Clear the parameter collection.
				mcolParameters = New Collection

				' Get the standard function parameter definitions.
				For Each objParameter In Functions.GetById(mlngFunctionID).Parameters

					mcolParameters.Add(New clsExprComponent(SessionInfo))
					With mcolParameters.Item(mcolParameters.Count())
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.ComponentType = ExpressionComponentTypes.giCOMPONENT_EXPRESSION
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item(mcolParameters.Count).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Component.Name = objParameter.Name
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item(mcolParameters.Count).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Component.ReturnType = objParameter.ParameterType
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID
					End With

				Next

				Dim rsParameters As DataTable

				' Get the customised function parameter definitions if they exist.
				If mobjBaseComponent.ComponentID > 0 And mcolParameters.Count > 0 Then

					iIndex = 1
					sSQL = "SELECT ExprID FROM ASRSysExpressions WHERE parentComponentID = " & Trim(Str(mobjBaseComponent.ComponentID)) & " ORDER BY exprID"

					rsParameters = DB.GetDataTable(sSQL, CommandType.Text)
					For Each objRow As DataRow In rsParameters.Rows

						' Instantiate a new component object.
						objNewParameter = New clsExprComponent(SessionInfo)

						' Construct the hierarchy of objects that define the parameter.
						objNewParameter.ComponentType = ExpressionComponentTypes.giCOMPONENT_EXPRESSION
						'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ExpressionID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						objNewParameter.Component.ExpressionID = objRow("ExprID")
						'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.BaseTableID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						objNewParameter.Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID

						' JDM - 06/11/2003 - Fault 7193 - Get field from database record not working
						'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ExpressionType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						objNewParameter.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType

						'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ConstructExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						fOK = objNewParameter.Component.ConstructExpression

						If fOK Then
							'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ExpressionType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							objNewParameter.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType

							' Reset the new expression's return type.
							'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							objNewParameter.Component.ReturnType = mcolParameters.Item(iIndex).Component.ReturnType

							' Insert the new expression into the function's parameter array.
							mcolParameters.Remove(iIndex)
							If mcolParameters.Count() >= iIndex Then
								mcolParameters.Add(objNewParameter, , iIndex)
							Else
								mcolParameters.Add(objNewParameter)
							End If
						End If

						iIndex = iIndex + 1
						'UPGRADE_NOTE: Object objNewParameter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objNewParameter = Nothing

					Next
				End If

			Catch ex As Exception
				mcolParameters = New Collection
				Throw
			End Try

		End Sub

		Private Function ReadParameterDefinition() As Boolean
			' Read the function's paramter definition from the database,
			' and create an array of components to represent the parameters.

			Dim fOK As Boolean = True
			Dim iIndex As Integer
			Dim sSQL As String

			Dim objNewParameter As clsExprComponent

			' Clear the parameter collection.
			mcolParameters = New Collection

			' Get the standard function parameter definitions.
			For Each objParameter In Functions.GetById(mlngFunctionID).Parameters

				mcolParameters.Add(New clsExprComponent(SessionInfo))
				With mcolParameters.Item(mcolParameters.Count())
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ComponentType = ExpressionComponentTypes.giCOMPONENT_EXPRESSION
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item(mcolParameters.Count).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Component.Name = objParameter.Name
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item(mcolParameters.Count).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Component.ReturnType = objParameter.ParameterType
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID
				End With

			Next

			Dim rsParameters As DataTable

			' Get the customised function parameter definitions if they exist.
			If mobjBaseComponent.ComponentID > 0 And mcolParameters.Count > 0 Then

				iIndex = 1
				sSQL = "SELECT ExprID FROM ASRSysExpressions WHERE parentComponentID = " & Trim(Str(mobjBaseComponent.ComponentID)) & " ORDER BY exprID"

				rsParameters = DB.GetDataTable(sSQL, CommandType.Text)
				For Each objRow As DataRow In rsParameters.Rows

					' Instantiate a new component object.
					objNewParameter = New clsExprComponent(SessionInfo)

					' Construct the hierarchy of objects that define the parameter.
					objNewParameter.ComponentType = ExpressionComponentTypes.giCOMPONENT_EXPRESSION
					'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ExpressionID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					objNewParameter.Component.ExpressionID = objRow("ExprID")
					'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.BaseTableID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					objNewParameter.Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID

					' JDM - 06/11/2003 - Fault 7193 - Get field from database record not working
					'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ExpressionType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					objNewParameter.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType

					'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ConstructExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fOK = objNewParameter.Component.ConstructExpression

					If fOK Then
						'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ExpressionType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						objNewParameter.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType

						' Reset the new expression's return type.
						'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						objNewParameter.Component.ReturnType = mcolParameters.Item(iIndex).Component.ReturnType

						' Insert the new expression into the function's parameter array.
						mcolParameters.Remove(iIndex)
						If mcolParameters.Count() >= iIndex Then
							mcolParameters.Add(objNewParameter, , iIndex)
						Else
							mcolParameters.Add(objNewParameter)
						End If
					End If

					iIndex = iIndex + 1
					'UPGRADE_NOTE: Object objNewParameter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objNewParameter = Nothing

				Next
			End If

TidyUpAndExit:
			'UPGRADE_NOTE: Object rsParameters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsParameters = Nothing
			ReadParameterDefinition = fOK
			Exit Function

ErrorTrap:
			' Clear the parameter collection.
			mcolParameters = New Collection
			fOK = False
			Resume TidyUpAndExit

		End Function

		Public Sub New(ByVal Value As SessionInfo)
			MyBase.New(Value)
			mcolParameters = New Collection
		End Sub

		Protected Overrides Sub Finalize()
			mcolParameters = Nothing
			MyBase.Finalize()
		End Sub

	End Class
End Namespace