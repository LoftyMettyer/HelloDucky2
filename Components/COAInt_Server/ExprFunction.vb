Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class clsExprFunction
	
	' Component definition variables.
	Private mlngFunctionID As Integer
	Private msFunctionName As String
	Private miReturnType As modExpression.ExpressionValueTypes
	Private msSPName As String
	
	' Definition for expanded/unexpanded status of the component
	Private mbExpanded As Boolean
	
	' Class handling variables.
	Private mobjBaseComponent As clsExprComponent
	Private mcolParameters As Collection
	Private mobjBadComponent As clsExprComponent
	
	
	Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
		' Retrun TRUE if the current expression (or any of its sub expressions)
		' contains the given expression. This ensures no cyclic expressions get created.
		'JPD 20040507 Fault 8600
		On Error GoTo ErrorTrap
		
		Dim objParameter As clsExprComponent
		Dim objSubExpression As clsExprExpression
		
		ContainsExpression = False
		
		For	Each objParameter In mcolParameters
			objSubExpression = objParameter.Component
			
			ContainsExpression = objSubExpression.ContainsExpression(plngExprID)
			
			'UPGRADE_NOTE: Object objSubExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objSubExpression = Nothing
			
			If ContainsExpression Then
				Exit For
			End If
		Next objParameter
		'UPGRADE_NOTE: Object objParameter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objParameter = Nothing
		
TidyUpAndExit: 
		Exit Function
		
ErrorTrap: 
		ContainsExpression = True
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	
	
	
	
	Public Function PrintComponent(ByRef piLevel As Short) As Boolean
		Dim Printer As New Printer
		' Print the component definition to the printer object.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim objParameter As clsExprComponent
		
		fOK = True
		
		' Position the printing.
		With Printer
			.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
			.CurrentY = .CurrentY + giPRINT_YSPACE
			Printer.Print("Function : " & ComponentDescription)
		End With
		
		' Print the function's parameter expressions.
		For	Each objParameter In mcolParameters
			'UPGRADE_WARNING: Couldn't resolve default property of object objParameter.Component.PrintComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			objParameter.Component.PrintComponent(piLevel + 1)
		Next objParameter
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object objParameter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objParameter = Nothing
		PrintComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		' Return the SQL code for the component.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim fFound As Boolean
		Dim fSrchColumnOK As Boolean
		Dim fRtnColumnOK As Boolean
		Dim iLoop As Short
		Dim lngSrchTableID As Integer
		Dim lngRtnTableID As Integer
		Dim sCode As String
		Dim sSQL As String
		Dim sRtnColumnCode As String
		Dim sSrchColumnCode As String
		Dim sSrchTableCode As String
		Dim sRealTableSource As String
		Dim sParamCode1 As String
		Dim sParamCode2 As String
		Dim sParamCode3 As String
		Dim sParamCode4 As String
		Dim sSrchColumnName As String
		Dim sRtnColumnName As String
		Dim sSrchTableName As String
		Dim rsInfo As ADODB.Recordset
		Dim objColumnPrivileges As CColumnPrivileges
		Dim objTableView As CTablePrivilege
		Dim asViews() As String
		Dim strRemainString As String
		Dim strTempTableName As String
		Dim strTempTableID As String
		Dim objBaseTable As CTablePrivilege
		
		'Currency Conversion Values
		Dim sCConvTable As String
		Dim sCConvExRateCol As String
		Dim sCConvCurrDescCol As String
		Dim sCConvDecCol As String
		
		fOK = True
		sCode = ""
		
		'UPGRADE_NOTE: clsGeneral was upgraded to clsGeneral_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim clsGeneral_Renamed As clsGeneral
		clsGeneral_Renamed = New clsGeneral
		
		' Get the first parameter's runtime code if required.
		If mcolParameters.Count() >= 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mcolParameters.Item(1).Component.RuntimeCode(sParamCode1, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues)
		End If
		
		' Get the second parameter's runtime code if required.
		If fOK And (mcolParameters.Count() >= 2) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mcolParameters.Item(2).Component.RuntimeCode(sParamCode2, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues)
		End If
		
		' Get the third parameter's runtime code if required.
		If fOK And (mcolParameters.Count() >= 3) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mcolParameters.Item(3).Component.RuntimeCode(sParamCode3, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues)
		End If
		
		' Get the fourth parameter's runtime code if required.
		If fOK And (mcolParameters.Count() >= 4) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mcolParameters.Item(4).Component.RuntimeCode(sParamCode4, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues)
		End If
		
		If fOK Then
			Select Case mlngFunctionID
				Case 1 ' System date
					sCode = "getdate()"
					
				Case 2 ' Convert to uppercase
					sCode = "upper(" & sParamCode1 & ")"
					
				Case 3 ' Convert numeric to string
					sCode = "IsNull(ltrim(str(" & sParamCode1 & ", 255, " & sParamCode2 & ")),'')"
					
				Case 4 ' If... Then... Else...
					sCode = "CASE WHEN (" & sParamCode1 & " = 1) THEN " & sParamCode2 & " ELSE " & sParamCode3 & " END"
					
				Case 5 ' Remove leading and trailing spaces
					sCode = "ltrim(rtrim(" & sParamCode1 & "))"
					
				Case 6 ' Extract characters from the left
					sCode = "left(" & sParamCode1 & ", " & sParamCode2 & ")"
					
				Case 7 ' Length of character field
					sCode = "len(" & sParamCode1 & ")"
					
				Case 8 ' Convert to lowercase
					sCode = "lower(" & sParamCode1 & ")"
					
				Case 9 ' Maximum
					sCode = "CASE WHEN (" & sParamCode1 & " > " & sParamCode2 & ") THEN " & sParamCode1 & " ELSE " & sParamCode2 & " END"
					
				Case 10 ' Minimum
					sCode = "CASE WHEN (" & sParamCode1 & " < " & sParamCode2 & ") THEN " & sParamCode1 & " ELSE " & sParamCode2 & " END"
					
				Case 11 ' Search for character string.
					sCode = "charindex(" & sParamCode2 & ", " & sParamCode1 & ")"
					
				Case 12 ' Capitalise Initials
					sCode = "(dbo.udf_ASRFn_CapitalizeInitials(" & sParamCode1 & "))"
					
				Case 13 ' Extract characters from the right
					sCode = "right(" & sParamCode1 & ", " & sParamCode2 & ")"
					
				Case 14 ' Extract part of a character string
					sCode = "substring(" & sParamCode1 & ", " & sParamCode2 & ", " & sParamCode3 & ")"
					
				Case 15 ' System Time
					sCode = "convert(varchar(50), getdate(), 8)"
					
				Case 16 ' Is field empty
					sCode = "(CASE WHEN ((" & sParamCode1 & ") IS NULL)"
					
					' Validate the sub-expression. This is done, not to  validate the expression,
					' but rather to determine the return type of the expression.
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mcolParameters.Item(1).Component.ValidateExpression(False)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters(1).ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Select Case mcolParameters.Item(1).ReturnType
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER
							sCode = sCode & " OR ((" & sParamCode1 & ") = '')"
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC
							sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC
							sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
					End Select
					
					sCode = sCode & " THEN 1 ELSE 0 END)"
					
				Case 17 ' Current user
					sCode = "system_user"
					
				Case 18 ' Whole Years Until Current Date
					sCode = "datediff(year," & sParamCode1 & ", getdate())" & " - case" & "       when datepart(month," & sParamCode1 & ") > datepart(month, getdate()) then 1" & "       when (datepart(month," & sParamCode1 & ") = datepart(month, getdate())) " & "           and (datepart(day," & sParamCode1 & ") > datepart(day, getdate())) then 1" & "       else 0" & "   end"
					
				Case 19 ' Remaining Months Since Whole Years
					sCode = "datepart(month, getdate())" & " - datepart(month, " & sParamCode1 & ")" & " - case" & "       when datepart(day," & sParamCode1 & ") > datepart(day, getdate()) then 1" & "       else 0" & "   end" & " + case" & "       when (datepart(month, getdate())" & "           - datepart(month," & sParamCode1 & ")" & "           - case" & "               when datepart(day," & sParamCode1 & ") > datepart(day, getdate()) then 1" & "               else 0" & "             end) < 0 then 12" & "       else 0" & "   end"
					
				Case 20 ' Capitalise Initials
					sCode = "(dbo.udf_ASRFn_InitialsFromForenames(" & sParamCode1 & "))"
					
				Case 21 ' First Name from Forenames
					sCode = "case" & "    when charindex(' ', ltrim(" & sParamCode1 & ")) > 0 then substring(ltrim(" & sParamCode1 & "), 1, charindex(' ', ltrim(" & sParamCode1 & "))-1)" & "    else ltrim(" & sParamCode1 & ")" & "end"
					
				Case 22 ' Weekdays From Start and End Dates
					
					'MH200600802 Fault 11395
					'"    when datediff(day, " & sParamCode1 & ", " & sParamCode2 & ") <= 0 then 0" & _
					'
					sCode = " case" & "    when datediff(day, " & sParamCode1 & ", " & sParamCode2 & ") < 0 then 0" & "    else datediff(day, " & sParamCode1 & ", " & sParamCode2 & " + 1)" & "        - (2 * (datediff(day, " & sParamCode1 & " - (datepart(dw, " & sParamCode1 & ") - 1), " & "                              " & sParamCode2 & " - (datepart(dw, " & sParamCode2 & ") - 1)) / 7))" & "        - case" & "              when datepart(dw, " & sParamCode1 & ") = 1 then 1" & "              else 0" & "          end" & "        - case" & "              when datepart(dw, " & sParamCode2 & ") = 7 then 1" & "              else 0" & "          end" & " end"
					
				Case 23 ' Add months to date
					sCode = "dateadd(month, " & sParamCode2 & ", " & sParamCode1 & ")"
					
				Case 24 ' Add years to date
					sCode = "dateadd(year, " & sParamCode2 & ", " & sParamCode1 & ")"
					
				Case 25 ' Convert character to numeric.
					'JPD 20041213 Fault 9568
					'sCode = _
					'" case" & _
					'"    when isnumeric(" & sParamCode1 & ") = 1 then convert(float, " & sParamCode1 & ")" & _
					'"    else 0" & _
					'" end"
					sCode = " case" & "    when isnumeric(" & sParamCode1 & ") = 1 then convert(float, convert(money, " & sParamCode1 & "))" & "    else 0" & " end"
					
				Case 26 ' Whole Months between 2 Dates.
					sCode = " case" & "    when " & sParamCode1 & " >= " & sParamCode2 & " then 0" & "    else datediff(month, " & sParamCode1 & ", " & sParamCode2 & ")" & "        - case" & "              when datepart(day, " & sParamCode2 & ") < datepart(day, " & sParamCode1 & ") then 1" & "              else 0" & "          end" & " end"
					
				Case 27 ' Parentheses
					sCode = sParamCode1
					
				Case 28 ' Day of the week
					sCode = "DATEPART(weekday, " & sParamCode1 & ")"
					
				Case 29 ' Working Days per week
					sCode = "(convert(float, len(replace(left(" & sParamCode1 & ", 14), ' ', ''))) / 2)"
					
				Case 30 ' Absence Duration
					'TM08102003
					If pfValidating Then
						strTempTableName = mobjBaseComponent.ParentExpression.BaseTableName
					Else
						strTempTableName = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName).RealSource
					End If
					
					If mobjBaseComponent.ParentExpression.BaseTableID = glngPersonnelTableID Then
						strTempTableID = "ID"
					Else
						strTempTableID = "ID_" & Trim(Str(glngPersonnelTableID))
					End If
					
					sCode = "(dbo.udf_ASRFn_AbsenceDuration(" & sParamCode1 & "," & sParamCode2 & "," & sParamCode3 & "," & sParamCode4 & "," & strTempTableName & "." & strTempTableID & "))"
					
				Case 31 ' Round down to nearest whole number.
					sCode = "round(" & sParamCode1 & ", 0, 1)"
					
				Case 32 ' Year of date.
					sCode = "datepart(year, " & sParamCode1 & ")"
					
				Case 33 ' Month of date.
					sCode = "datepart(month, " & sParamCode1 & ")"
					
				Case 34 ' Day of date.
					sCode = "datepart(day, " & sParamCode1 & ")"
					
				Case 35 ' Nice Date
					sCode = "datename(day, " & sParamCode1 & ") + ' ' + " & "datename(month, " & sParamCode1 & ") + ' ' + " & "datename(year, " & sParamCode1 & ")"
					
				Case 36 ' Nice Time
					'sCode = _
					'"convert(varchar(2), datepart(hour, convert(datetime, " & sParamCode1 & ")) % 12) + ':' " & _
					'"    + right('00' + datename(minute, convert(datetime, " & sParamCode1 & ")),2)" & _
					'"    + case" & _
					'"          when datepart(hour, convert(datetime, " & sParamCode1 & ")) > 11 then ' pm'" & _
					'"          else ' am'" & _
					'"      end"
					' JPD20020618 Fault 3999
					sCode = "case when len(ltrim(rtrim(" & sParamCode1 & "))) = 0 then ''" & " else case when isdate(" & sParamCode1 & ") = 0 then '***'" & " else (convert(varchar(2),((datepart(hour,convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)) + 11) % 12) + 1)" & " + ':' + right('00' + datename(minute, convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)),2)" & " + case when datepart(hour, convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)) > 11 then ' pm'" & " else ' am' end) end end"
					
				Case 37 ' Round Date to Start of nearest month
					sCode = " case" & "     when datediff(day, (" & sParamCode1 & " - datepart(day, " & sParamCode1 & ") + 1), " & sParamCode1 & ")" & "         <= datediff(day, " & sParamCode1 & ", (dateadd(month, 1, " & sParamCode1 & ") - datepart(day, dateadd(month, 1, " & sParamCode1 & ")) + 1))" & "         then " & sParamCode1 & " - datepart(day, " & sParamCode1 & ") + 1" & "     else dateadd(month, 1, " & sParamCode1 & ")" & "         - datepart(day, dateadd(month, 1, " & sParamCode1 & ")) + 1" & " end"
					
				Case 38 ' Is Between
					sCode = " case" & "     when (" & sParamCode1 & " >= " & sParamCode2 & ")" & "         and (" & sParamCode1 & " <= " & sParamCode3 & ") then 1" & "     else 0" & " end"
					
				Case 39 ' Service Years
					sCode = "   datepart(year, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)" & " - datepart(year, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & " - case" & "       when datepart(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & "          > datepart(month, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end) then 1" & "      when (datepart(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & "          = datepart(month, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end))" & "       and (datepart(day, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & "          > datepart(day, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)) then 1" & "      else 0" & "  end"
					
				Case 40 ' Service Months
					sCode = " (case" & "    when case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end" & "          >= case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end then 0" & "    else datediff(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)" & "        - case" & "              when datepart(day, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end) < datepart(day, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end) then 1" & "              else 0" & "          end" & " end) % 12"
					
				Case 42 ' Get field from database record.
					' Get the column parameter definitions.
					sSQL = "SELECT ASRSysColumns.columnID, ASRSysColumns.columnName, ASRSysTables.tableID, ASRSysTables.tableName" & " FROM ASRSysColumns" & " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & " WHERE ASRSysColumns.columnID IN (" & sParamCode1 & ", " & sParamCode3 & ")"
					rsInfo = datGeneral.GetRecords(sSQL)
					With rsInfo
						Do While Not .EOF
							If Trim(Str(.Fields("ColumnID").Value)) = sParamCode1 Then
								sSrchColumnName = .Fields("ColumnName").Value
								sSrchTableName = .Fields("TableName").Value
								lngSrchTableID = .Fields("TableID").Value
							End If
							
							If Trim(Str(.Fields("ColumnID").Value)) = sParamCode3 Then
								sRtnColumnName = .Fields("ColumnName").Value
								lngRtnTableID = .Fields("TableID").Value
							End If
							
							.MoveNext()
						Loop 
						
						.Close()
					End With
					'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					rsInfo = Nothing
					
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
							For	Each objTableView In gcoTablePrivileges.Collection
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
									
									sRtnColumnCode = sRtnColumnCode & " WHEN NOT " & "lookup_" & iLoop & "." & sRtnColumnName & " IS NULL THEN " & "lookup_" & iLoop & "." & sRtnColumnName
								Next iLoop
								
								sRtnColumnCode = sRtnColumnCode & " ELSE NULL" & " END"
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
							For	Each objTableView In gcoTablePrivileges.Collection
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
					
				Case 44 ' Add days to date.
					sCode = "dateadd(day, " & sParamCode2 & ", " & sParamCode1 & ")"
					
				Case 45 ' Days Between 2 Dates
					
					'MH20010220 Fault 1850
					'This function needs to be inclusive of both start and end
					'sCode = "datediff(dd, " & sParamCode1 & ", " & sParamCode2 & ")"
					sCode = "datediff(dd, " & sParamCode1 & ", " & sParamCode2 & ")+1"
					
				Case 46 'Working days between two dates
					
					If pfValidating Then
						strTempTableName = mobjBaseComponent.ParentExpression.BaseTableName
					Else
						strTempTableName = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName).RealSource
					End If
					
					If mobjBaseComponent.ParentExpression.BaseTableID = glngPersonnelTableID Then
						strTempTableID = "ID"
					Else
						strTempTableID = "ID_" & Trim(Str(glngPersonnelTableID))
					End If
					
					sCode = "(dbo.udf_ASRFn_WorkingDaysBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 & "," & strTempTableName & "." & strTempTableID & "))"
					
				Case 47 ' Absence between two dates
					
					If pfValidating Then
						strTempTableName = mobjBaseComponent.ParentExpression.BaseTableName
					Else
						strTempTableName = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName).RealSource
					End If
					
					If mobjBaseComponent.ParentExpression.BaseTableID = glngPersonnelTableID Then
						strTempTableID = "ID"
					Else
						strTempTableID = "ID_" & Trim(Str(glngPersonnelTableID))
					End If
					
					sCode = "(dbo.udf_ASRFn_AbsenceBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 & "," & sParamCode3 & "," & strTempTableName & "." & strTempTableID & "," & "convert(datetime,'" & VB6.Format(Now, "mm/dd/yyyy") & "')" & "))"
					
				Case 48 ' Round Up to nearest whole number.
					' JPD20030116 Fault 4910
					'sCode = "ceiling(" & sParamCode1 & ")"
					sCode = "CASE WHEN (" & sParamCode1 & ") < 0 THEN floor(" & sParamCode1 & ")" & " ELSE ceiling(" & sParamCode1 & ") END"
					
				Case 49 ' Round to nearest number.
					' JPD20020415 Fault 3701
					' Changed 'division by 2' to 'division by 2.0' to avoid SQL casting the result to an integer value.
					strRemainString = "(" & sParamCode1 & ") - ((floor(" & sParamCode1 & "/" & sParamCode2 & "))*" & sParamCode2 & ")"
					' JPD20030116 Fault 4910
					'sCode = "CASE WHEN (" + strRemainString + ")<((" + sParamCode2 + ")/2.0)" & _
					'" THEN (" + sParamCode1 + ")-(" + strRemainString + ")" & _
					'" ELSE (" + sParamCode1 + ")+(" + sParamCode2 + ")-(" + strRemainString + ") END"
					sCode = "CASE WHEN (((" & sParamCode1 & ")<0) AND ((" & strRemainString & ")<=((" & sParamCode2 & ")/2.0)))" & " OR (((" & sParamCode1 & ")>=0) AND ((" & strRemainString & ")<((" & sParamCode2 & ")/2.0)))" & " THEN (" & sParamCode1 & ")-(" & strRemainString & ")" & " ELSE (" & sParamCode1 & ")+(" & sParamCode2 & ")-(" & strRemainString & ") END"
					
					'MH20100629
					sCode = "CASE WHEN (" & sParamCode2 & " > 0) THEN " & sCode & " ELSE 0 END"
					
					'TM20011022 Currency Implementation
				Case 51
					'*********** runtime code to go here *************
					
					' Get the column parameter definitions.
					sSQL = "SELECT ASRSysModuleSetup.*, ASRSysColumns.ColumnName, ASRSysTables.TableName" & " FROM ASRSysModuleSetup" & "     INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID" & "     INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID" & " WHERE ASRSysModuleSetup.ModuleKey = 'MODULE_CURRENCY'"
					
					rsInfo = datGeneral.GetRecords(sSQL)
					sSQL = vbNullString
					
					With rsInfo
						If .RecordCount > 0 Then
							.MoveLast()
							.MoveFirst()
							Do While Not .EOF
								sCConvTable = .Fields("TableName").Value
								Select Case .Fields("ParameterKey").Value
									Case "Param_CurrencyNameColumn" : sCConvCurrDescCol = .Fields("ColumnName").Value
									Case "Param_ConversionValueColumn" : sCConvExRateCol = .Fields("ColumnName").Value
									Case "Param_DecimalColumn" : sCConvDecCol = .Fields("ColumnName").Value
								End Select
								.MoveNext()
							Loop 
							
							If (Len(sCConvTable) > 0) And (Len(sCConvCurrDescCol) > 0) And (Len(sCConvExRateCol) > 0) And (Len(sCConvDecCol) > 0) Then
								'              sCode = "(SELECT ROUND((" & sParamCode1
								'              sCode = sCode & "              / "
								'              sCode = sCode & "             (SELECT " & sCConvTable & "." & sCConvExRateCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode2 & ") "
								'              sCode = sCode & "              * "
								'              sCode = sCode & "             (SELECT " & sCConvTable & "." & sCConvExRateCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & ")) "
								'              sCode = sCode & "        , "
								'              sCode = sCode & "        (SELECT " & sCConvTable & "." & sCConvDecCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & ")) ) "
								
								'AE20071204 Fault #12669
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
						.Close()
					End With
					'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					rsInfo = Nothing
					
					' Last field change date
				Case 52
					objTableView = gcoTablePrivileges.FindTableID(CInt(clsGeneral_Renamed.GetColumnTable(CInt(sParamCode1))))
					sCode = "(SELECT Top 1 DateTimeStamp FROM ASRSysAuditTrail WHERE ColumnID = " & sParamCode1
					sCode = sCode & " And " & IIf(Not pfValidating, objTableView.RealSource & ".", "") & "ID = ASRSysAuditTrail.RecordID ORDER BY DateTimeStamp DESC)"
					
					' Field changed between two dates
				Case 53
					objTableView = gcoTablePrivileges.FindTableID(CInt(clsGeneral_Renamed.GetColumnTable(CInt(sParamCode1))))
					sCode = " case when " & " Exists(Select DateTimeStamp From ASRSysAuditTrail Where ColumnID = " & sParamCode1 & " And " & IIf(Not pfValidating, objTableView.RealSource & ".", "") & "ID = ASRSysAuditTrail.RecordID" & " And DateTimeStamp >= " & sParamCode2 & " And DateTimeStamp <= " & sParamCode3 & " + 1)" & " then 1 else 0 end"
					
					'Whole years between two dates
				Case 54
					sCode = " case " & " when " & sParamCode1 & " >= " & sParamCode2 & " then 0 " & " else " & "   datediff(year, " & sParamCode1 & ", " & sParamCode2 & ") " & "   - " & "   case " & "   when DatePart(Month, " & sParamCode2 & ") < DatePart(Month, " & sParamCode1 & ") " & "   then 1 " & "   else " & "     case " & "     when DatePart(Month, " & sParamCode2 & ") = DatePart(Month, " & sParamCode1 & ") " & "     then " & "       case " & "       when DatePart(Day, " & sParamCode2 & ") < DatePart(Day, " & sParamCode1 & ") " & "       then 1 " & "       else 0 " & "       end " & "     else " & "       0 " & "     end " & "   end " & " end "
					
					' JPD20021121 Fault 3177
				Case 55 ' First Day of Month - VERSION 2 FUNCTION
					sCode = "dateadd(dd, 1 - datepart(dd, " & sParamCode1 & "), " & sParamCode1 & ")"
					
					' JPD20021121 Fault 3177
				Case 56 ' Last Day of Month - VERSION 2 FUNCTION
					sCode = "dateadd(dd, -1, dateadd(mm, 1, dateadd(dd, 1 - datepart(dd, " & sParamCode1 & "), " & sParamCode1 & ")))"
					
					' JPD20021121 Fault 3177
				Case 57 ' First Day of Year - VERSION 2 FUNCTION
					sCode = "dateadd(dd, 1 - datepart(dy, " & sParamCode1 & "), " & sParamCode1 & ")"
					
					' JPD20021121 Fault 3177
				Case 58 ' Last Day of Year - VERSION 2 FUNCTION
					sCode = "dateadd(dd, -1, dateadd(yy, 1, dateadd(dd, 1 - datepart(dy, " & sParamCode1 & "), " & sParamCode1 & ")))"
					
					' JPD20021129 Fault 4337
				Case 59 ' Name of Month. - VERSION 2 FUNCTION
					sCode = "datename(month, " & sParamCode1 & ")"
					
					' JPD20021129 Fault 4337
				Case 60 ' Name of Day. - VERSION 2 FUNCTION
					sCode = "datename(weekday, " & sParamCode1 & ")"
					
					' JPD20021129 Fault 3606
				Case 61 ' Is field populated
					sCode = "(CASE WHEN ((" & sParamCode1 & ") IS NULL)"
					
					' Validate the sub-expression. This is done, not to  validate the expression,
					' but rather to determine the return type of the expression.
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mcolParameters.Item(1).Component.ValidateExpression(False)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters(1).ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Select Case mcolParameters.Item(1).ReturnType
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER
							sCode = sCode & " OR ((" & sParamCode1 & ") = '')"
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC
							sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC
							sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
					End Select
					
					sCode = sCode & " THEN 0 ELSE 1 END)"
					
					'Case 62 'PARENTAL LEAVE ENTITLEMENT
					'Case 63 'PARENTAL LEAVE TAKEN
					'Case 64 'MATERNITY EXPECTED RETURN DATE
					
				Case 65 ' Is Post Subordinate Of
					'If (Len(gsHierarchyTableName) > 0) Then
					'  Set objBaseTable = gcoTablePrivileges.FindTableID(glngHierarchyTableID)
					'  sCode = "CASE WHEN dbo.udf_ASRFn_IsPostSubordinateOf(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id) = 1 THEN 1 ELSE 0 END"
					'  Set objBaseTable = Nothing
					'Else
					'  sCode = "0"
					'End If
					
				Case 66 ' Is Post Subordinate Of User
					If (Len(gsHierarchyTableName) > 0) Then
						objBaseTable = gcoTablePrivileges.FindTableID(glngHierarchyTableID)
						sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_IsPostSubordinateOfUser()) THEN 1 ELSE 0 END"
						'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objBaseTable = Nothing
					Else
						sCode = "0"
					End If
					
				Case 67 ' Is Personnel Subordinate Of
					'If (Len(gsPersonnelTableName) > 0) Then
					'  Set objBaseTable = gcoTablePrivileges.FindTableID(glngPersonnelTableID)
					'  sCode = "CASE WHEN dbo.udf_ASRFn_IsPersonnelSubordinateOf(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id," & sParamCode2 & ") = 1 THEN 1 ELSE 0 END"
					'  Set objBaseTable = Nothing
					'Else
					'  sCode = "0"
					'End If
					
				Case 68 ' Is Personnel Subordinate Of User
					If (Len(gsPersonnelTableName) > 0) Then
						objBaseTable = gcoTablePrivileges.FindTableID(glngPersonnelTableID)
						sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_IsPersonnelSubordinateOfUser()) THEN 1 ELSE 0 END"
						'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objBaseTable = Nothing
					Else
						sCode = "0"
					End If
					
				Case 69 'Has Post Subordinate
					'If (Len(gsHierarchyTableName) > 0) Then
					'  Set objBaseTable = gcoTablePrivileges.FindTableID(glngHierarchyTableID)
					'  sCode = "CASE WHEN dbo.udf_ASRFn_HasPostSubordinate(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id) = 1 THEN 1 ELSE 0 END"
					'  Set objBaseTable = Nothing
					'Else
					'  sCode = "0"
					'End If
					
				Case 70 'Has Post Subordinate User
					If (Len(gsHierarchyTableName) > 0) Then
						objBaseTable = gcoTablePrivileges.FindTableID(glngHierarchyTableID)
						sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_HasPostSubordinateUser()) THEN 1 ELSE 0 END"
						'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objBaseTable = Nothing
					Else
						sCode = "0"
					End If
					
				Case 71 'Has Personnel Subordinate
					'If (Len(gsPersonnelTableName) > 0) Then
					'  Set objBaseTable = gcoTablePrivileges.FindTableID(glngPersonnelTableID)
					'  sCode = "CASE WHEN dbo.udf_ASRFn_HasPersonnelSubordinate(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id," & sParamCode2 & ") = 1 THEN 1 ELSE 0 END"
					'  Set objBaseTable = Nothing
					'Else
					'  sCode = "0"
					'End If
					
				Case 72 'Has Personnel Subordinate User
					If (Len(gsPersonnelTableName) > 0) Then
						objBaseTable = gcoTablePrivileges.FindTableID(glngPersonnelTableID)
						sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_HasPersonnelSubordinateUser()) THEN 1 ELSE 0 END"
						'UPGRADE_NOTE: Object objBaseTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objBaseTable = Nothing
					Else
						sCode = "0"
					End If
					
				Case 73 'Bradford Factor
					If pfValidating Then
						strTempTableName = mobjBaseComponent.ParentExpression.BaseTableName
					Else
						strTempTableName = gcoTablePrivileges.Item(mobjBaseComponent.ParentExpression.BaseTableName).RealSource
					End If
					
					If mobjBaseComponent.ParentExpression.BaseTableID = glngPersonnelTableID Then
						strTempTableID = "ID"
					Else
						strTempTableID = "ID_" & Trim(Str(glngPersonnelTableID))
					End If
					
					sCode = "(dbo.udf_ASRFn_BradfordFactor(" & sParamCode1 & "," & sParamCode2 & "," & sParamCode3 & "," & strTempTableName & "." & strTempTableID & "))"
					
				Case 77 ' Replace characters within a String
					sCode = "REPLACE(" & sParamCode1 & ", " & sParamCode2 & ", " & sParamCode3 & ")"
					
				Case Else
					fOK = False
					
			End Select
		End If
		
		If fOK Then
			' We need to convert date values to varchars in the format 'mm/dd/yyyy'.
			If miReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_DATE Then
				sCode = "convert(" & vbNewLine & "datetime, " & vbNewLine & "convert(" & vbNewLine & "varchar(20), " & vbNewLine & sCode & "," & vbNewLine & "101)" & vbNewLine & ")"
			End If
		End If
		
TidyUpAndExit: 
		If fOK Then
			' JDM - 20/03/02 - Fault 3667 - Needs some brackets around these functions
			'psRuntimeCode = sCode
			psRuntimeCode = "(" & sCode & ")"
		Else
			psRuntimeCode = ""
		End If
		RuntimeCode = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	
	
	Public Function WriteComponent() As Object
		' Write the component definition to the component recordset.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		Dim objParameter As clsExprComponent
		Dim objSubExpression As clsExprExpression
		
		fOK = True
		
		sSQL = "INSERT INTO ASRSysExprComponents" & " (componentID, exprID, type, functionID, valueLogic, ExpandedNode)" & " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & " " & Trim(Str(modExpression.ExpressionComponentTypes.giCOMPONENT_FUNCTION)) & "," & " " & Trim(Str(mlngFunctionID)) & "," & " 0," & IIf(mbExpanded, "1", "0") & ")"
		gADOCon.Execute(sSQL,  , ADODB.CommandTypeEnum.adCmdText)
		
		' Write the function parameter expressions.
		For	Each objParameter In mcolParameters
			objSubExpression = objParameter.Component
			objSubExpression.ParentComponentID = mobjBaseComponent.ComponentID
			objSubExpression.ExpressionID = 0
			fOK = objSubExpression.WriteExpression
			
			'UPGRADE_NOTE: Object objSubExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objSubExpression = Nothing
			
			If Not fOK Then
				Exit For
			End If
		Next objParameter
		'UPGRADE_NOTE: Object objParameter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objParameter = Nothing
		
TidyUpAndExit: 
		'UPGRADE_WARNING: Couldn't resolve default property of object WriteComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		WriteComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	
	Public ReadOnly Property BadComponent() As clsExprComponent
		Get
			' Return the component last caused the function to fail its validity check.
			BadComponent = mobjBadComponent
			
		End Get
	End Property
	
	Public ReadOnly Property ReturnType() As Short
		Get
			' Return the function's return type.
			ReturnType = miReturnType
			
		End Get
	End Property
	
	Public ReadOnly Property ComponentType() As Short
		Get
			' Return the 'function' component type.
			ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_FUNCTION
			
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
		Dim objFunctionCopy As New clsExprFunction
		
		' Copy the component's basic properties.
		' ie. the function id, not its parameters, etc.
		With objFunctionCopy
			.FunctionID = mlngFunctionID
		End With
		
		' JDM - 06/02/01 - Now copies it's children so that cut'n paste works
		' Copy all the child components
		Dim iCount As Short
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
	
	Public Function ValidateFunction() As Short
		' Validate the function. Return a code describing the validity.
		On Error GoTo BasicErrorTrap
		
		Dim iLoop As Short
		Dim iValidationCode As modExpression.ExprValidationCodes
		Dim iFunctionReturnType As modExpression.ExpressionValueTypes
		Dim sSQL As String
		Dim rsParameters As ADODB.Recordset
		Dim aiDummyValues(6) As Short
		Dim objSubExpression As clsExprExpression
		Dim objParameter As clsExprComponent
		
		iLoop = 0
		
		' Initialise the validation code.
		iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS
		'UPGRADE_NOTE: Object mobjBadComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjBadComponent = Nothing
		
		' Validate the function parameter expressions.
		For	Each objParameter In mcolParameters
			iLoop = iLoop + 1
			
			objSubExpression = objParameter.Component
			With objSubExpression
				' Validate the parameter expression.
				' NB. Reset the sub-expression's return type to that defined by the parameter definition
				' as it may be changeable. The evaluated return type will be determined when the
				' sub-expression is validated.
				sSQL = "SELECT parameterType FROM ASRSysFunctionParameters" & " WHERE functionID = " & Trim(Str(mlngFunctionID)) & " AND parameterIndex = " & Trim(Str(iLoop))
				rsParameters = datGeneral.GetRecords(sSQL)
				With rsParameters
					If Not (.BOF And .EOF) Then
						objSubExpression.ReturnType = .Fields("parameterType").Value
					End If
					
					.Close()
				End With
				'UPGRADE_NOTE: Object rsParameters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsParameters = Nothing
				
				iValidationCode = .ValidateExpression(False)
				
				If iValidationCode <> modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
					' Interpret the parameter sub-expression validation code to reflect
					' the fact that a function parameter was invalid.
					Select Case iValidationCode
						Case modExpression.ExprValidationCodes.giEXPRVALIDATION_NOCOMPONENTS
							iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_PARAMETERNOCOMPONENTS
						Case modExpression.ExprValidationCodes.giEXPRVALIDATION_SYNTAXERROR
							iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_PARAMETERSYNTAXERROR
						Case modExpression.ExprValidationCodes.giEXPRVALIDATION_EXPRTYPEMISMATCH
							iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_PARAMETERTYPEMISMATCH
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
		
		If iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
			If Not ValidateFunctionParameters(mlngFunctionID, iFunctionReturnType, aiDummyValues(1), aiDummyValues(2), aiDummyValues(3), aiDummyValues(4), aiDummyValues(5), aiDummyValues(6)) Then
				
				iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_PARAMETERTYPEMISMATCH
			Else
				miReturnType = iFunctionReturnType
			End If
		End If
		
TidyUpAndExit: 
		' Disassociate object variables.
		'UPGRADE_NOTE: Object objSubExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objSubExpression = Nothing
		'UPGRADE_NOTE: Object rsParameters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsParameters = Nothing
		ValidateFunction = iValidationCode
		Exit Function
		
BasicErrorTrap: 
		iValidationCode = modExpression.ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR
		Resume TidyUpAndExit
		
	End Function
	
	
	Private Function ReadFunction() As Boolean
		' Read the function definition from the database.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		' Read the function name, return type, etc.
		fOK = ReadFunctionDetails
		
		If fOK Then
			' Create the array of parameter components.
			fOK = ReadParameterDefinition
		End If
		
TidyUpAndExit: 
		ReadFunction = False
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Private Function ReadFunctionDetails() As Boolean
		' Read the function details (not parameter info) from the database.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		Dim rsFunction As ADODB.Recordset
		
		msFunctionName = "<unknown>"
		miReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED
		msSPName = vbNullString
		
		' Clear the parameter collection.
		ClearParameters()
		
		' Get the function definition.
		sSQL = "SELECT *" & " FROM ASRSysFunctions" & " WHERE functionID = " & Trim(Str(mlngFunctionID))
		rsFunction = datGeneral.GetRecords(sSQL)
		With rsFunction
			fOK = Not (.EOF And .BOF)
			
			If fOK Then
				msFunctionName = .Fields("functionName").Value
				miReturnType = .Fields("ReturnType").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				msSPName = IIf(IsDbNull(.Fields("SPName").Value), "", .Fields("SPName").Value)
			End If
			
			.Close()
		End With
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object rsFunction may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsFunction = Nothing
		ReadFunctionDetails = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	Private Function ReadParameterDefinition() As Boolean
		' Read the function's paramter definition from the database,
		' and create an array of components to represent the parameters.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim iIndex As Short
		Dim sSQL As String
		Dim rsParameters As ADODB.Recordset
		Dim objNewParameter As clsExprComponent
		
		fOK = True
		
		' Clear the parameter collection.
		ClearParameters()
		
		' Get the standard function parameter definitions.
		sSQL = "SELECT *" & " FROM ASRSysFunctionParameters" & " WHERE functionID = " & Trim(Str(mlngFunctionID)) & " ORDER BY parameterIndex"
		rsParameters = datGeneral.GetRecords(sSQL)
		With rsParameters
			Do While Not .EOF
				' Instantiate a component in the array to represent the parameter.
				mcolParameters.Add(New clsExprComponent)
				With mcolParameters.Item(mcolParameters.Count())
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_EXPRESSION
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item(mcolParameters.Count).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Component.Name = rsParameters.Fields("parameterName").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item(mcolParameters.Count).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Component.ReturnType = rsParameters.Fields("parameterType").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID
				End With
				
				.MoveNext()
			Loop 
			
			.Close()
		End With
		
		' Get the customised function parameter definitions if they exist.
		If (mobjBaseComponent.ComponentID > 0) Then
			
			iIndex = 1
			sSQL = "SELECT *" & " FROM ASRSysExpressions" & " WHERE parentComponentID = " & Trim(Str(mobjBaseComponent.ComponentID)) & " ORDER BY exprID"
			rsParameters = datGeneral.GetRecords(sSQL)
			With rsParameters
				Do While (Not .EOF) And fOK
					
					' Instantiate a new component object.
					objNewParameter = New clsExprComponent
					
					' Construct the hierarchy of objects that define the parameter.
					objNewParameter.ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_EXPRESSION
					'UPGRADE_WARNING: Couldn't resolve default property of object objNewParameter.Component.ExpressionID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					objNewParameter.Component.ExpressionID = .Fields("ExprID").Value
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
						
						'JDM - 16/03/01 - Fault 1935 - Load previous view
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						objNewParameter.ExpandedNode = IIf(IsDbNull(.Fields("ExpandedNode").Value), False, .Fields("ExpandedNode").Value)
						
						' Insert the new expression into the function's parameter array.
						mcolParameters.Remove(iIndex)
						If mcolParameters.Count() >= iIndex Then
							mcolParameters.Add(objNewParameter,  , iIndex)
						Else
							mcolParameters.Add(objNewParameter)
						End If
					End If
					
					iIndex = iIndex + 1
					'UPGRADE_NOTE: Object objNewParameter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objNewParameter = Nothing
					
					.MoveNext()
				Loop 
			End With
		End If
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object rsParameters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsParameters = Nothing
		ReadParameterDefinition = fOK
		Exit Function
		
ErrorTrap: 
		' Clear the parameter collection.
		ClearParameters()
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Sub ClearParameters()
		' Clear the function's parameter collection.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		' Remove all components from the collection.
		Do While mcolParameters.Count() > 0
			mcolParameters.Remove(1)
		Loop 
		'UPGRADE_NOTE: Object mcolParameters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolParameters = Nothing
		
		' Re-instantiate the collection.
		mcolParameters = New Collection
		Exit Sub
		
ErrorTrap: 
		fOK = False
		
	End Sub
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		' Create a new collection to hold the function's parameters.
		mcolParameters = New Collection
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		' Disassociate object variables.
		'UPGRADE_NOTE: Object mcolParameters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolParameters = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function UDFCode(ByRef psRuntimeCode() As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		Dim sParamCode1 As String
		Dim sParamCode2 As String
		Dim sParamCode3 As String
		Dim sParamCode4 As String
		Dim fOK As Boolean
		
		'JPD 20031031 Fault 7440
		fOK = True
		
		' Get the first parameter's runtime code if required.
		If mcolParameters.Count() >= 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mcolParameters.Item(1).Component.UDFCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
		End If
		
		' Get the second parameter's runtime code if required.
		If fOK And (mcolParameters.Count() >= 2) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mcolParameters.Item(2).Component.UDFCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
		End If
		
		' Get the third parameter's runtime code if required.
		If fOK And (mcolParameters.Count() >= 3) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mcolParameters.Item(3).Component.UDFCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
		End If
		
		' Get the fourth parameter's runtime code if required.
		If fOK And (mcolParameters.Count() >= 4) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolParameters.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mcolParameters.Item(4).Component.UDFCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
		End If
		
		UDFCode = fOK
		
	End Function
End Class