VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngFunctionID As Long
Private msFunctionName As String
Private miReturnType As ExpressionValueTypes
Private msSPName As String

' Definition for expanded/unexpanded status of the component
Private mbExpanded As Boolean

' Class handling variables.
Private mobjBaseComponent As clsExprComponent
Private mcolParameters As Collection
Private mobjBadComponent As clsExprComponent

Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  Dim objParameter As clsExprComponent
  Dim objSubExpression As clsExprExpression
  
  ContainsExpression = False
  
  For Each objParameter In mcolParameters
    Set objSubExpression = objParameter.Component
    
    ContainsExpression = objSubExpression.ContainsExpression(plngExprID)
        
    Set objSubExpression = Nothing
    
    If ContainsExpression Then
      Exit For
    End If
  Next objParameter
  Set objParameter = Nothing
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  COAMsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function







Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objParameter As clsExprComponent
  
  fOK = True
  
  ' Position the printing.
  With Printer
    .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    Printer.Print "Function : " & ComponentDescription
  End With
  
  ' Print the function's parameter expressions.
  For Each objParameter In mcolParameters
    objParameter.Component.PrintComponent (piLevel + 1)
  Next objParameter
  
TidyUpAndExit:
  Set objParameter = Nothing
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function




Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  ' Return the SQL code for the component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim fSrchColumnOK As Boolean
  Dim fRtnColumnOK As Boolean
  Dim iLoop As Integer
  Dim lngSrchTableID As Long
  Dim lngRtnTableID As Long
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
  Dim rsInfo As Recordset
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
  
  Dim clsGeneral As HRProDataMgr.clsGeneral
  Set clsGeneral = New HRProDataMgr.clsGeneral
  
  fOK = True
  sCode = ""
  
  ' Get the first parameter's runtime code if required.
  If mcolParameters.Count >= 1 Then
    fOK = mcolParameters.Item(1).Component.RuntimeCode(sParamCode1, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If
  
  ' Get the second parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 2) Then
    fOK = mcolParameters.Item(2).Component.RuntimeCode(sParamCode2, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If
  
  ' Get the third parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 3) Then
    fOK = mcolParameters.Item(3).Component.RuntimeCode(sParamCode3, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If
    
  ' Get the fourth parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 4) Then
    fOK = mcolParameters.Item(4).Component.RuntimeCode(sParamCode4, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If
    
  If fOK Then
    Select Case mlngFunctionID
      Case 1  ' System date
        sCode = "getdate()"
          
      Case 2  ' Convert to uppercase
        sCode = "upper(" & sParamCode1 & ")"
        
      Case 3  ' Convert numeric to string
        sCode = "IsNull(ltrim(str(" & sParamCode1 & ", 255, " & sParamCode2 & ")),'')"
  
      Case 4  ' If... Then... Else...
        sCode = "CASE WHEN (" & sParamCode1 & " = 1) THEN " & sParamCode2 & " ELSE " & sParamCode3 & " END"
    
      Case 5  ' Remove leading and trailing spaces
        sCode = "ltrim(rtrim(" & sParamCode1 & "))"
        
      Case 6  ' Extract characters from the left
        sCode = "left(" & sParamCode1 & ", " & sParamCode2 & ")"
        
      Case 7  ' Length of character field
        sCode = "len(" & sParamCode1 & ")"
        
      Case 8  ' Convert to lowercase
        sCode = "lower(" & sParamCode1 & ")"
  
      Case 9  ' Maximum
        sCode = "CASE WHEN (" & sParamCode1 & " > " & sParamCode2 & ") THEN " & sParamCode1 & " ELSE " & sParamCode2 & " END"
      
      Case 10  ' Minimum
        sCode = "CASE WHEN (" & sParamCode1 & " < " & sParamCode2 & ") THEN " & sParamCode1 & " ELSE " & sParamCode2 & " END"
      
      Case 11  ' Search for character string.
        sCode = "charindex(" & sParamCode2 & ", " & sParamCode1 & ")"
        
      Case 12  ' Capitalise Initials
        sCode = "(dbo.udf_ASRFn_CapitalizeInitials(" & sParamCode1 & "))"
        
      Case 13  ' Extract characters from the right
        sCode = "right(" & sParamCode1 & ", " & sParamCode2 & ")"
        
      Case 14  ' Extract part of a character string
        sCode = "substring(" & sParamCode1 & ", " & sParamCode2 & ", " & sParamCode3 & ")"
  
      Case 15  ' System Time
        sCode = "convert(varchar(50), getdate(), 8)"
    
      Case 16  ' Is field empty
        sCode = "(CASE WHEN ((" & sParamCode1 & ") IS NULL)"

        ' Validate the sub-expression. This is done, not to  validate the expression,
        ' but rather to determine the return type of the expression.
        mcolParameters(1).Component.ValidateExpression False
      
        Select Case mcolParameters(1).ReturnType
          Case giEXPRVALUE_CHARACTER
            sCode = sCode & " OR ((" & sParamCode1 & ") = '')"
          Case giEXPRVALUE_NUMERIC
            sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
          Case giEXPRVALUE_LOGIC
            sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
        End Select
        
        sCode = sCode & " THEN 1 ELSE 0 END)"
  
      Case 17  ' Current user
'        sCode = "user_name()"
        sCode = "system_user"
        
      Case 18  ' Whole Years Until Current Date
        sCode = "datediff(year," & sParamCode1 & ", getdate())" & _
          " - case" & _
          "       when datepart(month," & sParamCode1 & ") > datepart(month, getdate()) then 1" & _
          "       when (datepart(month," & sParamCode1 & ") = datepart(month, getdate())) " & _
          "           and (datepart(day," & sParamCode1 & ") > datepart(day, getdate())) then 1" & _
          "       else 0" & _
          "   end"
      
      Case 19  ' Remaining Months Since Whole Years
        sCode = "datepart(month, getdate())" & _
          " - datepart(month, " & sParamCode1 & ")" & _
          " - case" & _
          "       when datepart(day," & sParamCode1 & ") > datepart(day, getdate()) then 1" & _
          "       else 0" & _
          "   end" & _
          " + case" & _
          "       when (datepart(month, getdate())" & _
          "           - datepart(month," & sParamCode1 & ")" & _
          "           - case" & _
          "               when datepart(day," & sParamCode1 & ") > datepart(day, getdate()) then 1" & _
          "               else 0" & _
          "             end) < 0 then 12" & _
          "       else 0" & _
          "   end"
  
      Case 20  ' Capitalise Initials
        sCode = "(dbo.udf_ASRFn_InitialsFromForenames(" & sParamCode1 & "))"
  
      Case 21  ' First Name from Forenames
        sCode = _
          "case" & _
          "    when charindex(' ', ltrim(" & sParamCode1 & ")) > 0 then substring(ltrim(" & sParamCode1 & "), 1, charindex(' ', ltrim(" & sParamCode1 & "))-1)" & _
          "    else ltrim(" & sParamCode1 & ")" & _
          "end"
      
      Case 22  ' Weekdays From Start and End Dates
        
          'MH20060424 Fault 11054
          '"    when datediff(day, " & sParamCode1 & ", " & sParamCode2 & ") <= 0 then 0" & _

        sCode = _
          " case" & _
          "    when datediff(day, " & sParamCode1 & ", " & sParamCode2 & ") < 0 then 0" & _
          "    else datediff(day, " & sParamCode1 & ", " & sParamCode2 & " + 1)" & _
          "        - (2 * (datediff(day, " & sParamCode1 & " - (datepart(dw, " & sParamCode1 & ") - 1), " & _
          "                              " & sParamCode2 & " - (datepart(dw, " & sParamCode2 & ") - 1)) / 7))" & _
          "        - case" & _
          "              when datepart(dw, " & sParamCode1 & ") = 1 then 1" & _
          "              else 0" & _
          "          end" & _
          "        - case" & _
          "              when datepart(dw, " & sParamCode2 & ") = 7 then 1" & _
          "              else 0" & _
          "          end" & _
          " end"
    
      Case 23  ' Add months to date
        sCode = "dateadd(month, " & sParamCode2 & ", " & sParamCode1 & ")"
        
      Case 24  ' Add years to date
        sCode = "dateadd(year, " & sParamCode2 & ", " & sParamCode1 & ")"
  
      Case 25  ' Convert character to numeric.
        'JPD 20041213 Fault 9568
        'sCode = _
          " case" & _
          "    when isnumeric(" & sParamCode1 & ") = 1 then convert(float, " & sParamCode1 & ")" & _
          "    else 0" & _
          " end"
        sCode = _
          " case" & _
          "    when isnumeric(" & sParamCode1 & ") = 1 then convert(float, convert(money, " & sParamCode1 & "))" & _
          "    else 0" & _
          " end"
      
      Case 26  ' Whole Months between 2 Dates.
        sCode = _
          " case" & _
          "    when " & sParamCode1 & " >= " & sParamCode2 & " then 0" & _
          "    else datediff(month, " & sParamCode1 & ", " & sParamCode2 & ")" & _
          "        - case" & _
          "              when datepart(day, " & sParamCode2 & ") < datepart(day, " & sParamCode1 & ") then 1" & _
          "              else 0" & _
          "          end" & _
          " end"
    
      Case 27  ' Parentheses
        sCode = sParamCode1
  
      Case 28  ' Day of the week
        sCode = "DATEPART(weekday, " & sParamCode1 & ")"
      
      Case 29  ' Working Days per week
        sCode = "(convert(float, len(replace(left(" & sParamCode1 & ", 14), ' ', ''))) / 2)"
    
      Case 30  ' Absence Duration
      
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
        
        sCode = "(dbo.udf_ASRFn_AbsenceDuration(" & sParamCode1 & "," & sParamCode2 _
            & "," & sParamCode3 & "," & sParamCode4 & "," & strTempTableName & "." & strTempTableID & "))"
    
      Case 31  ' Round down to nearest whole number.
        sCode = "round(" & sParamCode1 & ", 0, 1)"
        
      Case 32  ' Year of date.
        sCode = "datepart(year, " & sParamCode1 & ")"
        
      Case 33  ' Month of date.
        sCode = "datepart(month, " & sParamCode1 & ")"
        
      Case 34  ' Day of date.
        sCode = "datepart(day, " & sParamCode1 & ")"
        
      Case 35  ' Nice Date
        sCode = _
          "datename(day, " & sParamCode1 & ") + ' ' + " & _
          "datename(month, " & sParamCode1 & ") + ' ' + " & _
          "datename(year, " & sParamCode1 & ")"
      
      Case 36  ' Nice Time
'        sCode = _
          "convert(varchar(2), datepart(hour, convert(datetime, " & sParamCode1 & ")) % 12) + ':' " & _
          "    + right('00' + datename(minute, convert(datetime, " & sParamCode1 & ")),2)" & _
          "    + case" & _
          "          when datepart(hour, convert(datetime, " & sParamCode1 & ")) > 11 then ' pm'" & _
          "          else ' am'" & _
          "      end"
        ' JPD20020617 Fault 3998
        sCode = _
          "case when len(ltrim(rtrim(" & sParamCode1 & "))) = 0 then ''" & _
          " else case when isdate(" & sParamCode1 & ") = 0 then '***'" & _
          " else (convert(varchar(2),((datepart(hour,convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)) + 11) % 12) + 1)" & _
          " + ':' + right('00' + datename(minute, convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)),2)" & _
          " + case when datepart(hour, convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)) > 11 then ' pm'" & _
          " else ' am' end) end end"
      
      Case 37  ' Round Date to Start of nearest month
        sCode = _
          " case" & _
          "     when datediff(day, (" & sParamCode1 & " - datepart(day, " & sParamCode1 & ") + 1), " & sParamCode1 & ")" & _
          "         <= datediff(day, " & sParamCode1 & ", (dateadd(month, 1, " & sParamCode1 & ") - datepart(day, dateadd(month, 1, " & sParamCode1 & ")) + 1))" & _
          "         then " & sParamCode1 & " - datepart(day, " & sParamCode1 & ") + 1" & _
          "     else dateadd(month, 1, " & sParamCode1 & ")" & _
          "         - datepart(day, dateadd(month, 1, " & sParamCode1 & ")) + 1" & _
          " end"
      
      Case 38  ' Is Between
        sCode = _
          " case" & _
          "     when (" & sParamCode1 & " >= " & sParamCode2 & ")" & _
          "         and (" & sParamCode1 & " <= " & sParamCode3 & ") then 1" & _
          "     else 0" & _
          " end"
      
      Case 39  ' Service Years
        sCode = _
          "   datepart(year, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)" & _
          " - datepart(year, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & _
          " - case" & _
          "       when datepart(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & _
          "          > datepart(month, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end) then 1" & _
          "      when (datepart(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & _
          "          = datepart(month, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end))" & _
          "       and (datepart(day, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & _
          "          > datepart(day, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)) then 1" & _
          "      else 0" & _
          "  end"
      
      Case 40  ' Service Months
        sCode = _
          " (case" & _
          "    when case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end" & _
          "          >= case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end then 0" & _
          "    else datediff(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)" & _
          "        - case" & _
          "              when datepart(day, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end) < datepart(day, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end) then 1" & _
          "              else 0" & _
          "          end" & _
          " end) % 12"
      
      Case 42  ' Get field from database record.
        ' Get the column parameter definitions.
        sSQL = "SELECT ASRSysColumns.columnID, ASRSysColumns.columnName, ASRSysTables.tableID, ASRSysTables.tableName" & _
          " FROM ASRSysColumns" & _
          " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
          " WHERE ASRSysColumns.columnID IN (" & sParamCode1 & ", " & sParamCode3 & ")"
        Set rsInfo = datGeneral.GetRecords(sSQL)
        With rsInfo
          Do While Not .EOF
            If Trim(Str(!ColumnID)) = sParamCode1 Then
              sSrchColumnName = !ColumnName
              sSrchTableName = !TableName
              lngSrchTableID = !TableID
            End If
            
            If Trim(Str(!ColumnID)) = sParamCode3 Then
              sRtnColumnName = !ColumnName
              lngRtnTableID = !TableID
            End If
            
            .MoveNext
          Loop
              
          .Close
        End With
        Set rsInfo = Nothing
        
        fOK = ((Len(sSrchColumnName) > 0) And (Len(sRtnColumnName) > 0)) And _
          (lngSrchTableID = lngRtnTableID)
    
        ' Construct the select statement to get the required field from the given table,
        ' incorporating permissions.
        If fOK Then
          ReDim asViews(2, 0)
          
          ' Check permissions on that return column.
          Set objColumnPrivileges = GetColumnPrivileges(sSrchTableName)
          sRealTableSource = gcoTablePrivileges.Item(sSrchTableName).RealSource
        
          fRtnColumnOK = objColumnPrivileges.IsValid(sRtnColumnName)
        
          If fRtnColumnOK Then
            fRtnColumnOK = objColumnPrivileges.Item(sRtnColumnName).AllowSelect
          End If
        
          If fRtnColumnOK Then
            ' The search column can be read direct from the table.
            sRtnColumnCode = sRealTableSource & "." & sRtnColumnName
          Else
            ' Then column cannot be read direct. If its from a parent, try parent views
            ' Loop thru the views on the table, seeing if any have read permis for the column
            ' Column 1 = view name
            ' Column 2 = "S" if the view is used for the search.
            For Each objTableView In gcoTablePrivileges.Collection
              If (Not objTableView.IsTable) And _
                (objTableView.TableID = lngSrchTableID) And _
                (objTableView.AllowSelect) Then
        
                ' Get the column permission for the view
                Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
        
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
            Set objTableView = Nothing
        
            ' Does the user have select permission thru ANY views ?
            If UBound(asViews, 2) = 0 Then
              ' The search column can be read neither from the table nor any views.
              sRtnColumnCode = "NULL"
            Else
              For iLoop = 1 To UBound(asViews, 2)
                If iLoop = 1 Then
                  sRtnColumnCode = "CASE"
                End If
        
                sRtnColumnCode = sRtnColumnCode & _
                  " WHEN NOT " & asViews(1, iLoop) & "." & sRtnColumnName & " IS NULL THEN " & asViews(1, iLoop) & "." & sRtnColumnName
              Next iLoop
        
              sRtnColumnCode = sRtnColumnCode & _
                " ELSE NULL" & _
                " END"
            End If
          End If
          
          
          Set objColumnPrivileges = GetColumnPrivileges(sSrchTableName)
          
          
          fSrchColumnOK = objColumnPrivileges.IsValid(sSrchColumnName)
        
          If fSrchColumnOK Then
            fSrchColumnOK = objColumnPrivileges.Item(sSrchColumnName).AllowSelect
          End If
        
          If fSrchColumnOK Then
            ' The search column can be read direct from the table.
            sSrchColumnCode = sRealTableSource & "." & sSrchColumnName & " = " & sParamCode2
          Else
            ' Then column cannot be read direct. If its from a parent, try parent views
            ' Loop thru the views on the table, seeing if any have read permis for the column
            ' Column 1 = view name
            ' Column 2 = "S" if the view is used for the search.
            For Each objTableView In gcoTablePrivileges.Collection
              If (Not objTableView.IsTable) And _
                (objTableView.TableID = lngSrchTableID) And _
                (objTableView.AllowSelect) Then
        
                ' Get the column permission for the view
                Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
        
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
            Set objTableView = Nothing
        
            sSrchColumnCode = ""
            For iLoop = 1 To UBound(asViews, 2)
              If asViews(2, iLoop) = "S" Then
                fSrchColumnOK = True
              
                If Len(sSrchColumnCode) > 0 Then
                  sSrchColumnCode = sSrchColumnCode & " OR "
                End If
          
                sSrchColumnCode = sSrchColumnCode & _
                  sRealTableSource & ".id IN (SELECT id FROM " & asViews(1, iLoop) & " WHERE " & sSrchColumnName & " = " & sParamCode2 & ")"
              End If
            Next iLoop
          End If
          
          sSrchTableCode = ""
          For iLoop = 1 To UBound(asViews, 2)
            sSrchTableCode = sSrchTableCode & _
              " LEFT OUTER JOIN " & asViews(1, iLoop) & " ON " & sRealTableSource & ".id = " & asViews(1, iLoop) & ".id"
          Next iLoop
          
          If fSrchColumnOK Then
            sCode = "(SELECT TOP 1 " & sRtnColumnCode & _
              " FROM " & sRealTableSource & _
              " " & sSrchTableCode & _
              " WHERE (" & sSrchColumnCode & "))"
          Else
            sCode = "null"
          End If
        End If

      Case 44  ' Add days to date.
        sCode = "dateadd(day, " & sParamCode2 & ", " & sParamCode1 & ")"
      
      Case 45  ' Days Between 2 Dates
            
        'MH20010220 Fault 1850
        'This function needs to be inclusive of both start and end
        'sCode = "datediff(dd, " & sParamCode1 & ", " & sParamCode2 & ")"
        sCode = "datediff(dd, " & sParamCode1 & ", " & sParamCode2 & ")+1"

      Case 46  'Working days between two dates
        
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

      Case 47  ' Absence between two dates
        
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
        'JPD 20041118 Fault 8231
        sCode = "(dbo.udf_ASRFn_AbsenceBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 _
            & "," & sParamCode3 & "," & strTempTableName & "." & strTempTableID & "," & "convert(datetime,'" & Replace(Format(Now(), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')" & "))"

      Case 48  ' Round Up to nearest whole number.
        ' JPD20030116 Fault 4910
        'sCode = "ceiling(" & sParamCode1 & ")"
        sCode = "CASE WHEN (" & sParamCode1 & ") < 0 THEN floor(" & sParamCode1 & ")" & _
          " ELSE ceiling(" & sParamCode1 & ") END"

      Case 49  ' Round to nearest number.
        ' JPD20020415 Fault 3701
        ' Changed 'division by 2' to 'division by 2.0' to avoid SQL casting the result to an integer value.
        strRemainString = "(" + sParamCode1 + ") - ((floor(" + sParamCode1 + "/" + sParamCode2 + "))*" + sParamCode2 + ")"
        ' JPD20030116 Fault 4910
        'sCode = "CASE WHEN (" + strRemainString + ")<((" + sParamCode2 + ")/2.0)" & _
          " THEN (" + sParamCode1 + ")-(" + strRemainString + ")" & _
          " ELSE (" + sParamCode1 + ")+(" + sParamCode2 + ")-(" + strRemainString + ") END"
        sCode = "CASE WHEN (((" + sParamCode1 + ")<0) AND ((" + strRemainString + ")<=((" + sParamCode2 + ")/2.0)))" & _
          " OR (((" + sParamCode1 + ")>=0) AND ((" + strRemainString + ")<((" + sParamCode2 + ")/2.0)))" & _
          " THEN (" + sParamCode1 + ")-(" + strRemainString + ")" & _
          " ELSE (" + sParamCode1 + ")+(" + sParamCode2 + ")-(" + strRemainString + ") END"
        
        'MH20100629
        sCode = "CASE WHEN (" & sParamCode2 & " > 0) THEN " & sCode & " ELSE 0 END"

      'TM20011022 Currency Implementation
      Case 51
      '*********** runtime code to go here *************

        ' Get the column parameter definitions.
        sSQL = vbNullString
        sSQL = sSQL & "SELECT ASRSysModuleSetup.*, ASRSysColumns.ColumnName, ASRSysTables.TableName "
        sSQL = sSQL & "FROM ASRSysModuleSetup "
        sSQL = sSQL & "     INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID "
        sSQL = sSQL & "                INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID "
        sSQL = sSQL & "WHERE ASRSysModuleSetup.ModuleKey = 'MODULE_CURRENCY' "
        
        Set rsInfo = datGeneral.GetRecords(sSQL)
        sSQL = vbNullString
        
        With rsInfo
          If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do While Not .EOF
              sCConvTable = !TableName
              Select Case !ParameterKey
              Case "Param_CurrencyNameColumn": sCConvCurrDescCol = !ColumnName
              Case "Param_ConversionValueColumn": sCConvExRateCol = !ColumnName
              Case "Param_DecimalColumn": sCConvDecCol = !ColumnName
              End Select
              .MoveNext
            Loop
                    
            If (Len(sCConvTable) > 0) And (Len(sCConvCurrDescCol) > 0) And _
                  (Len(sCConvExRateCol) > 0) And (Len(sCConvDecCol) > 0) Then
'              sCode = vbNullString
'              sCode = sCode & " ( CASE"
'              sCode = sCode & "  WHEN ((SELECT Currency_Conversion_Table.Units_per_Euro"
'              sCode = sCode & "                From Currency_Conversion_Table"
'              sCode = sCode & "                WHERE Currency_Conversion_Table.Currency_Name = " & sParamCode2 & ") <> 0) "
'              sCode = sCode & "  THEN (SELECT ROUND((" & sParamCode1
'              sCode = sCode & "              / "
'              sCode = sCode & "             (SELECT " & sCConvTable & "." & sCConvExRateCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode2 & ") "
'              sCode = sCode & "              * "
'              sCode = sCode & "             (SELECT " & sCConvTable & "." & sCConvExRateCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & ")) "
'              sCode = sCode & "        , "
'              sCode = sCode & "        (SELECT " & sCConvTable & "." & sCConvDecCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & ")) ) "
'              sCode = sCode & "  ELSE "
'              sCode = sCode & "    0"
'              sCode = sCode & "  END )"
              
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
          .Close
        End With
        Set rsInfo = Nothing

      ' Last field change date
      Case 52
        Set objTableView = gcoTablePrivileges.FindTableID(CLng(clsGeneral.GetColumnTable(CLng(sParamCode1))))
        sCode = "(SELECT Top 1 DateTimeStamp FROM ASRSysAuditTrail WHERE ColumnID = " & sParamCode1
        
        sCode = sCode & " And " & IIf(Not pfValidating, objTableView.RealSource & ".", "") & "ID = ASRSysAuditTrail.RecordID"
      
        sCode = sCode & " ORDER BY DateTimeStamp Desc)"
      
      ' Field changed between two dates
      Case 53
        Set objTableView = gcoTablePrivileges.FindTableID(CLng(clsGeneral.GetColumnTable(CLng(sParamCode1))))
        sCode = " case when " _
          & " Exists(Select DateTimeStamp From ASRSysAuditTrail Where ColumnID = " & sParamCode1 _
          & " And " & IIf(Not pfValidating, objTableView.RealSource & ".", "") & "ID = ASRSysAuditTrail.RecordID" _
          & " And DateTimeStamp >= " & sParamCode2 & " And DateTimeStamp < " & sParamCode3 & " + 1)" _
          & " then 1 else 0 end"
      
      'Whole years between two dates
      Case 54
        sCode = _
          " case " & _
          " when " & sParamCode1 & " >= " & sParamCode2 & _
          " then 0 " & _
          " else " & _
          "   datediff(year, " & sParamCode1 & ", " & sParamCode2 & ") " & _
          "   - " & _
          "   case " & _
          "   when DatePart(Month, " & sParamCode2 & ") < DatePart(Month, " & sParamCode1 & ") " & _
          "   then 1 " & _
          "   else " & _
          "     case " & _
          "     when DatePart(Month, " & sParamCode2 & ") = DatePart(Month, " & sParamCode1 & ") " & _
          "     then " & _
          "       case " & _
          "       when DatePart(Day, " & sParamCode2 & ") < DatePart(Day, " & sParamCode1 & ") " & _
          "       then 1 " & _
          "       else 0 " & _
          "       end " & _
          "     else " & _
          "       0 " & _
          "     end " & _
          "   end " & _
          " end "
      
      ' JPD20021121 Fault 3177
      Case 55  ' First Day of Month
        sCode = "dateadd(dd, 1 - datepart(dd, " & sParamCode1 & "), " & sParamCode1 & ")"
        
      ' JPD20021121 Fault 3177
      Case 56  ' Last Day of Month
        sCode = "dateadd(dd, -1, dateadd(mm, 1, dateadd(dd, 1 - datepart(dd, " & sParamCode1 & "), " & sParamCode1 & ")))"
        
      ' JPD20021121 Fault 3177
      Case 57  ' First Day of Year
        sCode = "dateadd(dd, 1 - datepart(dy, " & sParamCode1 & "), " & sParamCode1 & ")"
        
      ' JPD20021121 Fault 3177
      Case 58  ' Last Day of Year
        sCode = "dateadd(dd, -1, dateadd(yy, 1, dateadd(dd, 1 - datepart(dy, " & sParamCode1 & "), " & sParamCode1 & ")))"
      
      ' JPD20021129 Fault 4337
      Case 59  ' Name of Month.
        sCode = "datename(month, " & sParamCode1 & ")"
      
      ' JPD20021129 Fault 4337
      Case 60  ' Name of Day.
        sCode = "datename(weekday, " & sParamCode1 & ")"
      
      ' JPD20021129 Fault 3606
      Case 61  ' Is field populated
        sCode = "(CASE WHEN ((" & sParamCode1 & ") IS NULL)"

        ' Validate the sub-expression. This is done, not to  validate the expression,
        ' but rather to determine the return type of the expression.
        mcolParameters(1).Component.ValidateExpression False
      
        Select Case mcolParameters(1).ReturnType
          Case giEXPRVALUE_CHARACTER
            sCode = sCode & " OR ((" & sParamCode1 & ") = '')"
          Case giEXPRVALUE_NUMERIC
            sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
          Case giEXPRVALUE_LOGIC
            sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
        End Select
        
        sCode = sCode & " THEN 0 ELSE 1 END)"
  
      'Case 62 'PARENTAL LEAVE ENTITLEMENT
      'Case 63 'PARENTAL LEAVE TAKEN
      'Case 64 'MATERNITY EXPECTED RETURN DATE
      
      Case 65  ' Is Post Subordinate Of
        'If (Len(gsHierarchyTableName) > 0) Then
        '  Set objBaseTable = gcoTablePrivileges.FindTableID(glngHierarchyTableID)
        '  sCode = "CASE WHEN dbo.udf_ASRFn_IsPostSubordinateOf(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id) = 1 THEN 1 ELSE 0 END"
        '  Set objBaseTable = Nothing
        'Else
        '  sCode = "0"
        'End If
        
      Case 66  ' Is Post Subordinate Of User
        If (Len(gsHierarchyTableName) > 0) Then
          Set objBaseTable = gcoTablePrivileges.FindTableID(glngHierarchyTableID)
          sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_IsPostSubordinateOfUser()) THEN 1 ELSE 0 END"
          Set objBaseTable = Nothing
        Else
          sCode = "0"
        End If
        
      Case 67  ' Is Personnel Subordinate Of
        'If (Len(gsPersonnelTableName) > 0) Then
        '  Set objBaseTable = gcoTablePrivileges.FindTableID(glngPersonnelTableID)
        '  sCode = "CASE WHEN dbo.udf_ASRFn_IsPersonnelSubordinateOf(" & sParamCode1 & ", " & objBaseTable.RealSource & ".id," & sParamCode2 & ") = 1 THEN 1 ELSE 0 END"
        '  Set objBaseTable = Nothing
        'Else
        '  sCode = "0"
        'End If
      
      Case 68  ' Is Personnel Subordinate Of User
        If (Len(gsPersonnelTableName) > 0) Then
          Set objBaseTable = gcoTablePrivileges.FindTableID(glngPersonnelTableID)
          sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_IsPersonnelSubordinateOfUser()) THEN 1 ELSE 0 END"
          Set objBaseTable = Nothing
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
          Set objBaseTable = gcoTablePrivileges.FindTableID(glngHierarchyTableID)
          sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_HasPostSubordinateUser()) THEN 1 ELSE 0 END"
          Set objBaseTable = Nothing
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
          Set objBaseTable = gcoTablePrivileges.FindTableID(glngPersonnelTableID)
          sCode = "CASE WHEN " & objBaseTable.RealSource & ".id IN (SELECT id FROM dbo.udf_ASRFn_HasPersonnelSubordinateUser()) THEN 1 ELSE 0 END"
          Set objBaseTable = Nothing
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
        
        sCode = "(dbo.udf_ASRFn_BradfordFactor(" & sParamCode1 & "," & sParamCode2 _
            & "," & sParamCode3 & "," & strTempTableName & "." & strTempTableID & "))"
      
      Case 77 'Replace characters in a string
        sCode = "REPLACE(" & sParamCode1 & ", " & sParamCode2 & "," & sParamCode3 & ")"
      
      Case Else
        fOK = False
        
    End Select
  End If
  
  If fOK Then
    ' We need to convert date values to varchars in the format 'mm/dd/yyyy'.
    If miReturnType = giEXPRVALUE_DATE Then
      sCode = "convert(" & vbNewLine & _
        "datetime, " & vbNewLine & _
        "convert(" & vbNewLine & _
        "varchar(20), " & vbNewLine & _
        sCode & "," & vbNewLine & _
        "101)" & vbNewLine & _
        ")"
    End If
  End If
  
TidyUpAndExit:
  If fOK Then
    ' JDM - 26/11/01 - Fault 3204 - Needs some brackets around these functions
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






Public Function WriteComponent() As Boolean
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objParameter As clsExprComponent
  Dim objSubExpression As clsExprExpression
  
  fOK = True
  
  sSQL = "INSERT INTO ASRSysExprComponents" & _
    " (componentID, exprID, type, functionID, valueLogic, ExpandedNode)" & _
    " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & _
    " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & _
    " " & Trim(Str(giCOMPONENT_FUNCTION)) & "," & _
    " " & Trim(Str(mlngFunctionID)) & "," & _
    " 0," & IIf(mbExpanded, "1", "0") & ")"
  gADOCon.Execute sSQL, , adCmdText
  
  ' Write the function parameter expressions.
  For Each objParameter In mcolParameters
    Set objSubExpression = objParameter.Component
    objSubExpression.ParentComponentID = mobjBaseComponent.ComponentID
    objSubExpression.ExpressionID = 0
    fOK = objSubExpression.WriteExpression
        
    Set objSubExpression = Nothing
    
    If Not fOK Then
      Exit For
    End If
  Next objParameter
  Set objParameter = Nothing
  
TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function



Public Property Get BadComponent() As clsExprComponent
  ' Return the component last caused the function to fail its validity check.
  Set BadComponent = mobjBadComponent

End Property

Public Property Get ReturnType() As Integer
  ' Return the function's return type.
  ReturnType = miReturnType

End Property

Public Function CopyComponent()
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
    Dim iCount As Integer
    Dim objParameter As clsExprComponent
    For iCount = 1 To mcolParameters.Count
'        Set objParameter = New clsExprComponent
        objFunctionCopy.Parameters.Add (mcolParameters(iCount).CopyComponent)

'        objFunctionCopy.Parameters.Add (mcolParameters.Item(iCount))
    Next iCount
    Set objParameter = Nothing


  Set CopyComponent = objFunctionCopy
   
  ' Disassociate object variables.
  Set objFunctionCopy = Nothing
   
End Function

Public Property Get ComponentType() As Integer
  ' Return the 'function' component type.
  ComponentType = giCOMPONENT_FUNCTION
  
End Property

Public Property Get BaseComponent() As clsExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Function ValidateFunction() As Integer
  ' Validate the function. Return a code describing the validity.
  On Error GoTo BasicErrorTrap
  
  Dim iLoop As Integer
  Dim iValidationCode As ExprValidationCodes
  Dim iFunctionReturnType As ExpressionValueTypes
  Dim aiDummyValues(6) As Integer
  Dim objSubExpression As clsExprExpression
  Dim objParameter As clsExprComponent
  Dim objFunctionDefParameter As clsFunctionDefParameter
  Dim objFunctionDef As clsFunctionDef

  iLoop = 0
  gobjFunctionDefs.Initialise
  
  ' Initialise the validation code.
  iValidationCode = giEXPRVALIDATION_NOERRORS
  Set mobjBadComponent = Nothing

  If gobjFunctionDefs.IsValidID(mlngFunctionID) Then
    Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim(Str(mlngFunctionID)))
  
    ' Validate the function parameter expressions.
    For Each objParameter In mcolParameters
      iLoop = iLoop + 1
    
      Set objSubExpression = objParameter.Component
      With objSubExpression
        ' Validate the parameter expression.
        ' NB. Reset the sub-expression's return type to that defined by the parameter definition
        ' as it may be changeable. The evaluated return type will be determined when the
        ' sub-expression is validated.
        For Each objFunctionDefParameter In objFunctionDef.Parameters
          If objFunctionDefParameter.Index = iLoop Then
            objSubExpression.ReturnType = objFunctionDefParameter.ParameterType
            Exit For
          End If
        Next objFunctionDefParameter
        Set objFunctionDefParameter = Nothing
      
        iValidationCode = .ValidateExpression(False)
      
        If iValidationCode <> giEXPRVALIDATION_NOERRORS Then
          ' Interpret the parameter sub-expression validation code to reflect
          ' the fact that a function parameter was invalid.
          Select Case iValidationCode
            Case giEXPRVALIDATION_NOCOMPONENTS
              iValidationCode = giEXPRVALIDATION_PARAMETERNOCOMPONENTS
            Case giEXPRVALIDATION_SYNTAXERROR
              iValidationCode = giEXPRVALIDATION_PARAMETERSYNTAXERROR
            Case giEXPRVALIDATION_EXPRTYPEMISMATCH
              iValidationCode = giEXPRVALIDATION_PARAMETERTYPEMISMATCH
          End Select
          
          If .BadComponent Is Nothing Then
            Set mobjBadComponent = objParameter
          Else
            Set mobjBadComponent = .BadComponent
          End If
          Exit For
        End If
    
        ' Write the given return type into the array.
        aiDummyValues(iLoop) = .ReturnType
      End With
    Next objParameter
    Set objParameter = Nothing
  End If
  
  If iValidationCode = giEXPRVALIDATION_NOERRORS Then
    If Not ValidateFunctionParameters(mlngFunctionID, iFunctionReturnType, _
      aiDummyValues(1), aiDummyValues(2), aiDummyValues(3), _
      aiDummyValues(4), aiDummyValues(5), aiDummyValues(6)) Then
      
      iValidationCode = giEXPRVALIDATION_PARAMETERTYPEMISMATCH
    Else
      miReturnType = iFunctionReturnType
    End If
  End If

TidyUpAndExit:
  ' Disassociate object variables.
  Set objSubExpression = Nothing
  ValidateFunction = iValidationCode
  Exit Function
  
BasicErrorTrap:
  iValidationCode = giEXPRVALIDATION_UNKNOWNERROR
  Resume TidyUpAndExit

End Function


Public Property Set BaseComponent(ByVal pobjNewValue As clsExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property


Public Property Get FunctionID() As Long
  ' Return the function ID property.
  FunctionID = mlngFunctionID
  
End Property

Public Property Let FunctionID(ByVal plngNewValue As Long)
  ' Set the function ID property.
  mlngFunctionID = plngNewValue
  
  ' Read the function definition from the database.
  ReadFunction
  
End Property


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
  Dim objFunctionDef As clsFunctionDef
  
  fOK = True
  msFunctionName = "<unknown>"
  miReturnType = giEXPRVALUE_UNDEFINED
  msSPName = vbNullString
  
  ' Clear the parameter collection.
  ClearParameters

  ' Get the function definition.
  gobjFunctionDefs.Initialise
  If gobjFunctionDefs.IsValidID(mlngFunctionID) Then
    Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim(Str(mlngFunctionID)))
  
    msFunctionName = objFunctionDef.Name
    miReturnType = objFunctionDef.ReturnType
    msSPName = objFunctionDef.SPName
  Else
    fOK = False
  End If
   
TidyUpAndExit:
  ReadFunctionDetails = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function EvaluateComponent(palngColumnValues As Variant, pvResult As Variant) As Boolean
  ' Return the evaluated result of the function.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fSrchColumnOK As Boolean
  Dim fRtnColumnOK As Boolean
  Dim iLoop As Integer
  Dim sSQL As String
  Dim vParam1 As Variant
  Dim vParam2 As Variant
  Dim vParam3 As Variant
  Dim sFormatString As String
  Dim iNoYears As Integer
  Dim iNoMonths As Integer
  Dim iCount As Integer
  Dim dDateNextMonth As Date
  Dim dDateThisMonth As Date
  Dim lngMonths As Long
  Dim lngYears As Long
  Dim rsInfo As Recordset
  Dim lngSrchTableID As Long
  Dim lngRtnTableID As Long
  Dim sRtnColumnCode As String
  Dim sSrchColumnCode As String
  Dim sSrchColumnName As String
  Dim sRtnColumnName As String
  Dim iRtnColumnDataType As Integer
  Dim sSrchTableName As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim objTableView As CTablePrivilege
  Dim asViews() As String
  Dim sRealTableSource As String
  Dim dblRemainder As Double
  Dim fFound As Boolean
  Dim sSrchTableCode As String

  'Currency Conversion Values
  Dim sCConvTable As String
  Dim sCConvExRateCol As String
  Dim sCConvCurrDescCol As String
  Dim sCConvDecCol As String
  
  sSQL = vbNullString

  fOK = True
  
  ' Get the first parameter's runtime code if required.
  If mcolParameters.Count >= 1 Then
    fOK = mcolParameters.Item(1).Component.EvaluateExpression(palngColumnValues, vParam1)
  End If
  ' Get the second parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 2) Then
    fOK = mcolParameters.Item(2).Component.EvaluateExpression(palngColumnValues, vParam2)
  End If
  ' Get the third parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 3) Then
    fOK = mcolParameters.Item(3).Component.EvaluateExpression(palngColumnValues, vParam3)
  End If

  If fOK Then
    Select Case mlngFunctionID
      Case 1  ' System date
        pvResult = Date
          
      Case 2  ' Convert to uppercase
        pvResult = UCase(vParam1)
        
      Case 3  ' Convert numeric to string
        sFormatString = "#0"
        If vParam2 > 0 Then
          sFormatString = sFormatString & UI.GetSystemDecimalSeparator
  
          For iLoop = 1 To vParam2
            sFormatString = sFormatString & "0"
          Next iLoop
        End If
        pvResult = Format(Str(vParam1), sFormatString)
  
      Case 4  ' If... Then... Else...
        If vParam1 Then
          pvResult = vParam2
        Else
          pvResult = vParam3
        End If
    
      Case 5  ' Remove leading and trailing spaces
        pvResult = Trim(vParam1)
      
      Case 6  ' Extract characters from the left
        pvResult = Left(vParam1, vParam2)
        
      Case 7  ' Length of character field
        pvResult = Len(vParam1)
        
      Case 8  ' Convert to lowercase
        pvResult = LCase(vParam1)
  
      Case 9  ' Maximum
        pvResult = IIf(vParam1 > vParam2, vParam1, vParam2)
      
      Case 10  ' Minimum
        pvResult = IIf(vParam1 < vParam2, vParam1, vParam2)
      
      Case 11  ' Search for character string.
        pvResult = InStr(1, vParam1, vParam2)
        
      Case 13  ' Extract characters from the right
        pvResult = Right(vParam1, vParam2)
        
      Case 14  ' Extract part of a character string
        pvResult = Mid(vParam1, vParam2, vParam3)
  
      Case 15  ' System Time
        pvResult = CStr(Time)
    
      Case 16  ' Is field empty
        ' Validate the sub-expression. This is done, not to validate the expression,
        ' but rather to determine the return type of the expression.
        mcolParameters(1).Component.ValidateExpression False
  
        Select Case mcolParameters(1).ReturnType
          Case giEXPRVALUE_DATE
            pvResult = (vParam1 = vbNullString)
          Case giEXPRVALUE_NUMERIC
            pvResult = (vParam1 = 0)
          Case giEXPRVALUE_LOGIC
            pvResult = (vParam1 = 0)
          Case Else
            pvResult = ((vParam1 = "") Or (vParam1 = vbNullString))
        End Select
  
      Case 17  ' Current user
        ' JPD20021024 Fault 4652
        pvResult = RTrim(UCase(gsUserName))
        'pvResult = gsUserName
      
      Case 18  ' Whole Years Until Current Date
        ' Get the number of whole years
        vParam2 = Date
        
        If (vParam1 = vbNullString) Or _
          (vParam1 > vParam2) Then
          iNoYears = 0
        Else
          iNoYears = Year(vParam2) - Year(vParam1)
        
          ' See if the date passed in months are greater than todays month
          If Month(vParam1) > Month(vParam2) Then
            iNoYears = iNoYears - 1
          End If
        
          ' See if the months are equal and if they are test the day value
          If Month(vParam1) = Month(vParam2) Then
            If Day(vParam1) > Day(vParam2) Then
              iNoYears = iNoYears - 1
            End If
          End If
        End If
        pvResult = iNoYears
          
      Case 19  ' Remaining Months Since Whole Years
        vParam2 = Date
        
        ' Check the date isn't empty and is not greater than the system date
        If (vParam1 = vbNullString) Or _
          (vParam1 > vParam2) Then
          iNoMonths = 0
        Else
          iNoMonths = Month(vParam2) - Month(vParam1)
        
          ' Test the day value
          If Day(vParam1) > Day(vParam2) Then
            iNoMonths = iNoMonths - 1
          End If
          If iNoMonths < 0 Then
            iNoMonths = iNoMonths + 12
          End If
        End If
        
        pvResult = iNoMonths
  
      Case 21  ' First Name from Forenames
        pvResult = Left(vParam1 + " ", InStr(1, vParam1, " "))
      
      Case 22  ' Weekdays From Start and End Dates
        ' Loop through the days adding one for each weekday between the two days
        pvResult = 0
        
        If (vParam1 <> vbNullString) And _
          (vParam2 <> vbNullString) Then
        
          For iCount = 0 To (vParam2 - vParam1)
            If Not Weekday(vParam1 + iCount) = vbSunday And _
               Not Weekday(vParam1 + iCount) = vbSaturday Then
              pvResult = pvResult + 1
            End If
          Next iCount
        End If
      
      Case 23  ' Add months to date
        If (vParam1 = vbNullString) Then
          pvResult = vbNullString
        Else
          pvResult = DateAdd("m", vParam2, vParam1)
        End If
        
      Case 24  ' Add years to date
        If (vParam1 = vbNullString) Then
          pvResult = vbNullString
        Else
          pvResult = DateAdd("yyyy", vParam2, vParam1)
        End If
  
      Case 25  ' Convert character to numeric.
        pvResult = Val(vParam1)
      
      Case 26  ' Whole Months between 2 Dates.
        pvResult = 0
        
        If (vParam1 <> vbNullString) And _
          (vParam2 <> vbNullString) Then
        
          ' Check that date 1 is smaller than date 2
          If vParam1 < vParam2 Then
            ' Get the total number of months
            pvResult = DateDiff("m", vParam1, vParam2)
            
            ' See if the day field of vParam2 < vParam1 day field and if so - 1
            If Day(vParam2) < Day(vParam1) Then
              pvResult = pvResult - 1
            End If
          End If
        End If
    
      Case 27  ' Parentheses
        pvResult = vParam1
  
      Case 28  ' Day of the week
        If (vParam1 = vbNullString) Then
          pvResult = 0
        Else
          pvResult = Weekday(vParam1)
        End If
      
      Case 29  ' Working Days per week
        pvResult = 0
        
        For iLoop = 1 To Len(vParam1)
          If Mid(vParam1, iLoop, 1) <> " " Then
            pvResult = pvResult + 1
          End If
        Next iLoop
    
      Case 31  ' Round down to nearest whole number.
        pvResult = Fix(vParam1)
        
      Case 32  ' Year of date.
        If vParam1 = vbNullString Then
          pvResult = 0
        Else
          pvResult = Year(vParam1)
        End If
        
      Case 33  ' Month of date.
        If vParam1 = vbNullString Then
          pvResult = 0
        Else
          pvResult = Month(vParam1)
        End If
        
      Case 34  ' Day of date.
        If vParam1 = vbNullString Then
          pvResult = 0
        Else
          pvResult = Day(vParam1)
        End If
        
      Case 35  ' Nice Date
        If vParam1 = vbNullString Then
          pvResult = ""
        Else
          ' JPD20021024 Fault 4652
          pvResult = RTrim(UCase(Format(vParam1, "d mmmm yyyy")))
          'pvResult = Format(vParam1, "d mmmm yyyy")
        End If
      
      Case 36  ' Nice Time
        If vParam1 = vbNullString Then
          vParam1 = ""
        End If
          
        ' JPD20021024 Fault 4652
        pvResult = RTrim(UCase(Format(vParam1, "h:mm am/pm")))
        'pvResult = Format(vParam1, "h:mm am/pm")
      
      Case 37  ' Round Date to Start of nearest month
        If (vParam1 = vbNullString) Then
          pvResult = vParam1
        Else
          ' Create a date with one month added to the date and move it to the first day of that month
          dDateNextMonth = DateAdd("m", 1, vParam1)
          dDateNextMonth = DateAdd("d", -1 * (Day(dDateNextMonth) - 1), dDateNextMonth)
          
          ' Create a date which is the first of the month passed in
          dDateThisMonth = DateAdd("d", -1 * (Day(vParam1) - 1), vParam1)
          
          ' See which is the greatest gap between the two start month dates and the passed in date
          If (vParam1 - dDateThisMonth + 1) < (dDateNextMonth - vParam1) Then
            pvResult = dDateThisMonth
          Else
            pvResult = dDateNextMonth
          End If
        End If
      
      Case 38  ' Is Between
        If (vParam1 = vbNullString) Or _
          (vParam2 = vbNullString) Or _
          (vParam3 = vbNullString) Then
        
          pvResult = False
        Else
          If (vParam1 >= vParam2) And (vParam1 <= vParam3) Then
            pvResult = True
          Else
            pvResult = False
          End If
        End If
      
      Case 39  ' Service Years
        If (vParam2 = vbNullString) Then
          vParam2 = Date
        End If
      
        If (vParam1 = vbNullString) Or _
          (vParam1 > vParam2) Then
          iNoYears = 0
        Else
          iNoYears = Year(vParam2) - Year(vParam1)
        
          ' See if the date passed in months are greater than todays month
          If Month(vParam1) > Month(vParam2) Then
            iNoYears = iNoYears - 1
          End If
        
          ' See if the months are equal and if they are test the day value
          If Month(vParam1) = Month(vParam2) Then
            If Day(vParam1) > Day(vParam2) Then
              iNoYears = iNoYears - 1
            End If
          End If
        End If
        pvResult = iNoYears
      
      Case 40  ' Service Months
        If (vParam2 = vbNullString) Then
          vParam2 = Date
        End If
        
        lngMonths = 0
        If (vParam1 <> vbNullString) And _
          (vParam1 < vParam2) Then
          lngMonths = DateDiff("m", vParam1, vParam2)
          
          ' See if the day field of vParam2 < vParam1 day field and if so - 1
          If Day(vParam2) < Day(vParam1) Then
            lngMonths = lngMonths - 1
          End If
        End If
      
        lngYears = 0
        If (vParam1 <> vbNullString) And _
          (vParam1 < vParam2) Then
          lngYears = Year(vParam2) - Year(vParam1)
        
          ' See if the date passed in months are greater than todays month
          If Month(vParam1) > Month(vParam2) Then
            lngYears = lngYears - 1
          End If
        
          ' See if the months are equal and if they are test the day value
          If Month(vParam1) = Month(vParam2) Then
            If Day(vParam1) > Day(vParam2) Then
              lngYears = lngYears - 1
            End If
          End If
        End If
        
        pvResult = lngMonths - (12 * lngYears)
      
      Case 42  ' Get field from database record.
        ' Get the column parameter definitions.
        If (vParam1 = vbNullString) Or _
          (vParam3 = vbNullString) Then
          fOK = False
        Else
          sSQL = "SELECT ASRSysColumns.columnID, ASRSysColumns.columnName, ASRSysTables.tableID, ASRSysTables.tableName, ASRSysColumns.dataType" & _
            " FROM ASRSysColumns" & _
            " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
            " WHERE ASRSysColumns.columnID IN (" & vParam1 & ", " & vParam3 & ")"
          Set rsInfo = datGeneral.GetRecords(sSQL)
          With rsInfo
            Do While Not .EOF
              If !ColumnID = vParam1 Then
                sSrchColumnName = !ColumnName
                sSrchTableName = !TableName
                lngSrchTableID = !TableID
              End If
              
              If !ColumnID = vParam3 Then
                sRtnColumnName = !ColumnName
                lngRtnTableID = !TableID
                iRtnColumnDataType = !DataType
              End If
              
              .MoveNext
            Loop
                
            .Close
          End With
          Set rsInfo = Nothing
        
          fOK = ((Len(sSrchColumnName) > 0) And _
            (Len(sRtnColumnName) > 0)) And _
            (lngSrchTableID = lngRtnTableID)
        End If
        
        ' Construct the select statement to get the required field from the given table,
        ' incorporating permissions.
        If fOK Then
          ReDim asViews(2, 0)
          
          ' Check permissions on that return column.
          Set objColumnPrivileges = GetColumnPrivileges(sSrchTableName)
          sRealTableSource = gcoTablePrivileges.Item(sSrchTableName).RealSource
        
          fRtnColumnOK = objColumnPrivileges.IsValid(sRtnColumnName)
        
          If fRtnColumnOK Then
            fRtnColumnOK = objColumnPrivileges.Item(sRtnColumnName).AllowSelect
          End If
        
          If fRtnColumnOK Then
            ' The search column can be read direct from the table.
            sRtnColumnCode = sRealTableSource & "." & sRtnColumnName
          Else
            ' Then column cannot be read direct. If its from a parent, try parent views
            ' Loop thru the views on the table, seeing if any have read permis for the column
            ' Column 1 = view name
            ' Column 2 = "S" if the view is used for the search.
            For Each objTableView In gcoTablePrivileges.Collection
              If (Not objTableView.IsTable) And _
                (objTableView.TableID = lngSrchTableID) And _
                (objTableView.AllowSelect) Then
        
                ' Get the column permission for the view
                Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
        
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
            Set objTableView = Nothing
        
            ' Does the user have select permission thru ANY views ?
            If UBound(asViews, 2) = 0 Then
              ' The search column can be read neither from the table nor any views.
              sRtnColumnCode = "NULL"
            Else
              For iLoop = 1 To UBound(asViews, 2)
                If iLoop = 1 Then
                  sRtnColumnCode = "CASE"
                End If
        
                sRtnColumnCode = sRtnColumnCode & _
                  " WHEN NOT " & asViews(1, iLoop) & "." & sRtnColumnName & " IS NULL THEN " & asViews(1, iLoop) & "." & sRtnColumnName
              Next iLoop
        
              sRtnColumnCode = sRtnColumnCode & _
                " ELSE NULL" & _
                " END"
            End If
          End If
                
          Set objColumnPrivileges = GetColumnPrivileges(sSrchTableName)
          
          fSrchColumnOK = objColumnPrivileges.IsValid(sSrchColumnName)
        
          If fSrchColumnOK Then
            fSrchColumnOK = objColumnPrivileges.Item(sSrchColumnName).AllowSelect
          End If
        
          If fSrchColumnOK Then
            ' The search column can be read direct from the table.
            sSrchColumnCode = sRealTableSource & "." & sSrchColumnName & " = " & vParam2
          Else
            ' Then column cannot be read direct. If its from a parent, try parent views
            ' Loop thru the views on the table, seeing if any have read permis for the column
            ' Column 1 = view name
            ' Column 2 = "S" if the view is used for the search.
            For Each objTableView In gcoTablePrivileges.Collection
              If (Not objTableView.IsTable) And _
                (objTableView.TableID = lngSrchTableID) And _
                (objTableView.AllowSelect) Then
        
                ' Get the column permission for the view
                Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
        
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
            Set objTableView = Nothing
        
            sSrchColumnCode = ""
            For iLoop = 1 To UBound(asViews, 2)
              If asViews(2, iLoop) = "S" Then
                fSrchColumnOK = True
              
                If Len(sSrchColumnCode) > 0 Then
                  sSrchColumnCode = sSrchColumnCode & " OR "
                End If
          
                sSrchColumnCode = sSrchColumnCode & _
                  sRealTableSource & ".id IN (SELECT id FROM " & asViews(1, iLoop) & " WHERE " & sSrchColumnName & " = " & vParam2 & ")"
              End If
            Next iLoop
          End If
          
          sSrchTableCode = ""
          For iLoop = 1 To UBound(asViews, 2)
            sSrchTableCode = sSrchTableCode & _
              " LEFT OUTER JOIN " & asViews(1, iLoop) & " ON " & sRealTableSource & ".id = " & asViews(1, iLoop) & ".id"
          Next iLoop
          
          If fSrchColumnOK Then
            sSQL = "SELECT TOP 1 " & sRtnColumnCode & " AS [result]" & _
              " FROM " & sRealTableSource & _
              " " & sSrchTableCode & _
              " WHERE (" & sSrchColumnCode & ")"
              
            Set rsInfo = datGeneral.GetRecords(sSQL)
            With rsInfo
              If (.EOF And .BOF) Then
                fOK = False
              Else
                ' JPD20021024 Fault 4652
                If (iRtnColumnDataType = sqlOle) Or _
                  (iRtnColumnDataType = sqlVarChar) Or _
                  (iRtnColumnDataType = sqlVarBinary) Or _
                  (iRtnColumnDataType = sqlLongVarChar) Then
                  ' Convert OLE, character, photo and working pattern columns
                  ' (all essentially string type columns) to uppercase.
                  pvResult = RTrim(UCase(!Result))
                Else
                  pvResult = !Result
                End If
              End If
              .Close
            End With
            Set rsInfo = Nothing
          Else
            fOK = False
          End If
        End If
  
      Case 44  ' Add days to date.
        If (vParam1 = vbNullString) Then
          pvResult = vbNullString
        Else
          pvResult = DateAdd("d", vParam2, vParam1)
        End If
      
      Case 45  ' Days Between 2 Dates
        If (vParam1 = vbNullString) Or _
          (vParam2 = vbNullString) Then
          fOK = False
        Else
          pvResult = vParam2 - vParam1 + 1
        End If
  
      Case 48  ' Round Up to nearest whole number.
        ' JPD20030116 Fault 4910
        If vParam1 = Fix(vParam1) Then
          pvResult = Fix(vParam1)
        Else
          pvResult = Fix(vParam1) + IIf(vParam1 < 0, -1, 1)
        End If
  
      Case 49  ' Round to nearest number.
        ' JPD20030116 Fault 4910
        dblRemainder = vParam1 - (Int(vParam1 / vParam2) * vParam2)
        If ((vParam1 < 0) And (dblRemainder <= (vParam2 / 2))) Or _
          ((vParam1 >= 0) And (dblRemainder < (vParam2 / 2))) Then
          pvResult = vParam1 - dblRemainder
        Else
          pvResult = vParam1 + vParam2 - dblRemainder
        End If
        
      'TM20011022 Currency Implementation
      Case 51
      '*********** runtime code to go here *************

        ' Get the column parameter definitions.
        sSQL = vbNullString
        sSQL = sSQL & "SELECT ASRSysModuleSetup.*, ASRSysColumns.ColumnName, ASRSysTables.TableName "
        sSQL = sSQL & "FROM ASRSysModuleSetup "
        sSQL = sSQL & "     INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID "
        sSQL = sSQL & "                INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID "
        sSQL = sSQL & "WHERE ASRSysModuleSetup.ModuleKey = 'MODULE_CURRENCY' "
        
        Set rsInfo = datGeneral.GetRecords(sSQL)
        sSQL = vbNullString
        With rsInfo
          If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do While Not .EOF
              sCConvTable = !TableName
              Select Case !ParameterKey
              Case "Param_CurrencyNameColumn": sCConvCurrDescCol = !ColumnName
              Case "Param_ConversionValueColumn": sCConvExRateCol = !ColumnName
              Case "Param_DecimalColumn": sCConvDecCol = !ColumnName
              End Select
              .MoveNext
            Loop
            .Close
          Else
            pvResult = Null
          End If
        End With
        Set rsInfo = Nothing
        
        If (Len(sCConvTable) > 0) And (Len(sCConvCurrDescCol) > 0) And _
              (Len(sCConvExRateCol) > 0) And (Len(sCConvDecCol) > 0) Then
          sSQL = vbNullString
          sSQL = sSQL & "(SELECT ROUND((" & vParam1
          sSQL = sSQL & "              / "
          sSQL = sSQL & "             (SELECT " & sCConvTable & "." & sCConvExRateCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = '" & vParam2 & "') "
          sSQL = sSQL & "              * "
          sSQL = sSQL & "             (SELECT " & sCConvTable & "." & sCConvExRateCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = '" & vParam3 & "')) "
          sSQL = sSQL & "        , "
          sSQL = sSQL & "        (SELECT " & sCConvTable & "." & sCConvDecCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = '" & vParam3 & "')) AS [result] ) "
        
          Set rsInfo = datGeneral.GetRecords(sSQL)
          With rsInfo
            If (.EOF And .BOF) Then
              pvResult = Null
            Else
              pvResult = !Result
            End If
            .Close
          End With
          Set rsInfo = Nothing
        Else
          pvResult = Null
        End If
      
      
      Case 54 ' Whole months between two dates
        pvResult = 0
        
        If (vParam1 <> vbNullString) And _
          (vParam2 <> vbNullString) Then
        
          ' Check that date 1 is smaller than date 2
          If vParam1 < vParam2 Then
            ' Get the total number of years
            pvResult = DateDiff("y", vParam1, vParam2)
            
            If Month(vParam1) > Month(vParam2) Then
              pvResult = pvResult - 1
            End If
            
            If Month(vParam1) = Month(vParam2) Then
               ' See if the day field of vParam2 < vParam1 day field and if so - 1
              If Day(vParam1) > Day(vParam2) Then
                pvResult = pvResult - 1
              End If
            End If
            
          End If
        End If

      ' JPD20021121 Fault 3177
      Case 55  ' First Day of Month.
        If (vParam1 = vbNullString) Then
          pvResult = vbNullString
        Else
          pvResult = DateAdd("d", 1 - Day(vParam1), vParam1)
        End If
      
      ' JPD20021121 Fault 3177
      Case 56  ' Last Day of Month.
        If (vParam1 = vbNullString) Then
          pvResult = vbNullString
        Else
          pvResult = DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", 1 - Day(vParam1), vParam1)))
        End If
      
      ' JPD20021121 Fault 3177
      Case 57  ' First Day of Year.
        If (vParam1 = vbNullString) Then
          pvResult = vbNullString
        Else
          pvResult = DateAdd("m", 1 - Month(vParam1), DateAdd("d", 1 - Day(vParam1), vParam1))
        End If
      
      ' JPD20021121 Fault 3177
      Case 58  ' Last Day of Year.
        If (vParam1 = vbNullString) Then
          pvResult = vbNullString
        Else
          pvResult = DateAdd("d", -1, DateAdd("yyyy", 1, DateAdd("m", 1 - Month(vParam1), DateAdd("d", 1 - Day(vParam1), vParam1))))
        End If
      
      ' JPD20021129 Fault 4337
      Case 59  ' Name of Month.
        If vParam1 = vbNullString Then
          pvResult = vbNullString
        Else
          pvResult = MonthName(Month(vParam1))
        End If
      
      ' JPD20021129 Fault 4337
      Case 60  ' Name of Day.
        If vParam1 = vbNullString Then
          pvResult = vbNullString
        Else
          pvResult = WeekdayName(Weekday(vParam1), False, vbSunday)
        End If
      
      ' JPD20021129 Fault 3606
      Case 61  ' Is field populated
        ' Validate the sub-expression. This is done, not to validate the expression,
        ' but rather to determine the return type of the expression.
        mcolParameters(1).Component.ValidateExpression False
  
        Select Case mcolParameters(1).ReturnType
          Case giEXPRVALUE_DATE
            pvResult = (vParam1 <> vbNullString)
          Case giEXPRVALUE_NUMERIC
            pvResult = (vParam1 <> 0)
          Case giEXPRVALUE_LOGIC
            pvResult = (vParam1 <> 0)
          Case Else
            pvResult = ((vParam1 <> "") And (vParam1 <> vbNullString))
        End Select
      
      'Case 62 'PARENTAL LEAVE ENTITLEMENT
      'Case 63 'PARENTAL LEAVE TAKEN
      'Case 64 'MATERNITY EXPECTED RETURN DATE
  
      
      Case 77 ' Replace Characters in a String
         If vParam1 = vbNullString Then
          pvResult = vbNullString
        Else
          pvResult = Replace(vParam1, vParam2, vParam3)
        End If
              
      Case Else
        fOK = False
    End Select
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  EvaluateComponent = fOK
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
  Dim iIndex As Integer
  Dim sSQL As String
  Dim rsParameters As Recordset
  Dim objNewParameter As clsExprComponent
  Dim objFunctionDefParameter As clsFunctionDefParameter
  Dim objFunctionDef As clsFunctionDef
  Dim iMinIndex As Integer
  Dim iMaxIndex As Integer
  
  fOK = True
  
  ' Clear the parameter collection.
  ClearParameters

  ' Read the function parameter list from the collection.
  gobjFunctionDefs.Initialise
  
  If gobjFunctionDefs.IsValidID(mlngFunctionID) Then
    Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim(Str(mlngFunctionID)))
  
    iMinIndex = 0
    iMaxIndex = 0
    For Each objFunctionDefParameter In objFunctionDef.Parameters
      If objFunctionDefParameter.Index < iMinIndex Then iMinIndex = objFunctionDefParameter.Index
      If objFunctionDefParameter.Index > iMaxIndex Then iMaxIndex = objFunctionDefParameter.Index
    Next objFunctionDefParameter
    Set objFunctionDefParameter = Nothing
    If iMinIndex > iMaxIndex Then iMaxIndex = iMinIndex
    
    For iIndex = iMinIndex To iMaxIndex
      For Each objFunctionDefParameter In objFunctionDef.Parameters
        If objFunctionDefParameter.Index = iIndex Then
          mcolParameters.Add New clsExprComponent
          With mcolParameters.Item(mcolParameters.Count)
            .ComponentType = giCOMPONENT_EXPRESSION
            .Component.Name = objFunctionDefParameter.Name
            .Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType
            .Component.ReturnType = objFunctionDefParameter.ParameterType
            .Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID
            .Component.SecondTableID = mobjBaseComponent.ParentExpression.SecondTableID   'MH20021108
            .Component.ColumnList = mobjBaseComponent.ParentExpression.ColumnList
            Set .ParentExpression = mobjBaseComponent.ParentExpression
          End With
        End If
      Next objFunctionDefParameter
      Set objFunctionDefParameter = Nothing
    Next iIndex
  End If
    
  ' Get the customised function parameter definitions if they exist.
  If (mobjBaseComponent.ComponentID > 0) Then
  
    iIndex = 1
    sSQL = "SELECT *" & _
      " FROM ASRSysExpressions" & _
      " WHERE parentComponentID = " & Trim(Str(mobjBaseComponent.ComponentID)) & _
      " ORDER BY exprID"
    Set rsParameters = datGeneral.GetRecords(sSQL)
    With rsParameters
      Do While (Not .EOF) And fOK
        
        ' Instantiate a new component object.
        Set objNewParameter = New clsExprComponent
  
        ' Construct the hierarchy of objects that define the parameter.
        objNewParameter.ComponentType = giCOMPONENT_EXPRESSION
        objNewParameter.Component.ExpressionID = !ExprID
        objNewParameter.Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID
        objNewParameter.Component.SecondTableID = mobjBaseComponent.ParentExpression.SecondTableID  'MH20021108
        objNewParameter.Component.ColumnList = mobjBaseComponent.ParentExpression.ColumnList
        objNewParameter.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType
        Set objNewParameter.ParentExpression = mobjBaseComponent.ParentExpression
  
        fOK = objNewParameter.Component.ConstructExpression
        
        If fOK Then
          objNewParameter.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType
  
          ' Reset the new expression's return type.
          objNewParameter.Component.ReturnType = mcolParameters.Item(iIndex).Component.ReturnType

          'JDM - 16/03/01 - Fault 1935 - Load previous view
          objNewParameter.ExpandedNode = IIf(IsNull(!ExpandedNode), False, !ExpandedNode)
  
          ' Insert the new expression into the function's parameter array.
          mcolParameters.Remove iIndex
          If mcolParameters.Count >= iIndex Then
            mcolParameters.Add objNewParameter, , iIndex
          Else
            mcolParameters.Add objNewParameter
          End If
        End If
  
        iIndex = iIndex + 1
        Set objNewParameter = Nothing
          
        .MoveNext
      Loop
    End With
  End If
  
TidyUpAndExit:
  Set rsParameters = Nothing
  ReadParameterDefinition = fOK
  Exit Function
  
ErrorTrap:
  ' Clear the parameter collection.
  ClearParameters
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function HiddenElements() As Boolean
  On Error GoTo ErrorTrap
  
  Dim objParameter As clsExprComponent
  Dim objSubExpression As clsExprExpression
  
  HiddenElements = False
  
  For Each objParameter In mcolParameters
    Set objSubExpression = objParameter.Component
    
    HiddenElements = objSubExpression.HiddenElements
        
    Set objSubExpression = Nothing
    
    If HiddenElements Then
      Exit For
    End If
  Next objParameter
  Set objParameter = Nothing
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  COAMsgBox "Error checking for hidden components.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
  
End Function


Public Property Get ComponentDescription() As String
  ' Return the Function description.
  ComponentDescription = msFunctionName
  
End Property

Public Property Get Parameters() As Collection
  ' Return the parameter collection.
  Set Parameters = mcolParameters
  
End Property

Public Property Set Parameters(ByVal pcolCollection As Collection)
  ' Set the parameter collection.
  Set mcolParameters = pcolCollection

End Property

Public Sub ClearParameters()
  ' Clear the function's parameter collection.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Remove all components from the collection.
  Do While mcolParameters.Count > 0
    mcolParameters.Remove 1
  Loop
  Set mcolParameters = Nothing
  
  ' Re-instantiate the collection.
  Set mcolParameters = New Collection
  Exit Sub

ErrorTrap:
  fOK = False
  
End Sub



Private Sub Class_Initialize()
  ' Create a new collection to hold the function's parameters.
  Set mcolParameters = New Collection

End Sub


Private Sub Class_Terminate()
  ' Disassociate object variables.
  Set mcolParameters = Nothing

End Sub



















































Public Property Get ExpandedNode() As Boolean
    'Return whether this node is expanded or not
    ExpandedNode = mbExpanded

End Property

Public Property Let ExpandedNode(ByVal bNewValue As Boolean)
    'Set whether this component node is expanded or not
    mbExpanded = bNewValue

End Property

Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Copy the component definition to the clipboard object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objParameter As clsExprComponent
  Dim sClipboardMessage As String

  fOK = True
  sClipboardMessage = String(piLevel, vbTab) & "Function : " & ComponentDescription & vbNewLine
  
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
  ' Copy the function's parameter expressions to the clipboard
  For Each objParameter In mcolParameters
    fOK = objParameter.Component.CopyToClipboard(piLevel + 1)
  Next objParameter

TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function UDFCode(psRuntimeCode() As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
  Dim sParamCode1 As String
  Dim sParamCode2 As String
  Dim sParamCode3 As String
  Dim sParamCode4 As String
  Dim fOK As Boolean
  
  fOK = True
  
  ' Get the first parameter's runtime code if required.
  If mcolParameters.Count >= 1 Then
    fOK = mcolParameters.Item(1).Component.UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If
  
  ' Get the second parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 2) Then
    fOK = mcolParameters.Item(2).Component.UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If
  
  ' Get the third parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 3) Then
    fOK = mcolParameters.Item(3).Component.UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If
    
  ' Get the fourth parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 4) Then
    fOK = mcolParameters.Item(4).Component.UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If
 
  UDFCode = fOK
  
End Function
