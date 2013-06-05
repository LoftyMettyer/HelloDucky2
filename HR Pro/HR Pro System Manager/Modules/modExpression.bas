Attribute VB_Name = "modExpression"
Option Explicit

Public Function ExprTypeName(intExpressionType As ExpressionTypes) As String
  ' Return the description of the expression type.
  Select Case intExpressionType
    Case giEXPR_COLUMNCALCULATION
      ExprTypeName = "Column Calculation"
    
    Case giEXPR_GOTFOCUS ' NOT USED.
      ExprTypeName = "Field Entry Validation Clause"
    
    Case giEXPR_RECORDVALIDATION
      ExprTypeName = "Field Validation"
    
    Case giEXPR_DEFAULTVALUE
      ExprTypeName = "Default Value"
  
    Case giEXPR_STATICFILTER
      ExprTypeName = "Filter"
  
    Case giEXPR_PAGEBREAK ' NOT USED.
      ExprTypeName = "Page Break"
  
    Case giEXPR_ORDER ' NOT USED.
      ExprTypeName = "Order"
  
    Case giEXPR_RECORDDESCRIPTION
      ExprTypeName = "Record Description"
  
    Case giEXPR_OUTLOOKFOLDER
      ExprTypeName = "Calendar Folder"
  
    Case giEXPR_OUTLOOKSUBJECT
      ExprTypeName = "Subject Calculation"
  
    Case giEXPR_VIEWFILTER
      ExprTypeName = "View Filter"
  
    Case giEXPR_RUNTIMECALCULATION
      ExprTypeName = "Runtime Calculation"
    
    Case giEXPR_RECORDINDEPENDANTCALC
      ExprTypeName = "Record Independent Calculation"
      
    Case giEXPR_RUNTIMEFILTER
      ExprTypeName = "Filter"
  
    Case giEXPR_EMAIL
      'MH20031210 Fault
      '''MH20000727 Add Email
      'ExprTypeName = "Email Address"
      ExprTypeName = "Calculated Email Address"
    
    Case giEXPR_LINKFILTER
      ExprTypeName = "Link Filter"

    Case giEXPR_WORKFLOWCALCULATION
      ExprTypeName = "Workflow Calculation"

    Case giEXPR_WORKFLOWSTATICFILTER
      ExprTypeName = "Workflow Filter"

    Case giEXPR_WORKFLOWRUNTIMEFILTER
      ExprTypeName = "Workflow Filter"

    Case Else
      ExprTypeName = "Expression"
  
  End Select

End Function


Public Function ComponentTypeName(piComponentType As ExpressionComponentTypes) As String
  ' Return the description of the component type.
  Select Case piComponentType

    Case giCOMPONENT_FIELD
      ComponentTypeName = "Field"
    Case giCOMPONENT_FUNCTION
      ComponentTypeName = "Function"
    Case giCOMPONENT_CALCULATION
      ComponentTypeName = "Calculation"
    Case giCOMPONENT_VALUE
      ComponentTypeName = "Value"
    Case giCOMPONENT_OPERATOR
      ComponentTypeName = "Operator"
    Case giCOMPONENT_TABLEVALUE
      ComponentTypeName = "Lookup Table Value"
    Case giCOMPONENT_PROMPTEDVALUE
      ComponentTypeName = "Prompted Value"
    Case giCOMPONENT_CUSTOMCALC
      ComponentTypeName = "Custom Calculation"
    Case giCOMPONENT_EXPRESSION
      ComponentTypeName = "Expression"
    Case giCOMPONENT_FILTER
      ComponentTypeName = "Filter"
    Case giCOMPONENT_WORKFLOWVALUE
      ComponentTypeName = "Workflow Value"
    Case giCOMPONENT_WORKFLOWFIELD
      ComponentTypeName = "Workflow Identified Field"
    Case Else
      ComponentTypeName = "Component"
  
  End Select

End Function



Public Function GetExpressionUsageDesc(plngExprID As Long) As String

  Dim rsExpressions As dao.Recordset
  Dim rsComponents As dao.Recordset
  Dim strSQL As String
  Dim sUtilityName As String

  strSQL = "SELECT ParentComponentID, TableID, Type, Name, utilityID " & _
           "FROM   tmpExpressions " & _
           "WHERE  ExprID = " & CStr(plngExprID) & _
           "  AND  Type <> " & CStr(giEXPR_RUNTIMECALCULATION) & _
           "  AND  Type <> " & CStr(giEXPR_RUNTIMEFILTER)

  Set rsExpressions = daoDb.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)
  
  If Not (rsExpressions.BOF And rsExpressions.EOF) Then
    If rsExpressions!ParentComponentID > 0 Then
      strSQL = "SELECT ExprID " & _
               "FROM   tmpComponents " & _
               "WHERE  ComponentID = " & CStr(rsExpressions!ParentComponentID)
      Set rsComponents = daoDb.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)

      If Not (rsComponents.BOF And rsComponents.EOF) Then
        GetExpressionUsageDesc = GetExpressionUsageDesc(rsComponents!ExprID)
      End If

      rsComponents.Close
      Set rsComponents = Nothing
    Else
      If (rsExpressions!Type = giEXPR_WORKFLOWCALCULATION) _
        Or (rsExpressions!Type = giEXPR_WORKFLOWSTATICFILTER) _
        Or (rsExpressions!Type = giEXPR_WORKFLOWRUNTIMEFILTER) Then
      
        sUtilityName = GetWorkflowName(rsExpressions!UtilityID)

        GetExpressionUsageDesc = _
          ExprTypeName(rsExpressions!Type) & " : " & _
          rsExpressions!Name & _
          IIf(sUtilityName = "<unknown>", " <Unknown workflow>", " <'" & sUtilityName & "' workflow>")
      Else
        GetExpressionUsageDesc = _
          ExprTypeName(rsExpressions!Type) & " : " & _
          rsExpressions!Name & _
          " <" & GetTableName(rsExpressions!TableID) & ">"
      End If
    End If
  End If
  
  rsExpressions.Close
  Set rsExpressions = Nothing

End Function


Public Function GetExpressionUsageDescFromSQL(plngExprID As Long) As String

  Dim rsExpressions As ADODB.Recordset
  Dim rsComponents As ADODB.Recordset
  Dim strSQL As String


  strSQL = "SELECT ParentComponentID, TableID, Type, Name " & _
           "FROM   ASRSysExpressions " & _
           "WHERE  ExprID = " & CStr(plngExprID) & _
           "  AND  (Type = " & CStr(giEXPR_RUNTIMECALCULATION) & _
           "        OR Type = " & CStr(giEXPR_RUNTIMEFILTER) & _
           "        OR Type = " & CStr(giEXPR_EMAIL) & ")"

  Set rsExpressions = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsExpressions.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsExpressions.Open strSQL, gADOCon, adOpenStatic, adLockReadOnly
  
  If Not (rsExpressions.BOF And rsExpressions.EOF) Then
    If rsExpressions!ParentComponentID > 0 Then
      strSQL = "SELECT ExprID " & _
               "FROM   ASRSysExprComponents " & _
               "WHERE  ComponentID = " & CStr(rsExpressions!ParentComponentID)
      Set rsComponents = New ADODB.Recordset
      ' AE20080326 Fault #13043 - KB272358
      'rsComponents.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      rsComponents.Open strSQL, gADOCon, adOpenStatic, adLockReadOnly

      If Not (rsComponents.BOF And rsComponents.EOF) Then
        GetExpressionUsageDescFromSQL = GetExpressionUsageDescFromSQL(rsComponents!ExprID)
      End If

      rsComponents.Close
      Set rsComponents = Nothing

    Else
      GetExpressionUsageDescFromSQL = _
          ExprTypeName(rsExpressions!Type) & " : " & _
          rsExpressions!Name & _
          " <" & GetTableName(rsExpressions!TableID) & ">"
    
    End If

  End If
  rsExpressions.Close
  Set rsExpressions = Nothing

End Function


Public Function HasExpressionComponent(plngExprIDBeingSearched As Long, plngExprIDSearchedFor As Long) As Boolean
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap

  Dim rsExprComp As dao.Recordset
  Dim rsExpr As dao.Recordset
  Dim fHasExpr As Boolean
  Dim sSQL As String
  Dim lngSubExprID As Long
  
  HasExpressionComponent = (plngExprIDBeingSearched = plngExprIDSearchedFor)
  
  If Not HasExpressionComponent Then
    sSQL = "SELECT * FROM tmpComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
    Set rsExprComp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    With rsExprComp
      Do Until .EOF
        Select Case !Type
          Case giCOMPONENT_CALCULATION
            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)
      
            If lngSubExprID > 0 Then
              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
            End If
      
          Case giCOMPONENT_FILTER
            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)
      
            If lngSubExprID > 0 Then
              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
            End If
      
          Case giCOMPONENT_FIELD
            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
      
            If lngSubExprID > 0 Then
              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
            End If
        
          Case giCOMPONENT_FUNCTION
            sSQL = "SELECT exprID FROM tmpExpressions WHERE parentComponentID = " & CStr(!ComponentID)
            Set rsExpr = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            Do Until rsExpr.EOF
              HasExpressionComponent = HasExpressionComponent(rsExpr!ExprID, plngExprIDSearchedFor)
              
              If HasExpressionComponent Then
                Exit Do
              End If
              
              rsExpr.MoveNext
            Loop
            rsExpr.Close
            Set rsExpr = Nothing
        
          Case giCOMPONENT_WORKFLOWFIELD
            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
      
            If lngSubExprID > 0 Then
              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
            End If
        
        End Select
        
        If HasExpressionComponent Then
          Exit Do
        End If
        
        .MoveNext
      Loop
    End With
  
    rsExprComp.Close
  End If
  
TidyUpAndExit:
  Set rsExprComp = Nothing
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
    
End Function

Public Function ExpressionUsesColumn(plngExprIDBeingSearched As Long, plngColumnIDSearchedFor As Long) As Boolean

  'NB. This reads data from the sql db. i.e. not the access db.
  
  On Error GoTo ErrorTrap

  Dim rsExprComp As New ADODB.Recordset
  Dim rsExpr As New ADODB.Recordset
  Dim fHasExpr As Boolean
  Dim sSQL As String
  Dim lngSubExprID As Long
  Dim lngColumnID As Long
  
  ExpressionUsesColumn = False
  
  If Not ExpressionUsesColumn Then
    sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
    rsExprComp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
    With rsExprComp
      Do Until .EOF
        lngColumnID = 0
        lngSubExprID = 0
        
        Select Case !Type
          Case giCOMPONENT_CALCULATION
            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)
      
            If lngSubExprID > 0 Then
              ExpressionUsesColumn = ExpressionUsesColumn(lngSubExprID, plngColumnIDSearchedFor)
            End If
      
          Case giCOMPONENT_FILTER
            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)
      
            If lngSubExprID > 0 Then
              ExpressionUsesColumn = ExpressionUsesColumn(lngSubExprID, plngColumnIDSearchedFor)
            End If
      
          Case giCOMPONENT_FIELD
            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
            lngColumnID = IIf(IsNull(!fieldColumnID), 0, !fieldColumnID)
            
            If lngSubExprID > 0 Then
              ExpressionUsesColumn = ExpressionUsesColumn(lngSubExprID, plngColumnIDSearchedFor)
            ElseIf lngColumnID > 0 Then
              ExpressionUsesColumn = (lngColumnID = plngColumnIDSearchedFor)
            End If
          
          Case giCOMPONENT_TABLEVALUE
            lngColumnID = IIf(IsNull(!LookupColumnID), 0, !LookupColumnID)
            
            If lngColumnID > 0 Then
              ExpressionUsesColumn = (lngColumnID = plngColumnIDSearchedFor)
            End If
            
          Case giCOMPONENT_FUNCTION
            sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(!ComponentID)
            rsExpr.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

            Do Until rsExpr.EOF
              ExpressionUsesColumn = ExpressionUsesColumn(rsExpr!ExprID, plngColumnIDSearchedFor)
              
              If ExpressionUsesColumn Then
                Exit Do
              End If
              
              rsExpr.MoveNext
            Loop
            rsExpr.Close

        End Select
        
        If ExpressionUsesColumn Then
          Exit Do
        End If
        
        .MoveNext
      Loop
    End With
  
    rsExprComp.Close
  End If
  
TidyUpAndExit:
  Set rsExpr = Nothing
  Set rsExprComp = Nothing
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
    
End Function

Public Function ExpressionUsesRelationship_SQL(plngExprIDBeingSearched As Long, _
                                                  plngParentTableIDSearcherFor As Long, _
                                                  plngChildTableIDSearcherFor As Long) As Boolean

  'NB. This reads data from the sql db. i.e. not the access db.
  
  On Error GoTo ErrorTrap

  Dim rsExprComp As New ADODB.Recordset
  Dim rsExpr As New ADODB.Recordset
  Dim fHasExpr As Boolean
  Dim sSQL As String
  Dim lngSubExprID As Long
  Dim lngFieldTableID As Long
  Dim lngExprBaseTableID As Long
  
  ExpressionUsesRelationship_SQL = False
  
  If Not ExpressionUsesRelationship_SQL Then
    sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
    rsExprComp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With rsExprComp
      Do Until .EOF
        lngExprBaseTableID = IIf(IsNull(!TableID), 0, !TableID)
        lngSubExprID = 0
        
        Select Case !Type
          Case giCOMPONENT_CALCULATION
            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)

            If lngSubExprID > 0 Then
              ExpressionUsesRelationship_SQL = ExpressionUsesRelationship_SQL(lngSubExprID, _
                                                                              plngParentTableIDSearcherFor, _
                                                                              plngChildTableIDSearcherFor)
            End If
      
          Case giCOMPONENT_FILTER
            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)

            If lngSubExprID > 0 Then
              ExpressionUsesRelationship_SQL = ExpressionUsesRelationship_SQL(lngSubExprID, _
                                                                              plngParentTableIDSearcherFor, _
                                                                              plngChildTableIDSearcherFor)
            End If
      
          Case giCOMPONENT_FIELD
            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
            lngFieldTableID = IIf(IsNull(!fieldTableID), 0, !fieldTableID)

            If lngSubExprID > 0 Then
              ExpressionUsesRelationship_SQL = ExpressionUsesRelationship_SQL(lngSubExprID, _
                                                                              plngParentTableIDSearcherFor, _
                                                                              plngChildTableIDSearcherFor)
            Else
              ExpressionUsesRelationship_SQL = ((lngExprBaseTableID = plngParentTableIDSearcherFor) _
                                                  And (lngFieldTableID = plngChildTableIDSearcherFor)) _
                                                Or ((lngExprBaseTableID = plngChildTableIDSearcherFor) _
                                                  And (lngFieldTableID = plngParentTableIDSearcherFor))
            End If
          
          Case giCOMPONENT_FUNCTION
            sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(!ComponentID)
            rsExpr.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

            Do Until rsExpr.EOF
              ExpressionUsesRelationship_SQL = ExpressionUsesRelationship_SQL(lngSubExprID, _
                                                                              plngParentTableIDSearcherFor, _
                                                                              plngChildTableIDSearcherFor)

              If ExpressionUsesRelationship_SQL Then
                Exit Do
              End If

              rsExpr.MoveNext
            Loop
            rsExpr.Close

        End Select
        
        If ExpressionUsesRelationship_SQL Then
          Exit Do
        End If
        
        .MoveNext
      Loop
    End With
  
    rsExprComp.Close
  End If
  
TidyUpAndExit:
  Set rsExpr = Nothing
  Set rsExprComp = Nothing
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
    
End Function

Public Function ExpressionUsesRelationship(plngExprIDBeingSearched As Long, plngParentTableIDSearcherFor As Long, _
                                            plngChildTableIDSearcherFor As Long, lngExprBaseTableID As Long) As Boolean

  'NB. This reads data from the sql db. i.e. not the access db.
  
  On Error GoTo ErrorTrap

  Dim rsExprComp As dao.Recordset
  Dim rsExpr As dao.Recordset
  Dim fHasExpr As Boolean
  Dim sSQL As String
  Dim lngSubExprID As Long
  Dim lngFieldTableID As Long
  
  ExpressionUsesRelationship = False
  
  If Not ExpressionUsesRelationship Then
    sSQL = "SELECT * FROM tmpComponents WHERE tmpComponents.exprID = " & CStr(plngExprIDBeingSearched)
    Set rsExprComp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    With rsExprComp
      Do Until .EOF
        lngSubExprID = 0
        
        Select Case !Type
          Case giCOMPONENT_CALCULATION
            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)

            If lngSubExprID > 0 Then
              ExpressionUsesRelationship = ExpressionUsesRelationship(lngSubExprID, plngParentTableIDSearcherFor, _
                                                                      plngChildTableIDSearcherFor, lngExprBaseTableID)
            End If
      
          Case giCOMPONENT_FILTER
            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)

            If lngSubExprID > 0 Then
              ExpressionUsesRelationship = ExpressionUsesRelationship(lngSubExprID, plngParentTableIDSearcherFor, _
                                                                      plngChildTableIDSearcherFor, lngExprBaseTableID)
            End If
      
          Case giCOMPONENT_FIELD
            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
            lngFieldTableID = IIf(IsNull(!fieldTableID), 0, !fieldTableID)

            If lngSubExprID > 0 Then
              ExpressionUsesRelationship = ExpressionUsesRelationship(lngSubExprID, plngParentTableIDSearcherFor, _
                                                                      plngChildTableIDSearcherFor, lngExprBaseTableID)
            Else
              ExpressionUsesRelationship = ((lngExprBaseTableID = plngParentTableIDSearcherFor) _
                                            And (lngFieldTableID = plngChildTableIDSearcherFor)) _
                                          Or ((lngExprBaseTableID = plngChildTableIDSearcherFor) _
                                            And (lngFieldTableID = plngParentTableIDSearcherFor))
            End If
          
          Case giCOMPONENT_FUNCTION
            sSQL = "SELECT tmpExpressions.exprID FROM tmpExpressions WHERE tmpExpressions.parentComponentID = " & CStr(!ComponentID)
            Set rsExpr = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

            Do Until rsExpr.EOF
              ExpressionUsesRelationship = ExpressionUsesRelationship(rsExpr!ExprID, plngParentTableIDSearcherFor, _
                                                                      plngChildTableIDSearcherFor, lngExprBaseTableID)

              If ExpressionUsesRelationship Then
                Exit Do
              End If

              rsExpr.MoveNext
            Loop
            rsExpr.Close

        End Select
        
        If ExpressionUsesRelationship Then
          Exit Do
        End If
        
        .MoveNext
      Loop
    End With
  
    rsExprComp.Close
  End If
  
TidyUpAndExit:
  Set rsExpr = Nothing
  Set rsExprComp = Nothing
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
    
End Function

Public Sub CalculatedColumnsThatUseFunction(ByRef pvColumns As Variant, plngFunctionID As Long)
  On Error GoTo ErrorTrap
  
  Dim rsCheck As dao.Recordset
  Dim objComp As CExprComponent
  Dim sSQL As String
  Dim lngExprID As Long
  Dim objCalc As CExpression
  
  sSQL = "SELECT DISTINCT tmpComponents.componentID" & _
    " FROM tmpComponents, tmpExpressions " & _
    " WHERE tmpExpressions.exprid = tmpComponents.Exprid " & _
    "   AND tmpComponents.functionID = " & Trim$(Str$(plngFunctionID))

  Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  Do Until rsCheck.EOF
    Set objComp = New CExprComponent
    objComp.ComponentID = rsCheck!ComponentID
    lngExprID = objComp.RootExpressionID
    Set objComp = Nothing
      
    Set objCalc = New CExpression
    With objCalc
      ' Construct the Filter expression.
      .ExpressionID = lngExprID
      .CalculatedColumnsThatUseThisExpression pvColumns
    End With
    Set objCalc = Nothing
  
    rsCheck.MoveNext
  Loop
  ' Close the recordset.
  rsCheck.Close
  
TidyUpAndExit:
  
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
    
End Sub

Public Function TablesThatUseFunction(ByRef pvTables As Variant, plngFunctionID As Long) As Variant
  ' Return an array of the table IDs that use the 'AbsenceDuration' function in calculated columns.
  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  Dim iLoop As Integer
  Dim alngTempColumns() As Long
  Dim fFound As Boolean
  
  ' Work out which tables use the AbsenceDuration function
  ReDim alngTempColumns(0)
  CalculatedColumnsThatUseFunction alngTempColumns, plngFunctionID

  For iCount = 1 To UBound(alngTempColumns)
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", CLng(alngTempColumns(iCount))

      If Not .NoMatch Then
        fFound = False
        For iLoop = 1 To UBound(pvTables)
          If pvTables(iLoop) = !TableID Then
            fFound = True
            Exit For
          End If
        Next iLoop
        
        If Not fFound Then
          ReDim Preserve pvTables(UBound(pvTables) + 1)
          pvTables(UBound(pvTables)) = !TableID
        End If
      End If
    End With
  Next iCount

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function


