VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngTableID As Long
Private mlngColumnID As Long
Private miFieldPassType As Integer
Private miSelectionType As FieldSelectionTypes
Private mlngSelectionLine As Long
Private mlngSelOrderID As Long
Private mlngSelFilterID As Long
Private mstrUDFRuntimeCode As String

' Class handling variables.
Private mobjBaseComponent As clsExprComponent


Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  ContainsExpression = False
  
  If mlngSelFilterID > 0 Then
    ' Check if the calc component IS the one we're checking for.
    ContainsExpression = (plngExprID = mlngSelFilterID)
    
    If Not ContainsExpression Then
      ' The calc component IS NOT the one we're checking for.
      ' Check if it contains the one we're looking for.
      ContainsExpression = HasExpressionComponent(mlngSelFilterID, plngExprID)
    End If
  End If
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function







Public Function EvaluateComponent(palngColumnValues As Variant, pvResult As Variant) As Boolean
  ' Return the evaluated value of the Field component.
  On Error GoTo ErrorTrap

  Dim iIndex As Integer

  If (miFieldPassType = giPASSBY_REFERENCE) Then
    pvResult = mlngColumnID
    EvaluateComponent = True
    Exit Function
  Else
    ' Get the column value array passed in.
    For iIndex = 1 To UBound(palngColumnValues, 2)
      If palngColumnValues(1, iIndex) = mlngColumnID Then
        pvResult = palngColumnValues(2, iIndex)
        EvaluateComponent = True
        Exit Function
      End If
    Next iIndex
  End If
  
TidyUpAndExit:
  EvaluateComponent = False
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Public Function GenerateCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  pfUDFCode As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
  ' Return the SQL code for the component.
  ' Used for View Filter and Filter expressions.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim fColumnOK As Boolean
  Dim fParentField As Boolean
  Dim fNewSourceTable As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sCode As String
  Dim sOtherTableName As String
  Dim sOrderCode As String
  Dim sFilterCode As String
  Dim sColumnCode As String
  Dim rsInfo As ADODB.Recordset
  Dim asViews() As String
  Dim avOrderJoinTables() As Variant
  Dim strTableName As String
  Dim strColumnName As String
  Dim objFilterExpr As clsExprExpression
  Dim strUDFReturnType As String
  
  sCode = ""
  fOK = True
  
  If (miFieldPassType = giPASSBY_REFERENCE) Then
    sCode = Trim(Str(mlngColumnID))
  Else
    
    strTableName = modExpression.GetTableName(mlngTableID)
    strColumnName = modExpression.GetColumnName(mlngColumnID)
    
    If mobjBaseComponent.ParentExpression.BaseTableID = mlngTableID Then
      ' The field is in the expression's base table.
      sCode = strTableName & "." & strColumnName
    Else
      
      ' Check if the table is a child or parent of the expression's base table.
      sSQL = "SELECT *" & _
        " FROM ASRSysRelations" & _
        " WHERE parentID = " & Trim(Str(mlngTableID)) & _
        " AND childID = " & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID))
      Set rsInfo = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      With rsInfo
        fParentField = Not (.EOF And .BOF)
        .Close
      End With
      Set rsInfo = Nothing

      If fParentField Then
        ' The field is from a parent table of the expression's base table.
        sCode = strTableName & "." & strColumnName
      
        fNewSourceTable = True
        For iLoop = 1 To UBound(palngSourceTables, 2)
          If (palngSourceTables(1, iLoop) = 0) And _
            (palngSourceTables(2, iLoop) = mlngTableID) Then
            fNewSourceTable = False
            Exit For
          End If
        Next iLoop

        If fNewSourceTable Then
          iNextIndex = UBound(palngSourceTables, 2) + 1
          ReDim Preserve palngSourceTables(2, iNextIndex)
          palngSourceTables(1, iNextIndex) = 0
          palngSourceTables(2, iNextIndex) = mlngTableID
        End If
      
      Else
        
        ' The field is from a child table of the expression's base table.
        sCode = "(" & vbCrLf

        ' Construct the order code if required.
        ' Create an array of tables that need to be joined to make the order valid.
        ' Column 1 = table ID.
        ' Column 2 = table name.
        sOrderCode = ""
        ReDim avOrderJoinTables(2, 0)
        If mlngSelOrderID > 0 Then
          sSQL = "SELECT c.columnName, c.columnID, c.tableID, t.tableName, oi.ascending" & _
            " FROM ASRSysOrderItems oi" & _
            " JOIN ASRSysColumns c ON oi.columnID = c.columnID" & _
            " JOIN ASRSysTables t ON t.tableID = c.tableID" & _
            " WHERE oi.orderID = " & Trim(Str(mlngSelOrderID)) & _
            " AND oi.type = 'O'" & _
            " AND c.columnID = oi.columnID" & _
            " AND c.tableID = t.tableID" & _
            " AND c.dataType <> -4 AND c.datatype <> -3 " & _
            " ORDER BY oi.sequence"

          Set rsInfo = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
          With rsInfo
            
            Do While Not .EOF
              
              ' Construct the order code. Remember that if we are selecting the last record,
              ' we must reverse the ASC/DESC options.
              sOrderCode = sOrderCode & IIf(Len(sOrderCode) > 0, ", ", "") & _
                !TableName & "." & !ColumnName & _
                IIf(miSelectionType = giSELECT_LASTRECORD, IIf(!Ascending, " DESC", ""), IIf(!Ascending, "", " DESC"))
            
              If (!TableID <> mlngTableID) And _
                (!TableID <> mobjBaseComponent.ParentExpression.BaseTableID) Then
                
                ' Check if the table has already been added to the array of tables used in the order.
                fFound = False
                For iNextIndex = 1 To UBound(avOrderJoinTables, 2)
                  If avOrderJoinTables(1, iNextIndex) = !TableID And _
                    (avOrderJoinTables(2, iNextIndex) = !TableName) Then
                    
                    fFound = True
                    Exit For
                  End If
                Next iNextIndex

                If Not fFound Then
                  iNextIndex = UBound(avOrderJoinTables, 2) + 1
                  ReDim Preserve avOrderJoinTables(2, iNextIndex)
                  avOrderJoinTables(1, iNextIndex) = !TableID
                  avOrderJoinTables(2, iNextIndex) = !TableName
                End If
              End If

              .MoveNext
            Loop

            .Close
          End With
          Set rsInfo = Nothing
        End If

        If fOK Then
          ' Create the filter code if required.
          sFilterCode = ""
          If mlngSelFilterID > 0 Then
          
            If mlngSelFilterID = plngFixedExprID Then
              sFilterCode = psFixedSQLCode
            Else
              Set objFilterExpr = New clsExprExpression
              objFilterExpr.ExpressionID = mlngSelFilterID
              objFilterExpr.ConstructExpression
              fOK = objFilterExpr.RuntimeFilterCode(sFilterCode, pfApplyPermissions, pfValidating)
              Set objFilterExpr = Nothing
            End If
          End If
        End If
        
        If fOK Then
          Select Case miSelectionType
            Case giSELECT_FIRSTRECORD, giSELECT_LASTRECORD
              ' First and Last record selection uses the same code here.
              ' The difference is made when creating the 'order by' code above.
              sCode = sCode & _
                "SELECT TOP 1 " & strTableName & "." & strColumnName & vbCrLf & _
                "FROM " & strTableName & vbCrLf

              ' Add the JOIN code for the order.
              For iLoop = 1 To UBound(avOrderJoinTables, 2)
                sCode = sCode & _
                  "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & strTableName & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbCrLf
              Next iLoop
            
              sCode = sCode & _
                "WHERE " & mobjBaseComponent.ParentExpression.BaseTableName & ".id = " & strTableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbCrLf
            
              ' Add the filter code as required.
              If Len(sFilterCode) > 0 Then
                sCode = sCode & _
                  "AND " & strTableName & ".id IN" & vbCrLf & _
                  "(" & vbCrLf & _
                  sFilterCode & _
                  ")" & vbCrLf
              End If

              ' Add the order code as required.
              If mlngSelOrderID > 0 Then
                sCode = sCode & _
                  "ORDER BY " & sOrderCode & vbCrLf
              End If
              
            Case giSELECT_RECORDCOUNT
              ' No need to add the order code as it makes no differnt when selecting the record count.
              sCode = sCode & _
                "SELECT COUNT(" & strTableName & ".id)" & vbCrLf & _
                "FROM " & strTableName & vbCrLf & _
                "WHERE " & mobjBaseComponent.ParentExpression.BaseTableName & ".id = " & strTableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbCrLf
            
              ' Add the filter code as required.
              If Len(sFilterCode) > 0 Then
                sCode = sCode & _
                  "AND " & strTableName & ".id IN" & vbCrLf & _
                  "(" & vbCrLf & _
                  sFilterCode & _
                  ")" & vbCrLf
              End If


            Case giSELECT_RECORDTOTAL
              ' No need to add the order code as it makes no differnt when selecting the record total.
              sCode = sCode & _
                "SELECT SUM(" & strTableName & "." & strColumnName & ")" & vbCrLf & _
                "FROM " & strTableName & vbCrLf & _
                "WHERE " & mobjBaseComponent.ParentExpression.BaseTableName & ".id = " & strTableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbCrLf
            
              ' Add the filter code as required.
              If Len(sFilterCode) > 0 Then
                sCode = sCode & _
                  "AND " & strTableName & ".id IN" & vbCrLf & _
                  "(" & vbCrLf & _
                  sFilterCode & _
                  ")" & vbCrLf
              End If
                
            Case giSELECT_SPECIFICRECORD
              ' Specific for runtime filters.
              
              Select Case ReturnType
                Case giEXPRVALUE_DATE
                  strUDFReturnType = "datetime"
                
                Case giEXPRVALUE_CHARACTER
                  strUDFReturnType = "varchar(MAX)"
                
                Case giEXPRVALUE_NUMERIC
                  strUDFReturnType = "float"
                
                Case giEXPRVALUE_LOGIC
                  strUDFReturnType = "bit"

              End Select
              
              ' Create the udf code for this field
              mstrUDFRuntimeCode = "CREATE FUNCTION udf_ASRSys_" & Trim(Str(mobjBaseComponent.ComponentID)) & "(@PersonnelID float)" & vbCrLf & _
                                   "RETURNS " & strUDFReturnType & vbCrLf & _
                                   "AS" & vbCrLf & "BEGIN" & vbCrLf & _
                                   "   DECLARE @Result " & strUDFReturnType & vbCrLf & _
                                   "   DECLARE GetRecord CURSOR SCROLL FOR "
              
              mstrUDFRuntimeCode = mstrUDFRuntimeCode & _
                "SELECT " & strTableName & "." & strColumnName & vbCrLf & _
                "FROM " & strTableName & vbCrLf

              ' Add the JOIN code for the order.
              For iLoop = 1 To UBound(avOrderJoinTables, 2)
                mstrUDFRuntimeCode = mstrUDFRuntimeCode & _
                  "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & strTableName & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbCrLf
              Next iLoop
            
              mstrUDFRuntimeCode = mstrUDFRuntimeCode & _
                "WHERE @PersonnelID = " & strTableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbCrLf
            
              ' Add the filter code as required.
              If Len(sFilterCode) > 0 Then
                mstrUDFRuntimeCode = mstrUDFRuntimeCode & _
                  "AND " & strTableName & ".id IN" & vbCrLf & _
                  "(" & vbCrLf & _
                  sFilterCode & _
                  ")" & vbCrLf
              End If

              ' Add the order code as required.
              If mlngSelOrderID > 0 Then
                mstrUDFRuntimeCode = mstrUDFRuntimeCode & _
                  "ORDER BY " & sOrderCode & vbCrLf
              End If

              
              ' Finish off udf code
              mstrUDFRuntimeCode = mstrUDFRuntimeCode & _
                "OPEN GetRecord" & vbCrLf & _
                "FETCH ABSOLUTE " & Trim(Str(mlngSelectionLine)) & " FROM GetRecord INTO @Result" & vbCrLf & _
                "CLOSE GetRecord" & vbCrLf & _
                "DEALLOCATE GetRecord" & vbCrLf & _
                "RETURN @Result" & vbCrLf & _
                "END"
              
              sCode = sCode & " [dbo].udf_ASRSys_" & Trim(Str(mobjBaseComponent.ComponentID)) & "(" & mobjBaseComponent.ParentExpression.BaseTableName & ".id)"
              
              fOK = True

           Case Else
              ' Unrecognised child record selection option.
              fOK = False
          End Select

          sCode = sCode & ")"

          ' Add the table name to the list of source tables if it is not already there.
          fNewSourceTable = True
          For iLoop = 1 To UBound(palngSourceTables, 2)
            If (palngSourceTables(1, iLoop) = 0) And _
              (palngSourceTables(2, iLoop) = mlngTableID) Then
              fNewSourceTable = False
              Exit For
            End If
          Next iLoop
          
          If fNewSourceTable Then
            iNextIndex = UBound(palngSourceTables, 2) + 1
            ReDim Preserve palngSourceTables(2, iNextIndex)
            palngSourceTables(1, iNextIndex) = 0
            palngSourceTables(2, iNextIndex) = mlngTableID
          End If
        End If
      End If
    End If
  
    ' If the return type is a date, then convert the datetime value
    ' to a varchar, and then back to a datetime. This gets rid of the time part
    ' of the datetime value, which may cause errors when comparing datetime values.
    If ReturnType = giEXPRVALUE_DATE Then
      sCode = "convert(" & vbCrLf & _
        "datetime, " & vbCrLf & _
        "convert(" & vbCrLf & _
        "varchar(20), " & vbCrLf & _
        sCode & "," & vbCrLf & _
        "101)" & vbCrLf & _
        ")"
    End If
  
    If ReturnType = giEXPRVALUE_NUMERIC Then
      sCode = "convert(" & vbCrLf & _
        "float, " & vbCrLf & _
        sCode & vbCrLf & _
        ")"
    End If
  
    ' JDM - 19/12/01 - Fault 3299 - Problems concatenating strings
    If ReturnType = giEXPRVALUE_CHARACTER Then
      sCode = "IsNull((" & sCode & "),'')"
    End If
  
  End If
  
TidyUpAndExit:
  If fOK Then
    psRuntimeCode = IIf(pfUDFCode, mstrUDFRuntimeCode, sCode)
  Else
    psRuntimeCode = ""
  End If
  
  GenerateCode = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Position the printing.
  With Printer
    .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    Printer.Print ComponentDescription
  End With
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function WriteComponent()
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  sSQL = "INSERT INTO ASRSysExprComponents" & _
    " (componentID, exprID, type," & _
    " fieldTableID, fieldColumnID, fieldPassBy, fieldSelectionRecord," & _
    " fieldSelectionLine, fieldSelectionOrderID, fieldSelectionFilter, valueLogic)" & _
    " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & _
    " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & _
    " " & Trim(Str(giCOMPONENT_FIELD)) & "," & _
    " " & Trim(Str(mlngTableID)) & "," & _
    " " & Trim(Str(mlngColumnID)) & "," & _
    " " & Trim(Str(miFieldPassType)) & "," & _
    " " & Trim(Str(miSelectionType)) & "," & _
    " " & Trim(Str(mlngSelectionLine)) & "," & _
    " " & Trim(Str(mlngSelOrderID)) & "," & _
    " " & Trim(Str(mlngSelFilterID)) & "," & _
    " 0)"
  gADOCon.Execute sSQL, , adCmdText
  
TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objFieldCopy As New clsExprField
  
  ' Copy the component's basic properties.
  With objFieldCopy
    .ColumnID = mlngColumnID
    .FieldPassType = miFieldPassType
    .SelectionLine = mlngSelectionLine
    .SelectionOrderID = mlngSelOrderID
    .SelectionType = miSelectionType
    .SelectionFilterID = mlngSelFilterID
    .TableID = mlngTableID
  End With
  
  Set CopyComponent = objFieldCopy
   
  ' Disassociate object variables.
  Set objFieldCopy = Nothing
   
End Function

Public Property Get ComponentType() As Integer
  ' Return the component type.
  ComponentType = giCOMPONENT_FIELD
  
End Property


Public Property Get SelectionFilterID() As Long
  ' Return the Selection Filter property value.
  SelectionFilterID = mlngSelFilterID

End Property

Public Property Let SelectionFilterID(ByVal plngNewValue As Long)
  ' Set the Selection Filter property value.
  mlngSelFilterID = plngNewValue

End Property

Public Property Get ComponentDescription() As String
  ' Return a description of the field component.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim fChildField As Boolean
  Dim sSQL As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim sSelectionType As String
  Dim rsInfo As Recordset
  
  ' Get the column and table name.
  sSQL = "SELECT ASRSysColumns.columnName, ASRSysTables.tableName" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysColumns.columnID = " & Trim(Str(mlngColumnID))
  Set rsInfo = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsInfo
    fOK = Not (.EOF And .BOF)
  
    If fOK Then
      sColumnName = !ColumnName
      sTableName = !TableName
    Else
      sColumnName = "<unknown>"
      sTableName = "<unknown>"
    End If
  
    .Close
  End With
  Set rsInfo = Nothing

  If fOK Then
    ' Add the selection type description if required.
    If (miFieldPassType = giPASSBY_VALUE) Then
      ' Only give the full description if the field is in a child table of the
      ' expression's parent table.
    
      sSQL = "SELECT *" & _
        " FROM ASRSysRelations" & _
        " WHERE parentID = " & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & _
        " AND childID = " & Trim(Str(mlngTableID))
      Set rsInfo = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      With rsInfo
        fChildField = Not (.EOF And .BOF)
        
        .Close
      End With
      Set rsInfo = Nothing
      
      If fChildField Then
        Select Case miSelectionType
          Case giSELECT_FIRSTRECORD
            sSelectionType = "(first record"
          Case giSELECT_LASTRECORD
            sSelectionType = "(last record"
          Case giSELECT_SPECIFICRECORD
            sSelectionType = "(line " & Trim(Str(mlngSelectionLine))
          Case giSELECT_RECORDTOTAL
            sSelectionType = "(total"
          Case giSELECT_RECORDCOUNT
            sSelectionType = "(record count"
          Case Else
            sSelectionType = "(<unknown>"
        End Select
        
        If mlngSelOrderID > 0 Then
          ' Get the order name.
          sSQL = "SELECT name" & _
            " FROM ASRSysOrders" & _
            " WHERE orderID = " & Trim(Str(mlngSelOrderID))
          Set rsInfo = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
          With rsInfo
            If Not (.BOF And .EOF) Then
              sSelectionType = sSelectionType & ", order by '" & !Name & "'"
            End If
            
            .Close
          End With
          Set rsInfo = Nothing
        End If
      
        If mlngSelFilterID > 0 Then
          ' Get the filter name.
          sSQL = "SELECT name" & _
            " FROM ASRSysExpressions" & _
            " WHERE exprID = " & Trim(Str(mlngSelFilterID))
          Set rsInfo = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
          With rsInfo
            If Not (.BOF And .EOF) Then
              sSelectionType = sSelectionType & ", filter by '" & !Name & "'"
            End If
            
            .Close
          End With
          Set rsInfo = Nothing
        End If
      
        sSelectionType = sSelectionType & ")"
      End If
    End If
  End If
  
TidyUpAndExit:
  ' Return the component description (to be displayed in the expression treeview).
  If miSelectionType = giSELECT_RECORDCOUNT Then
    ComponentDescription = sTableName & " " & sSelectionType
  Else
    ComponentDescription = sTableName & " : " & sColumnName & " " & sSelectionType
  End If
  Exit Property
  
ErrorTrap:
  sTableName = "<unknown>"
  sColumnName = "<unknown>"
  sSelectionType = "<unknown>"
  fOK = False
  Resume TidyUpAndExit

End Property

Public Property Get ReturnType() As Integer
  ' Return the calculation's return type.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iType As ExpressionValueTypes
  Dim sSQL As String
  Dim rsColumn As Recordset
  
  fOK = True
  
  ' If the component returns the record count then
  ' the return type must be numeric; otherwise the
  ' return type is determined by the field type.
  If miSelectionType = giSELECT_RECORDCOUNT Then
    iType = giEXPRVALUE_NUMERIC
  Else
    ' Determine the field's type by creating an
    ' instance of the column class, and instructing
    ' it to read its own details (including type).
    sSQL = "SELECT dataType" & _
      " FROM ASRSysColumns" & _
      " WHERE columnID = " & Trim(Str(mlngColumnID))
    Set rsColumn = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsColumn
  
      fOK = Not (.EOF And .BOF)
    
      If fOK Then
        Select Case !DataType
          Case sqlNumeric, sqlInteger
            iType = giEXPRVALUE_NUMERIC
          Case sqlDate
            iType = giEXPRVALUE_DATE
          Case sqlVarchar, sqlLongVarChar
            iType = giEXPRVALUE_CHARACTER
          Case sqlBoolean
            iType = giEXPRVALUE_LOGIC
          Case sqlTypeOle
            iType = giEXPRVALUE_OLE
          Case sqlVarBinary
            iType = giEXPRVALUE_PHOTO
          Case Else
            fOK = False
        End Select
        
        If fOK Then
          If miFieldPassType = giPASSBY_REFERENCE Then
            iType = iType + giEXPRVALUE_BYREF_OFFSET
          End If
        End If
      End If
    
      .Close
    End With
    Set rsColumn = Nothing
  End If
  
TidyUpAndExit:
  If fOK Then
    ReturnType = iType
  Else
    ReturnType = giEXPRVALUE_UNDEFINED
  End If
  Exit Property
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
    
End Property




Public Property Get SelectionOrderID() As Long
  ' Return the Selection Order property value.
  SelectionOrderID = mlngSelOrderID

End Property

Public Property Let SelectionOrderID(ByVal plngNewValue As Long)
  ' Set the Selection Order property value.
  mlngSelOrderID = plngNewValue

End Property

Public Property Get SelectionType() As Integer
  ' Return the selection type.
  SelectionType = miSelectionType
  
End Property

Public Property Let SelectionType(ByVal piNewValue As Integer)
  ' Set the selection type.
'  If mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_STATICFILTER Then
'    miSelectionType = giSELECT_FIRSTRECORD
'  Else
    miSelectionType = piNewValue
'  End If
  
End Property

Public Property Get SelectionLine() As Long
  ' Return the record slection line property.
  SelectionLine = mlngSelectionLine
  
End Property

Public Property Let SelectionLine(ByVal plngNewValue As Long)
  ' Set the record slection line property.
  mlngSelectionLine = plngNewValue

End Property

Public Property Get FieldPassType() As Integer
  ' Return the field pass type property.
  FieldPassType = miFieldPassType
  
End Property

Public Property Let FieldPassType(ByVal piNewValue As Integer)
  ' Set the field pass type property.
  miFieldPassType = piNewValue

End Property

Public Property Get ColumnID() As Long
  ' Return the column id property.
  ColumnID = mlngColumnID

End Property

Public Property Let ColumnID(ByVal plngNewValue As Long)
  ' Set the column id property.
  mlngColumnID = plngNewValue

End Property

Public Property Get BaseComponent() As clsExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Set BaseComponent(ByVal pobjNewValue As clsExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property





Public Property Get TableID() As Long
  ' Return the table id property.
  TableID = mlngTableID
  
End Property

Public Property Let TableID(ByVal plngNewValue As Long)
  ' Set the table id property.
  mlngTableID = plngNewValue
  
End Property

Private Sub Class_Initialize()
  ' Initialise properties.
  miFieldPassType = giPASSBY_VALUE

End Sub

Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Copy the component definition to the clipboard object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sClipboardMessage As String
  
  fOK = True
  
  sClipboardMessage = String(piLevel, vbTab) & ComponentDescription & vbCrLf
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
TidyUpAndExit:
  CopyToClipboard = fOK
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
  
  RuntimeCode = GenerateCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, False, plngFixedExprID, psFixedSQLCode)

End Function


Public Function UDFCode(psRuntimeCode() As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
  Dim strUDFCode As String
  
  UDFCode = GenerateCode(strUDFCode, palngSourceTables, pfApplyPermissions, pfValidating, True, plngFixedExprID, psFixedSQLCode)
  
  If Len(strUDFCode) > 0 Then
    ReDim Preserve psRuntimeCode(UBound(psRuntimeCode) + 1)
    psRuntimeCode(UBound(psRuntimeCode)) = strUDFCode
  End If

End Function

