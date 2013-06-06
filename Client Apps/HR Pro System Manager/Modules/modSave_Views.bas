Attribute VB_Name = "modSave_Views"
Option Explicit

Public Function SaveViews(pfRefreshDatabase As Boolean) As Boolean
  ' Save any new or modified View definitions to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objFilter As CExpression
  Dim alngTempColumns() As Long
  Dim iCount As Integer
  Dim fChanged As Boolean
  
  fOK = True
  
  With recViewEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      If !Deleted Then
        fOK = ViewDelete
      End If
      
      .MoveNext
    Loop
  
  
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      If Not !Deleted Then
        If !New Then
          fOK = ViewNew
        ElseIf !Changed Or pfRefreshDatabase Then
          fOK = ViewSave
        Else
          ' JPD20021127 Fault 4325 - Check if the view's filter expression has changed.
          If !ExpressionID > 0 Then
            recExprEdit.Index = "idxExprID"
            recExprEdit.Seek "=", !ExpressionID, False
      
            If Not recExprEdit.NoMatch Then
              If recExprEdit!Changed Then
                fOK = ViewSave
              Else
                'JPD 20051122 Fault 10549
                  Set objFilter = New CExpression
        
                  objFilter.ExpressionID = !ExpressionID
                  If objFilter.ConstructExpression Then
                    ' Work out which columns are used in this filter.
                    ReDim alngTempColumns(0)
                    objFilter.ColumnsUsedInThisExpression alngTempColumns
                    
                    fChanged = False
                    For iCount = 1 To UBound(alngTempColumns)
                      With recColEdit
                        .Index = "idxColumnID"
                        .Seek "=", CLng(alngTempColumns(iCount))
  
                        If Not .NoMatch Then
                          If .Fields("changed").Value Then
                            fChanged = True
                            Exit For
                          End If
                        End If
                      End With
                    Next iCount
                    
                    If fChanged Then
                      fOK = ViewSave
                    End If
                  End If
                  Set objFilter = Nothing
              End If
            End If
          End If
        End If
      End If
      
      .MoveNext
    Loop
  
  
  End With

TidyUpAndExit:
  SaveViews = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error saving views"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ViewDelete() As Boolean
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  ' Delete the view info from the ASRSysViews table on the server.
  sSQL = "DELETE FROM ASRSysViews " & _
          "WHERE ViewID = " & recViewEdit.Fields("ViewID").Value
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  ' Delete the columns from the ASRSysViewColumns table on the server.
  sSQL = "DELETE FROM ASRSysViewColumns " & _
          "WHERE ViewID = " & recViewEdit.Fields("ViewID").Value
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  ' Delete the view screens from the ASRSysViewScreens table on the server.
  sSQL = "DELETE FROM ASRSysViewScreens " & _
          "WHERE ViewID = " & recViewEdit.Fields("ViewID").Value
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  ' Drop the view from the table on the server.
  ' AE20080307 Fault #12951
'  sSQL = "IF EXISTS " & _
'          "(SELECT Name " & _
'          "FROM sysobjects " & _
'          "WHERE id = object_id('dbo." & recViewEdit.Fields("OriginalViewName").Value & "') " & _
'          "AND sysstat & 0xf = 2) " & _
'          "DROP VIEW dbo." & recViewEdit.Fields("OriginalViewName").Value
  sSQL = "IF EXISTS " & _
          "(SELECT Name " & _
          "FROM sysobjects " & _
          "WHERE id = object_id('dbo." & recViewEdit.Fields("viewname").Value & "') " & _
          "AND sysstat & 0xf = 2) " & _
          "DROP VIEW dbo." & recViewEdit.Fields("viewname").Value
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  ViewDelete = fOK
  Exit Function

ErrorTrap:
  fOK = False
  'MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  OutputError "Error deleting view"
  Resume TidyUpAndExit

End Function

Private Function ViewNew() As Boolean
  On Error GoTo ErrorTrap
  ' Saves a new view definition to the server database.

  Dim fOK As Boolean
  Dim iNonSystemColumnsCount As Integer
  Dim sSQL As String
  Dim sTable As String
  Dim sColumns As String
  Dim sWhereClauseCode As String
  Dim rsColumns As dao.Recordset
  Dim objExpr As CExpression
  
  fOK = True
  
  'MH20020809 Remove reference to "viewAlternativeName" column
  ' Insert the view info into the ASRSysViews Table on the server.
  'sSQL = "INSERT INTO ASRSysViews" & _
    " (viewID, viewName, viewDescription, viewTableID, viewSQL, viewAlternativeName, expressionID)" & _
    "VALUES (" & recViewEdit.Fields("ViewID") & ", " & _
    "'" & recViewEdit.Fields("ViewName") & "', " & _
    "'" & recViewEdit.Fields("ViewDescription") & "', " & _
    recViewEdit.Fields("ViewTableID") & ", " & _
    "'" & recViewEdit.Fields("ViewSQL") & "', " & _
    "'" & recViewEdit.Fields("ViewAlternativeName") & "', " & _
    recViewEdit.Fields("ExpressionID") & ")"
  sSQL = "INSERT INTO ASRSysViews" & _
    " (viewID, viewName, viewDescription, viewTableID, viewSQL, expressionID)" & _
    "VALUES (" & recViewEdit.Fields("ViewID").Value & ", " & _
    "'" & recViewEdit.Fields("ViewName").Value & "', " & _
    "'" & recViewEdit.Fields("ViewDescription").Value & "', " & _
    recViewEdit.Fields("ViewTableID").Value & ", " & _
    "'" & recViewEdit.Fields("ViewSQL").Value & "', " & _
    recViewEdit.Fields("ExpressionID").Value & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  ' Insert the columns into the ASRSysViewColumns table on the server.
  With recViewColEdit
    .Index = "idxViewID"
    .Seek "=", recViewEdit.Fields("ViewID").Value
    If Not .NoMatch Then
      Do While Not .EOF
        
        If .Fields("viewID").Value <> recViewEdit.Fields("ViewID").Value Then
          Exit Do
        End If
      
        sSQL = "INSERT INTO ASRSysViewColumns" & _
          " (viewID, columnID, inView)" & _
          " VALUES (" & .Fields("ViewID").Value & ", " & _
          .Fields("ColumnID") & ", " & _
          IIf(.Fields("InView").Value, 1, 0) & ")"
        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        
        .MoveNext
      Loop
    End If
  End With
  
  ' Insert the view screens into the ASRSysViewScreens table on the server.
  With recViewScreens
    .Index = "idxViewID"
    .Seek "=", recViewEdit.Fields("ViewID").Value
    If Not .NoMatch Then
      Do While Not .EOF
        
        If .Fields("viewID").Value <> recViewEdit.Fields("ViewID").Value Then
          Exit Do
        End If
      
        sSQL = "INSERT INTO ASRSysViewScreens" & _
          " (screenID, viewID)" & _
          " VALUES (" & .Fields("ScreenID").Value & ", " & _
          .Fields("ViewID").Value & ") "
        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        
        .MoveNext
      Loop
    End If
  End With
                             
  ' Create the view in SQL Server.
  
  ' Now get the table name
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", recViewEdit.Fields("ViewTableID").Value
    sTable = Trim(recTabEdit.Fields("TableName").Value)
  End With
  
  ' First get the non-system and non-link columns.
  iNonSystemColumnsCount = 0
  sSQL = "SELECT tmpColumns.ColumnName" & _
    " FROM tmpViewColumns, tmpColumns" & _
    " WHERE (tmpViewColumns.ColumnID = tmpColumns.ColumnID" & _
    " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
    " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK)) & _
    " AND tmpViewColumns.InView = TRUE" & _
    " AND tmpViewColumns.ViewID = " & recViewEdit.Fields("ViewID").Value & ")" & _
    " ORDER BY tmpColumns.ColumnName"
  Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  sColumns = vbNullString
  With rsColumns
    While Not .EOF
      sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & "." & Trim(.Fields("ColumnName").Value) & vbNewLine
      iNonSystemColumnsCount = iNonSystemColumnsCount + 1
      .MoveNext
    Wend
  End With
  Set rsColumns = Nothing
  
  ' The must be at least one non-system/non-link column in the view.
  fOK = (iNonSystemColumnsCount > 0)
  
  If Not fOK Then
    MsgBox "At least one column must be included in the '" & recViewEdit!ViewName & "' view.", _
      vbCritical + vbOKOnly, App.Title
  Else
  
    ' Add System and Link columns.
    sSQL = "SELECT tmpColumns.ColumnName" & _
      " FROM tmpColumns" & _
      " WHERE (tmpColumns.columnType = " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
      " OR tmpColumns.columnType = " & Trim$(Str$(giCOLUMNTYPE_LINK)) & ")" & _
      " AND tmpColumns.tableID = " & Trim(Str(recViewEdit!ViewTableID)) & _
      " AND tmpColumns.deleted = FALSE" & _
      " ORDER BY tmpColumns.ColumnName"
    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsColumns
      While Not .EOF
        sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & "." & Trim(.Fields("ColumnName").Value) & vbNewLine
        .MoveNext
      Wend
    End With
    Set rsColumns = Nothing

    ' Add the TimeStamp column.
    sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & ".TimeStamp" & vbNewLine

    ' Get the 'where clause' code from the expression.
    Set objExpr = New CExpression
    objExpr.ExpressionID = recViewEdit!ExpressionID
    sWhereClauseCode = objExpr.ViewFilterCode
    Set objExpr = Nothing
  
    ' Now create the view
    sSQL = "CREATE VIEW dbo." & recViewEdit.Fields("ViewName").Value & vbNewLine & _
      "AS" & vbNewLine & _
      "    SELECT " & sColumns & vbNewLine & _
      "    FROM " & sTable & vbNewLine & _
      IIf(LenB(sWhereClauseCode) <> 0, "    WHERE " & sWhereClauseCode, vbNullString)
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  End If
  
TidyUpAndExit:
  Set rsColumns = Nothing
  Set objExpr = Nothing
  ViewNew = fOK
  Exit Function

ErrorTrap:
  fOK = False
  OutputError "Error creating view"
  Resume TidyUpAndExit

End Function



Private Function ViewSave() As Boolean
  ' Modify a view definition in the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iNonSystemColumnsCount As Integer
  Dim sSQL As String
  Dim sTable As String
  Dim sColumns As String
  Dim sWhereClauseCode As String
  Dim rsColumns As dao.Recordset
  Dim objExpr As CExpression
  
  fOK = True
  
  ' Update the view info in the ASRSysViews Table on the server.
  
  'MH20040426 Fault 8352
  'sSQL = "UPDATE ASRSysViews" & _
    " SET ViewDescription = '" & recViewEdit.Fields("ViewDescription") & "'," & _
    " ViewName = '" & recViewEdit.Fields("ViewName") & "'," & _
    " ExpressionID = " & recViewEdit.Fields("ExpressionID") & _
    " WHERE ViewID = " & recViewEdit.Fields("ViewID")
  sSQL = "UPDATE ASRSysViews" & _
    " SET ViewDescription = '" & Replace(recViewEdit.Fields("ViewDescription").Value, "'", "''") & "'," & _
    " ViewName = '" & recViewEdit.Fields("ViewName").Value & "'," & _
    " ExpressionID = " & recViewEdit.Fields("ExpressionID").Value & _
    " WHERE ViewID = " & recViewEdit.Fields("ViewID").Value
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  ' Update the columns in the ASRSysViewColumns table on the server.
  With recViewColEdit
    .Index = "idxViewID"
    .Seek "=", recViewEdit.Fields("ViewID").Value
    
    If Not .NoMatch Then
      Do While Not .EOF
      
        If .Fields("viewID").Value <> recViewEdit.Fields("ViewID").Value Then
          Exit Do
        End If
      
        If .Fields("changed").Value Then
          sSQL = "UPDATE ASRSysViewColumns" & _
            " SET inView=" & IIf(.Fields("InView").Value, 1, 0) & _
            " WHERE viewID=" & recViewEdit.Fields("ViewID").Value & _
            " AND columnID=" & .Fields("columnID").Value
          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        ElseIf .Fields("new").Value Then
          sSQL = "INSERT INTO ASRSysViewColumns" & _
            " (viewID, columnID, inView)" & _
            " VALUES (" & .Fields("ViewID").Value & ", " & _
            .Fields("ColumnID").Value & ", " & _
            IIf(.Fields("InView").Value, 1, 0) & ")"
          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
  ' Decide what to do with the view screens.
  With recViewScreens
    .Index = "idxViewID"
    .Seek "=", recViewEdit.Fields("ViewID").Value
    If Not .NoMatch Then
      Do While Not .EOF
        If .Fields("viewID").Value <> recViewEdit.Fields("ViewID").Value Then
          Exit Do
        End If
      
        ' Decide if they are new or should be deleted
        If .Fields("deleted").Value Then
          sSQL = "DELETE FROM ASRSysViewScreens " & _
                  "WHERE ScreenID = " & .Fields("ScreenID") & " " & _
                  "AND ViewID = " & .Fields("ViewID")
          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        ElseIf .Fields("new").Value Then
          sSQL = "INSERT INTO ASRSysViewScreens" & _
            " (screenID, viewID)" & _
            " VALUES (" & .Fields("ScreenID").Value & ", " & _
            .Fields("ViewID").Value & ") "
          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        End If
        .MoveNext
      Loop
    End If
  End With
  
  ' Now get the table name
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", recViewEdit.Fields("ViewTableID").Value
    sTable = Trim(recTabEdit.Fields("TableName").Value)
  End With
  
  ' Recreate the view in SQL Server
  
  ' First get the columns
  iNonSystemColumnsCount = 0
  sSQL = "SELECT tmpColumns.ColumnName" & _
    " FROM tmpViewColumns, tmpColumns" & _
    " WHERE (tmpViewColumns.ColumnID = tmpColumns.ColumnID" & _
    " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
    " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK)) & _
    " AND tmpViewColumns.InView = TRUE" & _
    " AND tmpViewColumns.ViewID = " & recViewEdit.Fields("ViewID").Value & ")" & _
    " ORDER BY tmpColumns.ColumnName"
  Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  sColumns = vbNullString
  With rsColumns
    While Not .EOF
      sColumns = sColumns & IIf(LenB(sColumns) = 0, vbNullString, ", ") & sTable & "." & Trim(.Fields("ColumnName").Value) & vbNewLine
      iNonSystemColumnsCount = iNonSystemColumnsCount + 1
      .MoveNext
    Wend
  End With
  
  ' The must be at least one non-system/non-link column in the view.
  fOK = (iNonSystemColumnsCount > 0)
  
  If Not fOK Then
    MsgBox "At least one column must be included in the '" & recViewEdit!ViewName & "' view.", _
      vbCritical + vbOKOnly, App.Title
  Else
  
    ' Add System and Link columns.
    sSQL = "SELECT tmpColumns.ColumnName" & _
      " FROM tmpColumns" & _
      " WHERE (tmpColumns.columnType = " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
      " OR tmpColumns.columnType = " & Trim$(Str$(giCOLUMNTYPE_LINK)) & ")" & _
      " AND tmpColumns.tableID = " & Trim(Str(recViewEdit!ViewTableID)) & _
      " AND tmpColumns.deleted = FALSE" & _
      " ORDER BY tmpColumns.ColumnName"
    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsColumns
      While Not .EOF
        sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & "." & Trim(.Fields("ColumnName").Value) & vbNewLine
        .MoveNext
      Wend
    End With
    Set rsColumns = Nothing
  
    ' Add the TimeStamp column.
    sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & ".TimeStamp" & vbNewLine
    
    If fOK Then
      ' Now drop the view if it exists
      ' Drop the view from SQL Server
      sSQL = "IF EXISTS " & _
              "(SELECT Name " & _
              "FROM sysobjects " & _
              "WHERE id = object_id('dbo." & recViewEdit.Fields("OriginalViewName").Value & "') " & _
              "AND sysstat & 0xf = 2) " & _
              "DROP VIEW dbo." & recViewEdit.Fields("OriginalViewName").Value
      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
      
      ' Get the 'where clause' code from the expression.
      Set objExpr = New CExpression
      objExpr.ExpressionID = recViewEdit!ExpressionID
      sWhereClauseCode = objExpr.ViewFilterCode
      Set objExpr = Nothing
  
      If fOK Then
        ' Now create the view
        sSQL = "CREATE VIEW dbo." & recViewEdit.Fields("ViewName").Value & vbNewLine & _
          "AS" & vbNewLine & _
          "    SELECT " & sColumns & vbNewLine & _
          "    FROM " & sTable & vbNewLine & _
          IIf(LenB(sWhereClauseCode) <> 0, "    WHERE " & sWhereClauseCode, vbNullString)
        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
      End If
    End If
  End If
  
TidyUpAndExit:
  Set objExpr = Nothing
  ViewSave = fOK
  Exit Function

ErrorTrap:
  fOK = False
  'MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  OutputError "Error updating view"
  Resume TidyUpAndExit

End Function


