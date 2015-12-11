Attribute VB_Name = "modSysMgr"
Option Explicit

'Constants
Const ChunkSize = 2 ^ 14

Private gasTableViewPrivileges() As String

Public Sub DisplayApplication()
  'JPD 20030908 Fault 5756
  
  'JPD 20030917 Fault 6991
  If (glngWindowLeft = 0) And _
    (glngWindowTop = 0) And _
    (glngWindowWidth = 0) And _
    (glngWindowHeight = 0) Then
    
    Exit Sub
  End If
  
  frmSysMgr.WindowState = IIf(giWindowState = vbMinimized, vbNormal, giWindowState)
  If frmSysMgr.WindowState = vbNormal Then
    frmSysMgr.Left = glngWindowLeft
    frmSysMgr.Top = glngWindowTop
    frmSysMgr.Width = glngWindowWidth
    frmSysMgr.Height = glngWindowHeight
  End If

End Sub

Public Function IsChildOfTable(lngParentTableID As Long, lngChildTableID As Long) As Boolean

  'Checks if the passed child table is a child of the passed parent table.

  IsChildOfTable = False
  
  With recRelEdit
    .Index = "idxChildParentID"
    .Seek ">=", lngChildTableID, lngParentTableID
    
    If Not .NoMatch Then
      Do While Not .EOF
        If !parentID = lngParentTableID And !childID = lngChildTableID Then
          IsChildOfTable = True
          Exit Do
        End If
        .MoveNext
      Loop
    Else
      IsChildOfTable = False
    End If
  End With

End Function

' Is this version of SQL 2008 or above
Public Function IsVersion10() As Boolean
  IsVersion10 = (glngSQLVersion >= 10)
End Function


Sub Main()

  gsApplicationPath = App.Path

  ' If we get problems, just in case...
  gbDisableCodeJock = (InStr(LCase(Command$), "/skin=false") > 0)

  m_fSafeForScripting = True
  m_fSafeForInitializing = True

  ASRDEVELOPMENT = Not vbCompiled

  ' Default logged on user information
  gstrWindowsCurrentDomain = Environ("USERDOMAIN")
  gstrWindowsCurrentUser = Environ("USERNAME")

  'Instantiate public classes
  Set Application = New SystemMgr.clsApplication
  Set ODBC = New SystemMgr.clsODBC

  'Instantiate Progress Bar class
  'Set gobjProgress = New COAProgress.COA_Progress
  Set gobjProgress = New clsProgress
  gobjProgress.StyleResource = CodeJockStylePath
  gobjProgress.StyleIni = CodeJockStyleIni
   
  If App.StartMode = vbSModeAutomation Then
    'If started via OLE automation, return control back to client application
    Exit Sub
  ElseIf App.StartMode = vbSModeStandalone Then
    'Login to database
    If Login Then
      'Display splash screen
      frmSplash.Show
      frmSplash.Refresh
      
      ApplyHotfixes (BEFORELOAD)
      
      'Activate System Manager
      Activate
      
      'Unload splash screen
      UnLoad frmSplash
      Screen.MousePointer = vbDefault

    End If
  End If
End Sub

Private Function GrantTableViewPrivileges(psTableViewName As String) As Boolean
  ' Restore the privilege settings in the global array gasTableViewPrivileges.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fGoodGroup As Boolean
  Dim fInsertGranted As Boolean
  Dim fDeleteGranted As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sUserName As String
  Dim sUserGroupName As String
  Dim sCurrentGroupName As String
  Dim rsUserInfo As ADODB.Recordset
  Dim rsGroups As ADODB.Recordset
  Dim rsSysRoles As ADODB.Recordset
  Dim asFixedRoles() As String

  fOK = True
  Set rsUserInfo = New ADODB.Recordset
  Set rsGroups = New ADODB.Recordset
  Set rsSysRoles = New ADODB.Recordset
  
  ' Clear the table/view privileges array if it describes privileges for a different table/view to the given one.
  If UBound(gasTableViewPrivileges, 2) > 0 Then
    If UCase$(Trim$(gasTableViewPrivileges(1, 1))) <> UCase$(Trim$(psTableViewName)) Then
      ReDim gasTableViewPrivileges(3, 0)
    End If
  End If
  
  ' Get a list of User Groups (Roles) from SQL Server
  rsGroups.Open "sp_helprole", gADOCon, adOpenForwardOnly, adLockReadOnly
  
  ' Creat an array of the standard system roles for SQL Server 7.0
  ReDim asFixedRoles(0)
  asFixedRoles(0) = "PUBLIC"
  rsSysRoles.Open "sp_helpdbfixedrole", gADOCon, adOpenForwardOnly, adLockReadOnly
  
  With rsSysRoles
    Do While Not .EOF
      iNextIndex = UBound(asFixedRoles) + 1
      ReDim Preserve asFixedRoles(iNextIndex)
      asFixedRoles(iNextIndex) = UCase(Trim(.Fields(0).value))
      .MoveNext
    Loop
    
    .Close
  End With
  
  ' Grant/Deny INSERT and DELETE privileges for each User Group (Role).
  With rsGroups
    If Not .EOF And Not .BOF Then
      While Not .EOF
        sCurrentGroupName = UCase(Trim(.Fields(0).value))
      
        ' Check that the group is valid (ie. not a system user Group (Role).
        fGoodGroup = True
        For iNextIndex = 0 To UBound(asFixedRoles)
          If asFixedRoles(iNextIndex) = sCurrentGroupName Then
            fGoodGroup = False
            Exit For
          End If
        Next iNextIndex
        
        If fGoodGroup Then
          fInsertGranted = False
          fDeleteGranted = False
            
          ' Find any descriptions of privileges for users in the current User Group in the 'gasTableViewPrivileges' array.
          For iLoop = 1 To UBound(gasTableViewPrivileges, 2)
            ' Get the primary User Group (Role) of the current user.
            sUserGroupName = vbNullString
            sUserName = UCase$(Trim(gasTableViewPrivileges(2, iLoop)))
            
            sSQL = "SELECT su1.name AS groupName" & _
              " FROM sysusers su1, sysusers su2" & _
              " WHERE su2.name = '" & sUserName & "'" & _
              " AND su2.gid = su1.uid"
            
            rsUserInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

            With rsUserInfo
              If Not (.EOF And .BOF) Then
                sUserGroupName = UCase(Trim(IIf(.Fields("groupName").value = "public", vbNullString, !GroupName!)))
              End If
              .Close
            End With
            
            If sCurrentGroupName = sUserGroupName Then
              If (gasTableViewPrivileges(3, iLoop) = "INSERT") Then
                fInsertGranted = True
              End If
              If (gasTableViewPrivileges(3, iLoop) = "DELETE") Then
                fDeleteGranted = True
              End If
            End If
            
            If fInsertGranted And fDeleteGranted Then
              Exit For
            End If
          Next iLoop
        
          ' Grant/Deny the INSERT privileges for this User Group (Role) as required.
          If fInsertGranted Then
            sSQL = "GRANT INSERT" & _
              " ON " & psTableViewName & _
              " TO [" & sCurrentGroupName & "]"
            gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
          Else
            sSQL = "DENY INSERT" & _
              " ON " & psTableViewName & _
              " TO [" & sCurrentGroupName & "]"
            gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
          End If
          
          ' Grant/Deny the DELETE privileges for this User Group (Role) as required.
          If fDeleteGranted Then
            sSQL = "GRANT DELETE" & _
              " ON " & psTableViewName & _
              " TO [" & sCurrentGroupName & "]"
            gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
          Else
            sSQL = "DENY DELETE" & _
              " ON " & psTableViewName & _
              " TO [" & sCurrentGroupName & "]"
            gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
          End If
        End If
        
        .MoveNext
      Wend
    End If
  
    .Close
  End With
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsUserInfo = Nothing
  Set rsSysRoles = Nothing
  Set rsGroups = Nothing
  
  GrantTableViewPrivileges = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

'Public Function ViewDelete() As Boolean
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  ' Delete the view info from the ASRSysViews table on the server.
'  sSQL = "DELETE FROM ASRSysViews " & _
'          "WHERE ViewID = " & recViewEdit.Fields("ViewID").Value
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Delete the columns from the ASRSysViewColumns table on the server.
'  sSQL = "DELETE FROM ASRSysViewColumns " & _
'          "WHERE ViewID = " & recViewEdit.Fields("ViewID").Value
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Delete the view screens from the ASRSysViewScreens table on the server.
'  sSQL = "DELETE FROM ASRSysViewScreens " & _
'          "WHERE ViewID = " & recViewEdit.Fields("ViewID").Value
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Drop the view from the table on the server.
'  sSQL = "IF EXISTS " & _
'          "(SELECT Name " & _
'          "FROM sysobjects " & _
'          "WHERE id = object_id('dbo." & recViewEdit.Fields("OriginalViewName").Value & "') " & _
'          "AND sysstat & 0xf = 2) " & _
'          "DROP VIEW dbo." & recViewEdit.Fields("OriginalViewName").Value
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'TidyUpAndExit:
'  ViewDelete = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  'MsgBox ODBC.FormatError(Err.Description), _
'    vbOKOnly + vbExclamation, Application.Name
'  OutputError "Error deleting view"
'  Resume TidyUpAndExit
'
'End Function
'
'Public Function ViewNew() As Boolean
'  On Error GoTo ErrorTrap
'  ' Saves a new view definition to the server database.
'
'  Dim fOK As Boolean
'  Dim iNonSystemColumnsCount As Integer
'  Dim sSQL As String
'  Dim sTable As String
'  Dim sColumns As String
'  Dim sWhereClauseCode As String
'  Dim rsColumns As dao.Recordset
'  Dim objExpr As CExpression
'
'  fOK = True
'
'  'MH20020809 Remove reference to "viewAlternativeName" column
'  ' Insert the view info into the ASRSysViews Table on the server.
'  'sSQL = "INSERT INTO ASRSysViews" & _
'    " (viewID, viewName, viewDescription, viewTableID, viewSQL, viewAlternativeName, expressionID)" & _
'    "VALUES (" & recViewEdit.Fields("ViewID") & ", " & _
'    "'" & recViewEdit.Fields("ViewName") & "', " & _
'    "'" & recViewEdit.Fields("ViewDescription") & "', " & _
'    recViewEdit.Fields("ViewTableID") & ", " & _
'    "'" & recViewEdit.Fields("ViewSQL") & "', " & _
'    "'" & recViewEdit.Fields("ViewAlternativeName") & "', " & _
'    recViewEdit.Fields("ExpressionID") & ")"
'  sSQL = "INSERT INTO ASRSysViews" & _
'    " (viewID, viewName, viewDescription, viewTableID, viewSQL, expressionID)" & _
'    "VALUES (" & recViewEdit.Fields("ViewID").Value & ", " & _
'    "'" & recViewEdit.Fields("ViewName").Value & "', " & _
'    "'" & recViewEdit.Fields("ViewDescription").Value & "', " & _
'    recViewEdit.Fields("ViewTableID").Value & ", " & _
'    "'" & recViewEdit.Fields("ViewSQL").Value & "', " & _
'    recViewEdit.Fields("ExpressionID").Value & ")"
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Insert the columns into the ASRSysViewColumns table on the server.
'  With recViewColEdit
'    .Index = "idxViewID"
'    .Seek "=", recViewEdit.Fields("ViewID").Value
'    If Not .NoMatch Then
'      Do While Not .EOF
'
'        If .Fields("viewID").Value <> recViewEdit.Fields("ViewID").Value Then
'          Exit Do
'        End If
'
'        sSQL = "INSERT INTO ASRSysViewColumns" & _
'          " (viewID, columnID, inView)" & _
'          " VALUES (" & .Fields("ViewID").Value & ", " & _
'          .Fields("ColumnID") & ", " & _
'          IIf(.Fields("InView").Value, 1, 0) & ")"
'        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'        .MoveNext
'      Loop
'    End If
'  End With
'
'  ' Insert the view screens into the ASRSysViewScreens table on the server.
'  With recViewScreens
'    .Index = "idxViewID"
'    .Seek "=", recViewEdit.Fields("ViewID").Value
'    If Not .NoMatch Then
'      Do While Not .EOF
'
'        If .Fields("viewID").Value <> recViewEdit.Fields("ViewID").Value Then
'          Exit Do
'        End If
'
'        sSQL = "INSERT INTO ASRSysViewScreens" & _
'          " (screenID, viewID)" & _
'          " VALUES (" & .Fields("ScreenID").Value & ", " & _
'          .Fields("ViewID").Value & ") "
'        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'        .MoveNext
'      Loop
'    End If
'  End With
'
'  ' Create the view in SQL Server.
'
'  ' Now get the table name
'  With recTabEdit
'    .Index = "idxTableID"
'    .Seek "=", recViewEdit.Fields("ViewTableID").Value
'    sTable = Trim(recTabEdit.Fields("TableName").Value)
'  End With
'
'  ' First get the non-system and non-link columns.
'  iNonSystemColumnsCount = 0
'  sSQL = "SELECT tmpColumns.ColumnName" & _
'    " FROM tmpViewColumns, tmpColumns" & _
'    " WHERE (tmpViewColumns.ColumnID = tmpColumns.ColumnID" & _
'    " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'    " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK)) & _
'    " AND tmpViewColumns.InView = TRUE" & _
'    " AND tmpViewColumns.ViewID = " & recViewEdit.Fields("ViewID").Value & ")" & _
'    " ORDER BY tmpColumns.ColumnName"
'  Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'  sColumns = vbNullString
'  With rsColumns
'    While Not .EOF
'      sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & "." & Trim(.Fields("ColumnName").Value) & vbNewLine
'      iNonSystemColumnsCount = iNonSystemColumnsCount + 1
'      .MoveNext
'    Wend
'  End With
'  Set rsColumns = Nothing
'
'  ' The must be at least one non-system/non-link column in the view.
'  fOK = (iNonSystemColumnsCount > 0)
'
'  If Not fOK Then
'    MsgBox "At least one column must be included in the '" & recViewEdit!ViewName & "' view.", _
'      vbCritical + vbOKOnly, App.Title
'  Else
'
'    ' Add System and Link columns.
'    sSQL = "SELECT tmpColumns.ColumnName" & _
'      " FROM tmpColumns" & _
'      " WHERE (tmpColumns.columnType = " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'      " OR tmpColumns.columnType = " & Trim$(Str$(giCOLUMNTYPE_LINK)) & ")" & _
'      " AND tmpColumns.tableID = " & Trim(Str(recViewEdit!ViewTableID)) & _
'      " AND tmpColumns.deleted = FALSE" & _
'      " ORDER BY tmpColumns.ColumnName"
'    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'    With rsColumns
'      While Not .EOF
'        sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & "." & Trim(.Fields("ColumnName").Value) & vbNewLine
'        .MoveNext
'      Wend
'    End With
'    Set rsColumns = Nothing
'
'    ' Add the TimeStamp column.
'    sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & ".TimeStamp" & vbNewLine
'
'    ' Get the 'where clause' code from the expression.
'    Set objExpr = New CExpression
'    objExpr.ExpressionID = recViewEdit!ExpressionID
'    sWhereClauseCode = objExpr.ViewFilterCode
'    Set objExpr = Nothing
'
'    ' Now create the view
'    sSQL = "CREATE VIEW dbo." & recViewEdit.Fields("ViewName").Value & vbNewLine & _
'      "AS" & vbNewLine & _
'      "    SELECT " & sColumns & vbNewLine & _
'      "    FROM " & sTable & vbNewLine & _
'      IIf(LenB(sWhereClauseCode) <> 0, "    WHERE " & sWhereClauseCode, vbNullString)
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  End If
'
'TidyUpAndExit:
'  Set rsColumns = Nothing
'  Set objExpr = Nothing
'  ViewNew = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error creating view"
'  Resume TidyUpAndExit
'
'End Function
'
'
'
'Public Function ViewSave() As Boolean
'  ' Modify a view definition in the server database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim iNonSystemColumnsCount As Integer
'  Dim sSQL As String
'  Dim sTable As String
'  Dim sColumns As String
'  Dim sWhereClauseCode As String
'  Dim rsColumns As dao.Recordset
'  Dim objExpr As CExpression
'
'  fOK = True
'
'  ' Update the view info in the ASRSysViews Table on the server.
'
'  'MH20040426 Fault 8352
'  'sSQL = "UPDATE ASRSysViews" & _
'    " SET ViewDescription = '" & recViewEdit.Fields("ViewDescription") & "'," & _
'    " ViewName = '" & recViewEdit.Fields("ViewName") & "'," & _
'    " ExpressionID = " & recViewEdit.Fields("ExpressionID") & _
'    " WHERE ViewID = " & recViewEdit.Fields("ViewID")
'  sSQL = "UPDATE ASRSysViews" & _
'    " SET ViewDescription = '" & Replace(recViewEdit.Fields("ViewDescription").Value, "'", "''") & "'," & _
'    " ViewName = '" & recViewEdit.Fields("ViewName").Value & "'," & _
'    " ExpressionID = " & recViewEdit.Fields("ExpressionID").Value & _
'    " WHERE ViewID = " & recViewEdit.Fields("ViewID").Value
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Update the columns in the ASRSysViewColumns table on the server.
'  With recViewColEdit
'    .Index = "idxViewID"
'    .Seek "=", recViewEdit.Fields("ViewID").Value
'
'    If Not .NoMatch Then
'      Do While Not .EOF
'
'        If .Fields("viewID").Value <> recViewEdit.Fields("ViewID").Value Then
'          Exit Do
'        End If
'
'        If .Fields("changed").Value Then
'          sSQL = "UPDATE ASRSysViewColumns" & _
'            " SET inView=" & IIf(.Fields("InView").Value, 1, 0) & _
'            " WHERE viewID=" & recViewEdit.Fields("ViewID").Value & _
'            " AND columnID=" & .Fields("columnID").Value
'          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'        ElseIf .Fields("new").Value Then
'          sSQL = "INSERT INTO ASRSysViewColumns" & _
'            " (viewID, columnID, inView)" & _
'            " VALUES (" & .Fields("ViewID").Value & ", " & _
'            .Fields("ColumnID").Value & ", " & _
'            IIf(.Fields("InView").Value, 1, 0) & ")"
'          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'        End If
'
'        .MoveNext
'      Loop
'    End If
'  End With
'
'  ' Decide what to do with the view screens.
'  With recViewScreens
'    .Index = "idxViewID"
'    .Seek "=", recViewEdit.Fields("ViewID").Value
'    If Not .NoMatch Then
'      Do While Not .EOF
'        If .Fields("viewID").Value <> recViewEdit.Fields("ViewID").Value Then
'          Exit Do
'        End If
'
'        ' Decide if they are new or should be deleted
'        If .Fields("deleted").Value Then
'          sSQL = "DELETE FROM ASRSysViewScreens " & _
'                  "WHERE ScreenID = " & .Fields("ScreenID") & " " & _
'                  "AND ViewID = " & .Fields("ViewID")
'          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'        ElseIf .Fields("new").Value Then
'          sSQL = "INSERT INTO ASRSysViewScreens" & _
'            " (screenID, viewID)" & _
'            " VALUES (" & .Fields("ScreenID").Value & ", " & _
'            .Fields("ViewID").Value & ") "
'          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'        End If
'        .MoveNext
'      Loop
'    End If
'  End With
'
'  ' Now get the table name
'  With recTabEdit
'    .Index = "idxTableID"
'    .Seek "=", recViewEdit.Fields("ViewTableID").Value
'    sTable = Trim(recTabEdit.Fields("TableName").Value)
'  End With
'
'  ' Recreate the view in SQL Server
'
'  ' First get the columns
'  iNonSystemColumnsCount = 0
'  sSQL = "SELECT tmpColumns.ColumnName" & _
'    " FROM tmpViewColumns, tmpColumns" & _
'    " WHERE (tmpViewColumns.ColumnID = tmpColumns.ColumnID" & _
'    " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'    " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK)) & _
'    " AND tmpViewColumns.InView = TRUE" & _
'    " AND tmpViewColumns.ViewID = " & recViewEdit.Fields("ViewID").Value & ")" & _
'    " ORDER BY tmpColumns.ColumnName"
'  Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'  sColumns = vbNullString
'  With rsColumns
'    While Not .EOF
'      sColumns = sColumns & IIf(LenB(sColumns) = 0, vbNullString, ", ") & sTable & "." & Trim(.Fields("ColumnName").Value) & vbNewLine
'      iNonSystemColumnsCount = iNonSystemColumnsCount + 1
'      .MoveNext
'    Wend
'  End With
'
'  ' The must be at least one non-system/non-link column in the view.
'  fOK = (iNonSystemColumnsCount > 0)
'
'  If Not fOK Then
'    MsgBox "At least one column must be included in the '" & recViewEdit!ViewName & "' view.", _
'      vbCritical + vbOKOnly, App.Title
'  Else
'
'    ' Add System and Link columns.
'    sSQL = "SELECT tmpColumns.ColumnName" & _
'      " FROM tmpColumns" & _
'      " WHERE (tmpColumns.columnType = " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'      " OR tmpColumns.columnType = " & Trim$(Str$(giCOLUMNTYPE_LINK)) & ")" & _
'      " AND tmpColumns.tableID = " & Trim(Str(recViewEdit!ViewTableID)) & _
'      " AND tmpColumns.deleted = FALSE" & _
'      " ORDER BY tmpColumns.ColumnName"
'    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'    With rsColumns
'      While Not .EOF
'        sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & "." & Trim(.Fields("ColumnName").Value) & vbNewLine
'        .MoveNext
'      Wend
'    End With
'    Set rsColumns = Nothing
'
'    ' Add the TimeStamp column.
'    sColumns = sColumns & IIf(LenB(sColumns) <> 0, ", ", vbNullString) & sTable & ".TimeStamp" & vbNewLine
'
'    If fOK Then
'      ' Now drop the view if it exists
'      ' Drop the view from SQL Server
'      sSQL = "IF EXISTS " & _
'              "(SELECT Name " & _
'              "FROM sysobjects " & _
'              "WHERE id = object_id('dbo." & recViewEdit.Fields("OriginalViewName").Value & "') " & _
'              "AND sysstat & 0xf = 2) " & _
'              "DROP VIEW dbo." & recViewEdit.Fields("OriginalViewName").Value
'      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'      ' Get the 'where clause' code from the expression.
'      Set objExpr = New CExpression
'      objExpr.ExpressionID = recViewEdit!ExpressionID
'      sWhereClauseCode = objExpr.ViewFilterCode
'      Set objExpr = Nothing
'
'      If fOK Then
'        ' Now create the view
'        sSQL = "CREATE VIEW dbo." & recViewEdit.Fields("ViewName").Value & vbNewLine & _
'          "AS" & vbNewLine & _
'          "    SELECT " & sColumns & vbNewLine & _
'          "    FROM " & sTable & vbNewLine & _
'          IIf(LenB(sWhereClauseCode) <> 0, "    WHERE " & sWhereClauseCode, vbNullString)
'        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'      End If
'    End If
'  End If
'
'TidyUpAndExit:
'  Set objExpr = Nothing
'  ViewSave = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  'MsgBox ODBC.FormatError(Err.Description), _
'    vbOKOnly + vbExclamation, Application.Name
'  OutputError "Error updating view"
'  Resume TidyUpAndExit
'
'End Function



Private Function DefinitionIntegrityCheck() As Boolean
  ' Performs an integrity check on the System databases.
  ' eg. any screen records without an associated table record will be deleted.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  '
  ' Tidy up the Columns table.
  '
  ' Delete any Column records that have no associated Table records.
  sSQL = "DELETE " & _
    " FROM ASRSysColumns" & _
    " WHERE tableID NOT IN (SELECT tableID FROM ASRSysTables)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  '
  ' Tidy up the Column Control Values table.
  '
  ' Delete any Column Control Value records that have no associated Column records.
  sSQL = "DELETE " & _
    " FROM ASRSysColumnControlValues" & _
    " WHERE columnID NOT IN (SELECT columnID FROM ASRSysColumns)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  '
  ' Tidy up the Controls table.
  '
  ' Delete any Control records that have no associated Table, Column or Screen records.
  sSQL = "DELETE " & _
    " FROM ASRSysControls" & _
    " WHERE columnID NOT IN (SELECT columnID FROM ASRSysColumns)" & _
    " OR screenID NOT IN (SELECT screenID FROM ASRSysScreens)" & _
    " OR tableID NOT IN (SELECT tableID FROM ASRSysTables)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the Expression Components table.
  '
  ' Delete any Expression Component records that have no associated Expression records.
  sSQL = "DELETE " & _
    " FROM ASRSysExprComponents" & _
    " WHERE exprID NOT IN (SELECT exprID FROM ASRSysExpressions)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the Screens table.
  '
  ' Delete any Screen records that have no associated Table records.
  sSQL = "DELETE " & _
    " FROM ASRSysScreens" & _
    " WHERE tableID NOT IN (SELECT tableID FROM ASRSysTables)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the History Screens table.
  '
  ' Delete any History Screen records that have no associated Screen records.
  sSQL = "DELETE " & _
    " FROM ASRSysHistoryScreens" & _
    " WHERE parentScreenID NOT IN (SELECT screenID FROM ASRSysScreens)" & _
    " OR historyScreenID NOT IN (SELECT screenID FROM ASRSysScreens)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the Page Captions table.
  '
  ' Delete any Page Caption records that have no associated Screen records.
  sSQL = "DELETE " & _
    " FROM ASRSysPageCaptions" & _
    " WHERE screenID NOT IN (SELECT screenID FROM ASRSysScreens)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the Relations table.
  '
  ' Delete any Relation records that have no associated Table records.
  sSQL = "DELETE " & _
    " FROM ASRSysRelations" & _
    " WHERE parentID NOT IN (SELECT tableID FROM ASRSysTables)" & _
    " OR childID NOT IN (SELECT tableID FROM ASRSysTables)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the Orders table.
  '
  ' Delete any Order records that have no associated Table records.
  sSQL = "DELETE " & _
    " FROM ASRSysOrders" & _
    " WHERE tableID NOT IN (SELECT tableID FROM ASRSysTables)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the Order Items table.
  '
  ' Delete any Order Item records that have no associated Order or Column records.
  sSQL = "DELETE " & _
    " FROM ASRSysOrderItems" & _
    " WHERE columnID NOT IN (SELECT orderID FROM ASRSysColumns)" & _
    " OR orderID NOT IN (SELECT orderID FROM ASRSysOrders)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the Views table.
  '
  ' Delete any View records that have no associated Table records.
  sSQL = "DELETE " & _
    " FROM ASRSysViews" & _
    " WHERE viewTableID NOT IN (SELECT tableID FROM ASRSysTables)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the View Columns table.
  '
  ' Delete any View Column records that have no associated View or Column records.
  sSQL = "DELETE " & _
    " FROM ASRSysViewColumns" & _
    " WHERE viewID NOT IN (SELECT viewID FROM ASRSysViews)" & _
    " OR columnID NOT IN (SELECT columnID FROM ASRSysColumns)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  '
  ' Tidy up the View Screens table.
  '
  ' Delete any View Screen records that have no associated View or Screen records.
  sSQL = "DELETE " & _
    " FROM ASRSysViewScreens" & _
    " WHERE viewID NOT IN (SELECT viewID FROM ASRSysViews)" & _
    " OR screenID NOT IN (SELECT screenID FROM ASRSysScreens)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  DefinitionIntegrityCheck = fOK
  Exit Function
  
ErrorTrap:
  MsgBox Err.Description, vbOKOnly + vbExclamation, Application.Name
  Err = False
  fOK = False
  Resume TidyUpAndExit
  
End Function

'Private Function CopyData() As Boolean
'  ' Copy the data to any cloned tables.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim iSourceColumnDataType As Integer
'  Dim iDestinationColumnSize As Integer
'  Dim iDestinationColumnDecimals As Integer
'  Dim iDestinationColumnDataType As Integer
'  Dim lngSourceTableID As Long
'  Dim lngDestinationTableID As Long
'  Dim dblMaxValue As Double
'  Dim sSQL As String
'  Dim sTempCopy As String
'  Dim sValueList As SystemMgr.cStringBuilder
'  Dim sColumnList As SystemMgr.cStringBuilder
'  Dim sSourceTableName As String
'  Dim sDestinationTableName As String
'  Dim rsTableName As dao.Recordset
'  Dim rsColumnTypes As New ADODB.Recordset
'  Dim rsCommonColumns As New ADODB.Recordset
'  Dim strColumnName As String
'
'  Set sValueList = New SystemMgr.cStringBuilder
'  Set sColumnList = New SystemMgr.cStringBuilder
'  fOK = True
'
'  With recTabEdit
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'
'    Do While (Not .EOF) And fOK
'
'      lngSourceTableID = 0
'      lngDestinationTableID = 0
'
'      If !copyDataTableID > 0 Then
'        lngDestinationTableID = !TableID
'        sDestinationTableName = !TableName
'        lngSourceTableID = !copyDataTableID
'      End If
'
'      If (lngSourceTableID > 0) And (lngDestinationTableID > 0) Then
'
'        ' Get the source table name.
'        sSQL = "SELECT tableName" & _
'          " FROM tmpTables" & _
'          " WHERE tableID=" & Trim$(Str$(lngSourceTableID))
'        Set rsTableName = daoDb.OpenRecordset(sSQL, _
'          dbOpenForwardOnly, dbReadOnly)
'        If Not (rsTableName.BOF And rsTableName.EOF) Then
'          sSourceTableName = rsTableName.Fields("tableName").Value
'        Else
'          fOK = False
'        End If
'        rsTableName.Close
'
'        If fOK Then
'          ' Copy the source table into a temporary table.
'          sTempCopy = GetTempTableName("Tmp_" & sSourceTableName)
'          fOK = Not (sTempCopy = vbNullString)
'        End If
'
'        If fOK Then
'
'          sSQL = "SELECT * INTO " & sTempCopy & _
'            " FROM " & sSourceTableName
'          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'          ' Build list of columns with which to re-populate this table.
'          sColumnList.TheString = vbNullString
'          sValueList.TheString = vbNullString
'
'          ' Get the names of the columns that are common to the source and destination tables.
'          sSQL = "SELECT DISTINCT columnName" & _
'            " FROM ASRSysColumns " & _
'            " WHERE tableID=" & Trim$(Str$(lngDestinationTableID)) & _
'            " AND columnName IN " & _
'            "   (SELECT columnName" & _
'            "     FROM ASRSysColumns" & _
'            "     WHERE tableID=" & Trim$(Str$(lngSourceTableID)) & ")"
'          rsCommonColumns.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly, adCmdText
'
'          With rsCommonColumns
'            While Not .EOF
'
'              ' Get the datatypes of the source and destination columns.
'              strColumnName = .Fields("ColumnName").Value
'              iSourceColumnDataType = 0
'              iDestinationColumnSize = 0
'              iDestinationColumnDecimals = 0
'              iDestinationColumnDataType = 0
'              sSQL = "SELECT tableID, dataType, size, decimals" & _
'                " FROM ASRSysColumns" & _
'                " WHERE columnName='" & strColumnName & "'" & _
'                " AND (tableID=" & Trim$(Str$(lngDestinationTableID)) & _
'                " OR tableID=" & Trim$(Str$(lngSourceTableID)) & ")"
'              rsColumnTypes.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'              With rsColumnTypes
'                While Not .EOF
'                  If !TableID = lngSourceTableID Then
'                    iSourceColumnDataType = !DataType
'                  ElseIf !TableID = lngDestinationTableID Then
'                    iDestinationColumnDataType = !DataType
'
'                    'TM20060615 - Fault 11085
'                    'Specify a size of '14' if a sqlLongVarChar(working pattern) as the size is not
'                    'stored in the ASRSysColumns table for this type.
'                    If !DataType = SQLDataType.sqlLongVarChar Then
'                      iDestinationColumnSize = 14
'                    Else
'                      iDestinationColumnSize = !Size
'                    End If
'                    iDestinationColumnDecimals = !Decimals
'                  End If
'                  .MoveNext
'                Wend
'                .Close
'              End With
'              Set rsColumnTypes = Nothing
'
'              fOK = (iSourceColumnDataType <> 0) And (iDestinationColumnDataType <> 0)
'
'              If fOK Then
'                If iDestinationColumnDataType = iSourceColumnDataType Then
'                  sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
'
'                  Select Case iDestinationColumnDataType
'                    ' Convert character.
'                    Case dtVARCHAR, dtLONGVARCHAR
'                      sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & _
'                        "CONVERT(varchar(" & Trim$(Str$(iDestinationColumnSize)) & ")," & strColumnName & ")"
'
'                    ' Convert numeric.
'                    Case dtNUMERIC
'                      ' Ensure that we don't try to copy any out of range data into the columns.
'                      dblMaxValue = 10 ^ (iDestinationColumnSize - iDestinationColumnDecimals)
'
'                      sSQL = "UPDATE " & sTempCopy & _
'                        " SET " & strColumnName & " = 0" & _
'                        " WHERE " & strColumnName & " >= " & Trim$(Str$(dblMaxValue)) & _
'                        " OR " & strColumnName & " <= -" & Trim$(Str$(dblMaxValue))
'                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'                      sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & _
'                        "CONVERT(numeric(" & Trim$(Str$(iDestinationColumnSize)) & "," & Trim$(Str$(iDestinationColumnDecimals)) & "), " & strColumnName & ")"
'
'                    Case Else
'                      sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & strColumnName
'
'                  End Select
'
'                Else
'                  Select Case iDestinationColumnDataType
'                    ' Convert data into character if possible.
'                    Case dtVARCHAR, dtLONGVARCHAR
'                      If (iSourceColumnDataType = dtTIMESTAMP) Or _
'                        (iSourceColumnDataType = dtINTEGER) Or _
'                        (iSourceColumnDataType = dtNUMERIC) Or _
'                        (iSourceColumnDataType = dtBIT) Then
'                        sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
'                        sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & "CONVERT(varchar(" & Trim$(Str$(iDestinationColumnSize)) & "), " & strColumnName & ")"
'                      End If
'
'                    ' Convert data into integer if possible.
'                    Case dtINTEGER
'                      If (iSourceColumnDataType = dtNUMERIC) Or _
'                        (iSourceColumnDataType = dtBIT) Then
'                        sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
'                        sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & "CONVERT(int, " & strColumnName & ")"
'                      End If
'
'                    ' Convert data into numeric if possible.
'                    Case dtNUMERIC
'                      If (iSourceColumnDataType = dtINTEGER) Or _
'                        (iSourceColumnDataType = dtBIT) Then
'                        sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
'                        sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & "CONVERT(numeric(" & Trim$(Str$(iDestinationColumnSize)) & "," & Trim$(Str$(iDestinationColumnDecimals)) & "), " & strColumnName & ")"
'                      End If
'
'                    ' Cannot convert any other datatype into bit, but we need to initialise it.
'                    Case dtBIT
'                      sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
'                      sValueList.Append IIf(sValueList.Length <> 0, ",0", "0")
'                  End Select
'                End If
'              End If
'
'              .MoveNext
'            Wend
'            .Close
'          End With
'          Set rsCommonColumns = Nothing
'
'          ' Get the names of the logic columns that are only in the destination table (ie. require initialising).
'          sSQL = "SELECT DISTINCT columnName" & _
'            " FROM ASRSysColumns " & _
'            " WHERE tableID=" & Trim$(Str$(lngDestinationTableID)) & _
'            " AND dataType=" & Trim$(Str$(dtBIT)) & _
'            " AND columnName NOT IN " & _
'            "   (SELECT columnName" & _
'            "     FROM ASRSysColumns" & _
'            "     WHERE tableID=" & Trim$(Str$(lngSourceTableID)) & ")"
'          rsCommonColumns.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly, adCmdText
'
'          With rsCommonColumns
'            While Not .EOF
'              sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & !ColumnName
'              sValueList.Append IIf(sValueList.Length <> 0, ",0", "0")
'
'              .MoveNext
'            Wend
'            .Close
'          End With
'          Set rsCommonColumns = Nothing
'
'          If (sColumnList.Length <> 0) And (sValueList.Length <> 0) Then
'
'
''MH20010403 The INSERT caused an error which include "IDENTITY INSERT IS OFF".
''           This was strange 'cos we just turned it on in the execute statement above.
''           Anyway, I combined the three execute statements and that seemed to fix
''           the error....... don't ask me!
'
'            ' Populate the destination table with data from the source table.
''            rdoCon.Execute "SET IDENTITY_INSERT " & sDestinationTableName & " ON"
''            sSQL = "INSERT INTO " & sDestinationTableName & " (" & sColumnList & ")" & _
''              " SELECT " & sValueList & " FROM " & sTempCopy
''            rdoCon.Execute sSQL, rdExecDirect
''            rdoCon.Execute "Set IDENTITY_INSERT " & sDestinationTableName & " OFF"
'
'            gADOCon.Execute _
'                "SET IDENTITY_INSERT " & sDestinationTableName & " ON" & vbNewLine & _
'                "INSERT INTO " & sDestinationTableName & " (" & sColumnList.ToString & ")" & _
'                " SELECT " & sValueList.ToString & " FROM " & sTempCopy & vbNewLine & _
'                "SET IDENTITY_INSERT " & sDestinationTableName & " OFF", , adCmdText + adExecuteNoRecords
'          End If
'        End If
'      End If
'
'      .MoveNext
'
'    Loop
'  End With
'
'TidyUpAndExit:
'
'  ' Drop the temporary table.
'  If LenB(sTempCopy) <> 0 Then
'    sSQL = "IF EXISTS (SELECT Name FROM dbo.sysobjects where id = object_id(N'[dbo].[" & sTempCopy & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" _
'          & " DROP TABLE " & sTempCopy
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'  End If
'
'  'If Not fOK Then
'  '  gobjProgress.Visible = False
'  '  MsgBox "Error copying data." & vbCr & vbCr & _
'  '    Err.Description, vbExclamation + vbOKOnly, App.ProductName
'  'End If
'  ' Disassociate object variables.
'  Set rsTableName = Nothing
'  Set rsCommonColumns = Nothing
'  CopyData = fOK
'  Exit Function
'
'ErrorTrap:
'  On Local Error Resume Next
'  OutputError "Error copying data"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function

''Private Function SaveTables(ByRef psErrMsg As String, pfRefreshDatabase As Boolean, pavOldColumns As Variant) As Boolean
''Private Function SaveTables(pfRefreshDatabase As Boolean, pavOldColumns As Variant) As Boolean
'Private Function SaveTables(pfRefreshDatabase As Boolean) ', pavOldColumns As Variant) As Boolean
'  ' Save the new or modified Table definitions.
'  On Error GoTo ErrorTrap
'
'  Dim objTable As Table
'  Dim fOK As Boolean
'  Dim fCreateMaxIDStoredProcedure As Boolean
'  Dim lngRecordCount As Long
'
'  fOK = True
'  fCreateMaxIDStoredProcedure = False
'
'  With recTabEdit
'    .Index = "idxTableID"
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'      lngRecordCount = .RecordCount
'    End If
'    Do While fOK And Not .EOF
'
'      'Do deleted ones first
'      If !Deleted Then
'        Set objTable = New Table
'        objTable.TableID = !TableID
'        Set mfrmUse = New frmUsage
'        mfrmUse.ResetList
'        If objTable.TableIsUsed(mfrmUse) Then
'          gobjProgress.Visible = False
'          Screen.MousePointer = vbDefault
'          Select Case !TableType
'            Case TableTypes.iTabParent
'              mfrmUse.ShowMessage !TableName & " Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.Table
'            Case TableTypes.iTabChild
'              mfrmUse.ShowMessage !TableName & " Child Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.ChildTable
'            Case TableTypes.iTabLookup
'              mfrmUse.ShowMessage !TableName & " Lookup Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.LookupTable
'          End Select
'
'          fOK = False
'        End If
'        UnLoad mfrmUse
'        Set mfrmUse = Nothing
'
'        gobjProgress.Visible = True
'
'        If fOK Then
'          OutputCurrentProcess2 "Deleting " & recTabEdit!TableName, lngRecordCount
'          gobjProgress.UpdateProgress2
'          fOK = TableDelete
'          fCreateMaxIDStoredProcedure = True
'        Else
'          Exit Do
'        End If
'
'      End If
'
'      fOK = fOK And Not gobjProgress.Cancelled
'      .MoveNext
'    Loop
'
'
'    .Index = "idxTableID"
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'      lngRecordCount = .RecordCount
'    End If
'    Do While fOK And Not .EOF
'
'      'Now do new and changed ones
'      If Not !Deleted Then
'        If !New Then
'          OutputCurrentProcess2 recTabEdit!TableName, lngRecordCount
'          gobjProgress.UpdateProgress2
'          fOK = TableNew
'          fCreateMaxIDStoredProcedure = True
'
'        ElseIf !Changed Or pfRefreshDatabase Then
'          OutputCurrentProcess2 recTabEdit!TableName, lngRecordCount
'          gobjProgress.UpdateProgress2
'          fOK = TableSave
'          fCreateMaxIDStoredProcedure = True
'        End If
'      End If
'
'      fOK = fOK And Not gobjProgress.Cancelled
'      .MoveNext
'    Loop
'
'
'  End With
'
'  ' JPD20030313 Fault 5159
'  If fOK And fCreateMaxIDStoredProcedure Then
'    fOK = CreateMaxIDStoredProcedure
'  End If
'
'TidyUpAndExit:
'  SaveTables = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error saving table definitions"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function

'Private Function CreateMaxIDStoredProcedure() As Boolean
'  ' Create the Max ID stored procedure.
'  ' JPD20030313 Fault 5159
'
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'  Dim sSPCode As String
'
'  Const sSPName = "dbo.spASRMaxID"
'
'  fOK = True
'
''  ' Drop the stored procedure if it already exists.
''  sSQL = "IF EXISTS" & _
''    " (SELECT Name" & _
''    "   FROM sysobjects" & _
''    "   WHERE id = object_id('" & sSPName & "')" & _
''    "     AND sysstat & 0xf = 4)" & _
''    " DROP PROCEDURE " & sSPName
''  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'  DropProcedure sSPName
'
'  ' Create the stored procedure.
'  sSPCode = "CREATE PROCEDURE " & sSPName & vbNewLine & _
'    "(" & vbNewLine & _
'    "    @piTableID integer,              /* Input variable to define the table ID. */" & vbNewLine & _
'    "    @piMaxRecordID integer OUTPUT   /* Output variable to hold the max record ID. */" & vbNewLine & _
'    ")" & vbNewLine & _
'    "AS" & vbNewLine & _
'    "BEGIN" & vbNewLine & _
'    "    SET @piMaxRecordID = 0" & vbNewLine & vbNewLine
'
'  With recTabEdit
'    .Index = "idxTableID"
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'    Do While fOK And Not .EOF
'      sSPCode = sSPCode & vbNewLine & _
'        "    IF @piTableID = " & Trim(Str(recTabEdit!TableID)) & vbNewLine & _
'        "    BEGIN" & vbNewLine & _
'        "        SELECT @piMaxRecordID = MAX(id) FROM " & recTabEdit!TableName & vbNewLine & _
'        "    END" & vbNewLine & vbNewLine
'
'      .MoveNext
'    Loop
'  End With
'
'  sSPCode = sSPCode & _
'    "END"
'
'  gADOCon.Execute sSPCode, , adCmdText + adExecuteNoRecords
'
'TidyUpAndExit:
'  CreateMaxIDStoredProcedure = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function



'Private Function SaveScreens() As Boolean
'  ' Save the new or modified screen definitions.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'
'  fOK = True
'
'  With recScrEdit
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'    Do While fOK And Not .EOF
'      If !Deleted Then
'        fOK = ScreenDelete
'      ElseIf !New Then
'        fOK = ScreenNew
'      ElseIf !Changed Then
'        fOK = ScreenSave
'      End If
'
'      .MoveNext
'    Loop
'  End With
'
'TidyUpAndExit:
'  SaveScreens = fOK
'  Exit Function
'
'ErrorTrap:
'  'MsgBox "Error saving screen definitions" & _
'         IIf(Trim(Err.Description) <> vbnullstring, "(" & Err.Description & ")", vbnullstring), vbCritical
'  OutputError "Error saving screen definitions"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'Private Function SaveWorkflows() As Boolean
'  ' Save the new or modified workflows definitions.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'
'  fOK = True
'
'  With recWorkflowEdit
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'    Do While fOK And Not .EOF
'      If !Deleted Then
'        fOK = WorkflowDelete
'      ElseIf !New Then
'        fOK = WorkflowNew
'      ElseIf !Changed Then
'        fOK = WorkflowSave
'      End If
'
'      .MoveNext
'    Loop
'  End With
'
'  If fOK Then
'    fOK = CreateSP_WorkflowCalculation
'  End If
'
'  If fOK Then
'    fOK = CreateSP_WorkflowParentRecord
'  End If
'
'TidyUpAndExit:
'  SaveWorkflows = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error saving workflow definitions"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function


'Private Function SaveOrders() As Boolean
'  ' Save the new or modified Order definitions to the server database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim objOrder As Order
'
'  fOK = True
'
'  With recOrdEdit
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'    Do While fOK And Not .EOF
'      If !Deleted Then
'        Set objOrder = New Order
'        objOrder.OrderID = !OrderID
'        Set mfrmUse = New frmUsage
'        mfrmUse.ResetList
'        If objOrder.OrderIsUsed(mfrmUse) Then
'          gobjProgress.Visible = False
'          Screen.MousePointer = vbDefault
'          mfrmUse.ShowMessage !Name & " Order", "The order cannot be deleted as the order is used by the following:", UsageCheckObject.Order
'          fOK = False
'        End If
'        UnLoad mfrmUse
'        Set mfrmUse = Nothing
'
'        gobjProgress.Visible = True
'
'        If fOK Then
'          fOK = OrderDelete
'        End If
'
'      ElseIf !New Then
'        fOK = OrderNew
'      ElseIf !Changed Then
'        fOK = OrderSave
'      End If
'
'      .MoveNext
'    Loop
'  End With
'
'TidyUpAndExit:
'  SaveOrders = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error saving orders"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function


'Private Function SaveEmailAddrs() As Boolean
'  ' Save the new or modified Email Address definitions to the server database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim objEmailAddr As clsEmailAddr
'
'  fOK = True
'
'  With recEmailAddrEdit
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'    Do While fOK And Not .EOF
'      If !Deleted Then
'        Set objEmailAddr = New clsEmailAddr
'        objEmailAddr.EmailID = !EmailID
'        Set mfrmUse = New frmUsage
'        mfrmUse.ResetList
'        If objEmailAddr.EmailIsUsed(mfrmUse) Then
'          gobjProgress.Visible = False
'          Screen.MousePointer = vbDefault
'          mfrmUse.ShowMessage !Name & " Email", "The email cannot be deleted as the email is used by the following:", UsageCheckObject.Email
'          fOK = False
'        End If
'        UnLoad mfrmUse
'        Set mfrmUse = Nothing
'
'        gobjProgress.Visible = True
'
'        If fOK Then
'          fOK = EmailAddrDelete
'        End If
'
'      ElseIf !New Then
'        fOK = EmailAddrNew
'      ElseIf !Changed Then
'        fOK = EmailAddrSave
'      End If
'
'      .MoveNext
'    Loop
'  End With
'
'TidyUpAndExit:
'  SaveEmailAddrs = fOK
'  Exit Function
'
'ErrorTrap:
'  'MsgBox "Error creating email addresses" & _
'         IIf(Trim(Err.Description) <> vbnullstring, "(" & Err.Description & ")", vbnullstring), vbCritical
'  OutputError "Error creating email addresses"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function


'Private Function SaveExpressions(pfRefreshDatabase As Boolean) As Boolean
'  ' Save the new and modified Expressions to the server database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim lngExprID As Long
'  Dim lngRecordCount As Long
'  Dim fSave As Boolean
'
'  fOK = True
'
'  With recExprEdit
'    .Index = "idxExprID"
'
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'      lngRecordCount = .RecordCount
'    End If
'
'    OutputCurrentProcess2 vbNullString, lngRecordCount
'
'    Do While fOK And Not .EOF
'      lngExprID = .Fields("exprID").Value
'
'      If !Deleted Then
'        OutputCurrentProcess2 .Fields("Name").Value
'        fOK = ExpressionDelete
'
'      ElseIf !New Then
'        OutputCurrentProcess2 .Fields("Name").Value
'        fOK = ExpressionNew
'
'      Else
'        fSave = !Changed _
'          Or pfRefreshDatabase _
'          Or Application.ChangedTableName _
'          Or Application.ChangedColumnName
'
'        If (Not fSave) _
'          And (!Type = giEXPR_WORKFLOWCALCULATION _
'            Or !Type = giEXPR_WORKFLOWSTATICFILTER _
'            Or !Type = giEXPR_WORKFLOWRUNTIMEFILTER) Then
'
'          ' Check if the workflow's changed.
'          recWorkflowEdit.Index = "idxWorkflowID"
'          recWorkflowEdit.Seek "=", !UtilityID
'
'          If Not recWorkflowEdit.NoMatch Then
'            fSave = recWorkflowEdit!Changed
'          End If
'
'        End If
'
'        If fSave Then
'          If .Fields("ParentComponentID").Value = 0 Then
'            OutputCurrentProcess2 .Fields("Name").Value
'          End If
'          fOK = ExpressionSave
'        Else
'          OutputCurrentProcess2 vbNullString
'        End If
'      End If
'
'      ' Ensure that we are positioned on the correct record
'      ' as the recExprEdit recordset may have been repositioned.
'      .Index = "idxExprID"
'      .Seek ">", lngExprID
'      .MoveNext
'      fOK = fOK And Not gobjProgress.Cancelled
'
'      gobjProgress.UpdateProgress2
'
'    Loop
'  End With
'
'TidyUpAndExit:
'  SaveExpressions = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error saving expressions"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'Private Function SaveViews(pfRefreshDatabase As Boolean) As Boolean
'  ' Save any new or modified View definitions to the server database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim objFilter As CExpression
'  Dim alngTempColumns() As Long
'  Dim iCount As Integer
'  Dim fChanged As Boolean
'
'  fOK = True
'
'  With recViewEdit
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'    Do While fOK And Not .EOF
'      If !Deleted Then
'        fOK = ViewDelete
'      End If
'
'      .MoveNext
'    Loop
'
'
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'    Do While fOK And Not .EOF
'      If Not !Deleted Then
'        If !New Then
'          fOK = ViewNew
'        ElseIf !Changed Or pfRefreshDatabase Then
'          fOK = ViewSave
'        Else
'          ' JPD20021127 Fault 4325 - Check if the view's filter expression has changed.
'          If !ExpressionID > 0 Then
'            recExprEdit.Index = "idxExprID"
'            recExprEdit.Seek "=", !ExpressionID, False
'
'            If Not recExprEdit.NoMatch Then
'              If recExprEdit!Changed Then
'                fOK = ViewSave
'              Else
'                'JPD 20051122 Fault 10549
'                  Set objFilter = New CExpression
'
'                  objFilter.ExpressionID = !ExpressionID
'                  If objFilter.ConstructExpression Then
'                    ' Work out which columns are used in this filter.
'                    ReDim alngTempColumns(0)
'                    objFilter.ColumnsUsedInThisExpression alngTempColumns
'
'                    fChanged = False
'                    For iCount = 1 To UBound(alngTempColumns)
'                      With recColEdit
'                        .Index = "idxColumnID"
'                        .Seek "=", CLng(alngTempColumns(iCount))
'
'                        If Not .NoMatch Then
'                          If .Fields("changed").Value Then
'                            fChanged = True
'                            Exit For
'                          End If
'                        End If
'                      End With
'                    Next iCount
'
'                    If fChanged Then
'                      fOK = ViewSave
'                    End If
'                  End If
'                  Set objFilter = Nothing
'              End If
'            End If
'          End If
'        End If
'      End If
'
'      .MoveNext
'    Loop
'
'
'  End With
'
'TidyUpAndExit:
'  SaveViews = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error saving views"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'
'Private Function SaveModuleDefinitions() As Boolean
'  ' Save the module definitions.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim rsModules As New ADODB.Recordset
'  Dim rsRelatedColumns As New ADODB.Recordset
'  Dim rsLinks As dao.Recordset
'  Dim rsAccord As dao.Recordset
'  Dim sSQL As String
'  Dim alngLinkIDs() As Long
'  Dim rsMaxLinkID As New ADODB.Recordset
'  Dim iLoop As Integer
'  Dim lngNewLinkID As Long
'
'  fOK = True
'
'  ' Default some Workflow setup parameters
'  DefaultWorkflowSetup
'
'  ' Delete any existing Module definitions.
'  gADOCon.Execute "DELETE FROM ASRSysModuleSetup", , adCmdText + adExecuteNoRecords
'
'  ' Open the Module Setup table.
'  rsModules.Open "ASRSysModuleSetup", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
'
'  With recModuleSetup
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'
'    Do While Not .EOF
'      rsModules.AddNew
'      rsModules!moduleKey = !moduleKey
'      rsModules!parameterkey = !parameterkey
'
'      If Not IsNull(!parametervalue) Then
'        rsModules!parametervalue = !parametervalue
'      End If
'
'      rsModules!ParameterType = !ParameterType
'      rsModules.Update
'
'      .MoveNext
'    Loop
'  End With
'
'  rsModules.Close
'
'  ' Delete any existing Related Column definitions.
'  gADOCon.Execute "DELETE FROM ASRSysModuleRelatedColumns", , adCmdText + adExecuteNoRecords
'
'  ' Open the Module Related Column table.
'  rsRelatedColumns.Open "ASRSysModuleRelatedColumns", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
'
'  With recModuleRelatedColumns
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'
'    Do While Not .EOF
'      rsRelatedColumns.AddNew
'      rsRelatedColumns!moduleKey = !moduleKey
'      rsRelatedColumns!parameterkey = !parameterkey
'      rsRelatedColumns!sourcecolumnid = !sourcecolumnid
'      rsRelatedColumns!destcolumnid = !destcolumnid
'      rsRelatedColumns.Update
'
'      .MoveNext
'    Loop
'  End With
'
'  rsRelatedColumns.Close
'
'  ' Delete any existing Self-service Intranet Link definitions.
'  gADOCon.Execute "DELETE FROM ASRSysSSIntranetLinks", , adCmdText + adExecuteNoRecords
'
'  ReDim alngLinkIDs(2, 0)
'
'  sSQL = "SELECT *" & _
'    " FROM tmpSSIntranetLinks"
'  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  While Not rsLinks.EOF
'    'JPD 20040630 Fault 8859
'    'sSQL = "INSERT INTO ASRSysSSIntranetLinks" & _
'      " (linkType, linkOrder, prompt, text, screenID, pageTitle, URL, startMode, utilityType, utilityID, viewID)" & _
'      " VALUES(" & _
'      CStr(rsLinks!LinkType) & "," & _
'      CStr(rsLinks!linkOrder) & "," & _
'      "'" & Replace(rsLinks!Prompt, "'", "''") & "'," & _
'      "'" & Replace(rsLinks!Text, "'", "''") & "'," & _
'      CStr(rsLinks!ScreenID) & "," & _
'      "'" & Replace(rsLinks!PageTitle, "'", "''") & "'," & _
'      "'" & Replace(IIf(IsNull(rsLinks!URL), "", rsLinks!URL), "'", "''") & "'," & _
'      CStr(rsLinks!StartMode) & "," & _
'      CStr(IIf(IsNull(rsLinks!UtilityType), 0, rsLinks!UtilityType)) & "," & _
'      CStr(IIf(IsNull(rsLinks!UtilityID), 0, rsLinks!UtilityID)) & "," & _
'      CStr(IIf(IsNull(rsLinks!ViewID), 0, rsLinks!ViewID)) & _
'      ")"
'
'    'NHRD31012007 Open in New Window Development ammendment.
'    'Added the newWindow variable
'    sSQL = "INSERT INTO ASRSysSSIntranetLinks" & _
'      " (linkType, linkOrder, prompt, text, screenID, pageTitle, URL, startMode, utilityType, utilityID, viewID, newWindow, tableID)" & _
'      " VALUES(" & _
'      CStr(rsLinks!LinkType) & "," & _
'      CStr(rsLinks!linkOrder) & "," & _
'      "'" & Replace(rsLinks!Prompt, "'", "''") & "'," & _
'      "'" & Replace(rsLinks!Text, "'", "''") & "'," & _
'      CStr(rsLinks!ScreenID) & "," & _
'      "'" & Replace(rsLinks!PageTitle, "'", "''") & "'," & _
'      "'" & Replace(IIf(IsNull(rsLinks!URL), "", rsLinks!URL), "'", "''") & "'," & _
'      CStr(rsLinks!StartMode) & "," & _
'      CStr(IIf(IsNull(rsLinks!UtilityType), 0, rsLinks!UtilityType)) & "," & _
'      CStr(IIf(IsNull(rsLinks!UtilityID), 0, rsLinks!UtilityID)) & "," & _
'      CStr(IIf(IsNull(rsLinks!ViewID), 0, rsLinks!ViewID)) & "," & _
'      IIf(IsNull(rsLinks!NewWindow), "0", IIf(rsLinks!NewWindow, "1", "0")) & "," & _
'      CStr(IIf(IsNull(rsLinks!TableID), 0, rsLinks!TableID)) & _
'      ")"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    sSQL = "SELECT MAX(id) AS [result]" & _
'      " FROM ASRSysSSIntranetLinks"
'    rsMaxLinkID.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    ReDim Preserve alngLinkIDs(2, UBound(alngLinkIDs, 2) + 1)
'    alngLinkIDs(1, UBound(alngLinkIDs, 2)) = rsLinks!ID
'    alngLinkIDs(2, UBound(alngLinkIDs, 2)) = rsMaxLinkID!result
'    rsMaxLinkID.Close
'
'    rsLinks.MoveNext
'  Wend
'  rsLinks.Close
'
'  ' Delete any existing Self-service Intranet Link Hidden Group records.
'  gADOCon.Execute "DELETE FROM ASRSysSSIHiddenGroups", , adCmdText + adExecuteNoRecords
'
'  sSQL = "SELECT *" & _
'    " FROM tmpSSIHiddenGroups"
'  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  While Not rsLinks.EOF
'    lngNewLinkID = 0
'    For iLoop = 1 To UBound(alngLinkIDs, 2)
'      If alngLinkIDs(1, iLoop) = rsLinks!LinkID Then
'        lngNewLinkID = alngLinkIDs(2, iLoop)
'        Exit For
'      End If
'    Next iLoop
'
'    sSQL = "INSERT INTO ASRSysSSIHiddenGroups" & _
'      " (linkID, groupName)" & _
'      " VALUES(" & _
'      CStr(lngNewLinkID) & "," & _
'      "'" & Replace(rsLinks!GroupName, "'", "''") & "'" & _
'      ")"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    rsLinks.MoveNext
'  Wend
'  rsLinks.Close
'
'  ' Delete any existing Self-service Intranet Views records.
'  gADOCon.Execute "DELETE FROM ASRSysSSIViews", , adCmdText + adExecuteNoRecords
'
'  sSQL = "SELECT *" & _
'    " FROM tmpSSIViews"
'  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'  While Not rsLinks.EOF
'
'    sSQL = "INSERT INTO ASRSysSSIViews" & _
'      " (viewID, " & _
'      "    buttonLinkPromptText, " & _
'      "    buttonLinkButtonText, " & _
'      "    hypertextLinkText, " & _
'      "    dropdownListLinkText, " & _
'      "    buttonLink, " & _
'      "    hypertextLink, " & _
'      "    dropdownListLink, " & _
'      "    singleRecordView, " & _
'      "    sequence, " & _
'      "    linksLinkText, " & _
'      "    pageTitle, " & _
'      "    tableID)"
'    sSQL = sSQL & _
'      " VALUES(" & _
'      CStr(rsLinks!ViewID) & "," & _
'      "'" & Replace(rsLinks!ButtonLinkPromptText, "'", "''") & "'," & _
'      "'" & Replace(rsLinks!ButtonLinkButtonText, "'", "''") & "'," & _
'      "'" & Replace(rsLinks!HypertextLinkText, "'", "''") & "'," & _
'      "'" & Replace(rsLinks!DropdownListLinkText, "'", "''") & "'," & _
'      IIf(rsLinks!ButtonLink, "1", "0") & "," & _
'      IIf(rsLinks!HypertextLink, "1", "0") & "," & _
'      IIf(rsLinks!DropdownListLink, "1", "0") & "," & _
'      IIf(rsLinks!SingleRecordView, "1", "0") & "," & _
'      CStr(rsLinks!Sequence) & "," & _
'      "'" & Replace(rsLinks!LinksLinkText, "'", "''") & "'," & _
'      "'" & IIf(IsNull(rsLinks!PageTitle), vbNullString, Replace(IIf(IsNull(rsLinks!PageTitle), vbNullString, rsLinks!PageTitle), "'", "''")) & "'," & _
'      CStr(rsLinks!TableID) & _
'      ")"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    rsLinks.MoveNext
'  Wend
'  rsLinks.Close
'
'  ' Store the Payroll Transfer Types
'  gADOCon.Execute "DELETE FROM ASRSysAccordTransferTypes", , adCmdText + adExecuteNoRecords
'
'  sSQL = "SELECT * FROM tmpAccordTransferTypes"
'  Set rsAccord = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  While Not rsAccord.EOF
'
'    sSQL = "INSERT INTO ASRSysAccordTransferTypes" & _
'      " (TransferTypeID, TransferType, FilterID, ASRBaseTableID, IsVisible, ForceAsUpdate)" & _
'      " VALUES (" & _
'      CStr(rsAccord!TransferTypeID) & "," & _
'      "'" & CStr(rsAccord!TransferType) & "'," & _
'      CStr(rsAccord!FilterID) & "," & _
'      CStr(rsAccord!ASRBaseTableID) & "," & _
'      IIf(rsAccord!IsVisible, "1", "0") & "," & _
'      IIf(rsAccord!ForceAsUpdate, "1", "0") & ")"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    rsAccord.MoveNext
'  Wend
'  rsAccord.Close
'
'  ' Store the Payroll mappings
'  gADOCon.Execute "DELETE FROM ASRSysAccordTransferFieldDefinitions", , adCmdText + adExecuteNoRecords
'
'  sSQL = "SELECT * FROM tmpAccordTransferFieldDefinitions"
'  Set rsAccord = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  While Not rsAccord.EOF
'
'    sSQL = "INSERT INTO ASRSysAccordTransferFieldDefinitions" & _
'      " (TransferFieldID, TransferTypeID, Mandatory, Description, ASRMapType, ASRTableID, ASRColumnID, ASRExprID, ASRValue, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ConvertData" & _
'      " , IsEmployeeName, IsDepartmentCode, IsDepartmentName, IsPayrollCode, GroupBy, PreventModify) " & _
'      " VALUES (" & _
'      CStr(rsAccord!TransferFieldID) & "," & _
'      CStr(rsAccord!TransferTypeID) & "," & _
'      IIf(rsAccord!Mandatory, "1", "0") & "," & _
'      "'" & Replace(rsAccord!Description, "'", "''") & "'," & _
'      IIf(IsNull(rsAccord!ASRMapType), "null", rsAccord!ASRMapType) & "," & _
'      IIf(IsNull(rsAccord!ASRTableID), "null", rsAccord!ASRTableID) & "," & _
'      IIf(IsNull(rsAccord!ASRColumnID), "null", rsAccord!ASRColumnID) & "," & _
'      IIf(IsNull(rsAccord!ASRExprID), "null", rsAccord!ASRExprID) & "," & _
'      "'" & Replace(IIf(IsNull(rsAccord!ASRValue), vbNullString, rsAccord!ASRValue), "'", "''") & "'," & _
'      IIf(rsAccord!IsCompanyCode, "1", "0") & "," & _
'      IIf(rsAccord!IsEmployeeCode, "1", "0") & "," & _
'      CStr(rsAccord!Direction) & "," & _
'      IIf(rsAccord!IsKeyField, "1", "0") & "," & _
'      IIf(rsAccord!AlwaysTransfer, "1", "0") & "," & _
'      IIf(rsAccord!ConvertData, "1", "0") & "," & _
'      IIf(rsAccord!IsEmployeeName, "1", "0") & "," & _
'      IIf(rsAccord!IsDepartmentCode, "1", "0") & "," & _
'      IIf(rsAccord!IsDepartmentName, "1", "0") & "," & _
'      IIf(rsAccord!IsPayrollCode, "1", "0") & "," & _
'      IIf(IsNull(rsAccord!GroupBy), "null", rsAccord!GroupBy) & ", " & _
'      IIf(rsAccord!PreventModify, "1", "0") & ")"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    rsAccord.MoveNext
'  Wend
'  rsAccord.Close
'
'  ' Store the Payroll Column Value Mappings
'  gADOCon.Execute "DELETE FROM ASRSysAccordTransferFieldMappings", , adCmdText + adExecuteNoRecords
'
'  sSQL = "SELECT * FROM tmpAccordTransferFieldMappings"
'  Set rsAccord = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  While Not rsAccord.EOF
'
'    sSQL = "INSERT INTO ASRSysAccordTransferFieldMappings" & _
'      " (TransferID, FieldID, HRProValue, AccordValue)" & _
'      " VALUES (" & _
'      CStr(rsAccord!TransferID) & "," & _
'      CStr(rsAccord!FieldID) & "," & _
'      "'" & rsAccord!HRProValue & "'," & _
'      "'" & rsAccord!AccordValue & "')"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    rsAccord.MoveNext
'  Wend
'  rsAccord.Close
'
'TidyUpAndExit:
'  Set rsAccord = Nothing
'  Set rsMaxLinkID = Nothing
'  Set rsModules = Nothing
'  Set rsRelatedColumns = Nothing
'  SaveModuleDefinitions = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error saving module setup"
'  Resume TidyUpAndExit
'
'End Function

''''Private Function ColumnNew() As Boolean
''''  ' Add the current column in the recColEdit recordset to the server databases.
''''  On Error GoTo ErrorTrap
''''
''''  Dim fOK As Boolean
''''  Dim sSQL As String
''''  Dim sColCreate As String
''''  Dim sName As String
''''  Dim iColumn As Integer
''''  Dim rsColumns As New ADODB.Recordset
''''  Dim rsDiaryLinks As New ADODB.Recordset
''''  Dim rsControlValues As New ADODB.Recordset
''''
''''  ' Open the server's column details table.
''''  rsColumns.Open "SELECT * FROM ASRSysColumns", gADOCon, adOpenDynamic, adLockOptimistic
''''
''''  ' Open the server's column control values table.
''''  rsControlValues.Open "SELECT * FROM ASRSysColumnControlValues", gADOCon, adOpenDynamic, adLockOptimistic
''''
''''  ' Open the server's Diary Links table.
''''  rsDiaryLinks.Open "SELECT * FROM ASRSysDiaryLinks", gADOCon, adOpenDynamic, adLockOptimistic
''''
''''  ' Add the column details to the server's ASRSysColumns table.
''''  With rsColumns
''''    .AddNew
''''
''''    For iColumn = 0 To .Fields.Count - 1
''''      sName = .Fields(iColumn).Name
''''      If Not IsNull(recColEdit.Fields(sName)) Then
''''        .Fields(iColumn) = recColEdit.Fields(sName)
''''      End If
''''    Next iColumn
''''
''''    .Update
''''  End With
''''
''''  ' Add the columns control values.
''''  With recContValEdit
''''    If Not (.BOF And .EOF) Then
''''      .MoveFirst
''''      Do While Not .EOF
''''        If !ColumnID = recColEdit!ColumnID Then
''''          rsControlValues.AddNew
''''          rsControlValues!ColumnID = !ColumnID
''''          rsControlValues!Value = !Value
''''          rsControlValues!Sequence = !Sequence
''''          rsControlValues.Update
''''        End If
''''        .MoveNext
''''      Loop
''''    End If
''''  End With
''''
''''  ' Add the diary link values.
''''  With recDiaryEdit
''''    If Not (.BOF And .EOF) Then
''''      .MoveFirst
''''      Do While Not .EOF
''''        If !ColumnID = recColEdit!ColumnID Then
''''          rsDiaryLinks.AddNew
''''          rsDiaryLinks!diaryID = !diaryID
''''          rsDiaryLinks!ColumnID = !ColumnID
''''          rsDiaryLinks!Comment = !Comment
''''          rsDiaryLinks!Offset = !Offset
''''          rsDiaryLinks!Period = !Period
''''          rsDiaryLinks!Reminder = !Reminder
''''          rsDiaryLinks!FilterID = !FilterID
''''          rsDiaryLinks!EffectiveDate = !EffectiveDate
''''          rsDiaryLinks!CheckLeavingDate = !CheckLeavingDate
''''          rsDiaryLinks.Update
''''        End If
''''        .MoveNext
''''      Loop
''''    End If
''''  End With
''''
''''  ' Create a SQL string to create the new column.
''''  If (recColEdit!DataType = dtVARBINARY) Or (recColEdit!DataType = dtLONGVARBINARY) Then
''''    sColCreate = GetColCreateString(recColEdit!ColumnName, dtVARCHAR, 255, 0)
''''  ElseIf (recColEdit!DataType = dtLONGVARCHAR) Then
''''    sColCreate = GetColCreateString(recColEdit!ColumnName, dtVARCHAR, 14, 0)
''''  Else
''''    sColCreate = GetColCreateString(recColEdit!ColumnName, recColEdit!DataType, recColEdit!Size, recColEdit!Decimals)
''''  End If
''''
''''  fOK = (Len(sColCreate) > 0)
''''
''''  If fOK Then
''''    sSQL = "ALTER TABLE " & recTabEdit!TableName & " ADD " & sColCreate
''''
''''    ' Add the code to define any required default.
''''    If Len(Trim(recColEdit!DefaultValue)) > 0 Then
''''      Select Case recColEdit!DataType
''''        Case dtVARCHAR, dtLONGVARCHAR
''''          sSQL = sSQL & " DEFAULT '" & recColEdit!DefaultValue & "' WITH VALUES"
''''        Case dtINTEGER, dtNUMERIC
''''          sSQL = sSQL & " DEFAULT " & recColEdit!DefaultValue & " WITH VALUES"
''''        Case dtBIT
''''          sSQL = sSQL & " DEFAULT " & IIf(recColEdit!DefaultValue = "TRUE", "1", "0") & " WITH VALUES"
''''      End Select
''''    Else
''''      If recColEdit!DataType = dtBIT Then
''''        sSQL = sSQL & " DEFAULT 0 WITH VALUES"
''''      End If
''''    End If
''''    gADOCon.Execute sSQL, , adExecuteNoRecords
''''
''''    ' Grant all privileges for this column to all groups.
''''    GrantColumnPermission recTabEdit!TableName, recColEdit!ColumnName
''''
''''  End If
''''
''''TidyUpAndExit:
''''  Set rsColumns = Nothing
''''  Set rsControlValues = Nothing
''''  Set rsDiaryLinks = Nothing
''''  ColumnNew = fOK
''''  Exit Function
''''
''''ErrorTrap:
''''  MsgBox "Error creating new column." & vbCr & vbCr & _
''''    Err.Description, vbExclamation + vbOKOnly, App.ProductName
''''  fOK = False
''''  Resume TidyUpAndExit
''''
''''End Function
''''Private Function ColumnChange() As Boolean
''''  ' Aleter the definition of the current column in the recColEdit recordset in the server databases.
''''  On Error GoTo ErrorTrap
''''
''''  Dim fOK As Boolean
''''  Dim fGoodGroup As Boolean
''''  Dim fSelectGranted As Boolean
''''  Dim fUpdateGranted As Boolean
''''  Dim fNameChanged As Boolean
''''  Dim fDataTypeChanged As Boolean
''''  Dim fSizeChanged As Boolean
''''  Dim fDecimalsChanged As Boolean
''''  Dim fDefaultValueChanged As Boolean
''''  Dim sSQL As String
''''  Dim sColCreate As String
''''  Dim sCurrentGroupName As String
''''  Dim sOldName As String
''''  Dim sUserGroupName As String
''''  Dim sUserName As String
''''  Dim sName As String
''''  Dim iLoop As Integer
''''  Dim iColumn As Integer
''''  Dim iOldType As Integer
''''  Dim iNextIndex As Integer
''''  Dim rsColumns As New ADODB.Recordset
''''  Dim rsDiaryLinks As New ADODB.Recordset
''''  Dim rsControlValues As New ADODB.Recordset
''''  Dim rsColumnDef As New ADODB.Recordset
''''  Dim rsPriv As New ADODB.Recordset
''''  Dim rsGroups As New ADODB.Recordset
''''  Dim rsSysRoles As New ADODB.Recordset
''''  Dim rsUserInfo As New ADODB.Recordset
''''  Dim asPrivileges() As String
''''
''''  fOK = True
''''
''''  ' Get the original column definition values.
''''  sSQL = "SELECT dataType, size, decimals, columnname, defaultValue" & _
''''    " FROM ASRSysColumns" & _
''''    " WHERE columnID = " & Trim(Str(recColEdit!ColumnID))
''''  rsColumnDef.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
''''
''''  With rsColumnDef
''''    fOK = Not (.EOF And .BOF)
''''    If fOK Then
''''      fNameChanged = (recColEdit!ColumnName <> !ColumnName)
''''      sOldName = !ColumnName
''''      fDataTypeChanged = (recColEdit!DataType <> !DataType)
''''      iOldType = !DataType
''''      fSizeChanged = (recColEdit!Size <> !Size)
''''      fDecimalsChanged = (recColEdit!Decimals <> !Decimals)
''''      fDefaultValueChanged = (recColEdit!DefaultValue <> !DefaultValue)
''''    End If
''''
''''    .Close
''''  End With
''''
''''  If fOK Then
''''    ' Delete column control values for this column.
''''    sSQL = "DELETE FROM ASRSysColumnControlValues" & _
''''      " WHERE columnID = " & Trim(Str(recColEdit!ColumnID))
''''    gADOCon.Execute sSQL, , adExecuteNoRecords
''''
''''    ' Delete diary links for this column.
''''    sSQL = "DELETE FROM ASRSysDiaryLinks" & _
''''      " WHERE columnID = " & Trim(Str(recColEdit!ColumnID))
''''    gADOCon.Execute sSQL, , adExecuteNoRecords
''''
''''    ' Delete the column definition.
''''    sSQL = "DELETE FROM ASRSysColumns" & _
''''      " WHERE columnID=" & Trim(Str(recColEdit!ColumnID))
''''    gADOCon.Execute sSQL, , adExecuteNoRecords
''''
''''    ' Open the server's column details table.
''''    rsColumns.Open "SELECT * FROM ASRSysColumns", gADOCon, adOpenDynamic, adLockOptimistic
''''    ' Open the server's column control values table.
''''    rsControlValues.Open "SELECT * FROM ASRSysColumnControlValues", gADOCon, adOpenDynamic, adLockOptimistic
''''    ' Open the server's Diary Links table.
''''    rsDiaryLinks.Open "SELECT * FROM ASRSysDiaryLinks", gADOCon, adOpenDynamic, adLockOptimistic
''''
''''    ' Change the column details to the server's ASRSysColumns table.
''''    With rsColumns
''''      .AddNew
''''      For iColumn = 0 To .Fields.Count - 1
''''        sName = .Fields(iColumn).Name
''''        If Not IsNull(recColEdit.Fields(sName)) Then
''''          .Fields(iColumn) = recColEdit.Fields(sName)
''''        End If
''''      Next iColumn
''''      .Update
''''    End With
''''
''''    ' Add the columns control values.
''''    With recContValEdit
''''      If Not (.BOF And .EOF) Then
''''        .MoveFirst
''''        Do While Not .EOF
''''          If !ColumnID = recColEdit!ColumnID Then
''''            rsControlValues.AddNew
''''            rsControlValues!ColumnID = !ColumnID
''''            rsControlValues!Value = !Value
''''            rsControlValues!Sequence = !Sequence
''''            rsControlValues.Update
''''          End If
''''          .MoveNext
''''        Loop
''''      End If
''''    End With
''''
''''    ' Add the diary link values.
''''    With recDiaryEdit
''''      If Not (.BOF And .EOF) Then
''''        .MoveFirst
''''        Do While Not .EOF
''''          If !ColumnID = recColEdit!ColumnID Then
''''            rsDiaryLinks.AddNew
''''            rsDiaryLinks!diaryID = !diaryID
''''            rsDiaryLinks!ColumnID = !ColumnID
''''            rsDiaryLinks!Comment = !Comment
''''            rsDiaryLinks!Offset = !Offset
''''            rsDiaryLinks!Period = !Period
''''            rsDiaryLinks!Reminder = !Reminder
''''            rsDiaryLinks!FilterID = !FilterID
''''            rsDiaryLinks!EffectiveDate = !EffectiveDate
''''            rsDiaryLinks!CheckLeavingDate = !CheckLeavingDate
''''            rsDiaryLinks.Update
''''          End If
''''          .MoveNext
''''        Loop
''''      End If
''''    End With
''''
''''    ' Change the column name if it has changed.
''''    If fNameChanged Then
''''      sSQL = "EXEC sp_rename '" & recTabEdit!TableName & "." & sOldName & "', '" & recColEdit!ColumnName & "', 'COLUMN'"
''''      gADOCon.Execute sSQL, , adExecuteNoRecords
''''    End If
''''
''''    ' Change the data type and size definition if it has changed.
''''    If fDataTypeChanged Or fSizeChanged Or fDecimalsChanged Then
''''      If (recColEdit!DataType = dtLONGVARBINARY) Or _
''''        (iOldType = dtLONGVARBINARY) Then
''''        ' Cannot use ALTER TABLE on OLE columns, so drop and recreate it.
''''        ' NB. We don't even try to copy OLE data.
''''        ' Read the privileges for this column.
''''
''''        ' Clear the array of column privileges.
''''        ReDim asPrivileges(2, 0)
''''
''''        ' Get a list of privileges for the given column.
''''        sSQL = "sp_column_privileges @table_name = '" & recTabEdit!TableName & "', @column_name = '" & recColEdit!ColumnName & "'"
''''        rsPriv.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
''''
''''        With rsPriv
''''        ' Read each privilege record into the array.
''''          While Not .EOF
''''
''''            ' Do not bother with 'dbo' privileges as these are created by default.
''''            ' Only bother with SELECT and UPDATE privileges as the others are
''''            ' handled at table level.
''''            If .Fields("GRANTEE") <> "dbo" And _
''''              ((.Fields("PRIVILEGE") = "SELECT") Or _
''''              (.Fields("PRIVILEGE") = "UPDATE")) Then
''''
''''              ' Add a new row onto the array, and populate it with privilege info.
''''              iNextIndex = UBound(asPrivileges, 2) + 1
''''              ReDim Preserve asPrivileges(2, iNextIndex)
''''
''''              asPrivileges(1, iNextIndex) = .Fields("GRANTEE")
''''              asPrivileges(2, iNextIndex) = .Fields("PRIVILEGE")
''''            End If
''''
''''            .MoveNext
''''          Wend
''''
''''          .Close
''''        End With
''''
''''        ' Drop the column.
''''        sSQL = "EXEC sp_ASRDropColumn " & recTabEdit!TableName & ", " & recColEdit!ColumnName
''''        gADOCon.Execute sSQL, , adExecuteNoRecords
''''        sColCreate = GetColCreateString(recColEdit!ColumnName, recColEdit!DataType, recColEdit!Size, recColEdit!Decimals)
''''        sSQL = "ALTER TABLE " & recTabEdit!TableName & " ADD " & sColCreate
''''        gADOCon.Execute sSQL, , adExecuteNoRecords
''''
''''        ' Set the privileges for this column.
''''        ' Get a list of User Groups (Roles) from SQL Server
''''        If IsVersion7 Then
''''          rsGroups.Open "sp_helprole", gADOCon, adOpenForwardOnly, adLockReadOnly
''''
''''          ' Create an array of the standard system roles for SQL Server 7.0
''''          ' which we want to leave alone.
''''          ReDim asFixedRoles(0)
''''          asFixedRoles(0) = "PUBLIC"
''''          rsSysRoles.Open "sp_helpdbfixedrole", gADOCon, adOpenForwardOnly, adLockReadOnly
''''
''''          With rsSysRoles
''''            Do While Not .EOF
''''              iNextIndex = UBound(asFixedRoles) + 1
''''              ReDim Preserve asFixedRoles(iNextIndex)
''''              asFixedRoles(iNextIndex) = UCase(Trim(.Fields(0).Value))
''''              .MoveNext
''''            Loop
''''
''''            .Close
''''          End With
''''        Else
''''          rsGroups.Open "sp_helpgroup", gADOCon, adOpenForwardOnly, adLockReadOnly
''''        End If
''''
''''        ' For each User Group (Role) ...
''''        With rsGroups
''''
''''          If Not .EOF And Not .BOF Then
''''
''''            While Not .EOF
''''              sCurrentGroupName = UCase(Trim(.Fields(0).Value))
''''              fSelectGranted = False
''''              fUpdateGranted = False
''''
''''              ' Check that the group is valid. ie. not a system User Group (Role).
''''              fGoodGroup = True
''''              If IsVersion7 Then
''''                For iNextIndex = 0 To UBound(asFixedRoles)
''''                  If asFixedRoles(iNextIndex) = sCurrentGroupName Then
''''                    fGoodGroup = False
''''                    Exit For
''''                  End If
''''                Next iNextIndex
''''              End If
''''
''''              If fGoodGroup Then
''''                For iLoop = 1 To UBound(asPrivileges, 2)
''''                  ' Get the primary User Group (Role) of the current user.
''''                  sUserGroupName = vbnullstring
''''                  sUserName = UCase$(Trim(asPrivileges(1, iLoop)))
''''
''''                  If IsVersion7 Then
''''                    sSQL = "SELECT su1.name AS groupName" & _
''''                      " FROM sysusers su1, sysusers su2" & _
''''                      " WHERE su2.name = '" & sUserName & "'" & _
''''                      " AND su2.gid = su1.uid"
''''                    rsUserInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
''''                    With rsUserInfo
''''                      If Not (.EOF And .BOF) Then
''''                        sUserGroupName = UCase(Trim(IIf(.Fields("groupName") = "public", vbnullstring, !GroupName!)))
''''                      End If
''''                      .Close
''''                    End With
''''                  Else
''''                    sSQL = "sp_helpuser '" & sUserName & "'"
''''                    rsUserInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
''''                    With rsUserInfo
''''                      If Not (.EOF And .BOF) Then
''''                        sUserGroupName = IIf(IsNull(.Fields("groupName")), "public", !GroupName)
''''                      End If
''''                      .Close
''''                    End With
''''                  End If
''''
''''                  If sCurrentGroupName = sUserGroupName Then
''''                    ' Mark the column as granted in our array of columns.
''''                    If asPrivileges(2, iLoop) = "SELECT" Then
''''                      fSelectGranted = True
''''                    ElseIf asPrivileges(2, iLoop) = "UPDATE" Then
''''                      fUpdateGranted = True
''''                    End If
''''                  End If
''''                Next iLoop
''''
''''                ' Grant/Deny the SELECT privileges for this User Group (Role) as required.
''''                If fSelectGranted Then
''''                  sSQL = "GRANT SELECT(" & recColEdit!ColumnName & ") ON " & recTabEdit!TableName & " TO [" & sCurrentGroupName & "]"
''''                  gADOCon.Execute sSQL, , adExecuteNoRecords
''''                Else
''''                  sSQL = "DENY SELECT(" & recColEdit!ColumnName & ") ON " & recTabEdit!TableName & " TO [" & sCurrentGroupName & "]"
''''                  gADOCon.Execute sSQL, , adExecuteNoRecords
''''                End If
''''                If fUpdateGranted Then
''''                  sSQL = "GRANT UPDATE(" & recColEdit!ColumnName & ") ON " & recTabEdit!TableName & " TO [" & sCurrentGroupName & "]"
''''                  gADOCon.Execute sSQL, , adExecuteNoRecords
''''                Else
''''                  sSQL = "DENY UPDATE(" & recColEdit!ColumnName & ") ON " & recTabEdit!TableName & " TO [" & sCurrentGroupName & "]"
''''                  gADOCon.Execute sSQL, , adExecuteNoRecords
''''                End If
''''              End If
''''
''''              .MoveNext
''''            Wend
''''          End If
''''
''''          .Close
''''        End With
''''      Else
''''        If (recColEdit!DataType = dtVARBINARY) Or (recColEdit!DataType = dtLONGVARBINARY) Then
''''          sColCreate = GetColCreateString(recColEdit!ColumnName, dtVARCHAR, 255, 0)
''''        ElseIf (recColEdit!DataType = dtLONGVARCHAR) Then
''''          sColCreate = GetColCreateString(recColEdit!ColumnName, dtVARCHAR, 14, 0)
''''        Else
''''          sColCreate = GetColCreateString(recColEdit!ColumnName, recColEdit!DataType, recColEdit!Size, recColEdit!Decimals)
''''        End If
''''
''''        fOK = (Len(sColCreate) > 0)
''''
''''        If fOK Then
''''          sSQL = "ALTER TABLE " & recTabEdit!TableName & " ALTER COLUMN " & sColCreate
''''          gADOCon.Execute sSQL, , adExecuteNoRecords
''''        End If
''''      End If
''''    End If
''''
''''    If fOK Then
''''      If fDefaultValueChanged Then
''''        ' Clear any existing default.
''''        sSQL = "EXEC sp_ASRDropColumnDefault '" & recColEdit!ColumnName & "', '" & recTabEdit!TableName & "'"
''''        gADOCon.Execute sSQL, , adExecuteNoRecords
''''        ' Set the new default.
''''        Select Case recColEdit!DataType
''''          Case dtVARCHAR, dtLONGVARCHAR
''''            sSQL = "ALTER TABLE " & recTabEdit!TableName & " WITH NOCHECK ADD DEFAULT '" & _
''''              recColEdit!DefaultValue & "' FOR " & recColEdit!ColumnName
''''            gADOCon.Execute sSQL, , adExecuteNoRecords
''''          Case dtINTEGER, dtNUMERIC
''''            sSQL = "ALTER TABLE " & recTabEdit!TableName & " WITH NOCHECK ADD DEFAULT " & _
''''              recColEdit!DefaultValue & " FOR " & recColEdit!ColumnName
''''            gADOCon.Execute sSQL, , adExecuteNoRecords
''''          Case dtBIT
''''            sSQL = "ALTER TABLE " & recTabEdit!TableName & " WITH NOCHECK ADD DEFAULT " & _
''''              IIf(recColEdit!DefaultValue = "TRUE", "1", "0") & " FOR " & recColEdit!ColumnName
''''            gADOCon.Execute sSQL, , adExecuteNoRecords
''''        End Select
''''      ElseIf (recColEdit!DataType = dtBIT) Then
''''        sSQL = "ALTER TABLE " & recTabEdit!TableName & " WITH NOCHECK ADD DEFAULT " & _
''''          IIf(recColEdit!DefaultValue = "TRUE", "1", "0") & " FOR " & recColEdit!ColumnName
''''        gADOCon.Execute sSQL, , adExecuteNoRecords
''''      End If
''''    End If
''''  End If
''''
''''TidyUpAndExit:
''''  Set rsColumns = Nothing
''''  Set rsDiaryLinks = Nothing
''''  Set rsControlValues = Nothing
''''  Set rsColumnDef = Nothing
''''  Set rsPriv = Nothing
''''  Set rsGroups = Nothing
''''  Set rsSysRoles = Nothing
''''  Set rsUserInfo = Nothing
''''
''''  ColumnChange = fOK
''''  Exit Function
''''
''''ErrorTrap:
''''  MsgBox "Error changing column definition." & vbCr & vbCr & _
''''    Err.Description, vbExclamation + vbOKOnly, App.ProductName
''''  fOK = False
''''  Resume TidyUpAndExit
''''
''''End Function
''''
''''
''''Private Sub GrantColumnPermission(psTableName As String, psColumnName As String, Optional pvPermission As Variant, Optional pvRole As Variant)
''''
''''  ' Grant the given permission to the given groups.
''''  Dim fGoodGroup As Boolean
''''  Dim iNextIndex As Integer
''''  Dim sSQL As String
''''  Dim sCurrentGroupName As String
''''  Dim rsGroups As New ADODB.Recordset
''''  Dim rsSysRoles As New ADODB.Recordset
''''  Dim asFixedRoles() As String
''''
''''  ' Get a list of User Groups (Roles) from SQL Server
''''  If IsVersion7 Then
''''    rsGroups.Open "sp_helprole", gADOCon, adOpenForwardOnly, adLockReadOnly
''''
''''    ' Create an array of the standard system roles for SQL Server 7.0
''''    ' which we want to leave alone.
''''    ReDim asFixedRoles(0)
''''    asFixedRoles(0) = "PUBLIC"
''''    rsSysRoles.Open "sp_helpdbfixedrole", gADOCon, adOpenForwardOnly, adLockReadOnly
''''
''''    With rsSysRoles
''''      Do While Not .EOF
''''        iNextIndex = UBound(asFixedRoles) + 1
''''        ReDim Preserve asFixedRoles(iNextIndex)
''''        asFixedRoles(iNextIndex) = UCase(Trim(.Fields(0).Value))
''''        .MoveNext
''''      Loop
''''
''''      .Close
''''    End With
''''  Else
''''    rsGroups.Open "sp_helpgroup", gADOCon, adOpenForwardOnly, adLockReadOnly
''''  End If
''''
''''  ' For each User Group (Role) ...
''''  With rsGroups
''''
''''    If Not .EOF And Not .BOF Then
''''
''''      While Not .EOF
''''        sCurrentGroupName = UCase(Trim(.Fields(0).Value))
''''
''''        ' Check that the group is valid. ie. not a system User Group (Role).
''''        If IsMissing(pvRole) Then
''''          fGoodGroup = True
''''          If IsVersion7 Then
''''            For iNextIndex = 0 To UBound(asFixedRoles)
''''              If asFixedRoles(iNextIndex) = sCurrentGroupName Then
''''                fGoodGroup = False
''''                Exit For
''''              End If
''''            Next iNextIndex
''''          End If
''''        Else
''''          fGoodGroup = (sCurrentGroupName = UCase(pvRole))
''''        End If
''''
''''        If fGoodGroup Then
''''          ' Grant the given permission(s).
''''          If IsMissing(pvPermission) Then
''''            sSQL = "GRANT SELECT(" & psColumnName & ") ON " & psTableName & " TO [" & sCurrentGroupName & "]"
''''            gADOCon.Execute sSQL, , adExecuteNoRecords
''''            sSQL = "GRANT UPDATE(" & psColumnName & ") ON " & psTableName & " TO [" & sCurrentGroupName & "]"
''''            gADOCon.Execute sSQL, , adExecuteNoRecords
''''          ElseIf pvPermission = "SELECT" Then
''''            sSQL = "GRANT SELECT(" & psColumnName & ") ON " & psTableName & " TO [" & sCurrentGroupName & "]"
''''            gADOCon.Execute sSQL, , adExecuteNoRecords
''''          Else
''''            sSQL = "GRANT UPDATE(" & psColumnName & ") ON " & psTableName & " TO [" & sCurrentGroupName & "]"
''''            gADOCon.Execute sSQL, , adExecuteNoRecords
''''          End If
''''        End If
''''
''''        .MoveNext
''''      Wend
''''    End If
''''
''''    .Close
''''  End With
''''
''''  Set rsGroups = Nothing
''''  Set rsSysRoles = Nothing
''''
''''End Sub
'
'Private Function ColumnDelete() As Boolean
'  ' Delete the current column in the recColEdit recordset from the server databases.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  ' Drop the deleted column's default if it exists.
'  sSQL = "EXEC sp_ASRDropColumnDefault '" & recColEdit!ColumnName & "', '" & recTabEdit!TableName & "'"
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Drop the deleted column.
'  sSQL = "EXEC sp_ASRDropColumn " & recTabEdit!TableName & ", " & recColEdit!ColumnName
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Delete column control values for this column.
'  sSQL = "DELETE FROM ASRSysColumnControlValues" & _
'    " WHERE columnID = " & Trim(Str(recColEdit!ColumnID))
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Delete diary links for this column.
'  sSQL = "DELETE FROM ASRSysDiaryLinks" & _
'    " WHERE columnID = " & Trim(Str(recColEdit!ColumnID))
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Delete the column definition.
'  sSQL = "DELETE FROM ASRSysColumns" & _
'    " WHERE columnID=" & Trim(Str(recColEdit!ColumnID))
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'TidyUpAndExit:
'  ColumnDelete = fOK
'  Exit Function
'
'ErrorTrap:
'  MsgBox "Error deleting column." & vbCr & vbCr & _
'    Err.Description, vbExclamation + vbOKOnly, App.ProductName
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function


Private Function RefreshChildView(plngChildViewID As Long, _
  psChildTableName As String, _
  pfColumnsDeleted As Boolean, _
  pfColumnsAdded As Boolean, _
  piType As Integer) As Boolean
  ' Refresh the given child view.
  ' Return TRUE if everything went okay.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'  Dim sParentJoinCode As String
'  Dim rsParentsInfo As rdoResultset
'  Dim lngLastParentID As Long
'
'  fOK = True
'
'  If fOK Then
'    If pfColumnsDeleted Then
'      ' Get the child view's parents.
'      sParentJoinCode = vbnullstring
'
'      sSQL = "SELECT ASRSysChildViewParents.parentTableID," & _
'        "    CASE" & _
'        "        WHEN parentType = 'UV' THEN ASRSysViews.viewName" & _
'        "        WHEN parentType = 'SV' THEN 'ASRSysChildView_' + convert(varchar(100), ASRSysChildViewParents.parentID)" & _
'        "        ELSE ASRSysTables.tableName" & _
'        "    END AS parentName " & _
'        " FROM ASRSysChildViewParents" & _
'        " LEFT OUTER JOIN ASRSysViews ON ASRSysChildViewParents.parentID = ASRSysViews.viewID" & _
'        " LEFT OUTER JOIN ASRSysTables ON ASRSysChildViewParents.parentID = ASRSysTables.tableID" & _
'        " WHERE childViewID = " & trim(str(plngChildViewID))
'
'      Set rsParentsInfo = rdoCon.OpenResultset(sSQL, _
'        rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'      fOK = Not (rsParentsInfo.EOF And rsParentsInfo.BOF)
'
'      If fOK Then
'        lngLastParentID = 0
'        Do While Not rsParentsInfo.EOF
'          If Len(sParentJoinCode) = 0 Then
'            sParentJoinCode = "        WHERE" & vbNewLine & _
'            "                (" & vbNewLine
'          Else
'            If piType = 1 Then
'              If lngLastParentID <> rsParentsInfo!parentTableID Then
'                lngLastParentID = rsParentsInfo!parentTableID
'
'                sParentJoinCode = sParentJoinCode & _
'                  "                )" & vbNewLine & _
'                  "                AND" & vbNewLine & _
'                  "                (" & vbNewLine
'              Else
'                sParentJoinCode = sParentJoinCode & _
'                  "                        OR" & vbNewLine
'              End If
'            Else
'              sParentJoinCode = sParentJoinCode & _
'                "                        OR" & vbNewLine
'            End If
'          End If
'
'          sParentJoinCode = sParentJoinCode & "                        " & _
'            psChildTableName & ".ID_" & trim(str(rsParentsInfo!parentTableID)) & " IN (SELECT ID FROM " & rsParentsInfo!parentName & ")" & vbNewLine
'
'          rsParentsInfo.MoveNext
'        Loop
'
'        sParentJoinCode = sParentJoinCode & "                )"
'      End If
'
'      rsParentsInfo.Close
'      Set rsParentsInfo = Nothing
'
'      If fOK Then
'        ' Refresh the view.
'        sSQL = "ALTER VIEW dbo." & "ASRSysChildView_" & trim(str(plngChildViewID)) & vbNewLine & _
'          "AS" & vbNewLine & _
'          "        SELECT " & psChildTableName & ".*" & vbNewLine & _
'          "        FROM " & psChildTableName & vbNewLine & _
'          sParentJoinCode
'        rdoCon.Execute sSQL, rdExecDirect
'      End If
'    ElseIf pfColumnsAdded Then
'      sSQL = "exec sp_refreshview 'dbo." & "ASRSysChildView_" & trim(str(plngChildViewID)) & "'"
'      rdoCon.Execute sSQL, rdExecDirect
'    End If
'  End If
'
'TidyUpAndExit:
'  RefreshChildView = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  'MsgBox ODBC.FormatError(Err.Description), _
'    vbOKOnly + vbExclamation, Application.Name
'  OutputError "Error refreshing child view"
'  Resume TidyUpAndExit
'
End Function

'Private Function LongestRouteToTopLevel(plngTableID As Long) As Integer
'  ' Return the given table's longest route to the top-level.
'  ' This is used when creating child views.
'  Dim iLongestRoute As Integer
'  Dim iParentsLongestRoute As Integer
'  Dim sSQL As String
'  Dim rsParents As dao.Recordset
'
'  iLongestRoute = 0
'
'  sSQL = "SELECT parentID" & _
'    " FROM tmpRelations" & _
'    " WHERE childID = " & Trim$(Str$(plngTableID))
'  Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'  With rsParents
'    Do While (Not .EOF)
'      iParentsLongestRoute = LongestRouteToTopLevel(.Fields(0).Value)
'
'      If (iParentsLongestRoute + 1) > iLongestRoute Then
'        iLongestRoute = (iParentsLongestRoute + 1)
'      End If
'
'      .MoveNext
'    Loop
'
'    .Close
'  End With
'  Set rsParents = Nothing
'
'  LongestRouteToTopLevel = iLongestRoute
'
'End Function


'Public Function HasExpressionComponent(plngExprIDBeingSearched As Long, plngExprIDSearchedFor As Long) As Boolean
'  'JPD 20040504 Fault 8599
'  On Error GoTo ErrorTrap
'
'  Dim rsExprComp As dao.Recordset
'  Dim rsExpr As dao.Recordset
'  Dim fHasExpr As Boolean
'  Dim sSQL As String
'  Dim lngSubExprID As Long
'
'  HasExpressionComponent = (plngExprIDBeingSearched = plngExprIDSearchedFor)
'
'  If Not HasExpressionComponent Then
'    sSQL = "SELECT * FROM tmpComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
'    Set rsExprComp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'    With rsExprComp
'      Do Until .EOF
'        Select Case !Type
'          Case giCOMPONENT_CALCULATION
'            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)
'
'            If lngSubExprID > 0 Then
'              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
'            End If
'
'          Case giCOMPONENT_FILTER
'            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)
'
'            If lngSubExprID > 0 Then
'              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
'            End If
'
'          Case giCOMPONENT_FIELD
'            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
'
'            If lngSubExprID > 0 Then
'              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
'            End If
'
'          Case giCOMPONENT_FUNCTION
'            sSQL = "SELECT exprID FROM tmpExpressions WHERE parentComponentID = " & CStr(!ComponentID)
'            Set rsExpr = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'            Do Until rsExpr.EOF
'              HasExpressionComponent = HasExpressionComponent(rsExpr!ExprID, plngExprIDSearchedFor)
'
'              If HasExpressionComponent Then
'                Exit Do
'              End If
'
'              rsExpr.MoveNext
'            Loop
'            rsExpr.Close
'            Set rsExpr = Nothing
'
'          Case giCOMPONENT_WORKFLOWFIELD
'            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
'
'            If lngSubExprID > 0 Then
'              HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
'            End If
'
'        End Select
'
'        If HasExpressionComponent Then
'          Exit Do
'        End If
'
'        .MoveNext
'      Loop
'    End With
'
'    rsExprComp.Close
'  End If
'
'TidyUpAndExit:
'  Set rsExprComp = Nothing
'
'  Exit Function
'
'ErrorTrap:
'  Resume TidyUpAndExit
'
'End Function
'
'Public Function ExpressionUsesColumn(plngExprIDBeingSearched As Long, plngColumnIDSearchedFor As Long) As Boolean
'
'  'NB. This reads data from the sql db. i.e. not the access db.
'
'  On Error GoTo ErrorTrap
'
'  Dim rsExprComp As New ADODB.Recordset
'  Dim rsExpr As New ADODB.Recordset
'  Dim fHasExpr As Boolean
'  Dim sSQL As String
'  Dim lngSubExprID As Long
'  Dim lngColumnID As Long
'
'  ExpressionUsesColumn = False
'
'  If Not ExpressionUsesColumn Then
'    sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
'    rsExprComp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'    With rsExprComp
'      Do Until .EOF
'        lngColumnID = 0
'        lngSubExprID = 0
'
'        Select Case !Type
'          Case giCOMPONENT_CALCULATION
'            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesColumn = ExpressionUsesColumn(lngSubExprID, plngColumnIDSearchedFor)
'            End If
'
'          Case giCOMPONENT_FILTER
'            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesColumn = ExpressionUsesColumn(lngSubExprID, plngColumnIDSearchedFor)
'            End If
'
'          Case giCOMPONENT_FIELD
'            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
'            lngColumnID = IIf(IsNull(!fieldColumnID), 0, !fieldColumnID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesColumn = ExpressionUsesColumn(lngSubExprID, plngColumnIDSearchedFor)
'            ElseIf lngColumnID > 0 Then
'              ExpressionUsesColumn = (lngColumnID = plngColumnIDSearchedFor)
'            End If
'
'          Case giCOMPONENT_TABLEVALUE
'            lngColumnID = IIf(IsNull(!LookupColumnID), 0, !LookupColumnID)
'
'            If lngColumnID > 0 Then
'              ExpressionUsesColumn = (lngColumnID = plngColumnIDSearchedFor)
'            End If
'
'          Case giCOMPONENT_FUNCTION
'            sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(!ComponentID)
'            rsExpr.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'            Do Until rsExpr.EOF
'              ExpressionUsesColumn = ExpressionUsesColumn(rsExpr!ExprID, plngColumnIDSearchedFor)
'
'              If ExpressionUsesColumn Then
'                Exit Do
'              End If
'
'              rsExpr.MoveNext
'            Loop
'            rsExpr.Close
'
'        End Select
'
'        If ExpressionUsesColumn Then
'          Exit Do
'        End If
'
'        .MoveNext
'      Loop
'    End With
'
'    rsExprComp.Close
'  End If
'
'TidyUpAndExit:
'  Set rsExpr = Nothing
'  Set rsExprComp = Nothing
'
'  Exit Function
'
'ErrorTrap:
'  Resume TidyUpAndExit
'
'End Function
'
'Public Function ExpressionUsesRelationship_SQL(plngExprIDBeingSearched As Long, _
'                                                  plngParentTableIDSearcherFor As Long, _
'                                                  plngChildTableIDSearcherFor As Long) As Boolean
'
'  'NB. This reads data from the sql db. i.e. not the access db.
'
'  On Error GoTo ErrorTrap
'
'  Dim rsExprComp As New ADODB.Recordset
'  Dim rsExpr As New ADODB.Recordset
'  Dim fHasExpr As Boolean
'  Dim sSQL As String
'  Dim lngSubExprID As Long
'  Dim lngFieldTableID As Long
'  Dim lngExprBaseTableID As Long
'
'  ExpressionUsesRelationship_SQL = False
'
'  If Not ExpressionUsesRelationship_SQL Then
'    sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
'    rsExprComp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    With rsExprComp
'      Do Until .EOF
'        lngExprBaseTableID = IIf(IsNull(!TableID), 0, !TableID)
'        lngSubExprID = 0
'
'        Select Case !Type
'          Case giCOMPONENT_CALCULATION
'            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesRelationship_SQL = ExpressionUsesRelationship_SQL(lngSubExprID, _
'                                                                              plngParentTableIDSearcherFor, _
'                                                                              plngChildTableIDSearcherFor)
'            End If
'
'          Case giCOMPONENT_FILTER
'            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesRelationship_SQL = ExpressionUsesRelationship_SQL(lngSubExprID, _
'                                                                              plngParentTableIDSearcherFor, _
'                                                                              plngChildTableIDSearcherFor)
'            End If
'
'          Case giCOMPONENT_FIELD
'            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
'            lngFieldTableID = IIf(IsNull(!fieldTableID), 0, !fieldTableID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesRelationship_SQL = ExpressionUsesRelationship_SQL(lngSubExprID, _
'                                                                              plngParentTableIDSearcherFor, _
'                                                                              plngChildTableIDSearcherFor)
'            Else
'              ExpressionUsesRelationship_SQL = ((lngExprBaseTableID = plngParentTableIDSearcherFor) _
'                                                  And (lngFieldTableID = plngChildTableIDSearcherFor)) _
'                                                Or ((lngExprBaseTableID = plngChildTableIDSearcherFor) _
'                                                  And (lngFieldTableID = plngParentTableIDSearcherFor))
'            End If
'
'          Case giCOMPONENT_FUNCTION
'            sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(!ComponentID)
'            rsExpr.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'            Do Until rsExpr.EOF
'              ExpressionUsesRelationship_SQL = ExpressionUsesRelationship_SQL(lngSubExprID, _
'                                                                              plngParentTableIDSearcherFor, _
'                                                                              plngChildTableIDSearcherFor)
'
'              If ExpressionUsesRelationship_SQL Then
'                Exit Do
'              End If
'
'              rsExpr.MoveNext
'            Loop
'            rsExpr.Close
'
'        End Select
'
'        If ExpressionUsesRelationship_SQL Then
'          Exit Do
'        End If
'
'        .MoveNext
'      Loop
'    End With
'
'    rsExprComp.Close
'  End If
'
'TidyUpAndExit:
'  Set rsExpr = Nothing
'  Set rsExprComp = Nothing
'
'  Exit Function
'
'ErrorTrap:
'  Resume TidyUpAndExit
'
'End Function
'
'Public Function ExpressionUsesRelationship(plngExprIDBeingSearched As Long, plngParentTableIDSearcherFor As Long, _
'                                            plngChildTableIDSearcherFor As Long, lngExprBaseTableID As Long) As Boolean
'
'  'NB. This reads data from the sql db. i.e. not the access db.
'
'  On Error GoTo ErrorTrap
'
'  Dim rsExprComp As dao.Recordset
'  Dim rsExpr As dao.Recordset
'  Dim fHasExpr As Boolean
'  Dim sSQL As String
'  Dim lngSubExprID As Long
'  Dim lngFieldTableID As Long
'
'  ExpressionUsesRelationship = False
'
'  If Not ExpressionUsesRelationship Then
'    sSQL = "SELECT * FROM tmpComponents WHERE tmpComponents.exprID = " & CStr(plngExprIDBeingSearched)
'    Set rsExprComp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'    With rsExprComp
'      Do Until .EOF
'        lngSubExprID = 0
'
'        Select Case !Type
'          Case giCOMPONENT_CALCULATION
'            lngSubExprID = IIf(IsNull(!CalculationID), 0, !CalculationID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesRelationship = ExpressionUsesRelationship(lngSubExprID, plngParentTableIDSearcherFor, _
'                                                                      plngChildTableIDSearcherFor, lngExprBaseTableID)
'            End If
'
'          Case giCOMPONENT_FILTER
'            lngSubExprID = IIf(IsNull(!FilterID), 0, !FilterID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesRelationship = ExpressionUsesRelationship(lngSubExprID, plngParentTableIDSearcherFor, _
'                                                                      plngChildTableIDSearcherFor, lngExprBaseTableID)
'            End If
'
'          Case giCOMPONENT_FIELD
'            lngSubExprID = IIf(IsNull(!FieldSelectionFilter), 0, !FieldSelectionFilter)
'            lngFieldTableID = IIf(IsNull(!fieldTableID), 0, !fieldTableID)
'
'            If lngSubExprID > 0 Then
'              ExpressionUsesRelationship = ExpressionUsesRelationship(lngSubExprID, plngParentTableIDSearcherFor, _
'                                                                      plngChildTableIDSearcherFor, lngExprBaseTableID)
'            Else
'              ExpressionUsesRelationship = ((lngExprBaseTableID = plngParentTableIDSearcherFor) _
'                                            And (lngFieldTableID = plngChildTableIDSearcherFor)) _
'                                          Or ((lngExprBaseTableID = plngChildTableIDSearcherFor) _
'                                            And (lngFieldTableID = plngParentTableIDSearcherFor))
'            End If
'
'          Case giCOMPONENT_FUNCTION
'            sSQL = "SELECT tmpExpressions.exprID FROM tmpExpressions WHERE tmpExpressions.parentComponentID = " & CStr(!ComponentID)
'            Set rsExpr = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'            Do Until rsExpr.EOF
'              ExpressionUsesRelationship = ExpressionUsesRelationship(rsExpr!ExprID, plngParentTableIDSearcherFor, _
'                                                                      plngChildTableIDSearcherFor, lngExprBaseTableID)
'
'              If ExpressionUsesRelationship Then
'                Exit Do
'              End If
'
'              rsExpr.MoveNext
'            Loop
'            rsExpr.Close
'
'        End Select
'
'        If ExpressionUsesRelationship Then
'          Exit Do
'        End If
'
'        .MoveNext
'      Loop
'    End With
'
'    rsExprComp.Close
'  End If
'
'TidyUpAndExit:
'  Set rsExpr = Nothing
'  Set rsExprComp = Nothing
'
'  Exit Function
'
'ErrorTrap:
'  Resume TidyUpAndExit
'
'End Function
'
'Public Sub CalculatedColumnsThatUseFunction(ByRef pvColumns As Variant, plngFunctionID As Long)
'  On Error GoTo ErrorTrap
'
'  Dim rsCheck As dao.Recordset
'  Dim objComp As CExprComponent
'  Dim sSQL As String
'  Dim lngExprID As Long
'  Dim objCalc As CExpression
'
'  sSQL = "SELECT DISTINCT tmpComponents.componentID" & _
'    " FROM tmpComponents, tmpExpressions " & _
'    " WHERE tmpExpressions.exprid = tmpComponents.Exprid " & _
'    "   AND tmpComponents.functionID = " & Trim$(Str$(plngFunctionID))
'
'  Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'  Do Until rsCheck.EOF
'    Set objComp = New CExprComponent
'    objComp.ComponentID = rsCheck!ComponentID
'    lngExprID = objComp.RootExpressionID
'    Set objComp = Nothing
'
'    Set objCalc = New CExpression
'    With objCalc
'      ' Construct the Filter expression.
'      .ExpressionID = lngExprID
'      .CalculatedColumnsThatUseThisExpression pvColumns
'    End With
'    Set objCalc = Nothing
'
'    rsCheck.MoveNext
'  Loop
'  ' Close the recordset.
'  rsCheck.Close
'
'TidyUpAndExit:
'
'  Exit Sub
'
'ErrorTrap:
'  Resume TidyUpAndExit
'
'End Sub
'
'Public Function TablesThatUseFunction(ByRef pvTables As Variant, plngFunctionID As Long) As Variant
'  ' Return an array of the table IDs that use the 'AbsenceDuration' function in calculated columns.
'  On Error GoTo ErrorTrap
'
'  Dim iCount As Integer
'  Dim iLoop As Integer
'  Dim alngTempColumns() As Long
'  Dim fFound As Boolean
'
'  ' Work out which tables use the AbsenceDuration function
'  ReDim alngTempColumns(0)
'  CalculatedColumnsThatUseFunction alngTempColumns, plngFunctionID
'
'  For iCount = 1 To UBound(alngTempColumns)
'    With recColEdit
'      .Index = "idxColumnID"
'      .Seek "=", CLng(alngTempColumns(iCount))
'
'      If Not .NoMatch Then
'        fFound = False
'        For iLoop = 1 To UBound(pvTables)
'          If pvTables(iLoop) = !TableID Then
'            fFound = True
'            Exit For
'          End If
'        Next iLoop
'
'        If Not fFound Then
'          ReDim Preserve pvTables(UBound(pvTables) + 1)
'          pvTables(UBound(pvTables)) = !TableID
'        End If
'      End If
'    End With
'  Next iCount
'
'TidyUpAndExit:
'  Exit Function
'
'ErrorTrap:
'  Resume TidyUpAndExit
'
'End Function

'Private Function ReadPermissions(ByRef psErrMsg As String) As Boolean
'  ' Create a collection of OpenHR user groups and their table/view/column permissions
'  ' Return TRUE if everything went okay.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim fSystemManager As Boolean
'  Dim fSecurityManager As Boolean
'  Dim sSQL As String
'  Dim sGroupName As String
'  Dim rsGroups As New ADODB.Recordset
'  Dim rsPermissions As New ADODB.Recordset
'  Dim objGroup As clsSecurityGroup
'  Dim lngAmountOfGroups As Long
'  Dim objPerformance As SystemMgr.clsPerformance
'
'  fOK = True
'  Set gObjGroups = Nothing
'  Set gObjGroups = New clsSecurityGroups
'
'  Set objPerformance = New SystemMgr.clsPerformance
'  objPerformance.ClearLogFile
'
'  'MH20040112 Fault 5627
'  'sSQL = "exec sp_ASRGetUserGroups"
'  sSQL = "SELECT name FROM sysusers " & _
'         "WHERE gid = uid AND gid > 0 " & _
'         "AND not (name like 'ASRSys%') AND not (name like 'db[_]%')"
'
'  ' Get the amount of records first
'  rsGroups.Open sSQL, gADOCon, adOpenKeyset, adLockReadOnly
'  lngAmountOfGroups = rsGroups.RecordCount
'  rsGroups.Close
'
'  OutputCurrentProcess2 vbNullString, lngAmountOfGroups + 1
'
'  rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'  With rsGroups
'
'    If Not .EOF And Not .BOF Then
'      While fOK And (Not .EOF)
'
'        sGroupName = Trim(.Fields(0).Value) 'Trim(!Name)
'
'        OutputCurrentProcess2 sGroupName
'        gobjProgress.UpdateProgress2
'
'        ' Add the group to the groups collection
'        Set objGroup = gObjGroups.Add(sGroupName)
'
'        ' Check if the group is permitted use of the System or Security managers.
'        fSystemManager = False
'        fSecurityManager = False
'        sSQL = "SELECT ASRSysGroupPermissions.permitted, ASRSysPermissionItems.itemKey" & _
'          " FROM ASRSysPermissionItems" & _
'          " INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
'          " INNER JOIN ASRSysGroupPermissions ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID" & _
'          " WHERE (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER' OR ASRSysPermissionItems.itemkey = 'SECURITYMANAGER')" & _
'          " AND ASRSysGroupPermissions.groupName = '" & sGroupName & "'"
'        rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'        Do While (Not rsPermissions.EOF)
'
'          If rsPermissions!permitted Then
'            If rsPermissions!ItemKey = "SYSTEMMANAGER" Then
'              fSystemManager = True
'            Else
'              fSecurityManager = True
'            End If
'          End If
'
'          rsPermissions.MoveNext
'        Loop
'        rsPermissions.Close
'
'        objGroup.SecurityManager = fSecurityManager
'        objGroup.SystemManager = fSystemManager
'
'        ' Initialise the user views collection.
'        objPerformance.StartClock sGroupName
'        fOK = SetupTablesCollection(objGroup)
'        objPerformance.LogSummary
'
'        fOK = fOK And Not gobjProgress.Cancelled
'
'        .MoveNext
'
'      Wend
'    End If
'
'    .Close
'  End With
'
'TidyUpAndExit:
'  'If (Not fOK) And (Len(psErrMsg) = 0) Then
'  If (Not fOK) And Not gobjProgress.Cancelled And _
'    (Len(psErrMsg) = 0) Then
'    OutputError "Error reading table permissions."
'  End If
'  Set rsPermissions = Nothing
'  Set rsGroups = Nothing
'  ReadPermissions = fOK
'  Exit Function
'
'ErrorTrap:
'  'psErrMsg = "Error reading table/view permissions." & vbCr & vbCr & _
'    Err.Description
'  OutputError "Error reading table/view permissions."
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'
'Private Function SetupTablesCollection(pobjGroup As clsSecurityGroup) As Boolean
'  ' Read the list of tables the current user has permission to see.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim fSysSecManager As Boolean
'  Dim fSelectAllPermission As Boolean
'  Dim fSelectNonePermission As Boolean
'  Dim fUpdateAllPermission As Boolean
'  Dim fUpdateNonePermission As Boolean
'  Dim iLoop As Integer
'  Dim lngNextIndex As Long
'  Dim lngRoleID As Long
'  Dim lngChildViewID As Long
'  Dim sSQL As String
'  Dim sLastRealSource As String
'  Dim sRealSourceList As SystemMgr.cStringBuilder
'  Dim sTableViewName As String
'  Dim rsInfo As ADODB.Recordset
'  Dim rsTables As ADODB.Recordset
'  Dim rsViews As ADODB.Recordset
'  Dim rsPermissions As ADODB.Recordset
'  Dim objColumn As clsSecurityColumn
'  Dim objColumns As clsSecurityColumns
'  Dim objTableView As clsSecurityTable
'  Dim avChildViews() As Variant
'  Dim iTemp As Integer
'  Dim fChildView As Boolean
'  Dim sPermissionName As String
'  Dim strObjectName As String
'  Dim strColumnName As String
'  Dim iAction As Integer
'  Dim iSelect As Integer
'  Dim iUpdate As Integer
'  Dim iTableType As Integer
'
'  Set sRealSourceList = New SystemMgr.cStringBuilder
'  Set rsInfo = New ADODB.Recordset
'  Set rsTables = New ADODB.Recordset
'  Set rsViews = New ADODB.Recordset
'  Set rsPermissions = New ADODB.Recordset
'
'  fOK = True
'  fSysSecManager = (pobjGroup.SecurityManager Or pobjGroup.SystemManager)
'
'  ' Create an array of child view IDs and their associated table names.
'  ' Column 1 - child view ID
'  ' Column 2 - associated table name
'  ' Column 3 - 0=OR, 1=AND
'  sSQL = "SELECT ASRSysChildViews2.childViewID, ASRSysTables.tableName, ASRSysChildViews2.type" & _
'    " FROM ASRSysChildViews2" & _
'    " INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID" & _
'    " WHERE ASRSysChildViews2.role = '" & pobjGroup.Name & "'"
'  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'  ReDim avChildViews(3, 100)
'  lngNextIndex = -1
'  If Not rsInfo.EOF Then
'    Do While Not rsInfo.EOF
'      lngNextIndex = lngNextIndex + 1
'      If lngNextIndex > UBound(avChildViews, 2) Then ReDim Preserve avChildViews(3, lngNextIndex + 100)
'      avChildViews(1, lngNextIndex) = rsInfo(0).Value
'      avChildViews(2, lngNextIndex) = rsInfo(1).Value
'      avChildViews(3, lngNextIndex) = IIf(IsNull(rsInfo(2).Value), 0, rsInfo(2).Value)
'      rsInfo.MoveNext
'    Loop
'    ReDim Preserve avChildViews(3, lngNextIndex)
'  End If
'  rsInfo.Close
'
'  ' Get the collection with items for each TABLE in the system.
'  sSQL = "SELECT tableID, tableName, tableType FROM ASRSysTables"
'  rsTables.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'  Do While Not rsTables.EOF
'    Set objColumns = New clsSecurityColumns
'    pobjGroup.Tables.Add objColumns, rsTables(1).Value, rsTables(2).Value
'    Set objColumns = Nothing
'
'    With pobjGroup.Tables(rsTables(1).Value)
'      .SelectPrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, IIf(rsTables(2).Value = iTabLookup, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED))
'      .UpdatePrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
'      .InsertPrivilege = IIf(fSysSecManager, True, False)
'      .DeletePrivilege = IIf(fSysSecManager, True, False)
'      .ParentJoinType = 0
'    End With
'    rsTables.MoveNext
'  Loop
'  rsTables.Close
'
'
'  ' Initialise the collection with items for each VIEW in the system.
'  sSQL = "SELECT ASRSysViews.viewName FROM ASRSysViews"
'  rsViews.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'  Do While Not rsViews.EOF
'    Set objColumns = New clsSecurityColumns
'
'    sTableViewName = rsViews(0).Value  ' ViewName
'    pobjGroup.Views.Add objColumns, sTableViewName, 0
'    Set objColumns = Nothing
'
'    With pobjGroup.Views(sTableViewName)
'      .SelectPrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
'      .UpdatePrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
'      .InsertPrivilege = IIf(fSysSecManager, True, False)
'      .DeletePrivilege = IIf(fSysSecManager, True, False)
'      .ParentJoinType = 0
'    End With
'   rsViews.MoveNext
'  Loop
'  rsViews.Close
'
'
'  ' Get the permissions for each table or view.
'  sRealSourceList.TheString = vbNullString
'  sLastRealSource = vbNullString
'
'  If Not fSysSecManager Then
'    ' If the user is NOT a 'system manager' or 'security manager'
'    ' read the table permissions from the server.
'    sSQL = "exec sp_ASRAllTablePermissionsForGroup '" & pobjGroup.Name & "'"
'    rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    Do While Not rsPermissions.EOF
'      Set objTableView = Nothing
'
'      sPermissionName = UCase(rsPermissions.Fields(0).Value)   ' Name
'      iAction = rsPermissions.Fields(1).Value                   ' Action
'
'      If sLastRealSource <> sPermissionName Then
'        sRealSourceList.Append IIf(sRealSourceList.Length <> 0, ", '", "'") & sPermissionName & "'"
'        sLastRealSource = sPermissionName
'      End If
'
'      If (iAction = 195) Or (iAction = 196) Then
'        fChildView = False
'
'        If Left$(sPermissionName, 8) = "ASRSYSCV" Then
'          fChildView = True
'          ' Determine which table the child view is for.
'          iTemp = InStr(sPermissionName, "#")
'          lngChildViewID = Val(Mid$(sPermissionName, 9, iTemp - 9))
'        End If
'
'        If fChildView Then
'          For lngNextIndex = 0 To UBound(avChildViews, 2)
'            If avChildViews(1, lngNextIndex) = lngChildViewID Then
'              Set objTableView = pobjGroup.Tables(avChildViews(2, lngNextIndex))
'              objTableView.ParentJoinType = avChildViews(3, lngNextIndex)
'              Exit For
'            End If
'          Next lngNextIndex
'        Else
'          If pobjGroup.Tables.IsValid(sPermissionName) Then
'            Set objTableView = pobjGroup.Tables(sPermissionName)
'          Else
'            Set objTableView = pobjGroup.Views(sPermissionName)
'          End If
'        End If
'
'        If Not objTableView Is Nothing Then
'          Select Case iAction
'            Case 195 ' Insert permission.
'              objTableView.InsertPrivilege = True
'            Case 196 ' Delete permission.
'              objTableView.DeletePrivilege = True
'          End Select
'        End If
'      End If
'
'      rsPermissions.MoveNext
'    Loop
'    rsPermissions.Close
'  End If
'
'  ' Get the list of all columns in all tables/views.
'  sSQL = "SELECT ASRSysColumns.columnName," & _
'    " ASRSysColumns.columnID," & _
'    " ASRSysTables.tableName AS tableViewName," & _
'    " ASRSysTables.tableType AS tableType" & _
'    " FROM ASRSysColumns" & _
'    " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
'    " WHERE ASRSysColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'    " AND ASRSysColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK)) & _
'    " UNION" & _
'    " SELECT ASRSysColumns.columnName," & _
'    " ASRSysColumns.columnID," & _
'    " ASRSysViews.viewName AS tableViewName," & _
'    " 0 AS tableType" & _
'    " FROM ASRSysColumns" & _
'    " INNER JOIN ASRSysViews ON ASRSysColumns.tableID = ASRSysViews.viewTableID" & _
'    " INNER JOIN ASRSysViewColumns ON (ASRSysColumns.columnID = ASRSysViewColumns.columnID AND ASRSysViews.viewID = ASRSysViewColumns.viewID)" & _
'    " WHERE ASRSysViewColumns.inView = 1" & _
'    " AND ASRSysColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'    " AND ASRSysColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK))
'  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'  Do While Not rsInfo.EOF
'
'    strColumnName = rsInfo.Fields(0).Value
'    iTableType = rsInfo.Fields(3).Value
'
'    If iTableType <> iTabView Then
'      Set objColumns = pobjGroup.Tables(rsInfo.Fields(2).Value).Columns
'    Else
'      Set objColumns = pobjGroup.Views(rsInfo.Fields(2).Value).Columns
'    End If
'
'    ' Add the column object to the collection.
'    Set objColumn = objColumns.Add(UCase$(Trim$(strColumnName)))
'
'    ' Set the security column properties
'    objColumn.Name = strColumnName
'    objColumn.ColumnID = rsInfo.Fields(1).Value
'    objColumn.SelectPrivilege = fSysSecManager Or (iTableType = iTabLookup)
'    objColumn.UpdatePrivilege = fSysSecManager
'
'    ' Release the security column
'    Set objColumn = Nothing
'    Set objColumns = Nothing
'
'    rsInfo.MoveNext
'
'  Loop
'  rsInfo.Close
'
'  ' If the current user is not a system/security manager then read the column permissions from SQL.
'  If (Not fSysSecManager) And (sRealSourceList.Length <> 0) Then
'    ' Get the SQL group id of the current user.
'    sSQL = "SELECT gid" & _
'      " FROM sysusers" & _
'      " WHERE name = '" & pobjGroup.Name & "'"
'    rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'    lngRoleID = rsInfo.Fields(0).Value
'    rsInfo.Close
'
'    sSQL = "EXEC dbo.[spASRGetAllTableAndViewColumnPermissionsForGroup] " & lngRoleID
'    rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'    Do While Not rsInfo.EOF
'
'      ' Get the current column's table/view name.
'      Set objTableView = Nothing
'
'      strObjectName = UCase(rsInfo.Fields(0).Value)
'      strColumnName = rsInfo.Fields(1).Value
'      iSelect = rsInfo.Fields(2).Value
'      iUpdate = rsInfo.Fields(3).Value
'
'      fChildView = False
'      If Left$(strObjectName, 8) = "ASRSYSCV" Then
'        fChildView = True
'        ' Determine which table the child view is for.
'        iTemp = InStr(strObjectName, "#")
'        lngChildViewID = Val(Mid$(strObjectName, 9, iTemp - 9))
'      End If
'
'      If fChildView Then
'        For lngNextIndex = 0 To UBound(avChildViews, 2)
'          If avChildViews(1, lngNextIndex) = lngChildViewID Then
'            Set objTableView = pobjGroup.Tables(avChildViews(2, lngNextIndex))
'            objTableView.ParentJoinType = avChildViews(3, lngNextIndex)
'            Exit For
'          End If
'        Next lngNextIndex
'
'      Else
'        If pobjGroup.Tables.IsValid(strObjectName) Then
'          Set objTableView = pobjGroup.Tables(strObjectName)
'        Else
'          Set objTableView = pobjGroup.Views(strObjectName)
'        End If
'      End If
'
'
'      If Not objTableView Is Nothing Then
'
'        If objTableView.Columns.IsValid(strColumnName) Then
'          objTableView.Columns(strColumnName).SelectPrivilege = iSelect
'          objTableView.Columns(strColumnName).UpdatePrivilege = iUpdate
'        End If
'
''        If iAction = 193 Then
''          If objTableView.Columns.IsValid(strColumnName) Then
''            objTableView.Columns(strColumnName).SelectPrivilege = bPermission
''          End If
''        End If
''
''        If iAction = 197 Then
''          If objTableView.Columns.IsValid(strColumnName) Then
''            objTableView.Columns(strColumnName).UpdatePrivilege = bPermission
''          End If
''        End If
'      End If
'
'      rsInfo.MoveNext
'    Loop
'    rsInfo.Close
'
'    ' Check if the table has SELECT/UPDATE ALL/SOME/NONE.
'    For Each objTableView In pobjGroup.Tables.Collection
'      fSelectAllPermission = True
'      fSelectNonePermission = True
'      fUpdateAllPermission = True
'      fUpdateNonePermission = True
'
'      For Each objColumn In objTableView.Columns.Collection
'        If objColumn.SelectPrivilege Then
'          fSelectNonePermission = False
'        Else
'          fSelectAllPermission = False
'        End If
'
'        If objColumn.UpdatePrivilege Then
'          fUpdateNonePermission = False
'        Else
'          fUpdateAllPermission = False
'        End If
'      Next objColumn
'
'      objTableView.SelectPrivilege = IIf(fSelectAllPermission, giPRIVILEGES_ALLGRANTED, IIf(fSelectNonePermission, giPRIVILEGES_NONEGRANTED, giPRIVILEGES_SOMEGRANTED))
'      objTableView.UpdatePrivilege = IIf(fUpdateAllPermission, giPRIVILEGES_ALLGRANTED, IIf(fUpdateNonePermission, giPRIVILEGES_NONEGRANTED, giPRIVILEGES_SOMEGRANTED))
'    Next objTableView
'  End If
'
'TidyUpAndExit:
'  Set rsInfo = Nothing
'  Set rsTables = Nothing
'  Set rsViews = Nothing
'  Set objTableView = Nothing
'  Set objColumn = Nothing
'  Set rsPermissions = Nothing
'
'  SetupTablesCollection = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error Reading Tables Collection"
'  Resume TidyUpAndExit
'
'End Function

Private Function PermittedChildView(plngTableID As Long, psGroupname As String) As Long
''Private Function PermittedChildView(psTableName As String, psGroupname As String) As Long
'  ' Return the ID of the child view on the given table that is appropriate for the given group (role).
'  ' Return 0 if no view is appropriate.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim fTableOK As Boolean
'  Dim fViewOK As Boolean
'  Dim iLoop As Integer
'  Dim iNextIndex As Integer
'  Dim lngChildViewID As Long
'  Dim lngParentViewID As Long
'  Dim sSQL As String
'  Dim sCode As String
'  Dim rsViews As dao.Recordset
'  Dim rsChildView As rdoResultset
'  Dim rsParents As dao.Recordset
'  Dim rsTable  As dao.Recordset
'  Dim avParents() As Variant
'  Dim iParentJoinType As Integer
'  Dim iParentCount As Integer
'  Dim iOKViewCount As Integer
'
'  fOK = True
'  lngChildViewID = 0
'  iParentCount = 0
'
'  ' Create an array of the parents of the given table that are accessible by the given group.
'  ' Column 1 = parent type (UT = top-level table
'  '                         UV = view of a top-level table
'  '                         SV = system view)
'  ' Column2 = parent ID
'  ReDim avParents(2, 0)
'
'  ' Get the given table's parents.
'  sSQL = "SELECT tmpTables.tableID, tmpTables.originalTableName, tmpTables.tableType," & _
'        " tmpTables.copySecurityTableID, tmpTables.copySecurityTableName, tmpTables.grantRead" & _
'        " FROM tmpRelations" & _
'        " INNER JOIN tmpTables ON tmpRelations.parentID = tmpTables.tableID" & _
'        " WHERE tmpRelations.childID = " & trim(str(pLngTableID))
'
'  Set rsParents = daoDb.OpenRecordset(sSQL, _
'    dbOpenForwardOnly, dbReadOnly)
'  With rsParents
'    ' Loop through the given table's parents, adding the permitted view of each to the array of parents.
'    Do While Not .EOF
'      If !TableType = iTabParent Then
'        ' Parent is a top-level table.
'        If gObjGroups(psGroupname).Tables.IsValid(!OriginalTableName) Then
'          ' Table is in the tables collection.
'          fTableOK = (gObjGroups(psGroupname).Tables(!OriginalTableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
'        Else
'          ' Table is NOT in the tables collection. Must be new,
'          ' so see if we're copying permissions from another table, or
'          ' if the default permissions are specified.
'          If !copySecurityTableID > 0 Then
'            fTableOK = (gObjGroups(psGroupname).Tables(!copySecurityTableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
'          Else
'            fTableOK = !GrantRead
'          End If
'        End If
'
'        If fTableOK Then
'          iParentCount = iParentCount + 1
'
'          ' The current group has permission to see all records in the parent table.
'          iNextIndex = UBound(avParents, 2) + 1
'          ReDim Preserve avParents(2, iNextIndex)
'          avParents(1, iNextIndex) = "UT"
'          avParents(2, iNextIndex) = !TableID
'        Else
'          ' The current group does NOT have permission to see all records in the parent table.
'          ' Get the permitted views on the table.
'          sSQL = "SELECT tmpViews.viewID, tmpViews.viewName, tmpViews.grantRead" & _
'            " FROM tmpViews" & _
'            " WHERE tmpViews.viewTableID = " & trim(str(!TableID)) & _
'            " AND tmpViews.deleted = FALSE"
'          iOKViewCount = 0
'
'          Set rsViews = daoDb.OpenRecordset(sSQL, _
'            dbOpenForwardOnly, dbReadOnly)
'          With rsViews
'            Do While Not .EOF
'              If gObjGroups(psGroupname).Views.IsValid(!ViewName) Then
'                ' View is in the views collection.
'                fViewOK = (gObjGroups(psGroupname).Views(!ViewName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
'              Else
'                ' Table is NOT in the tables collection. Must be new,
'                ' so get the default permissions.
'                fViewOK = !GrantRead
'              End If
'
'              If fViewOK Then
'                iOKViewCount = iOKViewCount + 1
'
'                iNextIndex = UBound(avParents, 2) + 1
'                ReDim Preserve avParents(2, iNextIndex)
'                avParents(1, iNextIndex) = "UV"
'                avParents(2, iNextIndex) = !ViewID
'              End If
'
'              .MoveNext
'            Loop
'
'            .Close
'          End With
'          Set rsViews = Nothing
'
'          If iOKViewCount > 0 Then
'            iParentCount = iParentCount + 1
'          End If
'        End If
'      ElseIf !TableType = iTabChild Then
'        ' Parent is not a top-level table.
'        'lngParentViewID = PermittedChildView(!TableName, psGroupname)
'        lngParentViewID = PermittedChildView(!TableID, psGroupname)
'
'        If lngParentViewID > 0 Then
'          iParentCount = iParentCount + 1
'
'          iNextIndex = UBound(avParents, 2) + 1
'          ReDim Preserve avParents(2, iNextIndex)
'          avParents(1, iNextIndex) = "SV"
'          avParents(2, iNextIndex) = lngParentViewID
'        End If
'      End If
'
'      .MoveNext
'    Loop
'
'    .Close
'  End With
'  Set rsParents = Nothing
'
'  iParentJoinType = 0
'  If iParentCount > 1 Then
'    ' More than 1 parent. Do we want the OR join child view, or the AND join child view ?
'    sSQL = "SELECT tmpTables.originalTableName" & _
'      " FROM tmpTables" & _
'      " WHERE tmpTables.tableID = " & trim(str(pLngTableID))
'
'    Set rsTable = daoDb.OpenRecordset(sSQL, _
'      dbOpenForwardOnly, dbReadOnly)
'    With rsTable
'      If gObjGroups(psGroupname).Tables.IsValid(!OriginalTableName) Then
'        iParentJoinType = gObjGroups(psGroupname).Tables(!OriginalTableName).ParentJoinType
'      End If
'    End With
'
'    rsTable.Close
'    Set rsTable = Nothing
'  End If
'
'  If UBound(avParents, 2) > 0 Then
'    ' Get the child view permutation that is configured for the permitted set of parents.
'    For iLoop = 1 To UBound(avParents, 2)
'      sCode = sCode & _
'        " INNER JOIN ASRSysChildViewParents tmpTable_" & trim(str(iLoop)) & _
'        " ON (ASRSysChildViews.childViewID = tmpTable_" & trim(str(iLoop)) & ".childViewID" & _
'        " AND tmpTable_" & trim(str(iLoop)) & ".parentType = '" & avParents(1, iLoop) & "'" & _
'        " AND tmpTable_" & trim(str(iLoop)) & ".parentID = " & trim(str(avParents(2, iLoop))) & ")"
'    Next iLoop
'
'    sSQL = "SELECT ASRSysChildViews.childViewID" & _
'      " FROM ASRSysChildViews" & _
'      sCode & _
'      " INNER JOIN ASRSysTables ON ASRSysChildViews.tableID = ASRSysTables.tableID" & _
'      " INNER JOIN ASRSysChildViewParents parentCount" & _
'      " ON (ASRSysChildViews.childViewID = parentCount.childViewID)" & _
'      " GROUP BY ASRSysTables.tableID, ASRSysChildViews.childViewID, ASRSysTables.tableName, ASRSysChildViews.type" & _
'      " HAVING ASRSysTables.tableID = " & trim(str(pLngTableID)) & _
'      " AND " & IIf(iParentJoinType = 0, "(ASRSysChildViews.type = 0 OR ASRSysChildViews.type IS NULL)", "ASRSysChildViews.type = 1") & _
'      " AND COUNT(parentCount.childViewID) = " & trim(str(UBound(avParents, 2)))
'
'    Set rsChildView = rdoCon.OpenResultset(sSQL)
'    fOK = Not (rsChildView.BOF And rsChildView.EOF)
'    If fOK Then
'      lngChildViewID = rsChildView!childViewID
'    End If
'  End If
'
'TidyUpAndExit:
'  If fOK Then
'    PermittedChildView = lngChildViewID
'  Else
'    PermittedChildView = 0
'  End If
'
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
End Function

'Private Function ApplyPermissions() As Boolean
'  ' Grant permissions to the tables/views/columns.
'  ' Return TRUE if everything passed off okay.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'
'  ' Apply the Table and Table Column permissions.
'  OutputCurrentProcess2 "Parent Tables", 3
'  fOK = ApplyPermissions_NonChildTables
'  fOK = fOK And Not gobjProgress.Cancelled
'  gobjProgress.UpdateProgress2
'
'  If fOK Then
'    ' Apply the View and View Column permissions.
'    OutputCurrentProcess2 "Views", 3
'    fOK = ApplyPermissions_UserViews
'    fOK = fOK And Not gobjProgress.Cancelled
'    gobjProgress.UpdateProgress2
'  End If
'
'  If fOK Then
'    ' Apply the new Child Table (and child view) permissions.
'    OutputCurrentProcess2 "Child Tables", 3
'    fOK = ApplyPermissions_ChildTables2
'    gobjProgress.UpdateProgress2
'  End If
'
'TidyUpAndExit:
'  ApplyPermissions = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error Applying Permissions"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'
'Private Function ApplyPermissions_NonChildTables() As Boolean
'  ' Apply the top-level and lookup Table and Table Column permissions to SQL Server database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'  Dim sGroupName As String
'  Dim sTableName As String
'  Dim objGroup As clsSecurityGroup
'  Dim rsTables As dao.Recordset
'  Dim rsColumns As dao.Recordset
'  Dim lngOriginalTableID As Long
'  Dim sOriginalTableName As String
'  Dim strOriginalColumnName As String
'  Dim strColumnName As String
'  Dim sSelectGrant As String
'  Dim sUpdateGrant As String
'
'  fOK = True
'
'  ' Get the set of top-level and lookup tables.
'  sSQL = "SELECT tableName, tableID, OriginalTableName, changed, new," & _
'    " CopySecurityTableID, CopySecurityTableName, tableType," & _
'    " GrantRead, GrantNew, GrantEdit, GrantDelete" & _
'    " FROM tmpTables" & _
'    " WHERE (tableType = " & Trim$(Str$(iTabParent)) & _
'    " OR tableType = " & Trim$(Str$(iTabLookup)) & ")" & _
'    " AND deleted = FALSE"
'  Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  Do While Not rsTables.EOF
'    For Each objGroup In gObjGroups.Collection
'      With objGroup
'        sGroupName = "[" & .Name & "]"
'        sTableName = rsTables.Fields(0).Value
'
'        If objGroup.SecurityManager Or _
'          objGroup.SystemManager Then
'
'          sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sTableName & " TO " & sGroupName
'          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'        Else
'
'          If (Not rsTables!New) Or (rsTables!copySecurityTableID > 0) Then
'            ' Initialise the Table Column permissions command strings.
'            sSelectGrant = vbNullString
'            sUpdateGrant = vbNullString
'
'            ' Get the set of non-system columns in the table.
'            If rsTables!copySecurityTableID > 0 Then
'              lngOriginalTableID = rsTables!copySecurityTableID
'              sOriginalTableName = rsTables!copySecurityTableName
'            Else
'              lngOriginalTableID = rsTables!TableID
'              sOriginalTableName = rsTables!OriginalTableName
'            End If
'
'            With .Tables.Item(sOriginalTableName)
'
'              If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Or .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
'
'                sSQL = "SELECT columnName, columnID, OriginalColumnName, new" & _
'                  " FROM tmpColumns" & _
'                  " WHERE tableID = " & Trim$(Str$(lngOriginalTableID)) & _
'                  " AND deleted = FALSE" & _
'                  " AND columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'                  " AND columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK))
'                Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'                Do While Not rsColumns.EOF
'
'                  'MH20060712 Fault 11313 - Addendum to JDM's work below.  :O)
'                  '"OriginalColumnName" is only used for existing columns hence moved into ELSE clause
'                  ''''JDM - 22/06/2006 - Fault 11186 - Addendum to MH's work below. Moved column generation list from outside loop to inside
'                  ''''strOriginalColumnName = UCase(Trim(rsColumns.Fields(2).Value))
'
'
'                  strColumnName = UCase(Trim(rsColumns.Fields(0).Value))
'
'                  If rsColumns.Fields(3).Value Then   ' Is new column
'
'                    ' Build string of columns that are allowed
'                    If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
'                      sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & strColumnName
'                    End If
'
'                    If .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
'                      sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & strColumnName
'                    End If
'
'                  Else
'                    'MH20060712 Fault 11313
'                    strOriginalColumnName = UCase(Trim(rsColumns.Fields(2).Value))
'
'                    ' Build string of columns that are revoked based on what they had before
'                    If .Columns.Item(strOriginalColumnName).SelectPrivilege Or .TableType = iTabLookup Then
'                      sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & strColumnName
'                    End If
'
'                    If .Columns.Item(strOriginalColumnName).UpdatePrivilege Then
'                      sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & strColumnName
'                    End If
'                  End If
'
'                  rsColumns.MoveNext
'                Loop
'
'                rsColumns.Close
'                Set rsColumns = Nothing
'              End If
'
'              ' Delete permissions
'              If .DeletePrivilege Then
'                sSQL = "GRANT DELETE ON " & sTableName & " TO " & sGroupName
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              End If
'
'              ' Insert permissions
'              If .InsertPrivilege Then
'                sSQL = "GRANT INSERT ON " & sTableName & " TO " & sGroupName
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              End If
'
'              ' Select permissions
'              If .SelectPrivilege = giPRIVILEGES_ALLGRANTED Or (.TableType = iTabLookup) Then
'                sSQL = "GRANT SELECT ON " & sTableName & " TO " & sGroupName
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              ElseIf .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
'                If LenB(sSelectGrant) <> 0 Then
'
'                  sSQL = "REVOKE SELECT ON " & sTableName & " TO " & sGroupName
'                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'                  'MH20060620 Fault 11186
'                  'sSQL = "GRANT SELECT(ID, " & sSelectGrant & ") ON " & sTableName & " TO " & sGroupName
'                  sSQL = "GRANT SELECT(ID,TimeStamp," & sSelectGrant & ") ON " & sTableName & " TO " & sGroupName
'                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                End If
'              Else
'                gADOCon.Execute "GRANT SELECT(id) ON " & sTableName & " TO " & sGroupName
'              End If
'
'              ' Update permissions
'              If .UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
'                sSQL = "GRANT UPDATE ON " & sTableName & " TO " & sGroupName
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              ElseIf .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
'                If LenB(sUpdateGrant) <> 0 Then
'
'                  sSQL = "REVOKE UPDATE ON " & sTableName & " TO " & sGroupName
'                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'                  'MH20060620 Fault 11186
'                  'sSQL = "GRANT UPDATE(" & sUpdateGrant & ") ON " & sTableName & " TO " & sGroupName
'                  sSQL = "GRANT UPDATE(ID,Timestamp," & sUpdateGrant & ") ON " & sTableName & " TO " & sGroupName
'                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                End If
'              End If
'
'            End With
'
'          Else
'            ' New table, not having permissions copied from another table.
'            ' Put user defined security settings on each group
'            If rsTables!GrantDelete Then
'              sSQL = "GRANT DELETE ON " & sTableName & " TO " & sGroupName
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'            End If
'
'            If rsTables!GrantNew Then
'              sSQL = "GRANT INSERT ON " & sTableName & " TO " & sGroupName
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'            End If
'
'            If rsTables!GrantRead Or (rsTables!TableType = iTabLookup) Then
'              sSQL = "GRANT SELECT ON " & sTableName & " TO " & sGroupName
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'            End If
'
'            gADOCon.Execute "GRANT SELECT(id) ON " & sTableName & " TO " & sGroupName
'
'            If rsTables!GrantEdit Then
'              sSQL = "GRANT UPDATE ON " & sTableName & " TO " & sGroupName
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'            End If
'
'          End If
'        End If
'      End With
'    Next objGroup
'    Set objGroup = Nothing
'
'    rsTables.MoveNext
'  Loop
'  rsTables.Close
'  Set rsTables = Nothing
'
'TidyUpAndExit:
'  Set objGroup = Nothing
'  ApplyPermissions_NonChildTables = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error Applying Permissions (Non-Child Tables)"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'
'
'
'Private Function ApplyPermissions_UserViews() As Boolean
'  ' Apply the user-defined views permissions to SQL Server database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'  Dim sGroupName As String
'  Dim sViewName As String
'  Dim sOriginalViewName As String
''  Dim sSelectDeny As String
''  Dim sUpdateDeny As String
'  Dim objGroup As clsSecurityGroup
'  Dim rsViews As dao.Recordset
'  Dim rsColumns As dao.Recordset
'  Dim sPermissionTypes As String
'  Dim strOriginalColumnName As String
'  Dim strColumnName As String
'  Dim sSelectGrant As String
'  Dim sUpdateGrant As String
'
'  fOK = True
'
'  ' Get the set of user defined views.
'  sSQL = "SELECT viewName, viewID, viewTableID, OriginalViewName, changed, new," & _
'    " GrantRead, GrantNew, GrantEdit, GrantDelete" & _
'    " FROM tmpViews" & _
'    " WHERE deleted = FALSE"
'  Set rsViews = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  Do While Not rsViews.EOF
'    For Each objGroup In gObjGroups.Collection
'      With objGroup
'        sGroupName = "[" & .Name & "]"
'        sViewName = rsViews!ViewName
'        sOriginalViewName = rsViews!OriginalViewName
'
'        If objGroup.SecurityManager Or _
'          objGroup.SystemManager Then
'          sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sViewName & " TO " & sGroupName
'          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'        Else
'          If Not rsViews!New Then
'
'            ' Initialise the column permissions command strings.
'            sSelectGrant = vbNullString
'            sUpdateGrant = vbNullString
'
'            With .Views.Item(sOriginalViewName)
'
'              ' Get the set of non-system columns in the view.
'              sSQL = "SELECT tmpColumns.columnName, tmpColumns.columnID, tmpColumns.OriginalColumnName, tmpColumns.new" & _
'                " FROM tmpViewColumns, tmpColumns" & _
'                " WHERE (tmpViewColumns.ColumnID = tmpColumns.ColumnID" & _
'                " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'                " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK)) & _
'                " AND tmpViewColumns.InView = TRUE" & _
'                " AND tmpViewColumns.ViewID = " & Trim(Str(rsViews!ViewID)) & ")"
'              Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'              Do While Not rsColumns.EOF
'
'                strOriginalColumnName = UCase(Trim(IIf(IsNull(rsColumns.Fields(2).Value), "", rsColumns.Fields(2).Value)))
'                strColumnName = UCase(Trim(rsColumns.Fields(0).Value))
'
'                If Not .Columns.IsValid(strOriginalColumnName) Then
'
'                  ' New column
'                  If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
'                    sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & strColumnName
'                  End If
'
'                  If .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
'                    sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & strColumnName
'                  End If
'                Else
'
'                  ' Existing column
'                  If .Columns.Item(strOriginalColumnName).SelectPrivilege Then
'                    sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & strColumnName
'                  End If
'
'                  If .Columns.Item(strOriginalColumnName).UpdatePrivilege Then
'                    sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & strColumnName
'                  End If
'                End If
'
'                rsColumns.MoveNext
'              Loop
'              rsColumns.Close
'              Set rsColumns = Nothing
'
'              ' Delete permission
'              If .DeletePrivilege Then
'                sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              End If
'
'              ' Insert permission
'              If .InsertPrivilege Then
'                sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              End If
'
'              ' Select permissions
'              If .SelectPrivilege = giPRIVILEGES_ALLGRANTED Then
'                sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              ElseIf .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
'                If LenB(sSelectGrant) <> 0 Then
'                  'MH20060620 Fault 11186
'                  'sSQL = "GRANT SELECT(" & sSelectGrant & ") ON " & sViewName & " TO " & sGroupName
'                  sSQL = "GRANT SELECT(ID,Timestamp," & sSelectGrant & ") ON " & sViewName & " TO " & sGroupName
'                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                End If
'              End If
'
'              ' Update permissions
'              If .UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
'                sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              ElseIf .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
'                If LenB(sUpdateGrant) <> 0 Then
'                  'MH20060620 Fault 11186
'                  'sSQL = "GRANT UPDATE(" & sUpdateGrant & ") ON " & sViewName & " TO " & sGroupName
'                  sSQL = "GRANT UPDATE(ID,Timestamp," & sUpdateGrant & ") ON " & sViewName & " TO " & sGroupName
'                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                End If
'              End If
'
'            End With
'
'          Else
'            ' New view, not having permissions copied from another view.
'            ' Put user defined security settings on each group
'            If rsViews!GrantDelete Then
'              sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'            End If
'
'            If rsViews!GrantNew Then
'              sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'            End If
'
'            If rsViews!GrantRead Then
'              sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'            End If
'
'            If rsViews!GrantEdit Then
'              sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'            End If
'
'
'          End If
'
'        End If
'      End With
'    Next objGroup
'    Set objGroup = Nothing
'
'    rsViews.MoveNext
'  Loop
'  rsViews.Close
'  Set rsViews = Nothing
'
'TidyUpAndExit:
'  Set objGroup = Nothing
'  ApplyPermissions_UserViews = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error Applying Permissions (User Views)"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'Private Function ApplyPermissions_ChildTables2() As Boolean
'  ' Apply the child table and permutated view permissions to SQL Server database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'  Dim sAllSQLCommands As String
'  Dim sGroupName As String
'  Dim sTableName As String
'  'Dim sSelectDeny As String
'  'Dim sUpdateDeny As String
'  Dim sSelectGrant As String
'  Dim sUpdateGrant As String
'  Dim objGroup As clsSecurityGroup
'  Dim rsTables As dao.Recordset
'  Dim rsColumns As dao.Recordset
'  Dim cmdChildView As ADODB.Command
'  Dim pmADO As ADODB.Parameter
'  Dim lngViewID As Long
'  Dim sViewName As String
'  Dim avChildTables() As Variant
'  Dim iMaxRouteLength As Integer
'  Dim rsChildren As dao.Recordset
'  Dim rsParents As dao.Recordset
'  Dim rsViews As dao.Recordset
'  Dim rsInfo As ADODB.Recordset
'  Dim rsChildViews As ADODB.Recordset
'  Dim iNextIndex As Integer
'  Dim iLoop As Integer
'  Dim iLoop1 As Integer
'  Dim iLoop2 As Integer
'  Dim iParentCount As Integer
'  Dim avParents() As Variant
'  Dim fTableOK As Boolean
'  Dim iOKViewCount As Integer
'  Dim fViewOK As Boolean
'  Dim sTemp As String
'  Dim iParentJoinType As Integer
'  Dim lngParentViewID As Long
'  Dim sCreatedChildViews As SystemMgr.cStringBuilder
'  Dim lngLastParentID As Long
'  Dim lngOriginalTableID As Long
'  Dim sOriginalTableName As String
'  Dim fTableReadable As Boolean
'  Dim sRelatedChildTables As String
'  Dim sSysSecRoles As String
'  Dim sNonSysSecRoles As String
'  Dim lngPreviousTimeOut As Long
'  Dim sColumnName As String
'  Dim sOriginalColumnName As String
'  Dim strParentIDs As String
'
'  Set sCreatedChildViews = New SystemMgr.cStringBuilder
'  Set rsInfo = New ADODB.Recordset
'  Set rsChildViews = New ADODB.Recordset
'
'  fOK = True
'  sCreatedChildViews.TheString = "0"
'  sRelatedChildTables = "0"
'  sAllSQLCommands = vbNullString
'
'  ' Drop all existing child views.
'  sSQL = "SELECT name FROM sysobjects WHERE name LIKE 'ASRSysCV%' AND xtype = 'V'"
'  rsInfo.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly
'
'  With rsInfo
'    Do While (Not .EOF)
'      sSQL = "DROP VIEW " & .Fields(0).Value
'      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'      .MoveNext
'    Loop
'
'    .Close
'  End With
'
'  ' Create an array of all child tables, and each child's longest route to the top-level.
'  ' eg.
'  '              Table A
'  '               / |
'  '              /  |
'  '             /   |
'  '       Table B   |
'  '             \   |
'  '              \  |
'  '               \ |
'  '              Table C
'  '
'  ' Table A is a top-level table.
'  ' Table B has a longest route to the top-level of 1.
'  ' Table C has a longest route to the top-level of 2.
'  '
'  ' We need to create views for the tables nearest the top-level first, as they might then
'  ' need to be propogated down. So, even though Tables B and C are both children of table A,
'  ' we need to create the views on Table B first.
'
'  ' Create an array of child tables.
'  ' Column 1 = table ID.
'  ' Column 2 = longest route to the top-level.
'  ReDim avChildTables(2, 0)
'  iMaxRouteLength = 0
'  sSQL = "SELECT DISTINCT tmpRelations.childID, tmpTables.tableName" & _
'    " FROM tmpRelations, tmpTables" & _
'    " WHERE tmpRelations.childID = tmpTables.tableID"
'  Set rsChildren = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  ReDim avChildTables(2, 100)
'  iNextIndex = -1
'  If Not rsChildren.EOF Then
'    Do While Not rsChildren.EOF
'      iNextIndex = iNextIndex + 1
'      If iNextIndex > UBound(avChildTables, 2) Then ReDim Preserve avChildTables(2, iNextIndex + 100)
'      avChildTables(1, iNextIndex) = rsChildren.Fields(0).Value
'      avChildTables(2, iNextIndex) = LongestRouteToTopLevel(rsChildren.Fields(0).Value)
'
'      sRelatedChildTables = sRelatedChildTables & "," & Trim(Str(rsChildren.Fields(0).Value))
'
'      If iMaxRouteLength < avChildTables(2, iNextIndex) Then
'        iMaxRouteLength = CInt(avChildTables(2, iNextIndex))
'      End If
'
'      rsChildren.MoveNext
'    Loop
'    ReDim Preserve avChildTables(2, iNextIndex)
'  End If
'  rsChildren.Close
'  Set rsChildren = Nothing
'
'
'  ' Deny non-SysMgr and non-SecMgr users access to orphaned child tables.
'  sSysSecRoles = vbNullString
'  sNonSysSecRoles = vbNullString
'  For Each objGroup In gObjGroups.Collection
'    sGroupName = IIf(IsVersion7, "[" & objGroup.Name & "]", objGroup.Name)
'
'    If (objGroup.SecurityManager) Or (objGroup.SystemManager) Then
'      sSysSecRoles = sSysSecRoles & IIf(LenB(sSysSecRoles) <> 0, ",", vbNullString) & sGroupName
'    Else
'      sNonSysSecRoles = sNonSysSecRoles & IIf(LenB(sNonSysSecRoles) <> 0, ",", vbNullString) & sGroupName
'    End If
'  Next objGroup
'
'  ' JPD20021120 Fault 4793
'  sSQL = "SELECT tmpTables.tableName" & _
'    " FROM tmpTables" & _
'    " WHERE tmpTables.tableType = " & Trim$(Str$(iTabChild)) & _
'    " AND tmpTables.tableID NOT IN (" & sRelatedChildTables & ")" & _
'    " AND tmpTables.deleted = FALSE"
'  Set rsChildren = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'  With rsChildren
'    Do While (Not .EOF)
'      If LenB(sSysSecRoles) <> 0 Then
'        sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & .Fields(0).Value & " TO " & sSysSecRoles
'        sAllSQLCommands = sAllSQLCommands & vbNewLine & sSQL
'        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'      End If
'
''''''      If Len(sNonSysSecRoles) > 0 Then
''''''        sSQL = "DENY DELETE, INSERT, SELECT, UPDATE ON " & .Fields(0).Value & " TO " & sNonSysSecRoles
''''''        sAllSQLCommands = sAllSQLCommands & vbNewLine & sSQL
''''''        gADOCon.Execute sSQL, , adExecuteNoRecords
''''''      End If
'
'      .MoveNext
'    Loop
'
'    .Close
'  End With
'  Set rsChildren = Nothing
'
'  ' For each child table (do those nearest to the top-level first).
'  For iLoop1 = 1 To iMaxRouteLength
'    ' For each table this distance from the top-level.
'    For iLoop2 = 0 To UBound(avChildTables, 2)
'      If CInt(avChildTables(2, iLoop2)) = iLoop1 Then
'
'        ' Get the child table info.
'        sSQL = "SELECT tableName, tableID, OriginalTableName, changed, new, GrantRead, GrantNew, GrantEdit, GrantDelete, CopySecurityTableID, CopySecurityTableName" & _
'          " FROM tmpTables" & _
'          " WHERE tableID = " & Trim$(Str$(avChildTables(1, iLoop2)))
'        Set rsTables = daoDb.OpenRecordset(sSQL, _
'          dbOpenForwardOnly, dbReadOnly)
'
'        ' Set permissions differently if we are a copied table
'        If rsTables!copySecurityTableID > 0 Then
'          lngOriginalTableID = rsTables!copySecurityTableID
'          sOriginalTableName = rsTables!copySecurityTableName
'        Else
'          lngOriginalTableID = rsTables!TableID
'          sOriginalTableName = rsTables!OriginalTableName
'        End If
'
'        For Each objGroup In gObjGroups.Collection
'          sGroupName = "[" & objGroup.Name & "]"
'          sTableName = rsTables!TableName
'
'          If (objGroup.SecurityManager) Or (objGroup.SystemManager) Then
'            sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sTableName & " TO " & sGroupName
'            gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
''''''          Else
''''''            sSQL = "DENY DELETE, INSERT, SELECT, UPDATE ON " & sTableName & " TO " & sGroupName
'          End If
'
'          If Not rsTables!New Or rsTables!copySecurityTableID > 0 Then
'            fTableReadable = (objGroup.Tables.Item(sOriginalTableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
'          Else
'            ' Is a new table, or a copied one with specified permissions
'            fTableReadable = rsTables!GrantRead
'          End If
'
'          If fTableReadable Then
'
'            ' Check which parents of this child table, the current role can see.
'            iParentCount = 0
'            ' Create an array of the parents of the given table that are accessible by the given group.
'            ' Column 1 = parent type (UT = top-level table
'            '                         UV = view of a top-level table
'            '                         SV = system view)
'            ' Column 2 = parent ID
'            ' Column 3 = parent table ID
'            ' Column 4 = parent name
'            ReDim avParents(4, 0)
'
'            ' Get the parents of the current child table.
'            sSQL = "SELECT tmpTables.tableID, tmpTables.tableName, tmpTables.originalTableName, tmpTables.tableType," & _
'              " tmpTables.copySecurityTableID, tmpTables.copySecurityTableName, tmpTables.grantRead" & _
'              " FROM tmpRelations" & _
'              " INNER JOIN tmpTables ON tmpRelations.parentID = tmpTables.tableID" & _
'              " WHERE tmpRelations.childID = " & Trim$(Str$(avChildTables(1, iLoop2)))
'
'            Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'            Do While (Not rsParents.EOF)
'              If rsParents!TableType = iTabParent Then
'                ' Parent is a top-level table.
'                If objGroup.Tables.IsValid(rsParents!OriginalTableName) Then
'                  ' Table is in the tables collection.
'                  fTableOK = (objGroup.Tables(rsParents!OriginalTableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
'                Else
'                  ' Table is NOT in the tables collection. Must be new,
'                  ' so see if we're copying permissions from another table, or
'                  ' if the default permissions are specified.
'                  If rsParents!copySecurityTableID > 0 Then
'                    fTableOK = (objGroup.Tables(rsParents!copySecurityTableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
'                  Else
'                    fTableOK = rsParents!GrantRead
'                  End If
'                End If
'
'                If fTableOK Then
'                  iParentCount = iParentCount + 1
'
'                  ' The current group has permission to see all records in the parent table.
'                  iNextIndex = UBound(avParents, 2) + 1
'                  ReDim Preserve avParents(4, iNextIndex)
'                  avParents(1, iNextIndex) = "UT"
'                  avParents(2, iNextIndex) = rsParents!TableID
'                  avParents(3, iNextIndex) = rsParents!TableID
'                  avParents(4, iNextIndex) = rsParents!TableName
'                Else
'                  ' The current group does NOT have permission to see all records in the parent table.
'                  ' Get the permitted views on the table.
'                  sSQL = "SELECT tmpViews.viewID, tmpViews.viewName, tmpViews.originalViewName, tmpViews.grantRead" & _
'                    " FROM tmpViews" & _
'                    " WHERE tmpViews.viewTableID = " & Trim(Str(rsParents!TableID)) & _
'                    " AND tmpViews.deleted = FALSE"
'                  iOKViewCount = 0
'
'                  Set rsViews = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'                  Do While Not rsViews.EOF
'                    If objGroup.Views.IsValid(rsViews!OriginalViewName) Then
'                      ' View is in the views collection.
'                      fViewOK = (objGroup.Views(rsViews!OriginalViewName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
'                    Else
'                      ' View is NOT in the views collection. Must be new,
'                      ' so get the default permissions.
'                      fViewOK = rsViews!GrantRead
'                    End If
'
'                    If fViewOK Then
'                      iOKViewCount = iOKViewCount + 1
'
'                      iNextIndex = UBound(avParents, 2) + 1
'                      ReDim Preserve avParents(4, iNextIndex)
'                      avParents(1, iNextIndex) = "UV"
'                      avParents(2, iNextIndex) = rsViews!ViewID
'                      avParents(3, iNextIndex) = rsParents!TableID
'                      avParents(4, iNextIndex) = rsViews!ViewName
'                    End If
'
'                    rsViews.MoveNext
'                  Loop
'
'                  rsViews.Close
'                  Set rsViews = Nothing
'
'                  If iOKViewCount > 0 Then
'                    iParentCount = iParentCount + 1
'                  End If
'                End If
'              ElseIf rsParents!TableType = iTabChild Then
'                ' Parent is not a top-level table.
'                lngParentViewID = 0
'                sSQL = "SELECT childViewID FROM ASRSysChildViews2 WHERE role = '" & objGroup.Name & "' AND tableID = " & Trim(Str(rsParents!TableID))
'                rsChildViews.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'                Do While (Not rsChildViews.EOF) And (lngParentViewID = 0)
'                  lngParentViewID = rsChildViews!childViewID
'
'                  ' Check if it really exists.
'                  sTemp = Left("ASRSysCV" & Trim$(Str$(lngParentViewID)) & "#" & Replace(rsParents!TableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)
'                  sSQL = "SELECT COUNT(Name) AS result FROM sysobjects WHERE name = '" & sTemp & "' AND xtype = 'V'"
'                  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'                  If rsInfo.Fields(0).Value = 0 Then
'                    lngParentViewID = 0
'                  End If
'
'                  rsInfo.Close
'
'                  rsChildViews.MoveNext
'                Loop
'
'                rsChildViews.Close
'
'                If lngParentViewID > 0 Then
'                  iParentCount = iParentCount + 1
'
'                  iNextIndex = UBound(avParents, 2) + 1
'                  ReDim Preserve avParents(4, iNextIndex)
'                  avParents(1, iNextIndex) = "SV"
'                  avParents(2, iNextIndex) = lngParentViewID
'                  avParents(3, iNextIndex) = rsParents!TableID
'                  avParents(4, iNextIndex) = Left("ASRSysCV" & Trim$(Str$(lngParentViewID)) & "#" & Replace(rsParents!TableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)
'                End If
'              End If
'
'              rsParents.MoveNext
'            Loop
'
'            rsParents.Close
'            Set rsParents = Nothing
'
'            If iParentCount > 0 Then
'              iParentJoinType = 0
'
'              ' More than 1 parent. Do we want the OR join child view, or the AND join child view ?
'              If objGroup.Tables.IsValid(rsTables!OriginalTableName) Then
'                iParentJoinType = objGroup.Tables(rsTables!OriginalTableName).ParentJoinType
'              End If
'
'              ' Enter the view definition in the ASRSysChildView2 table.
'              Set cmdChildView = New ADODB.Command
'              With cmdChildView
'                .CommandText = "dbo.sp_ASRInsertChildView2"
'                .CommandType = adCmdStoredProc
'                .CommandTimeout = 0
'                Set .ActiveConnection = gADOCon
'
'                Set pmADO = .CreateParameter("Result", adInteger, adParamOutput)
'                .Parameters.Append pmADO
'
'                Set pmADO = .CreateParameter("TableID", adInteger, adParamInput)
'                .Parameters.Append pmADO
'                pmADO.Value = rsTables!TableID
'
'                Set pmADO = .CreateParameter("JoinType", adInteger, adParamInput)
'                .Parameters.Append pmADO
'                pmADO.Value = iParentJoinType
'
'                Set pmADO = .CreateParameter("Name", adVarChar, adParamInput, 256)
'                .Parameters.Append pmADO
'                pmADO.Value = objGroup.Name
'
'                .Execute
'
'                lngViewID = IIf(IsNull(.Parameters(0).Value), vbNullString, .Parameters(0).Value)
'              End With
'              Set cmdChildView = Nothing
'
'
'              sCreatedChildViews.Append "," & Trim$(Str$(lngViewID))
'
'              ' Delete the existing entries in the ASRSysChildViewParents2 table.
'              sSQL = "DELETE FROM ASRSysChildViewParents2 WHERE childViewID = " & Trim$(Str$(lngViewID))
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'              For iNextIndex = 1 To UBound(avParents, 2)
'                sSQL = "INSERT INTO ASRSysChildViewParents2" & _
'                  " (childViewID, parentType, parentID, parentTableID)" & _
'                  " VALUES (" & Trim$(Str$(lngViewID)) & ", " & _
'                  "'" & avParents(1, iNextIndex) & "', " & _
'                  Trim$(Str$(avParents(2, iNextIndex))) & ", " & _
'                  Trim$(Str$(avParents(3, iNextIndex))) & ")"
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              Next iNextIndex
'
'              ' Create the view name.
'              sViewName = Left$("ASRSysCV" & Trim$(Str$(lngViewID)) & "#" & Replace(sTableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)
'
'              ' Create the view
'              sSQL = "CREATE VIEW dbo." & sViewName & vbNewLine & _
'                "AS" & vbNewLine & _
'                "        SELECT " & sTableName & ".*" & vbNewLine & _
'                "        FROM " & sTableName & vbNewLine & _
'                "        WHERE " & vbNewLine & _
'                "                (" & vbNewLine & _
'                "                        " & _
'                sTableName & ".ID_" & Trim$(Str$(avParents(3, 1))) & " IN (SELECT id FROM " & avParents(4, 1) & ")" & vbNewLine
'
'              lngLastParentID = avParents(3, 1)
'
'              strParentIDs = "ID_" & lngLastParentID
'              For iLoop = 2 To UBound(avParents, 2)
'
'                If lngLastParentID <> avParents(3, iLoop) Then
'                  lngLastParentID = avParents(3, iLoop)
'                  strParentIDs = strParentIDs & IIf(LenB(strParentIDs) <> 0, ",", "") & "ID_" & lngLastParentID
'
'                  sSQL = sSQL & _
'                    "                )" & vbNewLine & _
'                    "                " & IIf(iParentJoinType = 1, "AND", "OR") & vbNewLine & _
'                    "                (" & vbNewLine
'                Else
'                  sSQL = sSQL & _
'                    "                        OR" & vbNewLine
'                End If
'
'                sSQL = sSQL & "                        " & _
'                  sTableName & ".ID_" & Trim$(Str$(avParents(3, iLoop))) & " IN (SELECT id FROM " & avParents(4, iLoop) & ")" & vbNewLine
'              Next iLoop
'
'              sSQL = sSQL & "                )"
'
'              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'              If LenB(sSysSecRoles) <> 0 Then
'                sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sViewName & " TO " & sSysSecRoles
'                sAllSQLCommands = sAllSQLCommands & vbNewLine & sSQL
'                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'              End If
'
''''''              If Len(sNonSysSecRoles) > 0 Then
''''''                sSQL = "DENY DELETE, INSERT, SELECT, UPDATE ON " & sViewName & " TO " & sNonSysSecRoles
''''''                sAllSQLCommands = sAllSQLCommands & vbNewLine & sSQL
''''''                gADOCon.Execute sSQL, , adExecuteNoRecords
''''''              End If
'
'              If (Not objGroup.SecurityManager) And (Not objGroup.SystemManager) Then
'                ' Apply the configured permissions to the child view permutation.
'                ' Initialise the Table Column permissions command strings.
'                sSelectGrant = "ID,Timestamp," & strParentIDs
'                sUpdateGrant = "ID,Timestamp," & strParentIDs
'
'
'                If (Not rsTables!New) Or (rsTables!copySecurityTableID > 0) Then
'
'                  With objGroup.Tables.Item(sOriginalTableName)
'                    If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Or .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
'
'                      ' Get the set of non-system columns in the table.
'                      sSQL = "SELECT columnName, OriginalColumnName, new" & _
'                        " FROM tmpColumns" & _
'                        " WHERE tableID = " & Trim$(Str$(lngOriginalTableID)) & _
'                        " AND deleted = FALSE" & _
'                        " AND columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
'                        " AND columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK))
'                      Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'                      Do While Not rsColumns.EOF
'
'                        sColumnName = rsColumns.Fields(0).Value
'                        sOriginalColumnName = Trim(UCase(IIf(IsNull(rsColumns.Fields(1).Value), "", rsColumns.Fields(1).Value)))
'
'                        If rsColumns.Fields(2).Value Then ' Is New
'
'                          ' Build string of columns that are allowed
'                          If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
'                            sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & sColumnName
'                          End If
'
'                          If .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
'                            sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & sColumnName
'                          End If
'
'                        Else
'
'                          ' Build string of columns that are revoked based on what they had before
'                          If .Columns.Item(sOriginalColumnName).SelectPrivilege Or .TableType = iTabLookup Then
'                            sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & sColumnName
'                          End If
'
'                          If .Columns.Item(sOriginalColumnName).UpdatePrivilege Then
'                            sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & sColumnName
'                          End If
'                        End If
'
'                        rsColumns.MoveNext
'                      Loop
'
'                      rsColumns.Close
'                      Set rsColumns = Nothing
'                    End If
'
'
'                    ' Delete permissions
'                    If .DeletePrivilege Then
'                      sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
'                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                    End If
'
'                    ' Insert permissions
'                    If .InsertPrivilege Then
'                      sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
'                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                    End If
'
'                    ' Select permissions
'                    If .SelectPrivilege = giPRIVILEGES_ALLGRANTED Or (.TableType = iTabLookup) Then
'                      sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
'                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                    ElseIf .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
'                      If LenB(sSelectGrant) <> 0 Then
'                        'MH20060620 Fault 11186
'                        'sSQL = "GRANT SELECT(" & sSelectGrant & ") ON " & sViewName & " TO " & sGroupName
'                        sSQL = "GRANT SELECT(ID,Timestamp," & sSelectGrant & ") ON " & sViewName & " TO " & sGroupName
'                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                      End If
'                    End If
'
'                    ' Update permissions
'                    If .UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
'                      sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
'                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                    ElseIf .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
'                      If LenB(sUpdateGrant) <> 0 Then
'                        'MH20060620 Fault 11186
'                        'sSQL = "GRANT UPDATE(" & sUpdateGrant & ") ON " & sViewName & " TO " & sGroupName
'                        sSQL = "GRANT UPDATE(ID,Timestamp," & sUpdateGrant & ") ON " & sViewName & " TO " & sGroupName
'                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                      End If
'                    End If
'
'                  End With
'                Else
'                  ' Is a new table, or a copied one with specified permissions
'                  If rsTables!GrantRead Then
'                    sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
'                    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                  End If
'
'                  If rsTables!GrantEdit Then
'                    sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
'                    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                  End If
'
'                  If rsTables!GrantNew Then
'                    sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
'                    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                  End If
'
'                  If rsTables!GrantDelete Then
'                    sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
'                    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'                  End If
'
'                End If
'              End If
'            End If
'          End If
'        Next objGroup
'        Set objGroup = Nothing
'
'        rsTables.Close
'        Set rsTables = Nothing
'      End If
'    Next iLoop2
'  Next iLoop1
'
'  ' Delete invalid records from ASRSysChildViews2 and ASRSysChildViewParents2
'  sSQL = "DELETE FROM ASRSysChildViews2 WHERE childViewID NOT IN (" & sCreatedChildViews.ToString & ")"
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  sSQL = "DELETE FROM ASRSysChildViewParents2 WHERE childViewID NOT IN (" & sCreatedChildViews.ToString & ")"
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Drop all existing old style child views.
'  sSQL = "SELECT name FROM sysobjects WHERE name LIKE 'ASRSysChildView[_]%' AND xtype = 'V'"
'  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'  With rsInfo
'    Do While (Not .EOF)
'      sSQL = "DROP VIEW " & .Fields(0).Value
'      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'      .MoveNext
'    Loop
'
'    .Close
'  End With
'
'  ' Drop the old style child view tables.
'  sSQL = "SELECT COUNT(Name) AS result FROM sysobjects WHERE name = 'ASRSysChildViews' AND xtype = 'U'"
'  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'  If rsInfo!result > 0 Then
'    sSQL = "DROP TABLE ASRSysChildViews"
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'  End If
'  rsInfo.Close
'
'  sSQL = "SELECT COUNT(Name) AS result FROM sysobjects WHERE name = 'ASRSysChildViewParents' AND xtype = 'U'"
'  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'  If rsInfo!result > 0 Then
'    sSQL = "DROP TABLE ASRSysChildViewParents"
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'  End If
'  rsInfo.Close
'
'TidyUpAndExit:
'  Set cmdChildView = Nothing
'  Set objGroup = Nothing
'  Set rsInfo = Nothing
'  Set rsChildViews = Nothing
'
'  ApplyPermissions_ChildTables2 = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error Applying Permissions (Child Tables)"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'



Private Function vbCompiled() As Boolean
  
  'Dim nRtn As Long
  'Dim Buffer As String
  'Buffer = Space$(256)
  'nRtn = GetModuleFileNameA(0&, Buffer, Len(Buffer))
  'Buffer = UCase(Left(Buffer, nRtn))
  'vbCompiled = (Right(Buffer, 8) <> "\VB6.EXE")

  'Much better (and clever-er) !
  On Local Error Resume Next
  Err.Clear
  Debug.Print 1 / 0
  vbCompiled = (Err.Number = 0)

End Function


'{MH20000727
Public Sub SetComboItem(cboCombo As ComboBox, lItem As Long)

  Dim lCount As Long
    
  With cboCombo
    For lCount = 1 To .ListCount
      If .ItemData(lCount - 1) = lItem Then
        .ListIndex = lCount - 1
        Exit For
      End If
    Next
  End With

End Sub
'MH20000727}

Public Function GetComboItem(cboTemp As ComboBox) As Long
  
  GetComboItem = 0
  If cboTemp.ListIndex <> -1 Then
    GetComboItem = cboTemp.ItemData(cboTemp.ListIndex)
  End If
  
End Function


Public Function TimePeriod(intPeriod As TimePeriods) As String
  Select Case intPeriod
  Case iTimePeriodDays: TimePeriod = "day"
  Case iTimePeriodWeeks: TimePeriod = "week"
  Case iTimePeriodMonths: TimePeriod = "month"
  Case iTimePeriodYears: TimePeriod = "year"
  End Select
End Function


Public Sub OutputError(strError As String)

  Dim strOutput As String
  Dim fProgressVisible As Boolean


  strOutput = strError
  If Trim(Err.Description) <> vbNullString Then
    strOutput = strOutput & vbNewLine & _
                "(" & ODBC.FormatError(Err.Description) & ")"
  End If

  fProgressVisible = gobjProgress.Visible
  If fProgressVisible Then
    gobjProgress.Visible = False
  End If

  OutputCurrentProcess _
      vbNewLine & _
      "***** ERROR *****" & vbNewLine & _
      strOutput & vbNewLine & _
      "*****************"

  CheckIfNeedToReconnect

  Screen.MousePointer = vbDefault
  MsgBox strOutput, vbCritical + vbOKOnly, App.Title
  gobjProgress.Visible = fProgressVisible
  'fOK = False

End Sub


Private Sub CheckIfNeedToReconnect()

  Dim sConnect As String

  On Local Error Resume Next

  gADOCon.Execute "SELECT 'Testing Connection...'"

  If Err.Number = -2147467259 Then
    
    sConnect = gADOCon.ConnectionString
    
    gADOCon.RollbackTrans
    gADOCon.Close
    Set gADOCon = Nothing
  
    Set gADOCon = New ADODB.Connection
    With gADOCon
      .ConnectionString = sConnect
      .Provider = "SQLOLEDB"
      .CommandTimeout = 0
      .ConnectionTimeout = 0
      .CursorLocation = adUseServer
      .Mode = adModeReadWrite
      .Properties("Packet Size") = 32767
      .Open
    End With
    
    UnlockDatabase lckSaving
    UnlockDatabase lckReadWrite
    
    LockDatabase lckReadWrite
    LockDatabase lckSaving
    
    gADOCon.BeginTrans
    
  End If

End Sub


Public Function OutputMessage(pstrError As String) As Integer
  
  Dim fProgressVisible As Boolean
  Dim iAnswer As Integer
  Dim sCaption As String

  fProgressVisible = gobjProgress.Visible
  If fProgressVisible Then
    gobjProgress.Visible = False
    sCaption = gobjProgress.Bar1Caption
  End If

  Screen.MousePointer = vbDefault
  iAnswer = MsgBox(pstrError, vbQuestion + vbYesNo, App.Title)

  OutputCurrentProcess _
      vbNewLine & _
      "***** PROMPT *****" & vbNewLine & _
      pstrError & vbNewLine & _
      IIf(iAnswer = vbYes, "CONTINUE SAVING", "STOP SAVING") & vbNewLine & _
      "******************"

  If fProgressVisible And (iAnswer = vbYes) Then
    gobjProgress.Bar1Caption = sCaption
    gobjProgress.Visible = True
  End If

  OutputMessage = iAnswer
  
End Function



Public Sub OutputCurrentProcess(strInput As String, Optional blnOverwriteExisting As Boolean)

  'Ignore any errors in here...
  On Local Error GoTo LocalErr
  
  Dim strFileName As String

  If Trim$(strInput) <> vbNullString Then
    gobjProgress.Bar1Caption = Trim$(strInput) & " ..."
    gobjProgress.Bar2MaxValue = 100000
    gobjProgress.Bar2Value = 0
    gobjProgress.ResetBar2
    gobjProgress.UpdateProgress2
    gobjProgress.Bar2Value = 0
  End If

  strFileName = gsLogDirectory & "\savelog.txt"

  If blnOverwriteExisting Then
    If Dir(strFileName) <> vbNullString Then
      Kill strFileName
    End If
    
    Open strFileName For Append As #99
    Print #99, "Server    : " & gsServerName
    Print #99, "Database  : " & gsDatabaseName
    Print #99, "Username  : " & gsUserName
    Print #99, "Version   : " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    Print #99, vbNullString

  Else
    Open strFileName For Append As #99

  End If
  
  Print #99, Now & "  " & strInput
  Close #99

LocalErr:
  Err.Clear

End Sub



Public Sub ControlsDisableAll(objCurrent As Object, Optional blnEnabled As Boolean = False)

  Dim ctl As Control

  'Not all controls have a backcolor !
  On Local Error Resume Next

  If TypeOf objCurrent Is Form Then
    For Each ctl In objCurrent
      
      EnableControl ctl, blnEnabled
    Next

  Else
    EnableControl objCurrent, blnEnabled
    For Each ctl In objCurrent.Parent
      If ctl.Container.Name = objCurrent.Name Then
        EnableControl ctl, blnEnabled
      End If
    Next

  End If

End Sub

Public Function EnableControl(ctl As Control, blnEnabled As Boolean)

  ' JPD20020920 Added ActiveBar to the list of controls that are not enabled/disabled, as it has no
  ' 'enabled' property. Instead use the 'EnableActiveBar' method.
  
  'If Left(ctl.Name, 3) = "opt" Then
  '  Stop
  'End If
  
  
  If TypeOf ctl Is MSComctlLib.TabStrip Or _
         TypeOf ctl Is ComctlLib.TabStrip Or _
         TypeOf ctl Is ActiveBar Or _
         TypeOf ctl Is SSTab Then
    'Stop

  ElseIf TypeOf ctl Is Frame Or _
         TypeOf ctl Is PictureBox Or _
         TypeOf ctl Is MSComctlLib.ListView Or _
         TypeOf ctl Is ComctlLib.ListView Or _
         TypeOf ctl Is ListBox Or _
         TypeOf ctl Is Label Then
         
    'Just make container controls and scroll-able controls look disabled...
    '(NOTE: Code must be placed in drag-drop events etc. to disable it)
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    If (TypeOf ctl Is ListBox) Or _
        (TypeOf ctl Is MSComctlLib.ListView) Or _
        (TypeOf ctl Is ComctlLib.ListView) Then
        
      ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)    'SSDBGrid
    End If

    'TM20010822 Fault 2566
    'Controls of this type need to be enabled so that scroll/click functionality
    'is allowed.
    ctl.Enabled = True

  ElseIf TypeOf ctl Is SSDBGrid Then

    ctl.Enabled = blnEnabled
    ctl.BackColorEven = IIf(blnEnabled, vbWindowBackground, vbButtonFace)   'SSDBGrid
    ctl.BackColorOdd = IIf(blnEnabled, vbWindowBackground, vbButtonFace)   'SSDBGrid

  ElseIf TypeOf ctl Is CommandButton Then
    'Disable all CommandButtons except cancel...

    If ctl.Cancel = False Then
      ctl.Enabled = blnEnabled
    Else
      ctl.Enabled = True
    End If

  ElseIf (TypeOf ctl Is CheckBox) Or (TypeOf ctl Is OptionButton) Then

    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = vbButtonFace
    ctl.Enabled = blnEnabled

  ElseIf (TypeOf ctl Is UpDown) Then
    ctl.Enabled = blnEnabled

  ElseIf (TypeOf ctl Is TextBox) Then
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
    ctl.Locked = Not blnEnabled
    ctl.TabStop = blnEnabled
    ctl.Enabled = blnEnabled
  
  ElseIf (TypeOf ctl Is RichTextBox) Then
    ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
    ctl.Locked = Not blnEnabled
    ctl.TabStop = blnEnabled
    ctl.Enabled = blnEnabled
  
  Else
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
    ctl.Enabled = blnEnabled
  
  End If

End Function


'Public Function SaveChanges_LogoutCheck(blnSendMessageVisible As Boolean) As Boolean
'
'  Dim frmViewUsers As frmViewCurrentUsers
'  Dim blnCancelled As Boolean
'  Dim fOK As Boolean
'
'  Set frmViewUsers = New frmViewCurrentUsers
'  With frmViewUsers
'
'    fOK = .OkayToSave
'    If Not fOK And .grdUsers.Rows > 0 Then
'
'      gobjProgress.Visible = False
'      Screen.MousePointer = vbDefault
'
'      'NHRD20030425 Fault 4880 Added a check to determine 'which direction' theyre coming from.
'      'i.e. if they are doing a quick update then they wouldn't be logged in yet. Conversely if
'      'they are saving changes then they would already be logged in and I can engineer the message accordingly.
'      If Application.LoggedIn() Then
'        MsgBox "Making changes to the database will affect users who are currently logged into the system." & vbNewLine & vbNewLine & _
'               "You will need to ensure that all users are logged out and that you have locked the system " & vbNewLine & _
'               "before you can apply these changes.", vbInformation, "Saving Changes"
'      Else
'        MsgBox "Updating the system will affect users who are currently logged in." & vbNewLine & vbNewLine & _
'               "You will need to ensure that all users are logged out " & _
'               "before you can run the update process.", vbInformation, "Updating System"
'      End If
'
'      If .OkayToSave = False Then
'        .Enabled = True
'        .Saving = True
'        .SendMessageVisible = blnSendMessageVisible
'        .Show vbModal
'      End If
'      Screen.MousePointer = vbHourglass
'      fOK = Not .Cancelled
'    End If
'
'    SaveChanges_LogoutCheck = fOK
'
'  End With
'  UnLoad frmViewUsers
'  Set frmViewUsers = Nothing
'
'End Function
Private Function ConfigureHierarchySpecifics() As Boolean
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  fOK = True
  
  If Application.PersonnelModule And _
    gbEnableUDFFunctions Then
    
    fOK = modHierarchySpecifics.ConfigureHierarchySpecifics
  End If

TidyUpAndExit:
  ConfigureHierarchySpecifics = fOK
  Exit Function

ErrorTrap:
  OutputError "Error performing quick checks 3"
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function ValidateGTMaskDate(dtTemp As GTMaskDate.GTMaskDate) As Boolean

  Dim blnYearOkay As Boolean
  Dim sSysDateSeparator As String

  ValidateGTMaskDate = True

  sSysDateSeparator = UI.GetSystemDateSeparator
  
  With dtTemp
    If Trim(Replace(.Text, sSysDateSeparator, vbNullString)) <> vbNullString Then
  
      'MH20020423 Fault 3760 (Avoid changing 01/13/2002 to 13/01/2002)
      'If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Then
      'If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Or Left(.Text, 5) <> Left(.DateValue, 5) Then
      
      'MH20020423 Fault 3543 Also make sure that they enter a valid year
      blnYearOkay = (val(Mid(.Text, InStrRev(.Text, sSysDateSeparator) + 1)) >= 1753)
      
      If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Or _
          Format(.DateValue, DateFormat) <> .Text Or Not blnYearOkay Then

        Clipboard.Clear
        Clipboard.SetText .Text
        .DateValue = Null
        .Paste
        DoEvents
  
        .ForeColor = vbRed
        MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
        .ForeColor = vbWindowText
        .DateValue = Null
        If .Visible And .Enabled Then
          .SetFocus
        End If
        ValidateGTMaskDate = False
  
      End If
    End If
  End With

End Function

' Are the two passed in values within the given difference
Public Function IsWithin(plngValue1, plngValue2, plngDifference) As Boolean

  IsWithin = IIf(Abs(plngValue1 - plngValue2) > plngDifference Or Abs(plngValue1 - plngValue2) < plngDifference, False, True)

End Function

Public Function EnableUDFFunctions() As Boolean
  
  Dim sSQL As String
  Dim rsResult As New ADODB.Recordset

  sSQL = "exec sp_server_info 500"
  rsResult.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  Select Case val(rsResult(2))
    Case Is >= 8
      EnableUDFFunctions = True
    Case Else
      EnableUDFFunctions = False
  End Select

  rsResult.Close
  Set rsResult = Nothing

End Function

'Private Function CleanupDatabase() As Boolean
'
'  ' Clean any temporary stored procedures/functions/udfs that are lying around in the database
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  ' Delete the existing order definition from the server database.
'  sSQL = "EXEC spASRDropTempObjects"
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'  ' Clear out any junk that may be laying round in the messages table
'  sSQL = "DELETE FROM ASRSysMessages"
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'TidyUpAndExit:
'  CleanupDatabase = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error cleaning database"
'  Resume TidyUpAndExit
'
'End Function


Public Function GetTableName(lngTableID As Long) As String

  On Error GoTo ErrorTrap
  
  GetTableName = vbNullString
  
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", lngTableID
      
    If Not .NoMatch Then
      GetTableName = !TableName
    End If
  End With
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Public Function GetViewName(lngViewID As Long) As String

  On Error GoTo ErrorTrap
  
  GetViewName = vbNullString
  
  With recViewEdit
    .Index = "idxViewID"
    .Seek "=", lngViewID
      
    If Not .NoMatch Then
      GetViewName = !viewName
    End If
  End With
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Public Function GetColumnName(lngColumnID As Long, Optional pbJustColumnName As Boolean = False) As String
  
  On Error GoTo ErrorTrap
  
  GetColumnName = vbNullString
  
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", lngColumnID

    If Not .NoMatch Then
      GetColumnName = !ColumnName
    
      If pbJustColumnName = False Then
        With recTabEdit
          .Index = "idxTableID"
          .Seek "=", recColEdit!TableID
        
          If Not .NoMatch Then
            GetColumnName = Trim(!TableName) & "." & GetColumnName
          End If
        End With
      End If
    
    End If
  End With
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Public Function GetDataTypeName(piDataType As SQLDataType) As String
  
  On Error GoTo ErrorTrap
  
  GetDataTypeName = vbNullString
  
  Select Case piDataType
    Case sqlUnknown        ' ?
      GetDataTypeName = "Unknown"
    Case sqlOle            ' OLE columns
      GetDataTypeName = "OLE object"
    Case sqlBoolean        ' Logic columns
      GetDataTypeName = "Logic"
    Case sqlNumeric        ' Numeric columns
      GetDataTypeName = "Numeric"
    Case sqlInteger        ' Integer columns
      GetDataTypeName = "Integer"
    Case sqlDate           ' Date columns
      GetDataTypeName = "Date"
    Case sqlVarChar        ' Character columns
      GetDataTypeName = "Character"
    Case sqlVarBinary      ' Photo columns
      GetDataTypeName = "Photo"
    Case sqlLongVarChar    ' Working Pattern columns
      GetDataTypeName = "Working Pattern"
  End Select

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function


Public Function GetColumnDataType(lngColumnID As Long) As DataTypes
  
  On Error GoTo ErrorTrap
  
  GetColumnDataType = 0
  
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", lngColumnID

    If Not .NoMatch Then
      GetColumnDataType = !DataType
    End If
  End With
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Public Function GetColumnOLEType(lngColumnID As Long) As OLEType
  
  On Error GoTo ErrorTrap
  
  GetColumnOLEType = OLE_LOCAL
  
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", lngColumnID

    If Not .NoMatch Then
      GetColumnOLEType = !OLEType
    End If
  End With
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function


Public Function GetColumnSize(lngColumnID As Long, pbDecimals As Boolean) As Integer
  
  On Error GoTo ErrorTrap
  
  GetColumnSize = 0
  
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", lngColumnID

    If Not .NoMatch Then
      If pbDecimals Then
        GetColumnSize = !Decimals
      Else
        GetColumnSize = !Size
      End If
    End If
  End With
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Public Function CalculateBottomOfPage() As Long
  With Printer
    CalculateBottomOfPage = .ScaleHeight - (giPRINT_YINDENT)
  End With
End Function

Public Function CheckEndOfPage2(Optional mlngBottom As Long, Optional fReset As Boolean)
  If Printer.CurrentY > mlngBottom Then
    Call FooterText2
    Printer.NewPage
    
    If fReset Then glngPageNum = 0
    
    Printer.CurrentY = giPRINT_YINDENT
    Printer.CurrentX = giPRINT_XINDENT
  End If
End Function

Public Function FooterText2()
  
  Dim strPageNum As String
  
  glngPageNum = glngPageNum + 1
  strPageNum = "Page " & CStr(glngPageNum)

  Printer.FontSize = 8
  Printer.Print " "
  Printer.FontBold = False
  Printer.FontUnderline = False
  Printer.FontStrikethru = False
  
  Printer.CurrentX = giPRINT_XINDENT
  Printer.Print "Printed on " & Format(Now, DateFormat) & _
                " at " & Format(Now, "hh:nn") & " by " & gsUserName;
  
  Printer.CurrentX = (Printer.ScaleWidth - giPRINT_XINDENT) - Printer.TextWidth(strPageNum)
  Printer.Print strPageNum

  Printer.FontSize = 10

End Function


Public Sub RefreshAllControls(objContainer As Object)

  Dim ctl As Control

  On Local Error Resume Next

  If TypeOf objContainer Is Form Then
    For Each ctl In objContainer
      ctl.Refresh
    Next
  Else
    For Each ctl In objContainer.Parent
      If ctl.Container.Name = objContainer.Name Then
        ctl.Refresh
      End If
    Next
  End If
  
  objContainer.Refresh

End Sub


Public Function GetExpressionName(lngExprID As Long) As String

  GetExpressionName = vbNullString
  
  If lngExprID > 0 Then
  
    With recExprEdit
      .Index = "idxExprID"
      .Seek "=", lngExprID, False
    
      If Not .NoMatch Then
        GetExpressionName = Trim(!Name)
      End If
      
    End With
  End If

End Function


Public Function GetOrderName(plngOrderID As Long) As String
  On Error GoTo ErrorTrap
  
  Dim objOrder As Order

  GetOrderName = vbNullString
  
  If plngOrderID > 0 Then
    Set objOrder = New Order
    With objOrder
      .OrderID = plngOrderID
      
      ' Read the name of the current order.
      If .ConstructOrder Then
        GetOrderName = .OrderName
      End If
    End With
    Set objOrder = Nothing
  End If

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  GetOrderName = vbNullString
  Resume TidyUpAndExit
  
End Function



Public Function TimeFormat(dtInput As Date, strFormat As String) As String

  Dim lngIndex As Long
  
  TimeFormat = Format(dtInput, strFormat)

  For lngIndex = 1 To Len(strFormat)
    If Mid(strFormat, lngIndex, 1) = ":" Then
      Mid(TimeFormat, lngIndex, 1) = ":"
    End If
  Next

End Function


Public Function GetTableIDFromColumnID(lngColumnID As Long) As Long
  
  On Error GoTo ErrorTrap
  
  GetTableIDFromColumnID = 0
  
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", lngColumnID

    If Not .NoMatch Then
      GetTableIDFromColumnID = !TableID
    End If
  End With
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Public Sub OutputCurrentProcess2(strInput As String, Optional ByVal lngMaxSteps As Long)

  'Ignore any errors in here...
  On Local Error GoTo LocalErr
  
  Dim strFileName As String

  If Trim$(strInput) <> vbNullString Then
    gobjProgress.Bar2Caption = strInput & " ..."
  End If

  If lngMaxSteps > 0 Then
    gobjProgress.Bar2MaxValue = lngMaxSteps
  End If

  strFileName = gsLogDirectory & "\savelog.txt"

  If Trim$(strInput) <> vbNullString Then
    Open strFileName For Append As #99
    Print #99, Now & "    " & strInput
    Close #99
  End If

LocalErr:
  Err.Clear

End Sub


Function GetNextIdentitySeed(ByVal psTableName As String) As Long

  Dim sSQL As String
  Dim sErrorMessage As String
  Dim lngNewSeed As Long
  Dim iCharEnd As String
   
  sSQL = "DBCC CHECKIDENT ([" & psTableName & "], NORESEED)"
  gADOCon.Execute sSQL, -1, adExecuteNoRecords
  
  sErrorMessage = gADOCon.Errors(0).Description
  
  If InStr(1, UCase$(sErrorMessage), "NULL") Then
    lngNewSeed = 0
  Else
    If InStr(1, sErrorMessage, "Checking identity information: current identity value '") Then
      iCharEnd = InStr(1, sErrorMessage, "', current column value")
      lngNewSeed = Mid$(sErrorMessage, 56, iCharEnd - 56)
    End If
  End If


  gADOCon.Errors.Clear
  GetNextIdentitySeed = lngNewSeed

End Function

Public Function GetTagKeyFromCollection(pcolSSITableViews As clsSSITableViews, psTableViewName As String) As String
  
  Dim oSSITableView As clsSSITableView
  
  For Each oSSITableView In pcolSSITableViews.Collection
      
    If oSSITableView.TableViewName = psTableViewName Then
      GetTagKeyFromCollection = CreateTableViewTag(oSSITableView.TableID, oSSITableView.ViewID)
      Exit Function
    End If
  
  Next oSSITableView
  
End Function

Public Function GetTableIDFromCollection(pcolSSITableViews As clsSSITableViews, psTableViewName As String) As Long
  
  Dim oSSITableView As clsSSITableView
  
  For Each oSSITableView In pcolSSITableViews.Collection
      
    If oSSITableView.TableViewName = psTableViewName Then
      GetTableIDFromCollection = oSSITableView.TableID
      Exit Function
    End If
  
  Next oSSITableView
  
End Function

Public Function GetViewIDFromCollection(pcolSSITableViews As clsSSITableViews, psTableViewName As String) As Long
  
  Dim oSSITableView As clsSSITableView
  
  For Each oSSITableView In pcolSSITableViews.Collection
      
    If oSSITableView.TableViewName = psTableViewName Then
      GetViewIDFromCollection = oSSITableView.ViewID
      Exit Function
    End If
  
  Next oSSITableView
  
End Function

Public Function CreateTableViewTag(psTableID As String, psViewID As String) As String
   CreateTableViewTag = (psTableID & "_" & IIf(psViewID = vbNullString, "-1", psViewID))
End Function

Public Function CreateTableViewName(psTableName As String, psViewName As String) As String
  CreateTableViewName = psTableName & IIf(Len(psViewName) > 0, " (" & psViewName & " view)", vbNullString)
End Function

Public Function DecodeTag(psTag As String, pfView As Boolean) As String

  If pfView Then
    DecodeTag = Mid(psTag, InStr(1, psTag, "_") + 1)
  Else
    DecodeTag = Mid(psTag, 1, InStr(1, psTag, "_") - 1)
  End If

End Function

Public Function GetModuleSetupValue(sModuleKey As String, sParameterKey As String, strType As String) As String
    
  With recModuleSetup
    .Index = "idxModuleParameter"
  
    ' Get the Dependants table ID.
    GetModuleSetupValue = vbNullString
    .Seek "=", sModuleKey, sParameterKey
    If Not .NoMatch Then
      If Not IsNull(!parametervalue) Then
        Select Case strType
        Case "T"
          GetModuleSetupValue = GetTableName(val(!parametervalue))
        Case "C"
          GetModuleSetupValue = GetColumnName(val(!parametervalue), False)
        Case "ColumnNameOnly"
          GetModuleSetupValue = GetColumnName(val(!parametervalue), True)
        Case Else
          GetModuleSetupValue = !parametervalue
        End Select
      End If
    End If

  End With

End Function


Public Function GetClone(pavCloneRegister As Variant, strType As String, lngOldID As Long) As Long

  Dim lngNewID As Long
  Dim iIndex As Long
  
  lngNewID = lngOldID
  If lngOldID > 0 Then
    For iIndex = 1 To UBound(pavCloneRegister, 2)
      If pavCloneRegister(1, iIndex) = strType And pavCloneRegister(2, iIndex) = lngOldID Then
        lngNewID = pavCloneRegister(3, iIndex)
        Exit For
      End If
    Next
  End If

  GetClone = lngNewID

End Function

Public Function FormatGUID(ByRef GUID As Variant) As String

  If IsNull(GUID) Then
    FormatGUID = "NULL"
  Else
    FormatGUID = Replace(GUID, "{guid {", "")
    FormatGUID = Replace(FormatGUID, "}}", "")
  End If

End Function

Public Function FormatTableName(ByVal Owner As String, TableName As String) As String

  Dim strReturn As String
  strReturn = "[" & Owner & "].[" & TableName & "]"
  FormatTableName = strReturn

End Function

Public Sub GetObjectCategories(ByRef theCombo As ComboBox, UtilityType As UtilityType, UtilityID As Long, Optional TableID As Long)

  On Error GoTo ErrorTrap

  Dim rsTemp As New ADODB.Recordset
  Dim iListIndex As Integer
  
  ' Add <none>
  theCombo.AddItem "<None>"
  theCombo.ItemData(theCombo.NewIndex) = 0
  iListIndex = theCombo.NewIndex
          
  rsTemp.Open "EXEC dbo.spsys_getobjectcategories " & CStr(utlScreen) & ", " & CStr(UtilityID) & ", " & CStr(TableID) _
      , gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
      theCombo.AddItem rsTemp.Fields("category_name").value
      theCombo.ItemData(theCombo.NewIndex) = rsTemp.Fields("ID").value
      
      If rsTemp.Fields("Selected").value = 1 Then
        iListIndex = theCombo.NewIndex
      End If
      rsTemp.MoveNext
    Loop
  End If
  
  theCombo.Enabled = (theCombo.ListCount > 0)
    
  If iListIndex > -1 And UtilityID > 0 Then
    theCombo.ListIndex = iListIndex
  End If
  
TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Sub
  
ErrorTrap:
  GoTo TidyUpAndExit

End Sub

Public Function IsModuleEnabled(lngModuleCode As enum_Module) As Boolean
  IsModuleEnabled = (gobjLicence.Modules And lngModuleCode)
End Function


Public Function CheckLicence() As Boolean

  Dim sMsg As String
  Dim lActualCount As Long
  Dim lngCurrentHeadcount As Long
  Dim dToday As Date
  
  On Error GoTo Err_Trap
  
  CheckLicence = False
  dToday = DateValue(Now)
    
  ' Expiry date checks
  If gobjLicence.HasExpiryDate Then
    If (dToday > gobjLicence.ExpiryDate) Then
      sMsg = "Your licence to use this product has expired." & vbNewLine & _
            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible."
      gbLicenceExpired = True
      MsgBox sMsg, vbInformation
      GoTo Exit_Ok:
    End If
            
    If (dToday > DateAdd("d", -7, gobjLicence.ExpiryDate)) Then
      sMsg = "Your licence to use this product will expire on " & gobjLicence.ExpiryDate & "." & vbNewLine & vbNewLine & _
            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible."
      MsgBox sMsg, vbInformation
    End If
  End If
    
     
  ' Headcount checks
  If gobjLicence.Headcount > 0 Then
    Select Case gobjLicence.LicenceType
      Case LicenceType.Headcount, LicenceType.DMIConcurrencyAndHeadcount
        lngCurrentHeadcount = GetSystemSetting("Headcount", "current", 0)
        
      Case LicenceType.P14Headcount, LicenceType.DMIConcurrencyAndP14
       lngCurrentHeadcount = GetSystemSetting("Headcount", "P14", 0)

    End Select
   
    If lngCurrentHeadcount >= gobjLicence.Headcount Then
      sMsg = "You have reached or exceeded the headcount limit set within the terms of your licence agreement." & vbNewLine & vbNewLine & _
                            "You are no longer able to add new employee records, but you may access the system for other purposes." & vbNewLine & vbNewLine & _
                            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible to increase the licence headcount number."
      MsgBox sMsg, vbCritical
    
    ElseIf lngCurrentHeadcount >= gobjLicence.Headcount * 0.95 Then
      
      If DisplayWarningToUser(gsUserName, Headcount95Percent, 7) Then
        sMsg = "You are currently within 95% (" & lngCurrentHeadcount & " of " & gobjLicence.Headcount & " employees) of reaching the headcount limit set within the terms of your licence agreement." & vbNewLine & vbNewLine & _
                              "Once this limit is reached, you will no longer be able to add new employee records to the system." & vbNewLine & vbNewLine & _
                              "If you wish to increase the headcount number, please contact OpenHR Customer Services on 08451 609 999 as soon as possible."
        MsgBox sMsg, vbInformation
      End If
    
    End If
  End If
    
Exit_Ok:
  CheckLicence = True
  Exit Function
    
Exit_Fail:
  Screen.MousePointer = vbDefault
  MsgBox sMsg, vbCritical
  CheckLicence = False
  Exit Function

Err_Trap:
  CheckLicence = False

End Function

Private Function DisplayWarningToUser(userName As String, WarningType As WarningType, warningRefreshRate As Integer) As Boolean

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim bResult As Boolean

  On Error GoTo ErrorTrap
  
  ' Run the stored procedure to see if the given record has changed
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRUpdateWarningLog"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
                      
    Set pmADO = .CreateParameter("Username", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.value = userName
    
    Set pmADO = .CreateParameter("WarningType", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.value = WarningType
          
    Set pmADO = .CreateParameter("WarningRefreshRate", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.value = warningRefreshRate
          
    Set pmADO = .CreateParameter("WarnUser", adBoolean, adParamOutput)
    .Parameters.Append pmADO
          
    Set pmADO = Nothing

    cmADO.Execute
    bResult = CBool(.Parameters(3).value)
    
  End With
               
TidyUpAndExit:
  DisplayWarningToUser = bResult
  Exit Function

ErrorTrap:
  bResult = False
  GoTo TidyUpAndExit

End Function

Public Function CleanName(sFileName As String) As String
  Const sInvalidChars As String = "/\|<>:*?"""
  Dim lCt As Long
  CleanName = sFileName
  For lCt = 1 To Len(sInvalidChars)
    CleanName = Replace(CleanName, Mid(sInvalidChars, lCt, 1), "-")
  Next lCt
End Function
