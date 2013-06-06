Attribute VB_Name = "modSave_QuickChecks"
Option Explicit


'Private Function QuickChecks_1(ByRef psErrMsg As String) As Boolean
Public Function QuickChecks_1() As Boolean
  On Error GoTo ErrorTrap
  
  ' Do some quick validation checks before the lengthy save procedure starts.
  Dim fOK As Boolean
  Dim fTableOK As Boolean
  
  fOK = True
  
  ' Check that all tables have a default order.
  With recTabEdit
    .MoveFirst
    
    Do While (Not .EOF) And fOK
      fTableOK = (!Deleted) Or _
        (!defaultOrderID > 0)
  
      If Not fTableOK Then
        'psErrMsg = "A primary order must be defined for the '" & !TableName & "' table."
        OutputError "A primary order must be defined for the '" & !TableName & "' table."
        fOK = False
      End If
      
      .MoveNext
    Loop
  End With
    
  QuickChecks_1 = fOK

Exit Function

ErrorTrap:
  OutputError "Error performing quick checks 1"
  fOK = False

End Function


Public Function QuickChecks_2() As Boolean
  ' Do some quick validation checks before the lengthy save procedure starts.
  Dim fOK As Boolean
  Dim fAudit_TableDone As Boolean
  Dim fDiaryLink_TableDone As Boolean
  Dim iNextIndex As Integer
  Dim sMessage As String
  Dim asAuditTables() As String
  Dim asDiaryLinkTables() As String
  
  On Error GoTo ErrorTrap
  
  fOK = True
  ReDim asAuditTables(0)
  ReDim asDiaryLinkTables(0)
  
  If fOK Then
    ' Check that all all tables with Audit or Diary linked columns have a
    ' record description defined.
    With recTabEdit
      .MoveFirst
        
      Do While (Not .EOF) And fOK
        fAudit_TableDone = False
        fDiaryLink_TableDone = False
      
        If (Not !Deleted) And (Not !RecordDescExprID > 0) Then
          ' The current table has no record description.
          ' Check if it has any Audit or Diary linked columns.
          recColEdit.Index = "idxName"
          recColEdit.Seek ">=", !TableID
          
          If Not recColEdit.NoMatch Then
            Do While (Not recColEdit.EOF) And fOK
              If recColEdit!TableID <> !TableID Then
                Exit Do
              End If
              
              If (Not recColEdit!Deleted) Then
                If (recColEdit!Audit) And (Not fAudit_TableDone) Then
                  iNextIndex = UBound(asAuditTables) + 1
                  ReDim Preserve asAuditTables(iNextIndex)
                  asAuditTables(iNextIndex) = !TableName
                  
                  fAudit_TableDone = True
                End If
              
                If fOK And (recColEdit!DataType = dtTIMESTAMP) And (Not fDiaryLink_TableDone) Then
                  ' The column is a date column. Check if its got any Diary links.
                  recDiaryEdit.Index = "idxColumnID"
                  recDiaryEdit.Seek "=", recColEdit!ColumnID
                  
                  If Not recDiaryEdit.NoMatch Then
                    iNextIndex = UBound(asDiaryLinkTables) + 1
                    ReDim Preserve asDiaryLinkTables(iNextIndex)
                    asDiaryLinkTables(iNextIndex) = !TableName
                    
                    fDiaryLink_TableDone = True
                  End If
                End If
              End If
              
              recColEdit.MoveNext
            Loop
          End If
        End If
      
        .MoveNext
      Loop
    End With
  End If
  
  ' Prompt the user if they want to continue or not, if any tables have failed the quick checks.
  sMessage = vbNullString
  If UBound(asAuditTables) > 0 Then
    sMessage = "The following tables have audited columns that require the definition of a record description." & vbNewLine

    For iNextIndex = 1 To UBound(asAuditTables)
      sMessage = sMessage & vbNewLine & _
        "    " & asAuditTables(iNextIndex)
    Next iNextIndex
  End If
  
  If UBound(asDiaryLinkTables) > 0 Then
    sMessage = sMessage & IIf(UBound(asAuditTables) > 0, vbNewLine & vbNewLine, vbNullString) & _
      "The following tables have diary linked columns that require the definition of a record description." & vbNewLine

    For iNextIndex = 1 To UBound(asDiaryLinkTables)
      sMessage = sMessage & vbNewLine & _
        "    " & asDiaryLinkTables(iNextIndex)
    Next iNextIndex
  End If
  If LenB(sMessage) <> 0 Then
    'gobjProgress.Visible = False
    'fOK = (MsgBox(sMessage & vbNewLine & vbNewLine & _
    '  "Continue saving changes ?", _
    '  vbQuestion + vbYesNo, App.Title) = vbYes)
    'If fOK Then
    '  gobjProgress.Visible = True
    'End If
    fOK = (OutputMessage(sMessage & vbNewLine & vbNewLine & "Continue saving changes ?") = vbYes)
  End If

TidyAndExit:
  QuickChecks_2 = fOK

Exit Function

ErrorTrap:
  OutputError "Error performing quick checks 2"
  fOK = False
  Resume TidyAndExit

End Function

Public Function QuickChecks_3() As Boolean

  ' Do some quick validation checks before the lengthy save procedure starts.
  Dim LocalRecTabEdit As dao.Recordset
  Dim fOK As Boolean
  Dim iNextIndex As Integer
  Dim sMessage As String
  Dim asOrphanedTables() As String
  
  On Error GoTo ErrorTrap
  
  fOK = True
  ReDim asOrphanedTables(0)
  
  If fOK Then
    Set LocalRecTabEdit = recTabEdit.Clone
      
    ' Check if all the child tables have parents.
    With recTabEdit
      .MoveFirst
        
      Do While (Not .EOF) And fOK
        
        If (!TableType = enum_TableTypes.iTabChild) Then
          If (Not !Deleted) And Not (recRelEdit.BOF And recRelEdit.EOF) Then
            ' Check if it has any relationships.
            recRelEdit.MoveFirst
            recRelEdit.Index = "idxChildID"
            
            recRelEdit.Seek "=", !TableID
            
            If recRelEdit.NoMatch Then
              LocalRecTabEdit.MoveFirst
              LocalRecTabEdit.Index = "idxTableID"
              LocalRecTabEdit.Seek ">=", !TableID
              
              If (Not LocalRecTabEdit!Deleted) Then
                iNextIndex = UBound(asOrphanedTables) + 1
                ReDim Preserve asOrphanedTables(iNextIndex)
                asOrphanedTables(iNextIndex) = LocalRecTabEdit!TableName
              End If
              
            End If
          End If
        End If
        
        .MoveNext
      Loop
    End With
  End If
  
  ' Prompt the user if they want to continue or not, if any tables have failed the quick checks.
  sMessage = vbNullString
  If UBound(asOrphanedTables) > 0 Then
    sMessage = "The following child tables do not have any relationships." & vbNewLine

    For iNextIndex = 1 To UBound(asOrphanedTables)
      sMessage = sMessage & vbNewLine & _
        "    " & asOrphanedTables(iNextIndex)
    Next iNextIndex
    
    sMessage = sMessage & vbNewLine & vbNewLine & "Default permissions will not be applied."
  End If
  
  If LenB(sMessage) <> 0 Then
    fOK = (OutputMessage(sMessage & vbNewLine & vbNewLine & "Continue saving changes ?") = vbYes)
  End If

TidyAndExit:
  Set LocalRecTabEdit = Nothing
  QuickChecks_3 = fOK

Exit Function

ErrorTrap:
  OutputError "Error performing quick checks 3"
  fOK = False
  Resume TidyAndExit

End Function


Public Function QuickChecks_4() As Boolean

  Dim strEncrypted As String
  Dim rstTest As ADODB.Recordset
  Dim bOK As Boolean
  Dim strLogon As String

  On Error GoTo LocalErr

  'MH20061024 Fault 11610
  If Not (glngProcessMethod = iPROCESSADMIN_SERVICEACCOUNT Or glngProcessMethod = iPROCESSADMIN_SQLACCOUNT) Or glngSQLVersion < 9 Then
    QuickChecks_4 = True
    Exit Function
  End If


  ' Get the encrypted login string
  With recModuleSetup
    bOK = False
    .Index = "idxModuleParameter"
      
    ' Get the Login Name.
    .Seek "=", gsMODULEKEY_SQL, gsPARAMETERKEY_LOGINDETAILS
    If .NoMatch Then
      strEncrypted = ""
    Else
      ' AE20080410 Fault #13087
      'strEncrypted = IIf(IsNull(!ParameterValue) Or Len(!ParameterValue) = 0, 0, !ParameterValue)
      strEncrypted = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, vbNullString, !parametervalue)
    End If
  End With

  ' Test the encrypted logon
  Set rstTest = New ADODB.Recordset
  rstTest.Open "SELECT dbo.udfASRNetIsProcessValid('" & Replace(strEncrypted, "'", "''") & "')", gADOCon, adOpenForwardOnly, adLockReadOnly
  bOK = (rstTest.Fields(0).value = True)
  rstTest.Close

  
  If bOK = False Then
    GoTo LocalErr
  End If

TidyUpAndExit:
  Set rstTest = Nothing
  QuickChecks_4 = bOK
  Exit Function

LocalErr:
  bOK = False
  Screen.MousePointer = vbDefault
  Err.Clear
  OutputError "The SQL process account has not been defined or is invalid." & vbNewLine & _
         "Please go to Configuration on the Administration menu and define a valid processing account."
  GoTo TidyUpAndExit

End Function


Public Function RegenerateProcessAccount() As Boolean

  Dim strEncrypted As String
  Dim bOK As Boolean
  Dim strLogon As String
  Dim sName As String
  Dim sPassword As String
  Dim sDatabase As String
  Dim sServer As String

  On Error GoTo LocalErr

  'MH20061024 Fault 11610
  If glngSQLVersion < 9 Then
    RegenerateProcessAccount = True
    Exit Function
  End If

  With recModuleSetup
    bOK = False
    .Index = "idxModuleParameter"
      
    ' Get the Login Name.
    .Seek "=", gsMODULEKEY_SQL, gsPARAMETERKEY_LOGINDETAILS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_SQL
      !parameterkey = gsPARAMETERKEY_LOGINDETAILS
      !ParameterType = gsPARAMETERTYPE_ENCYPTED
      !parametervalue = EncryptLogonDetails("", "", gsDatabaseName, gsServerName)
      .Update
      
      glngProcessMethod = iPROCESSADMIN_SERVICEACCOUNT
      
    Else
      strEncrypted = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
      
      DecryptLogonDetails strEncrypted, sName, sPassword, sDatabase, sServer
      strEncrypted = EncryptLogonDetails(sName, sPassword, gsDatabaseName, gsServerName)
      
      .Edit
      !parametervalue = strEncrypted
      .Update
      
    End If
  End With

TidyUpAndExit:
  RegenerateProcessAccount = bOK
  Exit Function

LocalErr:
  bOK = False
  GoTo TidyUpAndExit

End Function

