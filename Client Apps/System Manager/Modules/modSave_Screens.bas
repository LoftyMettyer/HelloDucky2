Attribute VB_Name = "modSave_Screens"
Option Explicit

Public Function SaveScreens() As Boolean
  ' Save the new or modified screen definitions.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
    
  With recScrEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      If !Deleted Then
        fOK = ScreenDelete
      ElseIf !New Then
        fOK = ScreenNew
      ElseIf !Changed Then
        fOK = ScreenSave
      End If
      
      .MoveNext
    Loop
  End With

TidyUpAndExit:
  SaveScreens = fOK
  Exit Function
  
ErrorTrap:
  'MsgBox "Error saving screen definitions" & _
         IIf(Trim(Err.Description) <> vbnullstring, "(" & Err.Description & ")", vbnullstring), vbCritical
  OutputError "Error saving screen definitions"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function ScreenDelete() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngScreenID As Long
  
  lngScreenID = recScrEdit!ScreenID
  
  gADOCon.Execute "DELETE FROM ASRSysScreens WHERE screenID=" & lngScreenID, , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysPageCaptions WHERE screenID=" & lngScreenID, , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysControls WHERE screenID=" & lngScreenID, , adCmdText + adExecuteNoRecords

  fOK = True
  
TidyUpAndExit:
  ScreenDelete = fOK
  Exit Function

ErrorTrap:
  'MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  OutputError "Error deleting screen"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ScreenNew() As Boolean
  ' Save the current screen definition to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iColumn As Integer
  Dim sName As String
  Dim rsScreens As New ADODB.Recordset
  Dim rsControls As New ADODB.Recordset
  Dim rsPageCaptions As New ADODB.Recordset

  Dim strAddToError As String
  
  
  'MH20010305
  'Several people have had errors saving changes in System Manager and
  'they all seems to imply that they have been doing things with screens.
  'Chances are that it is crashing out somewhere in here so I've added
  'the "strAddToError" variable so that if they do get an error we can
  'see which area of this sub they error occured.
  
  
  fOK = True
  
  strAddToError = "1"
  ' Open the Screens table on the server.
  rsScreens.Open "ASRSysScreens", gADOCon, adOpenDynamic, adLockOptimistic, adCmdTableDirect
    
  With rsScreens
    .AddNew
    strAddToError = "2"
    For iColumn = 0 To .Fields.Count - 1
      sName = .Fields(iColumn).Name
      If Not IsNull(recScrEdit.Fields(sName)) Then
        .Fields(iColumn) = recScrEdit.Fields(sName)
      End If
    Next iColumn
    .Update
    .Close
  End With
  
  strAddToError = "3"
  ' Open the Screen Page Captions table on the server.
  rsPageCaptions.Open "ASRSysPageCaptions", gADOCon, adOpenDynamic, adLockOptimistic, adCmdTableDirect
    
  With recPageCaptEdit
    .Index = "idxScreenPage"
    .Seek "=", recScrEdit!ScreenID, 1
  
    If Not .NoMatch Then
      Do While Not .EOF
        'If no more pages for this screen exit loop.
        If !ScreenID <> recScrEdit!ScreenID Then
          Exit Do
        End If
      
        rsPageCaptions.AddNew
        strAddToError = "4"
        For iColumn = 0 To rsPageCaptions.Fields.Count - 1
          sName = rsPageCaptions.Fields(iColumn).Name
          If Not IsNull(.Fields(sName)) Then
            rsPageCaptions.Fields(iColumn) = .Fields(sName)
          End If
        Next iColumn
        rsPageCaptions.Update
      
        'Get next control definition
        .MoveNext
      Loop
    End If
  End With
  rsPageCaptions.Close
  
  strAddToError = "5"
  ' Open the Screen Controls table on the server.
  rsControls.Open "ASRSysControls", gADOCon, adOpenDynamic, adLockOptimistic, adCmdTableDirect

  With recCtrlEdit
    .Index = "idxScreenID"
    .Seek ">=", recScrEdit!ScreenID
  
    If Not .NoMatch Then
      Do While Not .EOF
        'If no more controls for this screen exit loop
        If !ScreenID <> recScrEdit!ScreenID Then
          Exit Do
        End If
        
        'Add control details to Controls table
        rsControls.AddNew
        strAddToError = "6"
        For iColumn = 0 To rsControls.Fields.Count - 1
          sName = rsControls.Fields(iColumn).Name
          If Not IsNull(.Fields(sName)) Then
            rsControls.Fields(iColumn) = .Fields(sName)
          End If
        Next iColumn
        rsControls.Update
        
        'Get next control definition
        .MoveNext
      Loop
    End If
  End With
  rsControls.Close
  
  strAddToError = "7"
  
TidyUpAndExit:
  Set rsScreens = Nothing
  Set rsPageCaptions = Nothing
  Set rsControls = Nothing
  ScreenNew = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating screen (" & strAddToError & ")"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ScreenSave() As Boolean
  ' Save the current screen record to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = ScreenDelete
  If fOK Then
    fOK = ScreenNew
  End If

TidyUpAndExit:
  ScreenSave = fOK
  Exit Function
  
ErrorTrap:
  'MsgBox "Error saving screen" & _
         IIf(Trim(Err.Description) <> vbnullstring, "(" & Err.Description & ")", vbnullstring), vbCritical
  OutputError "Error updating screen"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function SaveHistoryScreens() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim rsHistoryScreens As ADODB.Recordset
  
  Set rsHistoryScreens = New ADODB.Recordset
  fOK = True
  
  'Delete any existing relation definitions
  gADOCon.Execute "DELETE FROM ASRSysHistoryScreens", , adCmdText + adExecuteNoRecords
  
  'Open ASR History Screens table
  rsHistoryScreens.Open "ASRSysHistoryScreens", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
  
  With recHistScrEdit
  
    'Loop through History Screens table in local database
    If Not (.BOF And .EOF) Then
    
      .MoveFirst
      Do While Not .EOF
        rsHistoryScreens.AddNew
        rsHistoryScreens!ID = !ID
        rsHistoryScreens!parentScreenID = !parentScreenID
        rsHistoryScreens!historyScreenID = !historyScreenID
        rsHistoryScreens.Update
        
        .MoveNext
      Loop
    End If
  End With
  
  rsHistoryScreens.Close
  
TidyUpAndExit:
  Set rsHistoryScreens = Nothing
  SaveHistoryScreens = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  gobjProgress.Visible = False
  'MsgBox ODBC.FormatError(Err.Description), vbOKOnly + vbExclamation, Application.Name
  OutputError "Error saving history screens"
  Resume TidyUpAndExit

End Function
