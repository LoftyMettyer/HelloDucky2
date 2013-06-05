Attribute VB_Name = "modOutlookCalendar"
Option Explicit


'MH20040322
Public Function SaveOutlookLinks(lngTableID As Long) As Boolean

  Dim rsOutlookLinks As New ADODB.Recordset
  Dim rsOutlookLinksDestinations As New ADODB.Recordset
  Dim rsOutlookLinksColumns As New ADODB.Recordset
  Dim sSQL As String

  rsOutlookLinks.Open "SELECT * FROM ASRSysOutlookLinks", gADOCon, adOpenDynamic, adLockOptimistic
  rsOutlookLinksDestinations.Open "SELECT * FROM ASRSysOutlookLinksDestinations", gADOCon, adOpenDynamic, adLockOptimistic
  rsOutlookLinksColumns.Open "SELECT * FROM ASRSysOutlookLinksColumns", gADOCon, adOpenDynamic, adLockOptimistic

  With recOutlookLinks
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do While Not .EOF
        If !TableID = lngTableID Then
          If !Deleted Then
            sSQL = "UPDATE ASRSysOutlookEvents SET Deleted = 1 " & _
              "WHERE LinkID = " & CStr(!LinkID)
            gADOCon.Execute sSQL, , adExecuteNoRecords

          Else
            rsOutlookLinks.AddNew
            
            rsOutlookLinks!LinkID = !LinkID
            rsOutlookLinks!TableID = !TableID
            rsOutlookLinks!Title = !Title
            rsOutlookLinks!FilterID = !FilterID
            rsOutlookLinks!BusyStatus = !BusyStatus
            rsOutlookLinks!StartDate = !StartDate
            rsOutlookLinks!EndDate = !EndDate
            rsOutlookLinks!TimeRange = !TimeRange
            rsOutlookLinks!FixedStartTime = !FixedStartTime
            rsOutlookLinks!FixedEndTime = !FixedEndTime
            rsOutlookLinks!ColumnStartTime = !ColumnStartTime
            rsOutlookLinks!ColumnEndTime = !ColumnEndTime
            rsOutlookLinks!Subject = !Subject
            rsOutlookLinks!content = IIf(IsNull(!content), vbNullString, !content)
            rsOutlookLinks!Reminder = !Reminder
            rsOutlookLinks!ReminderOffset = !ReminderOffset
            rsOutlookLinks!ReminderPeriod = !ReminderPeriod
  
            rsOutlookLinks.Update
            rsOutlookLinks.MoveLast
  
            'Add the outlook attachment values.
            With recOutlookLinksDestinations
              If Not (.BOF And .EOF) Then
                .MoveFirst
  
                Do While Not .EOF
                  If !LinkID = recOutlookLinks!LinkID Then
                    rsOutlookLinksDestinations.AddNew
                    rsOutlookLinksDestinations!LinkID = !LinkID
                    rsOutlookLinksDestinations!FolderID = !FolderID
                    rsOutlookLinksDestinations.Update
                  End If
  
                  .MoveNext
                Loop
  
              End If
            End With
      
      
            'Add the outlook attachment values.
            With recOutlookLinksColumns
              If Not (.BOF And .EOF) Then
                .MoveFirst
  
                Do While Not .EOF
                  If !LinkID = recOutlookLinks!LinkID Then
                    rsOutlookLinksColumns.AddNew
                    rsOutlookLinksColumns!LinkID = !LinkID
                    rsOutlookLinksColumns!ColumnID = !ColumnID
                    rsOutlookLinksColumns!Heading = !Heading
                    rsOutlookLinksColumns!Sequence = !Sequence
                    rsOutlookLinksColumns.Update
                  End If
  
                  .MoveNext
                Loop
              End If
            End With
    
          End If
        End If

        .MoveNext
      Loop
    End If
  End With

  SaveOutlookLinks = True
  
TidyUpAndExit:
  Set rsOutlookLinks = Nothing
  Set rsOutlookLinksDestinations = Nothing
  Set rsOutlookLinksColumns = Nothing
  
  Exit Function

LocalErr:
  SaveOutlookLinks = False
  Resume TidyUpAndExit

End Function


Public Function SaveOutlookFolders() As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  With recOutlookFolders
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      If !Deleted Then
'        Dim objOutlookFolder As New clsOutlookFolder
'        objOutlookFolder.FolderID = !FolderID
'        Set mfrmUse = New frmUsage
'        mfrmUse.ResetList
'        If objOutlookFolder.OutlookIsUsed(mfrmUse) Then
'          gobjProgress.Visible = False
'          Screen.MousePointer = vbNormal
'          mfrmUse.ShowMessage !Name & " Outlook", "The Outlook cannot be deleted as the Outlook is used by the following:", UsageCheckObject.Outlook
'          fOK = False
'        End If
'        UnLoad mfrmUse
'        Set mfrmUse = Nothing
'
'        gobjProgress.Visible = True
'
'        If fOK Then
          fOK = OutlookFolderDelete
'        End If
        
      ElseIf !New Then
        fOK = OutlookFolderNew
      ElseIf !Changed Then
        fOK = OutlookFolderDelete
        If fOK Then
          fOK = OutlookFolderNew
        End If
      End If
      
      .MoveNext
    Loop
  End With
  
TidyUpAndExit:
  SaveOutlookFolders = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error creating Outlook addresses"
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function OutlookFolderDelete() As Boolean
  ' Delete the current Order definition from the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  sSQL = "DELETE FROM ASRSysOutlookFolders" & _
    " WHERE FolderID = " & CStr(recOutlookFolders!FolderID)
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
TidyUpAndExit:
  OutlookFolderDelete = fOK
  Exit Function

ErrorTrap:
  fOK = False
  OutputError "Error Deleting Outlook address"
  Resume TidyUpAndExit
  
End Function

Private Function OutlookFolderNew() As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iColumn As Integer
  Dim sName As String
  Dim rsOutlookFolders As New ADODB.Recordset
  
  fOK = True
  
  rsOutlookFolders.Open "SELECT * FROM ASRSysOutlookFolders", gADOCon, adOpenDynamic, adLockOptimistic
  With rsOutlookFolders
    .AddNew
    .Fields("FolderID") = !FolderID

    For iColumn = 0 To .Fields.Count - 1
      sName = .Fields(iColumn).Name
      
'      If Not UCase(Trim(sName)) = "TIMESTAMP" Then
        If Not IsNull(recOutlookFolders.Fields(sName)) Then
          .Fields(iColumn) = recOutlookFolders.Fields(sName)
        End If
'      End If
    Next iColumn
    .Update
  End With
  rsOutlookFolders.Close
  
TidyUpAndExit:
  Set rsOutlookFolders = Nothing
  OutlookFolderNew = fOK
  Exit Function

ErrorTrap:
  OutputError "Error saving new Outlook address"
  fOK = False
  Resume TidyUpAndExit
  
End Function



Public Function CreateOutlookEventsForTable(lngTableID As Long, sCurrentTable As String, lngRecordDescExprID As Long) As Boolean

  Dim strSQLTemp As String
  Dim strSQLLinks As String
  Dim strSQLFilter As String
  Dim lngLinkID As Long
  Dim lngFolderID As Long
  Dim strChildTableName As String

  strSQLLinks = vbNullString

  With recOutlookLinks

    If Not .BOF Or Not .EOF Then
      .Index = "idxTableID"
      .Seek "=", lngTableID

      If Not .NoMatch Then
        Do While Not .EOF
          If !TableID <> lngTableID Then
            Exit Do
          End If

          If Not !Deleted Then

            lngLinkID = !LinkID
  
            strSQLTemp = vbNullString
  
            With recOutlookLinksDestinations
              '.MoveFirst
              .Index = "idxLinkID"
              .Seek "=", lngLinkID
  
              If Not .NoMatch Then
                Do While Not .EOF
                  
                  If !LinkID <> lngLinkID Then
                    Exit Do
                  End If
  
                  lngFolderID = !FolderID
                  
                  strSQLTemp = strSQLTemp & _
                    "      EXEC spASROutlookEventRefresh " & CStr(lngLinkID) & ", " & CStr(lngFolderID) & ", " & CStr(lngTableID) & ", @recordID" & vbCrLf
                  .MoveNext
                Loop
              End If
  
            End With
  
  
            If strSQLTemp = vbNullString Then
              OutputError "Outlook link: " & !Title & " <" & sCurrentTable & "> has no destinations specified."
              CreateOutlookEventsForTable = False
              Exit Function
            End If
  
  
            If !FilterID > 0 Then
              strSQLFilter = GetSQLFilter(!FilterID, sCurrentTable)
              ''Need to restore current records after getting filter
              '.Index = "idxID"
              '.Seek "=", lngLinkID
              'recColEdit.Index = "idxColumnID"
              'recColEdit.Seek "=", lngColumnID
  
              strSQLTemp = _
              "    IF " & strSQLFilter & vbCrLf & _
              "    BEGIN" & vbCrLf & _
              strSQLTemp & _
              "    END" & vbCrLf & _
              "    ELSE" & vbCrLf & _
              "      UPDATE ASRSysOutlookEvents SET Deleted = 1 " & _
                  "WHERE LinkID = " & CStr(lngLinkID) & _
                  " AND RecordID = @recordID" & vbCrLf
            End If
  
            strSQLTemp = _
            "  IF NOT (SELECT " & GetColumnName(!StartDate) & " FROM [" & sCurrentTable & "] WHERE ID = @RecordID) IS NULL" & vbCrLf & _
            "  BEGIN" & vbCrLf & _
            strSQLTemp & _
            "  END" & vbCrLf & _
            "  ELSE" & vbCrLf & _
            "    UPDATE ASRSysOutlookEvents SET Deleted = 1 " & _
               " WHERE LinkID = " & CStr(lngLinkID) & _
               " AND RecordID = @recordID" & vbCrLf
  
            strSQLLinks = strSQLLinks & _
              "  --" & !Title & vbCrLf & _
              strSQLTemp & vbCrLf
          
          End If
          
          .MoveNext
        Loop
    
      End If
    End If

  End With
  
  
  
  With recRelEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If !parentID = lngTableID Then
        If TableHasOutlookLinks(!childID) Then

          strChildTableName = GetTableName(!childID)

          strSQLLinks = strSQLLinks & vbCrLf & _
            "  -- " & strChildTableName & vbCrLf & _
            "  DECLARE HRProCursor" & CStr(lngTableID) & " CURSOR" & vbCrLf & _
            "  FOR SELECT ID FROM [" & strChildTableName & "] WHERE ID_" & CStr(lngTableID) & " = @RecordID" & vbCrLf & _
            "  OPEN HRProCursor" & CStr(lngTableID) & vbCrLf & vbCrLf & _
            "  FETCH NEXT FROM HRProCursor" & CStr(lngTableID) & " INTO @ChildID" & vbCrLf & _
            "  WHILE (@@fetch_status <> -1)" & vbCrLf & _
            "  BEGIN" & vbCrLf & _
            "    IF (@@fetch_status <> -2)" & vbCrLf & _
            "      EXEC dbo.spASROutlook_" & CStr(!childID) & " @ChildID" & vbCrLf & _
            "    FETCH NEXT FROM HRProCursor" & CStr(lngTableID) & " INTO @ChildID" & vbCrLf & _
            "  END" & vbCrLf & vbCrLf & _
            "  CLOSE HRProCursor" & CStr(lngTableID) & vbCrLf & _
            "  DEALLOCATE HRProCursor" & CStr(lngTableID) & vbCrLf

        End If
      End If

      .MoveNext
    Loop
  End With
  
  
  
  
  'If Application.ChangedOutlookLink Then

'    strSQLTemp = _
'      "IF EXISTS (SELECT Name FROM sysobjects" & _
'      "    WHERE id = object_id('spASROutlook_" & CStr(lngTableID) & "')" & _
'      "    AND sysstat & 0xf = 4)" & _
'      "  DROP PROCEDURE spASROutlook_" & CStr(lngTableID)
'    gADOCon.Execute strSQLTemp, , adExecuteNoRecords
    DropProcedure "spASROutlook_" & CStr(lngTableID)

    If strSQLLinks <> vbNullString Then
    
      strSQLTemp = _
        "/* ------------------------------------------------------------------------------- */" & vbCrLf & _
        "/* HR Pro Outlook address stored procedure.                       */" & vbCrLf & _
        "/* Automatically generated by the System Manager.   */" & vbCrLf & _
        "/* ------------------------------------------------------------------------------- */" & vbCrLf & _
        "CREATE PROCEDURE dbo.spASROutlook_" & CStr(lngTableID) & vbCrLf & _
        "(@RecordID int)" & vbCrLf & _
        "AS" & vbCrLf & _
        "BEGIN" & vbCrLf & vbCrLf & _
        "  DECLARE @ChildID int" & vbCrLf & vbCrLf & _
        strSQLLinks & vbCrLf & _
        "END"

      gADOCon.Execute strSQLTemp, , adExecuteNoRecords

    End If

  'End If
  
  
  CreateOutlookEventsForTable = True

End Function


Private Function GetSQLFilter(lngFilterID As Long, sCurrentTable As String) As String

  Dim fOK As Boolean
  Dim objExpr As CExpression
  Dim strFilterRunTimeCode As String

  GetSQLFilter = vbNullString

  'Filter
  Set objExpr = New CExpression
  With objExpr

    objExpr.ExpressionID = lngFilterID
    objExpr.ConstructExpression
    fOK = objExpr.RuntimeFilterCode(strFilterRunTimeCode, False)

    strFilterRunTimeCode = Replace(strFilterRunTimeCode, vbCrLf, " ")
      
    GetSQLFilter = "@recordID IN " & _
          "(" & strFilterRunTimeCode & ")"
    
  End With
  Set objExpr = Nothing

End Function


Private Function TableHasOutlookLinks(lngTableID As Long) As Boolean

  TableHasOutlookLinks = False

  With recOutlookLinks

    If Not .BOF Or Not .EOF Then
      .Index = "idxTableID"
      .Seek "=", lngTableID

      If Not .NoMatch Then
        Do While Not .EOF
          If !TableID <> lngTableID Then
            Exit Do
          End If
          If Not !Deleted Then
            TableHasOutlookLinks = True
            Exit Do
          End If
          .MoveNext
        Loop
      End If
    End If

  End With

End Function
