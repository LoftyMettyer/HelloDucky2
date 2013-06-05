Attribute VB_Name = "modSave_EmailLinks"
Option Explicit

Public glngExpressionTableIDForDeleteTrigger As Boolean

Private rsEmailLinks As ADODB.Recordset
Private rsEmailRecipients As ADODB.Recordset
Private rsEmailColumns As ADODB.Recordset
Private rsLinkContent As ADODB.Recordset


'MH20090520
Public Const strDelimStart As String = "«"   'asc = 171
Public Const strDelimStop As String = "»"    'asc = 187


Public glngEmailMethod As Long
Public gstrEmailProfile As String
Public gstrEmailServer As String
Public gstrEmailAccount As String

Public glngEmailDateFormat As Long
Public gstrEmailAttachmentPath As String
Public gstrEmailTestAddr As String

Public gstrInsertEmailCode As String
Public gstrUpdateEmailCode As String
Public gstrDeleteEmailCode As String

Public Enum EmailType
  LinkRecord = 0
  LinkColumn = 1
  LinkOffset = 2
  LinkAmendment = 3
  LinkRebuild = 4
End Enum



Public Function OpenEmailRecordsets()

  Set rsEmailLinks = New ADODB.Recordset
  rsEmailLinks.Open "ASRSysEmailLinks", gADOCon, adOpenKeyset, adLockOptimistic, adCmdTableDirect
  
  Set rsEmailRecipients = New ADODB.Recordset
  rsEmailRecipients.Open "ASRSysEmailLinksRecipients", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

  Set rsLinkContent = New ADODB.Recordset
  rsLinkContent.Open "ASRSysLinkContent", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

  Set rsEmailColumns = New ADODB.Recordset
  rsEmailColumns.Open "ASRSysEmailLinksColumns", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

End Function


Public Function SaveEmailLinks(lngTableID As Long) As Boolean

  On Local Error GoTo LocalErr

  With recEmailLinksEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
  
      Do While Not .EOF
        If !TableID = lngTableID Then
          If Not !Deleted Then
            rsEmailLinks.AddNew
            rsEmailLinks!LinkID = !LinkID
            rsEmailLinks!TableID = !TableID
            rsEmailLinks!Title = IIf(IsNull(!Title), vbNullString, !Title)
            rsEmailLinks!FilterID = !FilterID
            rsEmailLinks!EffectiveDate = !EffectiveDate
            rsEmailLinks!Attachment = IIf(IsNull(!Attachment), vbNullString, !Attachment)
            rsEmailLinks!Type = IIf(IsNull(!Type), 0, !Type)
            
            rsEmailLinks!SubjectContentID = IIf(IsNull(!SubjectContentID), 0, !SubjectContentID)
            rsEmailLinks!BodyContentID = IIf(IsNull(!BodyContentID), 0, !BodyContentID)
  
            rsEmailLinks!RecordInsert = !RecordInsert
            rsEmailLinks!RecordDelete = !RecordDelete
            rsEmailLinks!RecordUpdate = !RecordUpdate
  
            rsEmailLinks!DateColumnID = !DateColumnID
            rsEmailLinks!DateOffset = !DateOffset
            rsEmailLinks!DatePeriod = !DatePeriod
            rsEmailLinks!DateAmendment = !DateAmendment
            
            rsEmailLinks.Update
            rsEmailLinks.MoveLast
  
  
            SaveLinkContent IIf(IsNull(recEmailLinksEdit!SubjectContentID), 0, recEmailLinksEdit!SubjectContentID)
            SaveLinkContent IIf(IsNull(recEmailLinksEdit!BodyContentID), 0, recEmailLinksEdit!BodyContentID)
  
  
            ' Add references to email recipients
            With recEmailRecipientsEdit
              If Not (.BOF And .EOF) Then
                .MoveFirst
  
                Do While Not .EOF
                  If !LinkID = recEmailLinksEdit!LinkID Then
                    rsEmailRecipients.AddNew
                    rsEmailRecipients!LinkID = recEmailLinksEdit!LinkID
                    rsEmailRecipients!RecipientID = !RecipientID
                    rsEmailRecipients!Mode = !Mode
                    rsEmailRecipients.Update
                  End If
    
                  .MoveNext
                Loop
              End If
            End With
    
    
            ' Add references to email recipients
            With recEmailLinksColumnsEdit
              If Not (.BOF And .EOF) Then
                .MoveFirst
  
                Do While Not .EOF
                  If !LinkID = recEmailLinksEdit!LinkID Then
                    rsEmailColumns.AddNew
                    rsEmailColumns!LinkID = recEmailLinksEdit!LinkID
                    rsEmailColumns!ColumnID = !ColumnID
                    rsEmailColumns.Update
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

  SaveEmailLinks = True
  
Exit Function

LocalErr:
  MsgBox "Error saving email links" & vbCrLf & Err.Description, vbCritical
  SaveEmailLinks = False

End Function


Private Function DeleteLinkContent(lngContentID As Long)

  If lngContentID > 0 Then
  
    With rsLinkContent
      If Not .BOF Or Not .EOF Then
        .MoveFirst
        Do While Not .EOF
          If !ContentID = lngContentID Then
            rsLinkContent.Delete
          End If
          rsLinkContent.MoveNext
        Loop
      End If
    
    End With
  
  End If

End Function


Private Function SaveLinkContent(lngContentID As Long)

  If lngContentID > 0 Then
  
    With recLinkContentEdit
      .Index = "idxContentIDSequence"
      .Seek ">=", lngContentID, 0
    
      If Not .NoMatch Then
        Do While Not .EOF
          
          If !ContentID <> lngContentID Then
            Exit Do
          End If

          rsLinkContent.AddNew
          rsLinkContent!id = !id
          rsLinkContent!ContentID = !ContentID
          rsLinkContent!Sequence = !Sequence
          rsLinkContent!FixedText = !FixedText
          rsLinkContent!FieldCode = !FieldCode
          rsLinkContent!FieldID = !FieldID
          rsLinkContent.Update

          .MoveNext
        Loop
      End If
    
    End With
  
  End If

End Function


Public Function CloseEmailRecordsets()

  If rsEmailLinks.State <> adStateClosed Then
    rsEmailLinks.Close
  End If
  Set rsEmailLinks = Nothing
  
  If rsEmailRecipients.State <> adStateClosed Then
    rsEmailRecipients.Close
  End If
  Set rsEmailRecipients = Nothing

  If rsEmailColumns.State <> adStateClosed Then
    rsEmailColumns.Close
  End If
  Set rsEmailColumns = Nothing

  If rsLinkContent.State <> adStateClosed Then
    rsLinkContent.Close
  End If
  Set rsLinkContent = Nothing

End Function


Private Function ApplyFilter(lngFilterID As Long, lngTableID As Long, strTableName As String, dtEffectiveDate As Date, strSQL As String) As String

  Dim fOK As Boolean
  Dim objExpr As CExpression
  Dim strFilter As String

  strFilter = vbNullString
  If lngFilterID > 0 Then
    Set objExpr = New CExpression
    With objExpr

      'MH20100324
      'Don't like this fix but also didn't like the idea of changing the expression builder too much... :o(
      glngExpressionTableIDForDeleteTrigger = IIf(strTableName = "deleted", lngTableID, 0)

      objExpr.ExpressionID = lngFilterID
      objExpr.ConstructExpression
      fOK = objExpr.RuntimeFilterCode(strFilter, False)
      
      glngExpressionTableIDForDeleteTrigger = False

      strFilter = Replace(strFilter, vbNewLine, " ")
      If Trim(strFilter) <> vbNullString Then
        strFilter = "@recordID IN (" & strFilter & ")"
      End If
      
    End With
    Set objExpr = Nothing
  End If
  
  
  If Not IsNull(dtEffectiveDate) Then
    strFilter = "DateDiff(day, '" & Replace(Format(dtEffectiveDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "', @emailDate) >= 0" & _
      IIf(strFilter <> vbNullString, vbNewLine & "              AND " & strFilter, "")
  End If


  If strFilter <> vbNullString Then
    strSQL = _
      "            IF " & strFilter & vbNewLine & _
      "            BEGIN" & vbNewLine & _
      strSQL & vbNewLine & _
      "            END" & vbNewLine
  End If
  
  
  ApplyFilter = strSQL

End Function



Public Sub CreateEmailProcsForTable(lngTableID As Long, _
  sCurrentTable As String, _
  lngRecordDescExprID As Long, _
  ByRef alngAuditColumns As Variant, _
  ByRef sDeclareInsCols As HRProSystemMgr.cStringBuilder, _
  ByRef sDeclareDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsCols2 As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sFetchInsCols As HRProSystemMgr.cStringBuilder, _
  ByRef sFetchDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsLargeCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsLargeCols2 As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectDelLargeCols As HRProSystemMgr.cStringBuilder)
  
  
  Dim strSQL As String
  Dim lngLinkID As Long
  Dim sTemp As String
  
  Dim strInsertUpdateOne As String
  Dim strDeleteOne As String
  Dim strRebuildAll As String
  Dim strRebuildOne As String
  Dim strSendCode As String
  Dim strInsertIntoQueue As String
  Dim strTemp As String

  Dim strLinkTitle As String
  Dim lngSubjectID As Long
  Dim lngBodyID As Long
  
  Dim strInsCol As String
  Dim strDelCol As String
  Dim strCheckColumns As String
  Dim lngColumnID As Long
  Dim lngColumnType As Long
  Dim strColumn As String


  On Error GoTo LocalErr

  strRebuildAll = vbNullString
  gstrInsertEmailCode = vbNullString
  gstrUpdateEmailCode = vbNullString
  gstrDeleteEmailCode = vbNullString


  
  'Loop through all of the email links for this table
  With recEmailLinksEdit

    If Not .BOF Or Not .EOF Then
      .Index = "idxID"
      .MoveFirst

      Do While Not .EOF

        strInsertUpdateOne = vbNullString
        strDeleteOne = vbNullString
        strCheckColumns = vbNullString

        If !TableID = lngTableID Then
          
          lngLinkID = !LinkID
          strLinkTitle = IIf(IsNull(!Title), vbNullString, !Title)
          lngSubjectID = IIf(IsNull(!SubjectContentID), 0, !SubjectContentID)
          lngBodyID = IIf(IsNull(!BodyContentID), 0, !BodyContentID)
          
          
          'Debug.Print sCurrentTable & "." & strLinkTitle
          
          CreateEmailProcedure lngTableID, sCurrentTable, lngLinkID, strLinkTitle, lngSubjectID, lngBodyID, !Attachment, !FilterID, !EffectiveDate
          
          
          'DATE RELATED
          If !Type = 2 Then

            lngColumnID = !DateColumnID
            Select Case GetColumnDataType(lngColumnID)
            Case dtNUMERIC, dtINTEGER, dtBIT
              strInsCol = "isnull(@insCol_" & CStr(lngColumnID) & ",0)"
              strDelCol = "isnull(@delCol_" & CStr(lngColumnID) & ",0)"
            Case Else
              strInsCol = "isnull(@insCol_" & CStr(lngColumnID) & ",'')"
              strDelCol = "isnull(@delCol_" & CStr(lngColumnID) & ",'')"
            End Select
            strCheckColumns = IIf(strCheckColumns <> vbNullString, strCheckColumns & " OR ", "") & _
                strInsCol & " <> " & strDelCol


            strInsertUpdateOne = _
                GetSQLForOffsetEmail(lngLinkID, lngTableID, sCurrentTable, !FilterID, !EffectiveDate, strInsCol, lngColumnID, !DateOffset, !DatePeriod, !DateAmendment) & vbNewLine
            
            AddColumnToTrigger _
                  lngColumnID, alngAuditColumns, sDeclareInsCols, sDeclareDelCols, _
                  sSelectInsCols2, sSelectDelCols, sFetchInsCols, sFetchDelCols, _
                  sSelectInsLargeCols, sSelectInsLargeCols2, sSelectDelLargeCols
          
            
            'Rebuild
            strRebuildOne = _
                GetSQLForRebuild(lngLinkID, lngTableID, sCurrentTable, !FilterID, !EffectiveDate, lngColumnID, !DateOffset, !DatePeriod)
            If strRebuildOne <> vbNullString Then
              strRebuildAll = strRebuildAll & vbNewLine & _
              "            -- " & strLinkTitle & vbNewLine & _
                strRebuildOne & vbNewLine
            End If
          
           
          Else

            'RECORD RELATED
            If !Type = 1 Then
              strColumn = "null"
              strInsCol = "''"
              strDelCol = "''"
            
              'Get the content and store it now because the record is being deleted!
              If !RecordDelete Then
                strDeleteOne = _
                    GetSQLEmailContent(lngTableID, "deleted", lngRecordDescExprID, lngLinkID, lngSubjectID, lngBodyID, !Attachment) & _
                    GetInsertCommand(!Type, lngLinkID, lngTableID, True, False, "null", strDelCol, True)
                strDeleteOne = _
                  "                    SELECT @emailDate = getDate()" & vbNewLine & _
                  ApplyFilter(!FilterID, lngTableID, "deleted", !EffectiveDate, strDeleteOne)
              End If

            Else              'Column Related
              
              lngColumnID = IsRelatedToSingleColumn(lngLinkID)
              strColumn = CStr(lngColumnID)
              If lngColumnID = 0 Then
                strInsCol = "null"
              Else
                Select Case GetColumnDataType(lngColumnID)
                Case dtNUMERIC, dtINTEGER, dtBIT
                  strInsCol = "isnull(@insCol_" & CStr(lngColumnID) & ",0)"
                Case Else
                  strInsCol = "isnull(@insCol_" & CStr(lngColumnID) & ",'')"
                End Select
              End If
            
            End If

            strInsertUpdateOne = _
              "            DELETE FROM ASRSysEmailQueue WHERE DateSent IS NULL AND recordID = @recordID AND LinkID = " & CStr(lngLinkID) & vbNewLine & _
              GetInsertCommand(!Type, lngLinkID, lngTableID, True, True, strColumn, strInsCol, False)
            
            strInsertUpdateOne = _
              "                    SELECT @emailDate = getDate()" & vbNewLine & _
              ApplyFilter(!FilterID, lngTableID, sCurrentTable, !EffectiveDate, strInsertUpdateOne)
            
            
            strCheckColumns = GetSQLForImmediateEmails(lngLinkID, _
                  alngAuditColumns, sDeclareInsCols, sDeclareDelCols, _
                  sSelectInsCols2, sSelectDelCols, sFetchInsCols, sFetchDelCols, _
                  sSelectInsLargeCols, sSelectInsLargeCols2, sSelectDelLargeCols)

          End If
          
          
          If strCheckColumns <> vbNullString Then
            strInsertUpdateOne = _
              "            IF " & strCheckColumns & vbNewLine & _
              "            BEGIN" & vbNewLine & _
              strInsertUpdateOne & vbNewLine & _
              "            END" & vbNewLine
          End If

          strInsertUpdateOne = _
            "            -- " & strLinkTitle & vbNewLine & _
            strInsertUpdateOne & vbNewLine
          


          If !Type <> 1 Then
            gstrInsertEmailCode = gstrInsertEmailCode & strInsertUpdateOne
            gstrUpdateEmailCode = gstrUpdateEmailCode & strInsertUpdateOne
          
          Else  'Record Related
            
            If !RecordInsert Then
              gstrInsertEmailCode = gstrInsertEmailCode & strInsertUpdateOne
            End If
            
            If !RecordUpdate Then
              gstrUpdateEmailCode = gstrUpdateEmailCode & _
                "          IF @iTriggerLevel = 1" & vbCrLf & _
                "          BEGIN" & vbCrLf & _
                strInsertUpdateOne & vbCrLf & _
                "          END"
            End If
            
            If !RecordDelete Then
              gstrDeleteEmailCode = gstrDeleteEmailCode & _
                "            -- " & strLinkTitle & vbNewLine & _
                strDeleteOne & vbNewLine
            End If

          End If

        End If




        .MoveNext
      Loop  'Next Link
    End If
  End With
  
  
  strTemp = _
      "            DECLARE @emailDate datetime" & vbNewLine & _
      "            DECLARE @purgeDate datetime" & vbNewLine & _
      "            DECLARE @LastSent varchar(max)" & vbNewLine & _
      "            DECLARE @sColumnValue varchar(max)" & vbNewLine & _
      "            DECLARE @username varchar(max)" & vbNewLine & vbNewLine & _
      "            SELECT @username = CASE WHEN UPPER(LEFT(APP_NAME(), " & Len(gsWORKFLOWAPPLICATIONPREFIX) & ")) = '" & UCase(gsWORKFLOWAPPLICATIONPREFIX) & "' THEN '" & gsWORKFLOWAPPLICATIONPREFIX & "' ELSE rtrim(system_user) END" & vbNewLine & _
      "            EXEC sp_ASRPurgeDate @purgedate OUTPUT, 'EMAIL'" & vbNewLine & vbNewLine & _
      "            UPDATE ASRSysEmailQueue SET RecordDesc = @recordDesc WHERE RecordID = @recordID AND TableID = " & CStr(lngTableID) & vbNewLine & vbNewLine

  gstrInsertEmailCode = IIf(gstrInsertEmailCode <> vbNullString, strTemp & gstrInsertEmailCode, vbNullString)
  gstrUpdateEmailCode = IIf(gstrUpdateEmailCode <> vbNullString, strTemp & gstrUpdateEmailCode, vbNullString)
  
  
  strTemp = strTemp & _
      "  DECLARE @Recip varchar(max)" & vbNewLine & _
      "  DECLARE @To varchar(max)" & vbNewLine & _
      "  DECLARE @Cc varchar(max)" & vbNewLine & _
      "  DECLARE @Bcc varchar(max)" & vbNewLine & _
      "  DECLARE @Subject varchar(max)" & vbNewLine & _
      "  DECLARE @Message varchar(max)" & vbNewLine & _
      "  DECLARE @Attachment varchar(max)" & vbNewLine & vbNewLine
  
  gstrDeleteEmailCode = IIf(gstrDeleteEmailCode <> vbNullString, strTemp & gstrDeleteEmailCode, vbNullString)


  DropProcedure "spASREmailRebuild_" & CStr(lngTableID)
  If strRebuildAll <> vbNullString Then
    
    strRebuildAll = _
      "/* ------------------------------------------------ */" & vbNewLine & _
      "/* HR Pro email address stored procedure.           */" & vbNewLine & _
      "/* Automatically generated by the System Manager.   */" & vbNewLine & _
      "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE PROCEDURE dbo.spASREmailRebuild_" & CStr(lngTableID) & vbNewLine & _
      "(@recordid int)" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN" & vbNewLine & vbNewLine & _
      "  DECLARE @hResult bit" & vbNewLine & _
      "  DECLARE @dateValue datetime" & vbNewLine & _
      GetSQLForRecordDescription(lngRecordDescExprID) & vbCrLf & vbCrLf & _
      strTemp & vbNewLine & vbNewLine & vbNewLine & _
      "  DELETE FROM ASRSysEmailQueue WHERE Immediate = 0 AND DateSent IS NULL AND recordID = @recordID AND TableID = " & CStr(lngTableID) & vbNewLine & vbNewLine & _
      strRebuildAll & vbNewLine & _
      "END"
    
    gADOCon.Execute strRebuildAll, , adExecuteNoRecords
  
  End If
      
Exit Sub

LocalErr:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
    Stop
    Resume Next
  End If

End Sub



Private Function AddColumnToTrigger( _
  lngColumnID As Long, ByRef alngAuditColumns As Variant, _
  ByRef sDeclareInsCols As HRProSystemMgr.cStringBuilder, _
  ByRef sDeclareDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsCols2 As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sFetchInsCols As HRProSystemMgr.cStringBuilder, _
  ByRef sFetchDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsLargeCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsLargeCols2 As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectDelLargeCols As HRProSystemMgr.cStringBuilder)
                  
                  
  Dim fColFound As Boolean
  Dim iLoop As Integer
  Dim sConvertInsCols As String
    
  Dim strColumnName As String
  Dim iDataType As Integer
  Dim lngSize As Long
  Dim iDecimals As Integer
    
    
    ' JPD20020913 - instead of making multiple queries to the triggered table, and
    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
    ' the cursor that we used to loop through to get just the id of each record being
    ' inserted/updated/deleted.
    ' Here we are adding the email columns to the SELECT statement that is used
    ' to create the cursor, the FETCH statement that used to loop through the cursor,
    ' and the DECLARE statements that are needed.
    ' The email check code is modified for the new implementation.
    ' NB. an array of columns that have been added to the SELECT statement is used
    ' to ensure that columns aren't added more than once. As well as audit columns,
    ' we're also going to add email columns and calculated columns later on.
    ' This change was driven by the performance degradation reported by
    ' Islington.
    fColFound = False
    sConvertInsCols = ""
    
    
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", lngColumnID
    
      strColumnName = !ColumnName
      iDataType = !DataType
      lngSize = !Size
      iDecimals = !Decimals
    
    End With
    
    
    ' Check if the column has already been declared and added to the select and fetch strings
    For iLoop = 1 To UBound(alngAuditColumns)
      If alngAuditColumns(iLoop) = lngColumnID Then
        fColFound = True
        Exit For
      End If
    Next iLoop
  
    If Not fColFound Then
      ReDim Preserve alngAuditColumns(UBound(alngAuditColumns) + 1)
      alngAuditColumns(UBound(alngAuditColumns)) = lngColumnID
  
      If (iDataType <> dtVARCHAR) Or (lngSize <= VARCHARTHRESHOLD) Then
  
        'sSelectInsCols.Append "," & vbNewLine & "        inserted." & strColumnName
        sSelectInsCols2.Append ",@insCol_" & CStr(lngColumnID) & "=" & strColumnName
        sSelectDelCols.Append "," & vbNewLine & "        deleted." & strColumnName
  
        'sFetchInsCols.Append "," & vbNewLine & "        @insCol_" & CStr(lngColumnID)
        sFetchDelCols.Append "," & vbNewLine & "        @delCol_" & CStr(lngColumnID)
      Else
        sSelectInsLargeCols.Append ",@insCol_" & CStr(lngColumnID) & "=inserted." & strColumnName
        sSelectInsLargeCols2.Append ",@insCol_" & CStr(lngColumnID) & "=" & strColumnName
        sSelectDelLargeCols.Append ",@delCol_" & CStr(lngColumnID) & "=deleted." & strColumnName
      End If
  
      sDeclareInsCols.Append "," & vbNewLine & "        @insCol_" & CStr(lngColumnID)
      sDeclareDelCols.Append "," & vbNewLine & "        @delCol_" & CStr(lngColumnID)
    End If
  
    Select Case iDataType
      Case dtVARCHAR
        If Not fColFound Then
          sDeclareInsCols.Append " varchar(MAX)"
          sDeclareDelCols.Append " varchar(MAX)"
        End If
        sConvertInsCols = "ISNULL(CONVERT(varchar(max), @insCol_" & CStr(lngColumnID) & "), '')"
  
      Case dtLONGVARCHAR
        If Not fColFound Then
          sDeclareInsCols.Append " varchar(14)"
          sDeclareDelCols.Append " varchar(14)"
        End If
        sConvertInsCols = "ISNULL(CONVERT(varchar(max), @insCol_" & CStr(lngColumnID) & "), '')"
  
      Case dtINTEGER
        If Not fColFound Then
          sDeclareInsCols.Append " integer"
          sDeclareDelCols.Append " integer"
        End If
        sConvertInsCols = "ISNULL(CONVERT(varchar(max), @insCol_" & CStr(lngColumnID) & "), '')"
  
      Case dtNUMERIC
        If Not fColFound Then
          sDeclareInsCols.Append " numeric(" & CStr(lngSize) & ", " & CStr(iDecimals) & ")"
          sDeclareDelCols.Append " numeric(" & CStr(lngSize) & ", " & CStr(iDecimals) & ")"
        End If
        sConvertInsCols = "ISNULL(CONVERT(varchar(max), @insCol_" & CStr(lngColumnID) & "), '')"
  
        ' JDM - 12/09/03 - Fault 5605 - Use Separator in emails
        If recColEdit.Fields("Use1000Separator").value = True Then
          sConvertInsCols = sConvertInsCols & vbNewLine _
              & vbTab & "SET @sColumnValue = REVERSE(SUBSTRING(@sColumnValue,0,CHARINDEX('.',@sColumnValue)))" & vbNewLine _
              & vbTab & "SET @itemp = 3" & vbNewLine _
              & vbTab & "WHILE @itemp < LEN(@sColumnValue)" & vbNewLine _
              & vbTab & "BEGIN" & vbNewLine _
              & vbTab & vbTab & "SET @sColumnValue = LEFT(@sColumnValue, @itemp) + ',' + SUBSTRING(@sColumnValue,@itemp+1,len(@sColumnValue))" & vbNewLine _
              & vbTab & vbTab & "SET @itemp = @itemp + 4" & vbNewLine _
              & vbTab & "END" & vbNewLine _
              & vbTab & "SET @sColumnValue = REVERSE(@sColumnValue) + SUBSTRING(@sColumnValue,CHARINDEX('.',@sColumnValue),LEN(@sColumnValue))" & vbNewLine
        End If
  
      Case dtTIMESTAMP
        If Not fColFound Then
          sDeclareInsCols.Append " datetime"
          sDeclareDelCols.Append " datetime"
        End If
        sConvertInsCols = "ISNULL(CONVERT(varchar(max), LEFT(DATENAME(month, @insCol_" & CStr(lngColumnID) & "),3) + ' ' + CONVERT(varchar(max),DATEPART(day, @insCol_" & CStr(lngColumnID) & ")) + ' ' + CONVERT(varchar(max),DATEPART(year, @insCol_" & CStr(lngColumnID) & "))), '')"
  
      Case dtBIT
        If Not fColFound Then
          sDeclareInsCols.Append " bit"
          sDeclareDelCols.Append " bit"
        End If
        sConvertInsCols = "ISNULL(CONVERT(varchar(max), CASE @insCol_" & CStr(lngColumnID) & " WHEN 1 THEN 'True' WHEN 0 THEN 'False' END), '')"
  
      Case Else   'dtVARBINARY, dtLONGVARBINARY
        If Not fColFound Then
          sDeclareInsCols.Append " varchar(max)"
          sDeclareDelCols.Append " varchar(max)"
        End If
        sConvertInsCols = "ISNULL(CONVERT(varchar(max), @insCol_" & CStr(lngColumnID) & "), '')"
    End Select
    
End Function


Private Function GetSQLEmailContent(lngTableID As Long, strTableName As String, lngRecDescID As Long, lngLinkID As Long, lngSubjectID As Long, lngBodyID As Long, strAttachment As String)

  Dim content As clsLinkContent
  Dim recEmailRecipients As New ADODB.Recordset
  Dim strResult As String
  Dim strType As String
  Dim strSQL As String
  'Dim blnFoundToRecipient As Boolean
  Dim objExpr As New CExpression

  
  On Local Error GoTo LocalErr  'Resume Next


  '@To, @Cc, @Bcc
  strSQL = "  SELECT @To='', @Cc='', @Bcc=''" & vbNewLine
  
  'blnFoundToRecipient = False
  With recEmailRecipientsEdit

    .Index = "idxLinkID"
    .Seek ">=", lngLinkID

    If Not .NoMatch Then
      Do While Not .EOF
        If !LinkID <> lngLinkID Then
          Exit Do
        End If
  
        strType = Choose(!Mode + 1, "@To", "@Cc", "@Bcc")
        
        If strTableName <> "deleted" Then
          strSQL = strSQL & _
              "    EXEC @hResult = dbo.spASRSysEmailAddr @Recip OUTPUT, " & CStr(!RecipientID) & ", @recordID" & vbNewLine & _
              "    SELECT " & strType & " = " & strType & " + RTrim(@Recip) + ';'" & vbNewLine & vbNewLine
        Else
        
          With recEmailAddrEdit
            .Index = "idxID"
            .Seek ">=", recEmailRecipientsEdit!RecipientID
        
            If Not .NoMatch Then

              Select Case !Type
              Case 0  'Fixed
                strSQL = strSQL & _
                  "    SELECT " & strType & " = " & strType & _
                  " + ltrim(rtrim(Fixed))+';' FROM ASRSysEmailAddress WHERE EmailID = " & CStr(recEmailRecipientsEdit!RecipientID) & ";" & vbNewLine & vbNewLine
              Case 1  'Column
                strSQL = strSQL & _
                  "    SELECT " & strType & " = " & strType & _
                  " + rtrim(isnull(" & GetColumnName(!ColumnID, True) & ",''))+';' FROM deleted WHERE ID = @recordID;" & vbNewLine & vbNewLine
              Case 2  'Calc
                Set objExpr = New CExpression
                With objExpr
                  .ExpressionID = recEmailAddrEdit!ExprID
                  If .ConstructExpression Then
                    glngExpressionTableIDForDeleteTrigger = lngTableID
                    strSQL = strSQL & _
                      .StoredProcedureCode("@strTemp", "deleted") & vbNewLine & _
                      "    SELECT " & strType & " = " & strType & " + @strTemp+';';" & vbNewLine & vbNewLine
                    glngExpressionTableIDForDeleteTrigger = 0
                  End If
                  Set objExpr = Nothing
                End With
              End Select
            End If
        
          End With
        
        End If

        'blnFoundToRecipient = blnFoundToRecipient Or (!Mode = 0)
        .MoveNext
      Loop
  
    End If

  End With
  
  
  
  '@Subject
  Set content = New clsLinkContent
  content.ReadDetail lngSubjectID
  strSQL = strSQL & _
      content.GetSQL(lngTableID, strTableName, lngRecDescID, "@Subject") & vbNewLine & vbNewLine
  Set content = Nothing
  
  
  '@Message
  Set content = New clsLinkContent
  content.ReadDetail lngBodyID
  strSQL = strSQL & _
      content.GetSQL(lngTableID, strTableName, lngRecDescID, "@Message") & vbNewLine & vbNewLine
  Set content = Nothing



  '@Attachment
  strSQL = strSQL & _
      "    SET @Attachment = '" & strAttachment & "'" & vbNewLine


  strSQL = RemoveDuplicateDeclares(strSQL)
  GetSQLEmailContent = strSQL

Exit Function

LocalErr:
  MsgBox Err.Description, vbCritical
  GetSQLEmailContent = "SELECT @To='', @Cc='', @Bcc='', @Subject='', @Message='', @Attachment=''"

End Function



Private Function CreateEmailProcedure(lngTableID As Long, strTableName As String, lngLinkID As Long, strLinkTitle As String, lngSubjectID As Long, lngBodyID As Long, strAttachment As String, lngFilterID As Long, dtEffectiveDate As Date)

  Dim strSPEmailUpdateContent As String
  Dim strSQL As String
  Dim blnAttachment As Boolean

  On Local Error Resume Next

  
  strSQL = _
    GetSQLEmailContent(lngTableID, strTableName, 0, lngLinkID, lngSubjectID, lngBodyID, strAttachment) & vbNewLine & _
    "  UPDATE ASRSysEmailQueue SET RepTo = @To, RepCC = @CC, RepBCC = @BCC, Subject = @Subject, Msgtext = @Message, Attachment = @Attachment" & vbNewLine & _
    "  WHERE QueueID = @QueueID" & vbNewLine & vbNewLine
    
  strSQL = _
    ApplyFilter(lngFilterID, lngTableID, strTableName, dtEffectiveDate, strSQL) & vbNewLine & _
    "  ELSE" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    DELETE FROM ASRSysEmailQueue WHERE QueueID = @QueueID" & vbNewLine & _
    "    SET @hResult = 1" & vbNewLine & _
    "  END"


  'Drop Procedure
  strSPEmailUpdateContent = "dbo.spASREmail_" & CStr(lngLinkID)
  DropProcedure strSPEmailUpdateContent
  
  
  'Create Procedure
  strSQL = _
    "/* ----------------------------- */" & vbNewLine & _
    "/* HR Pro Email stored procedure */" & vbNewLine & _
    "/* ----------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE " & strSPEmailUpdateContent & vbNewLine & _
    "  ( @queueID int                   " & vbNewLine & _
    "  , @recordID int                  " & vbNewLine & _
    "  , @UserName varchar(max)         " & vbNewLine & _
    "  , @To varchar(max)         OUTPUT" & vbNewLine & _
    "  , @Cc varchar(max)         OUTPUT" & vbNewLine & _
    "  , @Bcc varchar(max)        OUTPUT" & vbNewLine & _
    "  , @Subject varchar(max)    OUTPUT" & vbNewLine & _
    "  , @Message varchar(max)    OUTPUT" & vbNewLine & _
    "  , @Attachment varchar(max) OUTPUT" & vbNewLine & _
    "  )" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "  DECLARE @ID int                  " & vbNewLine & _
    "  DECLARE @hResult int             " & vbNewLine & _
    "  DECLARE @Recip varchar(max)      " & vbNewLine & _
    "  DECLARE @emailDate datetime      " & vbNewLine & vbNewLine & _
    "  SET @emailDate = getdate()       " & vbNewLine & vbNewLine & _
    strSQL & vbNewLine & vbNewLine & _
    "END"
    
   gADOCon.Execute strSQL, , adExecuteNoRecords


Exit Function

LocalErr:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
    Stop
    Resume Next
  End If

End Function


Private Function GetSQLForRebuild(lngLinkID As Long, lngTableID As Long, sCurrentTable As String, lngFilterID As Long, dtEffectiveDate As Date, lngDateColumnID As Long, lngOffset As Long, lngPeriod As Long) As String

  Dim strOutput As String
  Dim strDateValue As String
  
  strDateValue = "isnull(convert(varchar(max),@DateValue," & CStr(glngEmailDateFormat) & "),'')"
    
  strOutput = _
    "                IF (DateDiff(day, @purgeDate, @emailDate) >= 0 OR @PurgeDate IS NULL)" & vbNewLine & _
    "                 AND IsNull(@LastSent,'') <> " & strDateValue & vbNewLine & _
    "                BEGIN" & vbNewLine & _
    GetInsertCommand(LinkRebuild, lngLinkID, lngTableID, False, True, CStr(lngDateColumnID), strDateValue, False) & vbNewLine & _
    "                END" & vbNewLine

  strOutput = ApplyFilter(lngFilterID, lngTableID, sCurrentTable, dtEffectiveDate, strOutput)
  
  
  strOutput = _
      "            SELECT @dateValue = " & GetColumnName(lngDateColumnID, True) & " FROM " & sCurrentTable & " WHERE id = @recordID" & vbNewLine & _
      "            SELECT @emailDate = " & ApplyOffset("@dateValue", lngPeriod, lngOffset) & vbNewLine & _
      GetLastSent(lngLinkID) & _
      strOutput
  
  GetSQLForRebuild = strOutput

End Function


Private Function GetSQLForOffsetEmail(lngLinkID As Long, lngTableID As Long, sCurrentTable As String, lngFilterID As Long, dtEffectiveDate As Date, strInsVar As String, lngDateColumnID As Long, lngOffset As Long, lngPeriod As Long, blnAmendment As Boolean) As String

  Dim strOutput As String
    
    
  strOutput = vbNullString
  
  If blnAmendment Then
    strOutput = vbNewLine & _
      "                  IF (DateDiff(day, getdate(), @emailDate) > 0) AND (@LastSent IS NOT NULL)" & vbNewLine & _
      "                  BEGIN" & vbNewLine & _
      "                    SELECT @emailDate = getDate()" & vbNewLine & _
      GetInsertCommand(LinkAmendment, lngLinkID, lngTableID, False, True, CStr(lngDateColumnID), strInsVar, False) & vbNewLine & _
      "                  END" & vbNewLine & vbNewLine
  End If

  strOutput = _
    "                IF (DateDiff(day, @purgeDate, @emailDate) >= 0 OR @PurgeDate IS NULL)" & vbNewLine & _
    "                    OR (@LastSent IS NOT NULL) " & vbNewLine & _
    "                BEGIN" & vbNewLine & _
    GetInsertCommand(LinkOffset, lngLinkID, lngTableID, False, True, CStr(lngDateColumnID), strInsVar, False) & vbNewLine & _
    strOutput & _
    "                END" & vbNewLine
  
  strOutput = _
      "            DELETE FROM ASRSysEmailQueue WHERE DateSent IS NULL AND recordID = @recordID AND LinkID = " & CStr(lngLinkID) & vbNewLine & _
      "            SELECT @emailDate   = " & ApplyOffset(strInsVar, lngPeriod, lngOffset) & vbNewLine & _
      GetLastSent(lngLinkID) & _
      ApplyFilter(lngFilterID, lngTableID, sCurrentTable, dtEffectiveDate, strOutput)

  GetSQLForOffsetEmail = strOutput


End Function


Private Function ApplyOffset(strInput As String, lngPeriod As Long, lngOffset As Long) As String

  strInput = "IsNull(convert(datetime," & strInput & "),getdate())"
  
  If Abs(lngOffset) > 0 Then
    strInput = "dateadd(" & Choose(lngPeriod + 1, "dd", "ww", "mm", "yy") & "," & CStr(lngOffset) & "," & strInput & ")"
  End If
  
  ApplyOffset = strInput

End Function


Private Function GetLastSent(lngLinkID As Long) As String
  GetLastSent = _
    "            SELECT @LastSent     = (SELECT TOP 1 [ColumnValue] FROM ASRSysEmailQueue " & vbNewLine & _
    "                                    WHERE recordid = @recordid AND LinkID = " & CStr(lngLinkID) & vbNewLine & _
    "                                    ORDER BY DateSent DESC)" & vbNewLine
End Function


Private Function GetSQLForImmediateEmails(lngLinkID As Long, _
  ByRef alngAuditColumns As Variant, _
  ByRef sDeclareInsCols As HRProSystemMgr.cStringBuilder, _
  ByRef sDeclareDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsCols2 As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sFetchInsCols As HRProSystemMgr.cStringBuilder, _
  ByRef sFetchDelCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsLargeCols As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectInsLargeCols2 As HRProSystemMgr.cStringBuilder, _
  ByRef sSelectDelLargeCols As HRProSystemMgr.cStringBuilder) As String
  
  
  Dim strCheckColumns As String
  Dim lngColumnID As Long
  
  With recEmailLinksColumnsEdit
    .Index = "idxLinkID"
    .Seek ">=", lngLinkID
    If Not .NoMatch Then

      Do While Not .EOF
        If !LinkID <> lngLinkID Then
          Exit Do
        End If

        lngColumnID = !ColumnID
        
        Select Case GetColumnDataType(lngColumnID)
        Case dtNUMERIC, dtINTEGER, dtBIT
          strCheckColumns = IIf(strCheckColumns <> vbNullString, strCheckColumns & vbNewLine & "             OR ", "") & _
                "isnull(@insCol_" & CStr(lngColumnID) & ",0) <> isnull(@delCol_" & CStr(lngColumnID) & ",0)"
        Case Else
          strCheckColumns = IIf(strCheckColumns <> vbNullString, strCheckColumns & vbNewLine & "             OR ", "") & _
                "isnull(@insCol_" & CStr(lngColumnID) & ",'') <> isnull(@delCol_" & CStr(lngColumnID) & ",'')"
        End Select
          
        AddColumnToTrigger _
              lngColumnID, alngAuditColumns, sDeclareInsCols, sDeclareDelCols, _
              sSelectInsCols2, sSelectDelCols, sFetchInsCols, sFetchDelCols, _
              sSelectInsLargeCols, sSelectInsLargeCols2, sSelectDelLargeCols
        
        .MoveNext
      Loop

    End If
  End With

  GetSQLForImmediateEmails = strCheckColumns

End Function


Public Function RemoveDuplicateDeclares(strSQL As String) As String

  Dim colDeclares As Collection
  Dim strSQLLines() As String
  Dim strLine As String
  Dim lngIndex As Long

  strSQLLines = Split(strSQL, vbCrLf)


  Set colDeclares = New Collection
  For lngIndex = 0 To UBound(strSQLLines)
  
    strLine = UCase(Trim(strSQLLines(lngIndex)))
    If Left(strLine, 8) = "DECLARE " Then
      If Exists(colDeclares, strLine) Then
        strSQLLines(lngIndex) = vbNullString
      Else
        colDeclares.Add strLine, strLine
      End If
    End If
  
  Next

  RemoveDuplicateDeclares = Join(strSQLLines, vbCrLf)

End Function


Public Function Exists(col As Collection, id As String) As Boolean
  On Local Error GoTo LocalErr
  Exists = (Trim(col(id)) <> vbNullString)
Exit Function
LocalErr:
  Exists = False
End Function


Private Function IsRelatedToSingleColumn(lngLinkID As Long) As Long

  Dim lngID As Long
  
  lngID = 0

  With recEmailLinksColumnsEdit
    .Index = "idxLinkID"
    .Seek ">=", lngLinkID
    If Not .NoMatch Then

      Do While Not .EOF
        If !LinkID <> lngLinkID Then
          Exit Do
        End If
        
        If lngID = 0 Then
          'First Column
          lngID = !ColumnID
        Else
          'Second Column
          lngID = 0
          Exit Do
        End If
        
        .MoveNext
      Loop

    End If
  End With

  IsRelatedToSingleColumn = lngID

End Function


Private Function GetInsertCommand(intType As EmailType, lngLinkID As Long, lngTableID As Long, blnImmediate As Boolean, blnRecalcRecDesc As Boolean, strColumn As String, strColumnValue As String, blnIncludeContent As Boolean) As String

  Dim strOutput As String

  strOutput = _
    "                    " & _
    "INSERT ASRSysEmailQueue" & _
    "(LinkID" & _
    ",TableID" & _
    ",RecordID" & _
    ",DateDue" & _
    ",UserName" & _
    ",[Immediate]" & _
    ",Type" & _
    ",RecalculateRecordDesc" & _
    ",RecordDesc" & _
    ",ColumnID" & _
    ",ColumnValue"

  If blnIncludeContent Then
    strOutput = strOutput & _
      ",RepTo,RepCc,RepBcc,[Subject],MsgText,Attachment"
  End If

  strOutput = strOutput & _
    ")" & vbNewLine & _
    "                    " & _
    "VALUES " & _
    "(" & CStr(lngLinkID) & _
    "," & CStr(lngTableID) & _
    ",@recordID" & _
    ",@emailDate" & _
    ",@username" & _
    "," & IIf(blnImmediate, "1", "0") & _
    "," & CStr(intType) & _
    "," & IIf(blnRecalcRecDesc, "1", "0") & _
    ",@recordDesc" & _
    "," & strColumn

  Select Case GetColumnDataType(val(strColumn))
  Case dtTIMESTAMP
    strOutput = strOutput & _
      ",isnull(convert(varchar(max)," & strColumnValue & "," & CStr(glngEmailDateFormat) & "),'')"
  Case dtBIT
    strOutput = strOutput & _
      ",CASE WHEN " & strColumnValue & " = 1 THEN 'Yes' ELSE 'No' END"
  Case Else
    strOutput = strOutput & _
      "," & strColumnValue
  End Select

  If blnIncludeContent Then
    strOutput = strOutput & _
      ",@To,@Cc,@Bcc,@Subject,@Message,@Attachment"
  End If
  
  strOutput = strOutput & ")"

  GetInsertCommand = strOutput


End Function
