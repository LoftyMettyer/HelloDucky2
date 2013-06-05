Attribute VB_Name = "modEmail"
'Option Explicit
'
'Private rsEmailLinks As ADODB.Recordset
'Private rsEmailRecipients As ADODB.Recordset
'Private rsEmailColumns As ADODB.Recordset
'Private rsLinkContent As ADODB.Recordset
'
'
''MH20090520
'Public Const strDelimStart As String = "«"   'asc = 171
'Public Const strDelimStop As String = "»"    'asc = 187
'
'
'Public glngEmailMethod As Long
'Public gstrEmailProfile As String
'Public gstrEmailServer As String
'Public gstrEmailAccount As String
'
'Public glngEmailDateFormat As Long
'Public gstrEmailAttachmentPath As String
'Public gstrEmailTestAddr As String
''Public gstrEmailEventLogToAddr As String
'
'Public strInsertEmailCode As String
'Public strUpdateEmailCode As String
'Public strDeleteEmailCode As String
'
'Private mstrRebuildCode As String
'
'Private mstrInsertEmailTemp As String
'Private mstrUpdateEmailTemp As String
'
'
''Public Function GetEmailCommand( _
''      lngMethod As Long, _
''      blnResult As Boolean, _
''      strServer As String, _
''      strAccount As String, _
''      strTo As String, _
''      strCc As String, _
''      strBcc As String, _
''      strSubject As String, _
''      strBody As String, _
''      strAttachment As String) As String
''
''  Dim sSQL As String
''
''
''
''  sSQL = IIf(blnResult, "EXEC @hResult = ", "EXEC ")
''
''  Select Case lngMethod
''  Case 0
''    sSQL = IIf(blnResult, "SET @hResult = 1", "PRINT '1'")
''
''  Case 1
''    sSQL = sSQL & _
''        "master..xp_sendmail " & vbCrLf & _
''        "@recipients=" & strTo & ", " & vbCrLf & _
''        "@copy_recipients=" & strCc & ", " & vbCrLf & _
''        "@blind_copy_recipients=" & strBcc & ", " & vbCrLf & _
''        "@subject=" & strSubject & ", " & vbCrLf & _
''        "@message=" & strBody & ", " & vbCrLf & _
''        "@attachments=" & strAttachment
''
''  Case 2
''    sSQL = sSQL & _
''        "msdb.dbo.sp_send_dbmail " & vbCrLf & _
''        "@recipients=" & strTo & ", " & vbCrLf & _
''        "@copy_recipients=" & strCc & ", " & vbCrLf & _
''        "@blind_copy_recipients=" & strBcc & ", " & vbCrLf & _
''        "@subject=" & strSubject & ", " & vbCrLf & _
''        "@body=" & strBody & ", " & vbCrLf & _
''        "@file_attachments=" & strAttachment
''
''  Case 3
''    sSQL = sSQL & _
''        "master..xp_SMTPsendmail80 " & vbCrLf & _
''        "@address=" & strServer & ", " & vbCrLf & _
''        "@from=" & strAccount & ", " & vbCrLf & _
''        "@recipient=" & strTo & ", " & vbCrLf & _
''        "@copy_recipients=" & strCc & ", " & vbCrLf & _
''        "@blind_copy_recipients=" & strBcc & ", " & vbCrLf & _
''        "@subject=" & strSubject & ", " & vbCrLf & _
''        "@body=" & strBody & ", " & vbCrLf & _
''        "@attachments=" & strAttachment
''
''  End Select
''
''  GetEmailCommand = sSQL
''
''End Function
'
'
'Public Function OpenEmailRecordsets()
'
'  Set rsEmailLinks = New ADODB.Recordset
'  rsEmailLinks.Open "ASRSysEmailLinks", gADOCon, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'
'  Set rsEmailRecipients = New ADODB.Recordset
'  rsEmailRecipients.Open "ASRSysEmailLinksRecipients", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
'
'  Set rsLinkContent = New ADODB.Recordset
'  rsLinkContent.Open "ASRSysLinkContent", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
'
'  Set rsEmailColumns = New ADODB.Recordset
'  rsEmailColumns.Open "ASRSysEmailLinksColumns", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
'
'End Function
'
''Public Function DeleteEmailLinks(lngTableID As Long) As Boolean
''
''  On Local Error GoTo LocalErr
''
''  With rsEmailLinks
''    rsEmailLinks.Requery          'MH20061010 Fault 11470
''    If Not (.BOF And .EOF) Then
''      .MoveFirst
''      Do While Not .EOF
''        If !TableID = lngTableID Then
''
''          DeleteLinkContent IIf(IsNull(!SubjectContentID), 0, !SubjectContentID)
''          DeleteLinkContent IIf(IsNull(!BodyContentID), 0, !BodyContentID)
''
''
''          With rsEmailRecipients
''            If Not .BOF Or Not .EOF Then
''              .MoveFirst
''              Do While Not .EOF
''                If !LinkID = rsEmailLinks!LinkID Then
''                  .Delete
''                End If
''                .MoveNext
''              Loop
''            End If
''          End With
''
''          With rsEmailColumns
''            If Not .BOF Or Not .EOF Then
''              .MoveFirst
''              Do While Not .EOF
''                If !LinkID = rsEmailLinks!LinkID Then
''                  .Delete
''                End If
''                .MoveNext
''              Loop
''            End If
''          End With
''
''          .Delete
''
''        End If
''        .MoveNext
''      Loop
''    End If
''  End With
''
''  DeleteEmailLinks = True
''
''Exit Function
''
''LocalErr:
''  MsgBox "Error deleting email links" & vbCrLf & Err.Description, vbCritical
''  DeleteEmailLinks = False
''
''End Function
'
'
'Public Function SaveEmailLinks(lngTableID As Long) As Boolean
'
'  On Local Error GoTo LocalErr
'
'  With recEmailLinksEdit
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'
'      Do While Not .EOF
'        If !TableID = lngTableID Then
'          If Not !Deleted Then
'            rsEmailLinks.AddNew
'            rsEmailLinks!LinkID = !LinkID
'            rsEmailLinks!TableID = !TableID
'            rsEmailLinks!Title = IIf(IsNull(!Title), vbNullString, !Title)
'            rsEmailLinks!FilterID = !FilterID
'            rsEmailLinks!EffectiveDate = !EffectiveDate
'            rsEmailLinks!Attachment = IIf(IsNull(!Attachment), vbNullString, !Attachment)
'            rsEmailLinks!Type = IIf(IsNull(!Type), 0, !Type)
'
'            rsEmailLinks!SubjectContentID = IIf(IsNull(!SubjectContentID), 0, !SubjectContentID)
'            rsEmailLinks!BodyContentID = IIf(IsNull(!BodyContentID), 0, !BodyContentID)
'
'            rsEmailLinks!RecordInsert = !RecordInsert
'            rsEmailLinks!RecordDelete = !RecordDelete
'            rsEmailLinks!RecordUpdate = !RecordUpdate
'
'            rsEmailLinks!DateColumnID = !DateColumnID
'            rsEmailLinks!DateOffset = !DateOffset
'            rsEmailLinks!DatePeriod = !DatePeriod
'
'            rsEmailLinks.Update
'            rsEmailLinks.MoveLast
'
'
'            SaveLinkContent IIf(IsNull(recEmailLinksEdit!SubjectContentID), 0, recEmailLinksEdit!SubjectContentID)
'            SaveLinkContent IIf(IsNull(recEmailLinksEdit!BodyContentID), 0, recEmailLinksEdit!BodyContentID)
'
'
'            ' Add references to email recipients
'            With recEmailRecipientsEdit
'              If Not (.BOF And .EOF) Then
'                .MoveFirst
'
'                Do While Not .EOF
'                  If !LinkID = recEmailLinksEdit!LinkID Then
'                    rsEmailRecipients.AddNew
'                    rsEmailRecipients!LinkID = recEmailLinksEdit!LinkID
'                    rsEmailRecipients!RecipientID = !RecipientID
'                    rsEmailRecipients!Mode = !Mode
'                    rsEmailRecipients.Update
'                  End If
'
'                  .MoveNext
'                Loop
'              End If
'            End With
'
'
'            ' Add references to email recipients
'            With recEmailLinksColumnsEdit
'              If Not (.BOF And .EOF) Then
'                .MoveFirst
'
'                Do While Not .EOF
'                  If !LinkID = recEmailLinksEdit!LinkID Then
'                    rsEmailColumns.AddNew
'                    rsEmailColumns!LinkID = recEmailLinksEdit!LinkID
'                    rsEmailColumns!ColumnID = !ColumnID
'                    rsEmailColumns.Update
'                  End If
'
'                  .MoveNext
'                Loop
'              End If
'            End With
'
'          End If
'        End If
'
'        .MoveNext
'      Loop
'    End If
'  End With
'
'  SaveEmailLinks = True
'
'Exit Function
'
'LocalErr:
'  MsgBox "Error saving email links" & vbCrLf & Err.Description, vbCritical
'  SaveEmailLinks = False
'
'End Function
'
'
''Public Function SaveEmailLinksForColumn(lngColumnID As Long) As Boolean
''
''  With rsEmailLinks
''    rsEmailLinks.Requery          'MH20061010 Fault 11470
''    If Not (.BOF And .EOF) Then
''      .MoveFirst
''      Do While Not .EOF
''        If !ColumnID = lngColumnID Then
''
''          DeleteLinkContent IIf(IsNull(!SubjectContentID), 0, !SubjectContentID)
''          DeleteLinkContent IIf(IsNull(!BodyContentID), 0, !BodyContentID)
''
''
''          With rsEmailRecipients
''            If Not .BOF Or Not .EOF Then
''              .MoveFirst
''              Do While Not .EOF
''                If !LinkID = rsEmailLinks!LinkID Then
''                  .Delete
''                End If
''                .MoveNext
''              Loop
''            End If
''          End With
''
''          .Delete
''        End If
''        .MoveNext
''      Loop
''    End If
''  End With
''
''
''  With recEmailLinksEdit
''    If Not (.BOF And .EOF) Then
''      .MoveFirst
''
''      Do While Not .EOF
''        If !ColumnID = lngColumnID Then
''          rsEmailLinks.AddNew
''          rsEmailLinks!LinkID = !LinkID
''          rsEmailLinks!ColumnID = !ColumnID
''          rsEmailLinks!Title = IIf(IsNull(!Title), vbNullString, !Title)
''          rsEmailLinks!FilterID = !FilterID
''          rsEmailLinks!Immediate = !Immediate
''          rsEmailLinks!Offset = !Offset
''          rsEmailLinks!Period = !Period
''          rsEmailLinks!EffectiveDate = !EffectiveDate
''          rsEmailLinks!Subject = IIf(IsNull(!Subject), vbNullString, !Subject)
''          'rsEmailLinks!Importance = !Importance
''          'rsEmailLinks!Sensitivity = !Sensitivity
''          rsEmailLinks!IncRecDesc = !IncRecDesc
''          rsEmailLinks!IncColDetail = !IncColDetail
''          rsEmailLinks!IncUsername = !IncUsername
''          rsEmailLinks!EmailInsert = !EmailInsert
''          rsEmailLinks!EmailDelete = !EmailDelete
''          rsEmailLinks!EmailUpdate = !EmailUpdate
''
''          rsEmailLinks!Body = IIf(IsNull(!Body), vbNullString, !Body)
''          rsEmailLinks!Attachment = IIf(IsNull(!Attachment), vbNullString, !Attachment)
''
''          'MH20090521
''          rsEmailLinks!SubjectContentID = IIf(IsNull(!SubjectContentID), 0, !SubjectContentID)
''          rsEmailLinks!BodyContentID = IIf(IsNull(!BodyContentID), 0, !BodyContentID)
''
''          rsEmailLinks.Update
''          rsEmailLinks.MoveLast
''
''
''          SaveLinkContent IIf(IsNull(recEmailLinksEdit!SubjectContentID), 0, recEmailLinksEdit!SubjectContentID)
''          SaveLinkContent IIf(IsNull(recEmailLinksEdit!BodyContentID), 0, recEmailLinksEdit!BodyContentID)
''
''
''          ' Add references to email recipients
''          With recEmailRecipientsEdit
''            If Not (.BOF And .EOF) Then
''              .MoveFirst
''
''              Do While Not .EOF
''                If !LinkID = recEmailLinksEdit!LinkID Then
''                  rsEmailRecipients.AddNew
''                  rsEmailRecipients!LinkID = recEmailLinksEdit!LinkID
''                  rsEmailRecipients!RecipientID = !RecipientID
''                  rsEmailRecipients!Mode = !Mode
''                  rsEmailRecipients.Update
''                End If
''
''                .MoveNext
''              Loop
''            End If
''          End With
''
''        End If
''
''        .MoveNext
''      Loop
''    End If
''  End With
''
''End Function
'
'
'Private Function DeleteLinkContent(lngContentID As Long)
'
'  If lngContentID > 0 Then
'
'    With rsLinkContent
'      If Not .BOF Or Not .EOF Then
'        .MoveFirst
'        Do While Not .EOF
'          If !ContentID = lngContentID Then
'            rsLinkContent.Delete
'          End If
'          rsLinkContent.MoveNext
'        Loop
'      End If
'
'    End With
'
'  End If
'
'End Function
'
'
'Private Function SaveLinkContent(lngContentID As Long)
'
'  If lngContentID > 0 Then
'
'    With recLinkContentEdit
'      .Index = "idxContentIDSequence"
'      .Seek ">=", lngContentID, 0
'
'      If Not .NoMatch Then
'        Do While Not .EOF
'
'          If !ContentID <> lngContentID Then
'            Exit Do
'          End If
'
'          rsLinkContent.AddNew
'          rsLinkContent!id = !id
'          rsLinkContent!ContentID = !ContentID
'          rsLinkContent!Sequence = !Sequence
'          rsLinkContent!FixedText = !FixedText
'          rsLinkContent!FieldCode = !FieldCode
'          rsLinkContent!FieldID = !FieldID
'          rsLinkContent.Update
'
'          .MoveNext
'        Loop
'      End If
'
'    End With
'
'  End If
'
'End Function
'
'
'
'Public Function CloseEmailRecordsets()
'
'  If rsEmailLinks.State <> adStateClosed Then
'    rsEmailLinks.Close
'  End If
'  Set rsEmailLinks = Nothing
'
'  If rsEmailRecipients.State <> adStateClosed Then
'    rsEmailRecipients.Close
'  End If
'  Set rsEmailRecipients = Nothing
'
'  If rsEmailColumns.State <> adStateClosed Then
'    rsEmailColumns.Close
'  End If
'  Set rsEmailColumns = Nothing
'
'  If rsLinkContent.State <> adStateClosed Then
'    rsLinkContent.Close
'  End If
'  Set rsLinkContent = Nothing
'
'End Function
'
'
'
'Public Function CreateEmailSendStoredProcedure(strProcName As String, lngMethod As Long, strProfile As String, strServer As String, strAccount As String) As Boolean
'
'  Dim sSQL As String
'  Dim fOK As Boolean
'
'  On Local Error GoTo LocalErr
'
'  sSQL = vbNullString
'  If GetSystemSetting("email", "qa info", 0) = 1 Then
'
'    Select Case lngMethod
'    Case 1: sSQL = "xp_sendmail"
'    Case 2: sSQL = "sp_send_dbmail"
'    Case 3: sSQL = "xp_SMTPsendmail80"
'    End Select
'
'    'sSQL = "SELECT @Message = @Message+char(13)+char(13)+'QA Info: '+@@SERVERNAME+'\'+DB_NAME()+' (" & sSQL & ")'" & vbCrLf
'    sSQL = "SELECT @Message = @Message+char(13)+char(13)+'QA Info: '+CONVERT(varchar, SERVERPROPERTY('servername'))+'\'+DB_NAME()+' (" & sSQL & ")'" & vbCrLf
'  End If
'
'
'  Select Case lngMethod
'  Case 0
'    sSQL = "SET @hResult = 0"   'MH20061219 Mark as sent as per QA
'
'  Case 1
'    'MH20071026 Fault 12555
'    GrantExecuteToPublic "master", "xp_startmail"
'    GrantExecuteToPublic "master", "xp_sendmail"
'
'    sSQL = sSQL & _
'        "EXEC @hResult = master..xp_sendmail " & vbCrLf & _
'        "@recipients=@To, " & vbCrLf & _
'        "@copy_recipients=@CC, " & vbCrLf & _
'        "@blind_copy_recipients=@BCC, " & vbCrLf & _
'        "@subject=@Subject, " & vbCrLf & _
'        "@message=@Message, " & vbCrLf & _
'        "@attachments=@Attachment"
'
'  Case 2
'    'MH20071026 Fault 12555
'    GrantExecuteToPublic "msdb", "sp_send_dbmail"
'
'    ' AE20080215 Fault #12834
''    If Trim(sNewProfile) <> "<Use Default Profile>" And _
''        Trim(sNewProfile) <> "" Then
''      sNewProfile = "@profile_name = '" & sNewProfile & "' ," & vbCrLf
''    Else
''      sNewProfile = vbNullString
''    End If
'    Dim sNewProfile As String
'    sNewProfile = strProfile
'
'    If Trim(sNewProfile) = "<Use Default Profile>" Then
'      sNewProfile = vbNullString
'    ElseIf Trim(sNewProfile) <> vbNullString Then
'      sNewProfile = "@profile_name = '" & sNewProfile & "' ," & vbCrLf
'    End If
'
'    sSQL = sSQL & _
'        "EXEC @hResult = msdb.dbo.sp_send_dbmail " & vbCrLf & _
'        sNewProfile & _
'        "@recipients=@To, " & vbCrLf & _
'        "@copy_recipients=@CC, " & vbCrLf & _
'        "@blind_copy_recipients=@BCC, " & vbCrLf & _
'        "@subject=@Subject, " & vbCrLf & _
'        "@body=@Message, " & vbCrLf & _
'        "@file_attachments=@Attachment"
'
'  Case 3
'    'MH20071026 Fault 12555
'    GrantExecuteToPublic "master", "xp_SMTPsendmail80"
'
'    sSQL = sSQL & _
'        "EXEC @hResult = master..xp_SMTPsendmail80 " & vbCrLf & _
'        "@address='" & Replace(strServer, "'", "''") & "', " & vbCrLf & _
'        "@from='" & Replace(strAccount, "'", "''") & "', " & vbCrLf & _
'        "@recipient=@To, " & vbCrLf & _
'        "@copy_recipients=@CC, " & vbCrLf & _
'        "@blind_copy_recipients=@BCC, " & vbCrLf & _
'        "@subject=@Subject, " & vbCrLf & _
'        "@body=@Message, " & vbCrLf & _
'        "@attachments=@Attachment"
'
'  End Select
'
'  DropProcedure strProcName
'  sSQL = "CREATE PROCEDURE dbo.[" & strProcName & "](" & vbCrLf & _
'         "  @hResult int OUTPUT," & vbCrLf & _
'         "  @To varchar(8000)," & vbCrLf & _
'         "  @CC varchar(8000)," & vbCrLf & _
'         "  @BCC varchar(8000)," & vbCrLf & _
'         "  @Subject varchar(8000)," & vbCrLf & _
'         "  @Message varchar(8000)," & vbCrLf & _
'         "  @Attachment varchar(8000))" & vbCrLf & _
'         "AS " & vbCrLf & _
'         "BEGIN " & vbCrLf & _
'         sSQL & vbCrLf & _
'         "END"
'  gADOCon.Execute sSQL, , adExecuteNoRecords
'
'  fOK = True
'
'TidyUpAndExit:
'  CreateEmailSendStoredProcedure = fOK
'  Exit Function
'
'LocalErr:
'  OutputError "Error creating email send procedure"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'
''MH20071026 Fault 12555
'Public Sub GrantExecuteToPublic(strDatabase As String, strSPName As String)
'
'  On Local Error Resume Next
'
'  'JPD 20080714 Fault 13265
'  If gbIsUserSystemAdmin Then
'    gADOCon.Execute "USE [" & strDatabase & "]", , adExecuteNoRecords
'    gADOCon.Execute "GRANT EXECUTE ON " & strSPName & " TO PUBLIC", , adExecuteNoRecords
'    gADOCon.Execute "USE [" & gsDatabaseName & "]", , adExecuteNoRecords
'  End If
'
'End Sub
'
'
'Public Function CreateEmailAddrStoredProcedure() As Boolean
'
'  Const strSPName As String = "spASRSysEmailAddr"
'  Dim strSQL As String
'  Dim strTableName As String
'  Dim strColumnName As String
'  Dim strSPEmailCalc As String
'  Dim fOK As Boolean
'
'  On Error GoTo ErrorTrap
'
'  CreateEmailSendStoredProcedure "spASRSendMail", glngEmailMethod, gstrEmailProfile, gstrEmailServer, gstrEmailAccount
'
'  fOK = True
'
''  ' Drop any existing stored procedure.
''  strSQL = "IF EXISTS" & _
''           " (SELECT Name" & _
''           "   FROM sysobjects" & _
''           "   WHERE id = object_id('" & strSPName & "')" & _
''           "     AND sysstat & 0xf = 4)" & _
''           " DROP PROCEDURE dbo." & strSPName
''  gADOCon.Execute strSQL, , adExecuteNoRecords
'  DropProcedure strSPName
'
'
'  strSQL = vbNullString
'  With recEmailAddrEdit
'    .Index = "idxID"
'
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'
'    Do While Not .EOF
'
'      If Not !Deleted Then
'        If !Type <> 0 Then
'          strSQL = strSQL & _
'            IIf(strSQL <> vbNullString, "ELSE ", vbNullString) & _
'            "IF @EmailID = " & CStr(!EmailID) & vbNewLine & _
'            "BEGIN" & vbNewLine
'
'          Select Case !Type
'          Case 1    'Column
'            strTableName = GetTableName(!TableID)
'            strColumnName = GetColumnName(!ColumnID)
'
'            If strTableName <> vbNullString And strColumnName <> vbNullString Then
'              strSQL = strSQL & _
'                "    /* " & Trim(!Name) & " (Column) */" & vbNewLine & _
'                "    SET @hResult = (SELECT ltrim(rtrim(" & strColumnName & ")) FROM " & strTableName & " WHERE ID = @recordID)"
'            End If
'
'          Case 2    'Calculated
'
'            strSPEmailCalc = "sp_ASRExpr_" & CStr(!ExprID)
'
'            strSQL = strSQL & _
'              "    /* " & Trim(!Name) & " (Calculated) */" & vbNewLine & _
'              "    IF EXISTS (SELECT Name FROM sysobjects WHERE type = 'P'" & _
'              "        AND name = '" & strSPEmailCalc & "')" & vbNewLine & _
'              "    BEGIN" & vbNewLine & _
'              "        EXEC @hResult = " & strSPEmailCalc & " @EmailAddr OUTPUT, @recordID" & vbNewLine & _
'              "        IF @hResult <> 0 SET @EmailAddr = ''" & vbNewLine & _
'              "        SET @hResult = ltrim(rtrim(CONVERT(varchar(255), @EmailAddr)))" & vbNewLine & _
'              "    END"
'
'
'          End Select
'
'          strSQL = strSQL & vbNewLine & _
'            "END" & vbNewLine & vbNewLine
'
'        End If
'      End If
'
'      .MoveNext
'
'    Loop
'  End With
'
'
'  strSQL = _
'    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
'    "/* email address stored procedure.                       */" & vbNewLine & _
'    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
'    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
'    "CREATE PROCEDURE dbo." & strSPName & vbNewLine & _
'    "(" & vbNewLine & _
'    "    @hResult varchar(8000) OUTPUT," & vbNewLine & _
'    "    @EmailID integer," & vbNewLine & _
'    "    @recordID integer" & vbNewLine & _
'    ")" & vbNewLine & _
'    "AS" & vbNewLine & _
'    "BEGIN" & vbNewLine & vbNewLine & _
'    "DECLARE @EmailAddr char(255)" & vbNewLine & vbNewLine & _
'    strSQL & vbNewLine & _
'    IIf(Len(strSQL) > 0, vbNewLine & "ELSE" & vbNewLine, vbNullString) & _
'    "    SET @hResult = (SELECT ltrim(rtrim(Fixed)) From ASRSysEmailAddress WHERE EmailID = @EmailID)" & vbNewLine & _
'    "END"
'  gADOCon.Execute strSQL, , adExecuteNoRecords
'
'
'TidyUpAndExit:
'  CreateEmailAddrStoredProcedure = fOK
'  Exit Function
'
'ErrorTrap:
'  OutputError "Error creating email addresses"
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'
'Private Function GetColumnName(lngColumnID As Long) As String
'
'  On Error GoTo ErrorTrap
'
'  GetColumnName = vbNullString
'
'  With recColEdit
'    .Index = "idxColumnID"
'    .Seek "=", lngColumnID
'
'    If Not .NoMatch Then
'      GetColumnName = !ColumnName
'    End If
'  End With
'
'TidyUpAndExit:
'  Exit Function
'
'ErrorTrap:
'  Resume TidyUpAndExit
'
'End Function
'
'
'Private Function GetSQLFilter(lngFilterID As Long, sCurrentTable As String) As String
'
'  Dim fOK As Boolean
'  Dim objExpr As CExpression
'  Dim strFilterRunTimeCode As String
'
'  GetSQLFilter = vbNullString
'
'  'Filter
'  Set objExpr = New CExpression
'  With objExpr
'
'    objExpr.ExpressionID = lngFilterID
'    objExpr.ConstructExpression
'    fOK = objExpr.RuntimeFilterCode(strFilterRunTimeCode, False)
'
'    strFilterRunTimeCode = Replace(strFilterRunTimeCode, vbNewLine, " ")
'
'    GetSQLFilter = "@recordID IN " & _
'          "(" & strFilterRunTimeCode & ")"
'
'  End With
'  Set objExpr = Nothing
'
'End Function
'
'
'Private Function GetSQLRecipients(lngLinkID As Long) As String
'
'  Dim recEmailRecipients As New ADODB.Recordset
'  Dim strResult As String
'  Dim strSQL As String
'
'  strResult = vbNullString
'  With recEmailRecipientsEdit
'
'    If Not .BOF Or Not .EOF Then
'      .MoveFirst
'      Do While Not .EOF
'
'        If !LinkID = lngLinkID And !RecipientID > 0 Then
'
'          'strResult = strResult & _
'              "      EXEC @hResult = " & gsEMAILADDR & CStr(!RecipientID) & _
'              " @TempRecip OUTPUT, @recordID" & vbNewline & _
'              "      SELECT "
'
'          strResult = strResult & _
'              "      EXEC @hResult = dbo.spASRSysEmailAddr " & _
'              " @TempRecip OUTPUT, " & CStr(!RecipientID) & ", @recordID" & vbNewLine & _
'              "      SELECT "
'
'          Select Case !Mode
'          Case 0: strResult = strResult & "@RecipTo = @RecipTo"
'          Case 1: strResult = strResult & "@RecipCc = @RecipCc"
'          Case 2: strResult = strResult & "@RecipBcc = @RecipBcc"
'          End Select
'
'          strResult = strResult & " + RTrim(@TempRecip) + ';'" & vbNewLine & vbNewLine
'
'        End If
'
'        .MoveNext
'      Loop
'    End If
'
'  End With
'
'  If strResult <> vbNullString Then
'    strResult = "      SELECT @RecipTo = ''" & vbNewLine & _
'                "      SELECT @RecipCc = ''" & vbNewLine & _
'                "      SELECT @RecipBcc = ''" & vbNewLine & vbNewLine & _
'                strResult & vbNewLine & _
'                "      IF rtrim(replace(@RecipTo,';','')) = ''" & vbNewLine & _
'                "        RETURN 1" & vbNewLine & vbNewLine
'  End If
'
'  GetSQLRecipients = strResult
'
'End Function
'
'
'Public Sub CreateEmailProcsForTable(pLngCurrentTableID As Long, _
'  sCurrentTable As String, _
'  lngRecordDescExprID As Long, _
'  ByRef alngAuditColumns As Variant, _
'  ByRef sDeclareInsCols As SystemMgr.cStringBuilder, _
'  ByRef sDeclareDelCols As SystemMgr.cStringBuilder, _
'  ByRef sSelectInsCols2 As SystemMgr.cStringBuilder, _
'  ByRef sSelectDelCols As SystemMgr.cStringBuilder, _
'  ByRef sFetchInsCols As SystemMgr.cStringBuilder, _
'  ByRef sFetchDelCols As SystemMgr.cStringBuilder, _
'  ByRef sSelectInsLargeCols As SystemMgr.cStringBuilder, _
'  ByRef sSelectInsLargeCols2 As SystemMgr.cStringBuilder, _
'  ByRef sSelectDelLargeCols As SystemMgr.cStringBuilder)
'  ' JPD20020913 - instead of making multiple queries to the triggered table, and
'  ' the 'inserted' and 'deleted' tables, we now get all of the required information in
'  ' the cursor that we used to loop through to get just the id of each record being
'  ' inserted/updated/deleted.
'  ' Here we are passed a number of variables and an array from the table trigger creation
'  ' code so that the email columns can be added to the SELECT statement that is used
'  ' to create the cursor, the FETCH statement that used to loop through the cursor,
'  ' and the DECLARE statements that are needed.
'  ' The email check code is modified for the new implementation.
'  ' NB. an array of columns that have been added to the SELECT statement is used
'  ' to ensure that columns aren't added more than once. Audit columns, email columns
'  ' and calculated columns all use this method.
'  ' This change was driven by the performance degradation reported by
'  ' Islington.
'
'  Dim strTemp As String
'
'  On Error GoTo LocalErr
'
'  mstrRebuildCode = vbNullString
''  strInsertEmailCode = vbNullString
'  strUpdateEmailCode = vbNullString
'
'  With recColEdit
'    .Index = "idxColumnID"    'Fix for fault 1334
'    .MoveFirst
'
'    If Not .NoMatch Then
'      Do While Not .EOF
'        If !TableID = pLngCurrentTableID Then
'
'          ' JPD20020913 - instead of making multiple queries to the triggered table, and
'          ' the 'inserted' and 'deleted' tables, we now get all of the required information in
'          ' the cursor that we used to loop through to get just the id of each record being
'          ' inserted/updated/deleted.
'          ' Here we are passing a number of variables and an array to the email trigger creation
'          ' code so that the email columns can be added to the SELECT statement that is used
'          ' to create the cursor, the FETCH statement that used to loop through the cursor,
'          ' and the DECLARE statements that are needed.
'          ' The email check code is modified for the new implementation.
'          ' NB. an array of columns that have been added to the SELECT statement is used
'          ' to ensure that columns aren't added more than once. Audit columns, email columns
'          ' and calculated columns all use this method.
'          ' This change was driven by the performance degradation reported by
'          ' Islington.
'
''''
'''' MH 18/06/2009 - Comment this out until I implement the new stuff
''''
''''          CreateEmailProcsForColumn pLngCurrentTableID, sCurrentTable, !ColumnID, lngRecordDescExprID, alngAuditColumns, _
''''            sDeclareInsCols, sDeclareDelCols, _
''''            sSelectInsCols2, sSelectDelCols, _
''''            sFetchInsCols, sFetchDelCols, _
''''            sSelectInsLargeCols, sSelectInsLargeCols2, sSelectDelLargeCols
''''
''''          strInsertEmailCode = strInsertEmailCode & mstrInsertEmailTemp
''''          strUpdateEmailCode = strUpdateEmailCode & mstrUpdateEmailTemp
'
'        End If
'        .MoveNext
'      Loop  'Next Column
'    End If
'  End With
'
'
''  strTemp = "IF EXISTS" & _
''            " (SELECT Name" & _
''            "   FROM sysobjects" & _
''            "   WHERE id = object_id('dbo.spASREmailRebuild_" & CStr(pLngCurrentTableID) & "')" & _
''            "     AND sysstat & 0xf = 4)" & _
''            " DROP PROCEDURE dbo.spASREmailRebuild_" & CStr(pLngCurrentTableID)
''  gADOCon.Execute strTemp, , adExecuteNoRecords
'  DropProcedure "spASREmailRebuild_" & CStr(pLngCurrentTableID)
'
'
'  If strUpdateEmailCode <> vbNullString Then
'
'    strTemp = _
'        "            DECLARE @emailDate datetime" & vbNewLine & _
'        "            DECLARE @purgeDate datetime" & vbNewLine & _
'        "            DECLARE @LastSent varchar(8000)" & vbNewLine & _
'        "            DECLARE @username varchar(50)" & vbNewLine & _
'        "            SELECT @username = rtrim(system_user)" & vbNewLine & _
'        "            EXEC sp_ASRPurgeDate @purgedate OUTPUT, 'EMAIL'" & vbNewLine & vbNewLine
'
'    strInsertEmailCode = strTemp & strInsertEmailCode
'    strUpdateEmailCode = strTemp & strUpdateEmailCode
'
'  End If
'
'
'  If mstrRebuildCode <> vbNullString Then
'
'    mstrRebuildCode = _
'      "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
'      "/* email address stored procedure.                       */" & vbNewLine & _
'      "/* Automatically generated by the System Manager.   */" & vbNewLine & _
'      "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
'      "CREATE PROCEDURE dbo.spASREmailRebuild_" & CStr(pLngCurrentTableID) & vbNewLine & _
'      "(@recordid int)" & vbNewLine & _
'      "AS" & vbNewLine & _
'      "BEGIN" & vbNewLine & vbNewLine & _
'      "        DECLARE @strTemp varchar(8000)" & vbNewLine & _
'               strTemp & mstrRebuildCode & vbNewLine & _
'      "END"
'
''    gADOCon.Execute "IF EXISTS (SELECT Name FROM sysobjects" & _
''                   "   WHERE id = object_id('spASREmailRebuild_" & CStr(pLngCurrentTableID) & "')" & _
''                   "     AND sysstat & 0xf = 4)" & _
''                   " DROP PROCEDURE spASREmailRebuild_" & CStr(pLngCurrentTableID), , adExecuteNoRecords
'    DropProcedure "spASREmailRebuild_" & CStr(pLngCurrentTableID)
'
'    gADOCon.Execute mstrRebuildCode, , adExecuteNoRecords
'
'  End If
'
'Exit Sub
'
'LocalErr:
'  If ASRDEVELOPMENT Then
'    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
'    Stop
'  End If
'
'End Sub
'
'
'Private Sub CreateEmailProcsForColumn(lngTableID As Long, _
'  sCurrentTable As String, _
'  lngColumnID As Long, _
'  lngRecordDescExprID As Long, _
'  ByRef alngAuditColumns As Variant, _
'  ByRef sDeclareInsCols As SystemMgr.cStringBuilder, _
'  ByRef sDeclareDelCols As SystemMgr.cStringBuilder, _
'  ByRef sSelectInsCols2 As SystemMgr.cStringBuilder, _
'  ByRef sSelectDelCols As SystemMgr.cStringBuilder, _
'  ByRef sFetchInsCols As SystemMgr.cStringBuilder, _
'  ByRef sFetchDelCols As SystemMgr.cStringBuilder, _
'  ByRef sSelectInsLargeCols As SystemMgr.cStringBuilder, _
'  ByRef sSelectInsLargeCols2 As SystemMgr.cStringBuilder, _
'  ByRef sSelectDelLargeCols As SystemMgr.cStringBuilder)
'  ' JPD20020913 - instead of making multiple queries to the triggered table, and
'  ' the 'inserted' and 'deleted' tables, we now get all of the required information in
'  ' the cursor that we used to loop through to get just the id of each record being
'  ' inserted/updated/deleted.
'  ' Here we are passed a number of variables and an array from the table trigger creation
'  ' code so that the email columns can be added to the SELECT statement that is used
'  ' to create the cursor, the FETCH statement that used to loop through the cursor,
'  ' and the DECLARE statements that are needed.
'  ' The email check code is modified for the new implementation.
'  ' NB. an array of columns that have been added to the SELECT statement is used
'  ' to ensure that columns aren't added more than once. Audit columns, email columns
'  ' and calculated columns all use this method.
'  ' This change was driven by the performance degradation reported by
'  ' Islington.
'
'  'Dim recEmailLinks As rdo.rdoResultset
'  Dim strSQL As String
'  Dim strColumnName As String
'  Dim sConvertInsCols As String
'  Dim fColFound As Boolean
'  Dim iLoop As Integer
'  Dim iDataType As Integer
'  Dim iSize As Integer
'  Dim iDecimals As Integer
'  Dim lngLinkID As Long
'  Dim sTemp As String
'
'  On Error GoTo LocalErr
'
'  sConvertInsCols = ""
'
'  mstrInsertEmailTemp = vbNullString
'  mstrUpdateEmailTemp = vbNullString
'
'  strColumnName = recColEdit!ColumnName
'  iDataType = recColEdit!DataType
'  iSize = recColEdit!Size
'  iDecimals = recColEdit!Decimals
'
'  'Loop through all of the email links for this column
'  With recEmailLinksEdit
'
'    If Not .BOF Or Not .EOF Then
'      .Index = "idxID"    'MH20031104 Fault 7469
'      .MoveFirst
'
'      Do While Not .EOF
'
'        If !ColumnID = lngColumnID Then
'          mstrInsertEmailTemp = mstrInsertEmailTemp & _
'              CreateEmailProcsForLink(lngTableID, sCurrentTable, !ColumnID, lngRecordDescExprID)
'        End If
'
'        .MoveNext
'      Loop  'Next link
'
'    End If
'
'  End With
'
'
'  If mstrInsertEmailTemp <> vbNullString Then
'
'    ' JPD20020913 - instead of making multiple queries to the triggered table, and
'    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
'    ' the cursor that we used to loop through to get just the id of each record being
'    ' inserted/updated/deleted.
'    ' Here we are checking the old and new values of the email columns.
'    ' This change was driven by the performance degradation reported by
'    ' Islington.
'
'    ''Need to do a compare on insert so that if a link is set up on
'    ''leaving date a mail is not sent as soon as a record is inserted
'    ''"        SELECT @newValue = CONVERT(varchar(255), " & strColumnName & ") FROM inserted WHERE id = @recordID" & vbNewline
'    'mstrInsertEmailTemp = vbNewline & _
'      "        /* " & UCase(strColumnName) & " */" & vbNewline & _
'      "        SELECT @newValue = CONVERT(varchar(255), " & strColumnName & ") FROM " & sCurrentTable & " WHERE id = @recordID" & vbNewline & _
'      "        SELECT @oldValue = CONVERT(varchar(255), " & strColumnName & ") FROM deleted WHERE id = @recordID" & vbNewline & _
'      vbNewline & _
'      "        IF ISNULL(@oldValue,'') <> ISNULL(@newValue,'')" & vbNewline & _
'      "        BEGIN" & vbNewline & _
'      vbNewline & _
'      "          DELETE FROM ASRSYSEmailQueue" & vbNewline & _
'      "          WHERE datesent IS Null AND recordid = @recordid AND ColumnID = " & CStr(lngColumnID) & vbNewline & _
'      vbNewline & _
'      mstrInsertEmailTemp & vbNewline & _
'      "        END" & vbNewline
'
'
'    'MH20040707 We still need to check the old and new values for inserts as well as updates.
'    'This is because if you don't populate the leaving date when adding a new record then
'    'it shouldn't send an email notification about the leaving date!  With this code in place
'    'it compares the new value (null) with the old value (null) and decides that it hasn't
'    'changed so doesn't need to send an email.
'    mstrInsertEmailTemp = vbNewLine & _
'      "            IF (@insCol_" & Trim$(Str$(lngColumnID)) & " <> @delCol_" & Trim$(Str$(lngColumnID)) & ") OR " & vbNewLine & _
'      "                ((@insCol_" & Trim$(Str$(lngColumnID)) & " IS null) AND (NOT @delCol_" & Trim$(Str$(lngColumnID)) & " IS null)) OR " & vbNewLine & _
'      "                ((NOT @insCol_" & Trim$(Str$(lngColumnID)) & " IS null) AND (@delCol_" & Trim$(Str$(lngColumnID)) & " IS null))" & vbNewLine & _
'      "            BEGIN" & vbNewLine & _
'      "                DELETE FROM ASRSYSEmailQueue" & vbNewLine & _
'      "                WHERE datesent IS Null AND recordid = @recordid AND ColumnID = " & CStr(lngColumnID) & vbNewLine & _
'      mstrInsertEmailTemp & vbNewLine & _
'      "            END" & vbNewLine
'
'    ' Now we split insert and update - they begin to get different
'    mstrUpdateEmailTemp = mstrInsertEmailTemp
'
'    ' JPD20020913 - instead of making multiple queries to the triggered table, and
'    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
'    ' the cursor that we used to loop through to get just the id of each record being
'    ' inserted/updated/deleted.
'    ' Here we are adding the email columns to the SELECT statement that is used
'    ' to create the cursor, the FETCH statement that used to loop through the cursor,
'    ' and the DECLARE statements that are needed.
'    ' The email check code is modified for the new implementation.
'    ' NB. an array of columns that have been added to the SELECT statement is used
'    ' to ensure that columns aren't added more than once. As well as audit columns,
'    ' we're also going to add email columns and calculated columns later on.
'    ' This change was driven by the performance degradation reported by
'    ' Islington.
'    fColFound = False
'
'    ' Check if the column has already been declared and added to the select and fetch strings
'    For iLoop = 1 To UBound(alngAuditColumns)
'      If alngAuditColumns(iLoop) = lngColumnID Then
'        fColFound = True
'        Exit For
'      End If
'    Next iLoop
'
'    If Not fColFound Then
'      ReDim Preserve alngAuditColumns(UBound(alngAuditColumns) + 1)
'      alngAuditColumns(UBound(alngAuditColumns)) = lngColumnID
'
'      If (iDataType <> dtVARCHAR) Or (iSize <= VARCHARTHRESHOLD) Then
'
'        'sSelectInsCols.Append "," & vbNewLine & "        inserted." & strColumnName
'        sSelectInsCols2.Append ",@insCol_" & Trim$(Str$(lngColumnID)) & "=" & strColumnName
'        sSelectDelCols.Append "," & vbNewLine & "        deleted." & strColumnName
'
'        'sFetchInsCols.Append "," & vbNewLine & "        @insCol_" & Trim$(Str$(lngColumnID))
'        sFetchDelCols.Append "," & vbNewLine & "        @delCol_" & Trim$(Str$(lngColumnID))
'      Else
'        sSelectInsLargeCols.Append ",@insCol_" & Trim$(Str$(lngColumnID)) & "=inserted." & strColumnName
'        sSelectInsLargeCols2.Append ",@insCol_" & Trim$(Str$(lngColumnID)) & "=" & strColumnName
'        sSelectDelLargeCols.Append ",@delCol_" & Trim$(Str$(lngColumnID)) & "=deleted." & strColumnName
'      End If
'
'      sDeclareInsCols.Append "," & vbNewLine & "        @insCol_" & Trim$(Str$(lngColumnID))
'      sDeclareDelCols.Append "," & vbNewLine & "        @delCol_" & Trim$(Str$(lngColumnID))
'    End If
'
'    Select Case iDataType
'      Case dtVARCHAR
'        If Not fColFound Then
'          sDeclareInsCols.Append " varchar(" & Trim$(Str$(iSize)) & ")"
'          sDeclareDelCols.Append " varchar(" & Trim$(Str$(iSize)) & ")"
'        End If
'        sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"
'
'      Case dtLONGVARCHAR
'        If Not fColFound Then
'          sDeclareInsCols.Append " varchar(14)"
'          sDeclareDelCols.Append " varchar(14)"
'        End If
'        sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"
'
'      Case dtINTEGER
'        If Not fColFound Then
'          sDeclareInsCols.Append " integer"
'          sDeclareDelCols.Append " integer"
'        End If
'        sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"
'
'      Case dtNUMERIC
'        If Not fColFound Then
'          sDeclareInsCols.Append " numeric(" & Trim$(Str$(iSize)) & ", " & Trim$(Str$(iDecimals)) & ")"
'          sDeclareDelCols.Append " numeric(" & Trim$(Str$(iSize)) & ", " & Trim$(Str$(iDecimals)) & ")"
'        End If
'        sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"
'
'        ' JDM - 12/09/03 - Fault 5605 - Use Separator in emails
'        If recColEdit.Fields("Use1000Separator").value = True Then
'          sConvertInsCols = sConvertInsCols & vbNewLine _
'              & vbTab & "SET @strTemp = REVERSE(SUBSTRING(@strTemp,0,CHARINDEX('.',@strTemp)))" & vbNewLine _
'              & vbTab & "SET @itemp = 3" & vbNewLine _
'              & vbTab & "WHILE @itemp < LEN(@strTemp)" & vbNewLine _
'              & vbTab & "BEGIN" & vbNewLine _
'              & vbTab & vbTab & "SET @strTemp = LEFT(@strTemp, @itemp) + ',' + SUBSTRING(@strTemp,@itemp+1,len(@strTemp))" & vbNewLine _
'              & vbTab & vbTab & "SET @itemp = @itemp + 4" & vbNewLine _
'              & vbTab & "END" & vbNewLine _
'              & vbTab & "SET @strTemp = REVERSE(@strTemp) + SUBSTRING(@strTemp,CHARINDEX('.',@strTemp),LEN(@strTemp))" & vbNewLine
'        End If
'
'      Case dtTIMESTAMP
'        If Not fColFound Then
'          sDeclareInsCols.Append " datetime"
'          sDeclareDelCols.Append " datetime"
'        End If
'        sConvertInsCols = "ISNULL(CONVERT(varchar(3000), LEFT(DATENAME(month, @insCol_" & Trim$(Str$(lngColumnID)) & "),3) + ' ' + CONVERT(varchar(3000),DATEPART(day, @insCol_" & Trim$(Str$(lngColumnID)) & ")) + ' ' + CONVERT(varchar(3000),DATEPART(year, @insCol_" & Trim$(Str$(lngColumnID)) & "))), '')"
'
'      Case dtBIT
'        If Not fColFound Then
'          sDeclareInsCols.Append " bit"
'          sDeclareDelCols.Append " bit"
'        End If
'        sConvertInsCols = "ISNULL(CONVERT(varchar(3000), CASE @insCol_" & Trim$(Str$(lngColumnID)) & " WHEN 1 THEN 'True' WHEN 0 THEN 'False' END), '')"
'
'      Case dtVARBINARY, dtLONGVARBINARY
'        If Not fColFound Then
'          sDeclareInsCols.Append " varchar(3000)"
'          sDeclareDelCols.Append " varchar(3000)"
'        End If
'        sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"
'
'      Case Else
'        If Not fColFound Then
'          sDeclareInsCols.Append " varchar(8000)"
'          sDeclareDelCols.Append " varchar(8000)"
'        End If
'        sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"
'    End Select
'
'    If iDataType = dtTIMESTAMP Then
'      mstrInsertEmailTemp = vbNewLine & _
'        "            SET @strTemp = ISNULL(CONVERT(varchar(3000), @insCol_" & CStr(lngColumnID) & "," & CStr(glngEmailDateFormat) & "),'')" & vbNewLine & _
'        mstrInsertEmailTemp & vbNewLine
'
'      mstrUpdateEmailTemp = vbNewLine & _
'        "            SET @strTemp = ISNULL(CONVERT(varchar(3000), @insCol_" & CStr(lngColumnID) & "," & CStr(glngEmailDateFormat) & "),'')" & vbNewLine & _
'        mstrUpdateEmailTemp & vbNewLine
'
'    Else
'      mstrInsertEmailTemp = vbNewLine & _
'        "            SET @strTemp = " & sConvertInsCols & vbNewLine & _
'        mstrInsertEmailTemp & vbNewLine
'
'      mstrUpdateEmailTemp = vbNewLine & _
'        "            SET @strTemp = " & sConvertInsCols & vbNewLine & _
'        mstrUpdateEmailTemp & vbNewLine
'    End If
'
'    'Update is the same as insert for the moment !
'    'mstrUpdateEmailTemp = mstrInsertEmailTemp
'
'  End If
'
'Exit Sub
'
'LocalErr:
'  If ASRDEVELOPMENT Then
'    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
'    Stop
'  End If
'
'End Sub
'
'
'Private Function CreateEmailProcsForLink(lngTableID As Long, sCurrentTable As String, lngColumnID As Long, lngRecordDescExprID As Long) As String
'
''  'Dim recEmailLinks As rdo.rdoResultset
''  Dim strSQL As String
''
''  Dim strTriggerCode As String
''  Dim strSendCode As String
''  Dim strSendTemp As String
''  Dim strRebuildTemp As String
''  Dim strImmediate As String
''  Dim strTemp As String
''
''  Dim strSPEmailSend As String
''  Dim strColumnName As String
''  'Dim lngColumnType As Long
''  Dim blnDateColumn As Boolean
''  Dim lngLinkID As Long
''  Dim strLinkTitle As String
''  Dim strLinkFilter As String
''  Dim strLinkEffectiveDate As String
''  'Dim strLinkSubject As String
''  'Dim strLinkText As String
''  'Dim blnIncRecDesc As Boolean
''  'Dim blnIncColDetail As Boolean
''  'Dim blnIncUserName As Boolean
''  Dim blnAttachment As Boolean
''  Dim lngSubjectID As Long
''  Dim lngBodyID As Long
''
''
''  On Error GoTo LocalErr
''
''  With recEmailLinksEdit
''
''    strColumnName = recColEdit!ColumnName
''    'lngColumnType = recColEdit!DataType
''    blnDateColumn = (recColEdit!DataType = dtTIMESTAMP)
''    lngLinkID = !LinkID
''    strLinkTitle = IIf(IsNull(!Title), vbNullString, !Title)
''    strLinkSubject = IIf(IsNull(!Subject), vbNullString, !Subject)
''    strLinkText = IIf(IsNull(!Body), vbNullString, !Body)
''    blnIncRecDesc = !IncRecDesc
''    blnIncColDetail = !IncColDetail
''    blnIncUserName = !IncUsername
''    blnAttachment = (Trim(!Attachment) <> vbNullString)
''    lngSubjectID = IIf(IsNull(!SubjectContentID), 0, !SubjectContentID)
''    lngBodyID = IIf(IsNull(!BodyContentID), 0, !BodyContentID)
''
''    strSPEmailSend = gsEMAILSEND & CStr(lngLinkID)
''
''
'''    strSQL = "IF EXISTS " & _
'''             "(SELECT Name FROM sysobjects " & _
'''             "WHERE id = object_id('" & strSPEmailSend & "') " & _
'''             "AND sysstat & 0xf = 4) " & _
'''             "DROP PROCEDURE " & strSPEmailSend
'''    gADOCon.Execute strSQL, , adExecuteNoRecords
''    DropProcedure strSPEmailSend
''
''
''    If Not IsNull(!EffectiveDate) Then
''      strLinkEffectiveDate = Replace(Format(!EffectiveDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
''    Else
''      strLinkEffectiveDate = vbNullString
''    End If
''
''    If !FilterID > 0 Then
''      strLinkFilter = GetSQLFilter(!FilterID, sCurrentTable)
''
''      'Need to restore current records after getting filter
''      .Index = "idxID"
''      .Seek "=", lngLinkID
''      recColEdit.Index = "idxColumnID"
''      recColEdit.Seek "=", lngColumnID
''    End If
''
''
''    'Get Recipients
''    strSendTemp = GetSQLRecipients(lngLinkID)
''    If Trim(strSendTemp) = vbNullString Then
''      Exit Function
''    End If
''
''
'''''    'Body of message...
'''''    If blnIncColDetail Then
'''''      strSendTemp = strSendTemp & vbNewLine & _
'''''          "      IF @strTemp IS NULL or LTrim(@strTemp) = ''" & vbNewLine & _
'''''          "        SELECT @TempText = '" & Replace(Trim(strColumnName), "_", " ") & " : <Empty>'" & vbNewLine & _
'''''          "      ELSE" & vbNewLine & _
'''''          "        SELECT @TempText = '" & Replace(Trim(strColumnName), "_", " ") & " : ' + @strTemp" & vbNewLine
'''''    End If
'''''
'''''    strSendTemp = strSendTemp & vbNewLine & "      SELECT @TempText = "
'''''
'''''    If blnIncRecDesc Then
'''''      strSendTemp = strSendTemp & "'" & sCurrentTable & " : ' + @RecordDesc + char(13) + "
'''''    End If
'''''
'''''    If blnIncColDetail Then
'''''      strSendTemp = strSendTemp & "@TempText + char(13) + char(13) + "
'''''    End If
'''''
'''''    strSendTemp = strSendTemp & "char(13) + '" & Replace(strLinkText, "'", "''") & "'"
'''''
'''''    If blnIncUserName Then
'''''      ' NPG20090210 Fault 13398
'''''      ' strSendTemp = strSendTemp & vbNewLine & vbTab & "IF @UserName = '" & gsWORKFLOWAPPLICATIONPREFIX & "' SET @TempText = @TempText + char(13) + char(13) + 'Changed By : " & gsWORKFLOWAPPLICATIONPREFIX & "'" & vbNewLine & _
'''''      '   vbTab & "ELSE IF USER <> 'dbo' SET @TempText = @TempText + char(13) + char(13) + 'Changed By : ' + USER" & vbNewLine
'''''      strSendTemp = strSendTemp & vbNewLine & vbTab & "IF @UserName = '" & gsWORKFLOWAPPLICATIONPREFIX & "' SET @TempText = @TempText + char(13) + char(13) + 'Changed By : " & gsWORKFLOWAPPLICATIONPREFIX & "'" & vbNewLine & _
'''''        vbTab & "ELSE SET @TempText = @TempText + char(13) + char(13) + 'Changed By : ' + @UserName" & vbNewLine
'''''    End If
'''''
'''''
'''''  'ELSE IF USER <> 'dbo' SET @TempText = @TempText + char(13) + char(13) + 'Changed By : ' + USER
'''''
'''''  'ELSE SET @TempText = @TempText + char(13) + char(13) + 'Changed By : ' + @UserName
''
''
''
''    Dim content As clsLinkContent
''
''
''    Set content = New clsLinkContent
''    content.ReadDetail (lngSubjectID)
''    strSendTemp = strSendTemp & content.GetSQL(lngTableID, "@Subject") & vbNewLine & vbNewLine
''    Set content = Nothing
''
''
''    Set content = New clsLinkContent
''    content.ReadDetail (lngBodyID)
''    strSendTemp = strSendTemp & content.GetSQL(lngTableID, "@TempText") & vbNewLine & vbNewLine
''    Set content = Nothing
''
''
''
''
''
''    'Add Attachment where required.
''    strSendTemp = strSendTemp & vbNewLine & _
''    "      SET @hResult = 1" & vbNewLine & _
''    "      SELECT @Attachment = ''" & vbNewLine
''
''    If blnAttachment Then
''      'MH20061208 Fault 11733
''      'strSendTemp = strSendTemp & vbNewLine & vbNewLine & _
''        "      EXEC master..xp_fileexist '" & gstrEmailAttachmentPath & !Attachment & "'" & vbNewLine & vbNewLine
''      strSendTemp = strSendTemp & vbNewLine & vbNewLine & _
''        "      EXEC master..xp_fileexist '" & gstrEmailAttachmentPath & !Attachment & "', @AttachmentExists OUTPUT" & vbNewLine & vbNewLine
''
''      'strSendTemp = strSendTemp & vbNewLine & _
''        "      IF @AttachmentExists = 1" & vbNewLine & _
''        "        SELECT @attachment = '" & gstrEmailAttachmentPath & !Attachment & "'" & vbNewLine & _
''        "      ELSE" & vbNewLine & _
''        "        SELECT @TempText = @TempText + char(13)+ char(13)+ char(13)+" & _
''                 "'Attachment not found <" & gstrEmailAttachmentPath & !Attachment & ">'" & vbNewLine
''
''      strSendTemp = strSendTemp & vbNewLine & _
''        "      IF @AttachmentExists = 1" & vbNewLine & _
''        "      BEGIN" & vbNewLine & _
''        "        SELECT @attachment = '" & gstrEmailAttachmentPath & !Attachment & "'" & vbNewLine
''
''    End If
''
''
''    'Fault 1296 Blank subject defaults to "SQL server message"
''    'If blank then put in a tab instead !
''    'strSendTemp = strSendTemp & vbNewLine & _
''      "      SET @Subject = '" & IIf(Trim(strLinkSubject) = vbNullString, vbTab, Trim(Replace(strLinkSubject, "'", "''"))) & "'" & vbNewLine & _
''      "      EXEC spASRsendmail @hResult OUTPUT, @RecipTo, @RecipCC, @RecipBcc, @Subject, @TempText, @attachment" & vbNewLine
''    strSendTemp = strSendTemp & vbNewLine & _
''      "      IF rtrim(@Subject) = ''" & vbNewLine & _
''      "        SET @Subject = char(9)" & vbNewLine & _
''      "      EXEC spASRsendmail @hResult OUTPUT, @RecipTo, @RecipCC, @RecipBcc, @Subject, @TempText, @attachment" & vbNewLine
''
''
''    If blnAttachment Then
''      strSendTemp = strSendTemp & vbNewLine & _
''        "      END" & vbNewLine
''    End If
''
''
''
''    If strLinkEffectiveDate <> vbNullString Then
''      strSendTemp = _
''      "    IF DateDiff(day, '" & strLinkEffectiveDate & "', @emailDate) >= 0" & vbNewLine & _
''      "    BEGIN" & vbNewLine & _
''      strSendTemp & vbNewLine & _
''      "    END" & vbNewLine
''    End If
''
''    If strLinkFilter <> vbNullString Then
''      strSendTemp = _
''      "  IF " & strLinkFilter & vbNewLine & _
''      "  BEGIN" & vbNewLine & _
''      strSendTemp & vbNewLine & _
''      "  END" & vbNewLine
''    End If
''
''    strSendCode = strSendCode & vbNewLine & _
''      "  /* " & strLinkTitle & " */" & vbNewLine & _
''      "  SELECT @hResult = 1" & vbNewLine & _
''      strSendTemp & vbNewLine & _
''      "  RETURN @hResult"
''
''
''    strImmediate = _
''      "                INSERT ASRSysEmailQueue(LinkID, ColumnID, RecordID, ColumnValue, DateDue, UserName, [Immediate],RecalculateRecordDesc, RecordDesc)" & vbNewLine & _
''      "                VALUES (" & CStr(lngLinkID) & "," & CStr(lngColumnID) & ",@recordID,@strTemp,getdate()," & _
''      "CASE WHEN UPPER(LEFT(APP_NAME(), " & Len(gsWORKFLOWAPPLICATIONPREFIX) & ")) = '" & UCase(gsWORKFLOWAPPLICATIONPREFIX) & "' THEN '" & gsWORKFLOWAPPLICATIONPREFIX & "' ELSE @username END," & _
''      "1,@RecalculateRecordDesc, @recordDesc)"
''
''
''    strRebuildTemp = _
''      "                INSERT ASRSysEmailQueue(LinkID, ColumnID, RecordID, ColumnValue, DateDue, UserName, [Immediate],RecalculateRecordDesc)" & vbNewLine & _
''      "                VALUES (" & CStr(lngLinkID) & "," & CStr(lngColumnID) & ",@recordID,@strTemp,@emailDate," & _
''      "CASE WHEN UPPER(LEFT(APP_NAME(), " & Len(gsWORKFLOWAPPLICATIONPREFIX) & ")) = '" & UCase(gsWORKFLOWAPPLICATIONPREFIX) & "' THEN '" & gsWORKFLOWAPPLICATIONPREFIX & "' ELSE @username END," & _
''      "0,1)"
''
''    If !Immediate Then
''      strTriggerCode = strTriggerCode & vbNewLine & strImmediate
''
''      If strLinkFilter <> vbNullString Then
''        strTriggerCode = _
''          "            IF " & strLinkFilter & vbNewLine & _
''          "            BEGIN" & vbNewLine & _
''          strTriggerCode & vbNewLine & _
''          "            END" & vbNewLine
''      End If
''
''      If strLinkEffectiveDate <> vbNullString Then
''        strTriggerCode = _
''        "    IF DateDiff(day, '" & strLinkEffectiveDate & "', getdate()) >= 0" & vbNewLine & _
''        "    BEGIN" & vbNewLine & _
''        strTriggerCode & vbNewLine & _
''        "    END" & vbNewLine
''      End If
''
''    Else
''      strTriggerCode = strTriggerCode & _
''        vbNewLine & _
''        "                SET @hResult = '1'" & vbNewLine & _
''        vbNewLine
''
''      'If email should have already been sent or
''      'if email previously sent (which was not null)
''      'then send an email immediately
''      strTriggerCode = strTriggerCode & _
''        "                IF (DateDiff(day, @emailDate, getdate()) >= 0) OR" & vbNewLine & _
''        "                    (@LastSent IS NOT NULL)" & vbNewLine & _
''        "                BEGIN" & vbNewLine & _
''        strImmediate & vbNewLine & _
''        "                END" & vbNewLine & _
''        vbNewLine
''
''
''      'Check if we need to insert in queue for future email
''      strTriggerCode = strTriggerCode & _
''        "                IF (DateDiff(day, @emailDate, getdate()) < 0)" & vbNewLine & _
''        "                BEGIN" & vbNewLine & _
''        strRebuildTemp & vbNewLine & _
''        "                END" & vbNewLine & _
''        vbNewLine
''
''
''      strTemp = _
''        "                SELECT @LastSent = (SELECT TOP 1 [ColumnValue] FROM ASRSysEmailQueue " & vbNewLine & _
''        "                WHERE recordid = @recordid AND LinkID = " & CStr(lngLinkID) & " ORDER BY DateSent DESC)" & vbNewLine & _
''        vbNewLine & _
''        "                IF ((DateDiff(day, @purgeDate, @emailDate) >= 0 OR @PurgeDate IS NULL)" & vbNewLine
''
''
''      strTriggerCode = strTemp & _
''        "                    OR (@LastSent IS NOT NULL)) " & vbNewLine & _
''        IIf(strLinkEffectiveDate <> vbNullString, "                    AND (DateDiff(day, '" & strLinkEffectiveDate & "', @emailDate) >= 0)", "") & vbNewLine & _
''        "                BEGIN" & vbNewLine & _
''        strTriggerCode & vbNewLine & _
''        "                END" & vbNewLine
''
''      strRebuildTemp = strTemp & _
''        "                 AND IsNull(@LastSent,'') <> IsNull(@strTemp,''))" & vbNewLine & _
''        IIf(strLinkEffectiveDate <> vbNullString, "                 AND (DateDiff(day, '" & strLinkEffectiveDate & "', @emailDate) >= 0)", "") & vbNewLine & _
''        "              BEGIN" & vbNewLine & _
''        strRebuildTemp & vbNewLine & _
''        "              END" & vbNewLine
''
''
''      If Abs(!Offset) Then
''        strTemp = _
''          "              SELECT @emailDate = dateadd(" & _
''          Choose(!Period + 1, "dd", "ww", "mm", "yy") & _
''          "," & !Offset & ",@emailDate)" & vbNewLine
''
''        strTriggerCode = strTemp & strTriggerCode
''        strRebuildTemp = strTemp & strRebuildTemp
''
''      End If
''
''      'strTemp = "              SELECT @emailDate = IsNull(convert(datetime,@strTemp),getdate())" & vbNewline
''      'strTriggerCode = strTemp & strTriggerCode
''      'strRebuildTemp = _
''                "        SELECT @strTemp = CONVERT(varchar(255), " & strColumnName & ") FROM " & sCurrentTable & " WHERE id = @recordID" & vbNewline & _
''                strTemp & strRebuildTemp
''      strTriggerCode = _
''        "              SELECT @emailDate = IsNull(convert(datetime,@insCol_" & CStr(lngColumnID) & "),getdate())" & vbNewLine & _
''        strTriggerCode
''      strRebuildTemp = _
''                "        SELECT @strTemp = CONVERT(varchar(3000), " & strColumnName & _
''                IIf(blnDateColumn, "," & CStr(glngEmailDateFormat), "") & _
''                ") FROM " & sCurrentTable & " WHERE id = @recordID" & vbNewLine & _
''                "        SELECT @emailDate = IsNull(convert(datetime," & strColumnName & "),getdate()) FROM " & sCurrentTable & " WHERE id = @recordID" & vbNewLine & _
''                strRebuildTemp
''
''
''      If strLinkFilter <> vbNullString Then
''        strTriggerCode = _
''          "            IF " & strLinkFilter & vbNewLine & _
''          "            BEGIN" & vbNewLine & _
''          strTriggerCode & vbNewLine & _
''          "            END" & vbNewLine
''
''        strRebuildTemp = _
''          "            IF " & strLinkFilter & vbNewLine & _
''          "            BEGIN" & vbNewLine & _
''          strRebuildTemp & vbNewLine & _
''          "            END" & vbNewLine
''      End If
''
''      mstrRebuildCode = mstrRebuildCode & vbNewLine & _
''        "            /* " & strLinkTitle & " */" & vbNewLine & _
''        strRebuildTemp & vbNewLine
''
''    End If
''
''    strTriggerCode = vbNewLine & _
''      "            /* " & strLinkTitle & " */" & vbNewLine & _
''      strTriggerCode & vbNewLine
''
''  End With
''
''
''  If strSendCode <> vbNullString Then
''
''    'strSQL = _
''      "/* ------------------------------ */" & vbNewline & _
''      "/* Email stored procedure. */" & vbNewline & _
''      "/* ------------------------------ */" & vbNewline & _
''      "CREATE PROCEDURE " & strSPEmailSend & _
''      "(@recordID int, @recordDesc varChar(8000), @strTemp varChar(8000), @emailDate datetime, @UserName varchar(50))" & vbNewline & _
''      "AS" & vbNewline & _
''      "BEGIN" & vbNewline & _
''      "  DECLARE @hResult int," & vbNewline & _
''      "          @RecipTo varchar(8000)," & vbNewline & _
''      "          @RecipCc varchar(8000)," & vbNewline & _
''      "          @RecipBcc varchar(8000)," & vbNewline & _
''      "          @TempRecip varchar(8000)," & vbNewline & _
''      "          @TempText varchar(8000)," & vbNewline & _
''      "          @Attachment varchar(8000)," & vbNewline & _
''      "          @AttachmentExists int" & vbNewline & _
''      vbNewline & strSendCode & vbNewline & _
''      "END"
''    strSQL = _
''      "/* ------------------------------ */" & vbNewLine & _
''      "/* Email stored procedure. */" & vbNewLine & _
''      "/* ------------------------------ */" & vbNewLine & _
''      "CREATE PROCEDURE " & strSPEmailSend & vbNewLine & _
''      "(   @recordID int, " & vbNewLine & _
''      "    @recordDesc varChar(8000), " & vbNewLine & _
''      "    @strTemp varChar(8000), " & vbNewLine & _
''      "    @emailDate datetime, " & vbNewLine & _
''      "    @UserName varchar(50), " & vbNewLine & _
''      "    @RecipTo varchar(4000) OUTPUT, " & vbNewLine & _
''      "    @RecipCc varchar(4000) OUTPUT, " & vbNewLine & _
''      "    @RecipBcc varchar(4000) OUTPUT, " & vbNewLine & _
''      "    @Subject varchar(4000) OUTPUT, " & vbNewLine & _
''      "    @TempText varchar(8000) OUTPUT, " & vbNewLine & _
''      "    @Attachment varchar(4000) OUTPUT" & vbNewLine & _
''      ")" & vbNewLine & _
''      "AS" & vbNewLine & _
''      "BEGIN" & vbNewLine & _
''      "  DECLARE @hResult int," & vbNewLine & _
''      "          @TempRecip varchar(8000)," & vbNewLine & _
''      "          @AttachmentExists int" & vbNewLine & _
''      vbNewLine & strSendCode & vbNewLine & _
''      "END"
''
''    gADOCon.Execute strSQL, , adExecuteNoRecords
''
''  End If
''
''  CreateEmailProcsForLink = strTriggerCode
''
''Exit Function
''
''LocalErr:
''  If ASRDEVELOPMENT Then
''    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
''    Stop
''  End If
'
'End Function
'
'
'Public Function GetEmailAddressName(lngRecipientID As Long) As String
'
'  On Error GoTo ErrorTrap
'
'  GetEmailAddressName = vbNullString
'
'  With recEmailAddrEdit
'    .Index = "idxID"
'    .Seek "=", lngRecipientID
'    If Not .NoMatch Then
'      If Not !Deleted Then
'        GetEmailAddressName = !Name
'      End If
'    End If
'  End With
'
'  Exit Function
'
'ErrorTrap:
'
'End Function
'
