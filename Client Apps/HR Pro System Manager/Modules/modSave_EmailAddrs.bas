Attribute VB_Name = "modSave_EmailAddrs"
Option Explicit

Public Function SaveEmailAddrs(mfrmUse As frmUsage) As Boolean
  ' Save the new or modified Email Address definitions to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objEmailAddr As clsEmailAddr
  
  fOK = True
  
  With recEmailAddrEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      If !Deleted Then
        Set objEmailAddr = New clsEmailAddr
        objEmailAddr.EmailID = !EmailID
        Set mfrmUse = New frmUsage
        mfrmUse.ResetList
        If objEmailAddr.EmailIsUsed(mfrmUse) Then
          gobjProgress.Visible = False
          Screen.MousePointer = vbNormal
          mfrmUse.ShowMessage !Name & " Email", "The email cannot be deleted as the email is used by the following:", UsageCheckObject.Email
          fOK = False
        End If
        UnLoad mfrmUse
        Set mfrmUse = Nothing
       
        gobjProgress.Visible = True
        
        If fOK Then
          fOK = EmailAddrDelete
        End If
        
      ElseIf !New Then
        fOK = EmailAddrNew
      ElseIf !Changed Then
        fOK = EmailAddrSave
      End If
      
      .MoveNext
    Loop
  End With
  
TidyUpAndExit:
  SaveEmailAddrs = fOK
  Exit Function
  
ErrorTrap:
  'MsgBox "Error creating email addresses" & _
         IIf(Trim(Err.Description) <> vbnullstring, "(" & Err.Description & ")", vbnullstring), vbCritical
  OutputError "Error creating email addresses"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function EmailAddrDelete() As Boolean
  ' Delete the current Order definition from the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  ' Delete the existing order definition from the server database.
  sSQL = "DELETE FROM ASRSysEmailAddress" & _
    " WHERE EmailID=" & recEmailAddrEdit!EmailID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  EmailAddrDelete = fOK
  Exit Function

ErrorTrap:
  fOK = False
  OutputError "Error Deleting email address"
  Resume TidyUpAndExit
  
End Function

Private Function EmailAddrNew() As Boolean
  ' Write the current Order definition to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iColumn As Integer
  Dim sName As String
  Dim rsEmailAddrs As ADODB.Recordset
  
  Set rsEmailAddrs = New ADODB.Recordset
  fOK = True
  
  ' Open the order definition table on the server.
  rsEmailAddrs.Open "ASRSysEmailAddress", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

  ' Add the new order definition.
  With rsEmailAddrs
    .AddNew
    .Fields("EmailID").value = recEmailAddrEdit!EmailID

    For iColumn = 0 To .Fields.Count - 1
      sName = .Fields(iColumn).Name
      
      If Not UCase$(Trim$(sName)) = "TIMESTAMP" Then
        If Not IsNull(recEmailAddrEdit.Fields(sName).value) Then
          .Fields(iColumn).value = recEmailAddrEdit.Fields(sName).value
        End If
      End If
    Next iColumn
    .Update
  End With
  rsEmailAddrs.Close
  
TidyUpAndExit:
  Set rsEmailAddrs = Nothing
  EmailAddrNew = fOK
  Exit Function

ErrorTrap:
  OutputError "Error saving new email address"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function EmailAddrSave() As Boolean
  ' Save the current order to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Delete the existing record in the server database.
  fOK = EmailAddrDelete
  
  If fOK Then
    ' Create the new record in the server database.
    fOK = EmailAddrNew
  End If

TidyUpAndExit:
  EmailAddrSave = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Public Function CreateEmailAddrStoredProcedure() As Boolean

  Const strSPName As String = "spASRSysEmailAddr"
  Dim strSQL As String
  Dim strTableName As String
  Dim strColumnName As String
  Dim strSPEmailCalc As String
  Dim fOK As Boolean
  
  On Error GoTo ErrorTrap

  CreateEmailSendStoredProcedure "spASRSendMail", glngEmailMethod, gstrEmailProfile, gstrEmailServer, gstrEmailAccount
  
  fOK = True

'  ' Drop any existing stored procedure.
'  strSQL = "IF EXISTS" & _
'           " (SELECT Name" & _
'           "   FROM sysobjects" & _
'           "   WHERE id = object_id('" & strSPName & "')" & _
'           "     AND sysstat & 0xf = 4)" & _
'           " DROP PROCEDURE dbo." & strSPName
'  gADOCon.Execute strSQL, , adExecuteNoRecords
  DropProcedure strSPName


  strSQL = vbNullString
  With recEmailAddrEdit
    .Index = "idxID"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF

      If Not !Deleted Then
        If !Type <> 0 Then
          strSQL = strSQL & _
            IIf(strSQL <> vbNullString, "ELSE ", vbNullString) & _
            "IF @EmailID = " & CStr(!EmailID) & vbNewLine & _
            "BEGIN" & vbNewLine
      
          Select Case !Type
          Case 1    'Column
            strTableName = GetTableName(!TableID)
            strColumnName = GetColumnName(!ColumnID)
            
            If strTableName <> vbNullString And strColumnName <> vbNullString Then
              strSQL = strSQL & _
                "    /* " & Trim(!Name) & " (Column) */" & vbNewLine & _
                "    SET @hResult = (SELECT ltrim(rtrim(" & strColumnName & ")) FROM " & strTableName & " WHERE ID = @recordID)"
            End If
      
          Case 2    'Calculated
      
            strSPEmailCalc = "sp_ASRExpr_" & CStr(!ExprID)
            
            strSQL = strSQL & _
              "    /* " & Trim(!Name) & " (Calculated) */" & vbNewLine & _
              "    IF EXISTS (SELECT Name FROM sysobjects WHERE type = 'P'" & _
              "        AND name = '" & strSPEmailCalc & "')" & vbNewLine & _
              "    BEGIN" & vbNewLine & _
              "        EXEC @hResult = " & strSPEmailCalc & " @EmailAddr OUTPUT, @recordID" & vbNewLine & _
              "        IF @hResult <> 0 SET @EmailAddr = ''" & vbNewLine & _
              "        SET @hResult = ltrim(rtrim(CONVERT(varchar(255), @EmailAddr)))" & vbNewLine & _
              "    END"

          
          End Select
  
          strSQL = strSQL & vbNewLine & _
            "END" & vbNewLine & vbNewLine
  
        End If
      End If

      .MoveNext

    Loop
  End With


  strSQL = _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro email address stored procedure.                       */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE dbo." & strSPName & vbNewLine & _
    "(" & vbNewLine & _
    "    @hResult varchar(8000) OUTPUT," & vbNewLine & _
    "    @EmailID integer," & vbNewLine & _
    "    @recordID integer" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "DECLARE @EmailAddr char(255)" & vbNewLine & vbNewLine & _
    strSQL & vbNewLine & _
    IIf(Len(strSQL) > 0, vbNewLine & "ELSE" & vbNewLine, vbNullString) & _
    "    SET @hResult = (SELECT ltrim(rtrim(Fixed)) From ASRSysEmailAddress WHERE EmailID = @EmailID)" & vbNewLine & _
    "END"
  gADOCon.Execute strSQL, , adExecuteNoRecords


TidyUpAndExit:
  CreateEmailAddrStoredProcedure = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating email addresses"
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function CreateEmailSendStoredProcedure(strProcName As String, lngMethod As Long, strProfile As String, strServer As String, strAccount As String) As Boolean

  Dim sSQL As String
  Dim fOK As Boolean

  On Local Error GoTo LocalErr

  sSQL = vbNullString
  If GetSystemSetting("email", "qa info", 0) = 1 Then

    Select Case lngMethod
    Case 1: sSQL = "xp_sendmail"
    Case 2: sSQL = "sp_send_dbmail"
    Case 3: sSQL = "xp_SMTPsendmail80"
    End Select

    'sSQL = "SELECT @Message = @Message+char(13)+char(13)+'QA Info: '+@@SERVERNAME+'\'+DB_NAME()+' (" & sSQL & ")'" & vbCrLf
    sSQL = "    SELECT @Message = @Message+char(13)+char(13)+'QA Info: '+CONVERT(varchar, SERVERPROPERTY('servername'))+'\'+DB_NAME()+' (" & sSQL & ")'" & vbCrLf
  End If
  
  
  Select Case lngMethod
  Case 0
    sSQL = "    SET @hResult = 0"   'MH20061219 Mark as sent as per QA

  Case 1
    'MH20071026 Fault 12555
    GrantExecuteToPublic "master", "xp_startmail"
    GrantExecuteToPublic "master", "xp_sendmail"
    
    sSQL = sSQL & _
        "    DECLARE @To8000 varchar(8000)" & vbCrLf & _
        "    DECLARE @CC8000 varchar(8000)" & vbCrLf & _
        "    DECLARE @BCC8000 varchar(8000)" & vbCrLf & _
        "    DECLARE @Subject8000 varchar(8000)" & vbCrLf & _
        "    DECLARE @Message8000 varchar(8000)" & vbCrLf & vbCrLf

    sSQL = sSQL & _
        "    SET @To8000 = Left(@To,8000)" & vbCrLf & _
        "    SET @CC8000 = Left(@CC,8000)" & vbCrLf & _
        "    SET @BCC8000 = Left(@BCC,8000)" & vbCrLf & _
        "    SET @Subject8000 = Left(@Subject,8000)" & vbCrLf & _
        "    SET @Message8000 = Left(@Message,8000)" & vbCrLf & vbCrLf
    
    sSQL = sSQL & _
        "    EXEC @hResult = master..xp_sendmail " & vbCrLf & _
        "    @recipients=@To8000, " & vbCrLf & _
        "    @copy_recipients=@CC8000, " & vbCrLf & _
        "    @blind_copy_recipients=@BCC8000, " & vbCrLf & _
        "    @subject=@Subject8000, " & vbCrLf & _
        "    @message=@Message8000, " & vbCrLf & _
        "    @attachments=@Attachment"

  Case 2
    'MH20071026 Fault 12555
    GrantExecuteToPublic "msdb", "sp_send_dbmail"
           
    ' AE20080215 Fault #12834
'    If Trim(sNewProfile) <> "<Use Default Profile>" And _
'        Trim(sNewProfile) <> "" Then
'      sNewProfile = "@profile_name = '" & sNewProfile & "' ," & vbCrLf
'    Else
'      sNewProfile = vbNullString
'    End If
    Dim sNewProfile As String
    sNewProfile = strProfile

    If Trim(sNewProfile) = "<Use Default Profile>" Then
      sNewProfile = vbNullString
    ElseIf Trim(sNewProfile) <> vbNullString Then
      sNewProfile = "    @profile_name = '" & sNewProfile & "' ," & vbCrLf
    End If

    sSQL = sSQL & _
        "    EXEC @hResult = msdb.dbo.sp_send_dbmail " & vbCrLf & _
        sNewProfile & _
        "    @recipients=@To, " & vbCrLf & _
        "    @copy_recipients=@CC, " & vbCrLf & _
        "    @blind_copy_recipients=@BCC, " & vbCrLf & _
        "    @subject=@Subject, " & vbCrLf & _
        "    @body=@Message, " & vbCrLf & _
        "    @file_attachments=@Attachment"

  Case 3
    'MH20071026 Fault 12555
    GrantExecuteToPublic "master", "xp_SMTPsendmail80"
    
    sSQL = sSQL & _
        "    EXEC @hResult = master..xp_SMTPsendmail80 " & vbCrLf & _
        "    @address='" & Replace(strServer, "'", "''") & "', " & vbCrLf & _
        "    @from='" & Replace(strAccount, "'", "''") & "', " & vbCrLf & _
        "    @recipient=@To, " & vbCrLf & _
        "    @copy_recipients=@CC, " & vbCrLf & _
        "    @blind_copy_recipients=@BCC, " & vbCrLf & _
        "    @subject=@Subject, " & vbCrLf & _
        "    @body=@Message, " & vbCrLf & _
        "    @attachments=@Attachment"

  End Select
  
  
  'MH20090925 HRPRO-280
  sSQL = _
       "    IF rtrim(replace(isnull(@Attachment,''),';','')) <> ''" & vbCrLf & _
       "    BEGIN" & vbCrLf & _
       "        EXEC master..xp_fileexist @Attachment, @AttachmentExists OUTPUT" & vbCrLf & _
       "        IF @AttachmentExists = 0" & vbCrLf & _
       "            RETURN" & vbCrLf & _
       "    END" & vbCrLf & vbCrLf & _
       sSQL
  
  
  'MH20090827 HRPRO-304
  sSQL = _
       "    IF rtrim(replace(isnull(@To,''),';','')) = ''" & vbCrLf & _
       "        RETURN" & vbCrLf & vbCrLf & _
       sSQL
  
  
  DropProcedure strProcName
  sSQL = "CREATE PROCEDURE dbo.[" & strProcName & "](" & vbCrLf & _
         "  @hResult int OUTPUT," & vbCrLf & _
         "  @To varchar(max)," & vbCrLf & _
         "  @CC varchar(max)," & vbCrLf & _
         "  @BCC varchar(max)," & vbCrLf & _
         "  @Subject varchar(max)," & vbCrLf & _
         "  @Message varchar(max)," & vbCrLf & _
         "  @Attachment varchar(8000))" & vbCrLf & _
         "AS " & vbCrLf & _
         "BEGIN " & vbCrLf & vbCrLf & _
         "    DECLARE @AttachmentExists int" & vbCrLf & vbCrLf & _
         "    SET @hResult = 1" & vbCrLf & vbCrLf & _
         sSQL & vbCrLf & vbCrLf & _
         "END"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  fOK = True

TidyUpAndExit:
  CreateEmailSendStoredProcedure = fOK
  Exit Function

LocalErr:
  OutputError "Error creating email send procedure"
  fOK = False
  Resume TidyUpAndExit

End Function


'MH20071026 Fault 12555
Public Sub GrantExecuteToPublic(strDatabase As String, strSPName As String)

  On Local Error Resume Next

  'JPD 20080714 Fault 13265
  If gbIsUserSystemAdmin Then
    gADOCon.Execute "USE [" & strDatabase & "]", , adExecuteNoRecords
    gADOCon.Execute "GRANT EXECUTE ON " & strSPName & " TO PUBLIC", , adExecuteNoRecords
    gADOCon.Execute "USE [" & gsDatabaseName & "]", , adExecuteNoRecords
  End If
  
End Sub
