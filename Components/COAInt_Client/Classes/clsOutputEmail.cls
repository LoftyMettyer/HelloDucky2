VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''Option Explicit
'''
'''
'''Public Function SendEmailFromClient(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String, blnPause As Boolean) As String
'''
'''  'HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles
'''
'''  If blnPause Then
'''    SendEmailFromClientUsingShell strTo, strCC, strBCC, strSubject, strMsgText, strAttachment
'''
'''  ElseIf GetUserOrSystemSetting("Email CDO", "Enabled", "0") = 1 Then
'''    SendEmailFromClientUsingCDO strTo, strCC, strBCC, strSubject, strMsgText, strAttachment
'''
'''  Else
'''    SendEmailFromClientUsingMAPI strTo, strCC, strBCC, strSubject, strMsgText, strAttachment
'''
'''  End If
'''
'''End Function
'''
'''
'''Private Function SendEmailFromClientUsingShell(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String) As String
'''
'''  Dim strCommand As String
'''
'''  strCommand = "mailto:" & Replace(strTo, ";", ",") & " " & _
'''               "?cc=" & Replace(strCC, ";", ",") & " " & _
'''               "&bcc=" & Replace(strBCC, ";", ",") & " " & _
'''               "&subject=" & IIf(strSubject <> "", strSubject, " ") & _
'''               "&body=" & IIf(strMsgText <> "", Replace(strMsgText, vbCrLf, "%0A"), " ") & " "
'''
'''  ShellExecute 0&, vbNullString, strCommand, vbNullString, vbNullString, vbMaximizedFocus
'''
'''End Function
'''
'''
'''Private Function SendEmailFromClientUsingCDO(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String) As String
'''
'''  Dim iMsg As New CDO.Message
'''  Dim iConf As New CDO.Configuration
'''  Dim Flds As ADODB.Fields
'''  Dim strFrom As String
'''  Dim strArray() As String
'''  Dim lngIndex As Long
'''
'''  strFrom = GetUserOrSystemSetting("Email CDO", "From", "")
'''
'''  Set Flds = iConf.Fields
'''
'''  With Flds
'''    .Item(cdoSMTPServer) = GetUserOrSystemSetting("Email CDO", "SMTPServer", "")
'''    .Item(cdoSMTPUseSSL) = IIf(GetUserOrSystemSetting("Email CDO", "SMTPUseSSL", 0) = 1, True, False)
'''    .Item(cdoSMTPAuthenticate) = GetUserOrSystemSetting("Email CDO", "SMTPAuthenticate", 2)            'cdoAnonymous = 0, cdoBasic = 1, cdoNTLM = 2
'''    .Item(cdoSMTPConnectionTimeout) = GetUserOrSystemSetting("Email CDO", "SMTPConnectionTimeout", 10)
'''    .Item(cdoSendUserName) = GetUserOrSystemSetting("Email CDO", "SendUserName", "")
'''    .Item(cdoSendPassword) = GetUserOrSystemSetting("Email CDO", "SendPassword", "")
'''    .Item(cdoSendUsingMethod) = GetUserOrSystemSetting("Email CDO", "SendUsingMethod", 2)              'cdoSendUsingPickup = 1, cdoSendUsingPort = 2
'''    .Update
'''  End With
'''
'''  With iMsg
'''    Set .Configuration = iConf
'''    .To = strTo
'''    .From = strFrom
'''    .Subject = strSubject
'''    .TextBody = strMsgText
'''
'''    If (Trim(strAttachment) <> vbNullString) Then
'''      strArray = Split(strAttachment, ";")
'''      For lngIndex = LBound(strArray) To UBound(strArray)
'''        If Trim(strArray(lngIndex)) <> vbNullString Then
'''          .AddAttachment Trim(strArray(lngIndex))
'''        End If
'''      Next
'''    End If
'''
'''    .Send
'''  End With
'''
'''  Set iMsg = Nothing
'''  Set iConf = Nothing
'''  Set Flds = Nothing
'''
'''End Function
'''
'''
'''Private Function SendEmailFromClientUsingMAPI(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String) As String
'''
'''  Dim objMapiSession As MAPISession
'''  Dim objMapiMessages As MAPIMessages
'''
'''  Dim strError As String
'''  Dim strRecipients As String
'''  Dim lngRecipType As Long
'''
'''  Dim strArray() As String
'''  Dim lngIndex As Long
'''
'''  On Error GoTo LocalErr
'''
'''  Set objMapiSession = frmEmailSel.MAPISession1
'''  Set objMapiMessages = frmEmailSel.MAPIMessages1
'''
'''
'''  strError = vbNullString
'''
'''  If strTo & strCC & strBCC = vbNullString Then
'''    strError = "Please select recipient(s) to email"
'''    Exit Function
'''
'''  ElseIf strTo = vbNullString Then
'''    strError = "Please select a recipient from the TO column"
'''    Exit Function
'''
'''  End If
'''
'''
'''  If objMapiSession.SessionID = 0 Then
'''    objMapiSession.DownLoadMail = False
'''    objMapiSession.SignOn
'''    objMapiMessages.SessionID = objMapiSession.SessionID
'''  End If
'''
'''
'''  With objMapiMessages
'''    .Compose
'''
'''    For lngRecipType = 1 To 3
'''
'''      Select Case lngRecipType
'''      Case 1: strRecipients = strTo
'''      Case 2: strRecipients = strCC
'''      Case 3: strRecipients = strBCC
'''      End Select
'''
'''      If (Trim(strRecipients) <> vbNullString) Then
'''        strArray = Split(strRecipients, ";")
'''        For lngIndex = LBound(strArray) To UBound(strArray)
'''          If Trim(strArray(lngIndex)) <> vbNullString Then
'''            .RecipIndex = .RecipCount
'''            .RecipAddress = Trim(strArray(lngIndex))
'''            .RecipType = lngRecipType
'''            '.ResolveName
'''            'If intSendType = mapToList Then
'''            '  blnToRecipient = True
'''            'End If
'''          End If
'''        Next
'''      End If
'''
'''    Next
'''
'''    .MsgSubject = strSubject
'''    .MsgNoteText = strMsgText
'''
'''    If (Trim(strAttachment) <> vbNullString) Then
'''      strArray = Split(strAttachment, ";")
'''      For lngIndex = LBound(strArray) To UBound(strArray)
'''        If Trim(strArray(lngIndex)) <> vbNullString Then
'''          .AttachmentPosition = lngIndex
'''          .AttachmentType = 0
'''          .AttachmentPathName = strAttachment
'''          '.AttachmentName = mstrEmailAttachAs
'''        End If
'''      Next
'''    End If
'''
'''    .Send False
'''
'''  End With
'''
'''TidyAndExit:
'''  If objMapiSession.SessionID <> 0 Then
'''    objMapiSession.SignOff
'''  End If
'''  objMapiMessages.SessionID = 0
'''
'''  SendEmailFromClientUsingMAPI = strError
'''
'''Exit Function
'''
'''LocalErr:
'''  If Err.Number = 32001 Or Err = 32003 Then
'''    Resume Next
'''  Else
'''    strError = "Error sending email" & vbCrLf & _
'''           IIf(Err.Description <> vbNullString, "(" & Trim(Err.Description) & ")", "")
'''    GoTo TidyAndExit
'''  End If
'''
'''End Function
'''
'''
'''Public Function SendEmailFromServer(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String) As String
'''
'''  Dim adoCmd As ADODB.Command
'''  Dim strError As String
'''
'''  On Error GoTo LocalErr
'''
'''  strError = vbNullString
'''
'''  Set adoCmd = New ADODB.Command
'''  adoCmd.ActiveConnection = gADOCon
'''
'''  gADOCon.Errors.Clear
'''  adoCmd.CommandText = "exec spASRSendMail 0, " & _
'''                       "'" & Replace(strTo, "'", "''") & "', " & _
'''                       "'" & Replace(strCC, "'", "''") & "', " & _
'''                       "'" & Replace(strBCC, "'", "''") & "', " & _
'''                       "'" & Replace(strSubject, "'", "''") & "', " & _
'''                       "'" & Left(Replace(strMsgText, "'", "''"), 7000) & "', " & _
'''                       "'" & Replace(strAttachment, "'", "''") & "'"
'''  adoCmd.Execute
'''
'''TidyAndExit:
'''  SendEmailFromServer = strError
'''
'''Exit Function
'''
'''LocalErr:
'''  strError = Err.Description
'''  GoTo TidyAndExit
'''
'''End Function
'''
