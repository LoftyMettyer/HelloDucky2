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
Option Explicit



Public Function SendEmailFromServer(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String) As String

  Dim adoCmd As ADODB.Command
  Dim strError As String
  
  On Error GoTo LocalErr

  strError = vbNullString
  
  Set adoCmd = New ADODB.Command
  adoCmd.ActiveConnection = gADOCon
  
  gADOCon.Errors.Clear
  adoCmd.CommandText = "exec spASRSendMail 0, " & _
                       "'" & Replace(strTo, "'", "''") & "', " & _
                       "'" & Replace(strCC, "'", "''") & "', " & _
                       "'" & Replace(strBCC, "'", "''") & "', " & _
                       "'" & Replace(strSubject, "'", "''") & "', " & _
                       "'" & Left(Replace(strMsgText, "'", "''"), 7000) & "', " & _
                       "'" & Replace(strAttachment, "'", "''") & "'"
  adoCmd.Execute
  
TidyAndExit:
  SendEmailFromServer = strError

Exit Function

LocalErr:
  strError = Err.Description
  GoTo TidyAndExit

End Function


Public Function SendEmailFromClient(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String, blnPause As Boolean) As String

  Dim frm As frmEmailSel
  
  Dim strError As String
  Dim strRecipients As String
  Dim lngRecipType As Long
  
  Dim strArray() As String
  Dim lngIndex As Long
  
  On Error GoTo LocalErr

  strError = vbNullString
  
  If strTo & strCC & strBCC = vbNullString Then
    strError = "Please select recipient(s) to email"
    Exit Function
  
  ElseIf strTo = vbNullString Then
    strError = "Please select a recipient from the TO column"
    Exit Function
  
  End If
  
  
  frmEmailSel.MAPISignon
  If frmEmailSel.MAPISession1.SessionID <> 0 Then
    With frmEmailSel.MAPIMessages1
      .Compose
      
      For lngRecipType = 1 To 3
      
        Select Case lngRecipType
        Case 1: strRecipients = strTo
        Case 2: strRecipients = strCC
        Case 3: strRecipients = strBCC
        End Select
        
        If (Trim(strRecipients) <> vbNullString) Then
          strRecipArray = Split(strRecipients, ";")
          For lngIndex = LBound(strArray) To UBound(strArray)
            If Trim(strArray(lngIndex)) <> vbNullString Then
              .RecipIndex = .RecipCount
              .RecipAddress = Trim(strArray(lngIndex))
              .RecipType = lngRecipType
              .ResolveName
              'If intSendType = mapToList Then
              '  blnToRecipient = True
              'End If
            End If
          Next
        End If

      Next

      .MsgSubject = strSubject
      .MsgNoteText = strMsgText

'      If (Trim(strAttachment) <> vbNullString) Then
'        strArray = Split(strAttachment, ";")
'        For lngIndex = LBound(strArray) To UBound(strArray)
'          If Trim(strArray(lngIndex)) <> vbNullString Then
'            .AttachmentPosition = lngIndex
'            .AttachmentType = 0
'            .AttachmentPathName = strAttachment
''            .AttachmentName = mstrEmailAttachAs
'          End If
'        Next
'      End If

      .Send blnPause

    End With

  End If

TidyAndExit:
  frmEmailSel.MAPIsignoff
  SendEmailFromClient = strError

Exit Function

LocalErr:
  strError = Err.Description
  GoTo TidyAndExit

End Function


Private Function MAPISignon() As Boolean

  On Local Error GoTo LocalErr
  
  MAPISignon = True
  If MAPISession1.SessionID = 0 Then
    'No session currently exists
    MAPISession1.DownLoadMail = False
    MAPISession1.SignOn
    If Err > 0 Then
      If Err <> 32001 And Err <> 32003 Then
        'MH20020820 Fault 4317
        'MsgBox Error$, 48, "Mail Error"
        MsgBox "Email not configured correctly." & vbCrLf & _
               IIf(Err.Description <> vbNullString, "(" & Trim(Err.Description) & ")", ""), _
               vbExclamation, "Mail Error"
      End If
      MAPISignon = False
    Else
      MAPIMessages1.SessionID = MAPISession1.SessionID
    End If
  End If

Exit Function

LocalErr:


End Function

Private Function MAPIsignoff() As Integer
  
  On Error Resume Next
  MAPIsignoff = True
  If MAPISession1.SessionID <> 0 Then
    'Session currently exists
    MAPISession1.SignOff
    If Err > 0 Then
      MsgBox Error$, 48, "Mail Error"
      MAPIsignoff = False
    Else
      MAPIMessages1.SessionID = 0
    End If
  End If

End Function


