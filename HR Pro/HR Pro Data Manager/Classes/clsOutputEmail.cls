VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrTo As String
Private mstrCC As String
Private mstrBCC As String
Private mstrSubject As String
Private mstrMsgText As String
Private mstrAttachment As String
Private mblnPause As Boolean
Private mstrError As String

Public Sub SetRecipientsFromGrid(objGrid As SSDBGrid)
  
  Dim intCol As Integer
  Dim intRow As Integer
  Dim lngEmailID As Long
  Dim strEmailAddr As String
  
  mstrTo = vbNullString
  mstrCC = vbNullString
  mstrBCC = vbNullString
  
  objGrid.MoveFirst
  For intRow = 1 To objGrid.Rows
    For intCol = 0 To 2
      If objGrid.Columns(intCol).Value <> 0 Then

        lngEmailID = Val(objGrid.Columns(4).Value)
        strEmailAddr = GetEmailAddress(lngEmailID, mlngRecordID)

        If Trim(strEmailAddr) <> vbNullString Then
          Select Case intCol
            Case 0: mstrTo = mstrTo & IIf(mstrTo <> "", ";", "") & strEmailAddr
            Case 1: mstrCC = mstrCC & IIf(mstrCC <> "", ";", "") & strEmailAddr
            Case 2: mstrBCC = mstrBCC & IIf(mstrBCC <> "", ";", "") & strEmailAddr
          End Select
        Else
          strError = "Unable to use Email address <" & objGrid.Columns(3).Value & "> for this record as it is empty"
          Exit For
        End If

      End If
    Next

    objGrid.MoveNext
  Next

End Sub

Public Sub SetRecipients(strTo As String, strCC As String, strBCC As String)
  mstrTo = strTo
  mstrCC = strCC
  mstrBCC = strBCC
End Sub

Public Sub SetContent(strSubject As String, strMessage As String, strAttachment As String)
  mstrSubject = strSubject
  mstrMsgText = strMsgText
  mstrAttachment = strAttachment
End Sub


Public Function SendEmailFromServer() As String

  Dim adoCmd As ADODB.Command
  Dim strError As String
  
  On Error GoTo LocalErr

  strError = vbNullString
  
  Set adoCmd = New ADODB.Command
  adoCmd.ActiveConnection = gADOCon
  
  gADOCon.Errors.Clear
  adoCmd.CommandText = "exec spASRSendMail 0, " & _
                       "'" & Replace(mstrTo, "'", "''") & "', " & _
                       "'" & Replace(mstrCC, "'", "''") & "', " & _
                       "'" & Replace(mstrBCC, "'", "''") & "', " & _
                       "'" & Replace(mstrSubject, "'", "''") & "', " & _
                       "'" & Left(Replace(mstrMsgText, "'", "''"), 7000) & "', " & _
                       "'" & Replace(mstrAttachment, "'", "''") & "'"
  adoCmd.Execute
  
TidyAndExit:
  SendEmailFromServer = strError

Exit Function

LocalErr:
  strError = Err.Description
  GoTo TidyAndExit

End Function


Public Function SendEmailFromClient(blnPause As Boolean) As String

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
  
  
  MAPISignon
  If MAPISession1.SessionID <> 0 Then
    With MAPIMessages1
      .Compose
      
      For lngRecipType = 1 To 3
      
        Select Case lngRecipType
        Case 1: strRecipients = mstrTo
        Case 2: strRecipients = mstrCC
        Case 3: strRecipients = mstrBCC
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

      .MsgSubject = mstrSubject
      .MsgNoteText = mstrMsgText

'      If (Trim(mstrAttachment) <> vbNullString) Then
'        strArray = Split(mstrAttachment, ";")
'        For lngIndex = LBound(strArray) To UBound(strArray)
'          If Trim(strArray(lngIndex)) <> vbNullString Then
'            .AttachmentIndex = .AttachmentCount
'            .AttachmentPathName = Trim(strArray(lngIndex))
'          End If
'        Next
'      End If

      .Send blnPause

    End With

  End If

TidyAndExit:
  SendEmailFromClient = strError

Exit Function

LocalErr:
  strError = Err.Description
  GoTo TidyAndExit

End Function
