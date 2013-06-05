VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmEmailSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Email Recipients"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1034
   Icon            =   "frmEmailSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4335
      TabIndex        =   3
      Top             =   3590
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4335
      TabIndex        =   2
      Top             =   3100
      Width           =   1200
   End
   Begin VB.CheckBox chkIncRecDesc 
      Caption         =   "&Include Record Description in message"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4100
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin SSDataWidgets_B.SSDBGrid ssGrdRecipients 
      Height          =   3900
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4140
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   6
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      BalloonHelp     =   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   661
      Columns(0).Caption=   "To"
      Columns(0).Name =   "To"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   2
      Columns(1).Width=   741
      Columns(1).Caption=   "Cc"
      Columns(1).Name =   "CC"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Style=   2
      Columns(2).Width=   767
      Columns(2).Caption=   "Bcc"
      Columns(2).Name =   "BCC"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Style=   2
      Columns(3).Width=   6033
      Columns(3).Caption=   "Recipient"
      Columns(3).Name =   "Recipient"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "EmailID"
      Columns(4).Name =   "EmailID"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "IsGroup"
      Columns(5).Name =   "IsGroup"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   7302
      _ExtentY        =   6879
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   4335
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   4950
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmEmailSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum EmailSendTypes
  giEMAILSEND_RECORDEDIT = 0
  giEMAILSEND_EVENTLOG = 1
End Enum

Private mdatData As clsDataAccess          'DataAccess Class
Private mlngRecordID As Long
Private mlngTableID As Long
Private mlngDefaultEmailID As Long
Private mstrRecordDetails As String
Private miSendType As EmailSendTypes
Private mstrEventLogIDs As String

Public Sub Initialise(lngTableID As Long, lngRecordID As Long, strRecordDetails As String)
  
  Dim rsRecipients As Recordset
  Dim rsTemp As Recordset
  Dim strSQL As String
  
  mlngTableID = lngTableID
  mlngRecordID = lngRecordID
  mstrRecordDetails = strRecordDetails
  miSendType = giEMAILSEND_RECORDEDIT

  Set mdatData = New clsDataAccess

  Call GetTableDetails


  strSQL = "SELECT Name, EmailID " & _
           " FROM ASRSysEmailAddress " & _
           " WHERE tableID = 0 OR tableID = " & CStr(mlngTableID) & _
           " ORDER BY Name"   'MH20030516 Fault 4538
  
           'MH20030819 Fault 6242
           '" WHERE tableID = " & CStr(mlngTableID) &

  Set rsRecipients = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  With rsRecipients
  
    ssGrdRecipients.RemoveAll
    Do While Not .EOF

      ssGrdRecipients.AddItem _
          IIf(!EmailID = mlngDefaultEmailID, "1", "0") & _
          vbTab & "0" & vbTab & "0" & vbTab & _
          Trim(!Name) & vbTab & CStr(!EmailID) & vbTab
      
      .MoveNext
    Loop

    ssGrdRecipients.ScrollBars = IIf(ssGrdRecipients.Rows > 15, ssScrollBarsVertical, ssScrollBarsNone)

  End With

  rsRecipients.Close
  Set rsRecipients = Nothing
  
  chkIncRecDesc.Value = IIf(CBool(GetUserSetting("Email", "IncludeRecDesc", True)), vbChecked, vbUnchecked)
  
  'Refresh the buttons
  'RefreshButtons

End Sub

Public Sub SetupEventLogSend(pstrEventIDs As String)

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEmailSel.SetupEventLogSend()"

  Dim iCount As Integer
  Dim iNextPosition As Integer

  ' Set module to send event log entry(ies)
  miSendType = giEMAILSEND_EVENTLOG
  mstrEventLogIDs = pstrEventIDs

  ' Clear any exiting email addresses
  ssGrdRecipients.RemoveAll
   
  Dim rsEmailGroup As ADODB.Recordset
  Dim strEmailGroup As String
          
  strEmailGroup = vbNullString
  strEmailGroup = strEmailGroup & "SELECT [ASRSysEmailGroupName].[Name] AS 'Name'," & vbCrLf
  strEmailGroup = strEmailGroup & "       '0' + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       '0' + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       '0' + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       [ASRSysEmailGroupName].[Name] + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       CONVERT(varchar,[ASRSysEmailGroupName].[EmailGroupID]) + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       '1' AS 'EmailString'" & vbCrLf
  strEmailGroup = strEmailGroup & "FROM [ASRSysEmailGroupName]" & vbCrLf
  strEmailGroup = strEmailGroup & "UNION" & vbCrLf
  strEmailGroup = strEmailGroup & "SELECT (SELECT [ASRSysSystemSettings].[SettingValue]" & vbCrLf
  strEmailGroup = strEmailGroup & "        FROM [ASRSysSystemSettings]" & vbCrLf
  strEmailGroup = strEmailGroup & "        WHERE ([ASRSysSystemSettings].[Section] = 'Support')" & vbCrLf
  strEmailGroup = strEmailGroup & "           AND ([ASRSysSystemSettings].[SettingKey] = 'Email')" & vbCrLf
  strEmailGroup = strEmailGroup & "       ) AS 'Name'," & vbCrLf
  strEmailGroup = strEmailGroup & "       '0' + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       '0' + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       '0' + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       (SELECT [ASRSysSystemSettings].[SettingValue]" & vbCrLf
  strEmailGroup = strEmailGroup & "        FROM [ASRSysSystemSettings]" & vbCrLf
  strEmailGroup = strEmailGroup & "        WHERE ([ASRSysSystemSettings].[Section] = 'Support')" & vbCrLf
  strEmailGroup = strEmailGroup & "           AND ([ASRSysSystemSettings].[SettingKey] = 'Email')" & vbCrLf
  strEmailGroup = strEmailGroup & "       ) + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       (SELECT [ASRSysSystemSettings].[SettingValue]" & vbCrLf
  strEmailGroup = strEmailGroup & "        FROM [ASRSysSystemSettings]" & vbCrLf
  strEmailGroup = strEmailGroup & "        WHERE ([ASRSysSystemSettings].[Section] = 'Support')" & vbCrLf
  strEmailGroup = strEmailGroup & "           AND ([ASRSysSystemSettings].[SettingKey] = 'Email')" & vbCrLf
  strEmailGroup = strEmailGroup & "       ) + char(9) +" & vbCrLf
  strEmailGroup = strEmailGroup & "       '0' AS 'EmailString'" & vbCrLf
  strEmailGroup = strEmailGroup & "ORDER BY 'Name'" & vbCrLf

  Set rsEmailGroup = datGeneral.GetReadOnlyRecords(strEmailGroup)
  
  With rsEmailGroup
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        ssGrdRecipients.AddItem .Fields("EmailString").Value
        .MoveNext
      Loop
    End If
    .Close
  End With
  
  'Hide the include record description checkbox
  chkIncRecDesc.Visible = False

  ' Refresh the buttons
  'RefreshButtons

TidyUpAndExit:
  Set rsEmailGroup = Nothing
  gobjErrorStack.PopStack
  Exit Sub
  
ErrorTrap:
  gobjErrorStack.HandleError
  Resume TidyUpAndExit
  
End Sub


Private Function GetTableDetails() As Long
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT DefaultEmailID " & _
           " FROM ASRSysTables " & _
           " WHERE tableID = " & CStr(mlngTableID)
  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  mlngDefaultEmailID = 0
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    mlngDefaultEmailID = rsTemp!DefaultEmailID
  End If

  rsTemp.Close
  Set rsTemp = Nothing

End Function

Private Sub cmdCancel_Click()
  Unload Me
End Sub

'Function MAPISignon() As Integer
'  'Begin a MAPI session
'  Screen.MousePointer = 11
'  On Error Resume Next
'  MAPISignon = True
'  If MAPISession1.SessionID = 0 Then
'    'No session currently exists
'    MAPISession1.DownLoadMail = False
'    MAPISession1.SignOn
'    If Err > 0 Then
'      If Err <> 32001 And Err <> 32003 Then
'        'MH20020820 Fault 4317
'        'COAMsgBox Error$, 48, "Mail Error"
'        COAMsgBox "Email not configured correctly." & vbCrLf & _
'               IIf(Err.Description <> vbNullString, "(" & Trim(Err.Description) & ")", ""), _
'               vbExclamation, "Mail Error"
'      End If
'      MAPISignon = False
'    Else
'      MAPIMessages1.SessionID = MAPISession1.SessionID
'    End If
'  End If
'  Screen.MousePointer = 0
'End Function
'
'Function MAPIsignoff() As Integer
'  'End a MAPI session
'  Screen.MousePointer = 11
'  On Error Resume Next
'  MAPIsignoff = True
'  If MAPISession1.SessionID <> 0 Then
'    'Session currently exists
'    MAPISession1.SignOff
'    If Err > 0 Then
'      COAMsgBox Error$, 48, "Mail Error"
'      MAPIsignoff = False
'    Else
'      MAPIMessages1.SessionID = 0
'    End If
'  End If
'  Screen.MousePointer = 0
'End Function

Private Sub cmdOK_Click()
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEmailSel.cmdOK_Click()"

  ' depending on type of email run a different compose message function
  Select Case miSendType
    Case giEMAILSEND_EVENTLOG
      SendEventLog
    Case giEMAILSEND_RECORDEDIT
      SendRecord
  End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Function SendRecord() As Boolean

  Dim objOutputEmail As clsOutputEMail
  
  Dim strTo As String
  Dim strCC As String
  Dim strBCC As String
  Dim strSubject As String
  Dim strMsgText As String
  Dim strAttachment As String
  Dim blnPause As Boolean
  Dim strError As String
  
  Dim intCol As Integer
  Dim intRow As Integer
  Dim lngEmailID As Long
  Dim strEmailAddr As String
  
  On Error GoTo LocalErr

  Screen.MousePointer = vbHourglass
  
  strTo = vbNullString
  strCC = vbNullString
  strBCC = vbNullString
  
  ssGrdRecipients.Redraw = False
  ssGrdRecipients.Refresh
  
  ssGrdRecipients.MoveFirst
  For intRow = 1 To ssGrdRecipients.Rows
    For intCol = 0 To 2
      If ssGrdRecipients.Columns(intCol).Value <> 0 Then

        lngEmailID = Val(ssGrdRecipients.Columns(4).Value)
        strEmailAddr = GetEmailAddress(lngEmailID, mlngRecordID)

        If Trim(strEmailAddr) <> vbNullString Then
          Select Case intCol
            Case 0: strTo = strTo & IIf(strTo <> "", ";", "") & strEmailAddr
            Case 1: strCC = strCC & IIf(strCC <> "", ";", "") & strEmailAddr
            Case 2: strBCC = strBCC & IIf(strBCC <> "", ";", "") & strEmailAddr
          End Select
        Else
          strError = "Unable to use Email address <" & ssGrdRecipients.Columns(3).Value & "> for this record as it is empty"
          Exit For
        End If

      End If
    Next

    ssGrdRecipients.MoveNext
  Next

  If strError = vbNullString Then
    strMsgText = IIf(chkIncRecDesc.Value = vbChecked, mstrRecordDetails & vbCrLf & vbCrLf, "")
    
    Set objOutputEmail = New clsOutputEMail
    strError = objOutputEmail.SendEmailFromClient(strTo, strCC, strBCC, strSubject, strMsgText, strAttachment, True)
    Set objOutputEmail = Nothing
  End If
  
TidyAndExit:
  Screen.MousePointer = vbDefault
  
  If strError <> vbNullString Then
    ssGrdRecipients.Redraw = True
    COAMsgBox strError, vbExclamation + vbOKOnly, "Send Mail Message"
  Else
    Unload Me
  End If

Exit Function

LocalErr:
  strError = "Error sending email (" & Err.Description & ")"
  Resume TidyAndExit

End Function

Private Function SendEventLog() As Boolean

  Dim objOutputEmail As clsOutputEMail
  Dim colRecipients As clsEmailRecipients
  Dim objRecipient As clsEmailRecipient

  Dim strTo As String
  Dim strCC As String
  Dim strBCC As String
  Dim strSubject As String
  Dim strMsgText As String
  Dim strAttachment As String
  Dim blnPause As Boolean
  Dim strError As String
  
  Dim intCol As Integer
  Dim intRow As Integer
  Dim lngEmailID As Long
  Dim strEmailAddr As String
  
  On Error GoTo LocalErr

  Screen.MousePointer = vbHourglass
  
  strTo = vbNullString
  strCC = vbNullString
  strBCC = vbNullString
  
  ssGrdRecipients.Redraw = False
  ssGrdRecipients.Refresh
  
  ssGrdRecipients.MoveFirst
  For intRow = 1 To ssGrdRecipients.Rows
    For intCol = 0 To 2
    
      If ssGrdRecipients.Columns(intCol).Value <> 0 Then
        
        strEmailAddr = vbNullString
        If (ssGrdRecipients.Columns(5).Value = 1) Then
          Set colRecipients = New clsEmailRecipients
          colRecipients.Populate_Collection CStr(ssGrdRecipients.Columns(4).Value)

          strEmailAddr = vbNullString
          For Each objRecipient In colRecipients.Collection
            strEmailAddr = strEmailAddr & IIf(strEmailAddr <> "", ";", "") & Trim(objRecipient.FixedEmail)
          Next objRecipient

          Set objRecipient = Nothing
          Set colRecipients = Nothing

        ElseIf (Trim(ssGrdRecipients.Columns(4).Value) <> vbNullString) Then
          strEmailAddr = ssGrdRecipients.Columns(4).Value
        
        End If
        
        If Trim(strEmailAddr) <> vbNullString Then
          Select Case intCol
            Case 0: strTo = strTo & IIf(strTo <> "", ";", "") & strEmailAddr
            Case 1: strCC = strCC & IIf(strCC <> "", ";", "") & strEmailAddr
            Case 2: strBCC = strBCC & IIf(strBCC <> "", ";", "") & strEmailAddr
          End Select
        End If

      End If
        
    Next

    ssGrdRecipients.MoveNext
  Next

  If strError = vbNullString Then

    strSubject = GetSystemSetting("Licence", "Customer Name", "<<Unknown Customer>>") & " - OpenHR Event Log"
    strMsgText = GetEventDetails(mstrEventLogIDs)
    If Trim(strMsgText) = vbNullString Then
      COAMsgBox "The selected Event Log record(s) have been deleted by another User.", vbOKOnly + vbExclamation, "Event Log"
      GoTo TidyAndExit
    End If

    Set objOutputEmail = New clsOutputEMail
    strError = objOutputEmail.SendEmailFromClient(strTo, strCC, strBCC, strSubject, strMsgText, strAttachment, blnPause)
    Set objOutputEmail = Nothing
  End If

  SendEventLog = True

TidyAndExit:
  Screen.MousePointer = vbDefault
  
  If strError <> vbNullString Then
    ssGrdRecipients.Redraw = True
    COAMsgBox strError, vbExclamation + vbOKOnly, "Send Mail Message"
  Else
    Unload Me
  End If

Exit Function

LocalErr:
  strError = "Error sending email (" & Err.Description & ")"
  SendEventLog = False
  Resume TidyAndExit

End Function


Public Function SendBatchNotification(strSubject As String, strEventIDs As String, lngEmailGroupID As Long) As Boolean

  Dim objOutputEmail As clsOutputEMail
  Dim rsEmail As Recordset
  Dim strSQL As String
  
  Dim strTo As String
  Dim strMsgText As String
  
  Dim lngCount As Long
  Dim adoCmd As ADODB.Command
  Dim strErrors As String
  
  On Error GoTo LocalErr

  Screen.MousePointer = vbHourglass

  strSQL = "SELECT ASRSysEmailGroupItems.*," & _
           " ASRSysEmailAddress.Name as 'AddrName', ASRSysEmailAddress.Fixed as 'AddrFixed'" & _
           " FROM ASRSysEmailGroupItems" & _
           " JOIN ASRSysEmailAddress ON ASRSysEmailGroupItems.EmailDefID = ASRSysEmailAddress.EmailID" & _
           " WHERE EmailGroupID = " & CStr(lngEmailGroupID) & _
           " ORDER BY AddrName"
  Set rsEmail = datGeneral.GetReadOnlyRecords(strSQL)

  If rsEmail.BOF And rsEmail.EOF Then
    If Not gblnBatchJobsOnly Then
      COAMsgBox "Error retrieving email recipient(s)", vbExclamation, "Batch Job Notification"
    End If
    SendBatchNotification = False
    Exit Function
  End If

  Do While Not rsEmail.EOF
    strTo = strTo & _
      IIf(strTo <> vbNullString, ";", "") & _
      rsEmail!AddrFixed
    rsEmail.MoveNext
  Loop

  rsEmail.Close
  Set rsEmail = Nothing


  strMsgText = GetEventDetails(strEventIDs)
  Set objOutputEmail = New clsOutputEMail
  objOutputEmail.SendEmailFromServer strTo, "", "", strSubject, strMsgText, ""
  Set objOutputEmail = Nothing


  SendBatchNotification = True

Exit Function

LocalErr:
  If Not gblnBatchJobsOnly Then
    If gobjProgress.Visible Then
      gobjProgress.Visible = False
    End If
    COAMsgBox "Error sending email" & _
      IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString), vbExclamation, "Batch Job Notification"
  End If
  On Error Resume Next
  'frmEmailSel.MAPIsignoff
  'Set frmEmailSel = Nothing
  SendBatchNotification = False

End Function


Private Function GetEventDetails(pstrEventIDs As String) As String

  Dim strMessage As String
  Dim strSQL As String
  Dim rstDetailRecords As Recordset
  Dim rstBatchDetails As Recordset
  Dim blnNewEvent As Boolean
  Dim lngEventID As Long
  Dim lngDetailsCount As Long
  Dim lngCurrentLogCount As Long
  Dim strDateFormat As String
  
  Const NO_DETAILS = vbCrLf & "There are no details for this event log entry"
  
  strDateFormat = DateFormat
  
  'Retrieve the details for the selected records
  strSQL = vbNullString
  strSQL = strSQL & "SELECT [E].[ID], [E].[DateTime], [E].[EndTime], [E].[Duration], "
  strSQL = strSQL & "   [E].[Name], [E].[UserName], [E].[Mode], [D].[Notes] , [E].[Type], "
  strSQL = strSQL & "   [E].[SuccessCount], [E].[FailCount], [E].[Status], "
  strSQL = strSQL & "   [E].[BatchRunID], [E].[BatchJobID], [E].[BatchName], "
  strSQL = strSQL & " (SELECT COUNT(DISTINCT [C].[ID]) FROM [ASRSysEventLogDetails] [C] WHERE [C].[EventLogID] = [E].[ID]) AS 'DetailsCount' "
  strSQL = strSQL & "FROM [ASRSysEventLog] [E] "
  strSQL = strSQL & "       LEFT OUTER JOIN [ASRSysEventLogDetails] [D] "
  strSQL = strSQL & "       ON [D].[EventLogID] = [E].[ID] "
  strSQL = strSQL & "WHERE [E].[ID] IN (" & pstrEventIDs & ")"
  
  Set rstDetailRecords = datGeneral.GetReadOnlyRecords(strSQL)
  
  ' Add each note to the message string
  If Not (rstDetailRecords.EOF And rstDetailRecords.BOF) Then
    lngEventID = -1
    blnNewEvent = False
    strMessage = ""
    lngDetailsCount = 0
    lngCurrentLogCount = 0
    
    Do Until rstDetailRecords.EOF
    
      If lngEventID <> rstDetailRecords.Fields("ID").Value Then
        If lngEventID <> -1 Then
          strMessage = strMessage & vbCrLf & vbCrLf & vbCrLf
        End If
      
        lngEventID = rstDetailRecords.Fields("ID")
        blnNewEvent = True
        lngCurrentLogCount = 0
        
        strMessage = strMessage & String(Len(rstDetailRecords.Fields("Name")) + 30, "-") & vbCrLf
        strMessage = strMessage & "Event Name : " & rstDetailRecords.Fields("Name") & vbCrLf
        strMessage = strMessage & String(Len(rstDetailRecords.Fields("Name")) + 30, "-") & vbCrLf
        strMessage = strMessage & "Mode : " & vbTab & vbTab & IIf(rstDetailRecords.Fields("Mode").Value, "Batch", "Manual") & vbCrLf
        strMessage = strMessage & vbCrLf
        strMessage = strMessage & "Start Time : " & vbTab & Format(rstDetailRecords.Fields("DateTime"), strDateFormat & " hh:mm:ss") & vbCrLf
        strMessage = strMessage & "End Time : " & vbTab & IIf(IsNull(rstDetailRecords.Fields("EndTime")), "", Format(rstDetailRecords.Fields("EndTime"), strDateFormat & " hh:mm:ss")) & vbCrLf
        strMessage = strMessage & "Duration : " & vbTab & FormatEventDuration(IIf(IsNull(rstDetailRecords.Fields("Duration")), 0, rstDetailRecords.Fields("Duration"))) & vbCrLf
        strMessage = strMessage & vbCrLf
        strMessage = strMessage & "Type : " & vbTab & vbTab & GetUtilityType(rstDetailRecords.Fields("Type")) & vbCrLf
        strMessage = strMessage & "Status : " & vbTab & GetUtilityStatus(rstDetailRecords.Fields("Status")) & vbCrLf
        strMessage = strMessage & "User name : " & vbTab & rstDetailRecords.Fields("Username") & vbCrLf
        
        If Not IsNull(rstDetailRecords.Fields("BatchRunID")) Then
          If rstDetailRecords.Fields("BatchRunID") > 0 Then
            strSQL = vbNullString
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "FROM [ASRSysEventLog] "
            strSQL = strSQL & "WHERE [BatchRunID] = " & rstDetailRecords.Fields("BatchRunID")
            strSQL = strSQL & " ORDER BY [ID]"
            
            Set rstBatchDetails = datGeneral.GetReadOnlyRecords(strSQL)

            If Not (rstBatchDetails.BOF And rstBatchDetails.EOF) Then
              strMessage = strMessage & vbCrLf
              strMessage = strMessage & "Batch Job Name : " & rstBatchDetails.Fields("BatchName") & vbCrLf
              strMessage = strMessage & vbCrLf
              strMessage = strMessage & "All Jobs in Batch : " & vbCrLf & vbCrLf
              
              Do Until rstBatchDetails.EOF
                strMessage = strMessage & GetUtilityType(rstBatchDetails.Fields("Type")) & " - " & rstBatchDetails.Fields("Name") & " (" & GetUtilityStatus(rstBatchDetails.Fields("Status")) & ")" & vbCrLf
                rstBatchDetails.MoveNext
              Loop
            End If
            
          End If
        End If
        
        If (Not IsNull(rstDetailRecords.Fields("SuccessCount"))) And (Not IsNull(rstDetailRecords.Fields("FailCount"))) Then
          strMessage = strMessage & vbCrLf
          strMessage = strMessage & "Records Successful : " & vbTab & rstDetailRecords.Fields("SuccessCount") & vbCrLf
          strMessage = strMessage & "Records Failed : " & vbTab & rstDetailRecords.Fields("FailCount") & vbCrLf
        End If
        
        strMessage = strMessage & vbCrLf
        
        strMessage = strMessage & "Details : " & vbCrLf
        
        lngDetailsCount = rstDetailRecords.Fields("DetailsCount")
        lngCurrentLogCount = lngCurrentLogCount + 1
        If lngDetailsCount > 0 Then
          strMessage = strMessage & vbCrLf & vbCrLf & "***  Log entry " & lngCurrentLogCount & " of " & lngDetailsCount & "  ***" & vbCrLf
        End If
        strMessage = strMessage & IIf(((IsNull(rstDetailRecords.Fields("Notes"))) Or (Trim(rstDetailRecords.Fields("Notes")) = "")), NO_DETAILS, rstDetailRecords.Fields("Notes")) & vbCrLf
      Else
        blnNewEvent = False
        lngCurrentLogCount = lngCurrentLogCount + 1
        If lngDetailsCount > 0 Then
          strMessage = strMessage & vbCrLf & vbCrLf & "***  Log entry " & lngCurrentLogCount & " of " & lngDetailsCount & "  ***" & vbCrLf
        End If
        strMessage = strMessage & IIf(((IsNull(rstDetailRecords.Fields("Notes"))) Or (Trim(rstDetailRecords.Fields("Notes")) = "")), NO_DETAILS, rstDetailRecords.Fields("Notes")) & vbCrLf
      End If
      
      rstDetailRecords.MoveNext
    Loop
        
    Set rstDetailRecords = Nothing
  Else
  
''     JDM - 17/06/02 - Fault 3770 - Add details for successful records
'    strSQL = vbNullString
'    strSQL = strSQL & "SELECT [ASRSysEventLog].[Name], "
'    strSQL = strSQL & "       [ASRSysEventLog].[UserName], "
'    strSQL = strSQL & "       [ASRSysEventLog].[Mode], "
'    strSQL = strSQL & "       [ASRSysEventLog].[DateTime] "
'    strSQL = strSQL & "FROM  [ASRSysEventLog]"
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT [E].[ID], [E].[DateTime], [E].[EndTime], [E].[Duration], "
    strSQL = strSQL & "   [E].[Name], [E].[UserName], [E].[Mode], [E].[Type], "
    strSQL = strSQL & "   [E].[SuccessCount], [E].[FailCount], [E].[Status] "
    strSQL = strSQL & "FROM [ASRSysEventLog] [E] "
    strSQL = strSQL & "WHERE [E].[ID] IN (" & pstrEventIDs & ")"
  
    Set rstDetailRecords = datGeneral.GetReadOnlyRecords(strSQL)
    
    If Not rstDetailRecords.EOF And Not rstDetailRecords.BOF Then
      strMessage = strMessage & String(Len(rstDetailRecords.Fields("Name")) + 30, "-") & vbCrLf
      strMessage = strMessage & "Event Name : " & rstDetailRecords.Fields("Name") & vbCrLf
      strMessage = strMessage & String(Len(rstDetailRecords.Fields("Name")) + 30, "-") & vbCrLf
      strMessage = strMessage & "Mode : " & vbTab & vbTab & IIf(rstDetailRecords.Fields("Mode").Value, "Batch", "Manual") & vbCrLf
      strMessage = strMessage & vbCrLf
      strMessage = strMessage & "Start Time : " & vbTab & rstDetailRecords.Fields("DateTime") & vbCrLf
      strMessage = strMessage & "End Time : " & vbTab & IIf(IsNull(rstDetailRecords.Fields("EndTime")), "", rstDetailRecords.Fields("EndTime")) & vbCrLf
      strMessage = strMessage & "Duration : " & vbTab & FormatEventDuration(IIf(IsNull(rstDetailRecords.Fields("Duration")), 0, rstDetailRecords.Fields("Duration"))) & vbCrLf
      strMessage = strMessage & vbCrLf
      strMessage = strMessage & "Type : " & vbTab & vbTab & GetUtilityType(rstDetailRecords.Fields("Type")) & vbCrLf
      strMessage = strMessage & "Status : " & vbTab & vbTab & GetUtilityStatus(rstDetailRecords.Fields("Status")) & vbCrLf
      strMessage = strMessage & "User name : " & vbTab & rstDetailRecords.Fields("Username") & vbCrLf
      
      If (Not IsNull(rstDetailRecords.Fields("SuccessCount"))) And (Not IsNull(rstDetailRecords.Fields("FailCount"))) Then
        strMessage = strMessage & vbCrLf
        strMessage = strMessage & "Records Successful : " & vbTab & rstDetailRecords.Fields("SuccessCount") & vbCrLf
        strMessage = strMessage & "Records Failed : " & vbTab & vbTab & rstDetailRecords.Fields("FailCount") & vbCrLf
      End If
      
      strMessage = strMessage & vbCrLf
      
      strMessage = strMessage & "Details: " & vbCrLf & vbCrLf
      strMessage = strMessage & NO_DETAILS & vbCrLf
    
    End If
  
  End If

  GetEventDetails = strMessage
 
TidyUpAndExit:
  Set rstDetailRecords = Nothing
  Exit Function
  
ErrorTrap:
  GetEventDetails = ""
  GoTo TidyUpAndExit
  
End Function


Public Function SendEmail(strTo As String, strSubject As String, strMsgText As String, blnErrorMessage As Boolean, Optional blnPause As Boolean = False) As Boolean
  
  Dim objOutputEmail As clsOutputEMail
  Dim strError As String
  
  Set objOutputEmail = New clsOutputEMail
  strError = objOutputEmail.SendEmailFromClient(strTo, "", "", strSubject, strMsgText, "", blnPause)
  Set objOutputEmail = Nothing
  
  If strError <> "" And blnErrorMessage Then
    COAMsgBox strError, vbExclamation + vbOKOnly, "Send Mail Message"
  End If

End Function


Private Sub Form_Activate()
    Select Case miSendType
    Case 0
      'Personnel frmEmailSel
      Me.HelpContextID = 1115
    Case 1
      'Event Log frmEmailSel
      Me.HelpContextID = 1116
    End Select
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
End Sub


'Public Function ResendEmailQueueEntry(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String) As Boolean
'
'  Dim strError As String
'
'  strError = SendEmailFromServer(strTo, strCC, strBCC, strSubject, strMsgText, strAttachment)
'
'  If strError = vbNullString Then
'    COAMsgBox "Message resent.", vbInformation, "Email Queue"
'    ResendEmailQueueEntry = True
'  Else
'    COAMsgBox "Error sending email" & _
'      IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString), vbExclamation, "Email Queue"
'    ResendEmailQueueEntry = False
'  End If
'
'End Function
'
