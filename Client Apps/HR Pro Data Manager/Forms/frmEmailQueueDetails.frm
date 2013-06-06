VERSION 5.00
Begin VB.Form frmEmailQueueDetails 
   Caption         =   "Email Queue Message"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailQueueDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEmail 
      Caption         =   "&Resend..."
      Height          =   400
      Left            =   5280
      TabIndex        =   19
      Top             =   5200
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print..."
      Height          =   400
      Left            =   3960
      TabIndex        =   18
      Top             =   5200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   400
      Left            =   1440
      TabIndex        =   16
      Top             =   5200
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Pre&vious"
      Height          =   400
      Left            =   120
      TabIndex        =   15
      Top             =   5200
      Width           =   1215
   End
   Begin VB.Frame fraTop 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtDateSent 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtAttachment 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtSubject 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1400
         Width           =   6255
      End
      Begin VB.TextBox txtBCC 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1000
         Width           =   6255
      End
      Begin VB.TextBox txtCC 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   6255
      End
      Begin VB.TextBox txtTO 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   200
         Width           =   6255
      End
      Begin VB.Label lblDateSent 
         AutoSize        =   -1  'True
         Caption         =   "Date Sent :"
         Height          =   195
         Left            =   4800
         TabIndex        =   11
         Top             =   1860
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Attachment :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Subject :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1460
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bcc :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cc :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To :"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   285
      End
   End
   Begin VB.Frame fraBottom 
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   2440
      Width           =   7695
      Begin VB.TextBox txtMsgText 
         BackColor       =   &H8000000F&
         Height          =   2205
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6600
      TabIndex        =   17
      Top             =   5200
      Width           =   1215
   End
End
Attribute VB_Name = "frmEmailQueueDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdatData As clsDataAccess
Private mgrdEmailQueue As SSDBGrid


Private Function RemoveLastChar(strInput As String) As String
  If Right(strInput, 1) = ";" Then
    RemoveLastChar = Left(strInput, Len(strInput) - 1)
  Else
    RemoveLastChar = strInput
  End If
End Function

Private Sub ReadEmailContents(strDateSent As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strAttachment As String, strMsgText As String)

  txtTO.Text = RemoveLastChar(strTo)
  txtCC.Text = RemoveLastChar(strCC)
  txtBCC.Text = RemoveLastChar(strBCC)
  txtSubject.Text = IIf(strSubject = vbTab, "", strSubject)
  txtAttachment.Text = strAttachment
  txtMsgText.Text = strMsgText
  
  'Need to add line feeds after carriage returns...
  txtMsgText.Text = Replace(txtMsgText.Text, vbCr, vbCrLf)

  'Now it has added two line feeds (I dunno why??) so need to remove one
  txtMsgText.Text = Replace(txtMsgText.Text, Chr(13) & Chr(10) & Chr(10), vbCrLf)

End Sub

Public Sub Initialise(grdEmailQueue As SSDBGrid)

  Set mgrdEmailQueue = grdEmailQueue

  ShowDetails
  RefreshButtons
  Me.Show vbModal

End Sub

Private Sub ShowDetails()

  Dim lngLinkID As Long
  Dim lngRecordID As Long
  Dim strDateSent As String
  Dim strRecordDesc As String
  'Dim strColumnValue As String
  Dim strUserName As String
  Dim strTo As String
  Dim strCC As String
  Dim strBCC As String
  Dim strSubject As String
  Dim strAttachment As String
  Dim strMsgText As String
  Dim lngWorkflowInstanceID As Long
  Dim lngQueueID As Long
  Dim blnRecalculateRecordDesc As Boolean
  
  txtTO.Text = vbNullString
  txtTO.Tag = 0
  txtCC.Text = vbNullString
  txtCC.Tag = 0
  txtBCC.Text = vbNullString
  txtBCC.Tag = 0
    
  With mgrdEmailQueue
    If (.Rows > 0) And .SelBookmarks.Count = 1 Then

'      If (.Columns("LinkID").Value = vbNullString) _
'        And (.Columns("WorkflowInstanceID").Value = vbNullString) Then
'
'        .MovePrevious
'
'        If (.Columns("LinkID").Value = vbNullString) _
'          And (.Columns("WorkflowInstanceID").Value = vbNullString) Then
'
'          Exit Sub
'        End If
'      End If

      lngQueueID = .Columns("QueueID").Value
      lngLinkID = .Columns("LinkID").Value
      lngRecordID = .Columns("RecordID").Value
      strDateSent = .Columns("Email Sent").Value
      strTo = .Columns("RepTo").Value
      strCC = .Columns("RepCC").Value
      strBCC = .Columns("RepBCC").Value
      strSubject = .Columns("Subject").Value
      strMsgText = .Columns("MsgText").Value
      strAttachment = .Columns("Attachment").Value
      lngWorkflowInstanceID = .Columns("WorkflowInstanceID").Value
      blnRecalculateRecordDesc = .Columns("RecalculateRecordDesc").Value

      txtDateSent.Text = strDateSent
      If blnRecalculateRecordDesc Or strTo = vbNullString Then
        CalculateEmail lngLinkID, lngRecordID, lngQueueID   'lngLinkID, lngRecordID, strRecordDesc, "", strUserName, strDateSent
        If Not blnRecalculateRecordDesc Then
          MsgBox "No details have been stored for the current email.  This information has been calculated from the current email link but may not match the email which has actually been sent.", vbCritical
        End If
      Else
        ReadEmailContents strDateSent, strTo, strCC, strBCC, strSubject, strAttachment, strMsgText
      End If

    End If
  End With

End Sub

Private Sub cmdEmail_Click()

  Dim frmSendEmail As frmEmailSel
  Dim strMsgText As String


  If GetSystemSetting("email", "method", 1) = 0 Then
    MsgBox "Unable to resend this message as server side emails are currently disabled." & vbCrLf & _
           "Please contact your system administrator.", vbCritical, Me.Caption
    Exit Sub
  End If
  
  If MsgBox("Are you sure you want to resend this entry?", vbYesNo + vbQuestion, "Email Queue") <> vbYes Then
    Exit Sub
  End If

  strMsgText = txtMsgText.Text & vbCrLf & _
    "Resent on " & Format(Now, DateFormat) & _
    " at " & Format(Now, "hh:nn") & _
    " by " & gsUserName

  Set frmSendEmail = New frmEmailSel
  frmSendEmail.ResendEmailQueueEntry txtTO.Text, txtCC.Text, txtBCC.Text, txtSubject.Text, txtAttachment.Text, strMsgText
  Set frmSendEmail = Nothing

End Sub

Private Sub cmdNext_Click()

  With mgrdEmailQueue
    .MoveNext
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark
  End With
  ShowDetails
  RefreshButtons

End Sub

Private Sub cmdOK_Click()
  'Hiding rather than unload means it remembers the size for next time!
  'Unload Me
  Me.Hide
End Sub


Private Sub CalculateEmail(lngLinkID As Long, lngRecordID As Long, lngQueueID As Long)

  On Local Error GoTo LocalErr

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.spASREmail_" & CStr(lngLinkID)
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("hResult", adInteger, adParamReturnValue)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("QueueID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = lngQueueID

    Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = lngRecordID

    Set pmADO = .CreateParameter("Username", adVarChar, adParamInput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO
    pmADO.Value = gsUserName

    Set pmADO = .CreateParameter("To", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("Cc", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("Bcc", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("Subject", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("Message", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("Attachment", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    .Execute
    
    If .Parameters(0).Value > 0 Then
      MsgBox "This queue entry no longer matches the link filter and has been deleted.", vbInformation
      Unload Me
    Else
      txtTO.Text = .Parameters("To").Value
      txtCC.Text = .Parameters("Cc").Value
      txtBCC.Text = .Parameters("Bcc").Value
      txtSubject.Text = .Parameters("Subject").Value
      txtMsgText.Text = .Parameters("Message").Value
      txtAttachment.Text = .Parameters("Attachment").Value
      txtDateSent.Text = vbNullString
    End If

  End With

  Set pmADO = Nothing
  Set cmADO = Nothing

Exit Sub

LocalErr:
  MsgBox "Error populating email details", vbCritical

End Sub


Private Sub cmdPrevious_Click()
  
  With mgrdEmailQueue
    .MovePrevious
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark
  End With
  ShowDetails
  RefreshButtons

End Sub

Private Sub RefreshButtons()

  With mgrdEmailQueue
    cmdPrevious.Enabled = (.FirstRow + .Row > 1)
    cmdNext.Enabled = (.FirstRow + .Row < .Rows)
    cmdEmail.Enabled = (Me.txtDateSent.Text <> vbNullString)
  End With

End Sub

Private Sub cmdPrint_Click()

  Dim objPrintDef As clsPrintDef

  Set objPrintDef = New clsPrintDef

  With objPrintDef
    If objPrintDef.IsOK Then

      '.TabsOnPage = 3
      If .PrintStart(False) Then
        .PrintHeader "Email Queue Entry : " & _
                     mgrdEmailQueue.Columns("Email Title").Value & _
                     " (" & mgrdEmailQueue.Columns("RecDesc").Value & ")"

        .PrintNormal "Email Title: " & mgrdEmailQueue.Columns("Email Title").Value
        .PrintNormal "Record Description: " & mgrdEmailQueue.Columns("RecDesc").Value
        .PrintNormal "Column Name: " & mgrdEmailQueue.Columns("Column Name").Value
        .PrintNormal "Column Value: " & mgrdEmailQueue.Columns("Column Value").Value
        .PrintNormal "Email Due: " & mgrdEmailQueue.Columns("Email Due").Value
        .PrintNormal "Email Sent: " & mgrdEmailQueue.Columns("Email Sent").Value
        .PrintNormal
        .PrintNormal
        
        .PrintNormal "To: " & txtTO.Text
        .PrintNormal "Cc: " & txtCC.Text
        .PrintNormal "Bcc: " & txtBCC.Text
        .PrintNormal
        .PrintNormal "Subject: " & txtSubject.Text
        .PrintNormal
        .PrintNormal "Attachment: " & txtAttachment.Text
        .PrintNormal
        .PrintNormal "Message Text: " & txtMsgText.Text

        .PrintEnd
        .PrintConfirm "Email Queue Entry", "Email Queue Entry"

      End If
  
    End If
  End With
  
  Set objPrintDef = Nothing

End Sub

Private Sub Form_Load()
  RemoveIcon Me
  Hook Me.hWnd, (Me.Width - Me.ScaleWidth) + 6705, (Me.Height - Me.ScaleHeight) + 3660
End Sub

Private Sub Form_Resize()

  Dim lLeft As Long
  Dim lTop As Long
  Dim lWidth As Long
  Dim lHeight As Long

  Const GAP = 100
  Const GAP2 = 200

  'FORM
'  If Me.ScaleHeight < 3660 Then
'    Me.Height = (Me.Height - Me.ScaleHeight) + 3660
'    Exit Sub
'  End If
'  If Me.ScaleWidth < 6705 Then
'    Me.Width = (Me.Width - Me.ScaleWidth) + 6705
'    Exit Sub
'  End If


  'fraTOP
  lLeft = GAP
  lTop = GAP
  lWidth = Me.ScaleWidth - GAP2
  lHeight = fraTop.Height
  If lWidth >= 0 And lHeight >= 0 Then
    fraTop.Move lLeft, lTop, lWidth, lHeight
  End If

  'fraBOTTOM
  lTop = fraTop.Height + GAP2
  lHeight = Me.ScaleHeight - (lTop + cmdOK.Height + GAP2)
  If lWidth >= 0 And lHeight >= 0 Then
    fraBottom.Move lLeft, lTop, lWidth, lHeight
  End If

  'txtMSGTEXT
  lTop = GAP2
  lWidth = lWidth - GAP2
  lHeight = lHeight - (GAP * 3)
  If lWidth >= 0 And lHeight >= 0 Then
    txtMsgText.Move lLeft, lTop, lWidth, lHeight
  End If

  'txtTO, txtCC, txtBCC, txtSUBJECT
  lWidth = fraTop.Width - (txtTO.Left + GAP)
  If lWidth >= 0 Then
    txtTO.Width = lWidth
    txtCC.Width = lWidth
    txtBCC.Width = lWidth
    txtSubject.Width = lWidth
  End If

  'txtATTACHMENT, lblDATESENT, txtDATESENT
  lWidth = fraTop.Width - (txtAttachment.Left + GAP)
  lWidth = (lWidth - (lblDateSent.Width + txtDateSent.Width + GAP2))
  If lWidth >= 0 Then
    txtAttachment.Width = lWidth
  End If
  lLeft = txtAttachment.Left + txtAttachment.Width + GAP
  lblDateSent.Left = lLeft
  lLeft = lLeft + lblDateSent.Width + GAP
  txtDateSent.Left = lLeft

  'cmdOK
  lLeft = Me.ScaleWidth - (cmdOK.Width + GAP)
  lTop = Me.ScaleHeight - (cmdOK.Height + GAP)
  If lLeft > GAP And lTop > GAP Then
    cmdOK.Move lLeft, lTop
  End If

  'cmdEmail
  lLeft = lLeft - (cmdOK.Width + GAP)
  If lLeft > GAP And lTop > GAP Then
    cmdEmail.Move lLeft, lTop
  End If

  'cmdPrint
  lLeft = lLeft - (cmdEmail.Width + GAP)
  If lLeft > GAP And lTop > GAP Then
    cmdPrint.Move lLeft, lTop
  End If

  cmdPrevious.Top = lTop
  cmdNext.Top = lTop

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

