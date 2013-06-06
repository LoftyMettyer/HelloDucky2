VERSION 5.00
Begin VB.Form frmEmailLinkAttachmentSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Attachment"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1017
   Icon            =   "frmEmailLinkAttachmentSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4840
      Left            =   100
      TabIndex        =   2
      Top             =   100
      Width           =   5050
      Begin VB.TextBox txtAttachmentsPath 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000011&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   2970
      End
      Begin VB.ListBox lstAttachmentList 
         Height          =   3960
         Left            =   200
         TabIndex        =   3
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblAttachmentsPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachments Path :"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3950
      TabIndex        =   1
      Top             =   5080
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2640
      TabIndex        =   0
      Top             =   5080
      Width           =   1200
   End
End
Attribute VB_Name = "frmEmailLinkAttachmentSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrFileName As String
Private mblnCancelled As Boolean
Private mblnNoFiles As Boolean

Public Property Get FileName() As String
  FileName = mstrFileName
End Property

Public Property Let FileName(ByVal strNewValue As String)
  mstrFileName = strNewValue
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal blnNewValue As Boolean)
  mblnCancelled = blnNewValue
End Property

Private Sub Form_Load()
  mblnCancelled = True  'Cancelled unless click ok
  txtAttachmentsPath.Text = gstrEmailAttachmentPath
  lstAttachmentList_Populate
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  'lstAttachmentList.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
End Sub

Private Sub lstAttachmentList_Click()

  With lstAttachmentList
    If .ListIndex >= 0 Then
      mstrFileName = .List(.ListIndex)
    End If
  End With

End Sub

Private Sub lstAttachmentList_DblClick()
  cmdOK_Click
End Sub

Private Sub cmdOK_Click()
  If Not mblnNoFiles Then
    mblnCancelled = False
  End If
  UnLoad Me
End Sub

Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub lstAttachmentList_Populate()

  Dim rsTemp As New ADODB.Recordset
  Dim strSQL As String

  On Error GoTo LocalErr

  mblnNoFiles = True
  
  'strSQL = "master..xp_cmdshell 'dir " & Chr(34) & txtAttachmentsPath.Text & Chr(34) & " /a-d /b'"
  'Set rsTemp = rdoCon.OpenResultset(strSQL, rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
  'strSQL = "exec sp_ASRServerDir 3, '" & txtAttachmentsPath.Text & "'"
  strSQL = "master..xp_dirtree '" & txtAttachmentsPath.Text & "',1,1"
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  lstAttachmentList.Clear
  
  If Not rsTemp.BOF And Not rsTemp.EOF Then
  
    'Select Case Trim(rsTemp.rdoColumns(0).Value)
    'Case "The device is not ready.", _
    '     "The system cannot find the path specified.", _
    '     "File Not Found"
    '  lstAttachmentList.AddItem "<" & rsTemp.rdoColumns(0).Value & ">"
    'Case Else
      Do While Not rsTemp.EOF
        If rsTemp.Fields("File").value = 1 Then
          lstAttachmentList.AddItem rsTemp.Fields(0).value
        End If
        rsTemp.MoveNext
      Loop
      mblnNoFiles = False
    'End Select
  End If

  'Just double check that its not left blank
  If lstAttachmentList.ListCount = 0 Then
    lstAttachmentList.AddItem "No files found"
  End If

LocalErr:
  'If ASRDEVELOPMENT Then
  '  Debug.Print Err.Description
  '  Stop
  'End If
  If Not (rsTemp Is Nothing) Then
    If rsTemp.State = 1 Then
      rsTemp.Close
    End If
  End If
  Set rsTemp = Nothing

End Sub

