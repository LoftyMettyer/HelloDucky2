VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTechSupport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Technical Support"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1002
   Icon            =   "frmTechSupport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmTechSupport.frx":000C
   ScaleHeight     =   2655
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSupportContacts 
      Caption         =   "Contacts for Technical Support :"
      Height          =   1900
      Left            =   120
      TabIndex        =   1
      Top             =   100
      Width           =   4380
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone :"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   9
         Top             =   400
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax :"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   8
         Top             =   750
         Width           =   375
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   7
         Top             =   1095
         Width           =   600
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web site :"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   6
         Top             =   1450
         Width           =   735
      End
      Begin VB.Label lblTel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+44 (0) 01582 714800"
         Height          =   195
         Left            =   1350
         TabIndex        =   5
         Top             =   400
         Width           =   2445
      End
      Begin VB.Label lblFax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+44 (0) 01582 714814"
         Height          =   195
         Left            =   1350
         TabIndex        =   4
         Top             =   750
         Width           =   2265
      End
      Begin VB.Label lblEMail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "support"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1350
         MouseIcon       =   "frmTechSupport.frx":0316
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1100
         Width           =   2800
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1350
         MouseIcon       =   "frmTechSupport.frx":0468
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1450
         Width           =   2800
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3300
      TabIndex        =   0
      Top             =   2150
      Width           =   1200
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   700
      Top             =   2100
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
      Left            =   1395
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   150
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTechSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function MAPIsendMessage(ByVal MSubject As String, ByVal MSendTo As String, ByVal MMessage As String, ByVal MReceipt As Integer, MFORM As Form, ByVal SHOWCOMPOSE As Integer) As Integer
  Dim OldCap As String, UNam As String, PARSESTR As String
  
  On Error Resume Next
  
  'Log onto Mail server if not already logged onto it
  MAPISignon
  MAPIsendMessage = False
  If MAPISession1.SessionID <> 0 And Len(Trim(MSendTo)) > 0 Then
    With MAPIMessages1
      OldCap = MFORM.Caption
      .MsgIndex = -1
      .RecipDisplayName = MSendTo
      If Err = 0 Then
        .MsgSubject = MSubject
        If Len(Trim(MMessage)) > 0 Then
          .MsgNoteText = MMessage
        End If
        .MsgReceiptRequested = MReceipt
        MFORM.Caption = "Displaying Mail Message, Please Wait..."
        .Send SHOWCOMPOSE
        If Err > 0 Then
          MsgBox Error$, 48, "Send Mail Message"
        Else
          MAPIsendMessage = True
        End If
      Else
        MsgBox Error$, 48, "Send Mail Message"
      End If
      MFORM.Caption = OldCap
    End With
  End If
  
End Function
Function MAPIsignoff() As Integer
  'End a MAPI session
  Screen.MousePointer = 11
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
  Screen.MousePointer = 0
End Function

Function MAPISignon() As Integer
  'Begin a MAPI session
  Screen.MousePointer = 11
  On Error Resume Next
  MAPISignon = True
  If MAPISession1.SessionID = 0 Then
    'No session currently exists
    MAPISession1.DownLoadMail = False
    MAPISession1.SignOn
    If Err > 0 Then
      If Err <> 32001 And Err <> 32003 Then
        MsgBox Error$, 48, "Mail Error"
      End If
      MAPISignon = False
    Else
      MAPIMessages1.SessionID = MAPISession1.SessionID
    End If
  End If
  Screen.MousePointer = 0
End Function


Private Sub cmdOK_Click()
  Unload Me

End Sub

Private Sub Form_Load()
  
  'Get all the required support details
  lblTel.Caption = GetSystemSetting("Support", "Telephone No", "")
  lblFax.Caption = GetSystemSetting("Support", "Fax", "")
  lblEMail.Caption = GetSystemSetting("Support", "Email", "")
  lblURL.Caption = GetSystemSetting("Support", "Webpage", "")

End Sub


'Private Function GetSupportInfo() As Recordset
'
'  Dim sSQL As String
'  Dim rsSupport As rdo.rdoResultset
'
'  'Get all the required support details
'  sSQL = "SELECT * FROM ASRSysConfig"
'  Set rsSupport = rdoCon.OpenResultset(sSQL, _
'    rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'
'  If Not rsSupport.BOF And Not rsSupport.EOF Then
'    lblTel = Trim(rsSupport!supportTelNo)
'    lblFax = Trim(rsSupport!SupportFax)
'    lblEMail = Trim(rsSupport!SupportEMail)
'    lblURL = Trim(rsSupport!URL)
'  Else
'    MsgBox "No Support information found."
'  End If
'  rsSupport.Close
'  Set rsSupport = Nothing
'
'End Function


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Redo link colour
  lblEMail.ForeColor = &HFF0000
  lblURL.ForeColor = &HFF0000
  DoEvents
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub fraSupportContacts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Redo link colour
  lblEMail.ForeColor = &HFF0000
  lblURL.ForeColor = &HFF0000
  DoEvents
  
End Sub


Private Sub lblEMail_Click()
  
  'Show that the 'hyperlink' has been clicked on
  'lblEMail.ForeColor = &H800080
  
  'Show the compose message window
  MAPIsendMessage "HR Pro Support Query - Security Mgr", lblEMail.Caption, "", 0, Me, 1

End Sub


Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Highlight the link
  lblEMail.ForeColor = vbRed
  DoEvents

End Sub


Private Sub lblTel_Click()

End Sub

Private Sub lblTitle_Click(Index As Integer)

End Sub

Private Sub lblURL_Click()
  On Error GoTo ErrTrap

  Dim plngID As Integer
  
  'Show that the 'hyperlink' has been clicked on
  'lblURL.ForeColor = &H800080
  DoEvents
  
  ' Replaced the following line in the hope of making ShellExecute work on all PCs.
  ' Dont think it worked !
  plngID = ShellExecute(0&, vbNullString, Trim(lblURL.Caption), vbNullString, vbNullString, vbMaximizedFocus)
  
  If plngID = 0 Then
    ' Uh oh...the browser wasnt initiated...tell the user
    MsgBox "HR Pro cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"
  End If
  
  Exit Sub
  
ErrTrap:
    MsgBox "HR Pro cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"
  
End Sub


Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Highlight the link
  lblURL.ForeColor = vbRed
  DoEvents

End Sub


