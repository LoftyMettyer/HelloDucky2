VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{AD837810-DD1E-44E0-97C5-854390EA7D3A}#3.2#0"; "COA_Navigation.ocx"
Begin VB.Form frmTechSupport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Technical Support"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmTechSupport.frx":000C
   ScaleHeight     =   2355
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Contacts for Technical Support :"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   100
      Width           =   4580
      Begin COANavigation.COA_Navigation navSupport 
         Height          =   215
         Left            =   1350
         TabIndex        =   6
         Top             =   1110
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   370
         Caption         =   "navSupport"
         DisplayType     =   0
         NavigateIn      =   0
         NavigateTo      =   ""
         InScreenDesigner=   0   'False
         ColumnID        =   0
         ColumnName      =   ""
         Selected        =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontSize        =   8.25
         FontStrikethrough=   0   'False
         FontUnderline   =   -1  'True
         ForeColor       =   16711680
         NavigateOnSave  =   0   'False
      End
      Begin COANavigation.COA_Navigation navEmail 
         Height          =   215
         Left            =   1350
         TabIndex        =   7
         Top             =   750
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   370
         Caption         =   "navEmail"
         DisplayType     =   0
         NavigateIn      =   0
         NavigateTo      =   ""
         InScreenDesigner=   0   'False
         ColumnID        =   0
         ColumnName      =   ""
         Selected        =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontSize        =   8.25
         FontStrikethrough=   0   'False
         FontUnderline   =   -1  'True
         ForeColor       =   16711680
         NavigateOnSave  =   0   'False
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone :"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   5
         Top             =   400
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         Height          =   215
         Index           =   3
         Left            =   195
         TabIndex        =   4
         Top             =   750
         Width           =   600
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web site :"
         Height          =   215
         Index           =   4
         Left            =   195
         TabIndex        =   3
         Top             =   1110
         Width           =   870
      End
      Begin VB.Label lblTel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+44 (0) 01582 714800"
         Height          =   195
         Left            =   1350
         TabIndex        =   2
         Top             =   405
         Width           =   1935
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   600
      Top             =   2055
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
      Left            =   1215
      Top             =   2055
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3500
      TabIndex        =   0
      Top             =   1815
      Width           =   1200
   End
End
Attribute VB_Name = "frmTechSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    'Unload the form
    Unload Me

End Sub

Private Sub Form_Load()

'    Dim sSQL As String
'    Dim rsSupport As Recordset
'
'    'Get all the required support details
'    Set rsSupport = datGeneral.GetSupportInfo
'
'    If Not rsSupport.BOF And Not rsSupport.EOF Then
'        lblTel = rsSupport!SupportTelNo
'        lblFax = rsSupport!SupportFax
'        lblEMail = rsSupport!SupportEMail
'        lblURL = rsSupport!URL
'    Else
'        MsgBox "No Support information found."
'    End If
'    rsSupport.Close
'    Set rsSupport = Nothing
  
  lblTel.Caption = GetSystemSetting("Support", "Telephone No", "")
'  lblFax.Caption = GetSystemSetting("Support", "Fax", "")
  'lblEMail.Caption = GetSystemSetting("Support", "Email", "")
  'lblURL.Caption = GetSystemSetting("Support", "Webpage", "")

  navEmail.Caption = GetSystemSetting("Support", "Email", "")
  navEmail.NavigateTo = "mailto:" & navEmail.Caption
  
  navSupport.Caption = "http://" & GetSystemSetting("Support", "Webpage", "")
  navSupport.NavigateTo = navSupport.Caption

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

'Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'  ' Redo link colour
'  lblEMail.ForeColor = &HFF0000
'  lblURL.ForeColor = &HFF0000
'  DoEvents
'
'End Sub

'Private Sub lblEMail_Click()
'
'  'RH - June 99
'
'  'Show that the 'hyperlink' has been clicked on
'  'lblEMail.ForeColor = &H800080
'
'  'Show the compose message window
'  MAPIsendMessage "HR Pro Support Query - Data Manager", lblEMail.Caption, "", 0, Me, 1
'
'End Sub


'Private Sub lblURL_Click()
'  On Error GoTo ErrTrap
'
'  Dim plngID As Integer
'
'  'Show that the 'hyperlink' has been clicked on
'  'lblURL.ForeColor = &H800080
'  DoEvents
'
'  ' Replaced the following line in the hope of making ShellExecute work on all PCs.
'  ' Dont think it worked !
'  plngID = ShellExecute(0&, vbNullString, Trim(lblURL.Caption), vbNullString, vbNullString, vbMaximizedFocus)
'
'  If plngID = 0 Then
'    ' Uh oh...the browser wasnt initiated...tell the user
'    MsgBox "HR Pro cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"
'  End If
'
'  Exit Sub
'
'ErrTrap:
'    MsgBox "HR Pro cannot automatically open your default web browser." & vbCrLf & vbCrLf & "Please open your web browser manually and navigate to the " & vbCrLf & "web address which has been placed in your clipboard." & IIf(Err.Description = "", "", vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"), vbInformation + vbOKOnly, "Technical Support"
'
'End Sub

'Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'  ' Highlight the link
'  lblEMail.ForeColor = vbRed
'  DoEvents
'
'End Sub
'
'Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'  ' Highlight the link
'  lblURL.ForeColor = vbRed
'  DoEvents
'
'End Sub
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'  ' Redo link colour
'  lblEMail.ForeColor = &HFF0000
'  lblURL.ForeColor = &HFF0000
'  DoEvents
'
'End Sub

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
'        MsgBox Error$, 48, "Mail Error"
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
'      MsgBox Error$, 48, "Mail Error"
'      MAPIsignoff = False
'    Else
'      MAPIMessages1.SessionID = 0
'    End If
'  End If
'  Screen.MousePointer = 0
'End Function

'Function MAPIsendMessage(ByVal MSubject As String, ByVal MSendTo As String, ByVal MMessage As String, ByVal MReceipt As Integer, MFORM As Form, ByVal SHOWCOMPOSE As Integer) As Integer
'  Dim OldCap As String, UNam As String, ParseStr As String
'
'  On Error Resume Next
'
'  'Log onto Mail server if not already logged onto it
'  MAPISignon
'  MAPIsendMessage = False
'  If MAPISession1.SessionID <> 0 And Len(Trim(MSendTo)) > 0 Then
'    With MAPIMessages1
'      OldCap = MFORM.Caption
'      .MsgIndex = -1
'      .RecipDisplayName = MSendTo
'      If Err = 0 Then
'        .MsgSubject = MSubject
'        If Len(Trim(MMessage)) > 0 Then
'          .MsgNoteText = MMessage
'        End If
'        .MsgReceiptRequested = MReceipt
'        MFORM.Caption = "Initialising Mail, Please Wait..."
'        .Send SHOWCOMPOSE
'        If Err > 0 Then
'          MsgBox Error$, 48, "Send Mail Message"
'        Else
'          MAPIsendMessage = True
'        End If
'      Else
'        MsgBox Error$, 48, "Send Mail Message"
'      End If
'      MFORM.Caption = OldCap
'    End With
'  End If
'
'End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
  
End Sub


