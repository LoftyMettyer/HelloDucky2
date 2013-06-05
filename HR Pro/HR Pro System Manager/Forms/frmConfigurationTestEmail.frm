VERSION 5.00
Begin VB.Form frmConfigurationTestEmail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Email Message"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5084
   Icon            =   "frmConfigurationTestEmail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   3165
   End
   Begin VB.TextBox txtRecipient 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   3165
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3045
      TabIndex        =   2
      Top             =   3270
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   1770
      TabIndex        =   1
      Top             =   3270
      Width           =   1200
   End
   Begin VB.TextBox txtMessage 
      Height          =   1500
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1620
      Width           =   4000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject :"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   780
      Width           =   645
   End
   Begin VB.Label lblEmailServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   300
      Width           =   285
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "Message :"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "frmConfigurationTestEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMethod As Long
Private mstrProfile As String
Private mstrServer As String
Private mstrAccount As String

Public Sub Initialise(lngMethod As Long, strProfile As String, strServer As String, strAccount As String)

  mlngMethod = lngMethod
  mstrProfile = strProfile
  mstrServer = strServer
  mstrAccount = strAccount

  txtRecipient.Text = GetSystemSetting("test email", "recipient", "")
  txtSubject.Text = GetSystemSetting("test email", "subject", "Test Message")
  txtMessage.Text = Database.ServerName & " " & gsDatabaseName & vbCrLf & CStr(Now)

End Sub

Private Sub cmdOK_Click()

  Dim rsTemp As New ADODB.Recordset
  Dim strProcName As String
  Dim strSQL As String
  Dim fOK As Boolean
  Dim strErrorMessage As String

  On Local Error GoTo LocalErr

  Screen.MousePointer = vbHourglass
  strErrorMessage = vbNullString

  If Trim(txtRecipient.Text) = vbNullString Then
    MsgBox "Please enter a recipient.", vbCritical, Me.Caption
    txtRecipient.SetFocus
    Exit Sub
  End If


  SaveSystemSetting "test email", "recipient", txtRecipient.Text
  SaveSystemSetting "test email", "subject", txtSubject.Text

  strProcName = UniqueSQLObjectName("tmpsp_ASREmailTest", 4)
  fOK = CreateEmailSendStoredProcedure(strProcName, mlngMethod, mstrProfile, mstrServer, mstrAccount)

  If fOK Then

    If txtSubject.Text = vbNullString Then
      txtSubject.Text = vbTab
    End If

    strSQL = "EXEC " & strProcName & " 0, " & _
             "'" & Replace(txtRecipient.Text, "'", "''") & "', '', '', " & _
             "'" & Replace(txtSubject.Text, "'", "''") & "', " & _
             "'" & Replace(txtMessage.Text, "'", "''") & "', ''"
    'gADOCon.Execute strSQL
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
    If mlngMethod = 3 Then
      If Not IsNull(rsTemp.Fields("ErrorMessage").value) Then
        strErrorMessage = rsTemp.Fields("ErrorMessage").value
      End If
    End If
    
    If rsTemp.State <> adStateClosed Then
      rsTemp.Close
    End If
    Set rsTemp = Nothing
    
    DropUniqueSQLObject strProcName, 4
  
  End If

TidyAndExit:
  Screen.MousePointer = vbDefault
  If fOK Then
    Me.Hide
    If strErrorMessage <> vbNullString Then
      MsgBox strErrorMessage, vbExclamation, Me.Caption
    Else
      MsgBox "Mail Sent", vbInformation, Me.Caption
    End If
  End If
  
  UnLoad Me

Exit Sub

LocalErr:
  strErrorMessage = Err.Description
  Resume TidyAndExit

End Sub


Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()

End Sub
