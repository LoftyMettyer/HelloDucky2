VERSION 5.00
Begin VB.Form frmControlChange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Type Change"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5011
   Icon            =   "frmControlChange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1600
      TabIndex        =   3
      Top             =   2500
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3000
      TabIndex        =   4
      Top             =   2500
      Width           =   1200
   End
   Begin VB.OptionButton optChoice 
      Caption         =   "&Delete the controls from the screens"
      Height          =   315
      Index           =   1
      Left            =   200
      TabIndex        =   1
      Top             =   2000
      Width           =   3960
   End
   Begin VB.OptionButton optChoice 
      Caption         =   "C&hange the controls to the new type"
      Height          =   315
      Index           =   0
      Left            =   200
      TabIndex        =   0
      Top             =   1600
      Value           =   -1  'True
      Width           =   3960
   End
   Begin VB.Label lblInfo2 
      BackStyle       =   0  'Transparent
      Caption         =   "These screens will require reviewing."
      Height          =   315
      Left            =   195
      TabIndex        =   6
      Top             =   1200
      Width           =   3360
   End
   Begin VB.Label lblScreenList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<screen list>"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   795
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label lblInfo1 
      BackStyle       =   0  'Transparent
      Caption         =   "You are changing the control type that represents this column in the following screens :"
      Height          =   405
      Left            =   195
      TabIndex        =   2
      Top             =   195
      Width           =   4245
   End
End
Attribute VB_Name = "frmControlChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Properties.
Private gfCancelled As Boolean
Private gfChangeControls As Boolean

Private Sub cmdCancel_Click()
  ' Flag that the copy has been cancelled..
  gfCancelled = True
  
  ' Unload the form.
  UnLoad Me

End Sub

Private Sub cmdOK_Click()
  ' Flag that the change/deletion has been confirmed.
  gfCancelled = False
  
  ' Unload the form.
  UnLoad Me

End Sub


Private Sub Form_Initialize()
  ' Initialise variables.
  gfChangeControls = True

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
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr

End Sub


Public Property Get DeleteControls() As Boolean
  ' Return whether or not the controls are to be deleted.
  DeleteControls = Not gfChangeControls
  
End Property

Public Property Get ChangeControls() As Boolean
  ' Return whether or not the controls are to be changed.
  ChangeControls = gfChangeControls
  
End Property


Public Property Get Cancelled() As Boolean
  ' Return the 'cancelled' property.
  Cancelled = gfCancelled
  
End Property





Public Property Let ScreenList(ByVal vNewValue As Variant)
  ' Set the list of screens.
  Dim iIndex As Integer
  Dim lngYPosition As Long
  Dim sScreenList() As String
  Dim ctlScreenListLabel As Label
  
  Const iYSCREENLISTGAP = 250
  Const iYAFTERSCREENLISTGAP = 150
  Const iYBUTTONGAP = 500
  Const iYGAP = 400
  Const iYBORDERGAP = 200
  
  sScreenList = vNewValue
  
  ' Clear the screen list control array
  For Each ctlScreenListLabel In lblScreenList
    If ctlScreenListLabel.Index > 0 Then
      UnLoad ctlScreenListLabel
    End If
  Next ctlScreenListLabel
  Set ctlScreenListLabel = Nothing
  
  ' Load labels for each screen in the list.
  lngYPosition = lblScreenList(0).Top
  For iIndex = 1 To UBound(sScreenList)
    Load lblScreenList(iIndex)
    With lblScreenList(iIndex)
      .Caption = "'" & Trim(sScreenList(iIndex)) & "'"
      .Top = lngYPosition
      .Visible = True
    End With
    lngYPosition = lngYPosition + iYSCREENLISTGAP
  Next iIndex
  
  ' Position the other controls in relation to the list of screens.
  lblInfo2.Top = lngYPosition + iYAFTERSCREENLISTGAP
  optChoice(0).Top = lblInfo2.Top + iYGAP
  optChoice(1).Top = optChoice(0).Top + iYGAP
  cmdOK.Top = optChoice(1).Top + iYBUTTONGAP
  cmdCancel.Top = cmdOK.Top
  Me.Height = cmdOK.Top + 1000
  Me.Height = cmdOK.Top + cmdOK.Height + iYBORDERGAP + _
    (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))

End Property

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optChoice_Click(Index As Integer)
  ' Update the global variable.
  gfChangeControls = (Index = 0)
  
End Sub


