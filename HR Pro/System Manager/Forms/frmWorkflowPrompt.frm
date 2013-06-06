VERSION 5.00
Begin VB.Form frmWorkflowPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5054
   Icon            =   "frmWorkflowPrompt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optChoice 
      Caption         =   "???"
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   3500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2925
      TabIndex        =   1
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1515
      TabIndex        =   0
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmWorkflowPrompt.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblInfo1 
      BackStyle       =   0  'Transparent
      Caption         =   "Which of the following outbound flows do you wish the link to start from?"
      Height          =   405
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   4005
   End
End
Attribute VB_Name = "frmWorkflowPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Properties.
Private mfCancelled As Boolean
Private miOutboundFlowCode As Integer
Private miDecisionCaptionType As DecisionCaptionType
Private mwfElement As VB.Control

Public Property Let DecisionCaptionType(piDecisionCaptionType As DecisionCaptionType)
  miDecisionCaptionType = piDecisionCaptionType
End Property

Public Property Set Element(pwfElement As VB.Control)
  Set mwfElement = pwfElement

  If mwfElement.ElementType = elem_Decision Then
    DecisionCaptionType = mwfElement.DecisionCaptionType
  End If
  
  OutboundFlowInfo = mwfElement.OutboundFlows_Information
  
End Property

Public Property Get OutboundFlowCode() As Integer
  ' Return the selected code.
  OutboundFlowCode = miOutboundFlowCode
End Property

Public Property Let OutboundFlowCode(piNewValue As Integer)
  ' Set the selected code.
  Dim optTemp As OptionButton
  
  miOutboundFlowCode = piNewValue
  
  For Each optTemp In optChoice
    If (optTemp.Index > 0) _
      And val(optTemp.Tag) = miOutboundFlowCode Then
      
      optTemp.value = True
      Exit For
    End If
  Next optTemp
  Set optTemp = Nothing
  
End Property

Private Sub cmdCancel_Click()
  ' Flag that the copy has been cancelled..
  mfCancelled = True
  
  ' Unload the form.
  UnLoad Me

End Sub

Private Sub cmdOK_Click()
  ' Flag that the change/deletion has been confirmed.
  mfCancelled = False
  
  ' Unload the form.
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
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr
End Sub

Public Property Get Cancelled() As Boolean
  ' Return the 'cancelled' property.
  Cancelled = mfCancelled
End Property

Public Property Let OutboundFlowInfo(pavInfo As Variant)
  ' Load the different outbound flow info into the option group.
  Dim iLoop As Integer
  Dim sngLastTop As Single
  
  sngLastTop = optChoice(0).Top
  
  For iLoop = 1 To UBound(pavInfo, 2)
    Load optChoice(optChoice.UBound + 1)
    
    With optChoice(optChoice.UBound)
      If (mwfElement.ElementType = elem_WebForm) _
        Or (mwfElement.ElementType = elem_StoredData) Then
        .Caption = pavInfo(7, iLoop)
      Else
        .Caption = GetDecisionCaptionDescription(miDecisionCaptionType, (pavInfo(1, iLoop) = 1))
      End If
      
      .Tag = pavInfo(1, iLoop)
      .Top = sngLastTop
      .Left = 1080
      .Visible = True
      .TabIndex = iLoop - 1
    End With
    
    sngLastTop = sngLastTop + 360
  Next iLoop
  
  cmdOK.Top = sngLastTop + 180
  cmdCancel.Top = cmdOK.Top
  
  Me.Height = cmdOK.Top + cmdOK.Height + 600
  
  miOutboundFlowCode = CInt(optChoice(1).Tag)
  
End Property

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optChoice_Click(Index As Integer)
  ' Update the global variable.
  miOutboundFlowCode = CInt(optChoice(Index).Tag)

End Sub


