VERSION 5.00
Begin VB.Form frmTrigger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trigger"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14610
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTrigger.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraContent 
      Caption         =   "Content :"
      Height          =   5610
      Left            =   60
      TabIndex        =   3
      Top             =   2190
      Width           =   14500
      Begin VB.TextBox txtContent 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5130
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmTrigger.frx":000C
         Top             =   330
         Width           =   14175
      End
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details :"
      Height          =   2040
      Left            =   45
      TabIndex        =   2
      Top             =   60
      Width           =   14500
      Begin VB.ComboBox cboCodePosition 
         Height          =   315
         ItemData        =   "frmTrigger.frx":0012
         Left            =   1650
         List            =   "frmTrigger.frx":001C
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   825
         Width           =   2985
      End
      Begin VB.CheckBox chkIsSystem 
         Caption         =   "Is System Owned"
         Enabled         =   0   'False
         Height          =   300
         Left            =   7110
         TabIndex        =   7
         Top             =   315
         Width           =   2115
      End
      Begin VB.TextBox txtName 
         Height          =   330
         Left            =   1650
         TabIndex        =   6
         Top             =   315
         Width           =   4395
      End
      Begin VB.Label lblCodePosition 
         Caption         =   "Code Position :"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label lblName 
         Caption         =   "Name :"
         Height          =   225
         Left            =   225
         TabIndex        =   4
         Top             =   375
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   13335
      TabIndex        =   1
      Top             =   7935
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   12030
      TabIndex        =   0
      Top             =   7935
      Width           =   1200
   End
End
Attribute VB_Name = "frmTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjTriggerObject As clsTableTrigger

Private mbReadOnly As Boolean
Private mbCancelled As Boolean
Private mbLoading As Boolean
Private mbLocked As Boolean

Public Property Let Locked(ByRef bValue As Boolean)
  mbLocked = bValue
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(ByVal NewValue As Boolean)
  If Not mbLoading Then
    cmdOK.Enabled = NewValue And Not mbReadOnly
  End If
End Property

Public Property Get TriggerObject() As clsTableTrigger
  Set TriggerObject = mobjTriggerObject
End Property

Public Property Let TriggerObject(ByVal NewValue As clsTableTrigger)
  Set mobjTriggerObject = NewValue
End Property

Public Function PopulateControls() As Boolean
  
  On Error GoTo ErrorTrap

  Dim bOK As Boolean
  Dim lngTableID As Long

  bOK = True
  lngTableID = mobjTriggerObject.TableID
  mbLoading = True

  SetComboItem cboCodePosition, mobjTriggerObject.CodePosition
  chkIsSystem.value = IIf(mobjTriggerObject.IsSystem, vbChecked, vbUnchecked)
  txtName.Text = mobjTriggerObject.Name
  txtContent.Text = mobjTriggerObject.content
  
TidyUpAndExit:
  mbLoading = False
  PopulateControls = bOK
  Exit Function
  
ErrorTrap:
  bOK = False

End Function

Private Sub chkIsSystem_Click()
  Me.Changed = True
End Sub

Private Sub cmdCancel_Click()

  If Me.Changed Then
    Select Case MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
    Case vbYes
      cmdOK_Click
      Exit Sub
    Case vbCancel
      Exit Sub
    End Select
  End If

  Me.Hide

End Sub

Private Function ValidDefinition() As Boolean
  ValidDefinition = True
End Function

Private Sub SaveDefinition()

  TriggerObject.IsSystem = IIf(chkIsSystem.value = vbChecked, True, False)
  TriggerObject.Name = txtName.Text
  TriggerObject.CodePosition = TriggerCodePosition.AfterU02Update
  TriggerObject.content = txtContent.Text

End Sub

Private Sub cmdOK_Click()

  If ValidDefinition = False Then
    Exit Sub
  End If

  SaveDefinition
  mbCancelled = False
  Me.Hide

End Sub

Private Sub Form_Load()

  mbCancelled = True
  mbReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode) Or _
                  mbLocked
                  
End Sub

Private Sub txtContent_Change()
  Me.Changed = True
End Sub

Private Sub txtName_Change()
  Me.Changed = True
End Sub
