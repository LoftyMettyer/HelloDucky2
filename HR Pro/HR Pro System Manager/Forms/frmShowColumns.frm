VERSION 5.00
Begin VB.Form frmShowColumns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show Columns"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5063
   Icon            =   "frmShowColumns.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4335
      TabIndex        =   0
      Top             =   4605
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5550
      TabIndex        =   1
      Top             =   4605
      Width           =   1200
   End
   Begin VB.Frame fraDefinition 
      Caption         =   "Show Columns :"
      Height          =   4335
      Left            =   100
      TabIndex        =   2
      Top             =   100
      Width           =   6660
      Begin VB.CheckBox chkViewThisColumn 
         Caption         =   "Column Name"
         Height          =   195
         Index           =   0
         Left            =   200
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   2160
      End
   End
End
Attribute VB_Name = "frmShowColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const X_OFFSET = 350
Private Const Y_OFFSET = 2550
Private Const X_BORDER = 100
Private Const Y_BORDER = 100

Private mpropColumnTypes As HRProSystemMgr.Properties

' Load the property column bag that we're dealing with.
Public Property Let PropertySet(ppropColumnTypes As HRProSystemMgr.Properties)
  Set mpropColumnTypes = ppropColumnTypes
End Property

Private Sub cmdCancel_Click()

UnLoad Me

End Sub

Private Sub cmdOK_Click()

Dim iCount As Integer

' Save the selected column views
For iCount = 1 To chkViewThisColumn.Count - 1
  mpropColumnTypes(chkViewThisColumn(iCount).Caption).value = IIf(chkViewThisColumn(iCount).value = vbChecked, True, False)
Next iCount

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

  Dim iCount As Integer
  Dim iXOffset As Integer
  Dim iYOffset As Integer
  
  iXOffset = 0 'X_BORDER
  iYOffset = 0 'Y_BORDER
  
  For iCount = 1 To mpropColumnTypes.Count
    Load chkViewThisColumn(iCount)
    
    chkViewThisColumn(iCount).Top = chkViewThisColumn(0).Top + iXOffset
    chkViewThisColumn(iCount).Left = chkViewThisColumn(0).Left + iYOffset
    
    chkViewThisColumn(iCount).Visible = True
    chkViewThisColumn(iCount).Caption = mpropColumnTypes.Item(iCount).Name
    chkViewThisColumn(iCount).value = IIf(CBool(mpropColumnTypes.Item(iCount).value) = True, vbChecked, vbUnchecked)
  
    chkViewThisColumn(iCount).Enabled = (mpropColumnTypes.Item(iCount).Name <> "Name")
  
    If (iCount Mod 10) = 0 Then
      iXOffset = 0
      iYOffset = iYOffset + Y_OFFSET
    Else
      iXOffset = iXOffset + X_OFFSET
    End If
  
  Next iCount
  
  ' Make form look nice and pretty
  ResizeForm

End Sub

Private Sub ResizeForm()

  Dim iAcross As Integer
  Dim iDown As Integer
  
  iAcross = (chkViewThisColumn.Count \ 10) + 1
  iDown = IIf(iAcross > 1, 10, chkViewThisColumn.Count)
  
  ' Set frame size
  fraDefinition.Top = X_BORDER
  fraDefinition.Left = Y_BORDER
  fraDefinition.Width = (iAcross * Y_OFFSET)
  fraDefinition.Height = X_OFFSET * IIf(iDown = 10, iDown + 1, iDown) + 100
  
  If fraDefinition.Width < (cmdCancel.Left - cmdOK.Left + cmdCancel.Width) Then
    fraDefinition.Width = (cmdCancel.Left - cmdOK.Left + cmdCancel.Width)
  End If
  
  ' Set form size
  Me.Width = fraDefinition.Width + (Y_BORDER * 3)
  Me.Height = fraDefinition.Height + cmdOK.Height + 800
  
  ' Set button locations
  cmdOK.Top = fraDefinition.Top + fraDefinition.Height + X_BORDER
  cmdCancel.Top = cmdOK.Top
  cmdCancel.Left = fraDefinition.Left + fraDefinition.Width - cmdCancel.Width
  cmdOK.Left = cmdCancel.Left - 1350
  
  'cmdCancel.Left = cmdOk.Left + cmdOk.Width + 100


End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


