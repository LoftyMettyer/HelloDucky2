VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.Form frmDrop 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2835
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3720
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3720
   Begin XtremeSuiteControls.PushButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   2520
      TabIndex        =   4
      Top             =   1650
      Width           =   1095
      _Version        =   851969
      _ExtentX        =   1940
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdClear 
      Height          =   420
      Left            =   2520
      TabIndex        =   3
      Top             =   1140
      Width           =   1095
      _Version        =   851969
      _ExtentX        =   1940
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdSelect 
      Default         =   -1  'True
      Height          =   420
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _Version        =   851969
      _ExtentX        =   1940
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Select"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdNew 
      Height          =   420
      Left            =   2520
      TabIndex        =   2
      Top             =   630
      Width           =   1100
      _Version        =   851969
      _ExtentX        =   1931
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "New..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin MSComctlLib.ListView lsvList 
      Height          =   2505
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4419
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Manager"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbSelected As Boolean
Private msItem As String
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdClear_Click()

    Selected = True
    Item = ""
    Me.Hide

End Sub

Private Sub cmdNew_Click()

    Selected = True
    Item = "Add New Table Entry"
    Me.Hide

End Sub

Private Sub cmdSelect_Click()

    Selected = True
    Item = lsvList.SelectedItem.Text
    Me.Hide

End Sub

Public Property Get Selected() As Boolean

    Selected = mbSelected

End Property

Public Property Let Selected(ByVal bSelected As Boolean)

    mbSelected = bSelected

End Property

Public Property Get Item() As String

    Item = msItem

End Property

Public Property Let Item(ByVal sItem As String)

    msItem = sItem

End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  ' <Menu> Key - Dunno what this does, but its always been there and I'm afraid to take it out.
  If KeyCode = vbKeyMenu Then
    frmDrop.Hide
  End If

  ' Ctrl key pressed - pass back to parent control
  If (Shift And vbCtrlMask) > 0 Then
    frmDrop.Hide
    RaiseEvent KeyDown(KeyCode, Shift)
  End If

End Sub

Private Sub Form_Load()

  Selected = False

End Sub

Private Sub lsvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  If lsvList.SortKey = ColumnHeader.Index - 1 Then
    lsvList.SortOrder = IIf(lsvList.SortOrder = 0, 1, 0)
    Exit Sub
  End If
  
  lsvList.SortKey = ColumnHeader.Index - 1
  lsvList.SortOrder = lvwAscending

End Sub

Private Sub lsvList_DblClick()

  If lsvList.ListItems.Count > 0 Then
    cmdSelect_Click
  End If

End Sub

