VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl COASD_TabPage 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   870
      Left            =   225
      TabIndex        =   1
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1535
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox objContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   1665
      ScaleHeight     =   735
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "COASD_TabPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Click()
Public Event DblClick()

Private gfSelected As Boolean

Public ClientLeft As Long
Public ClientTop As Long
Public ClientHeight As Long
Public ClientWidth As Long
Public WFItemType As Integer

Public Sub AddTabPage(ByVal Caption As String)
  
  Dim iContainerIndex As Long
  iContainerIndex = objContainer.UBound + 1
  
  TabStrip1.Tabs.Add
  Load objContainer(iContainerIndex)
  With objContainer(iContainerIndex)
    .BackColor = UserControl.BackColor
    .BorderStyle = 0
    .Left = 50
    .Top = 50
    .Width = UserControl.Width - 100
    .Height = UserControl.Height - 100
    .Visible = True
    .ZOrder 1
  End With
  
  TabPages_Resize
  
End Sub

Public Property Get TabPage(ByVal PageNo As Long) As MSComctlLib.Tab
  Set TabPage = TabStrip1.Tabs(PageNo)
End Property

Public Property Get ControlPage(ByVal PageNo As Long) As StdPicture
  Set ControlPage = objContainer(PageNo)
End Property

Public Function Tabs() As Tabs
  Set Tabs = TabStrip1.Tabs
End Function

Private Sub objContainer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub objContainer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub objContainer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub TabStrip1_Click()
  
  Dim ctlPictureBox As PictureBox
  
  For Each ctlPictureBox In objContainer
    With ctlPictureBox
      If .Index = TabStrip1.SelectedItem.Index Then
        .Enabled = True
        .Visible = True
        .ZOrder 0
      Else
        .Enabled = False
        .Visible = False
      End If
    End With
  Next ctlPictureBox
  
  RaiseEvent Click

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal pbValue As Boolean)
  UserControl.Enabled = pbValue
End Property

Public Property Get Selected() As Boolean
  Selected = gfSelected
End Property

Public Property Let Selected(ByVal pfNewValue As Boolean)
  gfSelected = pfNewValue
End Property

Public Property Get SelectedItem() As MSComctlLib.Tab
  Set SelectedItem = TabStrip1.SelectedItem
End Property

Public Property Let BackColor(ByVal pColNewColor As OLE_COLOR)
  UserControl.BackColor = pColNewColor
End Property

Private Function TabPages_Resize() As Boolean
  ' Resize the tab pages.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim ctlPictureBox As PictureBox
  Dim Xframe As Long
  Dim YFrame As Long
  
  Xframe = 20
  YFrame = 20
  
  ' Position and size the tabstrip to fill the form's client area.
  TabStrip1.Move Xframe, YFrame, UserControl.ScaleWidth - (Xframe * 2), UserControl.ScaleHeight - (YFrame * 2)
  
  ' Position and size the picture box containers of the tabstrip.
  For Each ctlPictureBox In objContainer
    ctlPictureBox.Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
  Next ctlPictureBox

  fOK = True
  
TidyUpAndExit:
  ' Disassociate object variales.
  Set ctlPictureBox = Nothing
  TabPages_Resize = fOK
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Private Sub UserControl_Resize()

  With TabStrip1
    .Width = UserControl.Width
    .Height = UserControl.Height
  End With

End Sub
