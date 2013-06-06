VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl COAWF_TabPage 
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1965
   ScaleHeight     =   1545
   ScaleWidth      =   1965
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   870
      Left            =   225
      TabIndex        =   0
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
End
Attribute VB_Name = "COAWF_TabPage"
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
Private miWFItemType As Integer
Private miCurrentTab As Integer

Private Const Xframe = 20
Private Const YFrame = 20

Public ControlLevel As Long
Public MinimumWidth As Long
Public MinimumHeight As Long

Public Property Let WFItemType(New_Value As Integer)
  miWFItemType = New_Value
End Property

Public Property Get WFItemType() As Integer
  WFItemType = 21
End Property

Public Property Get ClientLeft() As Long
  ClientLeft = TabStrip1.ClientLeft
End Property

Public Property Get ClientTop() As Long
  ClientTop = TabStrip1.ClientTop
End Property

Public Property Get ClientHeight() As Long
  ClientHeight = TabStrip1.ClientHeight
End Property

Public Property Get ClientWidth() As Long
  ClientWidth = TabStrip1.ClientWidth
End Property

' Return the control's hWnd.
Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Sub AddTabPage(ByVal Caption As String)
  
  Dim iContainerIndex As Long
  Dim objAddedTab As MSComctlLib.Tab
  
  iContainerIndex = TabStrip1.Tabs.Count + 1
  
  Set objAddedTab = TabStrip1.Tabs.Add
  objAddedTab.Caption = Caption
   
  TabPages_Resize
  
End Sub

Public Property Get TabPage(ByVal PageNo As Long) As MSComctlLib.Tab
  Set TabPage = TabStrip1.Tabs(PageNo)
End Property

Public Property Get ControlPage(ByVal PageNo As Long) As StdPicture
  'Set ControlPage = objContainer(PageNo)
End Property

Public Function Tabs() As Tabs
  Set Tabs = TabStrip1.Tabs
End Function

Private Sub TabStrip1_Click()
  
  If TabStrip1.SelectedItem.Index <> miCurrentTab Then
    RaiseEvent Click
  End If

  miCurrentTab = TabStrip1.SelectedItem.Index

End Sub

Private Sub TabStrip1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
'  If TabStrip1.SelectedItem.Index <> miCurrentTab Then
'    RaiseEvent Click
'  Else
    RaiseEvent MouseDown(Button, Shift, x, y)
'  End If
'
'  miCurrentTab = TabStrip1.SelectedItem.Index
  
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
  
  ' Position and size the tabstrip to fill the form's client area.
  TabStrip1.Move Xframe, YFrame, UserControl.ScaleWidth - (Xframe * 2), UserControl.ScaleHeight - (YFrame * 2)
  
  fOK = True
  
TidyUpAndExit:
  TabPages_Resize = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub UserControl_Resize()

  With TabStrip1
    .Width = UserControl.Width
    .Height = UserControl.Height
  End With

End Sub

Public Function GetCaptions() As String
  
  Dim objTab As MSComctlLib.Tab
  Dim sCaptions As String
  
  For Each objTab In TabStrip1.Tabs
    sCaptions = sCaptions + objTab.Caption + ";"
  Next

  GetCaptions = sCaptions

End Function

Public Property Get Caption() As String
  Caption = GetCaptions
End Property

Public Property Let Caption(ByVal NewValue As String)

  Dim objTab As MSComctlLib.Tab
  Dim sCaption As String
  Dim aryCaptions() As String
  Dim lngCount As Long
  
  On Error GoTo ErrorTrap
  
  aryCaptions = Split(NewValue, ";")
  
  For lngCount = LBound(aryCaptions) To UBound(aryCaptions)
    TabStrip1.Tabs(lngCount + 1).Caption = aryCaptions(lngCount)
  Next

ErrorTrap:

End Property


