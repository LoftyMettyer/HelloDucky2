VERSION 5.00
Begin VB.UserControl COASD_ColourSelector 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   ScaleHeight     =   300
   ScaleWidth      =   1695
   ToolboxBitmap   =   "COASD_ColourSelector.ctx":0000
   Begin VB.PictureBox cboComboBox 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1410
      Picture         =   "COASD_ColourSelector.ctx":0312
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1425
   End
End
Attribute VB_Name = "COASD_ColourSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declare Windows API Functions.
Private Declare Function GetSystemMetricsAPI Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DblClick()

' Constant values.
Const gLngMinHeight = 200
Const gLngMinWidth = 200
Const SM_CXFRAME = 32

Private mlngColumnID As Long
Private gfSelected As Boolean
Private giControlLevel As Boolean
Private mblnReadOnly As Boolean
Private mlngTableID As Boolean
Private gbInScreenDesigner As Boolean

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Let TableID(New_Value As Long)
  mlngTableID = New_Value
End Property
Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Get Selected() As Boolean
  Selected = gfSelected
End Property

Public Property Let Selected(value As Boolean)
  gfSelected = value
End Property

Public Property Get Read_Only() As Boolean
  Read_Only = mblnReadOnly
End Property


Public Property Let Read_Only(blnValue As Boolean)
  mblnReadOnly = blnValue
End Property

Public Property Get ControlLevel() As Integer
  ControlLevel = giControlLevel
End Property

Public Property Let ControlLevel(ByVal piNewValue As Integer)
  giControlLevel = piNewValue
End Property

Public Property Get BackColor() As Long 'OLE_COLOR
  BackColor = lblColour.BackColor
End Property

Public Property Let BackColor(ByVal NewColor As Long) 'OLE_COLOR)
  lblColour.BackColor = NewColor
End Property

Public Property Let Enabled(ByVal value As Boolean)
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  UserControl.Enabled = value
  PropertyChanged "Enabled"
End Property


Private Sub UserControl_InitProperties()
  BackColor = vbWhite
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Enabled = PropBag.ReadProperty("Enabled", True)
  ColumnID = PropBag.ReadProperty("ColumnID", 0)
  BackColor = PropBag.ReadProperty("BackColor", vbWhite)
  Selected = PropBag.ReadProperty("Selected", False)
  ControlLevel = PropBag.ReadProperty("ControlLevel", 0)
  Read_Only = PropBag.ReadProperty("ReadOnly", False)
  TableID = PropBag.ReadProperty("TableID", 0)
End Sub


Private Sub UserControl_Resize()
  cboComboBox.Left = UserControl.Width - cboComboBox.Width
  lblColour.Width = cboComboBox.Left
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("ColumnID", ColumnID, 0)
  Call PropBag.WriteProperty("BackColor", lblColour.BackColor, vbWhite)
  Call PropBag.WriteProperty("Selected", gfSelected, False)
  Call PropBag.WriteProperty("ControlLevel", giControlLevel, 0)
  Call PropBag.WriteProperty("ReadOnly", mblnReadOnly, False)
  Call PropBag.WriteProperty("TableID", mlngTableID, 0)
End Sub

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
End Property

Public Property Let ColumnID(plngNewValue As Long)
  mlngColumnID = plngNewValue
End Property

Private Sub lblColour_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub lblColour_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblColour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblColour_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub cboComboBox_DblClick()
  RaiseEvent DblClick
End Sub

Public Property Get MinimumHeight() As Long
  
  Dim lngMinHeight As Long

  lngMinHeight = UserControl.TextHeight("X") + (Screen.TwipsPerPixelY * 8)

  MinimumHeight = IIf(lngMinHeight < gLngMinHeight, gLngMinHeight, lngMinHeight)

End Property

Public Property Get MinimumWidth() As Long
  
  Dim lngMinWidth As Long
  
  lngMinWidth = (4 * GetSystemMetricsAPI(SM_CXFRAME) * Screen.TwipsPerPixelX) + _
    cboComboBox.Width + _
    UserControl.TextWidth("W")
    
  MinimumWidth = IIf(lngMinWidth < gLngMinWidth, gLngMinWidth, lngMinWidth)
  
End Property

