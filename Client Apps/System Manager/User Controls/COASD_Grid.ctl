VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.UserControl COASD_Grid 
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   4800
   Begin SSDataWidgets_B.SSDBGrid grdGrid 
      Height          =   3045
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4365
      ScrollBars      =   0
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "ssetHighlight"
      stylesets(0).ForeColor=   -2147483634
      stylesets(0).BackColor=   -2147483635
      stylesets(0).Picture=   "COASD_Grid.ctx":0000
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      BalloonHelp     =   0   'False
      RowNavigation   =   3
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).Caption=   "Column Header"
      Columns(0).Name =   "ColumnHeader"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      TabNavigation   =   1
      _ExtentX        =   7699
      _ExtentY        =   5371
      _StockProps     =   79
      Caption         =   "Caption"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
Attribute VB_Name = "COASD_Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DblClick()

' Properties.
Private mlngColumnID As Long
Private mlngTableID As Long
Private miControlLevel As Integer
Private mfSelected As Boolean
Private msWFIdentifier As String
Private miWFItemType As Integer
Private mlngWFDatabaseRecord As Long
Private msWFWorkflowForm As String
Private msWFWorkflowValue As String
Private mlngWFRecordTableID As Long
Private mlngWFRecordOrderID As Long
Private mlngWFRecordFilterID As Long
Private mfMandatory As Boolean

Public UseAsTargetIdentifier As Boolean

Public Property Let WFWorkflowForm(New_Value As String)
  msWFWorkflowForm = New_Value
End Property
Public Property Get WFWorkflowForm() As String
  WFWorkflowForm = msWFWorkflowForm
End Property
Public Property Let WFWorkflowValue(New_Value As String)
  msWFWorkflowValue = New_Value
End Property
Public Property Get WFWorkflowValue() As String
  WFWorkflowValue = msWFWorkflowValue
End Property

Public Property Get HeadLines() As Integer
  HeadLines = grdGrid.HeadLines
End Property
Public Property Let HeadLines(ByVal New_Value As Integer)
  grdGrid.HeadLines = New_Value
End Property

Public Property Get Caption() As String
  Caption = grdGrid.Caption
End Property
Public Property Let Caption(ByVal New_Value As String)
  grdGrid.Caption = New_Value
  UserControl_Resize
End Property

Public Property Let WFDatabaseRecord(New_Value As Long)
  mlngWFDatabaseRecord = New_Value
End Property
Public Property Get WFDatabaseRecord() As Long
  WFDatabaseRecord = mlngWFDatabaseRecord
End Property

Public Property Get BackColorHighlight() As OLE_COLOR
  ' Return the grid control's background colour for highlighted rows.
  BackColorHighlight = grdGrid.StyleSets("ssetHighlight").BackColor
  
End Property
Public Property Get ForeColorHighlight() As OLE_COLOR
  ' Return the grid control's foreground colour for highlighted rows.
  ForeColorHighlight = grdGrid.StyleSets("ssetHighlight").ForeColor
  
End Property

Public Property Get BackColorEven() As OLE_COLOR
  ' Return the grid control's row colour for even rows.
  BackColorEven = grdGrid.BackColorEven
End Property

Public Property Let BackColorHighlight(ByVal pColNewColor As OLE_COLOR)
  ' Set the grid control's background colour for highlighted rows.
  grdGrid.StyleSets("ssetHighlight").BackColor = pColNewColor
  
End Property

Public Property Let ForeColorHighlight(ByVal pColNewColor As OLE_COLOR)
  ' Set the grid control's foreground colour for highlighted rows.
  grdGrid.StyleSets("ssetHighlight").ForeColor = pColNewColor
  
End Property


Public Property Let BackColorEven(ByVal pColNewColor As OLE_COLOR)
  ' Set the grid control's row colour for even rows.
  grdGrid.BackColorEven = pColNewColor
End Property


Public Property Get BackColorOdd() As OLE_COLOR
  ' Return the grid control's row colour for odd rows.
  BackColorOdd = grdGrid.BackColorOdd
End Property
Public Property Let BackColorOdd(ByVal pColNewColor As OLE_COLOR)
  ' Set the grid control's row colour for odd rows.
  grdGrid.BackColorOdd = pColNewColor
End Property

Public Property Get ForeColorEven() As OLE_COLOR
  ' Return the grid control's row colour for even rows.
  ForeColorEven = grdGrid.ForeColorEven
End Property
Public Property Let ForeColorEven(ByVal pColNewColor As OLE_COLOR)
  ' Set the grid control's row colour for even rows.
  grdGrid.ForeColorEven = pColNewColor
End Property

Public Property Get ForeColorOdd() As OLE_COLOR
  ' Return the grid control's row colour for odd rows.
  ForeColorOdd = grdGrid.ForeColorOdd
End Property
Public Property Let ForeColorOdd(ByVal pColNewColor As OLE_COLOR)
  ' Set the grid control's row colour for odd rows.
  grdGrid.ForeColorOdd = pColNewColor
End Property

Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  MinimumHeight = grdGrid.RowHeight
End Property

Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  MinimumWidth = grdGrid.RowHeight
End Property

Private Sub grdGrid_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub UserControl_Initialize()
  Caption = ""
  grdGrid.AddItem "This is an even line"
  grdGrid.AddItem "This is an odd line"
  
End Sub

Public Property Get TableID() As Long
  ' Return the control's table ID.
  TableID = mlngColumnID
End Property
Public Property Let TableID(ByVal New_Value As Long)
  ' Set the control's table ID.
  mlngColumnID = New_Value
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub grdGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub grdGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub grdGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get ColumnHeaders() As Boolean
  ' Return the grid ColumnHeaders property.
  ColumnHeaders = grdGrid.ColumnHeaders
End Property
Public Property Let ColumnHeaders(ByVal New_Value As Boolean)
  ' Set the grid ColumnHeaders property.
  grdGrid.ColumnHeaders = New_Value
End Property

Public Property Get Selected() As Boolean
  ' Return the Selected property.
  Selected = mfSelected
End Property
Public Property Let Selected(ByVal New_Value As Boolean)
  ' Set the Selected property.
  mfSelected = New_Value
End Property

Public Property Get ControlLevel() As Integer
  ' Return the control's level in the z-order.
  ControlLevel = miControlLevel
End Property
Public Property Let ControlLevel(ByVal New_Value As Integer)
  ' Set the control's level in the z-order.
  miControlLevel = New_Value
End Property

Public Property Get ColumnID() As Long
  ' Return the control's column ID.
  ColumnID = mlngColumnID
End Property
Public Property Let ColumnID(ByVal New_Value As Long)
  ' Set the control's column ID.
  mlngColumnID = New_Value
End Property

Public Property Let WFIdentifier(New_Value As String)
  msWFIdentifier = New_Value
End Property
Public Property Get WFIdentifier() As String
  WFIdentifier = msWFIdentifier
End Property

Public Property Let WFItemType(New_Value As Integer)
  miWFItemType = New_Value
End Property
Public Property Get WFItemType() As Integer
  WFItemType = miWFItemType
End Property

Public Property Get BackColor() As OLE_COLOR
  ' Return the control's background colour property.
  BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's background colour property.
  UserControl.BackColor = pColNewColor
  grdGrid.BackColor = pColNewColor
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
  ' Return the control's header back colour property.
  HeaderBackColor = grdGrid.BevelColorFace
End Property
Public Property Let HeaderBackColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's header back colour property.
  grdGrid.BevelColorFace = pColNewColor
  grdGrid.BevelColorFrame = vbBlack
  grdGrid.BevelColorHighlight = pColNewColor
  grdGrid.BevelColorShadow = pColNewColor
End Property

Public Property Get hWnd() As Long
  ' Return the control's hWnd.
  hWnd = UserControl.hWnd
End Property

Public Property Get Font() As Font
  ' Return the control's font property.
  Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal pObjNewValue As StdFont)
  
  ' Update the sub-controls.
  Set UserControl.Font = pObjNewValue
  Set grdGrid.Font = pObjNewValue
  
  UserControl_Resize
 
End Property

Public Property Get HeadFont() As Font
  ' Return the grid's headfont property.
  Set HeadFont = grdGrid.HeadFont
End Property
Public Property Set HeadFont(ByVal New_Value As StdFont)
  
  ' Update the grid's headfont property.
  Set grdGrid.HeadFont = New_Value
  
  UserControl_Resize
  
End Property

Public Property Get ForeColor() As OLE_COLOR
  ' Return the control's foreground colour property.
  ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's foreground colour property.
  UserControl.ForeColor = pColNewColor
  grdGrid.ForeColor = pColNewColor
End Property

Public Sub About()
  ' Display the About information.
  With App
    MsgBox .ProductName & " - " & .FileDescription & _
      vbCr & vbCr & .LegalCopyright, _
      vbOKOnly, "About " & .ProductName
  End With
End Sub

Private Sub UserControl_Resize()
  ' Resize the contained controls as the UserControl is resized.
  Dim lngControlWidth As Long
  Dim lngControlHeight As Long
  Dim lngMinHeight As Long
  Dim lngMinWidth As Long
  
  ' Do not let the user make the control too small.
  lngMinHeight = MinimumHeight
  lngMinWidth = MinimumWidth
  
  With UserControl
    If .Width < lngMinWidth Then
      .Width = lngMinWidth
    End If
    lngControlWidth = .Width
    
    lngControlHeight = .Height
  End With
  
  ' Resize the dummy spinner sub-controls.
  With grdGrid
    .Top = 0
    .Left = 0
    .Height = lngControlHeight
    .Columns(0).Width = lngControlWidth
    .Width = lngControlWidth
  End With
  
  With UserControl
    If .Height < lngMinHeight Then
      .Height = lngMinHeight
    End If
  End With

End Sub

Public Property Get WFRecordTableID() As Long
  WFRecordTableID = mlngWFRecordTableID

End Property

Public Property Let WFRecordTableID(ByVal plngNewValue As Long)
  mlngWFRecordTableID = plngNewValue
  
End Property

Public Property Get WFRecordOrderID() As Long
  WFRecordOrderID = mlngWFRecordOrderID

End Property

Public Property Let WFRecordOrderID(ByVal plngNewValue As Long)
  mlngWFRecordOrderID = plngNewValue

End Property

Public Property Get WFRecordFilterID() As Long
  WFRecordFilterID = mlngWFRecordFilterID

End Property

Public Property Let WFRecordFilterID(ByVal plngNewValue As Long)
  mlngWFRecordFilterID = plngNewValue

End Property

Public Property Get Mandatory() As Boolean
  Mandatory = mfMandatory
  
End Property

Public Property Let Mandatory(ByVal pfNewValue As Boolean)
  mfMandatory = pfNewValue
  
End Property
