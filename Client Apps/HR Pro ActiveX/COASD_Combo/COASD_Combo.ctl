VERSION 5.00
Begin VB.UserControl COASD_Combo 
   BackColor       =   &H80000005&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1875
   LockControls    =   -1  'True
   ScaleHeight     =   975
   ScaleWidth      =   1875
   Begin VB.PictureBox cboComboBox 
      Enabled         =   0   'False
      Height          =   315
      Left            =   960
      Picture         =   "COASD_Combo.ctx":0000
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   200
      Width           =   300
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   195
      Width           =   750
   End
End
Attribute VB_Name = "COASD_Combo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DblClick()

' Declare Windows API Functions.
Private Declare Function GetSystemMetricsAPI Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

' System metric constants.
Public Enum SystemMetrics
  SM_CYCAPTION = 4
  SM_CXBORDER = 5
  SM_CYBORDER = 6
  SM_CXFRAME = 32
  SM_CYFRAME = 33
  SM_CXICON = 11
  SM_CXICONSPACING = 38
  SM_CYICON = 12
  SM_CYICONSPACING = 39
  SM_CYSMCAPTION = 51
  SM_CXVSCROLL = 2
End Enum

Public Enum WorkflowValueTypes
  giWFVALUE_UNKNOWN = -1
  giWFVALUE_FIXED = 0
  giWFVALUE_WFVALUE = 1
  giWFVALUE_DBVALUE = 2
  giWFVALUE_CALC = 3
End Enum

' Constant values.
Const gLngMinHeight = 200
Const gLngMinWidth = 200

' Properties.
Private gLngColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private msWFIdentifier As String
Private miWFItemType As Integer
Private mdtWFDefaultValue As Date
Private msWFDateStringValue As String
Private mlngTableID As String
Private msDefaultStringValue As String
Private msControlValueList As String  'Tab delimited string of the list items in the dropdown
Private mblnReadOnly As Boolean   'NPG20071022

Private mForecolour As OLE_COLOR
Private mBackcolour As OLE_COLOR

Private mlngLookupTableID As Long
Private mlngLookupColumnID As Long
Private mlngLookupFilterColumn As Long
Private miLookupFilterOperator As Integer
Private msLookupFilterValue As String

Private mfMandatory As Boolean

Private mlngCalculationID As Long
Private miDefaultValueType As WorkflowValueTypes

Public Property Get DefaultValueType() As WorkflowValueTypes
  DefaultValueType = miDefaultValueType
  
End Property

Public Property Let DefaultValueType(ByVal piNewValue As WorkflowValueTypes)
  miDefaultValueType = piNewValue
  
End Property

Public Property Get CalculationID() As Long
  CalculationID = mlngCalculationID
  
End Property


Public Property Let CalculationID(ByVal plngNewValue As Long)
  mlngCalculationID = plngNewValue
  
End Property


Public Property Let LookupTableID(New_Value As Long)
  mlngLookupTableID = New_Value
End Property
Public Property Get LookupTableID() As Long
  LookupTableID = mlngLookupTableID
End Property
Public Property Let LookupColumnID(New_Value As Long)
  mlngLookupColumnID = New_Value
End Property
Public Property Get LookupColumnID() As Long
  LookupColumnID = mlngLookupColumnID
End Property


Public Property Let ControlValueList(New_Value As String)
  msControlValueList = New_Value
End Property
Public Property Get ControlValueList() As String
  ControlValueList = msControlValueList
End Property

Public Property Let DefaultStringValue(New_Value As String)
  msDefaultStringValue = New_Value
End Property
Public Property Get DefaultStringValue() As String
  DefaultStringValue = msDefaultStringValue
End Property



Public Property Let TableID(New_Value As Long)
  mlngTableID = New_Value
End Property
Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let WFDefaultValueDateString(New_Value As String)
  msWFDateStringValue = New_Value
End Property
Public Property Get WFDefaultValueDateString() As String
  WFDefaultValueDateString = msWFDateStringValue
End Property

Public Property Let WFDefaultValue(New_Value As Date)
  mdtWFDefaultValue = New_Value
End Property
Public Property Get WFDefaultValue() As Date
  WFDefaultValue = mdtWFDefaultValue
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

Public Property Get Selected() As Boolean
  ' Return the Selected property.
  Selected = gfSelected
  
End Property

Public Property Let Selected(ByVal pfNewValue As Boolean)
  ' Set the Selected property.
  gfSelected = pfNewValue
    
End Property

Public Property Get ControlLevel() As Integer
  ' Return the control's level in the z-order.
  ControlLevel = giControlLevel
  
End Property

Public Property Let ControlLevel(ByVal piNewValue As Integer)
  ' Set the control's level in the z-order.
  giControlLevel = piNewValue
  
End Property


Public Property Get ColumnID() As Long
  ' Return the control's column ID.
  ColumnID = gLngColumnID
  
End Property

Public Property Let ColumnID(ByVal plngNewValue As Long)
  ' Set the control's column ID.
  gLngColumnID = plngNewValue
  
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
  ' Display the About information.
  With App
    MsgBox .ProductName & " - " & .FileDescription & _
      vbCr & vbCr & .LegalCopyright, _
      vbOKOnly, "About " & .ProductName
  End With
  
End Sub

Private Sub cboComboBox_DblClick()
  RaiseEvent DblClick
End Sub


Private Sub lblLabel_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub lblLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub lblLabel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


Private Sub UserControl_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Public Property Get Caption() As String
  ' Return the Caption property.
  Caption = lblLabel.Caption
  
End Property

Public Property Let Caption(ByVal psNewValue As String)
  ' Set the Caption property if it has changed.
  lblLabel.Caption = psNewValue
  
End Property

Public Property Get BackColor() As OLE_COLOR
  ' Return the control's background colour property.
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's background colour property.
  UserControl.BackColor = pColNewColor
  lblLabel.BackColor = pColNewColor
  mBackcolour = pColNewColor
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
  ' Set the control's font property.
  Dim iLoop As Integer
  
  ' Update the sub-controls.
  Set UserControl.Font = pObjNewValue
  Set lblLabel.Font = pObjNewValue
  Set cboComboBox.Font = pObjNewValue
  
  UserControl_Resize
  
End Property

Public Property Get ForeColor() As OLE_COLOR
  ' Return the control's foreground colour property.
  ForeColor = UserControl.ForeColor
  
End Property

Public Property Let ForeColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's foreground colour property.
  UserControl.ForeColor = pColNewColor
  lblLabel.ForeColor = pColNewColor
  mForecolour = pColNewColor
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


Private Sub UserControl_Resize()
  ' Resize the contained controls as the UserControl is resized.
  Dim lngControlWidth As Long
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
  End With
  
  ' Resize the dummy combobox sub-controls.
  With lblLabel
    .Height = lngMinHeight
    .Left = 0
    .Top = 0
    .Width = lngControlWidth
  End With
  
  With cboComboBox
    .Height = lngMinHeight
    .Left = lngControlWidth - .Width
    .Top = 0
  End With

  UserControl.Height = lngMinHeight

End Sub

Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  Dim lngMinHeight As Long
  
  lngMinHeight = UserControl.TextHeight("X") + (Screen.TwipsPerPixelY * 8)
          
  MinimumHeight = IIf(lngMinHeight < gLngMinHeight, gLngMinHeight, lngMinHeight)
  
End Property

Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  Dim lngMinWidth As Long
  
  lngMinWidth = (4 * GetSystemMetricsAPI(SM_CXFRAME) * Screen.TwipsPerPixelX) + _
    cboComboBox.Width + _
    UserControl.TextWidth("W")
    
  MinimumWidth = IIf(lngMinWidth < gLngMinWidth, gLngMinWidth, lngMinWidth)
  
End Property



Public Property Get Mandatory() As Boolean
  Mandatory = mfMandatory

End Property

Public Property Let Mandatory(ByVal pfNewValue As Boolean)
  mfMandatory = pfNewValue
  
End Property

Public Property Get Read_Only() As Boolean
'NPG20071022
  Read_Only = mblnReadOnly

End Property


Public Property Let Read_Only(blnValue As Boolean)
'NPG20071022
  mblnReadOnly = blnValue

'  AE20080424 Fault #13124
'  With lblLabel
'    .BackColor = IIf(blnValue, vbButtonFace, vbWindowBackground)
'    .ForeColor = IIf(blnValue, vbGrayText, vbWindowText)
'  End With

  With lblLabel
    .BackColor = IIf(blnValue, vbButtonFace, mBackcolour)
    .ForeColor = IIf(blnValue, vbGrayText, mForecolour)
  End With
  
End Property

Public Property Get LookupFilterColumn() As Long
  LookupFilterColumn = mlngLookupFilterColumn
End Property

Public Property Let LookupFilterColumn(ByVal plngNewValue As Long)
  mlngLookupFilterColumn = plngNewValue
End Property

Public Property Get LookupFilterOperator() As Integer
  LookupFilterOperator = miLookupFilterOperator

End Property

Public Property Let LookupFilterOperator(ByVal piNewValue As Integer)
  miLookupFilterOperator = piNewValue
  
End Property

Public Property Get LookupFilterValue() As String
  LookupFilterValue = msLookupFilterValue

End Property

Public Property Let LookupFilterValue(ByVal psNewValue As String)
  msLookupFilterValue = psNewValue
  
End Property
