VERSION 5.00
Begin VB.UserControl COASD_Checkbox 
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2085
   ScaleHeight     =   930
   ScaleWidth      =   2085
   Begin VB.PictureBox picCheckBox 
      BackColor       =   &H80000005&
      Height          =   195
      Left            =   240
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   120
      Width           =   195
   End
   Begin VB.Image imgTicked 
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      MousePointer    =   1  'Arrow
      Picture         =   "COASD_CheckBox.ctx":0000
      Top             =   360
      Width           =   195
   End
   Begin VB.Image imgCheckBox 
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      MousePointer    =   1  'Arrow
      Picture         =   "COASD_CheckBox.ctx":024A
      Top             =   600
      Width           =   195
   End
   Begin VB.Label lblLabel 
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "COASD_Checkbox"
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

' Constant values.
Const gLngMinHeight = 195
Const gLngMinWidth = 200
Const gLngCheckBoxOffset = 360

' Properties.
Private giAlignment As Integer
Private gLngColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private msWFIdentifier As String
Private miWFItemType As Integer
Private mbWFDefaultValue As Boolean
Private mblnReadOnly As Boolean 'NPG20071022

Private mForecolour As OLE_COLOR
Private mBackcolour As OLE_COLOR

Public Enum ASRBackStyleConstants
  BACKSTYLE_TRANSPARENT = 0
  BACKSTYLE_OPAQUE = 1
End Enum

Public Enum WorkflowValueTypes
  giWFVALUE_UNKNOWN = -1
  giWFVALUE_FIXED = 0
  giWFVALUE_WFVALUE = 1
  giWFVALUE_DBVALUE = 2
  giWFVALUE_CALC = 3
End Enum

Private miBackStyle As ASRBackStyleConstants
Private mlngCalculationID As Long
Private miDefaultValueType As WorkflowValueTypes

Public Property Get CalculationID() As Long
  CalculationID = mlngCalculationID
  
End Property

Public Property Let CalculationID(ByVal plngNewValue As Long)
  mlngCalculationID = plngNewValue
  
End Property

Public Property Get DefaultValueType() As WorkflowValueTypes
  DefaultValueType = miDefaultValueType
  
End Property

Public Property Let DefaultValueType(ByVal piNewValue As WorkflowValueTypes)
  miDefaultValueType = piNewValue
  
End Property

Public Property Get BackStyle() As ASRBackStyleConstants
  BackStyle = miBackStyle

End Property

Public Property Let BackStyle(ByVal New_BackStyle As ASRBackStyleConstants)
  ' NB. This property is NOT applied to the UserControl and contained control(s)
  ' See MSDN Q179052 for the reason.
  miBackStyle = New_BackStyle
  PropertyChanged "BackStyle"
  
End Property



Public Property Let WFDefaultValue(New_Value As Boolean)
  mbWFDefaultValue = New_Value

  If mbWFDefaultValue Then
    picCheckBox.Visible = False
    imgCheckBox.Visible = False
    imgTicked.Visible = True
  Else
    picCheckBox.Visible = True
    imgCheckBox.Visible = True
    imgTicked.Visible = False
  End If

End Property
Public Property Get WFDefaultValue() As Boolean
  WFDefaultValue = mbWFDefaultValue
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
Public Property Get Caption() As String
  ' Return the Caption property.
  Caption = lblLabel.Caption
  
End Property

Public Property Let Caption(ByVal psNewValue As String)
  ' Set the Caption property if it has changed.
  lblLabel.Caption = psNewValue
  
  UserControl_Resize
  
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

Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  Dim lngMinHeight As Long
  
  lngMinHeight = UserControl.TextHeight("X")
          
  MinimumHeight = IIf(lngMinHeight < gLngMinHeight, gLngMinHeight, lngMinHeight)
  
End Property


Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  Dim lngMinWidth As Long
  
  lngMinWidth = gLngCheckBoxOffset + UserControl.TextWidth("W")
    
  MinimumWidth = IIf(lngMinWidth < gLngMinWidth, gLngMinWidth, lngMinWidth)
  
End Property
Public Property Get Alignment() As Integer
  ' Return the Alignment property.
  Alignment = giAlignment
  
End Property

Public Property Let Alignment(ByVal piNewValue As Integer)
  ' Set the Alignment property if it is valid.
  
  If (piNewValue = vbLeftJustify) Or (piNewValue = vbRightJustify) Then
        
    giAlignment = piNewValue
    
    ' Redraw the control with the new alignment.
    UserControl_Resize
      
  End If
  
End Property

Private Sub imgCheckBox_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub imgCheckBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub imgCheckBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub imgCheckBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub imgTicked_DblClick()
  RaiseEvent DblClick

End Sub


Private Sub lblLabel_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub picCheckBox_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub picCheckBox_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'##
Private Sub PicCheckBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub picCheckBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub picCheckBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

'##

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

Private Sub UserControl_Initialize()
  giAlignment = vbLeftJustify
  
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
  miBackStyle = PropBag.ReadProperty("BackStyle", BACKSTYLE_OPAQUE)

End Sub

Private Sub UserControl_Resize()
  ' Resize the contained controls as the UserControl is resized.
  Dim lngControlHeight As Long
  Dim lngControlWidth As Long
  Dim lngMinHeight As Long
  Dim lngMinWidth As Long
  Dim lngTextHeight As Long
  Dim lngTextWidth As Long
  Dim iExtraLinesNeeded As Integer
  Dim iLoop As Integer
  
  lngMinHeight = MinimumHeight
  lngMinWidth = MinimumWidth
  
  If mbWFDefaultValue Then
    picCheckBox.Visible = False
    imgCheckBox.Visible = False
    imgTicked.Visible = True
  Else
    picCheckBox.Visible = True
    imgCheckBox.Visible = True
    imgTicked.Visible = False
  End If
  
  ' Do not let the user make the control too small.
  With UserControl
    If .Height < lngMinHeight Then
      .Height = lngMinHeight
    End If

    If .Width < lngMinWidth Then
      .Width = lngMinWidth
    End If
      
    lngControlHeight = .Height
    lngControlWidth = .Width
  
    lngTextHeight = .TextHeight(lblLabel.Caption)
    lngTextWidth = .TextWidth(lblLabel.Caption)
  End With
  
  ' Resize the checkbox control.

  ' Calculate the dimensions of the label.
  With lblLabel

    .Height = lngTextHeight
    .Width = lngControlWidth - gLngCheckBoxOffset

    iExtraLinesNeeded = Int(lngTextWidth / .Width)
    For iLoop = 1 To iExtraLinesNeeded
      .Height = .Height + lngTextHeight
    Next iLoop
        
    If giAlignment = vbLeftJustify Then
      ' Left aligned.
      .Left = gLngCheckBoxOffset
      imgCheckBox.Left = 0
      imgTicked.Left = 0
      picCheckBox.Left = 0
      
    Else
      ' Right aligned.
      .Left = 0
      imgCheckBox.Left = UserControl.Width - imgCheckBox.Width
      imgTicked.Left = UserControl.Width - imgCheckBox.Width
      picCheckBox.Left = UserControl.Width - picCheckBox.Width
    
    End If
        
    .Top = (lngControlHeight - .Height) / 2
    imgCheckBox.Top = (lngControlHeight - imgCheckBox.Height) / 2
    imgTicked.Top = (lngControlHeight - imgCheckBox.Height) / 2
    picCheckBox.Top = (lngControlHeight - picCheckBox.Height) / 2
      
  End With
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  
  Call PropBag.WriteProperty("BackStyle", miBackStyle, BACKSTYLE_OPAQUE)

End Sub

Public Property Get Read_Only() As Boolean
'NPG20071022
  Read_Only = mblnReadOnly

End Property

Public Property Let Read_Only(blnValue As Boolean)
'NPG20071022
  mblnReadOnly = blnValue

  With imgCheckBox
    .Enabled = Not blnValue
  End With
  
  With imgTicked
    .Enabled = Not blnValue
  End With

'  JPD 20090102 Fault 13335
'  AE20080424 Fault #13124
  With picCheckBox
    .BackColor = IIf(blnValue, vbButtonFace, vbWindowBackground)
    .ForeColor = IIf(blnValue, vbGrayText, vbWindowText)
  End With
  
  With lblLabel
    .Enabled = Not blnValue
  End With
    
End Property

