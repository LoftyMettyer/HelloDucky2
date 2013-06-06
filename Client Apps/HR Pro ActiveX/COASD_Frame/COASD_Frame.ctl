VERSION 5.00
Begin VB.UserControl COASD_Frame 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Frame fraFrame 
      Caption         =   "Frame1"
      Height          =   1290
      Left            =   675
      TabIndex        =   0
      Top             =   660
      Width           =   2010
   End
End
Attribute VB_Name = "COASD_Frame"
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
Const gLngMinHeight = 180
Const gLngMinWidth = 180

' Properties.
Private gLngColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private msWFIdentifier As String
Private miWFItemType As Integer

Public Enum ASRBackStyleConstants
  BACKSTYLE_TRANSPARENT = 0
  BACKSTYLE_OPAQUE = 1
End Enum

Private miBackStyle As ASRBackStyleConstants
Public Property Get BackStyle() As ASRBackStyleConstants
  BackStyle = miBackStyle

End Property


Public Property Let BackStyle(ByVal New_BackStyle As ASRBackStyleConstants)
  ' NB. This property is NOT applied to the UserControl and contained control(s)
  miBackStyle = New_BackStyle
  PropertyChanged "BackStyle"
  
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

Public Sub About()
Attribute About.VB_UserMemId = -552
  ' Display the About information.
  With App
    MsgBox .ProductName & " - " & .FileDescription & _
      vbCr & vbCr & .LegalCopyright, _
      vbOKOnly, "About " & .ProductName
  End With
  
End Sub

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

Public Property Let ColumnID(ByVal pLngNewValue As Long)
  ' Set the control's column ID.
  gLngColumnID = pLngNewValue
  
End Property
Public Property Get hWnd() As Long
  ' Return the control's hWnd.
  hWnd = UserControl.hWnd
  
End Property

Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  MinimumHeight = gLngMinHeight
  
End Property

Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  MinimumWidth = gLngMinWidth
  
End Property

Private Sub fraFrame_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub fraFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub fraFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub fraFrame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
  With fraFrame
    .Top = 0
    .Left = 0
    .Height = lngControlHeight
    .Width = lngControlWidth
  End With
  
  With UserControl
    If .Height < lngMinHeight Then
      .Height = lngMinHeight
    End If
  End With
  
End Sub

Public Property Get Caption() As String
  ' Return the Caption property.
  Caption = fraFrame.Caption
  
End Property

Public Property Get BackColor() As OLE_COLOR
  ' Return the control's background colour property.
  BackColor = UserControl.BackColor
  
End Property

Public Property Get ForeColor() As OLE_COLOR
  ' Return the control's foreground colour property.
  ForeColor = UserControl.ForeColor
  
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
  Set fraFrame.Font = pObjNewValue
  
  UserControl_Resize
  
End Property

Public Property Let ForeColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's foreground colour property.
  UserControl.ForeColor = pColNewColor
  fraFrame.ForeColor = pColNewColor
  
End Property

Public Property Let BackColor(ByVal pColNewColor As OLE_COLOR)
  ' Set the control's background colour property.
  UserControl.BackColor = pColNewColor
  fraFrame.BackColor = pColNewColor

End Property

Public Property Let Caption(ByVal psNewValue As String)
  ' Set the Caption property if it has changed.
  fraFrame.Caption = psNewValue
  
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackStyle", miBackStyle, BACKSTYLE_OPAQUE)

End Sub


