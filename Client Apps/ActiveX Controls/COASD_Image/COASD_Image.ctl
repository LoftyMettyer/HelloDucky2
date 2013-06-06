VERSION 5.00
Begin VB.UserControl COASD_Image 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image imgDefaultImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   2430
      Picture         =   "COASD_Image.ctx":0000
      Stretch         =   -1  'True
      Top             =   2535
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image imgImage 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1320
      Left            =   780
      Picture         =   "COASD_Image.ctx":41EF2
      Stretch         =   -1  'True
      Top             =   615
      Width           =   2250
   End
End
Attribute VB_Name = "COASD_Image"
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
Private gLngPictureID As Long
Private gsPicture As String
Private gLngColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private msWFIdentifier As String
Private miWFItemType As Integer
Private miVOffsetBehave As Integer
Private miHOffsetBehave As Integer
Private miVOffset As Integer
Private miHOffset As Integer
Private miHBehave As Integer
Private miWBehave As Integer

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

Public Property Get BorderStyle() As Integer
  ' Return the control's border style.
  BorderStyle = imgImage.BorderStyle
  
End Property
Public Property Let BorderStyle(piNewValue As Integer)
  ' Set the control's border style.
  If ((piNewValue = vbBSNone) Or (piNewValue = vbFixedSingle)) Then
    
    imgImage.BorderStyle = piNewValue
    
    ' Resize the user control to allow for the new border.
    UserControl_Resize
    
  End If
  
End Property

Private Sub imgDefaultImage_DblClick()
  RaiseEvent DblClick
  
End Sub


Private Sub imgImage_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub imgImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub imgImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub imgImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

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


Private Sub UserControl_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Public Property Get Picture() As String
  ' Return the picture property.
  Picture = gsPicture
  
End Property

Public Property Let Picture(ByVal psNewValue As String)
  ' Set the control's picture property.
  On Error GoTo ErrorTrap
  
  ' Update the sub-controls.
  imgImage.Picture = LoadPicture(psNewValue)
  gsPicture = psNewValue
  
  Exit Property
  
ErrorTrap:
  imgImage.Picture = imgDefaultImage.Picture
  
End Property
Public Property Get PictureID() As Long
  ' Return the picture ID property.
  PictureID = gLngPictureID
  
End Property

Public Property Let PictureID(ByVal pLngNewValue As Long)
  ' Set the control's picture ID property.
  gLngPictureID = pLngNewValue

End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
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
  With imgImage
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

Public Property Get VerticalOffsetBehaviour() As Integer
  VerticalOffsetBehaviour = miVOffsetBehave
End Property

Public Property Let VerticalOffsetBehaviour(ByVal iNewValue As Integer)
  miVOffsetBehave = iNewValue
End Property

Public Property Get HorizontalOffsetBehaviour() As Integer
  HorizontalOffsetBehaviour = miHOffsetBehave
End Property

Public Property Let HorizontalOffsetBehaviour(ByVal iNewValue As Integer)
  miHOffsetBehave = iNewValue
End Property

Public Property Get VerticalOffset() As Integer
  VerticalOffset = miVOffset
End Property

Public Property Let VerticalOffset(ByVal iNewValue As Integer)
  miVOffset = iNewValue
End Property

Public Property Get HorizontalOffset() As Integer
  HorizontalOffset = miHOffset
End Property

Public Property Let HorizontalOffset(ByVal iNewValue As Integer)
  miHOffset = iNewValue
End Property

Public Property Get HeightBehaviour() As Integer
  HeightBehaviour = miHBehave
End Property

Public Property Let HeightBehaviour(ByVal iNewValue As Integer)
  miHBehave = iNewValue
End Property

Public Property Get WidthBehaviour() As Integer
  WidthBehaviour = miWBehave
End Property

Public Property Let WidthBehaviour(ByVal iNewValue As Integer)
  miWBehave = iNewValue
End Property

