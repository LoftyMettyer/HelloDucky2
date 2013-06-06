VERSION 5.00
Begin VB.UserControl COASD_SelectionBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ScaleHeight     =   2055
   ScaleWidth      =   3675
   Begin VB.PictureBox pctVerticalRight 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   3390
      ScaleHeight     =   1530
      ScaleWidth      =   210
      TabIndex        =   3
      Top             =   240
      Width           =   210
      Begin VB.Line linVerticalRight 
         X1              =   60
         X2              =   60
         Y1              =   1455
         Y2              =   135
      End
   End
   Begin VB.PictureBox pctVerticalLeft 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   45
      ScaleHeight     =   1530
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   225
      Width           =   210
      Begin VB.Line linVerticalLeft 
         X1              =   60
         X2              =   60
         Y1              =   1455
         Y2              =   135
      End
   End
   Begin VB.PictureBox pctHorizontalBottom 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   300
      ScaleHeight     =   270
      ScaleWidth      =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   2880
      Begin VB.Line linHorizontalBottom 
         X1              =   45
         X2              =   2700
         Y1              =   105
         Y2              =   105
      End
   End
   Begin VB.PictureBox pctHorizontalTop 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   405
      ScaleHeight     =   270
      ScaleWidth      =   2820
      TabIndex        =   0
      Top             =   90
      Width           =   2820
      Begin VB.Line linHorizontalTop 
         X1              =   45
         X2              =   2700
         Y1              =   105
         Y2              =   105
      End
   End
End
Attribute VB_Name = "COASD_SelectionBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event KeyDown(KeyCode As Integer, Shift As Integer)


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
  BorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
  BorderStyle = PropBag.ReadProperty("BorderStyle", vbBSSolid)

End Sub

Public Property Get hWnd() As Long
  ' Return the control's hWnd.
  hWnd = UserControl.hWnd
  
End Property

Private Sub UserControl_Resize()
  
  ' Top Line
  With pctHorizontalTop
    .Left = 0
    .Top = 0
    .Width = UserControl.Width
    .Height = 15
  End With
  
  With linHorizontalTop
    .X1 = 0
    .X2 = UserControl.Width
    .Y1 = 0
    .Y2 = 0
  End With

  ' Bottom
  With pctHorizontalBottom
    .Left = 0
    .Top = UserControl.Height - 15
    .Width = UserControl.Width
    .Height = 15
  End With
  
  With linHorizontalBottom
    .X1 = 0
    .X2 = UserControl.Width
    .Y1 = 0
    .Y2 = 0
  End With

  ' Left Line
  With pctVerticalLeft
    .Left = 0
    .Top = 0
    .Width = 15
    .Height = UserControl.Height
  End With
  
  With linVerticalLeft
    .X1 = 0
    .X2 = 0
    .Y1 = 0
    .Y2 = UserControl.Height
  End With
  
  ' Right Line
  With pctVerticalRight
    .Left = UserControl.Width - 15
    .Top = 0
    .Width = 15
    .Height = UserControl.Height
  End With
  
  With linVerticalRight
    .X1 = 0
    .X2 = 0
    .Y1 = 0
    .Y2 = UserControl.Height
  End With
  
End Sub



Public Property Get BorderStyle() As Integer
  ' Return the BorderStyle property.
  BorderStyle = linHorizontalTop.BorderStyle
End Property

Public Property Let BorderStyle(ByVal piNewValue As Integer)
  ' Set the BorderStyle property (if it is a valid value).
  If (piNewValue = vbTransparent) Or _
    (piNewValue = vbBSSolid) Or _
    (piNewValue = vbBSDash) Or _
    (piNewValue = vbBSDot) Or _
    (piNewValue = vbBSDashDot) Or _
    (piNewValue = vbBSDashDotDot) Or _
    (piNewValue = vbBSInsideSolid) Then
    
    linHorizontalTop.BorderStyle = piNewValue
    linHorizontalBottom.BorderStyle = piNewValue
    linVerticalLeft.BorderStyle = piNewValue
    linVerticalRight.BorderStyle = piNewValue
  
  End If
  
End Property

Public Property Get BorderColor() As Variant
  ' Return the BorderColor property.
  'BorderColor = shpBox.BorderColor
  BorderColor = linHorizontalTop.BorderColor
  
End Property

Public Property Let BorderColor(ByVal polecolNewValue As Variant)
  
  ' Set the BorderStyle property (if it is a valid value).
  'shpBox.BorderColor = polecolNewValue
  linHorizontalTop.BorderColor = polecolNewValue
  linHorizontalBottom.BorderColor = polecolNewValue
  linVerticalLeft.BorderColor = polecolNewValue
  linVerticalRight.BorderColor = polecolNewValue

End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  
  Call PropBag.WriteProperty("BorderColor", BorderColor, vbBlack)
  Call PropBag.WriteProperty("BorderStyle", BorderStyle, vbBSSolid)

End Sub

Public Sub About()
  ' Display the About information.
  With App
    MsgBox .ProductName & " - " & .FileDescription & _
      vbCr & vbCr & .LegalCopyright, _
      vbOKOnly, "About " & .ProductName
  End With
  
End Sub



