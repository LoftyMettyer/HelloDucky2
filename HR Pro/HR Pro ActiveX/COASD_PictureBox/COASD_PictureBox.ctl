VERSION 5.00
Begin VB.UserControl COASD_PictureBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picPicture 
      Height          =   975
      Left            =   1125
      ScaleHeight     =   915
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   570
      Width           =   1515
   End
End
Attribute VB_Name = "COASD_PictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
'Const m_def_Picture = ""
Const m_def_PictureLocation = 0
Const m_def_PictureID = 0
Const m_def_MinimumHeight = 0
Const m_def_MinimumWidth = 0
Const m_def_VerticalOffsetBehaviour = 0
Const m_def_HorizontalOffsetBehaviour = 0
Const m_def_VerticalOffset = 0
Const m_def_HorizontalOffset = 0
Const m_def_HeightBehaviour = 0
Const m_def_WidthBehaviour = 0
Const m_def_WFItemType = 0
Const m_def_WFIdentifier = ""
Const m_def_Selected = 0
'Property Variables:
'Dim m_Picture As String
Private gsPicture As String
Dim m_PictureLocation As Long
Dim m_PictureID As Long
Dim m_MinimumHeight As Long
Dim m_MinimumWidth As Long
Dim m_VerticalOffsetBehaviour As Integer
Dim m_HorizontalOffsetBehaviour As Integer
Dim m_VerticalOffset As Integer
Dim m_HorizontalOffset As Integer
Dim m_HeightBehaviour As Integer
Dim m_WidthBehaviour As Integer
Dim m_WFItemType As Integer
Dim m_WFIdentifier As String
Dim m_Selected As Boolean
'Event Declarations:
Event Click() 'MappingInfo=picPicture,picPicture,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=picPicture,picPicture,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=picPicture,picPicture,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=picPicture,picPicture,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=picPicture,picPicture,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picPicture,picPicture,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picPicture,picPicture,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picPicture,picPicture,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPicture,picPicture,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = picPicture.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  picPicture.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPicture,picPicture,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = picPicture.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  picPicture.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPicture,picPicture,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = picPicture.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  picPicture.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPicture,picPicture,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = picPicture.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set picPicture.Font = New_Font
  PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPicture,picPicture,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = picPicture.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  picPicture.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPicture,picPicture,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  picPicture.Refresh
End Sub

Private Sub picPicture_Click()
  RaiseEvent Click
End Sub

Private Sub picPicture_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub picPicture_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picPicture_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picPicture_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=picPicture,picPicture,-1,Picture
'Public Property Get Picture() As Picture
'  Set Picture = picPicture.Picture
'End Property
'
'Public Property Set Picture(ByVal New_Picture As Picture)
'  Set picPicture.Picture = New_Picture
'  PropertyChanged "Picture"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Selected() As Boolean
  Selected = m_Selected
End Property

Public Property Let Selected(ByVal New_Selected As Boolean)
  m_Selected = New_Selected
  PropertyChanged "Selected"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_Selected = m_def_Selected
  m_WFItemType = m_def_WFItemType
  m_WFIdentifier = m_def_WFIdentifier
  m_MinimumHeight = m_def_MinimumHeight
  m_MinimumWidth = m_def_MinimumWidth
  m_VerticalOffsetBehaviour = m_def_VerticalOffsetBehaviour
  m_HorizontalOffsetBehaviour = m_def_HorizontalOffsetBehaviour
  m_VerticalOffset = m_def_VerticalOffset
  m_HorizontalOffset = m_def_HorizontalOffset
  m_HeightBehaviour = m_def_HeightBehaviour
  m_WidthBehaviour = m_def_WidthBehaviour
  m_PictureID = m_def_PictureID
  m_PictureLocation = m_def_PictureLocation
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  picPicture.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  picPicture.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  picPicture.Enabled = PropBag.ReadProperty("Enabled", True)
  Set picPicture.Font = PropBag.ReadProperty("Font", Ambient.Font)
  picPicture.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  m_Selected = PropBag.ReadProperty("Selected", m_def_Selected)
  m_WFItemType = PropBag.ReadProperty("WFItemType", m_def_WFItemType)
  m_WFIdentifier = PropBag.ReadProperty("WFIdentifier", m_def_WFIdentifier)
  m_MinimumHeight = PropBag.ReadProperty("MinimumHeight", m_def_MinimumHeight)
  m_MinimumWidth = PropBag.ReadProperty("MinimumWidth", m_def_MinimumWidth)
  m_VerticalOffsetBehaviour = PropBag.ReadProperty("VerticalOffsetBehaviour", m_def_VerticalOffsetBehaviour)
  m_HorizontalOffsetBehaviour = PropBag.ReadProperty("HorizontalOffsetBehaviour", m_def_HorizontalOffsetBehaviour)
  m_VerticalOffset = PropBag.ReadProperty("VerticalOffset", m_def_VerticalOffset)
  m_HorizontalOffset = PropBag.ReadProperty("HorizontalOffset", m_def_HorizontalOffset)
  m_HeightBehaviour = PropBag.ReadProperty("HeightBehaviour", m_def_HeightBehaviour)
  m_WidthBehaviour = PropBag.ReadProperty("WidthBehaviour", m_def_WidthBehaviour)
  m_PictureID = PropBag.ReadProperty("PictureID", m_def_PictureID)
  m_PictureLocation = PropBag.ReadProperty("PictureLocation", m_def_PictureLocation)
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
  With picPicture
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

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", picPicture.BackColor, &H8000000F)
  Call PropBag.WriteProperty("ForeColor", picPicture.ForeColor, &H80000012)
  Call PropBag.WriteProperty("Enabled", picPicture.Enabled, True)
  Call PropBag.WriteProperty("Font", picPicture.Font, Ambient.Font)
  Call PropBag.WriteProperty("BorderStyle", picPicture.BorderStyle, 1)
  Call PropBag.WriteProperty("Selected", m_Selected, m_def_Selected)
  Call PropBag.WriteProperty("WFItemType", m_WFItemType, m_def_WFItemType)
  Call PropBag.WriteProperty("WFIdentifier", m_WFIdentifier, m_def_WFIdentifier)
  Call PropBag.WriteProperty("MinimumHeight", m_MinimumHeight, m_def_MinimumHeight)
  Call PropBag.WriteProperty("MinimumWidth", m_MinimumWidth, m_def_MinimumWidth)
  Call PropBag.WriteProperty("VerticalOffsetBehaviour", m_VerticalOffsetBehaviour, m_def_VerticalOffsetBehaviour)
  Call PropBag.WriteProperty("HorizontalOffsetBehaviour", m_HorizontalOffsetBehaviour, m_def_HorizontalOffsetBehaviour)
  Call PropBag.WriteProperty("VerticalOffset", m_VerticalOffset, m_def_VerticalOffset)
  Call PropBag.WriteProperty("HorizontalOffset", m_HorizontalOffset, m_def_HorizontalOffset)
  Call PropBag.WriteProperty("HeightBehaviour", m_HeightBehaviour, m_def_HeightBehaviour)
  Call PropBag.WriteProperty("WidthBehaviour", m_WidthBehaviour, m_def_WidthBehaviour)
  Call PropBag.WriteProperty("PictureID", m_PictureID, m_def_PictureID)
  Call PropBag.WriteProperty("PictureLocation", m_PictureLocation, m_def_PictureLocation)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get WFItemType() As Integer
  WFItemType = m_WFItemType
End Property

Public Property Let WFItemType(ByVal New_WFItemType As Integer)
  m_WFItemType = New_WFItemType
  PropertyChanged "WFItemType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get WFIdentifier() As String
  WFIdentifier = m_WFIdentifier
End Property

Public Property Let WFIdentifier(ByVal New_WFIdentifier As String)
  m_WFIdentifier = New_WFIdentifier
  PropertyChanged "WFIdentifier"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinimumHeight() As Long
  MinimumHeight = m_MinimumHeight
End Property

Public Property Let MinimumHeight(ByVal New_MinimumHeight As Long)
  m_MinimumHeight = New_MinimumHeight
  PropertyChanged "MinimumHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinimumWidth() As Long
  MinimumWidth = m_MinimumWidth
End Property

Public Property Let MinimumWidth(ByVal New_MinimumWidth As Long)
  m_MinimumWidth = New_MinimumWidth
  PropertyChanged "MinimumWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get VerticalOffsetBehaviour() As Integer
  VerticalOffsetBehaviour = m_VerticalOffsetBehaviour
End Property

Public Property Let VerticalOffsetBehaviour(ByVal New_VerticalOffsetBehaviour As Integer)
  m_VerticalOffsetBehaviour = New_VerticalOffsetBehaviour
  PropertyChanged "VerticalOffsetBehaviour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get HorizontalOffsetBehaviour() As Integer
  HorizontalOffsetBehaviour = m_HorizontalOffsetBehaviour
End Property

Public Property Let HorizontalOffsetBehaviour(ByVal New_HorizontalOffsetBehaviour As Integer)
  m_HorizontalOffsetBehaviour = New_HorizontalOffsetBehaviour
  PropertyChanged "HorizontalOffsetBehaviour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get VerticalOffset() As Integer
  VerticalOffset = m_VerticalOffset
End Property

Public Property Let VerticalOffset(ByVal New_VerticalOffset As Integer)
  m_VerticalOffset = New_VerticalOffset
  PropertyChanged "VerticalOffset"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get HorizontalOffset() As Integer
  HorizontalOffset = m_HorizontalOffset
End Property

Public Property Let HorizontalOffset(ByVal New_HorizontalOffset As Integer)
  m_HorizontalOffset = New_HorizontalOffset
  PropertyChanged "HorizontalOffset"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get HeightBehaviour() As Integer
  HeightBehaviour = m_HeightBehaviour
End Property

Public Property Let HeightBehaviour(ByVal New_HeightBehaviour As Integer)
  m_HeightBehaviour = New_HeightBehaviour
  PropertyChanged "HeightBehaviour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get WidthBehaviour() As Integer
  WidthBehaviour = m_WidthBehaviour
End Property

Public Property Let WidthBehaviour(ByVal New_WidthBehaviour As Integer)
  m_WidthBehaviour = New_WidthBehaviour
  PropertyChanged "WidthBehaviour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get PictureID() As Long
  PictureID = m_PictureID
End Property

Public Property Let PictureID(ByVal New_PictureID As Long)
  m_PictureID = New_PictureID
  PropertyChanged "PictureID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get PictureLocation() As Long
  PictureLocation = m_PictureLocation
End Property

Public Property Let PictureLocation(ByVal New_PictureLocation As Long)
  m_PictureLocation = New_PictureLocation
  PropertyChanged "PictureLocation"
End Property


Public Property Get Picture() As String
  ' Return the picture property.
  Picture = gsPicture
  
End Property

Public Property Let Picture(ByVal psNewValue As String)
  ' Set the control's picture property.
  On Error GoTo ErrorTrap
  
  ' Update the sub-controls.
  picPicture.Picture = LoadPicture(psNewValue)
  gsPicture = psNewValue
  
  Exit Property
  
ErrorTrap:
  picPicture.Picture = Nothing
  
End Property
