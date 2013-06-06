VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl COASD_OLE 
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ScaleHeight     =   2055
   ScaleWidth      =   3810
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COASD_OLE.ctx":0000
            Key             =   "ENABLED"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COASD_OLE.ctx":0962
            Key             =   "DISABLED"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "(OLE)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   345
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.PictureBox picCaptionContainer 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   2000
      ScaleHeight     =   1155
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   480
      Width           =   1695
      Begin VB.Label lblCaption 
         Caption         =   "Caption"
         Height          =   600
         Left            =   105
         TabIndex        =   1
         Top             =   150
         Width           =   930
      End
   End
End
Attribute VB_Name = "COASD_OLE"
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

' Constant values.
Const gLngMinHeight = 330
Const gLngMinWidth = 330

' Properties.
Private gLngColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private mblnReadOnly As Boolean 'NPG20071022


Public Sub About()
Attribute About.VB_UserMemId = -552
  ' Display the About information.
  With App
    MsgBox .ProductName & " - " & .FileDescription & _
      vbCr & vbCr & .LegalCopyright, _
      vbOKOnly, "About " & .ProductName
  End With
  
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
    
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
    Caption_Refresh
    
End Property

Public Property Let ButtonCaption(sNewCaption As String)

  cmdCommand.Caption = sNewCaption

End Property

Public Property Get ButtonCaption() As String

  ButtonCaption = cmdCommand.Caption
  
End Property

Public Property Get ColumnID() As Long
  ' Return the control's column ID.
  ColumnID = gLngColumnID
  
End Property

Public Property Let ColumnID(ByVal pLngNewValue As Long)
  ' Set the control's column ID.
  gLngColumnID = pLngNewValue
  
End Property
Public Property Get ControlLevel() As Integer
  ' Return the control's level in the z-order.
  ControlLevel = giControlLevel
  
End Property

Public Property Let ControlLevel(ByVal piNewValue As Integer)
  ' Set the control's level in the z-order.
  giControlLevel = piNewValue
  
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
    Set Font = lblCaption.Font
    
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
    Caption_Refresh
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = lblCaption.FontBold
    
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    lblCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    Caption_Refresh
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = lblCaption.FontItalic
    
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
  UserControl.FontItalic() = New_FontItalic
  lblCaption.FontItalic() = New_FontItalic
  PropertyChanged "FontItalic"
  Caption_Refresh
  
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = lblCaption.FontSize
    
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
  UserControl.FontSize() = New_FontSize
  lblCaption.FontSize() = New_FontSize
  PropertyChanged "FontSize"
  Caption_Refresh
  
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = lblCaption.FontStrikethru
    
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
  UserControl.FontStrikethru() = New_FontStrikethru
  lblCaption.FontStrikethru() = New_FontStrikethru
  PropertyChanged "FontStrikethru"
  Caption_Refresh
  
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = lblCaption.FontUnderline
    
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
  UserControl.FontUnderline() = New_FontUnderline
  lblCaption.FontUnderline() = New_FontUnderline
  PropertyChanged "FontUnderline"
  Caption_Refresh
  
End Property


Private Sub cmdCommand_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cmdCommand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub cmdCommand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub cmdCommand_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub picCaptionContainer_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picCaptionContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub picCaptionContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub picCaptionContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Initialize()
  cmdCommand.DisabledPicture = ImageList1.ListImages("ENABLED").picture
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
  'Load property values from storage
  lblCaption.Caption = PropBag.ReadProperty("Caption", "")
  Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
  lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
  lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
  lblCaption.FontSize = PropBag.ReadProperty("FontSize", 8)
  lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
  lblCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)

End Sub


Private Sub UserControl_Resize()
  ' Resize the contained controls as the UserControl is resized.
  Dim lngControlWidth As Long
  Dim lngControlHeight As Long
  Dim lngMinHeight As Long
  Dim lngMinWidth As Long
  Dim lngTextHeight As Long
  Dim lngTextWidth As Long
  Dim lngButtonBorderWidth  As Long

  ' Do not let the user make the control too small.
  lngMinHeight = MinimumHeight
  lngMinWidth = MinimumWidth
  lngButtonBorderWidth = (GetSystemMetricsAPI(SM_CYBORDER) * Screen.TwipsPerPixelY * 2)
  
  With UserControl
    lngControlWidth = .Width
    lngControlHeight = .Height
    
    If .Height < lngMinHeight Then
      .Height = lngMinHeight
    End If
    If .Width < lngMinWidth Then
      .Width = lngMinWidth
    End If
    
    lngTextHeight = .TextHeight(lblCaption.Caption)
    lngTextWidth = .TextWidth(lblCaption.Caption)
  End With
  
  ' Resize the sub-controls.
  With cmdCommand
    .Top = 0
    .Left = 0
    .Height = UserControl.Height
    .Width = UserControl.Width
  End With
  
  With picCaptionContainer
    .Top = lngButtonBorderWidth
    .Left = lngButtonBorderWidth
    .Height = UserControl.Height - (lngButtonBorderWidth * 2)
    .Width = UserControl.Width - (lngButtonBorderWidth * 2)
  End With
  
  Caption_Refresh
  
End Sub

Public Property Get hWnd() As Long
  ' Return the control's hWnd.
  hWnd = UserControl.hWnd
  
End Property

Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  Dim lngMinHeight As Long
  
  MinimumHeight = gLngMinHeight
 
End Property
Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  MinimumWidth = gLngMinWidth
  
End Property

Public Property Get Selected() As Boolean
  ' Return the Selected property.
  Selected = gfSelected
  
End Property

Public Property Let Selected(ByVal pfNewValue As Boolean)
  ' Set the Selected property.
  gfSelected = pfNewValue
    
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
  Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
  Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
  Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
  Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 8)
  Call PropBag.WriteProperty("FontStrikethru", lblCaption.FontStrikethru, 0)
  Call PropBag.WriteProperty("FontUnderline", lblCaption.FontUnderline, 0)

End Sub



Private Sub Caption_Refresh()
  ' Resize the Caption control.
  Dim lngTextHeight As Long
  Dim lngTextWidth As Long
  Dim iLoop As Integer
  Dim iExtraLinesNeeded As Integer
  
  With UserControl
    lngTextHeight = .TextHeight(lblCaption.Caption)
    lngTextWidth = .TextWidth(lblCaption.Caption)
  End With
  
  With lblCaption
    .Height = lngTextHeight
    
    iExtraLinesNeeded = Int(lngTextWidth / picCaptionContainer.Width)
    For iLoop = 1 To iExtraLinesNeeded
      .Height = .Height + lngTextHeight
    Next iLoop
    .Top = (picCaptionContainer.Height - .Height) / 2
    
    .Left = (picCaptionContainer.Width - lngTextWidth) / 2
    .Width = lngTextWidth
    If (iExtraLinesNeeded > 0) Or (.Left < 0) Then
      .Left = 0
      .Width = picCaptionContainer.Width
    End If
  End With
  
End Sub

Public Property Get picture() As String

'  picture = cmdCommand.picture

End Property

Public Property Let picture(sNewPicture As String)

'  If Dir(sNewPicture) <> "" Then
'    cmdCommand.picture = LoadPicture(sNewPicture)
'    cmdCommand.picture = LoadPicture(sNewPicture)
'  End If
  
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Public Property Get Read_Only() As Boolean
'NPG20071022
  Read_Only = mblnReadOnly

End Property


Public Property Let Read_Only(blnValue As Boolean)
'NPG20071022
  mblnReadOnly = blnValue

  'cmdCommand.Enabled = Not blnValue
  'lblCaption.ForeColor = IIf(blnValue, vbButtonFace, vbWindowBackground)
  
  cmdCommand.DisabledPicture = ImageList1.ListImages(IIf(blnValue, "DISABLED", "ENABLED")).picture
  
  
End Property
