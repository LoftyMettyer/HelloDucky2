VERSION 5.00
Begin VB.UserControl COAWF_Junction 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   2685
   ScaleWidth      =   2970
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   510
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgPicture 
      Height          =   525
      Index           =   6
      Left            =   480
      Picture         =   "COAWF_Junction.ctx":0000
      Top             =   480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgMask 
      Height          =   525
      Index           =   6
      Left            =   1200
      Picture         =   "COAWF_Junction.ctx":0F08
      Top             =   480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgHighlight 
      Height          =   525
      Index           =   6
      Left            =   1920
      Picture         =   "COAWF_Junction.ctx":1E10
      Top             =   480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgPicture 
      Height          =   525
      Index           =   7
      Left            =   480
      Picture         =   "COAWF_Junction.ctx":2D18
      Top             =   1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgMask 
      Height          =   525
      Index           =   7
      Left            =   1200
      Picture         =   "COAWF_Junction.ctx":3C20
      Top             =   1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgHighlight 
      Height          =   525
      Index           =   7
      Left            =   1920
      Picture         =   "COAWF_Junction.ctx":4B28
      Top             =   1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgPicture 
      Height          =   285
      Index           =   8
      Left            =   720
      Picture         =   "COAWF_Junction.ctx":5A30
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgMask 
      Height          =   285
      Index           =   8
      Left            =   1200
      Picture         =   "COAWF_Junction.ctx":5EE8
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgHighlight 
      Height          =   285
      Index           =   8
      Left            =   1920
      Picture         =   "COAWF_Junction.ctx":63A0
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgPicture 
      Height          =   285
      Index           =   9
      Left            =   720
      Picture         =   "COAWF_Junction.ctx":6858
      Top             =   2040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgMask 
      Height          =   285
      Index           =   9
      Left            =   1200
      Picture         =   "COAWF_Junction.ctx":6D10
      Top             =   2040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgHighlight 
      Height          =   285
      Index           =   9
      Left            =   1920
      Picture         =   "COAWF_Junction.ctx":71C8
      Top             =   2040
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "COAWF_Junction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

' Globals
Private msIdentifier As String
Private miElementType As ElementType
Private mfHighlighted As Boolean
Private miConnectorPairIndex As Integer
Private mavOutboundFlowInfo() As Variant
Private miControlIndex As Integer

' Declare public enums.
Public Enum ElementType
  elem_Begin = 0
  elem_Terminator = 1
  elem_WebForm = 2
  elem_Email = 3
  elem_Decision = 4
  elem_StoredData = 5
  elem_SummingJunction = 6
  elem_Or = 7
  elem_Connector1 = 8
  elem_Connector2 = 9
End Enum

' App Version properties
Private miAppMajor As Integer
Private miAppMinor As Integer
Private miAppRevision As Integer

Public Enum LineDirection
  lineDirection_Down = 0
  lineDirection_Left = 1
  lineDirection_Right = 2
  lineDirection_Up = 3
End Enum

Public Sub About()
  ' Display the 'About' box.
  MsgBox App.ProductName & " - " & App.FileDescription & _
    vbCr & vbCr & App.LegalCopyright, _
    vbOKOnly, "About " & App.ProductName
End Sub

Public Property Get ControlIndex() As Integer
  ControlIndex = miControlIndex
End Property

Public Property Let ControlIndex(ByVal piIndex As Integer)
  miControlIndex = piIndex
End Property

Public Property Get Identifier() As String
  ' Return the Identifier
  Identifier = msIdentifier
End Property

Public Property Let Identifier(ByVal psNewValue As String)
  ' Set the identifier
  msIdentifier = psNewValue
  PropertyChanged "Identifier"
End Property

Private Sub UserControl_Click()
'  Highlighted = True
End Sub

Private Sub UserControl_DblClick()
  ' Pass the DblClick event to the parent form.
  RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the KeyDown event to the parent form.
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

Private Sub UserControl_Resize()
  ResizeElement
End Sub

Private Sub UserControl_InitProperties()
  ' Initialise the properties.
  On Error Resume Next
  
  ElementType = elem_SummingJunction
  Caption = "A"
  Set Font = Ambient.Font
    
  AppMajor = 0
  AppMinor = 0
  AppRevision = 0
  
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Load property values from storage.
  On Error Resume Next

  ' Read the previous set of properties.
  ElementType = PropBag.ReadProperty("ElementType", elem_SummingJunction)
  Highlighted = PropBag.ReadProperty("Highlighted", False)
  Caption = PropBag.ReadProperty("Caption", "A")
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  
  AppMajor = PropBag.ReadProperty("AppMajor", 0)
  AppMinor = PropBag.ReadProperty("AppMinor", 0)
  AppRevision = PropBag.ReadProperty("AppRevision", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
  
  ' Save the current set of properties.
  Call PropBag.WriteProperty("ElementType", miElementType, elem_SummingJunction)
  Call PropBag.WriteProperty("Highlighted", mfHighlighted, False)
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    
  Call PropBag.WriteProperty("AppMajor", miAppMajor, 0)
  Call PropBag.WriteProperty("AppMinor", miAppMinor, 0)
  Call PropBag.WriteProperty("AppRevision", miAppRevision, 0)
End Sub

Public Property Get ElementPicture() As StdPicture
  Set ElementPicture = imgPicture(miElementType).Picture
End Property

Private Sub lblCaption_Click()
  UserControl_Click
End Sub

Private Sub lblCaption_DblClick()
  UserControl_DblClick
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl_MouseUp Button, Shift, X, Y
End Sub

Public Property Get ElementTypeDescription() As String
  ' Return the current element type description.
  Select Case miElementType
    Case elem_Connector1
      ElementTypeDescription = "Connector (part 1)"
    Case elem_Connector2
      ElementTypeDescription = "Connector (part 2)"
    Case elem_Or
      ElementTypeDescription = "Or"
    Case elem_SummingJunction
      ElementTypeDescription = "And"
    Case Else
      ElementTypeDescription = "<Unknown>"
  End Select
  
End Property

Public Property Get Highlighted() As Boolean
  ' Return the 'highlighted' property.
  Highlighted = mfHighlighted
End Property

Public Property Let Highlighted(ByVal pfNewValue As Boolean)
  ' Set the 'highlighted' property.
  mfHighlighted = pfNewValue
  PropertyChanged "Highlighted"
  
  ' Change the picture as required.
  With UserControl
    If mfHighlighted Then
      .Picture = imgHighlight(miElementType).Picture
    Else
      .Picture = imgPicture(miElementType).Picture
    End If
  End With

End Property

Public Property Get Caption() As String
  ' Return the caption
  Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal psNewValue As String)
  ' Set the caption
  lblCaption.Caption = psNewValue
  PropertyChanged "Caption"
  PositionCaption
End Property

Public Property Get CaptionWidth() As Single
  ' Return the caption's width
  CaptionWidth = lblCaption.Width
End Property

Public Property Get CaptionHeight() As Single
  ' Return the caption's height
  CaptionHeight = lblCaption.Height
End Property

Public Property Get CaptionVerticalPosition() As Single
  ' Return the caption's vertical position
  CaptionVerticalPosition = lblCaption.Top
End Property

Public Property Get CaptionHorizontalPosition() As Single
  ' Return the caption's Horizontal position
  CaptionHorizontalPosition = lblCaption.Left
End Property

Public Property Get Font() As Font
  ' Return the caption font.
  Set Font = lblCaption.Font
End Property

Private Sub PositionCaption()
  ' Display the appropriate picture.
  Dim sngSingleLineLength As Single
  
  With lblCaption
    If Len(.Caption) > 0 Then
      ' Disable wordwrap to get the string length in a single line
      .WordWrap = False
      sngSingleLineLength = .Width
      .WordWrap = True
      
      Select Case miElementType
        Case elem_Connector1, elem_Connector2
          If sngSingleLineLength > UserControl.Width Then
            .Width = UserControl.Width
          End If
      
          .Left = (UserControl.Width - .Width) / 2
          .Top = (UserControl.Height - .Height) / 2
      End Select

      .Visible = True
    Else
      .Visible = False
    End If
  End With
  
End Sub

Public Property Get AppMajor() As Integer
  AppMajor = miAppMajor
End Property

Public Property Let AppMajor(ByVal piNewValue As Integer)
  miAppMajor = piNewValue
End Property

Public Property Get AppMinor() As Integer
  AppMinor = miAppMinor
End Property

Public Property Let AppMinor(ByVal piNewValue As Integer)
  miAppMinor = piNewValue
End Property

Public Property Get AppRevision() As Integer
  AppRevision = miAppRevision
End Property

Public Property Let AppRevision(ByVal piNewValue As Integer)
  miAppRevision = piNewValue
End Property

Public Property Get ElementType() As ElementType
  ' Return the current element type.
  ElementType = miElementType
End Property

Public Property Let ElementType(ByVal piNewValue As ElementType)
  ' Set the current type.
  miElementType = piNewValue
  PropertyChanged "ElementType"

  ' Display the appropriate picture.
  DrawElement
  
End Property

Private Sub ResizeElement()
  ' Resize to the current picture.
  With UserControl
    .Height = imgPicture(miElementType).Height
    .Width = imgPicture(miElementType).Width
  End With
End Sub

Private Sub DrawElement()
  ' Display the appropriate picture.
  With UserControl
    If mfHighlighted Then
      .Picture = imgHighlight(miElementType).Picture
    Else
      .Picture = imgPicture(miElementType).Picture
    End If
    .MaskPicture = imgMask(miElementType).Picture
  End With
  
  ResizeElement
  Caption = ""
  
End Sub

Public Function OutboundFlows_Information() As Variant
  ' Return an array defining all required parameters for the outbound flows
  ' of the element.
  
  ' Redimension the array that holds the outbound flow information.
  ' Column 1 = Tag (see enums, or 0 if there's only a single outbound flow)
  ' Column 2 = Direction
  ' Column 3 = XOffset
  ' Column 4 = YOffset
  ' Column 5 = Maximum
  ' Column 6 = Minimum
  ' Column 7 = Description
  ReDim mavOutboundFlowInfo(7, 0)
  
  Select Case miElementType
    Case elem_Connector1
      ' No outbound flows.
      
    Case elem_Connector2
      ReDim mavOutboundFlowInfo(7, 1)
      mavOutboundFlowInfo(1, 1) = 0
      mavOutboundFlowInfo(2, 1) = lineDirection_Down
      mavOutboundFlowInfo(3, 1) = (UserControl.Width / 2)
      mavOutboundFlowInfo(4, 1) = UserControl.Height
      mavOutboundFlowInfo(5, 1) = -1     ' -1 indicates no maximum
      mavOutboundFlowInfo(6, 1) = 1
      mavOutboundFlowInfo(7, 1) = ""
    
    Case elem_Or
      ReDim mavOutboundFlowInfo(7, 1)
      mavOutboundFlowInfo(1, 1) = 0
      mavOutboundFlowInfo(2, 1) = lineDirection_Down
      mavOutboundFlowInfo(3, 1) = (UserControl.Width / 2)
      mavOutboundFlowInfo(4, 1) = UserControl.Height
      mavOutboundFlowInfo(5, 1) = -1     ' -1 indicates no maximum
      mavOutboundFlowInfo(6, 1) = 1
      mavOutboundFlowInfo(7, 1) = ""
    
    Case elem_SummingJunction
      ReDim mavOutboundFlowInfo(7, 1)
      mavOutboundFlowInfo(1, 1) = 0
      mavOutboundFlowInfo(2, 1) = lineDirection_Down
      mavOutboundFlowInfo(3, 1) = (UserControl.Width / 2)
      mavOutboundFlowInfo(4, 1) = UserControl.Height
      mavOutboundFlowInfo(5, 1) = -1     ' -1 indicates no maximum
      mavOutboundFlowInfo(6, 1) = 1
      mavOutboundFlowInfo(7, 1) = ""
    
  End Select
  
  OutboundFlows_Information = mavOutboundFlowInfo
  
End Function

Public Property Get InboundFlow_Direction() As LineDirection
  ' Return the line direction for inbound flows.
  InboundFlow_Direction = lineDirection_Up
End Property

Public Property Get InboundFlows_Maximum() As Integer
  ' Return the maximum number of inbound flows for the element.
  Select Case miElementType
    Case elem_Connector1
      InboundFlows_Maximum = 1
    Case elem_Connector2
      InboundFlows_Maximum = 0
    Case elem_Or
      InboundFlows_Maximum = -1 ' -1 indicates no maximum
    Case elem_SummingJunction
      InboundFlows_Maximum = -1 ' -1 indicates no maximum
    Case Else
      InboundFlows_Maximum = 0
  End Select
End Property

Public Property Get InboundFlows_Minimum() As Integer
  ' Return the minimum number of inbound flows for the element.
  Select Case miElementType
    Case elem_Connector1
      InboundFlows_Minimum = 1
    Case elem_Connector2
      InboundFlows_Minimum = 0
    Case elem_Or
      InboundFlows_Minimum = 2
    Case elem_SummingJunction
      InboundFlows_Minimum = 2
    Case Else
      InboundFlows_Minimum = 0
  End Select
End Property

Public Property Get InboundFlow_XOffset() As Single
  ' Return the XOffset for inbound flows.
  InboundFlow_XOffset = (UserControl.Width / 2)
End Property

Public Property Get InboundFlow_YOffset() As Single
  ' Return the YOffset for inbound flows.
  InboundFlow_YOffset = 0
End Property

Public Property Get ConnectorPairIndex() As Integer
  ' Return the miConnectorPairIndex
  ConnectorPairIndex = miConnectorPairIndex
End Property

Public Property Let ConnectorPairIndex(ByVal piNewValue As Integer)
  ' Set the miConnectorPairIndex
  miConnectorPairIndex = piNewValue
End Property

Public Property Get hWnd() As Variant
  hWnd = UserControl.hWnd
End Property
