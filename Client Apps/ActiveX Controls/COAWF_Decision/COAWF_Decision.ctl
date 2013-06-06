VERSION 5.00
Begin VB.UserControl COAWF_Decision 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "COAWF_Decision.ctx":0000
   ScaleHeight     =   2175
   ScaleWidth      =   4920
   Begin VB.Image imgHighlight 
      Height          =   1230
      Left            =   3360
      Picture         =   "COAWF_Decision.ctx":6434
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgMask 
      Height          =   1230
      Left            =   1680
      Picture         =   "COAWF_Decision.ctx":C868
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgPicture 
      Height          =   1230
      Left            =   0
      Picture         =   "COAWF_Decision.ctx":12C9C
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblDecisionCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label lblDecisionCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   120
   End
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
      Left            =   2220
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   510
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "COAWF_Decision"
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

' App Version properties
Private miAppMajor As Integer
Private miAppMinor As Integer
Private miAppRevision As Integer

' Globals
Private msIdentifier As String
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

Public Enum LineDirection
  lineDirection_Down = 0
  lineDirection_Left = 1
  lineDirection_Right = 2
  lineDirection_Up = 3
End Enum

Public Enum DecisionFlowTypes
  decisionFlowType_Button = 0
  decisionFlowType_Expression = 1
End Enum

Public Enum DecisionOutboundFlows
  decisionOutFlow_False = 0
  decisionOutFlow_True = 1
End Enum

Public Enum DecisionCaptionType
  decisionCaption_T_F = 0
  decisionCaption_Y_N = 1
  decisionCaption_1_0 = 2
  decisionCaption_tick_cross = 3
End Enum

' Decision specific properties
Private miDecisionCaptionType As DecisionCaptionType
Private miDecisionFlowType As DecisionFlowTypes
Private msTrueFlowIdentifier As String
Private mlngDecisionFlowExprID As Long

Public Sub About()
  ' Display the 'About' box.
  MsgBox App.ProductName & " - " & App.FileDescription & _
    vbCr & vbCr & App.LegalCopyright, _
    vbOKOnly, "About " & App.ProductName
End Sub

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

Private Sub UserControl_Initialize()
  StaticCaptionInitialize
  DrawElement
End Sub

Private Sub UserControl_InitProperties()
  ' Initialise the properties.
  On Error Resume Next
  
  Caption = "Decision"
  Identifier = ""
  Set Font = Ambient.Font
  
  DecisionCaptionType = decisionCaption_T_F
  DecisionFlowType = decisionFlowType_Button
  TrueFlowIdentifier = ""
  DecisionFlowExpressionID = 0
  
  AppMajor = 0
  AppMinor = 0
  AppRevision = 0
  
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Load property values from storage.
  On Error Resume Next

  ' Read the previous set of properties.
  Caption = PropBag.ReadProperty("Caption", "Decision")
  Identifier = PropBag.ReadProperty("Identifier", "")
  Highlighted = PropBag.ReadProperty("Highlighted", False)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  
  DecisionCaptionType = PropBag.ReadProperty("DecisionCaptionType", decisionCaption_T_F)
  DecisionFlowType = PropBag.ReadProperty("DecisionFlowType", decisionFlowType_Button)
  TrueFlowIdentifier = PropBag.ReadProperty("TrueFlowIdentifier", "")
  DecisionFlowExpressionID = PropBag.ReadProperty("DecisionFlowExpressionID", 0)

  AppMajor = PropBag.ReadProperty("AppMajor", 0)
  AppMinor = PropBag.ReadProperty("AppMinor", 0)
  AppRevision = PropBag.ReadProperty("AppRevision", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
  
  ' Save the current set of properties.
  Call PropBag.WriteProperty("Highlighted", mfHighlighted, False)
  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
  Call PropBag.WriteProperty("Identifier", msIdentifier, "")
  
  Call PropBag.WriteProperty("DecisionCaptionType", miDecisionCaptionType, decisionCaption_T_F)
  Call PropBag.WriteProperty("DecisionFlowType", miDecisionFlowType, decisionFlowType_Button)
  Call PropBag.WriteProperty("TrueFlowIdentifier", msTrueFlowIdentifier, "")
  Call PropBag.WriteProperty("DecisionFlowExpressionID", mlngDecisionFlowExprID, 0)
  
  Call PropBag.WriteProperty("AppMajor", miAppMajor, 0)
  Call PropBag.WriteProperty("AppMinor", miAppMinor, 0)
  Call PropBag.WriteProperty("AppRevision", miAppRevision, 0)
  
End Sub

Public Property Get ElementPicture() As StdPicture
  Set ElementPicture = imgPicture.Picture
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

Public Property Get ElementType() As ElementType
  ' Return the current element type.
  ElementType = elem_Decision
End Property

Public Property Get ElementTypeDescription() As String
  ' Return the current element type description.
  ElementTypeDescription = "Decision"
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
      .Picture = imgHighlight.Picture
    Else
      .Picture = imgPicture.Picture
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

Public Property Get Identifier() As String
  ' Return the Identifier
  Identifier = msIdentifier
End Property

Public Property Let Identifier(ByVal psNewValue As String)
  ' Set the identifier
  msIdentifier = psNewValue
  PropertyChanged "Identifier"
End Property

Public Property Get ControlIndex() As Integer
  ControlIndex = miControlIndex
End Property

Public Property Let ControlIndex(ByVal piIndex As Integer)
  miControlIndex = piIndex
End Property

Private Sub ResizeElement()
  ' Resize to the current picture.
  With UserControl
    .Height = imgPicture.Height
    .Width = imgPicture.Width
  End With
End Sub

Private Sub DrawElement()
  ' Display the appropriate picture.
  Dim lblTemp As Label
  
  With UserControl
    If mfHighlighted Then
      .Picture = imgHighlight.Picture
    Else
      .Picture = imgPicture.Picture
    End If
    .MaskPicture = imgMask.Picture
  End With
  
  For Each lblTemp In lblDecisionCaption
    lblTemp.Visible = True
  Next lblTemp
  
  ResizeElement
  PositionMiniCaptions
  
  Caption = ""
  
End Sub

Private Sub SetMiniCaptionFont(psFontName As String, piSize As Integer)

  On Error GoTo ErrorTrap
  
  Dim ctlLabel As Label
  
  For Each ctlLabel In lblDecisionCaption
    ctlLabel.Font.Name = psFontName
    ctlLabel.Font.Size = piSize
  Next ctlLabel
  Set ctlLabel = Nothing
  
  If (miDecisionCaptionType = decisionCaption_tick_cross) _
    And (UCase(Trim(lblDecisionCaption(lblDecisionCaption.LBound).Font.Name)) <> UCase(Trim(psFontName))) Then
    
    ' Couldn't set the font to be a graphic one (eg. Wingdings) so try another one.
    Select Case UCase(Trim(psFontName))
      Case UCase(Trim("Wingdings 2"))
        lblDecisionCaption(0).Caption = "ü"
        lblDecisionCaption(1).Caption = "û"
    
        SetMiniCaptionFont "Wingdings", 8
    
      Case Else
        DecisionCaptionType = decisionCaption_T_F
        
    End Select
  End If

ErrorTrap:
End Sub

Private Sub StaticCaptionInitialize()
  On Error GoTo ErrorTrap
  
  PositionMiniCaptions
  
  Exit Sub
      
ErrorTrap:

End Sub

Private Sub PositionCaption()
  ' Display the appropriate picture.
  Dim sngSingleLineLength As Single
  
  With lblCaption
    If Len(.Caption) > 0 Then
      ' Disable wordwrap to get the string length in a single line
      .WordWrap = False
      sngSingleLineLength = .Width
      .WordWrap = True
      
      If sngSingleLineLength > (UserControl.Width * 0.6) Then
        .Width = (UserControl.Width * 0.6)
      End If
    
      .Left = (UserControl.Width - .Width) / 2
      .Top = (UserControl.Height - .Height) / 2

      .Visible = True
    Else
      .Visible = False
    End If
  End With
  
End Sub

Private Sub PositionMiniCaptions()
  Const VERTICALGAP = 50
  Const HORIZONTALGAP = 200
  Const HORIZONTALGAP_SMALL = 50
  
  ' Decision element mini-captions
  With lblDecisionCaption(0)
    .Left = (UserControl.Width - .Width) / 2
    .Top = UserControl.Height - .Height - VERTICALGAP
  End With

  With lblDecisionCaption(1)
    .Left = UserControl.Width - .Width - HORIZONTALGAP
    .Top = (UserControl.Height - .Height) / 2
  End With
    
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
  
  ReDim mavOutboundFlowInfo(7, 2)
  mavOutboundFlowInfo(1, 1) = decisionOutFlow_True
  mavOutboundFlowInfo(2, 1) = lineDirection_Down
  mavOutboundFlowInfo(3, 1) = (UserControl.Width / 2)
  mavOutboundFlowInfo(4, 1) = UserControl.Height
  mavOutboundFlowInfo(5, 1) = -1     ' -1 indicates no maximum
  mavOutboundFlowInfo(6, 1) = 1
  mavOutboundFlowInfo(7, 1) = "True"

  mavOutboundFlowInfo(1, 2) = decisionOutFlow_False
  mavOutboundFlowInfo(2, 2) = lineDirection_Right
  mavOutboundFlowInfo(3, 2) = UserControl.Width
  mavOutboundFlowInfo(4, 2) = (UserControl.Height / 2)
  mavOutboundFlowInfo(5, 2) = -1     ' -1 indicates no maximum
  mavOutboundFlowInfo(6, 2) = 1
  mavOutboundFlowInfo(7, 2) = "False"

  OutboundFlows_Information = mavOutboundFlowInfo
  
End Function


Public Property Get InboundFlow_Direction() As LineDirection
  ' Return the line direction for inbound flows.
  InboundFlow_Direction = lineDirection_Up
End Property

Public Property Get InboundFlows_Maximum() As Integer
  ' Return the maximum number of inbound flows for the element.
  InboundFlows_Maximum = 1
End Property

Public Property Get InboundFlows_Minimum() As Integer
  ' Return the minimum number of inbound flows for the element.
  InboundFlows_Minimum = 1
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

Public Property Get DecisionCaptionType() As DecisionCaptionType
  DecisionCaptionType = miDecisionCaptionType
End Property

Public Property Let DecisionCaptionType(ByVal piNewValue As DecisionCaptionType)
  On Error GoTo ErrorTrap
  
  miDecisionCaptionType = piNewValue
  PropertyChanged "DecisionCaptionType"
  
  Select Case miDecisionCaptionType
    Case decisionCaption_Y_N
      lblDecisionCaption(0).Caption = "Y"
      lblDecisionCaption(1).Caption = "N"
    
    Case decisionCaption_1_0
      lblDecisionCaption(0).Caption = "1"
      lblDecisionCaption(1).Caption = "0"
    
    Case decisionCaption_tick_cross
      lblDecisionCaption(0).Caption = "P"
      lblDecisionCaption(1).Caption = "O"
    
    Case Else ' decisionCaption_T_F
      lblDecisionCaption(0).Caption = "T"
      lblDecisionCaption(1).Caption = "F"
  End Select
  
  If miDecisionCaptionType = decisionCaption_tick_cross Then
    SetMiniCaptionFont "Wingdings 2", 8
  Else
    SetMiniCaptionFont "Small Fonts", 5
  End If
  
  PositionMiniCaptions
  
  Exit Property
  
ErrorTrap:

End Property

Public Property Get DecisionFlowType() As DecisionFlowTypes
  DecisionFlowType = miDecisionFlowType
End Property

Public Property Let DecisionFlowType(ByVal piNewValue As DecisionFlowTypes)
  miDecisionFlowType = piNewValue
  PropertyChanged "DecisionFlowType"
  
  Select Case miDecisionFlowType
    Case decisionFlowType_Expression
      msTrueFlowIdentifier = ""
    
    Case Else ' decisionFlowType_Button
      mlngDecisionFlowExprID = 0
  End Select
End Property

Public Property Get TrueFlowIdentifier() As String
  ' Return the TrueFlowIdentifier
  TrueFlowIdentifier = msTrueFlowIdentifier
End Property

Public Property Let TrueFlowIdentifier(ByVal psNewValue As String)
  ' Set the TrueFlowIdentifier
  msTrueFlowIdentifier = psNewValue
  PropertyChanged "TrueFlowIdentifier"
End Property

Public Property Get DecisionCaptionFont() As Font
  ' Return the decision caption font.
  Set DecisionCaptionFont = MiniCaptionFont
End Property

Public Property Get MiniCaptionFont() As Font
  ' Return the mini caption font.
  Set MiniCaptionFont = lblDecisionCaption(0).Font
End Property

Public Property Get DecisionFalseCaption() As String
  ' Return the false decision caption
  DecisionFalseCaption = MiniCaption(1)
End Property

Public Property Get MiniCaption(piIndex As Integer) As String
  ' Return the mini caption.
  MiniCaption = lblDecisionCaption(piIndex).Caption
End Property

Public Property Get DecisionFalseCaptionHeight() As Single
  ' Return the decision false caption's height
  DecisionFalseCaptionHeight = MiniCaptionHeight(1)
End Property

Public Property Get MiniCaptionHeight(piIndex As Integer) As Single
  ' Return the mini caption height.
  MiniCaptionHeight = lblDecisionCaption(piIndex).Height
End Property

Public Property Get DecisionFalseCaptionHorizontalPosition() As Single
  ' Return the decision false caption's Horizontal position
  DecisionFalseCaptionHorizontalPosition = MiniCaptionHorizontalPosition(1)
End Property

Public Property Get MiniCaptionHorizontalPosition(piIndex As Integer) As Single
  ' Return the mini caption Horizontal position.
  MiniCaptionHorizontalPosition = lblDecisionCaption(piIndex).Left
End Property

Public Property Get DecisionFalseCaptionVerticalPosition() As Single
  ' Return the decision false caption's Vertical position
  DecisionFalseCaptionVerticalPosition = MiniCaptionVerticalPosition(1)
End Property

Public Property Get MiniCaptionVerticalPosition(piIndex As Integer) As Single
  ' Return the mini caption Vertical position.
  MiniCaptionVerticalPosition = lblDecisionCaption(piIndex).Top
End Property

Public Property Get DecisionFalseCaptionWidth() As Single
  ' Return the decision false caption's Width
  DecisionFalseCaptionWidth = MiniCaptionWidth(1)
End Property

Public Property Get MiniCaptionWidth(piIndex As Integer) As Single
  ' Return the mini caption width.
  MiniCaptionWidth = lblDecisionCaption(piIndex).Width
End Property

Public Property Get DecisionFlowExpressionID() As Long
  DecisionFlowExpressionID = mlngDecisionFlowExprID
End Property

Public Property Let DecisionFlowExpressionID(ByVal plngNewValue As Long)
  mlngDecisionFlowExprID = plngNewValue
End Property

Public Property Get DecisionTrueCaption() As String
  ' Return the true decision caption
  DecisionTrueCaption = MiniCaption(0)
End Property

Public Property Get DecisionTrueCaptionHeight() As Single
  ' Return the decision true caption's height
  DecisionTrueCaptionHeight = MiniCaptionHeight(0)
End Property

Public Property Get DecisionTrueCaptionHorizontalPosition() As Single
  ' Return the decision true caption's Horizontal position
  DecisionTrueCaptionHorizontalPosition = MiniCaptionHorizontalPosition(0)
End Property

Public Property Get DecisionTrueCaptionVerticalPosition() As Single
  ' Return the decision true caption's Vertical position
  DecisionTrueCaptionVerticalPosition = MiniCaptionVerticalPosition(0)
End Property

Public Property Get DecisionTrueCaptionWidth() As Single
  ' Return the decision true caption's Width
  DecisionTrueCaptionWidth = MiniCaptionWidth(0)
End Property

Public Property Get hWnd() As Variant
  hWnd = UserControl.hWnd
End Property
