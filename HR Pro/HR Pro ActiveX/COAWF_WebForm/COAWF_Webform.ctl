VERSION 5.00
Begin VB.UserControl COAWF_Webform 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   1920
   ScaleWidth      =   5040
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      Picture         =   "COAWF_Webform.ctx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1560
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   1680
      Picture         =   "COAWF_Webform.ctx":40DC
      ScaleHeight     =   795
      ScaleWidth      =   1560
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picHighlight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   3360
      Picture         =   "COAWF_Webform.ctx":81B8
      ScaleHeight     =   795
      ScaleWidth      =   1560
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1560
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
      Left            =   1980
      TabIndex        =   4
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   510
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWebFormCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¹"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "COAWF_Webform"
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

Public Enum WebFormOutboundFlows
  webFormOutFlow_Normal = 0
  webFormOutFlow_Timeout = 1
End Enum

Public Enum TimeoutPeriod
  TIMEOUT_MINUTE = 0
  TIMEOUT_HOUR = 1
  TIMEOUT_DAY = 2
  TIMEOUT_WEEK = 3
  TIMEOUT_MONTH = 4
  TIMEOUT_YEAR = 5
End Enum

Public Enum MessageType
  MESSAGE_SYSTEMDEFAULT = 0
  MESSAGE_CUSTOM = 1
  MESSAGE_NONE = 2
End Enum

' WebForm specific properties
Private mlngWebFormBGImageID As Long
Private mlngWebFormBGImageLocation As Long
Private mlngWebFormFGColor As Long
Private mlngWebFormBGColor As Long
Private mfntWebFormDefaultFont As StdFont
Private mlngWebFormWidth As Long
Private mlngWebFormHeight As Long
Private mlngWebFormTimeoutFrequency As Long
Private miWebFormTimeoutPeriod As TimeoutPeriod
Private mfTimeoutExcludeWeekend As Boolean
Private mlngDescriptionExprID As Long
Private mfDescriptionHasWorkflowName As Boolean
Private mfDescriptionHasElementCaption As Boolean
Private masValidations() As String
Private miCompletionMessageType As MessageType
Private msCompletionMessage As String
Private miSavedForLaterMessageType As MessageType
Private msSavedForLaterMessage As String
Private miFollowOnFormsMessageType As MessageType
Private msFollowOnFormsMessage As String

' WebForm/Email specific properties
Private masItems() As String

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
  ReDim masItems(0, 0)
  ReDim masValidations(0, 0)
  ReDim mavDataColumns(0, 0)

  StaticCaptionInitialize
  DrawElement
End Sub

Private Sub UserControl_InitProperties()
  ' Initialise the properties.
  On Error Resume Next
  
  Caption = "Web Form"
  Identifier = ""
  Set Font = Ambient.Font
    
  AppMajor = 0
  AppMinor = 0
  AppRevision = 0
  
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Load property values from storage.
  On Error Resume Next

  ' Read the previous set of properties.
  Caption = PropBag.ReadProperty("Caption", "Web Form")
  Identifier = PropBag.ReadProperty("Identifier", "")
  Highlighted = PropBag.ReadProperty("Highlighted", False)
  
  AppMajor = PropBag.ReadProperty("AppMajor", 0)
  AppMinor = PropBag.ReadProperty("AppMinor", 0)
  AppRevision = PropBag.ReadProperty("AppRevision", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
  
  ' Save the current set of properties.
  Call PropBag.WriteProperty("Highlighted", mfHighlighted, False)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
  Call PropBag.WriteProperty("Identifier", msIdentifier, "")
    
  Call PropBag.WriteProperty("AppMajor", miAppMajor, 0)
  Call PropBag.WriteProperty("AppMinor", miAppMinor, 0)
  Call PropBag.WriteProperty("AppRevision", miAppRevision, 0)
End Sub

Public Property Get ElementPicture() As StdPicture
  Set ElementPicture = picPicture.Picture
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
  ElementType = elem_WebForm
End Property

Public Property Get ElementTypeDescription() As String
  ' Return the current element type description.
  ElementTypeDescription = "Web Form"
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
      .Picture = picHighlight.Picture
    Else
      .Picture = picPicture.Picture
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

Private Sub ResizeElement()
  ' Resize to the current picture.
  With UserControl
    .Height = picPicture.Height
    .Width = picPicture.Width
  End With
End Sub

Private Sub DrawElement()
  ' Display the appropriate picture.
  With UserControl
    If mfHighlighted Then
      .Picture = picHighlight.Picture
    Else
      .Picture = picPicture.Picture
    End If
    .MaskPicture = picMask.Picture
  End With
    
  lblWebFormCaption.Visible = (mlngWebFormTimeoutFrequency > 0)
  
  ResizeElement
  PositionMiniCaptions
  
  Caption = vbNullString
End Sub

Private Sub SetMiniCaptionFont(psFontName As String, piSize As Integer)

  On Error GoTo ErrorTrap
  
  Dim ctlLabel As VB.Label
  
  lblWebFormCaption.Font.Name = psFontName
  lblWebFormCaption.Font.Size = piSize

  If (UCase(Trim(lblWebFormCaption.Font.Name)) <> UCase(Trim(psFontName))) Then
    ' Couldn't set the font to be a graphic one (eg. Wingdings) so try another one.
    lblWebFormCaption.Caption = "T"
    SetMiniCaptionFont "Small Fonts", 5
  End If

ErrorTrap:
End Sub

Private Sub StaticCaptionInitialize()
  On Error GoTo ErrorTrap
  
  lblWebFormCaption.Caption = "¹"
  SetMiniCaptionFont "Wingdings", 7
  PositionMiniCaptions
  
  Exit Sub
      
ErrorTrap:
  Exit Sub
End Sub

Public Function TimeoutPeriodDescription(plngFrequency As Long, piPeriod As TimeoutPeriod) As String
  Dim fKnownPeriod As Boolean
  Dim sDescription As String
  
  fKnownPeriod = False
  sDescription = "<unknown>"
  
  Select Case piPeriod
    Case TIMEOUT_MINUTE
      sDescription = "minute"
      fKnownPeriod = True
    Case TIMEOUT_HOUR
      sDescription = "hour"
      fKnownPeriod = True
    Case TIMEOUT_DAY
      sDescription = "day"
      fKnownPeriod = True
    Case TIMEOUT_WEEK
      sDescription = "week"
      fKnownPeriod = True
    Case TIMEOUT_MONTH
      sDescription = "month"
      fKnownPeriod = True
    Case TIMEOUT_YEAR
      sDescription = "year"
      fKnownPeriod = True
  End Select
  
  If fKnownPeriod Then
    sDescription = CStr(plngFrequency) & " " & sDescription
    
    If (plngFrequency <> 1) Then
      sDescription = sDescription & "s"
    End If
  End If
  
  TimeoutPeriodDescription = sDescription
End Function

Private Sub PositionCaption()
  ' Display the appropriate picture.
  Dim sngSingleLineLength As Single
  
  With lblCaption
    If Len(.Caption) > 0 Then
      ' Disable wordwrap to get the string length in a single line
      .WordWrap = False
      sngSingleLineLength = .Width
      .WordWrap = True
      
      If sngSingleLineLength > ((UserControl.Width * 0.9) - 100) Then
        .Width = ((UserControl.Width * 0.9) - 100)
      End If
  
      .Left = ((UserControl.Width - .Width) / 2)
      .Top = ((UserControl.Height * 1.2) - .Height) / 2
      
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
  
  ' Web Form element mini-captions
  With lblWebFormCaption
    .Left = UserControl.Width - .Width - HORIZONTALGAP_SMALL
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
  ReDim mavOutboundFlowInfo(7, 2)
  mavOutboundFlowInfo(1, 1) = webFormOutFlow_Normal
  mavOutboundFlowInfo(2, 1) = lineDirection_Down
  mavOutboundFlowInfo(3, 1) = (UserControl.Width / 2)
  mavOutboundFlowInfo(4, 1) = UserControl.Height
  mavOutboundFlowInfo(5, 1) = -1     ' -1 indicates no maximum
  mavOutboundFlowInfo(6, 1) = 1
  mavOutboundFlowInfo(7, 1) = "Normal"

  mavOutboundFlowInfo(1, 2) = webFormOutFlow_Timeout
  mavOutboundFlowInfo(2, 2) = lineDirection_Right
  mavOutboundFlowInfo(3, 2) = UserControl.Width
  mavOutboundFlowInfo(4, 2) = (UserControl.Height / 2)
  mavOutboundFlowInfo(5, 2) = -1     ' -1 indicates no maximum
  mavOutboundFlowInfo(6, 2) = 0
  mavOutboundFlowInfo(7, 2) = "Timeout"

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
  InboundFlow_YOffset = 75
End Property

Public Property Get ConnectorPairIndex() As Integer
  ' Return the miConnectorPairIndex
  ConnectorPairIndex = miConnectorPairIndex
End Property

Public Property Let ConnectorPairIndex(ByVal piNewValue As Integer)
  ' Set the miConnectorPairIndex
  miConnectorPairIndex = piNewValue
End Property

Public Property Get WebFormBGImageID() As Long
  'WebFormBGImageID should contain an image from the ASRSysPictures table.
  WebFormBGImageID = mlngWebFormBGImageID
End Property

Public Property Let WebFormBGImageID(New_ImageID As Long)
  'WebFormBGImageID should contain an image from the ASRSysPictures table.
  mlngWebFormBGImageID = New_ImageID
End Property

Public Property Get WebFormBGImageLocation() As Long
  WebFormBGImageLocation = mlngWebFormBGImageLocation
End Property

Public Property Let WebFormBGImageLocation(New_Location As Long)
  mlngWebFormBGImageLocation = New_Location
End Property

Public Property Get WebFormFGColor() As Long
  WebFormFGColor = mlngWebFormFGColor
End Property
Public Property Get WebFormBGColor() As Long
  WebFormBGColor = mlngWebFormBGColor
End Property

Public Property Let WebFormBGColor(New_Color As Long)
  mlngWebFormBGColor = New_Color
End Property

Public Property Let WebFormFGColor(New_Color As Long)
  mlngWebFormFGColor = New_Color
End Property

Public Property Get WebFormDefaultFont() As StdFont
  Set WebFormDefaultFont = mfntWebFormDefaultFont
End Property

Public Property Set WebFormDefaultFont(ByVal New_Font As StdFont)
  Set mfntWebFormDefaultFont = New_Font
End Property

Public Property Get WebFormWidth() As Long
  WebFormWidth = mlngWebFormWidth
End Property
Public Property Let WebFormWidth(ByVal New_Width As Long)
  mlngWebFormWidth = New_Width
End Property

Public Property Get WebFormHeight() As Long
  WebFormHeight = mlngWebFormHeight
End Property

Public Property Let WebFormHeight(ByVal New_Height As Long)
  mlngWebFormHeight = New_Height
End Property

Public Property Get WebFormTimeoutPeriod() As TimeoutPeriod
  WebFormTimeoutPeriod = miWebFormTimeoutPeriod
End Property

Public Property Let WebFormTimeoutPeriod(ByVal piNewValue As TimeoutPeriod)
  miWebFormTimeoutPeriod = piNewValue
End Property

Public Property Get WebFormTimeoutFrequency() As Long
  WebFormTimeoutFrequency = mlngWebFormTimeoutFrequency
End Property

Public Property Let WebFormTimeoutFrequency(ByVal plngNewValue As Long)
  mlngWebFormTimeoutFrequency = plngNewValue
  lblWebFormCaption.Visible = (mlngWebFormTimeoutFrequency > 0)
End Property

Public Property Get WebFormTimeoutExcludeWeekend() As Boolean
  WebFormTimeoutExcludeWeekend = mfTimeoutExcludeWeekend
End Property

Public Property Let WebFormTimeoutExcludeWeekend(ByVal pfNewValue As Boolean)
  mfTimeoutExcludeWeekend = pfNewValue
End Property

Public Property Get DescriptionExprID() As Long
  DescriptionExprID = mlngDescriptionExprID
End Property

Public Property Let DescriptionExprID(ByVal plngNewValue As Long)
  mlngDescriptionExprID = plngNewValue
End Property

Public Property Get DescriptionHasWorkflowName() As Boolean
  DescriptionHasWorkflowName = mfDescriptionHasWorkflowName
End Property

Public Property Let DescriptionHasWorkflowName(ByVal pfNewValue As Boolean)
  mfDescriptionHasWorkflowName = pfNewValue
End Property

Public Property Get DescriptionHasElementCaption() As Boolean
  DescriptionHasElementCaption = mfDescriptionHasElementCaption
End Property

Public Property Let DescriptionHasElementCaption(ByVal pfNewValue As Boolean)
  mfDescriptionHasElementCaption = pfNewValue
End Property

Public Property Get Validations() As Variant
  Validations = masValidations
End Property

Public Property Let Validations(ByVal pavNewValue As Variant)
  masValidations = pavNewValue
End Property

Public Property Get Items() As Variant
  Items = masItems
End Property

Public Property Let Items(ByVal pavNewValue As Variant)
  masItems = pavNewValue
End Property

Public Property Get TimeoutCaption() As String
  ' Return the Timeout caption
  TimeoutCaption = MiniCaption(0)
End Property

Public Property Get MiniCaption(piIndex As Integer) As String
  ' Return the mini caption.
  MiniCaption = lblWebFormCaption.Caption
End Property

Public Property Get TimeoutCaptionFont() As Font
  ' Return the Timeout caption font.
  Set TimeoutCaptionFont = MiniCaptionFont
End Property

Public Property Get MiniCaptionFont() As Font
  ' Return the mini caption font.
  Set MiniCaptionFont = lblWebFormCaption.Font
End Property

Public Property Get TimeoutCaptionHeight() As Single
  ' Return the Timeout caption's height
  TimeoutCaptionHeight = MiniCaptionHeight(0)
End Property

Public Property Get MiniCaptionHeight(piIndex As Integer) As Single
  ' Return the mini caption height.
  MiniCaptionHeight = lblWebFormCaption.Height
End Property

Public Property Get TimeoutCaptionHorizontalPosition() As Single
  ' Return the Timeout caption's Horizontal position
  TimeoutCaptionHorizontalPosition = MiniCaptionHorizontalPosition(0)
End Property

Public Property Get MiniCaptionHorizontalPosition(piIndex As Integer) As Single
  ' Return the mini caption Horizontal position.
  MiniCaptionHorizontalPosition = lblWebFormCaption.Left
End Property

Public Property Get TimeoutCaptionVerticalPosition() As Single
  ' Return the Timeout caption's Vertical position
  TimeoutCaptionVerticalPosition = MiniCaptionVerticalPosition(0)
End Property

Public Property Get MiniCaptionVerticalPosition(piIndex As Integer) As Single
  ' Return the mini caption Vertical position.
  MiniCaptionVerticalPosition = lblWebFormCaption.Top
End Property

Public Property Get TimeoutCaptionWidth() As Single
  ' Return the Timeout caption's Width
  TimeoutCaptionWidth = MiniCaptionWidth(0)
End Property

Public Property Get MiniCaptionWidth(piIndex As Integer) As Single
  ' Return the mini caption width.
  MiniCaptionWidth = lblWebFormCaption.Width
End Property

Public Property Get hWnd() As Variant
  hWnd = UserControl.hWnd
End Property

Public Property Get WFCompletionMessageType() As MessageType
  WFCompletionMessageType = miCompletionMessageType
  
End Property

Public Property Let WFCompletionMessageType(ByVal piNewValue As MessageType)
  miCompletionMessageType = piNewValue

End Property

Public Property Get WFCompletionMessage() As String
  If miCompletionMessageType = MESSAGE_CUSTOM Then
    WFCompletionMessage = msCompletionMessage
  Else
    WFCompletionMessage = ""
  End If
  
End Property

Public Property Let WFCompletionMessage(ByVal psNewValue As String)
  msCompletionMessage = psNewValue

End Property

Public Property Get WFSavedForLaterMessageType() As MessageType
  WFSavedForLaterMessageType = miSavedForLaterMessageType

End Property

Public Property Let WFSavedForLaterMessageType(ByVal piNewValue As MessageType)
  miSavedForLaterMessageType = piNewValue

End Property

Public Property Get WFSavedForLaterMessage() As String
  If miSavedForLaterMessageType = MESSAGE_CUSTOM Then
    WFSavedForLaterMessage = msSavedForLaterMessage
  Else
    WFSavedForLaterMessage = ""
  End If

End Property

Public Property Let WFSavedForLaterMessage(ByVal psNewValue As String)
  msSavedForLaterMessage = psNewValue

End Property

Public Property Get WFFollowOnFormsMessageType() As MessageType
  WFFollowOnFormsMessageType = miFollowOnFormsMessageType

End Property

Public Property Let WFFollowOnFormsMessageType(ByVal piNewValue As MessageType)
  miFollowOnFormsMessageType = piNewValue

End Property

Public Property Get WFFollowOnFormsMessage() As String
  If miFollowOnFormsMessageType = MESSAGE_CUSTOM Then
    WFFollowOnFormsMessage = msFollowOnFormsMessage
  Else
    WFFollowOnFormsMessage = ""
  End If
  

End Property

Public Property Let WFFollowOnFormsMessage(ByVal psNewValue As String)
  msFollowOnFormsMessage = psNewValue

End Property
