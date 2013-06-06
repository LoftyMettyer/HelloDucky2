VERSION 5.00
Begin VB.UserControl COAWF_StoredData 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
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
   ScaleHeight     =   1830
   ScaleWidth      =   5190
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
      Left            =   2340
      TabIndex        =   2
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   510
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblStoredDataCaption 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   120
   End
   Begin VB.Label lblStoredDataCaption 
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
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   60
   End
   Begin VB.Image imgPicture 
      Height          =   780
      Left            =   120
      Picture         =   "COAWF_StoredData.ctx":0000
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgMask 
      Height          =   780
      Left            =   1800
      Picture         =   "COAWF_StoredData.ctx":3FA4
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgHighlight 
      Height          =   780
      Left            =   3480
      Picture         =   "COAWF_StoredData.ctx":7F48
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "COAWF_StoredData"
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

Public Enum StoredDataOutboundFlows
  storedDataOutFlow_Success = 0
  storedDataOutFlow_Failure = 1
End Enum

Public Enum DataAction
  DATAACTION_INSERT = 0
  DATAACTION_UPDATE = 1
  DATAACTION_DELETE = 2
End Enum

' StoredData specific properties
Private mavDataColumns() As Variant
Private miDataAction As DataAction
Private mlngDataTableID As Long

Private miDataRecord As Integer
Private msRecordSelectorIdentifier As String
Private msRecordSelectorWebFormIdentifier As String
Private mlngDataRecordTableID As Long

Private miSecondaryDataRecord As Integer
Private msSecondaryRecordSelectorIdentifier As String
Private msSecondaryRecordSelectorWebFormIdentifier As String
Private mlngSecondaryDataRecordTableID As Long

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
  
  Caption = "Stored Data"
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
  Caption = PropBag.ReadProperty("Caption", "Stored Data")
  Identifier = PropBag.ReadProperty("Identifier", "")
  Highlighted = PropBag.ReadProperty("Highlighted", False)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  
  RecordSelectorIdentifier = PropBag.ReadProperty("RecordSelectorIdentifier", "")
  RecordSelectorWebFormIdentifier = PropBag.ReadProperty("RecordSelectorWebFormIdentifier", "")
  SecondaryRecordSelectorIdentifier = PropBag.ReadProperty("SecondaryRecordSelectorIdentifier", "")
  SecondaryRecordSelectorWebFormIdentifier = PropBag.ReadProperty("SecondaryRecordSelectorWebFormIdentifier", "")

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
    
  Call PropBag.WriteProperty("RecordSelectorIdentifier", msRecordSelectorIdentifier, "")
  Call PropBag.WriteProperty("RecordSelectorWebFormIdentifier", msRecordSelectorWebFormIdentifier, "")
  Call PropBag.WriteProperty("SecondaryRecordSelectorIdentifier", msSecondaryRecordSelectorIdentifier, "")
  Call PropBag.WriteProperty("SecondaryRecordSelectorWebFormIdentifier", msSecondaryRecordSelectorWebFormIdentifier, "")

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
  ElementType = elem_StoredData
End Property

Public Property Get ElementTypeDescription() As String
  ' Return the current element type description.
  ElementTypeDescription = "Stored Data"
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

Public Property Get MiniCaption(piIndex As Integer) As String
  ' Return the mini caption.
  MiniCaption = lblStoredDataCaption(piIndex).Caption
End Property

Public Property Get MiniCaptionFont() As Font
  ' Return the mini caption font.
  Set MiniCaptionFont = lblStoredDataCaption(0).Font
End Property

Public Property Get MiniCaptionHeight(piIndex As Integer) As Single
  ' Return the mini caption height.
  MiniCaptionHeight = lblStoredDataCaption(piIndex).Height
End Property

Public Property Get MiniCaptionWidth(piIndex As Integer) As Single
  ' Return the mini caption width.
  MiniCaptionWidth = lblStoredDataCaption(piIndex).Width
End Property

Public Property Get MiniCaptionHorizontalPosition(piIndex As Integer) As Single
  ' Return the mini caption Horizontal position.
  MiniCaptionHorizontalPosition = lblStoredDataCaption(piIndex).Left
End Property

Public Property Get MiniCaptionVerticalPosition(piIndex As Integer) As Single
  ' Return the mini caption Vertical position.
  MiniCaptionVerticalPosition = lblStoredDataCaption(piIndex).Top
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
  
  ResizeElement
  PositionMiniCaptions
  
  Caption = ""
  
End Sub

Private Sub SetMiniCaptionFont(piElementType As ElementType, _
  psFontName As String, _
  piSize As Integer)

  On Error GoTo ErrorTrap
  
  Dim ctlLabel As Label
  
  For Each ctlLabel In lblStoredDataCaption
    ctlLabel.Font.Name = psFontName
    ctlLabel.Font.Size = piSize
  Next ctlLabel

  If (UCase(Trim(lblStoredDataCaption(lblStoredDataCaption.LBound).Font.Name)) <> UCase(Trim(psFontName))) Then
    ' Couldn't set the font to be a graphic one (eg. Wingdings) so try another one.
    Select Case UCase(Trim(psFontName))
      Case UCase(Trim("Wingdings 2"))
        lblStoredDataCaption(storedDataOutFlow_Success).Caption = "ü"
        lblStoredDataCaption(storedDataOutFlow_Failure).Caption = "û"
    
        SetMiniCaptionFont piElementType, "Wingdings", 8

      Case Else
        lblStoredDataCaption(storedDataOutFlow_Success).Caption = "OK"
        lblStoredDataCaption(storedDataOutFlow_Failure).Caption = "F"
        
        SetMiniCaptionFont piElementType, "Small Fonts", 5
    End Select
  End If
  
ErrorTrap:

End Sub

Private Sub StaticCaptionInitialize()
  On Error GoTo ErrorTrap
  
  StaticCaptionInitialize_StoredData
  
  PositionMiniCaptions
  
  Exit Sub
      
ErrorTrap:
  Exit Sub

End Sub

Private Sub StaticCaptionInitialize_StoredData()
  On Error GoTo ErrorTrap
  
  lblStoredDataCaption(storedDataOutFlow_Success).Caption = "P"
  lblStoredDataCaption(storedDataOutFlow_Failure).Caption = "O"
  
  SetMiniCaptionFont elem_StoredData, "Wingdings 2", 8
  
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
      
      If sngSingleLineLength > (UserControl.Width * 0.7) Then
        .Width = (UserControl.Width * 0.7)
      End If
    
      .Left = (UserControl.Width * 0.1) + ((UserControl.Width * 0.7) - .Width) / 2
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
  
  ' Stored Data element mini-captions
  With lblStoredDataCaption(0)
    .Left = (UserControl.Width - .Width) / 2
    .Top = UserControl.Height - .Height
  End With

  With lblStoredDataCaption(1)
    .Left = (UserControl.Width * 0.95) - .Width - HORIZONTALGAP
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
  mavOutboundFlowInfo(1, 1) = storedDataOutFlow_Success
  mavOutboundFlowInfo(2, 1) = lineDirection_Down
  mavOutboundFlowInfo(3, 1) = (UserControl.Width / 2)
  mavOutboundFlowInfo(4, 1) = UserControl.Height
  mavOutboundFlowInfo(5, 1) = -1     ' -1 indicates no maximum
  mavOutboundFlowInfo(6, 1) = 1
  mavOutboundFlowInfo(7, 1) = "Success"

  mavOutboundFlowInfo(1, 2) = storedDataOutFlow_Failure
  mavOutboundFlowInfo(2, 2) = lineDirection_Right
  mavOutboundFlowInfo(3, 2) = 1330
  mavOutboundFlowInfo(4, 2) = (UserControl.Height / 2)
  mavOutboundFlowInfo(5, 2) = -1     ' -1 indicates no maximum
  mavOutboundFlowInfo(6, 2) = 0
  mavOutboundFlowInfo(7, 2) = "Failure"
  
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

Public Property Get RecordSelectorIdentifier() As String
  ' Return the RecordSelectorIdentifier
  RecordSelectorIdentifier = msRecordSelectorIdentifier
End Property

Public Property Let RecordSelectorIdentifier(ByVal psNewValue As String)
  ' Set the RecordSelectorIdentifier
  msRecordSelectorIdentifier = psNewValue
  PropertyChanged "RecordSelectorIdentifier"
End Property

Public Property Get SecondaryRecordSelectorIdentifier() As String
  ' Return the SecondaryRecordSelectorIdentifier
  SecondaryRecordSelectorIdentifier = msSecondaryRecordSelectorIdentifier
End Property

Public Property Let SecondaryRecordSelectorIdentifier(ByVal psNewValue As String)
  ' Set the SecondaryRecordSelectorIdentifier
  msSecondaryRecordSelectorIdentifier = psNewValue
  PropertyChanged "SecondaryRecordSelectorIdentifier"
End Property

Public Property Get RecordSelectorWebFormIdentifier() As String
  ' Return the RecordSelectorWebFormIdentifier
  RecordSelectorWebFormIdentifier = msRecordSelectorWebFormIdentifier
End Property

Public Property Let RecordSelectorWebFormIdentifier(ByVal psNewValue As String)
  ' Set the RecordSelectorWebFormIdentifier
  msRecordSelectorWebFormIdentifier = psNewValue
  PropertyChanged "RecordSelectorWebFormIdentifier"
End Property

Public Property Get SecondaryRecordSelectorWebFormIdentifier() As String
  ' Return the SecondaryRecordSelectorWebFormIdentifier
  SecondaryRecordSelectorWebFormIdentifier = msSecondaryRecordSelectorWebFormIdentifier
End Property

Public Property Let SecondaryRecordSelectorWebFormIdentifier(ByVal psNewValue As String)
  ' Set the SecondaryRecordSelectorWebFormIdentifier
  msSecondaryRecordSelectorWebFormIdentifier = psNewValue
  PropertyChanged "SecondaryRecordSelectorWebFormIdentifier"
End Property

Public Property Get DataColumns() As Variant
  DataColumns = mavDataColumns
End Property

Public Property Let DataColumns(ByVal pavNewValue As Variant)
  mavDataColumns = pavNewValue
End Property

Public Property Get DataAction() As DataAction
  DataAction = miDataAction
End Property

Public Property Let DataAction(ByVal piNewValue As DataAction)
  miDataAction = piNewValue
End Property

Public Property Get DataTableID() As Long
  DataTableID = mlngDataTableID
End Property

Public Property Let DataTableID(ByVal plngNewValue As Long)
  mlngDataTableID = plngNewValue
End Property

Public Property Get DataRecord() As Integer
  DataRecord = miDataRecord
End Property

Public Property Let DataRecord(ByVal piNewValue As Integer)
  miDataRecord = piNewValue
End Property

Public Property Get SecondaryDataRecord() As Integer
  SecondaryDataRecord = miSecondaryDataRecord
End Property

Public Property Let SecondaryDataRecord(ByVal piNewValue As Integer)
  miSecondaryDataRecord = piNewValue
End Property

Public Property Get SecondaryDataRecordTableID() As Long
  SecondaryDataRecordTableID = mlngSecondaryDataRecordTableID
End Property

Public Property Let SecondaryDataRecordTableID(ByVal plngNewValue As Long)
  mlngSecondaryDataRecordTableID = plngNewValue
End Property

Public Property Get DataRecordTableID() As Long
  DataRecordTableID = mlngDataRecordTableID
End Property

Public Property Let DataRecordTableID(ByVal plngNewValue As Long)
  mlngDataRecordTableID = plngNewValue
End Property

Public Property Get hWnd() As Variant
  hWnd = UserControl.hWnd
End Property
