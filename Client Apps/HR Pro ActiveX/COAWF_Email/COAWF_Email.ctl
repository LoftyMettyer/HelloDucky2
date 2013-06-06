VERSION 5.00
Begin VB.UserControl COAWF_Email 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
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
   ScaleHeight     =   1800
   ScaleWidth      =   5205
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
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   510
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgPicture 
      Height          =   780
      Left            =   120
      Picture         =   "COAWF_Email.ctx":0000
      Top             =   840
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image imgMask 
      Height          =   780
      Left            =   1800
      Picture         =   "COAWF_Email.ctx":3FA4
      Top             =   840
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image imgHighlight 
      Height          =   780
      Left            =   3480
      Picture         =   "COAWF_Email.ctx":7F48
      Top             =   840
      Visible         =   0   'False
      Width           =   1560
   End
End
Attribute VB_Name = "COAWF_Email"
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

Public Enum WorkflowEmailAttachmentTypes
  ' Numbering out of sequence as we started using the WorkflowWebFormItemTypes
  ' enum for this, which included differenty items. Sorry.
  giWFEMAILITEM_UNKNOWN = -1
  giWFEMAILITEM_DBVALUE = 1
  giWFEMAILITEM_WFVALUE = 4
  giWFEMAILITEM_FILEATTACHMENT = 18 ' NB. Only used in emails.
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

' Email specific properties
Private mlngEmailID As Long
Private mlngEmailCCID As Long
Private mlngEmailBCCID As Long
Private miEmailRecord As Integer
Private msEmailSubject As String

Private msRecordSelectorIdentifier As String
Private msRecordSelectorWebFormIdentifier As String

Private masItems() As String

Private miAttachment_Type As WorkflowEmailAttachmentTypes
Private msAttachment_File As String
Private msAttachment_WFElementIdentifier As String
Private msAttachment_WFValueIdentifier As String
Private mlngAttachment_DBColumnID As Long
Private miAttachment_DBRecord As Integer
Private msAttachment_DBElement As String
Private msAttachment_DBValue As String


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
  
  DrawElement
End Sub

Private Sub UserControl_InitProperties()
  ' Initialise the properties.
  On Error Resume Next
  
  Caption = "Email"
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
  Caption = PropBag.ReadProperty("Caption", "Email")
  Identifier = PropBag.ReadProperty("Identifier", "")
  Highlighted = PropBag.ReadProperty("Highlighted", False)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  
  RecordSelectorIdentifier = PropBag.ReadProperty("RecordSelectorIdentifier", "")
  RecordSelectorWebFormIdentifier = PropBag.ReadProperty("RecordSelectorWebFormIdentifier", "")

  AppMajor = PropBag.ReadProperty("AppMajor", 0)
  AppMinor = PropBag.ReadProperty("AppMinor", 0)
  AppRevision = PropBag.ReadProperty("AppRevision", 0)
  
  Attachment_Type = PropBag.ReadProperty("Attachment_Type", giWFEMAILITEM_UNKNOWN)
  Attachment_File = PropBag.ReadProperty("Attachment_File", "")
  Attachment_WFElementIdentifier = PropBag.ReadProperty("Attachment_WFElementIdentifier", "")
  Attachment_WFValueIdentifier = PropBag.ReadProperty("Attachment_WFValueIdentifier", "")
  Attachment_DBColumnID = PropBag.ReadProperty("Attachment_DBColumnID", 0)
  Attachment_DBRecord = PropBag.ReadProperty("Attachment_DBRecord", 0)
  Attachment_DBElement = PropBag.ReadProperty("Attachment_DBElement", "")
  Attachment_DBValue = PropBag.ReadProperty("Attachment_DBValue", "")
  
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

  Call PropBag.WriteProperty("AppMajor", miAppMajor, 0)
  Call PropBag.WriteProperty("AppMinor", miAppMinor, 0)
  Call PropBag.WriteProperty("AppRevision", miAppRevision, 0)
  
  
  Call PropBag.WriteProperty("Attachment_Type", miAttachment_Type, giWFEMAILITEM_UNKNOWN)
  Call PropBag.WriteProperty("Attachment_File", msAttachment_File, "")
  Call PropBag.WriteProperty("Attachment_WFElementIdentifier", msAttachment_WFElementIdentifier, "")
  Call PropBag.WriteProperty("Attachment_WFValueIdentifier", msAttachment_WFValueIdentifier, "")
  Call PropBag.WriteProperty("Attachment_DBColumnID", mlngAttachment_DBColumnID, 0)
  Call PropBag.WriteProperty("Attachment_DBRecord", miAttachment_DBRecord, 0)
  Call PropBag.WriteProperty("Attachment_DBElement", msAttachment_DBElement, "")
  Call PropBag.WriteProperty("Attachment_DBValue", msAttachment_DBValue, "")
  
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
  ElementType = elem_Email
End Property

Public Property Get ElementTypeDescription() As String
  ' Return the current element type description.
  ElementTypeDescription = "Email"
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

Private Sub ResizeElement()
  ' Resize to the current picture.
  With UserControl
    .Height = imgPicture.Height
    .Width = imgPicture.Width
  End With
End Sub

Private Sub DrawElement()
  ' Display the appropriate picture.
  With UserControl
    If mfHighlighted Then
      .Picture = imgHighlight.Picture
    Else
      .Picture = imgPicture.Picture
    End If
    .MaskPicture = imgMask.Picture
  End With
  
  ResizeElement
    
  Caption = vbNullString
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
      
      If sngSingleLineLength > (UserControl.Width * 0.9) Then
        .Width = (UserControl.Width * 0.9)
      End If
  
      .Left = (UserControl.Width - .Width) / 2
      .Top = ((UserControl.Height * 0.8) - .Height) / 2
      
      .Visible = True
    Else
      .Visible = False
    End If
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
  ReDim mavOutboundFlowInfo(7, 1)
  mavOutboundFlowInfo(1, 1) = 0
  mavOutboundFlowInfo(2, 1) = lineDirection_Down
  mavOutboundFlowInfo(3, 1) = (UserControl.Width / 2)
  mavOutboundFlowInfo(4, 1) = 650
  mavOutboundFlowInfo(5, 1) = -1     ' -1 indicates no maximum
  mavOutboundFlowInfo(6, 1) = 1
  mavOutboundFlowInfo(7, 1) = ""
  
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

Public Property Get Items() As Variant
  Items = masItems
End Property

Public Property Let Items(ByVal pavNewValue As Variant)
  masItems = pavNewValue
End Property

Public Property Get EmailRecord() As Integer
  EmailRecord = miEmailRecord
End Property

Public Property Let EmailRecord(ByVal piNewValue As Integer)
  miEmailRecord = piNewValue
End Property

Public Property Get EmailSubject() As String
  EmailSubject = msEmailSubject
End Property

Public Property Let EmailSubject(ByVal psNewValue As String)
  msEmailSubject = psNewValue
End Property

Public Property Get EmailID() As Long
  EmailID = mlngEmailID
End Property

Public Property Let EmailID(ByVal plngNewValue As Long)
  mlngEmailID = plngNewValue
End Property

Public Property Get EmailCCID() As Long
  EmailCCID = mlngEmailCCID
End Property

Public Property Let EmailCCID(ByVal plngNewValue As Long)
  mlngEmailCCID = plngNewValue
End Property

Public Property Get EmailBCCID() As Long
  EmailBCCID = mlngEmailBCCID
End Property

Public Property Let EmailBCCID(ByVal plngNewValue As Long)
  mlngEmailBCCID = plngNewValue
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

Public Property Get RecordSelectorWebFormIdentifier() As String
  ' Return the RecordSelectorWebFormIdentifier
  RecordSelectorWebFormIdentifier = msRecordSelectorWebFormIdentifier
End Property

Public Property Let RecordSelectorWebFormIdentifier(ByVal psNewValue As String)
  ' Set the RecordSelectorWebFormIdentifier
  msRecordSelectorWebFormIdentifier = psNewValue
  PropertyChanged "RecordSelectorWebFormIdentifier"
End Property

Public Property Get hWnd() As Variant
  hWnd = UserControl.hWnd
End Property

Public Property Get Attachment_Type() As WorkflowEmailAttachmentTypes
  Attachment_Type = miAttachment_Type
  
End Property

Public Property Let Attachment_Type(ByVal piNewValue As WorkflowEmailAttachmentTypes)
  miAttachment_Type = piNewValue
  PropertyChanged "Attachment_Type"

End Property

Public Property Get Attachment_File() As String
  Attachment_File = msAttachment_File
  
End Property

Public Property Let Attachment_File(ByVal psNewValue As String)
  msAttachment_File = psNewValue
  PropertyChanged "Attachment_File"
  
End Property

Public Property Get Attachment_WFElementIdentifier() As String
  Attachment_WFElementIdentifier = msAttachment_WFElementIdentifier
  
End Property

Public Property Let Attachment_WFElementIdentifier(ByVal psNewValue As String)
  msAttachment_WFElementIdentifier = psNewValue
  PropertyChanged "Attachment_WFElementIdentifier"

End Property

Public Property Get Attachment_WFValueIdentifier() As String
  Attachment_WFValueIdentifier = msAttachment_WFValueIdentifier
  
End Property

Public Property Let Attachment_WFValueIdentifier(ByVal psNewValue As String)
  msAttachment_WFValueIdentifier = psNewValue
  PropertyChanged "Attachment_WFValueIdentifier"

End Property

Public Property Get Attachment_DBColumnID() As Long
  Attachment_DBColumnID = mlngAttachment_DBColumnID
  
End Property

Public Property Let Attachment_DBColumnID(ByVal plngNewValue As Long)
  mlngAttachment_DBColumnID = plngNewValue
  PropertyChanged "Attachment_DBColumnID"

End Property

Public Property Get Attachment_DBRecord() As Integer
  Attachment_DBRecord = miAttachment_DBRecord
  
End Property

Public Property Let Attachment_DBRecord(ByVal piNewValue As Integer)
  miAttachment_DBRecord = piNewValue
  PropertyChanged "Attachment_DBRecord"

End Property

Public Property Get Attachment_DBElement() As String
  Attachment_DBElement = msAttachment_DBElement
  
End Property

Public Property Let Attachment_DBElement(ByVal psNewValue As String)
  msAttachment_DBElement = psNewValue
  PropertyChanged "Attachment_DBElement"

End Property

Public Property Get Attachment_DBValue() As String
  Attachment_DBValue = msAttachment_DBValue
  
End Property

Public Property Let Attachment_DBValue(ByVal psNewValue As String)
  msAttachment_DBValue = psNewValue
  PropertyChanged "Attachment_DBValue"

End Property
