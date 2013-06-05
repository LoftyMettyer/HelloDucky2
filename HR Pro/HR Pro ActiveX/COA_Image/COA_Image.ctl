VERSION 5.00
Begin VB.UserControl COA_Image 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1485
   ScaleHeight     =   1485
   ScaleWidth      =   1485
   ToolboxBitmap   =   "COA_Image.ctx":0000
   Begin VB.PictureBox picPicture 
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   1320
         X2              =   1320
         Y1              =   165
         Y2              =   1060
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   50
         X2              =   50
         Y1              =   50
         Y2              =   945
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   150
         X2              =   1305
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   50
         X2              =   1000
         Y1              =   50
         Y2              =   50
      End
      Begin VB.Image imgImage 
         Height          =   1455
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "COA_Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

Public Enum OLEType
  OLE_LOCAL = 0
  OLE_SERVER = 1
  OLE_EMBEDDED = 2
  OLE_UNC = 3
End Enum

'Default Property Values:
Const m_def_ASRDataField = 0
Const m_def_ForeColor = 0

'Property Variables:
Dim m_ASRDataField As Variant
Dim m_ForeColor As Long

Dim miOLEType As OLEType
Dim mlngColumnID As Long
Dim mobjEmbeddedStream As ADODB.Stream

Private mblnShowSelectionMarkers As Boolean
Private mblnSelecting As Boolean

'Event Declarations:
Event Click() 'MappingInfo=imgImage,imgImage,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Resize() 'MappingInfo=picPicture,picPicture,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

Event SpacePressed() ' RH 14/07/00 - To allow keypres

Public Property Get BackColor() As OLE_COLOR
  ' Return the back colour of the control
  BackColor = picPicture.BackColor
  
End Property
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
  ' Set the back colour of the individual controls
  picPicture.BackColor = NewColor
  
End Property
Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    pdwSupportedOptions As Long, _
                                                    pdwEnabledOptions As Long)

    Dim Rc      As Long
    Dim rClsId  As udtGUID
    Dim IID     As String
    Dim bIID()  As Byte

    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or _
                          INTERFACESAFE_FOR_UNTRUSTED_DATA

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                pdwEnabledOptions = IIf(m_fSafeForScripting, _
              INTERFACESAFE_FOR_UNTRUSTED_CALLER, 0)
                Exit Sub
            Case IID_IPersistStorage, IID_IPersistStream, _
               IID_IPersistPropertyBag
                pdwEnabledOptions = IIf(m_fSafeForInitializing, _
              INTERFACESAFE_FOR_UNTRUSTED_DATA, 0)
                Exit Sub
            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    ByVal dwOptionsSetMask As Long, _
                                                    ByVal dwEnabledOptions As Long)
    Dim Rc          As Long
    Dim rClsId      As udtGUID
    Dim IID         As String
    Dim bIID()      As Byte

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
             INTERFACESAFE_FOR_UNTRUSTED_CALLER) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForScripting Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case IID_IPersistStorage, IID_IPersistStream, _
          IID_IPersistPropertyBag
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
              INTERFACESAFE_FOR_UNTRUSTED_DATA) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForInitializing Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Public Property Let Selecting(ByVal Value As Boolean)
  
  mblnSelecting = Value
  
End Property

Public Property Get Selecting() As Boolean
  
  Selecting = mblnSelecting

End Property

Public Property Let ShowSelectionMarkers(ByVal Value As Boolean)
  
  If Value = True Then
    If Selecting = True Then
      mblnShowSelectionMarkers = True
    End If
  Else
    If Selecting = True Then
      Value = True
    End If
  End If
  
  mblnShowSelectionMarkers = Value
  
  Line1.Visible = Value
  Line2.Visible = Value
  Line3.Visible = Value
  Line4.Visible = Value
  
End Property

Public Property Get ShowSelectionMarkers() As Boolean
  
  ShowSelectionMarkers = mblnShowSelectionMarkers

End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
  m_ForeColor = New_ForeColor
  PropertyChanged "ForeColor"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgImage,imgImage,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = imgImage.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  imgImage.Enabled = New_Enabled
  picPicture.Enabled = New_Enabled
  UserControl.Enabled = New_Enabled
  
  PropertyChanged "Enabled"
  
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgImage,imgImage,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = imgImage.BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  imgImage.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"

End Property

Private Sub imgImage_Click()
  
Selecting = True
RaiseEvent Click
  
End Sub

Private Sub picPicture_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeySpace Then
    Selecting = True
    RaiseEvent SpacePressed
  End If
  
  KeyCode = 0
  Shift = 0

End Sub

Private Sub picPicture_Resize()
  RaiseEvent Resize
  
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgImage,imgImage,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
  Set Picture = imgImage.Picture

End Property

Public Property Set Picture(ByVal New_Picture As Picture)
  Set imgImage.Picture = New_Picture
  PropertyChanged "Picture"

End Property

Public Sub SetPicturePath(ByVal psNewValue As String)
  ' Used by the intranet
  On Error GoTo ErrorTrap
  
  Set imgImage.Picture = LoadPicture(psNewValue)
  PropertyChanged "Picture"

  Exit Sub
  
ErrorTrap:
  Set imgImage.Picture = LoadPicture()
  PropertyChanged "Picture"
End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_ForeColor = m_def_ForeColor
  m_ASRDataField = m_def_ASRDataField
  
  BackColor = vbButtonFace
  
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Load property values from storage.
  ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  Enabled = PropBag.ReadProperty("Enabled", True)
  BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  DMIAppearance = PropBag.ReadProperty("DMIAppearance", 1)
  Set Picture = PropBag.ReadProperty("Picture", Nothing)
  ASRDataField = PropBag.ReadProperty("ASRDataField", m_def_ASRDataField)
  
  BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)

End Sub

Private Sub UserControl_Resize()
  ' Resize the constituent controls.
  picPicture.Height = UserControl.Height
  imgImage.Height = picPicture.Height
  picPicture.Width = UserControl.Width
  imgImage.Width = picPicture.Width
  
  ' Top horizontal
  Line1.X1 = 50
  Line1.X2 = (picPicture.ScaleWidth - 50)
  Line1.Y1 = 50
  Line1.Y2 = 50
  
  ' Bottom horizontal
  Line2.X1 = 50
  Line2.X2 = (picPicture.ScaleWidth - 50)
  Line2.Y1 = (picPicture.ScaleHeight - 50)
  Line2.Y2 = (picPicture.ScaleHeight - 50)
  
  ' Left Vertical
  Line3.X1 = 50
  Line3.X2 = 50
  Line3.Y1 = 50
  Line3.Y2 = (picPicture.ScaleHeight - 50)
  
  ' Right Vertical
  Line4.X1 = (picPicture.ScaleWidth - 50)
  Line4.X2 = (picPicture.ScaleWidth - 50)
  Line4.Y1 = 50
  Line4.Y2 = (picPicture.ScaleHeight - 50)
  
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  ' Write property values to the property bag.
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("Enabled", imgImage.Enabled, True)
  Call PropBag.WriteProperty("BorderStyle", imgImage.BorderStyle, 0)
  Call PropBag.WriteProperty("DMIAppearance", picPicture.Appearance, 1)
  Call PropBag.WriteProperty("Picture", Picture, Nothing)
  Call PropBag.WriteProperty("ASRDataField", m_ASRDataField, m_def_ASRDataField)
  
  Call PropBag.WriteProperty("BackColor", picPicture.BackColor, vbButtonFace)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ASRDataField() As Variant
  ASRDataField = m_ASRDataField

End Property

Public Property Let ASRDataField(ByVal New_ASRDataField As Variant)
  m_ASRDataField = New_ASRDataField
  PropertyChanged "ASRDataField"

End Property


Public Property Let DMIAppearance(ByVal piNewValue As Integer)
  picPicture.Appearance = piNewValue
  PropertyChanged "DMIAppearance"
End Property



Public Property Let EmbeddedStream(ByRef pobjStream As ADODB.Stream)
  Set mobjEmbeddedStream = pobjStream
End Property

Public Property Get EmbeddedStream() As ADODB.Stream
  Set EmbeddedStream = mobjEmbeddedStream
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
End Property

Public Property Let ColumnID(plngNewValue As Long)
  mlngColumnID = plngNewValue
End Property

Public Property Get OLEType() As OLEType
  OLEType = miOLEType
End Property
  
Public Property Let OLEType(ByVal piNewValue As OLEType)
  miOLEType = piNewValue
End Property

