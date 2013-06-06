VERSION 5.00
Begin VB.UserControl COA_ColourSelector 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   ScaleHeight     =   300
   ScaleWidth      =   1755
   ToolboxBitmap   =   "COA_ColourSelector.ctx":0000
   Begin VB.CommandButton cmdSelect 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1425
      TabIndex        =   1
      ToolTipText     =   "Select Colour"
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1425
   End
End
Attribute VB_Name = "COA_ColourSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

' Declare Windows API Functions.
Private Declare Function GetSystemMetricsAPI Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

' Declare public events.
Public Event Click()

' Constant values.
Const gLngMinHeight = 200
Const gLngMinWidth = 200
Const SM_CXFRAME = 32

'Public Enum OLEType
'  OLE_LOCAL = 0
'  OLE_SERVER = 1
'  OLE_EMBEDDED = 2
'  OLE_UNC = 3
'End Enum

'Default Property Values:
'Const m_def_ASRDataField = 0
'Const m_def_ForeColor = 0

'Property Variables:
'Dim m_ASRDataField As Variant
'Dim m_ForeColor As Long

'Dim miOLEType As OLEType
'Dim mobjEmbeddedStream As ADODB.Stream

Private mlngColumnID As Long
Private gfSelected As Boolean
Private giControlLevel As Boolean
Private mblnReadOnly As Boolean
Private mlngTableID As Boolean

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Let TableID(New_Value As Long)
  mlngTableID = New_Value
End Property
Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Get Selected() As Boolean
  Selected = gfSelected
End Property

Public Property Let Selected(value As Boolean)
  gfSelected = value
End Property

Public Property Get Read_Only() As Boolean
  Read_Only = mblnReadOnly
End Property


Public Property Let Read_Only(blnValue As Boolean)
  mblnReadOnly = blnValue
End Property

Public Property Get ControlLevel() As Integer
  ' Return the control's level in the z-order.
  ControlLevel = giControlLevel
  
End Property

Public Property Let ControlLevel(ByVal piNewValue As Integer)
  ' Set the control's level in the z-order.
  giControlLevel = piNewValue
  
End Property




'Private mblnShowSelectionMarkers As Boolean
'Private mblnSelecting As Boolean

'Event Declarations:
'Event Click() 'MappingInfo=imgImage,imgImage,-1,Click
'Event Resize() 'MappingInfo=picPicture,picPicture,-1,Resize

'Event SpacePressed() ' RH 14/07/00 - To allow keypres

Public Property Get BackColor() As Long 'OLE_COLOR
  ' Return the back colour of the control
  BackColor = lblColour.BackColor
  
End Property
Public Property Let BackColor(ByVal NewColor As Long) 'OLE_COLOR)
  ' Set the back colour of the individual controls
  lblColour.BackColor = NewColor
  
End Property


Private Sub cmdSelect_Click()
  RaiseEvent Click
End Sub


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

'Public Property Let Selecting(ByVal Value As Boolean)
'
'  mblnSelecting = Value
'
'End Property
'
'Public Property Get Selecting() As Boolean
'
'  Selecting = mblnSelecting
'
'End Property
'
'Public Property Let ShowSelectionMarkers(ByVal Value As Boolean)
'
'  If Value = True Then
'    If Selecting = True Then
'      mblnShowSelectionMarkers = True
'    End If
'  Else
'    If Selecting = True Then
'      Value = True
'    End If
'  End If
'
'  mblnShowSelectionMarkers = Value
'
'  Line1.Visible = Value
'  Line2.Visible = Value
'  Line3.Visible = Value
'  Line4.Visible = Value
'
'End Property
'
'Public Property Get ShowSelectionMarkers() As Boolean
'
'  ShowSelectionMarkers = mblnShowSelectionMarkers
'
'End Property


''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,0
'Public Property Get ForeColor() As Long
'  ForeColor = m_ForeColor
'
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As Long)
'  m_ForeColor = New_ForeColor
'  PropertyChanged "ForeColor"
'
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgImage,imgImage,-1,Enabled
Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal value As Boolean)
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  'imgImage.Enabled = New_Enabled
  'picPicture.Enabled = New_Enabled
  UserControl.Enabled = value
  cmdSelect.Enabled = Not mblnReadOnly
  
  PropertyChanged "Enabled"
  
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=imgImage,imgImage,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'  BorderStyle = imgImage.BorderStyle
'
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'  imgImage.BorderStyle() = New_BorderStyle
'  PropertyChanged "BorderStyle"
'End Property

'Private Sub imgImage_Click()
'
'Selecting = True
'RaiseEvent Click
'
'End Sub
'
'Private Sub picPicture_KeyDown(KeyCode As Integer, Shift As Integer)
'
'  If KeyCode = vbKeySpace Then
'    Selecting = True
'    RaiseEvent SpacePressed
'  End If
'
'  KeyCode = 0
'  Shift = 0
'
'End Sub

'Private Sub picPicture_Resize()
'  RaiseEvent Resize
'
'End Sub

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=imgImage,imgImage,-1,Picture
'Public Property Get Picture() As Picture
'  Set Picture = imgImage.Picture
'
'End Property
'
'Public Property Set Picture(ByVal New_Picture As Picture)
'  Set imgImage.Picture = New_Picture
'  PropertyChanged "Picture"
'
'End Property
'
'Public Sub SetPicturePath(ByVal psNewValue As String)
'  ' Used by the intranet
'  On Error GoTo ErrorTrap
'
'  Set imgImage.Picture = LoadPicture(psNewValue)
'  PropertyChanged "Picture"
'
'  Exit Sub
'
'ErrorTrap:
'  Set imgImage.Picture = LoadPicture()
'  PropertyChanged "Picture"
'End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  'm_ForeColor = m_def_ForeColor
  'm_ASRDataField = m_def_ASRDataField
  BackColor = vbWhite
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Load property values from storage.
  'ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  Enabled = PropBag.ReadProperty("Enabled", True)
  'BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  'DMIAppearance = PropBag.ReadProperty("DMIAppearance", 1)
  'Set Picture = PropBag.ReadProperty("Picture", Nothing)
  'ASRDataField = PropBag.ReadProperty("ASRDataField", m_def_ASRDataField)
  ColumnID = PropBag.ReadProperty("ColumnID", 0)
  BackColor = PropBag.ReadProperty("BackColor", vbWhite)
  Selected = PropBag.ReadProperty("Selected", False)
  ControlLevel = PropBag.ReadProperty("ControlLevel", 0)
  Read_Only = PropBag.ReadProperty("ReadOnly", False)
  TableID = PropBag.ReadProperty("TableID", 0)
End Sub


Private Sub UserControl_Resize()
  If UserControl.Width < MinimumWidth Then UserControl.Width = MinimumWidth
  cmdSelect.Left = UserControl.Width - cmdSelect.Width
  lblColour.Width = cmdSelect.Left
End Sub


'Private Sub UserControl_Resize()
'  ' Resize the constituent controls.
'  picPicture.Height = UserControl.Height
'  imgImage.Height = picPicture.Height
'  picPicture.Width = UserControl.Width
'  imgImage.Width = picPicture.Width
'
'  ' Top horizontal
'  Line1.X1 = 50
'  Line1.X2 = (picPicture.ScaleWidth - 50)
'  Line1.Y1 = 50
'  Line1.Y2 = 50
'
'  ' Bottom horizontal
'  Line2.X1 = 50
'  Line2.X2 = (picPicture.ScaleWidth - 50)
'  Line2.Y1 = (picPicture.ScaleHeight - 50)
'  Line2.Y2 = (picPicture.ScaleHeight - 50)
'
'  ' Left Vertical
'  Line3.X1 = 50
'  Line3.X2 = 50
'  Line3.Y1 = 50
'  Line3.Y2 = (picPicture.ScaleHeight - 50)
'
'  ' Right Vertical
'  Line4.X1 = (picPicture.ScaleWidth - 50)
'  Line4.X2 = (picPicture.ScaleWidth - 50)
'  Line4.Y1 = 50
'  Line4.Y2 = (picPicture.ScaleHeight - 50)
'
'End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  ' Write property values to the property bag.
  'Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("ColumnID", ColumnID, 0)
  'Call PropBag.WriteProperty("BorderStyle", imgImage.BorderStyle, 0)
  'Call PropBag.WriteProperty("DMIAppearance", picPicture.Appearance, 1)
  'Call PropBag.WriteProperty("Picture", Picture, Nothing)
  'Call PropBag.WriteProperty("ASRDataField", m_ASRDataField, m_def_ASRDataField)
  
  Call PropBag.WriteProperty("BackColor", lblColour.BackColor, vbWhite)
  
  Call PropBag.WriteProperty("Selected", gfSelected, False)
  Call PropBag.WriteProperty("ControlLevel", giControlLevel, 0)
  Call PropBag.WriteProperty("ReadOnly", mblnReadOnly, False)
  Call PropBag.WriteProperty("TableID", mlngTableID, 0)

End Sub

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get ASRDataField() As Variant
'  ASRDataField = m_ASRDataField
'
'End Property

'Public Property Let ASRDataField(ByVal New_ASRDataField As Variant)
'  m_ASRDataField = New_ASRDataField
'  PropertyChanged "ASRDataField"
'
'End Property


'Public Property Let DMIAppearance(ByVal piNewValue As Integer)
'  picPicture.Appearance = piNewValue
'  PropertyChanged "DMIAppearance"
'End Property



'Public Property Let EmbeddedStream(ByRef pobjStream As ADODB.Stream)
'  Set mobjEmbeddedStream = pobjStream
'End Property
'
'Public Property Get EmbeddedStream() As ADODB.Stream
'  Set EmbeddedStream = mobjEmbeddedStream
'End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
End Property

Public Property Let ColumnID(plngNewValue As Long)
  mlngColumnID = plngNewValue
End Property

'Public Property Get OLEType() As OLEType
'  OLEType = miOLEType
'End Property
'
'Public Property Let OLEType(ByVal piNewValue As OLEType)
'  miOLEType = piNewValue
'End Property


'Private Sub lblColour_DblClick()
'  RaiseEvent DblClick
'End Sub
'
'Private Sub lblColour_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub
'
'Private Sub lblColour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub
'
'Private Sub lblColour_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseUp(Button, Shift, X, Y)
'End Sub
'
'Private Sub UserControl_DblClick()
'  RaiseEvent DblClick
'End Sub
'
'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'  RaiseEvent KeyDown(KeyCode, Shift)
'End Sub
'
'Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub
'
'Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub
'
'Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseUp(Button, Shift, X, Y)
'End Sub
'
'Private Sub cboComboBox_DblClick()
'  RaiseEvent DblClick
'End Sub
'
'Private Sub cboComboBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub
'
'Private Sub cboComboBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub
'
'Private Sub cboComboBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  RaiseEvent MouseUp(Button, Shift, X, Y)
'End Sub

Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  Dim lngMinHeight As Long

  lngMinHeight = UserControl.TextHeight("X") + (Screen.TwipsPerPixelY * 8)

  MinimumHeight = IIf(lngMinHeight < gLngMinHeight, gLngMinHeight, lngMinHeight)

End Property

Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  Dim lngMinWidth As Long
  
  lngMinWidth = (4 * GetSystemMetricsAPI(SM_CXFRAME) * Screen.TwipsPerPixelX) + _
    cmdSelect.Width + _
    UserControl.TextWidth("W")
    
  MinimumWidth = IIf(lngMinWidth < gLngMinWidth, gLngMinWidth, lngMinWidth)
  
End Property


