VERSION 5.00
Begin VB.UserControl COASD_Label 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label lblLabel 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "COASD_Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DblClick()

' Declare Windows API Functions.
Private Declare Function GetSystemMetricsAPI Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long



Public Enum WorkflowValueTypes
  giWFVALUE_UNKNOWN = -1
  giWFVALUE_FIXED = 0
  giWFVALUE_WFVALUE = 1
  giWFVALUE_DBVALUE = 2
  giWFVALUE_CALC = 3
End Enum

' Constant values.
Const gLngMinHeight = 180
Const gLngMinWidth = 180

' Properties.
Private gLngColumnID As Long
Private giControlLevel As Integer
Private gfSelected As Boolean
Private msWFIdentifier As String
Private miWFItemType As Integer
Private mlngWFDatabaseRecord As Long
Private msWFWorkflowForm As String
Private msWFWorkflowValue As String
Private mdblWFDefaultNumericValue As Double
Private msWFDefaultCharValue As String
Private mlngWFInputSize As Long
Private mlngWFInputDecimals As Long
Private mfMandatory As Boolean
Private mblnReadOnly As Boolean   'NPG20071023

Private mlngCalculationID As Long
Private miCaptionType As WorkflowValueTypes
Private miDefaultValueType As WorkflowValueTypes

Private mForecolour As OLE_COLOR
Private mBackcolour As OLE_COLOR

Public Enum ASRBackStyleConstants
  BACKSTYLE_TRANSPARENT = 0
  BACKSTYLE_OPAQUE = 1
End Enum

Private miBackStyle As ASRBackStyleConstants
Private mfPasswordType As Boolean

Public Property Get BackStyle() As ASRBackStyleConstants
  BackStyle = miBackStyle

End Property

Public Property Let BackStyle(ByVal New_BackStyle As ASRBackStyleConstants)
  ' NB. This property is NOT applied to the UserControl and contained control(s)
  ' See MSDN Q179052 for the reason.
  miBackStyle = New_BackStyle
  PropertyChanged "BackStyle"
  
End Property




Public Property Let WFInputDecimals(New_Value As Long)
  mlngWFInputDecimals = New_Value
End Property
Public Property Get WFInputDecimals() As Long
  WFInputDecimals = mlngWFInputDecimals
End Property

Public Property Let WFInputSize(New_Value As Long)
  mlngWFInputSize = New_Value
End Property
Public Property Get WFInputSize() As Long
  WFInputSize = mlngWFInputSize
End Property

Public Property Let WFDefaultNumericValue(New_Value As Double)
  mdblWFDefaultNumericValue = New_Value
End Property
Public Property Get WFDefaultNumericValue() As Double
  WFDefaultNumericValue = mdblWFDefaultNumericValue
End Property

Public Property Let WFDefaultCharValue(New_Value As String)
  msWFDefaultCharValue = New_Value
End Property
Public Property Get WFDefaultCharValue() As String
  WFDefaultCharValue = msWFDefaultCharValue
End Property

Public Property Let WFWorkflowForm(New_Value As String)
  msWFWorkflowForm = New_Value
End Property
Public Property Get WFWorkflowForm() As String
  WFWorkflowForm = msWFWorkflowForm
End Property
Public Property Let WFWorkflowValue(New_Value As String)
  msWFWorkflowValue = New_Value
End Property
Public Property Get WFWorkflowValue() As String
  WFWorkflowValue = msWFWorkflowValue
End Property

Public Property Let WFDatabaseRecord(New_Value As Long)
  mlngWFDatabaseRecord = New_Value
End Property
Public Property Get WFDatabaseRecord() As Long
  WFDatabaseRecord = mlngWFDatabaseRecord
End Property

Public Property Let WFIdentifier(New_Value As String)
  msWFIdentifier = New_Value
End Property
Public Property Get WFIdentifier() As String
  WFIdentifier = msWFIdentifier
End Property

Public Property Let WFItemType(New_Value As Integer)
  miWFItemType = New_Value
End Property
Public Property Get WFItemType() As Integer
  WFItemType = miWFItemType
End Property

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    pdwSupportedOptions As Long, _
                                                    pdwEnabledOptions As Long)

    Dim rc      As Long
    Dim rClsId  As udtGUID
    Dim IID     As String
    Dim bIID()  As Byte

    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or _
                          INTERFACESAFE_FOR_UNTRUSTED_DATA

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), rc)

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
    Dim rc          As Long
    Dim rClsId      As udtGUID
    Dim IID         As String
    Dim bIID()      As Byte

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), rc)

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

Public Property Get Selected() As Boolean
  ' Return the Selected property.
  Selected = gfSelected
  
End Property

Public Property Let Selected(ByVal pfNewValue As Boolean)
  ' Set the Selected property.
  gfSelected = pfNewValue
    
End Property

Public Property Get ControlLevel() As Integer
  ' Return the control's level in the z-order.
  ControlLevel = giControlLevel
  
End Property

Public Property Let ControlLevel(ByVal piNewValue As Integer)
  ' Set the control's level in the z-order.
  giControlLevel = piNewValue
  
End Property


Public Property Get ColumnID() As Long
  ' Return the control's column ID.
  ColumnID = gLngColumnID
  
End Property

Public Property Let ColumnID(ByVal plngNewValue As Long)
  ' Set the control's column ID.
  gLngColumnID = plngNewValue
  
End Property

Public Sub About()
  ' Display the About information.
  With App
    MsgBox .ProductName & " - " & .FileDescription & _
      vbCr & vbCr & .LegalCopyright, _
      vbOKOnly, "About " & .ProductName
  End With
  
End Sub
'Public Property Get Caption() As String
'  ' Return the Caption property.
'  Caption = lblLabel.Caption
'
'End Property
'
'Public Property Let Caption(ByVal psNewValue As String)
'  ' Set the Caption property if it has changed.
'  lblLabel.Caption = psNewValue
'
'End Property
'
'Public Property Get Font() As Font
'  ' Return the control's font property.
'  Set Font = UserControl.Font
'
'End Property
'
'Public Property Set Font(ByVal pObjNewValue As StdFont)
'  ' Set the control's font property.
'  Dim iLoop As Integer
'
'  ' Update the sub-controls.
'  Set UserControl.Font = pObjNewValue
'  Set lblLabel.Font = pObjNewValue
'
'  UserControl_Resize
'
'End Property
'
'Public Property Get ForeColor() As OLE_COLOR
'  ' Return the control's foreground colour property.
'  ForeColor = UserControl.ForeColor
'
'End Property
'
'Public Property Let ForeColor(ByVal pColNewColor As OLE_COLOR)
'  ' Set the control's foreground colour property.
'  UserControl.ForeColor = pColNewColor
'  lblLabel.ForeColor = pColNewColor
'
'End Property
Public Property Get hWnd() As Long
  ' Return the control's hWnd.
  hWnd = UserControl.hWnd
  
End Property
Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  Dim lngMinHeight As Long
  
  lngMinHeight = UserControl.TextHeight("X")
  If lblLabel.BorderStyle = vbFixedSingle Then
    lngMinHeight = lngMinHeight + (8 * Screen.TwipsPerPixelY)
  End If
          
  MinimumHeight = IIf(lngMinHeight < gLngMinHeight, gLngMinHeight, lngMinHeight)
  
End Property
Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  Dim lngMinWidth As Long
  
  lngMinWidth = UserControl.TextWidth("mn") / 2
  If lblLabel.BorderStyle = vbFixedSingle Then
    lngMinWidth = lngMinWidth + (8 * Screen.TwipsPerPixelX)
  End If
  
  MinimumWidth = IIf(lngMinWidth < gLngMinWidth, gLngMinWidth, lngMinWidth)
  
End Property
'
'
'Public Property Get BackColor() As OLE_COLOR
'  ' Return the control's background colour property.
'  BackColor = UserControl.BackColor
'
'End Property
'
'Public Property Let BackColor(ByVal pColNewColor As OLE_COLOR)
'  ' Set the control's background colour property.
'  UserControl.BackColor = pColNewColor
'  lblLabel.BackColor = pColNewColor
'
'End Property

Private Sub lblLabel_DblClick()
 RaiseEvent DblClick

End Sub

Private Sub lblLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub lblLabel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_DblClick()
 RaiseEvent DblClick

End Sub

Private Sub UserControl_Initialize()
  mBackcolour = vbWhite
  mForecolour = vbBlack
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
  With lblLabel
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
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
  BackColor = mBackcolour
End Property

Public Property Get Alignment() As AlignmentConstants
  Alignment = lblLabel.Alignment
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

  If Not mblnReadOnly Then
    lblLabel.BackColor() = New_BackColor
  End If
  
  mBackcolour = New_BackColor
  PropertyChanged "BackColor"

End Property

Public Property Let Alignment(ByVal pNewValue As AlignmentConstants)
  lblLabel.Alignment = pNewValue
  PropertyChanged "Alignment"
  
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
  ForeColor = mForecolour
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

  If Not mblnReadOnly Then
    lblLabel.ForeColor() = New_ForeColor
  End If
  
  mForecolour = New_ForeColor
  PropertyChanged "ForeColor"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,Font
Public Property Get Font() As Font
    Set Font = lblLabel.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set lblLabel.Font = New_Font
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
  UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,BorderStyle
Public Property Get BorderStyle() As Integer
    BorderStyle = lblLabel.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    lblLabel.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,Caption
Public Property Get Caption() As String
    Caption = lblLabel.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblLabel.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = lblLabel.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblLabel.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = lblLabel.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblLabel.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = lblLabel.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblLabel.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = lblLabel.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lblLabel.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLabel,lblLabel,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = lblLabel.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lblLabel.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblLabel.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    lblLabel.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set lblLabel.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblLabel.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    lblLabel.Caption = PropBag.ReadProperty("Caption", "")
    lblLabel.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblLabel.FontItalic = PropBag.ReadProperty("FontItalic", 0)
'    lblLabel.FontSize = PropBag.ReadProperty("FontSize", 0)
    lblLabel.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    lblLabel.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    lblLabel.Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)

    miBackStyle = PropBag.ReadProperty("BackStyle", BACKSTYLE_OPAQUE)
    mfPasswordType = PropBag.ReadProperty("PasswordType", False)

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", lblLabel.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", lblLabel.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", lblLabel.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", lblLabel.BorderStyle, 1)
    Call PropBag.WriteProperty("Caption", lblLabel.Caption, "")
    Call PropBag.WriteProperty("FontBold", lblLabel.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblLabel.FontItalic, 0)
    Call PropBag.WriteProperty("FontSize", lblLabel.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblLabel.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", lblLabel.FontUnderline, 0)
    Call PropBag.WriteProperty("Alignment", lblLabel.Alignment, vbLeftJustify)

    Call PropBag.WriteProperty("BackStyle", miBackStyle, BACKSTYLE_OPAQUE)
    Call PropBag.WriteProperty("PasswordType", mfPasswordType, False)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Public Property Get Mandatory() As Boolean
  Mandatory = mfMandatory
  
End Property

Public Property Let Mandatory(ByVal pfNewValue As Boolean)
  mfMandatory = pfNewValue
  
End Property

Public Property Get CalculationID() As Long
  CalculationID = mlngCalculationID
  
End Property

Public Property Let CalculationID(ByVal plngNewValue As Long)
  mlngCalculationID = plngNewValue
  
End Property

Public Property Get CaptionType() As WorkflowValueTypes
  CaptionType = miCaptionType
  
End Property
Public Property Get DefaultValueType() As WorkflowValueTypes
  DefaultValueType = miDefaultValueType
  
End Property

Public Property Let DefaultValueType(ByVal piNewValue As WorkflowValueTypes)
  miDefaultValueType = piNewValue
  
End Property
Public Property Let CaptionType(ByVal piNewValue As WorkflowValueTypes)
  miCaptionType = piNewValue
  
End Property

Public Property Get Read_Only() As Boolean
'NPG20071023
  Read_Only = mblnReadOnly

End Property


Public Property Let Read_Only(blnValue As Boolean)
'NPG20071023
  mblnReadOnly = blnValue
  
  With lblLabel
    .BackColor = IIf(blnValue, vbButtonFace, mBackcolour)
    .ForeColor = IIf(blnValue, vbGrayText, mForecolour)
  End With
End Property

Public Property Get PasswordType() As Boolean
  PasswordType = mfPasswordType
  
End Property

Public Property Let PasswordType(ByVal pfNewValue As Boolean)
  mfPasswordType = pfNewValue
  PropertyChanged "PasswordType"
  
End Property


