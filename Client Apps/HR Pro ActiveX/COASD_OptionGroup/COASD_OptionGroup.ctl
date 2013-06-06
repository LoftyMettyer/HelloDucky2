VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.UserControl COASD_OptionGroup 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   LockControls    =   -1  'True
   ScaleHeight     =   1695
   ScaleWidth      =   3285
   Begin XtremeSuiteControls.GroupBox fraOptGroup 
      Height          =   1050
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   2670
      _Version        =   851969
      _ExtentX        =   4710
      _ExtentY        =   1852
      _StockProps     =   79
      Caption         =   "No Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.Frame fraInternal 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Enabled         =   0   'False
         Height          =   465
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   1590
         Begin XtremeSuiteControls.RadioButton Option1 
            Height          =   330
            Index           =   0
            Left            =   45
            TabIndex        =   2
            Top             =   45
            Width           =   1500
            _Version        =   851969
            _ExtentX        =   2646
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "RadioButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "COASD_OptionGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declare control events
' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Click()
Public Event DblClick()

'Declare windows API types
Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type

Public Enum WorkflowValueTypes
  giWFVALUE_UNKNOWN = -1
  giWFVALUE_FIXED = 0
  giWFVALUE_WFVALUE = 1
  giWFVALUE_DBVALUE = 2
  giWFVALUE_CALC = 3
End Enum

'Declare windows API function
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long

'Declare local variables
Private InResize As Boolean

Private mvarMaxLength As Integer

' For new property (Alignment)
Private miAlignment As Integer
Const gLngMinHeight = 300
Const gLngMinWidth = 300
Private giControlLevel As Integer
Private gLngColumnID As Long
Private gfSelected As Boolean

Private msWFIdentifier As String
Private miWFItemType As Integer
Private msWFDefaultValue As String
Private mlngCalculationID As Long
Private miDefaultValueType As WorkflowValueTypes

Private mBackcolour As OLE_COLOR
Private mForecolour As OLE_COLOR

Public Enum ASRBackStyleConstants
  BACKSTYLE_TRANSPARENT = 0
  BACKSTYLE_OPAQUE = 1
End Enum

Private miBackStyle As ASRBackStyleConstants

Private msDefaultStringValue As String

Private mfNoOptions As Boolean

Private mblnReadOnly As Boolean  'NPG20071022

Public Property Get CalculationID() As Long
  CalculationID = mlngCalculationID
  
End Property
Public Property Let CalculationID(ByVal plngNewValue As Long)
  mlngCalculationID = plngNewValue
  
End Property

Public Property Get DefaultValueType() As WorkflowValueTypes
  DefaultValueType = miDefaultValueType
  
End Property

Public Property Let DefaultValueType(ByVal piNewValue As WorkflowValueTypes)
  miDefaultValueType = piNewValue
  
End Property

Public Sub ClearOptions()
  
  Do Until (Option1.Count = 1)
    Unload Option1(Option1.UBound)
  Loop
  
  Option1(0).Visible = False
  
End Sub

Public Property Let DefaultStringValue(New_Value As String)
  msDefaultStringValue = New_Value
End Property
Public Property Get DefaultStringValue() As String
  DefaultStringValue = msDefaultStringValue
End Property

Public Property Get ControlValueList() As String
  'Returns Tab delimited list of control values from the option group control array.
  Dim ctl As RadioButton
  Dim s As String
  
  For Each ctl In Option1
    s = s & ctl.Caption & vbTab
  Next ctl
  
  If Len(s) > 0 Then
    ControlValueList = Left(s, Len(s) - 1)
  End If
  
End Property

Public Property Get BackStyle() As ASRBackStyleConstants
  BackStyle = miBackStyle
End Property
Public Property Let BackStyle(ByVal New_BackStyle As ASRBackStyleConstants)
  ' NB. This property is NOT applied to the UserControl and contained control(s)
  miBackStyle = New_BackStyle
  PropertyChanged "BackStyle"
End Property

Public Property Get NoOptions() As Boolean
  NoOptions = mfNoOptions
End Property
Public Property Let NoOptions(New_Value As Boolean)
  If New_Value Then
    Me.ClearOptions
  End If
  mfNoOptions = New_Value
End Property

Public Function SelectOption(psValue As String) As Boolean

  Dim ctl As RadioButton
  
  For Each ctl In Option1
    If (psValue = ctl.Caption) Then
      ctl.Value = True
      SelectOption = True
    Else
      ctl.Value = False
    End If
  Next ctl
  
End Function

Public Property Let WFDefaultValue(New_Value As String)
  msWFDefaultValue = New_Value
End Property
Public Property Get WFDefaultValue() As String
 WFDefaultValue = Option1(GetSelectedOption()).Value
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

Private Sub fraInternal_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub fraOptGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub fraInternal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub Option1_DblClick(Index As Integer)
  RaiseEvent DblClick

End Sub

Private Sub Option1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub Option1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick

End Sub

Private Sub UserControl_Initialize()
  mForecolour = vbBlack
  mBackcolour = vbButtonFace
End Sub

Private Sub UserControl_InitProperties()
  On Error Resume Next
  Caption = Extender.Name
  Set Font = Ambient.Font
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  miBackStyle = PropBag.ReadProperty("BackStyle", BACKSTYLE_OPAQUE)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackStyle", miBackStyle, BACKSTYLE_OPAQUE)
  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
End Sub

Private Sub UserControl_Resize()

  Dim Index As Integer
  Dim intHeight As Integer, intWidth As Integer
  
Select Case miAlignment

  Case 0: 'Vertical
    If Not InResize Then
      InResize = True
      If BorderStyle = 0 Then
        intHeight = UserControl.TextHeight(Caption) * 0.5
        If intHeight < 200 Then intHeight = 200
        intWidth = 0
      Else
        intHeight = UserControl.TextHeight(Caption) * 1.5
        If intHeight < 400 Then intHeight = 400
        intWidth = UserControl.TextWidth(Caption) + 100
      End If
      For Index = Option1.LBound To Option1.UBound
        With Option1(Index)
          
          '.Width = 285 + UserControl.TextWidth(.Caption) * 1.25
          .Width = 350 + UserControl.TextWidth(.Caption) * 1.25
          
          .Height = IIf(UserControl.TextHeight(.Caption) < 240, _
            240, UserControl.TextHeight(.Caption))
            
          If Index = 0 Then
            fraInternal.Top = (.Height * (Index + BorderStyle)) + (.Height * 0.25)
            fraInternal.Left = GetAvgCharWidth(UserControl.hDC) * 2
            .Top = 0
          Else
            .Top = Option1(Index - 1).Top + UserControl.TextHeight(.Caption) + 50 '20
          End If
          
          .Left = 0
            
          intHeight = intHeight + .Height
          If .Width > intWidth Then intWidth = .Width
        
        End With
      Next Index
      
      intWidth = intWidth + (GetAvgCharWidth(UserControl.hDC) * 2)
      
      intHeight = Option1(Option1.UBound).Top + (Option1(Option1.UBound).Height * 2) + (fraInternal.Top / 2)
      
      With UserControl
        .Height = intHeight
        .Width = intWidth
      End With
      
      fraOptGroup.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
      fraInternal.Width = intWidth - GetAvgCharWidth(UserControl.hDC) * 3
      fraInternal.Height = Option1(Option1.UBound).Top + (Option1(Option1.UBound).Height)
      InResize = False
    
    End If

  Case 1: 'Horizonal
    If Not InResize Then
      InResize = True
      If BorderStyle = 0 Then
        intHeight = UserControl.TextHeight(Caption) * 0.5
        If intHeight < 200 Then intHeight = 200
        intWidth = 0
      Else
        intHeight = UserControl.TextHeight(Caption) * 1.5
        If intHeight < 400 Then intHeight = 400
        intWidth = UserControl.TextWidth(Caption)
      End If
      Dim temp As Integer
      For Index = Option1.LBound To Option1.UBound
        With Option1(Index)
          .Width = 285 + UserControl.TextWidth(.Caption) + GetAvgCharWidth(UserControl.hDC)
          .Height = IIf(UserControl.TextHeight(.Caption) < 240, _
            240, UserControl.TextHeight(.Caption))
          
          If Index = 0 Then
            fraInternal.Top = (.Height * (Index + BorderStyle)) + (.Height * 0.25)
            fraInternal.Left = GetAvgCharWidth(UserControl.hDC) * 2
            .Top = 0
          Else
            .Top = Option1(0).Top
          End If
          
          If Index > 0 Then
            .Left = Option1(Index - 1).Left + Option1(Index - 1).Width + UserControl.TextWidth("WW")
          End If
          
          intWidth = intWidth + .Width
          
        End With
      Next Index
      
      intHeight = (fraInternal.Top * 1.5) + Option1(Option1.UBound).Height
      
      If BorderStyle = 1 Then
        intWidth = (Option1(Option1.UBound).Left) + Option1(Option1.UBound).Width + UserControl.TextWidth("WW")
        If intWidth < (UserControl.TextWidth(Caption) + UserControl.TextWidth("WW")) Then intWidth = UserControl.TextWidth(Caption) + UserControl.TextWidth("WW")
      Else
        intWidth = (Option1(Option1.UBound).Left) + Option1(Option1.UBound).Width + UserControl.TextWidth("W")
      End If
      
      With UserControl
        .Height = intHeight
        ' AE20080519 Fault #13022
        '.Width = intWidth
        .Width = intWidth + UserControl.TextWidth("W")
      End With

      fraOptGroup.Width = UserControl.Width
      fraOptGroup.Height = UserControl.Height
      fraInternal.Width = (Option1(Option1.UBound).Left) + Option1(Option1.UBound).Width
      fraInternal.Height = Option1(Option1.UBound).Height
      
      InResize = False
    End If

End Select

End Sub

Public Property Get Alignment() As Integer
  
  Alignment = miAlignment

End Property

Public Property Let Alignment(ByVal vNewValue As Integer)

  If vNewValue = 0 Or vNewValue = 1 Then
    miAlignment = vNewValue
    UserControl_Resize
  End If
  
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = mBackcolour
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
  Dim Index As Integer
  
  mBackcolour = NewColor
  
  If Not mblnReadOnly Then
  
    fraOptGroup.BackColor = mBackcolour
    For Index = Option1.LBound To Option1.UBound
      Option1(Index).BackColor = mBackcolour
    Next Index
    
    'JDM - 16/08/01 - Fault 2694 - Not setting background colour
    fraInternal.BackColor = mBackcolour
    
  End If
  
End Property

Public Property Get BorderStyle() As Integer
  BorderStyle = IIf(fraOptGroup.BorderStyle = xtpFrameBorder, 1, 0)
End Property

Public Property Let BorderStyle(ByVal NewValue As Integer)
  fraOptGroup.BorderStyle = IIf(NewValue = 1, xtpFrameBorder, xtpFrameNone)
  UserControl_Resize
End Property

Public Property Get Caption() As String
  Caption = fraOptGroup.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
  fraOptGroup.Caption = NewCaption
  '
  UserControl_Resize
  '
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
  Dim ctlTemp As Control
  
  UserControl.Enabled() = NewEnabled
  If Not NewEnabled Then
    For Each ctlTemp In UserControl.Controls
        If TypeOf ctlTemp Is RadioButton Then
            ctlTemp.Enabled = False
        End If
    Next
  End If
  
End Property

Public Property Get Font() As Font
  Set Font = fraOptGroup.Font
End Property

Public Property Set Font(ByVal NewFont As Font)

  Dim Index As Integer
  
  Set UserControl.Font = NewFont
  
  Set fraOptGroup.Font = UserControl.Font
  
  For Index = Option1.LBound To Option1.UBound
    Set Option1(Index).Font = UserControl.Font
  Next Index
  
  UserControl_Resize

  PropertyChanged "Font"

End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = mForecolour
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
  Dim Index As Integer
  
  mForecolour = NewColor
  
  If Not mblnReadOnly Then
    fraOptGroup.ForeColor = mForecolour
    For Index = Option1.LBound To Option1.UBound
      Option1(Index).ForeColor = mForecolour
    Next Index
  End If

End Property

Public Property Get hWnd() As Long
  hWnd = fraOptGroup.hWnd
End Property

Public Property Get MaxLength() As Integer
  MaxLength = mvarMaxLength
End Property

Public Property Let MaxLength(intMaxLen As Integer)
  mvarMaxLength = IIf(intMaxLen > 0, intMaxLen, 0)
End Property

Public Function SetOptions(ByRef pasOptions As Variant)
  Dim iX As Integer
  Dim iIndex As Integer
  Dim iArrayDim As Integer
  Dim iTemp As Integer

  iIndex = 0
  ' Decide if we have a one or 2 dimension array
  iArrayDim = 2
  On Error GoTo err_SetOptions
  iTemp = UBound(pasOptions, 2) > 0
  If iArrayDim = 2 Then
    For iX = LBound(pasOptions, 2) To UBound(pasOptions, 2)
      If Option1.Count - 1 < iX Then
        Load Option1(iX)
      End If
      With Option1(iIndex)
        .Caption = pasOptions(0, iX)
        .Visible = True
      End With
      iIndex = iIndex + 1
    Next iX
  Else
    For iX = LBound(pasOptions, 1) To UBound(pasOptions, 1)
      If Option1.Count - 1 < iX Then
        Load Option1(iX)
      End If
      With Option1(iIndex)
        .Caption = pasOptions(iX)
        .Visible = True
      End With
      iIndex = iIndex + 1
    Next iX
  End If
  UserControl_Resize
Exit Function

err_SetOptions:
  If Err.Number = 9 Then
    iArrayDim = 1
    Resume Next
  Else
    Err.Raise Err.Number, "SetOptions", Err.Description
  End If
End Function

Public Property Get Text() As String
  Dim Index As Integer
  
  Index = GetSelectedOption
  If Index >= 0 Then
    If MaxLength > 0 Then
      Text = Left(Option1(Index).Caption, MaxLength)
    Else
      Text = Option1(Index).Caption
    End If
  Else
    Text = vbNullString
  End If
End Property

Public Property Let Text(ByVal NewValue As String)
  Dim Index As Integer

  NewValue = UCase(Trim(NewValue))
  For Index = Option1.LBound To Option1.UBound
    If UCase(Left(Option1(Index).Caption, Len(NewValue))) = NewValue Then
      Option1(Index).Value = True
      Exit For
    Else
      Option1(Index).Value = False
    End If
  Next Index
  
End Property

Public Property Get Value() As Integer
  Value = GetSelectedOption
End Property

Public Property Let Value(ByVal NewValue As Integer)
  Dim Index As Integer

  For Index = Option1.LBound To Option1.UBound
    Option1(Index).Value = (Index = NewValue)
  Next Index
  
End Property

Public Sub Refresh()
  UserControl_Resize
  UserControl.Refresh
End Sub

Private Function GetSelectedOption() As Integer
  Dim Index As Integer
  Dim Selected As Integer
  
  Selected = -1
  For Index = Option1.LBound To Option1.UBound
    If Option1(Index).Value = True Then
      Selected = Index
      Exit For
    End If
  Next Index

  GetSelectedOption = Selected
End Function

Private Function GetAvgCharWidth(ByVal hDC As Long) As Integer
  Dim typTxtMetrics As TEXTMETRIC
  
  Call GetTextMetrics(hDC, typTxtMetrics)
  GetAvgCharWidth = (typTxtMetrics.tmAveCharWidth * Screen.TwipsPerPixelX)
End Function

Public Sub About()
  MsgBox App.ProductName & " - " & App.FileDescription & _
    vbCr & vbCr & App.LegalCopyright, _
    vbOKOnly, "About " & App.ProductName
End Sub



Private Sub fraOptGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub fraOptGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


Private Sub fraInternal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub fraInternal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


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

Public Property Get MinimumWidth() As Long
  ' Return the minimum height of the control.
  MinimumWidth = gLngMinWidth
  
End Property

Public Property Get MinimumHeight() As Long
  ' Return the minimum height of the control.
  MinimumHeight = gLngMinHeight
  
End Property

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

Public Property Get Selected() As Boolean
  ' Return the Selected property.
  Selected = gfSelected
End Property

Public Property Let Selected(ByVal pfNewValue As Boolean)
  ' Set the Selected property.
  gfSelected = pfNewValue
End Property

Public Property Get Read_Only() As Boolean
'NPG20071022
  Read_Only = mblnReadOnly

End Property

Public Property Let Read_Only(blnValue As Boolean)
  
  Dim lngIndex As Long

  mblnReadOnly = blnValue

  For lngIndex = Option1.LBound To Option1.UBound
    Option1(lngIndex).ForeColor = IIf(blnValue, vbGrayText, mForecolour)
    Option1(lngIndex).BackColor = IIf(blnValue, vbButtonFace, mBackcolour)
    fraOptGroup.ForeColor = IIf(blnValue, vbGrayText, mForecolour)
    fraOptGroup.BackColor = IIf(blnValue, vbButtonFace, mBackcolour)
    
    UserControl.BackColor = IIf(blnValue, vbButtonFace, mBackcolour)
  Next

End Property
