VERSION 5.00
Begin VB.UserControl COA_ColourPicker 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   Picture         =   "COA_ColourPicker.ctx":0000
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   31
   ToolboxBitmap   =   "COA_ColourPicker.ctx":030A
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "COA_ColourPicker.ctx":0404
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "COA_ColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

'Declare Public Events
Public Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Public Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Public Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Public Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize

'Enums
Public Enum cpAppearanceConstants
    Flat
    [3D]
End Enum

'API function & constant declarations
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'Module specific variable declarations
Private RClr As RECT
Private RBut As RECT

Private IsInFocus As Boolean
Private IsButDown As Boolean

'Default Property Values:
Private Const m_def_ShowToolTips = True
Private Const m_def_ShowSysColorButton = True
Private Const m_def_ShowDefault = True
Private Const m_def_ShowCustomColors = True
Private Const m_def_ShowMoreColors = True
Private Const m_def_DefaultCaption = "Default"
Private Const m_def_MoreColorsCaption = "Custom Colours..."
Private Const m_def_BackColor = &H8000000C
Private Const m_def_Appearance = cpAppearanceConstants.[3D]
Private Const m_def_Color = &HFFFFFF
Private Const m_def_DefaultColor = &HFFFFFF

'Property Variables:
Private m_ShowToolTips As Boolean
Private m_ShowSysColorButton    As Boolean
Private m_ShowDefault           As Boolean
Private m_ShowCustomColors      As Boolean
Private m_ShowMoreColors        As Boolean
Private m_DefaultCaption        As String
Private m_MoreColorsCaption     As String
Private m_BackColor             As OLE_COLOR
Private m_Appearance            As cpAppearanceConstants
Private m_Color                 As OLE_COLOR
Private m_DefaultColor          As OLE_COLOR


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    
    If Button = 1 Then
        If (X >= RBut.Left And X <= RBut.Right) And (Y >= RBut.Top And Y <= RBut.Bottom) Then
            IsButDown = True
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    
    If IsButDown Then
        If Not ((X >= RBut.Left And X <= RBut.Right) And (Y >= RBut.Top And Y <= RBut.Bottom)) Then
            IsButDown = False
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    
    If Button = 1 And IsButDown Then
            IsButDown = False
    End If
End Sub

Public Sub ShowPalette()
    Dim ClrCtrlPos As RECT
    Dim Pt As POINTAPI
    
    Call GetWindowRect(hwnd, ClrCtrlPos)
    
    DefClr = m_DefaultColor
    CurClr = m_Color
    
    DefCap = m_DefaultCaption
    MorCap = m_MoreColorsCaption
    
    ShwDef = m_ShowDefault
    ShwMor = m_ShowMoreColors
    ShwCus = m_ShowCustomColors
    ShwSys = m_ShowSysColorButton
    ShwTip = m_ShowToolTips

    Load frmColorPalette
    With frmColorPalette
        Call GetCursorPos(Pt)
        
        .Top = (Pt.Y * Screen.TwipsPerPixelY) - .Height / 2
        .Left = (Pt.X * Screen.TwipsPerPixelX) - .Width / 2
        
        If (.Top + .Height) > Screen.Height Then
            .Top = (Pt.Y * Screen.TwipsPerPixelY) - .Height
        ElseIf (.Top < 0) Then
            .Top = 0
        End If
        
        If (.Left + .Width) > Screen.Width Then
          .Left = (Pt.X * Screen.TwipsPerPixelX) - .Width
        ElseIf (.Left < 0) Then
          .Left = 0
        End If
        
        .SelectedColor = m_Color
       .Show vbModal
        
        If Not .IsCanceled Then m_Color = .SelectedColor
    
    End With
    Unload frmColorPalette
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    ' Display the About information.
  With App
    MsgBox .ProductName & " - " & .FileDescription & _
      vbCr & vbCr & .LegalCopyright, _
      vbOKOnly, "About " & .ProductName
  End With
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_DefaultColor = m_def_DefaultColor
    m_Color = m_def_Color
    m_Appearance = m_def_Appearance
    m_BackColor = m_def_BackColor
    m_ShowDefault = m_def_ShowDefault
    m_ShowCustomColors = m_def_ShowCustomColors
    m_ShowMoreColors = m_def_ShowMoreColors
    m_DefaultCaption = m_def_DefaultCaption
    m_MoreColorsCaption = m_def_MoreColorsCaption
    m_ShowSysColorButton = m_def_ShowSysColorButton
    m_ShowToolTips = m_def_ShowToolTips
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_DefaultColor = PropBag.ReadProperty("DefaultColor", m_def_DefaultColor)
    m_Color = PropBag.ReadProperty("Value", m_def_Color)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ShowDefault = PropBag.ReadProperty("ShowDefault", m_def_ShowDefault)
    m_ShowCustomColors = PropBag.ReadProperty("ShowCustomColors", m_def_ShowCustomColors)
    m_ShowMoreColors = PropBag.ReadProperty("ShowMoreColors", m_def_ShowMoreColors)
    m_DefaultCaption = PropBag.ReadProperty("DefaultCaption", m_def_DefaultCaption)
    m_MoreColorsCaption = PropBag.ReadProperty("MoreColorsCaption", m_def_MoreColorsCaption)
    m_ShowSysColorButton = PropBag.ReadProperty("ShowSysColorButton", m_def_ShowSysColorButton)
    m_ShowToolTips = PropBag.ReadProperty("ShowToolTips", m_def_ShowToolTips)
End Sub

Private Sub UserControl_Resize()
  If Height <> 465 Then Height = 465
  If Width <> 465 Then Width = 465
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DefaultColor", m_DefaultColor, m_def_DefaultColor)
    Call PropBag.WriteProperty("Value", m_Color, m_def_Color)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ShowDefault", m_ShowDefault, m_def_ShowDefault)
    Call PropBag.WriteProperty("ShowCustomColors", m_ShowCustomColors, m_def_ShowCustomColors)
    Call PropBag.WriteProperty("ShowMoreColors", m_ShowMoreColors, m_def_ShowMoreColors)
    Call PropBag.WriteProperty("DefaultCaption", m_DefaultCaption, m_def_DefaultCaption)
    Call PropBag.WriteProperty("MoreColorsCaption", m_MoreColorsCaption, m_def_MoreColorsCaption)
    Call PropBag.WriteProperty("ShowSysColorButton", m_ShowSysColorButton, m_def_ShowSysColorButton)
    Call PropBag.WriteProperty("ShowToolTips", m_ShowToolTips, m_def_ShowToolTips)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get DefaultColor() As OLE_COLOR
Attribute DefaultColor.VB_Description = "Returns/Sets  the default color"
Attribute DefaultColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DefaultColor = m_DefaultColor
End Property

Public Property Let DefaultColor(ByVal New_DefaultColor As OLE_COLOR)
    m_DefaultColor = New_DefaultColor
    PropertyChanged "DefaultColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/Sets the selected color"
Attribute Color.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Color.VB_UserMemId = 0
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As OLE_COLOR)
    m_Color = New_Color
    PropertyChanged "Color"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,cpAppearanceConstants.[3D]
Public Property Get Appearance() As cpAppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As cpAppearanceConstants)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000C&
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowDefault() As Boolean
Attribute ShowDefault.VB_Description = "Returns/Sets whether default button will be shown or not"
Attribute ShowDefault.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowDefault = m_ShowDefault
End Property

Public Property Let ShowDefault(ByVal New_ShowDefault As Boolean)
    m_ShowDefault = New_ShowDefault
    PropertyChanged "ShowDefault"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowCustomColors() As Boolean
Attribute ShowCustomColors.VB_Description = "Returns/Sets whether custom colors will be shown or not"
Attribute ShowCustomColors.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowCustomColors = m_ShowCustomColors
End Property

Public Property Let ShowCustomColors(ByVal New_ShowCustomColors As Boolean)
    m_ShowCustomColors = New_ShowCustomColors
    PropertyChanged "ShowCustomColors"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowMoreColors() As Boolean
Attribute ShowMoreColors.VB_Description = "Returns/Sets whether More Colors button will be shown or not"
Attribute ShowMoreColors.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowMoreColors = m_ShowMoreColors
End Property

Public Property Let ShowMoreColors(ByVal New_ShowMoreColors As Boolean)
    m_ShowMoreColors = New_ShowMoreColors
    PropertyChanged "ShowMoreColors"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Default
Public Property Get DefaultCaption() As String
Attribute DefaultCaption.VB_Description = "Returns/Sets the caption in default button"
Attribute DefaultCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DefaultCaption = m_DefaultCaption
End Property

Public Property Let DefaultCaption(ByVal New_DefaultCaption As String)
    m_DefaultCaption = New_DefaultCaption
    PropertyChanged "DefaultCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,More Colors...
Public Property Get MoreColorsCaption() As String
Attribute MoreColorsCaption.VB_Description = "Returns/Sets the caption in the More button"
Attribute MoreColorsCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MoreColorsCaption = m_MoreColorsCaption
End Property

Public Property Let MoreColorsCaption(ByVal New_MoreColorsCaption As String)
    m_MoreColorsCaption = New_MoreColorsCaption
    PropertyChanged "MoreColorsCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowSysColorButton() As Boolean
Attribute ShowSysColorButton.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowSysColorButton = m_ShowSysColorButton
End Property

Public Property Let ShowSysColorButton(ByVal New_ShowSysColorButton As Boolean)
    m_ShowSysColorButton = New_ShowSysColorButton
    PropertyChanged "ShowSysColorButton"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowToolTips() As Boolean
Attribute ShowToolTips.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowToolTips = m_ShowToolTips
End Property

Public Property Let ShowToolTips(ByVal New_ShowToolTips As Boolean)
    m_ShowToolTips = New_ShowToolTips
    PropertyChanged "ShowToolTips"
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


