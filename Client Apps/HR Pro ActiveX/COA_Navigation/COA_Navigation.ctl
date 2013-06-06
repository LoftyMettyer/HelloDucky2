VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.UserControl COA_Navigation 
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2010
   ScaleWidth      =   2985
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   1305
      Width           =   1005
      ExtentX         =   1773
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox picHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1800
      Picture         =   "COA_Navigation.ctx":0000
      ScaleHeight     =   480
      ScaleWidth      =   660
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox picBrowser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1620
      ScaleHeight     =   570
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   1125
      Width           =   870
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1275
      _Version        =   851969
      _ExtentX        =   2249
      _ExtentY        =   1138
      _StockProps     =   79
      Caption         =   "Navigate..."
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblHyperlink 
      Caption         =   "http://"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   990
      Width           =   2400
   End
End
Attribute VB_Name = "COA_Navigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

Public Event ToolClickRequest(ByVal Tool As String)
Public Event DBExecuteRequest(ByVal SQL As String)

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Public Enum enum_DisplayType
  Hyperlink = 0
  Button = 1
  Browser = 2
  Hidden = 3
End Enum

Public Enum enum_NavigateIn
  URL = 0
  MenuBar = 1
  DB = 2
End Enum

Private miDisplayType As enum_DisplayType
Private miNavigateIn As enum_NavigateIn
Private mstrNavigateTo As String
Private mstrCaption As String
Private mlngColumnID As Long
Private msColumnName As String
Private mbSelected As Boolean
Private mbEnabled As Boolean
Private miControlLevel As Integer
Private mForeColor As OLE_COLOR
Private mbNavigateOnSave As Boolean
Private mBackColor As OLE_COLOR

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

Private Sub lblHyperlink_Click()
  Navigate mstrNavigateTo
End Sub

Private Sub PushButton1_Click()

  If miNavigateIn = MenuBar Then
    RaiseEvent ToolClickRequest(NavigateTo)
  ElseIf miNavigateIn = DB Then
    RaiseEvent DBExecuteRequest(NavigateTo)
  Else
    Navigate mstrNavigateTo
  End If

End Sub

Private Sub UserControl_Initialize()
  miDisplayType = enum_DisplayType.Button
  mstrCaption = "Navigate..."
  gbInScreenDesigner = True
  mlngColumnID = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the keydown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
  RefreshControls
End Sub

Public Sub RefreshControls()

  Dim sDisplayText As String
  
  sDisplayText = IIf(mstrCaption = vbNullString, mstrNavigateTo, mstrCaption)

  ' Button properties
  PushButton1.Font = UserControl.Font
  PushButton1.Font.Bold = UserControl.Font.Bold
  PushButton1.Font.Size = UserControl.Font.Size
  PushButton1.Font.Strikethrough = UserControl.Font.Strikethrough
  PushButton1.Font.Underline = UserControl.Font.Underline
  PushButton1.Width = UserControl.Width
  PushButton1.Height = UserControl.Height
  PushButton1.Caption = sDisplayText
  PushButton1.ToolTipText = mstrNavigateTo
  PushButton1.ForeColor = mForeColor
  PushButton1.Enabled = mbEnabled
  PushButton1.Visible = (miDisplayType = enum_DisplayType.Button)

  ' Hyperlink properties
  lblHyperlink.Font = UserControl.Font
  lblHyperlink.Font.Bold = UserControl.Font.Bold
  lblHyperlink.Font.Size = UserControl.Font.Size
  lblHyperlink.Font.Underline = True
  lblHyperlink.ForeColor = vbBlue
  lblHyperlink.Top = 10
  lblHyperlink.Left = 10
  lblHyperlink.Width = UserControl.TextWidth(sDisplayText)
  lblHyperlink.Height = UserControl.TextHeight(sDisplayText)
  lblHyperlink.Caption = sDisplayText
  lblHyperlink.ToolTipText = mstrNavigateTo
  lblHyperlink.ForeColor = mForeColor
  lblHyperlink.Enabled = mbEnabled
  lblHyperlink.Visible = (miDisplayType = enum_DisplayType.Hyperlink)
  lblHyperlink.BackColor = mBackColor

  ' Hidden control properties
  picHidden.Top = 0
  picHidden.Left = 0
  picHidden.Width = UserControl.Width
  picHidden.Height = UserControl.Height
  picHidden.Visible = (miDisplayType = enum_DisplayType.Hidden) And gbInScreenDesigner

  ' Dummy browser properties
  picBrowser.Top = 0
  picBrowser.Left = 0
  picBrowser.Width = UserControl.Width
  picBrowser.Height = UserControl.Height
  picBrowser.Visible = (miDisplayType = enum_DisplayType.Browser) And gbInScreenDesigner

  ' Browser properties
  WebBrowser1.Top = 0
  WebBrowser1.Left = 0
  WebBrowser1.Width = UserControl.Width
  WebBrowser1.Height = UserControl.Height
  WebBrowser1.Offline = Not mbEnabled
  WebBrowser1.Visible = (miDisplayType = enum_DisplayType.Browser) And Not gbInScreenDesigner

  If miDisplayType = Browser And gbInScreenDesigner = False Then
    If Len(mstrNavigateTo) = 0 Then
      WebBrowser1.Navigate "about:blank"
    Else
      WebBrowser1.Navigate mstrNavigateTo
    End If
  End If

  UserControl.BackColor = mBackColor

End Sub

Public Property Let InScreenDesigner(ByVal NewValue As Boolean)
  gbInScreenDesigner = NewValue
End Property

Public Property Get InScreenDesigner() As Boolean
  InScreenDesigner = gbInScreenDesigner
End Property

Public Property Get DisplayType() As enum_DisplayType
  DisplayType = miDisplayType
End Property

Public Property Let DisplayType(ByVal NewValue As enum_DisplayType)

  If miDisplayType <> NewValue Then
    If NewValue = Hyperlink Then mForeColor = vbBlue
    If NewValue = Button Then mForeColor = vbBlack
  End If

  miDisplayType = NewValue
  RefreshControls
End Property

Public Property Get NavigateIn() As enum_NavigateIn
  NavigateIn = miNavigateIn
End Property

Public Property Let NavigateIn(ByVal NewValue As enum_NavigateIn)
  miNavigateIn = NewValue
End Property

Public Property Get NavigateTo() As String
  If gbInScreenDesigner And mlngColumnID > 0 Then
    NavigateTo = "http://<<" & msColumnName & ">>"
  Else
    NavigateTo = mstrNavigateTo
  End If
End Property

Public Property Let NavigateTo(ByVal NewValue As String)
  mstrNavigateTo = NewValue
  RefreshControls
'  DoEvents
End Property

Public Property Let NavigateOnSave(ByVal NewValue As Boolean)
  mbNavigateOnSave = NewValue
End Property

Public Property Get NavigateOnSave() As Boolean
  NavigateOnSave = mbNavigateOnSave
End Property

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  mstrCaption = PropBag.ReadProperty("Caption", "Navigate...")
  miDisplayType = PropBag.ReadProperty("DisplayType", enum_DisplayType.Button)
  miNavigateIn = PropBag.ReadProperty("NavigateIn", enum_NavigateIn.URL)
  mstrNavigateTo = PropBag.ReadProperty("NavigateTo", "about:blank")
  gbInScreenDesigner = PropBag.ReadProperty("InScreenDesigner", True)
  mlngColumnID = PropBag.ReadProperty("ColumnID", 0)
  msColumnName = PropBag.ReadProperty("ColumnName", vbNullString)
  mbSelected = PropBag.ReadProperty("Selected", False)
  mbEnabled = PropBag.ReadProperty("Enabled", True)
  mbNavigateOnSave = PropBag.ReadProperty("NavigateOnSave", False)

  UserControl.Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
  UserControl.Font.Bold = PropBag.ReadProperty("FontBold", UserControl.Ambient.Font.Bold)
  UserControl.Font.Size = PropBag.ReadProperty("FontSize", UserControl.Ambient.Font.Size)
  UserControl.Font.Strikethrough = PropBag.ReadProperty("FontStrikethrough", UserControl.Ambient.Font.Strikethrough)
  UserControl.Font.Underline = PropBag.ReadProperty("FontUnderline", UserControl.Ambient.Font.Underline)
  mForeColor = PropBag.ReadProperty("ForeColor", UserControl.Ambient.ForeColor)
  mBackColor = PropBag.ReadProperty("BackColor", UserControl.Ambient.BackColor)

  RefreshControls
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Caption", mstrCaption)
  Call PropBag.WriteProperty("DisplayType", miDisplayType)
  Call PropBag.WriteProperty("NavigateIn", miNavigateIn)
  Call PropBag.WriteProperty("NavigateTo", mstrNavigateTo)
  Call PropBag.WriteProperty("InScreenDesigner", gbInScreenDesigner)
  Call PropBag.WriteProperty("ColumnID", mlngColumnID)
  Call PropBag.WriteProperty("ColumnName", msColumnName)
  Call PropBag.WriteProperty("Selected", mbSelected)
  Call PropBag.WriteProperty("Enabled", mbEnabled)

  Call PropBag.WriteProperty("Font", lblHyperlink.Font)
  Call PropBag.WriteProperty("FontBold", lblHyperlink.Font.Bold)
  Call PropBag.WriteProperty("FontSize", lblHyperlink.Font.Size)
  Call PropBag.WriteProperty("FontStrikethrough", lblHyperlink.Font.Strikethrough)
  Call PropBag.WriteProperty("FontUnderline", lblHyperlink.Font.Underline)
  Call PropBag.WriteProperty("ForeColor", mForeColor)
  Call PropBag.WriteProperty("BackColor", mBackColor)
  Call PropBag.WriteProperty("NavigateOnSave", NavigateOnSave)

End Sub


' ----------------------------------------
' Mouse Handling Events
' ----------------------------------------
Private Sub picBrowser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub picBrowser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseMove(Button, Shift, X, Y)
  End If
End Sub

Private Sub picBrowser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseUp(Button, Shift, X, Y)
  End If
End Sub

Private Sub picHidden_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub picHidden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseMove(Button, Shift, X, Y)
  End If
End Sub

Private Sub picHidden_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseUp(Button, Shift, X, Y)
  End If
End Sub

Private Sub lblHyperlink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub lblHyperlink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseMove(Button, Shift, X, Y)
  Else
    SetCursor LoadCursor(0, IDC_HAND)
  End If
End Sub

Private Sub lblHyperlink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseUp(Button, Shift, X, Y)
  End If
End Sub

Private Sub PushButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub PushButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseMove(Button, Shift, X, Y)
  End If
End Sub

Private Sub PushButton1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If gbInScreenDesigner Then
    RaiseEvent MouseUp(Button, Shift, X, Y)
  End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Screen.MousePointer = vbDefault
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

' ----------------------------------------
' Standard control properties
' ----------------------------------------
Public Property Get Font() As Font
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  UserControl.Font.Bold = New_Font.Bold
  UserControl.Font.Size = New_Font.Size
  UserControl.Font.Strikethrough = New_Font.Strikethrough
  UserControl.Font.Underline = New_Font.Underline
  RefreshControls
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  mForeColor = New_ForeColor
  RefreshControls
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
End Property

Public Property Let ColumnID(ByVal NewValue As Long)
  mlngColumnID = NewValue
End Property

Public Property Get ColumnName() As String
  ColumnName = msColumnName
End Property

Public Property Let ColumnName(ByVal NewValue As String)
  msColumnName = NewValue
End Property

Public Property Get Selected() As Boolean
  Selected = mbSelected
End Property

Public Property Let Selected(ByVal NewValue As Boolean)
  mbSelected = NewValue
End Property

Public Property Get ControlLevel() As Integer
  ControlLevel = miControlLevel
End Property

Public Property Let ControlLevel(ByVal NewValue As Integer)
  miControlLevel = NewValue
End Property

Public Property Get Caption() As String
  Caption = mstrCaption
End Property

Public Property Let Caption(ByVal NewValue As String)
  mstrCaption = NewValue
  RefreshControls
'  DoEvents
End Property

Public Property Get Enabled() As Boolean
  Enabled = mbEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  mbEnabled = NewValue
  RefreshControls
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
  mBackColor = NewColor
  RefreshControls
End Property

Public Sub ExecutePostSave()

  If mbNavigateOnSave Then
    PushButton1_Click
  End If

End Sub

Public Property Get MinimumHeight() As Long
  MinimumHeight = UserControl.TextHeight("X")
End Property

Public Property Get MinimumWidth() As Long
  MinimumWidth = UserControl.TextWidth("W")
End Property
