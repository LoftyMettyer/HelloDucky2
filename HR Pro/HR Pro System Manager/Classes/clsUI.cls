VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

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

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type

'Windows version constants
Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

'Button message constants
Const BM_SETSTYLE = &HF4

'Combo box message constants
Const CB_ERR = (-1)
Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
Const CB_SELECTSTRING = &H14D

'List box message constants
Const LB_ERR = (-1)
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2
Const LB_SELECTSTRING = &H18C
Const LB_SETTABSTOPS = &H192

'Window constants
Const HWND_TOPMOST = -1
Const hWnd_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Const SPI_GETWORKAREA = 48

Const LOCALE_SYSTEM_DEFAULT = &H800
Const LOCALE_USER_DEFAULT = &H400
'Const LOCALE_SLONGDATE = &H20        '  long date format string
Const LOCALE_SDATE = &H1D        '  date separator
Const LOCALE_STIME = &H1E        '  time separator
Const LOCALE_SSHORTDATE = &H1F        '  short date format string
'Const LOCALE_STIMEFORMAT = &H1003      '  time format string
'Const LOCALE_IDATE = &H21        '  short date format ordering
'Const LOCALE_ILDATE = &H22        '  long date format ordering
'Const LOCALE_ITIME = &H23        '  time format specifier
'Const LOCALE_ICENTURY = &H24        '  century format specifier
'Const LOCALE_ITLZERO = &H25        '  leading zeros in time field
'Const LOCALE_IDAYLZERO = &H26        '  leading zeros in day field
'Const LOCALE_IMONLZERO = &H27        '  leading zeros in month field
'Const LOCALE_S1159 = &H28        '  AM designator
'Const LOCALE_S2359 = &H29        '  PM designator
Const LOCALE_SDECIMAL = &HE         '  decimal separator
Const LOCALE_STHOUSAND = &HF         '  thousand separator
Const LOCALE_IMEASURE = &HD          ' Measurement System

'ChildWindowFromPointEx constants
Const CWP_ALL = 0
Const CWP_SKIPINVISIBLE = 1
Const CWP_SKIPDISABLED = 2
Const CWP_SKIPTRANSPARENT = 4

'Window style constants
Const GWL_STYLE = (-16)
Const WS_THICKFRAME = &H40000

Const MAX_COMPUTERNAME_LENGTH = 15

'Keyboard constants
Const KEYEVENTF_KEYUP = &H2


'Windows API functions
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long, ByVal Flags As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDialogBaseUnits Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetricsAPI Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Private Declare Function SendMessageLng Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long


Function GetSystemMeasurement() As String

  On Error GoTo ErrorTrap
  
  ' Return the system measurement (metric or us).
  
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, sBuffer, 99)
  GetSystemMeasurement = Left(sBuffer, lngLength - 1)
  
  If GetSystemMeasurement = 1 Then
    GetSystemMeasurement = "us"
  Else
    GetSystemMeasurement = "metric"
  End If
  
TidyUpAndExit:
  Exit Function
ErrorTrap:

End Function



Public Sub ControlsDisableAll(frmCurrent As Form)

  'Not all controls have a backcolor !
  On Local Error Resume Next

  Dim ctl As Control
  For Each ctl In frmCurrent

    If TypeOf ctl Is ComctlLib.TabStrip Or _
           TypeOf ctl Is SSTab Then
      'Stop
    
    ElseIf TypeOf ctl Is Frame Or _
           TypeOf ctl Is PictureBox Or _
           TypeOf ctl Is ListView Or _
           TypeOf ctl Is ListBox Then
      'Just make container controls and scroll-able controls look disabled...
      '(NOTE: Code must be placed in drag-drop events etc. to disable it)
      ctl.ForeColor = vbGrayText
      ctl.BackColor = vbButtonFace
    
    ElseIf TypeOf ctl Is SSDBGrid Then
      ctl.Enabled = False
      ctl.BackColorEven = vbButtonFace    'SSDBGrid
      ctl.BackColorOdd = vbButtonFace     'SSDBGrid
    
    ElseIf TypeOf ctl Is CommandButton Then
      'Disable all CommandButtons except cancel...

      If ctl.Cancel = False Then
        ctl.Enabled = False
      Else
        ctl.Enabled = True
      End If

    Else
      ctl.BackColor = vbButtonFace
      ctl.Enabled = False
      ctl.ForeColor = vbGrayText
      ctl.BackColor = vbButtonFace
    
    End If

  Next

End Sub

Function cboFind(ByVal hWnd As Long, ByVal FindStr As String, ByVal Exact As Boolean) As Integer
  cboFind = SendMessageStr(hWnd, IIf(Exact, CB_FINDSTRINGEXACT, CB_FINDSTRING), -1, FindStr)
End Function

Function cboSelect(ByVal hWnd As Long, ByVal SelectStr As String) As Integer
  cboSelect = SendMessageStr(hWnd, CB_SELECTSTRING, -1, SelectStr)
End Function

Sub frmAtCenter(ThisForm As Form)
  On Error Resume Next
  
  With ThisForm
    .Top = Int((Screen.Height - .Height) / 2)
    .Left = Int((Screen.Width - .Width) / 2)
  End With
  
End Sub
' Position the given child form in the centre of the given
' parent from.
Sub frmAtCenterOfParent(pChildForm As Form, pParentForm As Form)
  Dim iTop As Integer
  Dim iLeft As Integer
  
  On Error Resume Next
  
  With pChildForm
    iTop = Int((pParentForm.Height - .Height) / 2)
    iLeft = Int((pParentForm.Width - .Width) / 2)
    
    ' If the child form is too big then just position it at the
    ' origin of the parent form.
    .Top = pParentForm.Top + IIf(iTop < 0, 0, iTop)
    .Left = pParentForm.Left + IIf(iLeft < 0, 0, iLeft)
    
  End With
  
End Sub

Sub frmAtMouse(pFrmThisForm As Form)
  On Error Resume Next
  
  Dim lngMouseX As Long
  Dim lngMouseY As Long
  Dim lngSizeX As Long
  Dim lngSizeY As Long
  
  If Not GetWorkAreaSize(lngSizeX, lngSizeY) Then
    lngSizeX = Screen.Width
    lngSizeY = Screen.Height
  End If
  
  If GetMousePos(lngMouseX, lngMouseY) Then
    With pFrmThisForm
      If lngMouseY + .Height > lngSizeY Then
        .Top = lngSizeY - .Height
      Else
        .Top = lngMouseY
      End If
      If lngMouseX + .Width > lngSizeX Then
        .Left = lngSizeX - .Width
      Else
        .Left = lngMouseX
      End If
    End With
  End If

End Sub

Function frmIsLoaded(ByVal psFormName As String) As Boolean
  Dim iLoop As Integer
  
  For iLoop = 0 To Forms.Count - 1
    If UCase(Forms(iLoop).Name) = UCase(psFormName) Then
      frmIsLoaded = True
      Exit For
    End If
  Next iLoop
  
End Function

Function frmTopmost(ByVal pLngHWnd As Long, pfTopMost As Boolean) As Boolean
  Dim lngFlags As Long
  
  lngFlags = SWP_NOMOVE Or SWP_NOSIZE
  
  frmTopmost = (SetWindowPos(pLngHWnd, _
    IIf(pfTopMost, HWND_TOPMOST, hWnd_NOTOPMOST), _
    0, 0, 0, 0, lngFlags) <> 0)
    
End Function

Function lstDedup(pObjSourceList As Object) As Boolean
  On Error GoTo ErrorTrap
  
  Dim iIndex As Integer
  
  With pObjSourceList
  
    If .ListCount > 1 Then
      
      iIndex = 0
      While iIndex < .ListCount
        If .List(iIndex) = .List(iIndex + 1) Then
          .RemoveItem (iIndex)
        Else
          iIndex = iIndex + 1
        End If
      Wend
    
    End If
  
  End With

  lstDedup = True
  Exit Function
  
ErrorTrap:
  lstDedup = False
  Err = False
  
End Function

Function lstFind(ByVal pLngHWnd As Long, ByVal psFindStr As String, ByVal pfExact As Boolean) As Integer
  
  lstFind = SendMessageStr(pLngHWnd, IIf(pfExact, LB_FINDSTRINGEXACT, LB_FINDSTRING), -1, psFindStr)

End Function

Function lstSelect(ByVal pLngHWnd As Long, ByVal psSelectStr As String) As Integer
  
  lstSelect = SendMessageStr(pLngHWnd, LB_SELECTSTRING, -1, psSelectStr)

End Function

Function lstSetTabs(ByVal pLngHWnd As Long, ByVal psTabs As String) As Boolean
  Dim iStart As Integer
  Dim iEnd As Integer
  Dim iTabCount As Integer
  Dim lngTabs() As Long
  Dim lngNewTab As Long
  Dim lngReturn As Long

  iTabCount = 0
  
  If Len(Trim(psTabs)) > 0 Then
  
    iStart = 1
    
    Do
    
      iEnd = InStr(iStart, psTabs, ",")
      If iEnd > iStart + 1 Then
        lngNewTab = CInt(Mid(psTabs, iStart, iEnd - iStart))
      Else
        lngNewTab = CInt(Mid(psTabs, iStart))
      End If
      
      If lngNewTab > 0 Then
        iTabCount = iTabCount + 1
        ReDim Preserve lngTabs(iTabCount)
        lngTabs(iTabCount - 1) = lngNewTab
      End If
      
      iStart = iEnd + 1
      
    Loop While iEnd > 0
    
    If iTabCount > 0 Then
      lngReturn = SendMessageLng(pLngHWnd, LB_SETTABSTOPS, iTabCount, lngTabs(0))
      lstSetTabs = True
    End If
    
  End If

End Function

Function treMoveNext(pTrvTreeView As ComctlLib.TreeView) As Long
  On Error GoTo ErrorTrap
  
  Dim ThisNode As ComctlLib.Node
      
  With pTrvTreeView
    
    If .SelectedItem.Children > 0 And .SelectedItem.Expanded Then
      
      .SelectedItem = .SelectedItem.Child
    
    Else
      
      If .SelectedItem <> .SelectedItem.LastSibling Then
        
        .SelectedItem = .SelectedItem.Next
      
      Else
        
        Set ThisNode = .SelectedItem
        
        Do While ThisNode <> ThisNode.Root
          If ThisNode.Parent <> ThisNode.Parent.LastSibling Then
            .SelectedItem = ThisNode.Parent.Next
            Exit Do
          Else
            Set ThisNode = ThisNode.Parent
          End If
        Loop
        
        Set ThisNode = Nothing
      
      End If
    End If
    
    treMoveNext = .SelectedItem.Index
  End With
  
  Exit Function

ErrorTrap:
        
  Set ThisNode = Nothing
  treMoveNext = 0
  Err = False
  
End Function

Function treMovePrevious(pTrvTreeView As ComctlLib.TreeView) As Long
  On Error GoTo ErrorTrap
  
  Dim ThisNode As ComctlLib.Node
  
  With pTrvTreeView
  
    If .SelectedItem <> .SelectedItem.Root Then
    
      If .SelectedItem <> .SelectedItem.FirstSibling Then
      
        Set ThisNode = .SelectedItem.Previous
        
        Do While ThisNode.Children > 0 And ThisNode.Expanded
          Set ThisNode = ThisNode.Child.LastSibling
        Loop
        
        .SelectedItem = ThisNode
        Set ThisNode = Nothing
        
      Else
        .SelectedItem = .SelectedItem.Parent
      End If
    End If
    
    treMovePrevious = .SelectedItem.Index
  End With
  
  Exit Function

ErrorTrap:
  Set ThisNode = Nothing
  treMovePrevious = 0
  Err = False

End Function

Function txtSelText() As String
  ' Select all of the text in the current control.
  
  If Not Screen.ActiveForm.ActiveControl Is Nothing Then
    
    ' Only select the text in text or maskedit controls.
    If TypeOf Screen.ActiveForm.ActiveControl Is TextBox Then
      With Screen.ActiveForm.ActiveControl
        .SelStart = 0
        .SelLength = Len(.Text)
        txtSelText = .SelText
      End With
    End If
  
'    If TypeOf Screen.ActiveForm.ActiveControl Is MaskEdBox Then
'      With Screen.ActiveForm.ActiveControl
'        .SelStart = 0
'        .SelLength = Len(.FormattedText)
'        txtSelText = .SelText
'      End With
'    End If
  
  End If

End Function

Function GetAvgCharWidth(ByVal pLngHDC As Long) As Double
  Dim typTxtMetrics As TEXTMETRIC
  
  Call GetTextMetrics(pLngHDC, typTxtMetrics)
  
  GetAvgCharWidth = (typTxtMetrics.tmAveCharWidth * Screen.TwipsPerPixelX)

End Function
Function GetMaxCharWidth(ByVal pLngHDC As Long) As Double
  Dim typTxtMetrics As TEXTMETRIC
  
  Call GetTextMetrics(pLngHDC, typTxtMetrics)
  
  GetMaxCharWidth = (typTxtMetrics.tmMaxCharWidth * Screen.TwipsPerPixelX)

End Function

Function GetCaption(pObjControl As Object) As String
  On Error Resume Next
  
  'Attempt to set the Caption property
  GetCaption = pObjControl.Caption
  If Err Then
    'Attempt to set the Text property
    GetCaption = pObjControl.Text
  End If
  Err = False

End Function

Public Function CaptionHeight() As Double
  ' Return the height of a form's caption bar.
  CaptionHeight = GetSystemMetricsAPI(SM_CYSMCAPTION) * Screen.TwipsPerPixelY

End Function

Public Function YBorder() As Double
  ' Return the height of a control border.
  YBorder = GetSystemMetricsAPI(SM_CYBORDER) * Screen.TwipsPerPixelY

End Function

Function GetCharHeight(ByVal pLngHDC As Long) As Double
  Dim typTxtMetrics As TEXTMETRIC
  
  Call GetTextMetrics(pLngHDC, typTxtMetrics)
  
  GetCharHeight = (typTxtMetrics.tmHeight * Screen.TwipsPerPixelY)

End Function

Function GetDeskTopSize(pLngSizeX As Long, pLngSizeY As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim lngHWndDeskTop As Long
  Dim RectDeskTop As Rect
  
  lngHWndDeskTop = GetDesktopWindow()
  If GetClientRect(lngHWndDeskTop, RectDeskTop) Then
    pLngSizeX = (RectDeskTop.Right - RectDeskTop.Left) * Screen.TwipsPerPixelX
    pLngSizeY = (RectDeskTop.Bottom - RectDeskTop.Top) * Screen.TwipsPerPixelY
  End If
  
  GetDeskTopSize = True
  Exit Function

ErrorTrap:
  GetDeskTopSize = False
  Err = False

End Function

Function GetInverseColor(ByVal pLngColor As Long) As Long
  Dim iR As Integer
  Dim iG As Integer
  Dim iB As Integer
  
  If Int(pLngColor / (2 ^ 24)) <> 0 Then
    pLngColor = GetSysColor((pLngColor And ((2 ^ 24) - 1)))
  End If
  pLngColor = (((2 ^ 24) - 1) Xor pLngColor)
  
  GetRGB pLngColor, iR, iG, iB
  If Abs(128 - iR) < 30 And Abs(128 - iG) < 30 And Abs(128 - iB) < 30 Then
    GetInverseColor = RGB(255, 255, 255)
  Else
    GetInverseColor = pLngColor
  End If
  
End Function

Function GetMousePos(pLngMouseX As Long, pLngMouseY As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim MousePos As POINTAPI
  
  GetCursorPos MousePos
  pLngMouseX = MousePos.x * Screen.TwipsPerPixelX
  pLngMouseY = MousePos.y * Screen.TwipsPerPixelY

  GetMousePos = True
  Exit Function
  
ErrorTrap:
  GetMousePos = False
  Err = False

End Function

Function GetOSName() As String
  Dim typOSVer As OSVERSIONINFO
  Dim sOSName As String
  
  typOSVer.dwOSVersionInfoSize = Len(typOSVer)
  
  If GetVersionEx(typOSVer) Then
    Select Case typOSVer.dwPlatformId
      Case VER_PLATFORM_WIN32s
        sOSName = "Windows 32s"
      Case VER_PLATFORM_WIN32_WINDOWS
        sOSName = "Windows 95"
      Case VER_PLATFORM_WIN32_NT
        sOSName = "Windows NT"
    End Select
  Else
    sOSName = vbNullString
  End If

  GetOSName = sOSName
  
End Function

Function GetOSVersion() As Long
  Dim typOSVer As OSVERSIONINFO
  Dim sOSVersion As String
  
  typOSVer.dwOSVersionInfoSize = Len(typOSVer)
  
  If GetVersionEx(typOSVer) Then
    sOSVersion = Trim(Str(typOSVer.dwMajorVersion)) & "." & _
      Trim(Str(typOSVer.dwMinorVersion))
  Else
    sOSVersion = vbNullString
  End If

  GetOSVersion = sOSVersion
  
End Function

Function GetHostName() As String
  Dim sBuffer As String
  Dim lngSize As Long
  
  'MH20010205 Fault 1167
  'W95/W98 not getting Host Name.
  'Need to add one to max computer name length (don't ask!)

  'sBuffer = String(MAX_COMPUTERNAME_LENGTH, 0)
  'lngSize = MAX_COMPUTERNAME_LENGTH
  sBuffer = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
  lngSize = MAX_COMPUTERNAME_LENGTH + 1
  
  If GetComputerName(sBuffer, lngSize) Then
    GetHostName = Left(sBuffer, lngSize)
  Else
    GetHostName = vbNullString
  End If

End Function


Function GetRGB(ByVal plngRGBValue As Long, piRValue As Integer, piGValue As Integer, piBValue As Integer) As Boolean
  Dim lngRGB As Long
  Dim iRed As Integer
  Dim iGreen As Integer
  Dim iBlue As Integer
  
  lngRGB = (plngRGBValue And ((2 ^ 24) - 1))
  iBlue = Int(lngRGB / (2 ^ 16))
  iGreen = Int((lngRGB And ((2 ^ 16) - 1)) / (2 ^ 8))
  iRed = (lngRGB And ((2 ^ 8) - 1))
  
  piRValue = iRed
  piGValue = iGreen
  piBValue = iBlue
  
  GetRGB = True
  
End Function

' Converts RGB value into a hex code for IExplorer
Function GetHexColour(iRed As Integer, iGreen As Integer, iBlue As Integer) As String

  GetHexColour = "#" & Right("0" & Hex(iRed), 2) & Right("0" & Hex(iGreen), 2) & Right("0" & Hex(iBlue), 2)

End Function

Function SysColorToHex(ByVal plngRGBValue As Long)
  Dim iRed As Integer
  Dim iGreen As Integer
  Dim iBlue As Integer
  Dim fOK As Boolean

  fOK = GetRGB(plngRGBValue, iRed, iGreen, iBlue)
  SysColorToHex = UI.GetHexColour(iRed, iGreen, iBlue)

End Function

Function HexToSysColor(ByVal sHexColor As String) As String
  'The input at this point could be HexColor = "#00FF1F"
  Dim iRed As Integer
  Dim iGreen As Integer
  Dim iBlue As Integer
  
  sHexColor = Replace(sHexColor, "#", "")
  
  iRed = val("&H" & Mid(sHexColor, 1, 2))
  iGreen = val("&H" & Mid(sHexColor, 3, 2))
  iBlue = val("&H" & Mid(sHexColor, 5, 2))
  
  HexToSysColor = RGB(iRed, iGreen, iBlue)
      
End Function

Function GetSystemColor(ByVal pLngColorDef As Long) As Long
  
  GetSystemColor = GetSysColor(pLngColorDef)

End Function
Function GetSystemDateFormat() As String
  ' Return the system data format.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer, 99)
  GetSystemDateFormat = Left(sBuffer, lngLength - 1)
  
End Function

Function GetSystemDateSeparator() As String
  ' Return the system data separator.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, sBuffer, 99)
  GetSystemDateSeparator = Left(sBuffer, lngLength - 1)
  
End Function


Function GetSystemTimeSeparator() As String
  ' Return the system data separator.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIME, sBuffer, 99)
  GetSystemTimeSeparator = Left(sBuffer, lngLength - 1)
  
End Function


Function GetSystemDecimalSeparator() As String
  ' Return the system data separator.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, 99)
  GetSystemDecimalSeparator = Left(sBuffer, lngLength - 1)
  
End Function

Function GetSystemThousandSeparator() As String
  ' Return the system data separator.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sBuffer, 99)
  GetSystemThousandSeparator = Left(sBuffer, lngLength - 1)
  
End Function


Function GetSystemMetrics(ByVal pIndex As SystemMetrics) As Long
  
  GetSystemMetrics = GetSystemMetricsAPI(pIndex)

End Function

Function GetControlAtPoint(pCtlThisControl As Control) As Control
  Dim frmParentForm As Form
  Dim objParentControl As Object
  Dim WndRect As Rect
  Dim lngHWnd As Long
  Dim lngHWndFound As Long
  Dim iLoop As Integer
  Dim pt As POINTAPI
  Dim iX As Integer
  Dim iY As Integer
  
  lngHWnd = 0
  lngHWndFound = 0
  
  'Get current mouse position, relative to entire screen
  Call GetCursorPos(pt)
  
  'Check if this controls container is its parent form
  If pCtlThisControl.Container Is pCtlThisControl.Parent Then
    'Set pointer to controls parent
    Set objParentControl = pCtlThisControl.Parent
    
    'Get position of controls parent, relative to entire screen
    Call GetWindowRect(objParentControl.hWnd, WndRect)
    
    'Calculate current mouse position within the controls parent
    iX = pt.x - (WndRect.Left + GetSystemMetrics(SM_CXFRAME))
    iY = pt.y - (WndRect.Top + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYFRAME))
    
  Else
    'Set pointer to controls container
    Set objParentControl = pCtlThisControl.Container
    
    'Get position of controls container, relative to entire screen
    Call GetWindowRect(objParentControl.hWnd, WndRect)
    
    'Calculate current mouse position within the controls container
    iX = pt.x - (WndRect.Left + GetSystemMetrics(SM_CXBORDER))
    iY = pt.y - (WndRect.Top + GetSystemMetrics(SM_CYBORDER))
  End If
  
  'Attempt to find a visible child window at mouse position
  lngHWnd = ChildWindowFromPointEx(objParentControl.hWnd, iX, iY, CWP_SKIPINVISIBLE)
  
  ' Disassociate object variable.
  Set objParentControl = Nothing
  
  'Check if a visible child window was found
  Do While lngHWnd > 0 And lngHWnd <> lngHWndFound
    'Save handle of found window
    lngHWndFound = lngHWnd
    
    'Get position of found window, relative to entire screen
    Call GetWindowRect(lngHWndFound, WndRect)
    
    'Calculate mouse position within the found window
    iX = pt.x - (WndRect.Left + GetSystemMetrics(SM_CXBORDER))
    iY = pt.y - (WndRect.Top + GetSystemMetrics(SM_CYBORDER))
  
    'Attempt to find a visible child window at mouse position
    lngHWnd = ChildWindowFromPointEx(lngHWndFound, iX, iY, CWP_SKIPINVISIBLE)
  Loop
  
  'Check if a visible child window was found
  If lngHWndFound > 0 Then
    'Set pointer to parent form
    Set frmParentForm = pCtlThisControl.Parent
    
    On Error Resume Next
    
    'Loop through parent forms controls
    For iLoop = 0 To frmParentForm.Controls.Count - 1
      'Check is this controls handle is the found windows handle
      If frmParentForm.Controls(iLoop).hWnd = lngHWndFound Then
        If Err = 0 Then
          'Return a pointer to the found control
          Set GetControlAtPoint = frmParentForm.Controls(iLoop)
          Exit For
        Else
          Err = False
        End If
      End If
    Next iLoop
    
    ' Disassocaite object variable.
    Set frmParentForm = Nothing
    
  End If
  
End Function

Public Function YFrame() As Double
  ' Return the height of a control frame.
  YFrame = GetSystemMetricsAPI(SM_CYFRAME) * Screen.TwipsPerPixelY

End Function

Function GetWorkAreaSize(pLngSizeX As Long, pLngSizeY As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim RectWorkArea As Rect

  If SystemParametersInfo(SPI_GETWORKAREA, vbEmpty, RectWorkArea, vbEmpty) Then
    pLngSizeX = (RectWorkArea.Right - RectWorkArea.Left) * Screen.TwipsPerPixelX
    pLngSizeY = (RectWorkArea.Bottom - RectWorkArea.Top) * Screen.TwipsPerPixelY
  End If
  
  GetWorkAreaSize = True
  Exit Function

ErrorTrap:
  GetWorkAreaSize = False
  Err = False

End Function

Function LockWindow(ByVal pLngHWnd As Long) As Boolean

  'Unlock any window currently locked
  UnlockWindow
  
  'Lock required window
  LockWindow = LockWindowUpdate(pLngHWnd)

End Function

Function GetChildWindowHWnd(ByVal pLngSourceHWnd As Long, ByVal pLngCommand As Long) As Long
  ' Return the child/sibling window of the source window as defined by the Command.
  ' Return 0 if no window is found.
  GetChildWindowHWnd = GetWindow(pLngSourceHWnd, pLngCommand)

End Function


Sub SendKeys(ByVal psSendString As String, Optional ByVal pLngHWnd As Long)
  Dim iChar As Integer
  Dim iAscii As Integer
  Dim iVKey As Integer
  Dim iOEMScan As Integer
  Dim sOEMChar As String
  
  If Len(psSendString) < 1 Then
    Exit Sub
  End If
  
  If pLngHWnd > 0 Then
    'Set focus to required window
    Call SetFocusAPI(pLngHWnd)
  End If
  
  For iChar = 1 To Len(psSendString)
    iAscii = Asc(Mid(psSendString, iChar, 1))
    
    'Get the virtual key code for this character
    iVKey = VkKeyScan(iAscii) And &HFF
    
    'Get the OEM character
    sOEMChar = Space(2)
    CharToOem Chr(iAscii), sOEMChar
    
    'Get the OEM scan code
    iOEMScan = OemKeyScan(Asc(sOEMChar)) And &HFF
    
    'Send the down key
    keybd_event iVKey, iOEMScan, 0, 0
    
    'Send the up key
    keybd_event iVKey, iOEMScan, KEYEVENTF_KEYUP, 0
  Next iChar
  
End Sub

Public Sub SetCaption(pObjControl As Object, ByVal psCaption As String)
  On Error Resume Next
  
  'Attempt to set the Caption property
  pObjControl.Caption = psCaption
  
  If Err Then
    'Attempt to set the Text property
    pObjControl.Text = psCaption
  End If
  
  Err = False

End Sub

Sub SetFocus(ByVal pLngHWnd As Long)

  'Set focus to required window
  Call SetFocusAPI(pLngHWnd)

End Sub

Public Function SetThickFrame(ByVal pLngHWnd As Long, ByVal pfThickFrame As Boolean)
  Dim lngWndStyle As Long
  
  'Get window style info
  lngWndStyle = GetWindowLong(pLngHWnd, GWL_STYLE)
  
  'Change window style info
  If pfThickFrame Then
    lngWndStyle = (lngWndStyle Or WS_THICKFRAME)
  Else
    lngWndStyle = (lngWndStyle Xor WS_THICKFRAME)
  End If
  
  'Update window style
  If SetWindowLong(pLngHWnd, GWL_STYLE, lngWndStyle) <> 0 Then
    SetThickFrame = True
  End If

End Function

Public Function UnlockWindow() As Boolean

  'Unlock any window currently locked
  UnlockWindow = LockWindowUpdate(0&)

End Function


Public Function ConvertNumberForSQL(ByVal strInput As String) As String
  'Get a number in the correct format for a SQL string
  '(e.g. on french systems replace decimal comma for a decimal point)
  ConvertNumberForSQL = Replace(strInput, UI.GetSystemDecimalSeparator, ".")
End Function

Public Function ConvertNumberForDisplay(ByVal strInput As String) As String
  'Get a number in the correct format for display
  '(e.g. on french systems replace decimal point for a decimal comma)
  ConvertNumberForDisplay = Replace(strInput, ".", UI.GetSystemDecimalSeparator)
End Function

Public Function ConvertDateLocaleToSQL(psLocaleDate As String) As String
  Dim sSQLDate As String
  
  sSQLDate = Format(psLocaleDate, "mm/dd/yyyy")
  sSQLDate = Replace(sSQLDate, UI.GetSystemDateSeparator, "/")

  ConvertDateLocaleToSQL = sSQLDate

End Function

Public Sub FormatTDBNumberControl(objInput As Object)
  
  'MH20010130
  If TypeOf objInput Is TDBNumber6Ctl.TDBNumber Then
    objInput.Separator = "x"
    objInput.DecimalPoint = UI.GetSystemDecimalSeparator
    objInput.Separator = UI.GetSystemThousandSeparator
  
  ElseIf (TypeOf objInput Is TDBNumberCtrl.TDBNumber) Then
    'Need to set to dummy incase new thousand is same as old decimal (cases error)
    objInput.SepThousand = "x"
    objInput.SepDecimal = UI.GetSystemDecimalSeparator
    objInput.SepThousand = UI.GetSystemThousandSeparator
  
  End If

End Sub


Public Sub FormatGTDateControl(objInput As Object)
  Dim sFormat As String
  
  If TypeOf objInput Is GTMaskDate.GTMaskDate Then
    sFormat = DateFormat
    sFormat = Replace(UCase(sFormat), "D", "_")
    sFormat = Replace(UCase(sFormat), "M", "_")
    sFormat = Replace(UCase(sFormat), "Y", "_")
    objInput.NullText = sFormat
  End If

End Sub

' Center the mouse over the control.
Public Sub CenterMouseOverControl(ByVal ctl As Control, frm As Form)
On Error GoTo CenterErr

Dim pt As POINTAPI

    ' Get the control's upper left position
    ' in screen coordinates.
    ' Note that this function converts the values
    ' in pt into screen coordinates. Because we just
    ' allocated pt, the values start as 0.
    ClientToScreen ctl.hWnd, pt

    ' Position the cursor. SetCursorPos uses screen
    ' coordinates (pixels).
    SetCursorPos _
        pt.x + frm.ScaleX(ctl.Width / 2, frm.ScaleMode, vbPixels), _
        pt.y + frm.ScaleY(ctl.Height / 2, frm.ScaleMode, vbPixels)
        
  Exit Sub
  
CenterErr:
  
End Sub


