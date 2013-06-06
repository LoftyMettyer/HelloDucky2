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

Private retVAL As Variant
Private DeskhWnd As Variant

Private frmHeight As Long
Private frmWidth As Long

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



Function RedrawScreen(ByVal hWnd As Long) As Boolean
  ' Redraw the required window
  Dim WndRect As Rect
  
  Call GetWindowRect(hWnd, WndRect)
  
  RedrawScreen = RedrawWindow(hWnd, WndRect, 0, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)

End Function



Public Function CaptionHeight() As Double
  ' Return the height of a form's caption bar.
  CaptionHeight = GetSystemMetricsAPI(SM_CYSMCAPTION) * Screen.TwipsPerPixelY

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

Public Function GetSystemDateFormat() As String
  ' Return the system data format.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer, 99)
  GetSystemDateFormat = Left(sBuffer, lngLength - 1)
  
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

Function GetSystemMeasurement() As String

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "UI.GetSystemMeasurement()"
  
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
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function


Function cboFind(ByVal hWnd As Long, ByVal FindStr As String, ByVal Exact As Boolean) As Integer
  cboFind = SendMessage(hWnd, IIf(Exact, CB_FINDSTRINGEXACT, CB_FINDSTRING), -1, FindStr)
End Function

Function cboSelect(ByVal pobjControl As Control, ByVal psSelectStr As String) As Integer
  'JPD 20050921 Fault 10386
  Dim iLoop As Integer

  If TypeOf pobjControl Is XtremeSuiteControls.ComboBox Then
    ' First try and do an exact match.
    For iLoop = 0 To pobjControl.ListCount - 1
      If pobjControl.List(iLoop) = psSelectStr Then
        cboSelect = iLoop
        Exit Function
      End If
    Next iLoop

    ' No exact match, do things the old way.
    cboSelect = SendMessage(pobjControl.hWnd, CB_SELECTSTRING, -1, psSelectStr)
  End If

End Function

Function cboDropDown(ByVal hWnd As Long, bDropDown As Boolean) As Integer

  ' RH 31/07/00 - Automatically drop down a combo box.
  cboDropDown = SendMessage(hWnd, &H14F, bDropDown, "")

End Function

Sub frmAtCenter(ThisForm As Form)
  On Error Resume Next
  
  With ThisForm
    .Top = Int((Screen.Height - .Height) / 2)
    .Left = Int((Screen.Width - .Width) / 2)
  End With
  
End Sub

Sub frmAtMouse(ThisForm As Form)
  On Error Resume Next
  
  Dim MouseX As Long, MouseY As Long
  Dim SizeX As Long, SizeY As Long
  
  If Not GetWorkAreaSize(SizeX, SizeY) Then
    SizeX = Screen.Width
    SizeY = Screen.Height
  End If
  
  If GetMousePos(MouseX, MouseY) Then
    With ThisForm
      If MouseY + .Height > SizeY Then
        .Top = SizeY - .Height
      Else
        .Top = MouseY
      End If
      If MouseX + .Width > SizeX Then
        .Left = SizeX - .Width
      Else
        .Left = MouseX
      End If
    End With
  End If

End Sub

Function frmIsLoaded(ByVal FormName As String) As Boolean
  Dim f As Integer
  
  For f = 0 To Forms.Count - 1
    If UCase(Forms(f).Name) = UCase(FormName) Then
      frmIsLoaded = True
      Exit For
    End If
  Next f
  
End Function

Function frmTopmost(ByVal hWnd As Long, bTopMost As Boolean) As Boolean
  Dim lFlags As Long
  
  lFlags = SWP_NOMOVE Or SWP_NOSIZE
  
  frmTopmost = (SetWindowPos(hWnd, _
    IIf(bTopMost, HWND_TOPMOST, hWnd_NOTOPMOST), _
    0, 0, 0, 0, lFlags) <> 0)
End Function

Function lstFind(ByVal hWnd As Long, ByVal FindStr As String, ByVal Exact As Boolean) As Integer
  lstFind = SendMessage(hWnd, IIf(Exact, LB_FINDSTRINGEXACT, LB_FINDSTRING), -1, FindStr)
End Function

Function lstSelect(ByVal hWnd As Long, ByVal SelectStr As String) As Integer
  lstSelect = SendMessage(hWnd, LB_SELECTSTRING, -1, SelectStr)
End Function

Function lstDedup(SourceList As Object) As Boolean
  On Error GoTo ErrorTrap
  
  Dim i As Integer
  
  With SourceList
    If .ListCount > 1 Then
      i = 0
      While i < .ListCount
        If .List(i) = .List(i + 1) Then
          .RemoveItem (i)
        Else
          i = i + 1
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

Function treMoveNext(TreeView As ComctlLib.TreeView) As Long
  On Error GoTo ErrorTrap
  
  Dim ThisNode As ComctlLib.Node
      
  With TreeView
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
  treMoveNext = 0
  Err = False
  
End Function

Function treMovePrevious(TreeView As ComctlLib.TreeView) As Long
  On Error GoTo ErrorTrap
  
  Dim ThisNode As ComctlLib.Node
  
  With TreeView
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
  treMovePrevious = 0
  Err = False

End Function

Function txtSelText() As String
  If TypeOf Screen.ActiveForm.ActiveControl Is TextBox Then
    With Screen.ActiveForm.ActiveControl
      .SelStart = 0
      .SelLength = Len(.Text)
      txtSelText = .SelText
    End With
  End If
End Function
Function txtSelNumeric() As String
  If TypeOf Screen.ActiveForm.ActiveControl Is TDBNumber Then
    With Screen.ActiveForm.ActiveControl
      .SelStart = 0
      .SelLength = Len(.Text)
      txtSelNumeric = .SelText
    End With
  End If
End Function


Public Function YBorder() As Double
  ' Return the height of a control border.
  YBorder = GetSystemMetricsAPI(SM_CYBORDER) * Screen.TwipsPerPixelY

End Function


Public Function YFrame() As Double
  ' Return the height of a control frame.
  YFrame = GetSystemMetricsAPI(SM_CYFRAME) * Screen.TwipsPerPixelY

End Function


Function GetAvgCharWidth(ByVal hDC As Long) As Integer
  Dim typTxtMetrics As TEXTMETRIC
  
  Call GetTextMetrics(hDC, typTxtMetrics)
  
  GetAvgCharWidth = (typTxtMetrics.tmAveCharWidth * Screen.TwipsPerPixelX)

End Function

Function GetMaxCharWidth(ByVal hDC As Long) As Integer
  Dim typTxtMetrics As TEXTMETRIC
  
  Call GetTextMetrics(hDC, typTxtMetrics)
  
  GetMaxCharWidth = (typTxtMetrics.tmMaxCharWidth * Screen.TwipsPerPixelX)

End Function

Function GetCaption(Control As Object) As String
  On Error Resume Next
  
  'Attempt to set the Caption property
  GetCaption = Control.Caption
  If Err Then
    'Attempt to set the Text property
    GetCaption = Control.Text
  End If
  Err = False

End Function

Function GetCharHeight(ByVal hDC As Long) As Integer
  Dim typTxtMetrics As TEXTMETRIC
  
  Call GetTextMetrics(hDC, typTxtMetrics)
  
  GetCharHeight = (typTxtMetrics.tmHeight * Screen.TwipsPerPixelY)

End Function

Function GetDeskTopSize(SizeX As Long, SizeY As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim hWndDeskTop As Long
  Dim RectDeskTop As Rect
  
  hWndDeskTop = GetDesktopWindow()
  If GetClientRect(hWndDeskTop, RectDeskTop) Then
    SizeX = (RectDeskTop.Right - RectDeskTop.Left) * Screen.TwipsPerPixelX
    SizeY = (RectDeskTop.Bottom - RectDeskTop.Top) * Screen.TwipsPerPixelY
  End If
  
  GetDeskTopSize = True
  Exit Function

ErrorTrap:
  GetDeskTopSize = False
  Err = False

End Function

Function GetInverseColor(ByVal Color As Long) As Long
  Dim r As Integer, g As Integer, b As Integer
  
  If Int(Color / (2 ^ 24)) <> 0 Then
    Color = GetSysColor((Color And ((2 ^ 24) - 1)))
  End If
  Color = (((2 ^ 24) - 1) Xor Color)
  
  GetRGB Color, r, g, b
  If Abs(128 - r) < 30 And Abs(128 - g) < 30 And Abs(128 - b) < 30 Then
    GetInverseColor = RGB(255, 255, 255)
  Else
    GetInverseColor = Color
  End If
End Function

Function GetMousePos(MouseX As Long, MouseY As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim MousePos As POINTAPI
  
  GetCursorPos MousePos
  MouseX = MousePos.x * Screen.TwipsPerPixelX
  MouseY = MousePos.y * Screen.TwipsPerPixelY

  GetMousePos = True
  Exit Function
  
ErrorTrap:
  GetMousePos = False
  Err = False

End Function

Function GetOSName() As String
  Dim typOSVer As OSVERSIONINFO
  Dim strOSName As String
  
  typOSVer.dwOSVersionInfoSize = Len(typOSVer)
  If GetVersionEx(typOSVer) Then
    Select Case typOSVer.dwPlatformId
      Case VER_PLATFORM_WIN32s
        strOSName = "Windows 32s"
      Case VER_PLATFORM_WIN32_WINDOWS
        strOSName = "Windows 95"
      Case VER_PLATFORM_WIN32_NT
        strOSName = "Windows NT"
    End Select
  Else
    strOSName = vbNullString
  End If

  GetOSName = strOSName
  
End Function

Function GetOSVersion() As String
  Dim typOSVer As OSVERSIONINFO
  Dim strOSVersion As String
  
  typOSVer.dwOSVersionInfoSize = Len(typOSVer)
  If GetVersionEx(typOSVer) Then
    strOSVersion = Trim(Str(typOSVer.dwMajorVersion)) & "." & _
      Trim(Str(typOSVer.dwMinorVersion)) & " " & _
      Trim(typOSVer.szCSDVersion)
  Else
    strOSVersion = vbNullString
  End If

  GetOSVersion = strOSVersion
  
End Function

Function GetRGB(ByVal rgbValue As Long, RValue As Integer, GValue As Integer, BValue As Integer) As Boolean
  Dim lngRGB As Long
  Dim intRed As Integer, intGreen As Integer, intBlue As Integer
  
  lngRGB = (rgbValue And ((2 ^ 24) - 1))
  intBlue = Int(lngRGB / (2 ^ 16))
  intGreen = Int((lngRGB And ((2 ^ 16) - 1)) / (2 ^ 8))
  intRed = (lngRGB And ((2 ^ 8) - 1))
  
  RValue = intRed
  GValue = intGreen
  BValue = intBlue
  
  GetRGB = True
  
End Function

Function GetSystemColor(ByVal ColorDef As Long) As Long
  GetSystemColor = GetSysColor(ColorDef)
End Function

Function GetSystemMetrics(ByVal Index As SystemMetrics) As Long
  GetSystemMetrics = GetSystemMetricsAPI(Index)
End Function

Function GetControlAtPoint(ThisControl As Control) As Control
  Dim ParentForm As Form
  Dim ParentControl As Object
  Dim WndRect As Rect
  Dim hWnd As Long, hWndFound As Long
  Dim i As Integer
  Dim pt As POINTAPI
  Dim x As Integer, y As Integer
  
  hWnd = 0
  hWndFound = 0
  
  'Get current mouse position, relative to entire screen
  Call GetCursorPos(pt)
  
  'Check if this controls container is its parent form
  If ThisControl.Container Is ThisControl.Parent Then
    'Set pointer to controls parent
    Set ParentControl = ThisControl.Parent
    
    'Get position of controls parent, relative to entire screen
    Call GetWindowRect(ParentControl.hWnd, WndRect)
    
    'Calculate current mouse position within the controls parent
    x = pt.x - (WndRect.Left + GetSystemMetrics(SM_CXFRAME))
    y = pt.y - (WndRect.Top + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYFRAME))
  Else
    'Set pointer to controls container
    Set ParentControl = ThisControl.Container
    
    'Get position of controls container, relative to entire screen
    Call GetWindowRect(ParentControl.hWnd, WndRect)
    
    'Calculate current mouse position within the controls container
    x = pt.x - (WndRect.Left + GetSystemMetrics(SM_CXBORDER))
    y = pt.y - (WndRect.Top + GetSystemMetrics(SM_CYBORDER))
  End If
  
  'Attempt to find a visible child window at mouse position
  hWnd = ChildWindowFromPointEx(ParentControl.hWnd, x, y, CWP_SKIPINVISIBLE)
  
  'Check if a visible child window was found
  Do While hWnd > 0 And hWnd <> hWndFound
    'Save handle of found window
    hWndFound = hWnd
    
    'Get position of found window, relative to entire screen
    Call GetWindowRect(hWndFound, WndRect)
    
    'Calculate mouse position within the found window
    x = pt.x - (WndRect.Left + GetSystemMetrics(SM_CXBORDER))
    y = pt.y - (WndRect.Top + GetSystemMetrics(SM_CYBORDER))
  
    'Attempt to find a visible child window at mouse position
    hWnd = ChildWindowFromPointEx(hWndFound, x, y, CWP_SKIPINVISIBLE)
  Loop
  
  'Check if a visible child window was found
  If hWndFound > 0 Then
    'Set pointer to parent form
    Set ParentForm = ThisControl.Parent
    
    On Error Resume Next
    
    'Loop through parent forms controls
    For i = 0 To ParentForm.Controls.Count - 1
      'Check is this controls handle is the found windows handle
      If ParentForm.Controls(i).hWnd = hWndFound Then
        If Err = 0 Then
          'Return a pointer to the found control
          Set GetControlAtPoint = ParentForm.Controls(i)
          Exit For
        Else
          Err = False
        End If
      End If
    Next i
  End If
  
End Function

Function GetWorkAreaSize(SizeX As Long, SizeY As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim RectWorkArea As Rect

  If SystemParametersInfo(SPI_GETWORKAREA, vbEmpty, RectWorkArea, vbEmpty) Then
    SizeX = (RectWorkArea.Right - RectWorkArea.Left) * Screen.TwipsPerPixelX
    SizeY = (RectWorkArea.Bottom - RectWorkArea.Top) * Screen.TwipsPerPixelY
  End If
  
  GetWorkAreaSize = True
  Exit Function

ErrorTrap:
  GetWorkAreaSize = False
  Err = False

End Function

Function LockWindow(ByVal hWnd As Long) As Boolean
  'Unlock any window currently locked
  UnlockWindow
  'Lock required window
  LockWindow = LockWindowUpdate(hWnd)
End Function

Sub SendKeys(ByVal SendString As String, Optional ByVal hWnd As Long)
  Dim intChar As Integer
  Dim intAscii As Integer
  Dim intVKey As Integer
  Dim intOEMScan As Integer
  Dim strOEMChar As String
  
  If Len(SendString) < 1 Then
    Exit Sub
  End If
  If hWnd > 0 Then
    'Set focus to required window
    Call SetFocusAPI(hWnd)
  End If
  For intChar = 1 To Len(SendString)
    intAscii = Asc(Mid(SendString, intChar, 1))
    
    'Get the virtual key code for this character
    intVKey = VkKeyScan(intAscii) And &HFF
    
    'Get the OEM character
    strOEMChar = Space(2)
    CharToOem Chr(intAscii), strOEMChar
    
    'Get the OEM scan code
    intOEMScan = OemKeyScan(Asc(strOEMChar)) And &HFF
    
    'Send the down key
    keybd_event intVKey, intOEMScan, 0, 0
    
    'Send the up key
    keybd_event intVKey, intOEMScan, KEYEVENTF_KEYUP, 0
  Next intChar
  
End Sub

Sub SetCaption(Control As Object, ByVal Caption As String)
  On Error Resume Next
  
  'Attempt to set the Caption property
  Control.Caption = Caption
  If Err Then
    'Attempt to set the Text property
    Control.Text = Caption
  End If
  Err = False

End Sub

Sub SetFocus(ByVal hWnd As Long)
  'Set focus to required window
  Call SetFocusAPI(hWnd)
End Sub

Function SetThickFrame(ByVal hWnd As Long, ByVal ThickFrame As Boolean)
  Dim WndStyle As Long
  
  'Get window style info
  WndStyle = GetWindowLong(hWnd, GWL_STYLE)
  
  'Change window style info
  If ThickFrame Then
    WndStyle = (WndStyle Or WS_THICKFRAME)
  Else
    WndStyle = (WndStyle Xor WS_THICKFRAME)
  End If
  
  'Update window style
  If SetWindowLong(hWnd, GWL_STYLE, WndStyle) <> 0 Then
    SetThickFrame = True
  End If

End Function

Function UnlockWindow() As Boolean
  'Unlock any window currently locked
  UnlockWindow = LockWindowUpdate(0&)
End Function

Function GetHostName() As String
  Dim sBuffer As String
  Dim lngSize As Long
  
  'MH20010206
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

Private Function GetX()
    Dim Point As POINTAPI, retVAL As Long
    retVAL = GetCursorPos(Point)
    GetX = Point.x
End Function

Private Function GetY()
    Dim Point As POINTAPI, retVAL As Long
    retVAL = GetCursorPos(Point)
    GetY = Point.y
End Function

Private Sub SetClipVars(MINHEIGHT As Long, MINWIDTH As Long)
    frmHeight = MINHEIGHT
    frmWidth = MINWIDTH
End Sub

Public Sub ClipForForm(frm As Form, MINHEIGHT As Long, MINWIDTH As Long)
    Dim ResizeREC As Rect
    Dim DesktopREC As Rect
    Dim GetClipREC As Rect
    ResizeREC.Top = (frm.Top + MINHEIGHT) / Screen.TwipsPerPixelY - 2
    ResizeREC.Bottom = (frm.Top + (frm.Height - MINHEIGHT)) / Screen.TwipsPerPixelY + 2
    ResizeREC.Left = (frm.Left + MINWIDTH) / Screen.TwipsPerPixelX - 2
    ResizeREC.Right = (frm.Left + (frm.Width - MINWIDTH)) / Screen.TwipsPerPixelX + 2
    If frm.Width <> frmWidth And frm.WindowState = 0 Then
        retVAL = GetClipCursor(GetClipREC)
        DeskhWnd = GetDesktopWindow()
        retVAL = GetWindowRect(DeskhWnd, DesktopREC)
        If GetX > ((frm.Left + (MINWIDTH / 2)) / Screen.TwipsPerPixelX) And GetClipREC.Right = DesktopREC.Right Then
            If GetClipREC.Left = ResizeREC.Left And (GetClipREC.Top = ResizeREC.Top Or GetClipREC.Bottom = ResizeREC.Bottom) Then
                frmHeight = frm.Height
                frmWidth = frm.Width
                Exit Sub
            ElseIf GetClipREC.Left = ResizeREC.Left And (GetClipREC.Top <> ResizeREC.Top Or GetClipREC.Bottom <> ResizeREC.Bottom) Then
                If frm.Height <> frmHeight Then
                    If GetY > (frm.Top / Screen.TwipsPerPixelY) + 25 Then
                        DesktopREC.Top = ResizeREC.Top
                    Else
                        DesktopREC.Bottom = ResizeREC.Bottom
                    End If
                    frmHeight = frm.Height
                End If
            ElseIf GetClipREC.Left <> ResizeREC.Left And (GetClipREC.Top = ResizeREC.Top Or GetClipREC.Bottom = ResizeREC.Bottom) Then
                If GetClipREC.Top = ResizeREC.Top Then DesktopREC.Top = GetClipREC.Top
                If GetClipREC.Bottom = ResizeREC.Bottom Then DesktopREC.Bottom = GetClipREC.Bottom
                DesktopREC.Left = ResizeREC.Left
            End If
            DesktopREC.Left = ResizeREC.Left
        Else
            If GetClipREC.Right = ResizeREC.Right And (GetClipREC.Bottom = ResizeREC.Bottom Or GetClipREC.Top = ResizeREC.Top) Then
                frmHeight = frm.Height
                frmWidth = frm.Width
                Exit Sub
            ElseIf GetClipREC.Right = ResizeREC.Right And (GetClipREC.Bottom <> ResizeREC.Bottom Or GetClipREC.Top <> ResizeREC.Top) Then
                If frm.Height <> frmHeight Then
                    If GetY > (frm.Top / Screen.TwipsPerPixelY) + 25 Then
                        DesktopREC.Top = ResizeREC.Top
                    Else
                        DesktopREC.Bottom = ResizeREC.Bottom
                    End If
                    frmHeight = frm.Height
                End If
            ElseIf GetClipREC.Right <> ResizeREC.Right And (GetClipREC.Bottom = ResizeREC.Bottom Or GetClipREC.Top = ResizeREC.Top) Then
                If GetClipREC.Top = ResizeREC.Top Then DesktopREC.Top = GetClipREC.Top
                If GetClipREC.Bottom = ResizeREC.Bottom Then DesktopREC.Bottom = GetClipREC.Bottom
            End If
            DesktopREC.Right = ResizeREC.Right
        End If
        retVAL = ClipCursor(DesktopREC)
        frmWidth = frm.Width
    ElseIf frm.Height <> frmHeight And frm.WindowState = 0 Then
        retVAL = GetClipCursor(GetClipREC)
        DeskhWnd = GetDesktopWindow()
        retVAL = GetWindowRect(DeskhWnd, DesktopREC)
        If GetY > ((frm.Top + (MINHEIGHT / 2)) / Screen.TwipsPerPixelY) And GetClipREC.Bottom <> ResizeREC.Bottom Then
            If GetClipREC.Top = ResizeREC.Top And (GetClipREC.Left = ResizeREC.Left Or GetClipREC.Right = ResizeREC.Right) Then
                frmHeight = frm.Height
                frmWidth = frm.Width
                Exit Sub
            ElseIf GetClipREC.Top = ResizeREC.Top And (GetClipREC.Left <> ResizeREC.Left Or GetClipREC.Right <> ResizeREC.Right) Then
                If frm.Width <> frmWidth Then
                    If GetX > (frm.Left / Screen.TwipsPerPixelX) + 15 Then
                        DesktopREC.Left = ResizeREC.Left
                    Else
                        DesktopREC.Right = ResizeREC.Right
                    End If
                    frmWidth = frm.Width
                End If
            ElseIf GetClipREC.Top <> ResizeREC.Top And (GetClipREC.Left = ResizeREC.Left Or GetClipREC.Right = ResizeREC.Right) Then
                If GetClipREC.Left = ResizeREC.Left Then DesktopREC.Left = GetClipREC.Left
                If GetClipREC.Right = ResizeREC.Right Then DesktopREC.Right = GetClipREC.Right
            End If
            DesktopREC.Top = ResizeREC.Top
        Else
            If GetClipREC.Bottom = ResizeREC.Bottom And (GetClipREC.Right = ResizeREC.Right Or GetClipREC.Left = ResizeREC.Left) Then
                frmHeight = frm.Height
                frmWidth = frm.Width
                Exit Sub
            ElseIf GetClipREC.Bottom = ResizeREC.Bottom And (GetClipREC.Right <> ResizeREC.Right Or GetClipREC.Left <> ResizeREC.Left) Then
                If frm.Width <> frmWidth Then
                    If GetX > (frm.Left / Screen.TwipsPerPixelX) + 15 Then
                        DesktopREC.Left = ResizeREC.Left
                    Else
                        DesktopREC.Right = ResizeREC.Right
                    End If
                    frmWidth = frm.Width
                End If
            ElseIf GetClipREC.Bottom <> ResizeREC.Bottom And (GetClipREC.Right = ResizeREC.Right Or GetClipREC.Left = ResizeREC.Left) Then
                If GetClipREC.Left = ResizeREC.Left Then DesktopREC.Left = GetClipREC.Left
                If GetClipREC.Right = ResizeREC.Right Then DesktopREC.Right = GetClipREC.Right
            End If
            DesktopREC.Bottom = ResizeREC.Bottom
        End If
        retVAL = ClipCursor(DesktopREC)
        frmHeight = frm.Height
    End If
End Sub

Public Sub RemoveClipping()
    Dim DesktopREC As Rect, retVAL As Long, DeskhWnd As Long
    DeskhWnd = GetDesktopWindow()
    retVAL = GetWindowRect(DeskhWnd, DesktopREC)
    retVAL = ClipCursor(DesktopREC)
End Sub

Public Function GetSystemLongDateFormat() As String
  ' Return the system data format.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE, sBuffer, 99)
  GetSystemLongDateFormat = Left(sBuffer, lngLength - 1)
  
End Function
