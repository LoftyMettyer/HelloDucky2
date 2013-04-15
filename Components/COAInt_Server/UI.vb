Option Strict Off
Option Explicit On
Friend Class clsUI
	
	Private Structure POINTAPI
		Dim x As Integer
		Dim y As Integer
	End Structure
	
	Private retVAL As Object
	Private DeskhWnd As Object
	
	Private Structure TEXTMETRIC
		Dim tmHeight As Integer
		Dim tmAscent As Integer
		Dim tmDescent As Integer
		Dim tmInternalLeading As Integer
		Dim tmExternalLeading As Integer
		Dim tmAveCharWidth As Integer
		Dim tmMaxCharWidth As Integer
		Dim tmWeight As Integer
		Dim tmOverhang As Integer
		Dim tmDigitizedAspectX As Integer
		Dim tmDigitizedAspectY As Integer
		Dim tmFirstChar As Byte
		Dim tmLastChar As Byte
		Dim tmDefaultChar As Byte
		Dim tmBreakChar As Byte
		Dim tmItalic As Byte
		Dim tmUnderlined As Byte
		Dim tmStruckOut As Byte
		Dim tmPitchAndFamily As Byte
		Dim tmCharSet As Byte
	End Structure
	
	Private Structure OSVERSIONINFO
		Dim dwOSVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char '  Maintenance string for PSS usage
	End Structure
	
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Private Structure MINMAXINFO
		Dim ptReserved As POINTAPI
		Dim ptMaxSize As POINTAPI
		Dim ptMaxPosition As POINTAPI
		Dim ptMinTrackSize As POINTAPI
		Dim ptMaxTrackSize As POINTAPI
	End Structure
	
	'Windows version constants
	Const VER_PLATFORM_WIN32s As Short = 0
	Const VER_PLATFORM_WIN32_WINDOWS As Short = 1
	Const VER_PLATFORM_WIN32_NT As Short = 2
	
	'Combo box constants
	Const CB_ERR As Short = (-1)
	Const CB_FINDSTRING As Integer = &H14C
	Const CB_FINDSTRINGEXACT As Integer = &H158
	Const CB_SELECTSTRING As Integer = &H14D
	
	'List box constants
	Const LB_ERR As Short = (-1)
	Const LB_FINDSTRING As Integer = &H18F
	Const LB_FINDSTRINGEXACT As Integer = &H1A2
	Const LB_SELECTSTRING As Integer = &H18C
	
	'Window constants
	Const HWND_TOPMOST As Short = -1
	Const hWnd_NOTOPMOST As Short = -2
	Const SWP_NOMOVE As Integer = &H2
	Const SWP_NOSIZE As Integer = &H1
	
	Const SPI_GETWORKAREA As Short = 48
	
	'ChildWindowFromPointEx constants
	Const CWP_ALL As Short = 0
	Const CWP_SKIPINVISIBLE As Short = 1
	Const CWP_SKIPDISABLED As Short = 2
	Const CWP_SKIPTRANSPARENT As Short = 4
	
	'Window style constants
	Const GWL_STYLE As Short = (-16)
	Const WS_THICKFRAME As Integer = &H40000
	
	'System metrics constants
	Public Enum SystemMetrics
		SM_CXVSCROLL = 2
		SM_CYCAPTION = 4
		SM_CXBORDER = 5
		SM_CYBORDER = 6
		SM_CXFRAME = 32
		SM_CYFRAME = 33
		SM_CYSMCAPTION = 51
	End Enum
	
	'Keyboard constants
	Const KEYEVENTF_KEYUP As Integer = &H2
	
	Const MAX_COMPUTERNAME_LENGTH As Short = 15
	
	Const LOCALE_SYSTEM_DEFAULT As Integer = &H800
	Const LOCALE_USER_DEFAULT As Integer = &H400
	Const LOCALE_SDATE As Integer = &H1D ' date separator
	Const LOCALE_SSHORTDATE As Integer = &H1F ' short date format string
	Const LOCALE_SDECIMAL As Integer = &HE ' decimal separator
	Const LOCALE_STHOUSAND As Integer = &HF ' thousand separator
	Const LOCALE_IMEASURE As Integer = &HD ' Measurement System
	
	
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function ClipCursor Lib "user32.dll" (ByRef lpRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetClipCursor Lib "user32.dll" (ByRef lprc As RECT) As Integer
	Private frmHeight As Integer
	Private frmWidth As Integer
	
	'Windows API functions
	Private Declare Function CharToOem Lib "user32"  Alias "CharToOemA"(ByVal lpszSrc As String, ByVal lpszDst As String) As Integer
	Private Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As Integer, ByVal xPoint As Integer, ByVal yPoint As Integer, ByVal Flags As Integer) As Integer
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Integer, ByRef lpPoint As POINTAPI) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
	Private Declare Function GetDesktopWindow Lib "user32" () As Integer
	Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Integer
	Private Declare Function GetSystemMetricsAPI Lib "user32"  Alias "GetSystemMetrics"(ByVal nIndex As Integer) As Integer
	'UPGRADE_WARNING: Structure TEXTMETRIC may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetTextMetrics Lib "gdi32"  Alias "GetTextMetricsA"(ByVal hdc As Integer, ByRef lpMetrics As TEXTMETRIC) As Integer
	'UPGRADE_WARNING: Structure OSVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetVersionEx Lib "kernel32"  Alias "GetVersionExA"(ByRef lpVersionInformation As OSVERSIONINFO) As Integer
	Private Declare Function GetWindowLong Lib "user32"  Alias "GetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
	Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Integer) As Integer
	Private Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
	Private Declare Function SetFocusAPI Lib "user32"  Alias "SetFocus"(ByVal hWnd As Integer) As Integer
	Private Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
  Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Integer, ByVal uParam As Integer, ByRef lpvParam As String, ByVal fuWinIni As Integer) As Integer
	'UPGRADE_NOTE: cChar was upgraded to cChar_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Declare Function VkKeyScan Lib "user32"  Alias "VkKeyScanA"(ByVal cChar_Renamed As Byte) As Short
	Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Integer, ByVal yPoint As Integer) As Integer
	Private Declare Function GetLocaleInfo Lib "kernel32"  Alias "GetLocaleInfoA"(ByVal locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
	Private Declare Function GetComputerName Lib "kernel32"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	
	
	Public Function CaptionHeight() As Double
		' Return the height of a form's caption bar.
		CaptionHeight = GetSystemMetricsAPI(SystemMetrics.SM_CYSMCAPTION) * VB6.TwipsPerPixelY
		
	End Function
	
  Shared Function GetSystemDateSeparator() As String
    ' Return the system data separator.
    Dim lngLength As Integer
    Dim sBuffer As New VB6.FixedLengthString(100)

    lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, sBuffer.Value, 99)
    GetSystemDateSeparator = Left(sBuffer.Value, lngLength - 1)

  End Function
	
  Shared Function GetSystemDateFormat() As String
    ' Return the system data format.
    Dim lngLength As Integer
    Dim sBuffer As New VB6.FixedLengthString(100)

    lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer.Value, 99)
    GetSystemDateFormat = Left(sBuffer.Value, lngLength - 1)

  End Function
	
  Shared Function GetSystemDecimalSeparator() As String
    ' Return the system data separator.
    Dim lngLength As Integer
    Dim sBuffer As New VB6.FixedLengthString(100)

    lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer.Value, 99)
    GetSystemDecimalSeparator = Left(sBuffer.Value, lngLength - 1)


  End Function
	
  Shared Function GetSystemThousandSeparator() As String
    ' Return the system data separator.
    Dim lngLength As Integer
    Dim sBuffer As New VB6.FixedLengthString(100)

    lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sBuffer.Value, 99)
    GetSystemThousandSeparator = Left(sBuffer.Value, lngLength - 1)

  End Function
	
	
	Function GetSystemMeasurement() As String
		' Return the system measurement (metric or us).
		
		Dim lngLength As Integer
		Dim sBuffer As New VB6.FixedLengthString(100)
		
		lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, sBuffer.Value, 99)
		GetSystemMeasurement = Left(sBuffer.Value, lngLength - 1)
		
		If CDbl(GetSystemMeasurement) = 1 Then
			GetSystemMeasurement = "us"
		Else
			GetSystemMeasurement = "metric"
		End If
		
	End Function
	
	
	Function cboFind(ByVal hWnd As Integer, ByVal FindStr As String, ByVal Exact As Boolean) As Short
		cboFind = SendMessage(hWnd, IIf(Exact, CB_FINDSTRINGEXACT, CB_FINDSTRING), -1, FindStr)
	End Function
	
	Function cboSelect(ByVal hWnd As Integer, ByVal SelectStr As String) As Short
		cboSelect = SendMessage(hWnd, CB_SELECTSTRING, -1, SelectStr)
	End Function
	
	Function cboDropDown(ByVal hWnd As Integer, ByRef bDropDown As Boolean) As Short
		
		' RH 31/07/00 - Automatically drop down a combo box.
		cboDropDown = SendMessage(hWnd, &H14F, bDropDown, "")
		
	End Function
	
  'Sub frmAtCenter(ByRef ThisForm As System.Windows.Forms.Form)
  '	On Error Resume Next

  '	With ThisForm
  '		.Top = VB6.TwipsToPixelsY(Int((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(.Height)) / 2))
  '		.Left = VB6.TwipsToPixelsX(Int((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(.Width)) / 2))
  '	End With

  'End Sub
	
  'Sub frmAtMouse(ByRef ThisForm As System.Windows.Forms.Form)
  '	On Error Resume Next

  '	Dim MouseX, MouseY As Integer
  '	Dim SizeX, SizeY As Integer

  '	If Not GetWorkAreaSize(SizeX, SizeY) Then
  '		SizeX = VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width)
  '		SizeY = VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height)
  '	End If

  '	If GetMousePos(MouseX, MouseY) Then
  '		With ThisForm
  '			If MouseY + VB6.PixelsToTwipsY(.Height) > SizeY Then
  '				.Top = VB6.TwipsToPixelsY(SizeY - VB6.PixelsToTwipsY(.Height))
  '			Else
  '				.Top = VB6.TwipsToPixelsY(MouseY)
  '			End If
  '			If MouseX + VB6.PixelsToTwipsX(.Width) > SizeX Then
  '				.Left = VB6.TwipsToPixelsX(SizeX - VB6.PixelsToTwipsX(.Width))
  '			Else
  '				.Left = VB6.TwipsToPixelsX(MouseX)
  '			End If
  '		End With
  '	End If

  'End Sub
	
  'Function frmIsLoaded(ByVal FormName As String) As Boolean
  '	Dim f As Short

  '   For f = 0 To My.Application.OpenForms.Count - 1
  '     If UCase(My.Application.OpenForms.Item(f).Name) = UCase(FormName) Then
  '       frmIsLoaded = True
  '       Exit For
  '     End If
  '   Next f

  'End Function
	
  'Function frmTopmost(ByVal hWnd As Integer, ByRef bTopMost As Boolean) As Boolean
  '	Dim lFlags As Integer

  '	lFlags = SWP_NOMOVE Or SWP_NOSIZE

  '	frmTopmost = (SetWindowPos(hWnd, IIf(bTopMost, HWND_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, lFlags) <> 0)
  'End Function
	
	Function lstFind(ByVal hWnd As Integer, ByVal FindStr As String, ByVal Exact As Boolean) As Short
		lstFind = SendMessage(hWnd, IIf(Exact, LB_FINDSTRINGEXACT, LB_FINDSTRING), -1, FindStr)
	End Function
	
	Function lstSelect(ByVal hWnd As Integer, ByVal SelectStr As String) As Short
		lstSelect = SendMessage(hWnd, LB_SELECTSTRING, -1, SelectStr)
	End Function
	
	Function lstDedup(ByRef SourceList As Object) As Boolean
		On Error GoTo ErrorTrap
		
		Dim i As Short
		
		With SourceList
			'UPGRADE_WARNING: Couldn't resolve default property of object SourceList.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .ListCount > 1 Then
				i = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object SourceList.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				While i < .ListCount
					'UPGRADE_WARNING: Couldn't resolve default property of object SourceList.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .List(i) = .List(i + 1) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object SourceList.RemoveItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.RemoveItem(i)
					Else
						i = i + 1
					End If
				End While
			End If
		End With
		
		lstDedup = True
		Exit Function
		
ErrorTrap: 
		lstDedup = False
		Err.Number = False
		
	End Function
	
	'Function treMoveNext(TreeView As ComctlLib.TreeView) As Long
	'  On Error GoTo ErrorTrap
	'
	'  Dim ThisNode As ComctlLib.Node
	'
	'  With TreeView
	'    If .SelectedItem.Children > 0 And .SelectedItem.Expanded Then
	'      .SelectedItem = .SelectedItem.Child
	'    Else
	'      If .SelectedItem <> .SelectedItem.LastSibling Then
	'        .SelectedItem = .SelectedItem.Next
	'      Else
	'        Set ThisNode = .SelectedItem
	'        Do While ThisNode <> ThisNode.Root
	'          If ThisNode.Parent <> ThisNode.Parent.LastSibling Then
	'            .SelectedItem = ThisNode.Parent.Next
	'            Exit Do
	'          Else
	'            Set ThisNode = ThisNode.Parent
	'          End If
	'        Loop
	'        Set ThisNode = Nothing
	'      End If
	'    End If
	'
	'    treMoveNext = .SelectedItem.Index
	'  End With
	'
	'  Exit Function
	'
	'ErrorTrap:
	'  treMoveNext = 0
	'  Err = False
	'
	'End Function
	'
	'Function treMovePrevious(TreeView As ComctlLib.TreeView) As Long
	'  On Error GoTo ErrorTrap
	'
	'  Dim ThisNode As ComctlLib.Node
	'
	'  With TreeView
	'    If .SelectedItem <> .SelectedItem.Root Then
	'      If .SelectedItem <> .SelectedItem.FirstSibling Then
	'        Set ThisNode = .SelectedItem.Previous
	'        Do While ThisNode.Children > 0 And ThisNode.Expanded
	'          Set ThisNode = ThisNode.Child.LastSibling
	'        Loop
	'        .SelectedItem = ThisNode
	'        Set ThisNode = Nothing
	'      Else
	'        .SelectedItem = .SelectedItem.Parent
	'      End If
	'    End If
	'
	'    treMovePrevious = .SelectedItem.Index
	'  End With
	'
	'  Exit Function
	'
	'ErrorTrap:
	'  treMovePrevious = 0
	'  Err = False
	'
	'End Function
	
  'Function txtSelText() As String
  '	If TypeOf System.Windows.Forms.Form.ActiveForm.ActiveControl Is System.Windows.Forms.TextBox Then
  '		With System.Windows.Forms.Form.ActiveForm.ActiveControl
  '			'UPGRADE_WARNING: Couldn't resolve default property of object Screen.ActiveForm.ActiveControl.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '			.SelStart = 0
  '			'UPGRADE_WARNING: Couldn't resolve default property of object Screen.ActiveForm.ActiveControl.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '			.SelLength = Len(.Text)
  '			'UPGRADE_WARNING: Couldn't resolve default property of object Screen.ActiveForm.ActiveControl.SelText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '			txtSelText = .SelText
  '		End With
  '	End If
  'End Function
	
	Public Function YBorder() As Double
		' Return the height of a control border.
		YBorder = GetSystemMetricsAPI(SystemMetrics.SM_CYBORDER) * VB6.TwipsPerPixelY
		
	End Function
	
	
	Public Function YFrame() As Double
		' Return the height of a control frame.
		YFrame = GetSystemMetricsAPI(SystemMetrics.SM_CYFRAME) * VB6.TwipsPerPixelY
		
	End Function
	
	
	Function GetAvgCharWidth(ByVal hdc As Integer) As Short
		Dim typTxtMetrics As TEXTMETRIC
		
		Call GetTextMetrics(hdc, typTxtMetrics)
		
		GetAvgCharWidth = (typTxtMetrics.tmAveCharWidth * VB6.TwipsPerPixelX)
		
	End Function
	
	Function GetMaxCharWidth(ByVal hdc As Integer) As Short
		Dim typTxtMetrics As TEXTMETRIC
		
		Call GetTextMetrics(hdc, typTxtMetrics)
		
		GetMaxCharWidth = (typTxtMetrics.tmMaxCharWidth * VB6.TwipsPerPixelX)
		
	End Function
	
	Function GetCaption(ByRef Control As Object) As String
		On Error Resume Next
		
		'Attempt to set the Caption property
		'UPGRADE_WARNING: Couldn't resolve default property of object Control.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCaption = Control.Caption
		If Err.Number Then
			'Attempt to set the Text property
			'UPGRADE_WARNING: Couldn't resolve default property of object Control.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetCaption = Control.Text
		End If
		Err.Number = False
		
	End Function
	
	Function GetCharHeight(ByVal hdc As Integer) As Short
		Dim typTxtMetrics As TEXTMETRIC
		
		Call GetTextMetrics(hdc, typTxtMetrics)
		
		GetCharHeight = (typTxtMetrics.tmHeight * VB6.TwipsPerPixelY)
		
	End Function
	
	Function GetDeskTopSize(ByRef SizeX As Integer, ByRef SizeY As Integer) As Boolean
		On Error GoTo ErrorTrap
		
		Dim hWndDeskTop As Integer
		Dim RectDeskTop As RECT
		
		hWndDeskTop = GetDesktopWindow()
		If GetClientRect(hWndDeskTop, RectDeskTop) Then
			SizeX = (RectDeskTop.Right_Renamed - RectDeskTop.Left_Renamed) * VB6.TwipsPerPixelX
			SizeY = (RectDeskTop.Bottom - RectDeskTop.Top) * VB6.TwipsPerPixelY
		End If
		
		GetDeskTopSize = True
		Exit Function
		
ErrorTrap: 
		GetDeskTopSize = False
		Err.Number = False
		
	End Function
	
	Function GetInverseColor(ByVal Color As Integer) As Integer
		Dim g, r, b As Short
		
		If Int(Color / (2 ^ 24)) <> 0 Then
			Color = GetSysColor(Color And ((2 ^ 24) - 1))
		End If
		Color = (((2 ^ 24) - 1) Xor Color)
		
		GetRGB(Color, r, g, b)
		If System.Math.Abs(128 - r) < 30 And System.Math.Abs(128 - g) < 30 And System.Math.Abs(128 - b) < 30 Then
			GetInverseColor = RGB(255, 255, 255)
		Else
			GetInverseColor = Color
		End If
	End Function
	
	Function GetMousePos(ByRef MouseX As Integer, ByRef MouseY As Integer) As Boolean
		On Error GoTo ErrorTrap
		
		Dim MousePos As POINTAPI
		
		GetCursorPos(MousePos)
		MouseX = MousePos.x * VB6.TwipsPerPixelX
		MouseY = MousePos.y * VB6.TwipsPerPixelY
		
		GetMousePos = True
		Exit Function
		
ErrorTrap: 
		GetMousePos = False
		Err.Number = False
		
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
			strOSVersion = Trim(Str(typOSVer.dwMajorVersion)) & "." & Trim(Str(typOSVer.dwMinorVersion)) & " " & Trim(typOSVer.szCSDVersion)
		Else
			strOSVersion = vbNullString
		End If
		
		GetOSVersion = strOSVersion
		
	End Function
	
	Function GetRGB(ByVal rgbValue As Integer, ByRef RValue As Short, ByRef GValue As Short, ByRef BValue As Short) As Boolean
		Dim lngRGB As Integer
		Dim intGreen, intRed, intBlue As Short
		
		lngRGB = (rgbValue And ((2 ^ 24) - 1))
		intBlue = Int(lngRGB / (2 ^ 16))
		intGreen = Int(CShort(lngRGB And ((2 ^ 16) - 1)) / 2 ^ 8)
		intRed = (lngRGB And ((2 ^ 8) - 1))
		
		RValue = intRed
		GValue = intGreen
		BValue = intBlue
		
		GetRGB = True
		
	End Function
	
	Function GetSystemColor(ByVal ColorDef As Integer) As Integer
		GetSystemColor = GetSysColor(ColorDef)
	End Function
	
	Function GetSystemMetrics(ByVal Index As SystemMetrics) As Integer
		GetSystemMetrics = GetSystemMetricsAPI(Index)
	End Function
	
	Function GetControlAtPoint(ByRef ThisControl As System.Windows.Forms.Control) As System.Windows.Forms.Control
		Dim ParentForm As System.Windows.Forms.Form
		Dim ParentControl As Object
		Dim WndRect As RECT
		Dim hWnd, hWndFound As Integer
		Dim i As Short
		Dim pt As POINTAPI
		Dim x, y As Short
		
		hWnd = 0
		hWndFound = 0
		
		'Get current mouse position, relative to entire screen
		Call GetCursorPos(pt)
		
		'Check if this controls container is its parent form
		'UPGRADE_WARNING: Control property ThisControl.Parent was upgraded to ThisControl.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		If ThisControl.Parent Is ThisControl.FindForm Then
			'Set pointer to controls parent
			'UPGRADE_WARNING: Control property ThisControl.Parent was upgraded to ThisControl.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			ParentControl = ThisControl.FindForm
			
			'Get position of controls parent, relative to entire screen
			'UPGRADE_WARNING: Couldn't resolve default property of object ParentControl.hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call GetWindowRect(ParentControl.hWnd, WndRect)
			
			'Calculate current mouse position within the controls parent
			x = pt.x - (WndRect.Left_Renamed + GetSystemMetrics(SystemMetrics.SM_CXFRAME))
			y = pt.y - (WndRect.Top + GetSystemMetrics(SystemMetrics.SM_CYCAPTION) + GetSystemMetrics(SystemMetrics.SM_CYFRAME))
		Else
			'Set pointer to controls container
			ParentControl = ThisControl.Parent
			
			'Get position of controls container, relative to entire screen
			'UPGRADE_WARNING: Couldn't resolve default property of object ParentControl.hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call GetWindowRect(ParentControl.hWnd, WndRect)
			
			'Calculate current mouse position within the controls container
			x = pt.x - (WndRect.Left_Renamed + GetSystemMetrics(SystemMetrics.SM_CXBORDER))
			y = pt.y - (WndRect.Top + GetSystemMetrics(SystemMetrics.SM_CYBORDER))
		End If
		
		'Attempt to find a visible child window at mouse position
		'UPGRADE_WARNING: Couldn't resolve default property of object ParentControl.hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hWnd = ChildWindowFromPointEx(ParentControl.hWnd, x, y, CWP_SKIPINVISIBLE)
		
		'Check if a visible child window was found
		Do While hWnd > 0 And hWnd <> hWndFound
			'Save handle of found window
			hWndFound = hWnd
			
			'Get position of found window, relative to entire screen
			Call GetWindowRect(hWndFound, WndRect)
			
			'Calculate mouse position within the found window
			x = pt.x - (WndRect.Left_Renamed + GetSystemMetrics(SystemMetrics.SM_CXBORDER))
			y = pt.y - (WndRect.Top + GetSystemMetrics(SystemMetrics.SM_CYBORDER))
			
			'Attempt to find a visible child window at mouse position
			hWnd = ChildWindowFromPointEx(hWndFound, x, y, CWP_SKIPINVISIBLE)
		Loop 
		
		'Check if a visible child window was found
		If hWndFound > 0 Then
			'Set pointer to parent form
			'UPGRADE_WARNING: Control property ThisControl.Parent was upgraded to ThisControl.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			ParentForm = ThisControl.FindForm
			
			On Error Resume Next
			
			'Loop through parent forms controls
			'UPGRADE_WARNING: Controls method Controls.Count has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			For i = 0 To ParentForm.Controls.Count() - 1
				'Check is this controls handle is the found windows handle
				If CType(ParentForm.Controls(i), Object).Handle.ToInt32 = hWndFound Then
					If Err.Number = 0 Then
						'Return a pointer to the found control
						GetControlAtPoint = CType(ParentForm.Controls(i), Object)
						Exit For
					Else
						Err.Number = False
					End If
				End If
			Next i
		End If
		
	End Function
	
	Function GetWorkAreaSize(ByRef SizeX As Integer, ByRef SizeY As Integer) As Boolean
		On Error GoTo ErrorTrap
		
		Dim RectWorkArea As RECT
		
		'UPGRADE_WARNING: Couldn't resolve default property of object RectWorkArea. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    If SystemParametersInfo(SPI_GETWORKAREA, VariantType.Empty, RectWorkArea.ToString(), VariantType.Empty) Then
      SizeX = (RectWorkArea.Right_Renamed - RectWorkArea.Left_Renamed) * VB6.TwipsPerPixelX
      SizeY = (RectWorkArea.Bottom - RectWorkArea.Top) * VB6.TwipsPerPixelY
    End If
		
		GetWorkAreaSize = True
		Exit Function
		
ErrorTrap: 
		GetWorkAreaSize = False
		Err.Number = False
		
	End Function
	
	Function LockWindow(ByVal hWnd As Integer) As Boolean
		'Unlock any window currently locked
		UnlockWindow()
		'Lock required window
		LockWindow = LockWindowUpdate(hWnd)
	End Function
	
	Sub SendKeys(ByVal SendString As String, Optional ByVal hWnd As Integer = 0)
		Dim intChar As Short
		Dim intAscii As Short
		Dim intVKey As Short
		Dim intOEMScan As Short
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
			CharToOem(Chr(intAscii), strOEMChar)
			
			'Get the OEM scan code
			intOEMScan = OemKeyScan(Asc(strOEMChar)) And &HFF
			
			'Send the down key
			keybd_event(intVKey, intOEMScan, 0, 0)
			
			'Send the up key
			keybd_event(intVKey, intOEMScan, KEYEVENTF_KEYUP, 0)
		Next intChar
		
	End Sub
	
	Sub SetCaption(ByRef Control As Object, ByVal Caption As String)
		On Error Resume Next
		
		'Attempt to set the Caption property
		'UPGRADE_WARNING: Couldn't resolve default property of object Control.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Control.Caption = Caption
		If Err.Number Then
			'Attempt to set the Text property
			'UPGRADE_WARNING: Couldn't resolve default property of object Control.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Control.Text = Caption
		End If
		Err.Number = False
		
	End Sub
	
	Sub SetFocus(ByVal hWnd As Integer)
		'Set focus to required window
		Call SetFocusAPI(hWnd)
	End Sub
	
	Function SetThickFrame(ByVal hWnd As Integer, ByVal ThickFrame As Boolean) As Object
		Dim WndStyle As Integer
		
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
			'UPGRADE_WARNING: Couldn't resolve default property of object SetThickFrame. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SetThickFrame = True
		End If
		
	End Function
	
	Function UnlockWindow() As Boolean
		'Unlock any window currently locked
		UnlockWindow = LockWindowUpdate(0)
	End Function
	
	Function GetHostName() As String
		Dim sBuffer As String
		Dim lngSize As Integer
		
		'MH20010206
		'W95/W98 not getting Host Name.
		'Need to add one to max computer name length (don't ask!)
		
		'sBuffer = String(MAX_COMPUTERNAME_LENGTH, 0)
		'lngSize = MAX_COMPUTERNAME_LENGTH
		sBuffer = New String(Chr(0), MAX_COMPUTERNAME_LENGTH + 1)
		lngSize = MAX_COMPUTERNAME_LENGTH + 1
		
		If GetComputerName(sBuffer, lngSize) Then
			GetHostName = Left(sBuffer, lngSize)
		Else
			GetHostName = vbNullString
		End If
		
	End Function
	
	Private Function GetX() As Object
		Dim Point As POINTAPI
		Dim retVAL As Integer
		retVAL = GetCursorPos(Point)
		'UPGRADE_WARNING: Couldn't resolve default property of object GetX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetX = Point.x
	End Function
	
	Private Function GetY() As Object
		Dim Point As POINTAPI
		Dim retVAL As Integer
		retVAL = GetCursorPos(Point)
		'UPGRADE_WARNING: Couldn't resolve default property of object GetY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetY = Point.y
	End Function
	
	Private Sub SetClipVars(ByRef MinHeight As Integer, ByRef MinWidth As Integer)
		frmHeight = MinHeight
		frmWidth = MinWidth
	End Sub
	
  'Public Sub ClipForForm(ByRef frm As System.Windows.Forms.Form, ByRef MinHeight As Integer, ByRef MinWidth As Integer)
  '	Dim ResizeREC As RECT
  '	Dim DesktopREC As RECT
  '	Dim GetClipREC As RECT
  '	ResizeREC.Top = (VB6.PixelsToTwipsY(frm.Top) + MinHeight) / VB6.TwipsPerPixelY - 2
  '	ResizeREC.Bottom = (VB6.PixelsToTwipsY(frm.Top) + (VB6.PixelsToTwipsY(frm.Height) - MinHeight)) / VB6.TwipsPerPixelY + 2
  '	ResizeREC.Left_Renamed = (VB6.PixelsToTwipsX(frm.Left) + MinWidth) / VB6.TwipsPerPixelX - 2
  '	ResizeREC.Right_Renamed = (VB6.PixelsToTwipsX(frm.Left) + (VB6.PixelsToTwipsX(frm.Width) - MinWidth)) / VB6.TwipsPerPixelX + 2
  '	If VB6.PixelsToTwipsX(frm.Width) <> frmWidth And frm.WindowState = 0 Then
  '		'UPGRADE_WARNING: Couldn't resolve default property of object retVAL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		retVAL = GetClipCursor(GetClipREC)
  '		'UPGRADE_WARNING: Couldn't resolve default property of object DeskhWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		DeskhWnd = GetDesktopWindow()
  '		'UPGRADE_WARNING: Couldn't resolve default property of object DeskhWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		'UPGRADE_WARNING: Couldn't resolve default property of object retVAL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		retVAL = GetWindowRect(DeskhWnd, DesktopREC)
  '		If GetX > ((VB6.PixelsToTwipsX(frm.Left) + (MinWidth / 2)) / VB6.TwipsPerPixelX) And GetClipREC.Right_Renamed = DesktopREC.Right_Renamed Then
  '			If GetClipREC.Left_Renamed = ResizeREC.Left_Renamed And (GetClipREC.Top = ResizeREC.Top Or GetClipREC.Bottom = ResizeREC.Bottom) Then
  '				frmHeight = VB6.PixelsToTwipsY(frm.Height)
  '				frmWidth = VB6.PixelsToTwipsX(frm.Width)
  '				Exit Sub
  '			ElseIf GetClipREC.Left_Renamed = ResizeREC.Left_Renamed And (GetClipREC.Top <> ResizeREC.Top Or GetClipREC.Bottom <> ResizeREC.Bottom) Then 
  '				If VB6.PixelsToTwipsY(frm.Height) <> frmHeight Then
  '					If GetY > (VB6.PixelsToTwipsY(frm.Top) / VB6.TwipsPerPixelY) + 25 Then
  '						DesktopREC.Top = ResizeREC.Top
  '					Else
  '						DesktopREC.Bottom = ResizeREC.Bottom
  '					End If
  '					frmHeight = VB6.PixelsToTwipsY(frm.Height)
  '				End If
  '			ElseIf GetClipREC.Left_Renamed <> ResizeREC.Left_Renamed And (GetClipREC.Top = ResizeREC.Top Or GetClipREC.Bottom = ResizeREC.Bottom) Then 
  '				If GetClipREC.Top = ResizeREC.Top Then DesktopREC.Top = GetClipREC.Top
  '				If GetClipREC.Bottom = ResizeREC.Bottom Then DesktopREC.Bottom = GetClipREC.Bottom
  '				DesktopREC.Left_Renamed = ResizeREC.Left_Renamed
  '			End If
  '			DesktopREC.Left_Renamed = ResizeREC.Left_Renamed
  '		Else
  '			If GetClipREC.Right_Renamed = ResizeREC.Right_Renamed And (GetClipREC.Bottom = ResizeREC.Bottom Or GetClipREC.Top = ResizeREC.Top) Then
  '				frmHeight = VB6.PixelsToTwipsY(frm.Height)
  '				frmWidth = VB6.PixelsToTwipsX(frm.Width)
  '				Exit Sub
  '			ElseIf GetClipREC.Right_Renamed = ResizeREC.Right_Renamed And (GetClipREC.Bottom <> ResizeREC.Bottom Or GetClipREC.Top <> ResizeREC.Top) Then 
  '				If VB6.PixelsToTwipsY(frm.Height) <> frmHeight Then
  '					If GetY > (VB6.PixelsToTwipsY(frm.Top) / VB6.TwipsPerPixelY) + 25 Then
  '						DesktopREC.Top = ResizeREC.Top
  '					Else
  '						DesktopREC.Bottom = ResizeREC.Bottom
  '					End If
  '					frmHeight = VB6.PixelsToTwipsY(frm.Height)
  '				End If
  '			ElseIf GetClipREC.Right_Renamed <> ResizeREC.Right_Renamed And (GetClipREC.Bottom = ResizeREC.Bottom Or GetClipREC.Top = ResizeREC.Top) Then 
  '				If GetClipREC.Top = ResizeREC.Top Then DesktopREC.Top = GetClipREC.Top
  '				If GetClipREC.Bottom = ResizeREC.Bottom Then DesktopREC.Bottom = GetClipREC.Bottom
  '			End If
  '			DesktopREC.Right_Renamed = ResizeREC.Right_Renamed
  '		End If
  '		'UPGRADE_WARNING: Couldn't resolve default property of object retVAL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		retVAL = ClipCursor(DesktopREC)
  '		frmWidth = VB6.PixelsToTwipsX(frm.Width)
  '	ElseIf VB6.PixelsToTwipsY(frm.Height) <> frmHeight And frm.WindowState = 0 Then 
  '		'UPGRADE_WARNING: Couldn't resolve default property of object retVAL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		retVAL = GetClipCursor(GetClipREC)
  '		'UPGRADE_WARNING: Couldn't resolve default property of object DeskhWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		DeskhWnd = GetDesktopWindow()
  '		'UPGRADE_WARNING: Couldn't resolve default property of object DeskhWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		'UPGRADE_WARNING: Couldn't resolve default property of object retVAL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		retVAL = GetWindowRect(DeskhWnd, DesktopREC)
  '		If GetY > ((VB6.PixelsToTwipsY(frm.Top) + (MinHeight / 2)) / VB6.TwipsPerPixelY) And GetClipREC.Bottom <> ResizeREC.Bottom Then
  '			If GetClipREC.Top = ResizeREC.Top And (GetClipREC.Left_Renamed = ResizeREC.Left_Renamed Or GetClipREC.Right_Renamed = ResizeREC.Right_Renamed) Then
  '				frmHeight = VB6.PixelsToTwipsY(frm.Height)
  '				frmWidth = VB6.PixelsToTwipsX(frm.Width)
  '				Exit Sub
  '			ElseIf GetClipREC.Top = ResizeREC.Top And (GetClipREC.Left_Renamed <> ResizeREC.Left_Renamed Or GetClipREC.Right_Renamed <> ResizeREC.Right_Renamed) Then 
  '				If VB6.PixelsToTwipsX(frm.Width) <> frmWidth Then
  '					If GetX > (VB6.PixelsToTwipsX(frm.Left) / VB6.TwipsPerPixelX) + 15 Then
  '						DesktopREC.Left_Renamed = ResizeREC.Left_Renamed
  '					Else
  '						DesktopREC.Right_Renamed = ResizeREC.Right_Renamed
  '					End If
  '					frmWidth = VB6.PixelsToTwipsX(frm.Width)
  '				End If
  '			ElseIf GetClipREC.Top <> ResizeREC.Top And (GetClipREC.Left_Renamed = ResizeREC.Left_Renamed Or GetClipREC.Right_Renamed = ResizeREC.Right_Renamed) Then 
  '				If GetClipREC.Left_Renamed = ResizeREC.Left_Renamed Then DesktopREC.Left_Renamed = GetClipREC.Left_Renamed
  '				If GetClipREC.Right_Renamed = ResizeREC.Right_Renamed Then DesktopREC.Right_Renamed = GetClipREC.Right_Renamed
  '			End If
  '			DesktopREC.Top = ResizeREC.Top
  '		Else
  '			If GetClipREC.Bottom = ResizeREC.Bottom And (GetClipREC.Right_Renamed = ResizeREC.Right_Renamed Or GetClipREC.Left_Renamed = ResizeREC.Left_Renamed) Then
  '				frmHeight = VB6.PixelsToTwipsY(frm.Height)
  '				frmWidth = VB6.PixelsToTwipsX(frm.Width)
  '				Exit Sub
  '			ElseIf GetClipREC.Bottom = ResizeREC.Bottom And (GetClipREC.Right_Renamed <> ResizeREC.Right_Renamed Or GetClipREC.Left_Renamed <> ResizeREC.Left_Renamed) Then 
  '				If VB6.PixelsToTwipsX(frm.Width) <> frmWidth Then
  '					If GetX > (VB6.PixelsToTwipsX(frm.Left) / VB6.TwipsPerPixelX) + 15 Then
  '						DesktopREC.Left_Renamed = ResizeREC.Left_Renamed
  '					Else
  '						DesktopREC.Right_Renamed = ResizeREC.Right_Renamed
  '					End If
  '					frmWidth = VB6.PixelsToTwipsX(frm.Width)
  '				End If
  '			ElseIf GetClipREC.Bottom <> ResizeREC.Bottom And (GetClipREC.Right_Renamed = ResizeREC.Right_Renamed Or GetClipREC.Left_Renamed = ResizeREC.Left_Renamed) Then 
  '				If GetClipREC.Left_Renamed = ResizeREC.Left_Renamed Then DesktopREC.Left_Renamed = GetClipREC.Left_Renamed
  '				If GetClipREC.Right_Renamed = ResizeREC.Right_Renamed Then DesktopREC.Right_Renamed = GetClipREC.Right_Renamed
  '			End If
  '			DesktopREC.Bottom = ResizeREC.Bottom
  '		End If
  '		'UPGRADE_WARNING: Couldn't resolve default property of object retVAL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
  '		retVAL = ClipCursor(DesktopREC)
  '		frmHeight = VB6.PixelsToTwipsY(frm.Height)
  '	End If
  'End Sub
	
	Public Sub RemoveClipping()
		Dim DesktopREC As RECT
		Dim retVAL, DeskhWnd As Integer
		DeskhWnd = GetDesktopWindow()
		retVAL = GetWindowRect(DeskhWnd, DesktopREC)
		retVAL = ClipCursor(DesktopREC)
	End Sub
	
	
	Public Function DateFormat() As String
		' Returns the date format.
		' NB. Windows allows the user to configure totally stupid
		' date formats (eg. d/M/yyMydy !). This function does not cater
		' for such stupidity, and simply takes the first occurence of the
		' 'd', 'M', 'y' characters.
		Dim sSysFormat As String
		Dim sSysDateSeparator As String
		Dim sDateFormat As String
		Dim iLoop As Short
		Dim fDaysDone As Boolean
		Dim fMonthsDone As Boolean
		Dim fYearsDone As Boolean
		
		fDaysDone = False
		fMonthsDone = False
		fYearsDone = False
		sDateFormat = ""
		
		sSysFormat = GetSystemDateFormat
		sSysDateSeparator = GetSystemDateSeparator
		
		' Loop through the string picking out the required characters.
		For iLoop = 1 To Len(sSysFormat)
			
			Select Case Mid(sSysFormat, iLoop, 1)
				Case "d"
					If Not fDaysDone Then
						' Ensure we have two day characters.
						sDateFormat = sDateFormat & "dd"
						fDaysDone = True
					End If
					
				Case "M"
					If Not fMonthsDone Then
						' Ensure we have two month characters.
						sDateFormat = sDateFormat & "mm"
						fMonthsDone = True
					End If
					
				Case "y"
					If Not fYearsDone Then
						' Ensure we have four year characters.
						sDateFormat = sDateFormat & "yyyy"
						fYearsDone = True
					End If
					
				Case Else
					sDateFormat = sDateFormat & Mid(sSysFormat, iLoop, 1)
			End Select
			
		Next iLoop
		
		' Ensure that all day, month and year parts of the date
		' are present in the format.
		If Not fDaysDone Then
			If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
				sDateFormat = sDateFormat & sSysDateSeparator
			End If
			
			sDateFormat = sDateFormat & "dd"
		End If
		
		If Not fMonthsDone Then
			If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
				sDateFormat = sDateFormat & sSysDateSeparator
			End If
			
			sDateFormat = sDateFormat & "mm"
		End If
		
		If Not fYearsDone Then
			If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
				sDateFormat = sDateFormat & sSysDateSeparator
			End If
			
			sDateFormat = sDateFormat & "yyyy"
		End If
		
		' Return the date format.
		DateFormat = sDateFormat
		
	End Function
End Class