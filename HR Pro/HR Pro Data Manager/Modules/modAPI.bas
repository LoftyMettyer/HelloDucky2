Attribute VB_Name = "modAPI"
' Kernel API functions
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpoperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


'ODBC API functions
Public Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hdbc&, phstmt&) As Integer
Public Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hstmt&, ByVal fOption%) As Integer
Public Declare Function SQLGetData Lib "odbc32.dll" (ByVal hstmt&, ByVal icol%, ByVal fCType%, ByVal rgbValue As String, ByVal cbValueMax&, pcbValue&) As Integer
Public Declare Function SQLColumns Lib "odbc32.dll" (ByVal hstmt&, ByVal szTblQualifier As String, ByVal cbTblQualifier%, ByVal szTblOwner As String, ByVal cbTblOwner%, ByVal szTblName As String, ByVal cbTblName%, ByVal szColName As String, ByVal cbColName%) As Integer
Public Declare Function SQLBindCol Lib "odbc32.dll" (ByVal hstmt&, ByVal icol%, ByVal fCType%, rgbValue As Any, ByVal cbValueMax&, pcbValue&) As Integer
Public Declare Function SQLFetch Lib "odbc32.dll" (ByVal hstmt&) As Integer
Public Declare Function SQLGetInfo Lib "odbc32.dll" (ByVal hdbc&, ByVal fInfoType%, ByRef rgbInfoValue As Any, ByVal cbInfoMax%, cbInfoOut%) As Integer
Public Declare Function SQLGetInfoString Lib "odbc32.dll" Alias "SQLGetInfo" (ByVal hdbc&, ByVal fInfoType%, ByVal rgbInfoValue As String, ByVal cbInfoMax%, cbInfoOut%) As Integer

' Network API functions
Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

' Menu API functions
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

' Window API functions
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long, ByVal Flags As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

' Registy API functions
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long

' Keyboard API functions
Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer

' Misc API functions
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMetricsAPI Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

' Mouse API functions
Public Declare Function ClipCursor Lib "user32.dll" (lpRect As Rect) As Long
Public Declare Function GetClipCursor Lib "user32.dll" (lprc As Rect) As Long

' Icon extracting API functions
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

' Font handling API functions
Public Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

'Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
'Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'Private Const MF_BYPOSITION = &H400&
'Private Const MF_REMOVE = &H1000&


' API constants
Public Const ERROR_SUCCESS = 0

' Registry API constants
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const PROCESS_QUERY_INFORMATION = &H400

Public Const EM_GETLINECOUNT  As Long = &HBA



Public Const WM_LBUTTONDBLCLK = &H203


' Window constants
Public Const SC_CLOSE As Long = &HF060&
Public Const SC_MAXIMIZE As Long = &HF030&
Public Const SC_MINIMIZE As Long = &HF020&

Public Const xSC_CLOSE As Long = -10&
Public Const xSC_MAXIMIZE As Long = -11&
Public Const xSC_MINIMIZE As Long = -12&

Public Const MIIM_STATE As Long = &H1&
Public Const MIIM_ID As Long = &H2&
Public Const MFS_GRAYED As Long = &H3&
Public Const WM_NCACTIVATE As Long = &H86

Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100

' Keyboard constants
Public Const KEYEVENTF_KEYUP = &H2

Public Const MAX_COMPUTERNAME_LENGTH = 15

Public Const LOCALE_SYSTEM_DEFAULT = &H800
Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_SDATE = &H1D            ' date separator
Public Const LOCALE_STIME = &H1E            ' time separator
Public Const LOCALE_SSHORTDATE = &H1F       ' short date format string
Public Const LOCALE_SLONGDATE = &H20        ' long date format string
Public Const LOCALE_SDECIMAL = &HE          ' decimal separator
Public Const LOCALE_STHOUSAND = &HF         ' thousand separator
Public Const LOCALE_IMEASURE = &HD          ' Measurement System

'Windows version constants
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'Combo box constants
Public Const CB_ERR = (-1)
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_SELECTSTRING = &H14D

'List box constants
Public Const LB_ERR = (-1)
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SELECTSTRING = &H18C

'Window constants
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_FRAMECHANGED = &H20


Public Const SPI_GETWORKAREA = 48

'ChildWindowFromPointEx constants
Public Const CWP_ALL = 0
Public Const CWP_SKIPINVISIBLE = 1
Public Const CWP_SKIPDISABLED = 2
Public Const CWP_SKIPTRANSPARENT = 4

' Icon constants
Public Const WM_SETICON = &H80
Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1

' Window constants
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_EX_WINDOWEDGE As Long = &H100
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_STYLE As Long = (-16)

' File generation constants
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const ILD_TRANSPARENT = &H1
Public Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Const LB_SETHORIZONTALEXTENT = &H194

' Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

' tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4
Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

' EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

' ODBC constants
Public Const SQL_DROP As Long = 1
Public Const SQL_NTS As Long = -3
Public Const SQL_NULL_DATA As Long = -1

' ODBC datatype constants
Public Const SQL_SIGNED_OFFSET As Long = -20
Public Const SQL_CHAR As Long = 1
Public Const SQL_INTEGER As Long = 4
Public Const SQL_SMALLINT As Long = 5
Public Const SQL_C_CHAR As Long = SQL_CHAR
Public Const SQL_C_LONG As Long = SQL_INTEGER
Public Const SQL_C_SHORT As Long = SQL_SMALLINT
Public Const SQL_C_SSHORT As Long = SQL_C_SHORT + SQL_SIGNED_OFFSET
Public Const SQL_C_SLONG As Long = SQL_C_LONG + SQL_SIGNED_OFFSET

'ODBC return value constants
Public Const SQL_SUCCESS As Long = 0
Public Const SQL_SUCCESS_WITH_INFO As Long = 1

Public Const STR_LEN = 254


' Type definitions
Public Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type


Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type TEXTMETRIC
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

Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type

Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type NEWTEXTMETRIC
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
  ntmFlags As Long
  ntmSizeEM As Long
  ntmCellHeight As Long
  ntmAveWidth As Long
End Type


'Public Sub HideSystemMenu(frm As Form)
'
'  Dim hMenu As Long
'  Dim lngCount As Long
'  Dim lngIndex As Long
'
'  hMenu = GetSystemMenu(frm.hWnd, 0)
'  lngCount = GetMenuItemCount(hMenu)
'  For lngIndex = 0 To lngCount
'    RemoveMenu hMenu, lngIndex, MF_REMOVE Or MF_BYPOSITION
'  Next
'
'End Sub
'
