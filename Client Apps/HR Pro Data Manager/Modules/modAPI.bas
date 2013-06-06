Attribute VB_Name = "modAPI"
' Windows API functions
Public Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpoperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'ODBC API functions
Public Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hdbc&, phstmt&) As Integer
Public Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hstmt&, ByVal fOption%) As Integer
Public Declare Function SQLGetData Lib "odbc32.dll" (ByVal hstmt&, ByVal icol%, ByVal fCType%, ByVal rgbValue As String, ByVal cbValueMax&, pcbValue&) As Integer
Public Declare Function SQLColumns Lib "odbc32.dll" (ByVal hstmt&, ByVal szTblQualifier As String, ByVal cbTblQualifier%, ByVal szTblOwner As String, ByVal cbTblOwner%, ByVal szTblName As String, ByVal cbTblName%, ByVal szColName As String, ByVal cbColName%) As Integer
Public Declare Function SQLBindCol Lib "odbc32.dll" (ByVal hstmt&, ByVal icol%, ByVal fCType%, rgbValue As Any, ByVal cbValueMax&, pcbValue&) As Integer
Public Declare Function SQLFetch Lib "odbc32.dll" (ByVal hstmt&) As Integer
Public Declare Function SQLGetInfo Lib "odbc32.dll" (ByVal hdbc&, ByVal fInfoType%, ByRef rgbInfoValue As Any, ByVal cbInfoMax%, cbInfoOut%) As Integer
Public Declare Function SQLGetInfoString Lib "odbc32.dll" Alias "SQLGetInfo" (ByVal hdbc&, ByVal fInfoType%, ByVal rgbInfoValue As String, ByVal cbInfoMax%, cbInfoOut%) As Integer
