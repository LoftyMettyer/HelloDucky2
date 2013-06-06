VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Net"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetUserName Lib "ADVAPI32.DLL" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal _
    hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, _
    ByVal lpszAttributes As String) As Long

Const REG_SZ = 1
Const KEY_ALL_ACCESS = &H2003F
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Const ODBC_ADD_DSN = 1         ' Add User DSN
Const ODBC_REMOVE_DSN = 3      ' Remove User DSN
Const ODBC_ADD_SYS_DSN = 4     ' Add System DSN
Const ODBC_REMOVE_SYS_DSN = 6  ' Remove System DSN

Function UserName() As String
  Dim sName As String
  Dim lNameLen As Long
  
  sName = String(255, 0)
  lNameLen = Len(sName)
  GetUserName sName, lNameLen
  
  UserName = IIf(lNameLen < 1, vbNullString, Left(sName, lNameLen))
End Function

Function ComputerName() As String
  Dim sName As String
  Dim lNameLen As Long
  
  sName = String(255, 0)
  lNameLen = Len(sName)
  GetComputerName sName, lNameLen
  
  ComputerName = IIf(lNameLen < 1, vbNullString, Left(sName, lNameLen))
End Function


Public Function MakeDSN(ByVal strServer As String, ByVal strDatabase As String) As Boolean

  On Error GoTo ErrorTrap

  Dim sAttributes As String
  Dim sDBQ As String
  Dim lngRet As Long
  
  Dim hKey As Long
  Dim regValue As String
  Dim valueType As Long

  ' query the Registry to check whether the DSN is already installed
  ' open the key
  If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBC.INI\" & gstrWindowsAuthentication_DNSName, 0, _
    KEY_ALL_ACCESS, hKey) = 0 Then
    
    regValue = String$(1024, 0)
    
    If RegQueryValueEx(hKey, "DBQ", 0, valueType, regValue, Len(regValue)) = 0 Then
      If valueType = REG_SZ Then
        sDBQ = Left$(regValue, InStr(regValue, vbNullChar) - 1)
      End If
    End If

    RegCloseKey hKey
  End If

  ' Perform the action only if we're adding a DSN that doesn't exist
  ' or removing and existing DSN
  If sDBQ = "" Then
    sAttributes = "DSN=" & gstrWindowsAuthentication_DNSName _
                & vbNullChar & "Server=" & strServer _
                & vbNullChar & "DESCRIPTION=HRPro connection to database" _
                & vbNullChar & "DATABASE=" & strDatabase _
                & vbNullChar & "TRUSTED_CONNECTION=yes"

    lngRet = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, "SQL Server", sAttributes)
  End If
  
  MakeDSN = True
  Exit Function
  
ErrorTrap:
  MakeDSN = False
  
End Function


