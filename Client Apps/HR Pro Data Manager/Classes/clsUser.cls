VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Win32 APIs to determine OS information.
'
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
'
' Win32 NetAPIs.
'
Private Declare Function NetGetDCName Lib "netapi32.dll" (strServerName As Any, strDomainName As Any, pBuffer As Long) As Long
Private Declare Function NetUserChangePassword Lib "netapi32" (Domain As Any, User As Any, OldPass As Byte, NewPass As Byte) As Long
Private Declare Function NetUserGetInfo Lib "netapi32" (lpServer As Any, UserName As Byte, ByVal Level As Long, lpBuffer As Long) As Long
Private Declare Function NetUserGetGroups Lib "netapi32" (lpServer As Any, UserName As Byte, ByVal Level As Long, lpBuffer As Long, ByVal PrefMaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long) As Long
Private Declare Function NetUserGetLocalGroups Lib "netapi32" (lpServer As Any, UserName As Byte, ByVal Level As Long, ByVal Flags As Long, lpBuffer As Long, ByVal MaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal pBuffer As Long) As Long
Private Declare Function NetRemoteTOD Lib "netapi32.dll" (yServer As Any, pBuffer As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameW Lib "advapi32.dll" (lpBuffer As Byte, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long

' Kernel32 API Declares
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Byte, ByVal lpString2 As Long) As Long


Private Type USER_INFO_3_API
   Name As Long
   Password As Long
   PasswordAge As Long
   Privilege As Long
   HomeDir As Long
   Comment As Long
   Flags As Long
   ScriptPath As Long
   AuthFlags As Long
   FullName As Long
   UserComment As Long
   Parms As Long
   Workstations As Long
   LastLogon As Long
   LastLogoff As Long
   AcctExpires As Long
   MaxStorage As Long
   UnitsPerWeek As Long
   LogonHours As Long
   BadPwCount As Long
   NumLogons As Long
   LogonServer As Long
   CountryCode As Long
   CodePage As Long
   UserID As Long
   PrimaryGroupID As Long
   Profile As Long
   HomeDirDrive As Long
   PasswordExpired As Long
End Type

Private Type USER_INFO_3
   Name As String
   Password As String
   PasswordAge As Long
   Privilege As Long
   HomeDir As String
   Comment As String
   Flags As Long
   ScriptPath As String
   AuthFlags As Long
   FullName As String
   UserComment As String
   Parms As String
   Workstations As String
   LastLogon As Long
   LastLogoff As Long
   AcctExpires As Long
   MaxStorage As Long
   UnitsPerWeek As Long
   LogonHours(0 To 20) As Byte
   BadPwCount As Long
   NumLogons As Long
   LogonServer As String
   CountryCode As Long
   CodePage As Long
   UserID As Long
   PrimaryGroupID As Long
   Profile As String
   HomeDirDrive As String
   PasswordExpired As Boolean
End Type

Private Type GROUP_INFO_2_API
   Name As Long
   Comment As Long
   GroupID As Long
   Attributes As Long
End Type

Private Type GROUP_INFO_2
   Name As String
   Comment As String
   GroupID As Long
   Attributes As Long
End Type

Private Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Private Const FILTER_NORMAL_ACCOUNT As Long = &H2&
Private Const FILTER_PROXY_ACCOUNT As Long = &H4&
Private Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
Private Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
Private Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&

Private Const NERR_Success As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234&

Private Const TIMEQ_FOREVER = -1&             '((unsigned long) -1L)
Private Const USER_MAXSTORAGE_UNLIMITED = -1& '((unsigned long) -1L)
Private Const USER_NO_LOGOFF = -1&            '((unsigned long) -1L)
Private Const UNITS_PER_DAY = 24
Private Const UNITS_PER_WEEK = UNITS_PER_DAY * 7

Private Const USER_PRIV_MASK = 3
Private Const USER_PRIV_GUEST = 0
Private Const USER_PRIV_USER = 1
Private Const USER_PRIV_ADMIN = 2

Private Const UNLEN = 256         ' Maximum username length
Private Const GNLEN = UNLEN       ' Maximum groupname length
Private Const CNLEN = 15          ' Maximum computer name length
Private Const MAXCOMMENTSZ = 256  ' Multipurpose comment length
Private Const LG_INCLUDE_INDIRECT As Long = &H1&

Private m_UserInfo As USER_INFO_3
Private m_UserName As String
Private m_Server As String
Private m_Groups() As String
Private m_LocalGroups() As String
Private m_IsWinNT As Boolean
Private mstrDomainController As String

' *********************************************************
'  Initialization
' *********************************************************
Private Sub Class_Initialize()
   Dim os As OSVERSIONINFO
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   
   If os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      m_IsWinNT = True
      Server = Environ("LOGONSERVER")     'GetPrimaryDomainController(Environ("USERDOMAIN"))
      UserName = Environ("USERNAME")
   End If
End Sub

' *********************************************************
'  Public Properties
' *********************************************************
Public Property Get UserName() As String
   UserName = m_UserInfo.Name
End Property

Public Property Let UserName(NewVal As String)
   m_UserName = NewVal
   Me.Refresh
End Property

Public Property Get Server() As String
   Server = m_Server
End Property

Public Property Let Server(NewVal As String)
   m_Server = NewVal
End Property

Public Property Get GroupCount() As Long
   GroupCount = UBound(m_Groups) + 1
End Property

Public Property Get Group(ByVal Index As Long) As String
   If Index >= LBound(m_Groups) And Index <= UBound(m_Groups) Then
      Group = m_Groups(Index)
   End If
End Property

Public Property Get LocalGroupCount() As Long
   LocalGroupCount = UBound(m_LocalGroups) + 1
End Property

Public Property Get LocalGroup(ByVal Index As Long) As String
   If Index >= LBound(m_LocalGroups) And Index <= UBound(m_LocalGroups) Then
      LocalGroup = m_LocalGroups(Index)
   End If
End Property

Public Property Get Password()
   Err.Raise Number:=vbObjectError + 1, _
             Source:="CNetUser.Password", _
             Description:="Password property is write-only for security."
End Property

Public Property Get PasswordAge() As Long
   PasswordAge = m_UserInfo.PasswordAge
End Property

Public Property Get Privilege() As Long
   Privilege = m_UserInfo.Privilege
End Property

Public Property Get HomeDir() As String
   HomeDir = m_UserInfo.HomeDir
End Property

Public Property Get Comment() As String
   Comment = m_UserInfo.Comment
End Property

Public Property Get Flags() As Long
   Flags = m_UserInfo.Flags
End Property

Public Property Get ScriptPath() As String
   ScriptPath = m_UserInfo.ScriptPath
End Property

Public Property Get AuthFlags() As Long
   AuthFlags = m_UserInfo.AuthFlags
End Property

Public Property Get FullName() As String
   FullName = m_UserInfo.FullName
End Property

Public Property Get UserComment() As String
   UserComment = m_UserInfo.UserComment
End Property

Public Property Get Parms() As String
   Parms = m_UserInfo.Parms
End Property

Public Property Get Workstations() As String
   Workstations = m_UserInfo.Workstations
End Property

Public Property Get LastLogon() As Long
    LastLogon = m_UserInfo.LastLogon
End Property

Public Property Get LastLogoff() As Long
   LastLogoff = m_UserInfo.LastLogoff
End Property

Public Property Get LastLogonDate() As Double
    LastLogonDate = NetTimeToVbTime(m_UserInfo.LastLogon)
End Property

Public Property Get LastLogoffDate() As Double
   LastLogoffDate = NetTimeToVbTime(m_UserInfo.LastLogoff)
End Property

Public Property Get AcctExpires() As Long
   AcctExpires = m_UserInfo.AcctExpires
End Property

Public Property Get AcctExpiresDate() As Long
   AcctExpiresDate = NetTimeToVbTime(m_UserInfo.AcctExpires)
End Property

Public Property Get MaxStorage() As Long
   MaxStorage = m_UserInfo.MaxStorage
End Property

Public Property Get UnitsPerWeek() As Long
   UnitsPerWeek = m_UserInfo.UnitsPerWeek
End Property

Public Property Get LogonHours(ByVal Index As Long) As Byte
   If Index >= 0 And Index <= 20 Then
      LogonHours = m_UserInfo.LogonHours(Index)
   End If
End Property

Public Property Get BadPasswordCount() As Long
   BadPasswordCount = m_UserInfo.BadPwCount
End Property

Public Property Get NumLogons() As Long
   NumLogons = m_UserInfo.NumLogons
End Property

Public Property Get LogonServer() As String
   LogonServer = m_UserInfo.LogonServer
End Property

Public Property Get CountryCode() As Long
   CountryCode = m_UserInfo.CountryCode
End Property

Public Property Get CodePage() As Long
   CodePage = m_UserInfo.CodePage
End Property

Public Property Get UserID() As Long
   UserID = m_UserInfo.UserID
End Property

Public Property Get PrimaryGroupID() As Long
   PrimaryGroupID = m_UserInfo.PrimaryGroupID
End Property

Public Property Get Profile() As String
   Profile = m_UserInfo.Profile
End Property

Public Property Get HomeDirDrive() As String
   HomeDirDrive = m_UserInfo.HomeDirDrive
End Property

Public Property Get PasswordExpired() As Boolean
   PasswordExpired = m_UserInfo.PasswordExpired
End Property

' *********************************************************
'  Public Methods
' *********************************************************
Public Function Refresh() As Boolean
   Dim lpBuffer As Long
   Dim yUserName() As Byte
   Dim yServer() As Byte
   Dim uUserApi As USER_INFO_3_API
   Dim nRet As Long
   
   yUserName = m_UserName & vbNullChar
   If m_Server = "" Then
      nRet = NetUserGetInfo(ByVal 0&, yUserName(0), 3, lpBuffer)
   Else
      If InStr(m_Server, "\\") = 1 Then
         yServer = m_Server & vbNullChar
      Else
         yServer = "\\" & m_Server & vbNullChar
      End If
      nRet = NetUserGetInfo(yServer(0), yUserName(0), 3, lpBuffer)
   End If
   
   If nRet = NERR_Success Then
      CopyMem uUserApi, ByVal lpBuffer, Len(uUserApi)
      '
      ' Transfer data to VB structure
      '
      m_UserInfo.Name = PointerToStringW(uUserApi.Name)
      m_UserInfo.Password = PointerToStringW(uUserApi.Password)
      m_UserInfo.PasswordAge = uUserApi.PasswordAge
      m_UserInfo.Privilege = uUserApi.Privilege
      m_UserInfo.HomeDir = PointerToStringW(uUserApi.HomeDir)
      m_UserInfo.Comment = PointerToStringW(uUserApi.Comment)
      m_UserInfo.Flags = uUserApi.Flags
      m_UserInfo.ScriptPath = PointerToStringW(uUserApi.ScriptPath)
      m_UserInfo.AuthFlags = uUserApi.AuthFlags
      m_UserInfo.FullName = PointerToStringW(uUserApi.FullName)
      m_UserInfo.UserComment = PointerToStringW(uUserApi.UserComment)
      m_UserInfo.Parms = PointerToStringW(uUserApi.Parms)
      m_UserInfo.Workstations = PointerToStringW(uUserApi.Workstations)
      m_UserInfo.LastLogon = uUserApi.LastLogon
      m_UserInfo.LastLogoff = uUserApi.LastLogoff
      m_UserInfo.AcctExpires = uUserApi.AcctExpires
      m_UserInfo.MaxStorage = uUserApi.MaxStorage
      m_UserInfo.UnitsPerWeek = uUserApi.UnitsPerWeek
      CopyMem m_UserInfo.LogonHours(0), ByVal uUserApi.LogonHours, 21
      m_UserInfo.BadPwCount = uUserApi.BadPwCount
      m_UserInfo.NumLogons = uUserApi.NumLogons
      m_UserInfo.LogonServer = PointerToStringW(uUserApi.LogonServer)
      m_UserInfo.CountryCode = uUserApi.CountryCode
      m_UserInfo.CodePage = uUserApi.CodePage
      m_UserInfo.UserID = uUserApi.UserID
      m_UserInfo.PrimaryGroupID = uUserApi.PrimaryGroupID
      m_UserInfo.Profile = PointerToStringW(uUserApi.Profile)
      m_UserInfo.HomeDirDrive = PointerToStringW(uUserApi.HomeDirDrive)
      m_UserInfo.PasswordExpired = CBool(uUserApi.PasswordExpired)
      '
      ' Return success
      '
      Refresh = True
   End If
   '
   ' Clean up
   '
   If lpBuffer Then
      Call NetApiBufferFree(lpBuffer)
      RefreshGroups
      RefreshLocalGroups
   End If
End Function

Public Function NetTimeToVbTime(NetDate As Long) As Double
   Const BaseDate# = 25569   'DateSerial(1970, 1, 1)
   Const SecsPerDay# = 86400
   NetTimeToVbTime = BaseDate + (CDbl(NetDate) / SecsPerDay)
End Function

' *********************************************************
'  Private Methods
' *********************************************************
Private Sub RefreshLocalGroups()
   Dim lpBuffer As Long
   Dim yUserName() As Byte
   Dim yServer() As Byte
   Dim lpGroups() As Long
   Dim nRead As Long
   Dim nTotal As Long
   Dim nRet As Long
   Dim i As Long
   'Const Flags& = LG_INCLUDE_INDIRECT 'To get domain groups as well
   Const Flags& = 0
   
   yUserName = m_UserName & vbNullChar
   If m_Server = "" Then
      nRet = NetUserGetLocalGroups(ByVal 0&, yUserName(0), 0, Flags, lpBuffer, &H4000, nRead, nTotal)
   Else
      If InStr(m_Server, "\\") = 1 Then
         yServer = m_Server & vbNullChar
      Else
         yServer = "\\" & m_Server & vbNullChar
      End If
      nRet = NetUserGetLocalGroups(yServer(0), yUserName(0), 0, Flags, lpBuffer, &H400, nRead, nTotal)
   End If
      
'   If nRet = NERR_Success Then
'      ReDim lpGroups(0 To nRead - 1) As Long
'      ReDim m_LocalGroups(0 To nRead - 1) As String
'      CopyMem lpGroups(0), ByVal lpBuffer, nRead * 4
'      For i = 0 To nRead - 1
'         m_LocalGroups(i) = PointerToStringW(lpGroups(i))
'      Next i
'   End If
   '
   ' Clean up
   '
   If lpBuffer Then
      Call NetApiBufferFree(lpBuffer)
   End If
End Sub

Private Sub RefreshGroups()
   Dim lpBuffer As Long
   Dim yUserName() As Byte
   Dim yServer() As Byte
   Dim lpGroups() As Long
   Dim nRead As Long
   Dim nTotal As Long
   Dim nRet As Long
   Dim i As Long
   
   yUserName = m_UserName & vbNullChar
   If m_Server = "" Then
      nRet = NetUserGetGroups(ByVal 0&, yUserName(0), 0, lpBuffer, &H4000, nRead, nTotal)
   Else
      If InStr(m_Server, "\\") = 1 Then
         yServer = m_Server & vbNullChar
      Else
         yServer = "\\" & m_Server & vbNullChar
      End If
      nRet = NetUserGetGroups(yServer(0), yUserName(0), 0, lpBuffer, &H400, nRead, nTotal)
   End If
      
   If nRet = NERR_Success Then
      ReDim lpGroups(0 To nRead - 1) As Long
      ReDim m_Groups(0 To nRead - 1) As String
      CopyMem lpGroups(0), ByVal lpBuffer, nRead * 4
      For i = 0 To nRead - 1
         m_Groups(i) = PointerToStringW(lpGroups(i))
      Next i
   End If
   '
   ' Clean up
   '
   If lpBuffer Then
      Call NetApiBufferFree(lpBuffer)
   End If
End Sub

Private Function PointerToStringW(lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMem Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function

Private Function PointerToDWord(lpDWord As Long) As Long
   Dim nRet As Long
   If lpDWord Then
      CopyMem nRet, ByVal lpDWord, 4
      PointerToDWord = nRet
   End If
End Function

Private Function CurrentUserName() As String
   Dim Buffer As String
   Dim yBuffer() As Byte
   Dim nRet As Long
   Dim nLen As Long
   Const NameLength = UNLEN + 1
   
   nLen = NameLength * 2
   ReDim yBuffer(0 To nLen - 1) As Byte
   If GetUserNameW(yBuffer(0), nLen) Then
      Buffer = yBuffer
      CurrentUserName = Left(Buffer, nLen - 1)
   End If
End Function

Private Function CurrentMachineName() As String
   Dim Buffer As String
   Dim yBuffer() As Byte
   Dim nRet As Long
   Dim nLen As Long
   Const NameLength = CNLEN + 1
   
   nLen = NameLength * 2
   ReDim yBuffer(0 To nLen - 1) As Byte
   If GetComputerNameW(yBuffer(0), nLen) Then
      Buffer = yBuffer
      CurrentMachineName = Left(Buffer, nLen)
   End If
End Function

' JDM - NOTE : This function only works 66.66% of the time (avoid using if you value your sanity!!!)
Private Function GetPrimaryDomainController(ByVal pstrDomain As String) As String

  Dim lsServer As String   ' Holds the name of the server
  Dim llBuffer As Long     ' Holds pointer to name
  Dim laByte() As Byte     ' Holds name
  Dim llGotNameOK As Long  ' Reply from the API call
  Dim llBufferOK As Long   ' Reply from clear buffer call
  Dim aDomain() As Byte    ' Pass in required domain
  
  ' Put domain into byte array
  aDomain() = pstrDomain & vbNullString
     
  ' This will return nothing if the machine is not in a domain
  llGotNameOK = NetGetDCName(vbNullString, aDomain(0), llBuffer)

  If llGotNameOK = 0 Then ' success
   
    ReDim laByte(256)
    
    ' llBuffer is a pointer so copy contents using API call
    RtlMoveMemory laByte(0), llBuffer, 256
    ' Convert from byte array to string
    lsServer = laByte
    ' Set to return name strip off trailing rubbish
    GetPrimaryDomainController = NullTrim(lsServer)
    ' Free llBuffer - not in other samples but mentioned in documentation
    llBufferOK = NetApiBufferFree(llBuffer)
    
  Else
    GetPrimaryDomainController = vbNullString
  End If

End Function

Private Function NullTrim(ByVal Convert As Variant) As String

  Dim i As Integer          ' Counter
  Dim lsString  As String   ' Holds variable once converted to string

  '*** Errors passed up to calling routine

  ' Sort out any required convertion
  lsString = vbNullString
  Select Case VarType(Convert)
      Case vbByte, vbArray + vbByte
          '*** [GW need some work here]
          'lsString = CStr(Convert)
      Case vbString
          lsString = Convert
      Case vbEmpty
      Case vbNull
      Case vbInteger
          lsString = CStr(Convert)
      Case vbLong
          lsString = CStr(Convert)
      Case vbSingle
          lsString = CStr(Convert)
      Case vbDouble
          lsString = CStr(Convert)
      Case vbCurrency
          lsString = CStr(Convert)
      Case vbDate
          lsString = Format$(Convert, "General Date")
      Case vbBoolean
          lsString = Convert
      Case vbObject
      Case vbError
      Case vbVariant
      Case vbDataObject
      Case vbArray
  End Select
  
  ' Now check for any Null terminate
  i = InStr(lsString, vbNullChar)
  If i > 0 Then
      NullTrim = Trim(Left$(lsString, i - 1))
  Else
      NullTrim = Trim(lsString)
  End If
    
End Function
