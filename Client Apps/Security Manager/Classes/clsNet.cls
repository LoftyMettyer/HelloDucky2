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

'Utility class to perform network enumeration functions.
'Can locate all servers, printers and shares on a given network.
Option Explicit

Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_REMEMBERED As Long = &H3&

Private Const RESOURCEDISPLAYTYPE_DIRECTORY& = &H9
Private Const RESOURCEDISPLAYTYPE_DOMAIN& = &H1
Private Const RESOURCEDISPLAYTYPE_FILE& = &H4
Private Const RESOURCEDISPLAYTYPE_GENERIC& = &H0
Private Const RESOURCEDISPLAYTYPE_GROUP& = &H5
Private Const RESOURCEDISPLAYTYPE_NETWORK& = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT& = &H7
Private Const RESOURCEDISPLAYTYPE_SERVER& = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE& = &H3
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN& = &H8

Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCETYPE_PRINT As Long = &H2&
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&

Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000

Private Const NO_ERROR = 0
Private Const ERROR_MORE_DATA = 234
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type

Private Type NETRESOURCE_EXTENDED
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    sLocalName As String
    sRemoteName As String
    sComment As String
    sProvider As String
End Type

Private Type NAME_INFO
    Name As Long
End Type

'WNet API resources
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
'Private Declare Function VarPtrAny Lib "vb40032.dll" Alias "VarPtr" (lpObject As Any) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpTo As Any, lpFrom As Any, ByVal lLen As Long)
Private Declare Sub CopyMemByPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpTo As Long, ByVal lpFrom As Long, ByVal lLen As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function NetGetDCName Lib "netapi32.dll" (strServerName As Any, strDomainName As Any, pBuffer As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32.dll" (Buffer As Any) As Long
Private Declare Function NetUserEnum Lib "netapi32" (ByVal servername As Long, ByVal Level As Long, ByVal filter As Long, buf As Any, ByVal PrefMaxLen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal _
    hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, _
    ByVal lpszAttributes As String) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, _
    ByRef lpcbData As Long) As Long

Const REG_SZ = 1
Const KEY_ALL_ACCESS = &H2003F
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Const ODBC_ADD_DSN = 1         ' Add User DSN
Const ODBC_REMOVE_DSN = 3      ' Remove User DSN
Const ODBC_ADD_SYS_DSN = 4     ' Add System DSN
Const ODBC_REMOVE_SYS_DSN = 6  ' Remove System DSN

'Filtered immediate access storage
Private sUserName As String
Private sMachineName As String
Private sServerList As String
Private sPrinterList As String
Private sShareList As String
Private sDirectoryList As String
Private sDomainList As String
Private sFileList As String
Private sGenericList As String
Private sGroupList As String
Private sNetworkList As String
Private sRootList As String
Private sShareAdminList As String

'Sets the resource types that will be enumerated (disk,printer etc)
Private lResType As Long

'Limitation
Private Const MAX_RESOURCES = 256
Private Const NOT_A_CONTAINER = -1

'Stores the results of the enumeration
Private uNetApi(0 To MAX_RESOURCES) As NETRESOURCE
Private uNet() As NETRESOURCE_EXTENDED

Private Sub Class_Initialize()

'Default setting
Call SetResourceType(RESOURCETYPE_DISK)

End Sub

'Reloads the enumerated list of network entities
Public Sub Reset()
    
Dim bFirstTime As Boolean
Dim lReturn As Long
Dim hEnum As Long
Dim lCount As Long
Dim lMin As Long
Dim lLength As Long
Dim l As Long
Dim lBufferSize As Long
Dim lLastIndex As Long

bFirstTime = True
Do
  'Create an enumeration using the required resource type
  If bFirstTime Then
    lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, GetResourceType, RESOURCEUSAGE_ALL, ByVal 0&, hEnum)
    bFirstTime = False
  Else
    If uNet(lLastIndex).dwUsage And RESOURCEUSAGE_CONTAINER Then
      lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, GetResourceType, RESOURCEUSAGE_ALL, uNet(lLastIndex), hEnum)
    Else
      lReturn = NOT_A_CONTAINER
      hEnum = 0
    End If
    lLastIndex = lLastIndex + 1
  End If
  
  'Make sure that we have a good enumeration
  If lReturn = NO_ERROR Then
    lCount = RESOURCE_ENUM_ALL
    'Work through the enumeration until we run out
    Do
      lBufferSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
      lReturn = WNetEnumResource(hEnum, lCount, uNetApi(0), lBufferSize)
      If lCount > 0 Then
        ReDim Preserve uNet(0 To lMin + lCount - 1) As NETRESOURCE_EXTENDED
        For l = 0 To lCount - 1
          'Each Resource will appear here as uNet(i)
          uNet(lMin + l).dwScope = uNetApi(l).dwScope
          uNet(lMin + l).dwType = uNetApi(l).dwType
          uNet(lMin + l).dwDisplayType = uNetApi(l).dwDisplayType
          uNet(lMin + l).dwUsage = uNetApi(l).dwUsage
          
          'Get the name
          If uNetApi(l).pLocalName Then
            lLength = lstrlen(uNetApi(l).pLocalName)
            uNet(lMin + l).sLocalName = Space$(lLength)
            CopyMem ByVal uNet(lMin + l).sLocalName, ByVal uNetApi(l).pLocalName, lLength
          End If
          
          'Get the remote name
          If uNetApi(l).pRemoteName Then
            lLength = lstrlen(uNetApi(l).pRemoteName)
            uNet(lMin + l).sRemoteName = Space$(lLength)
            CopyMem ByVal uNet(lMin + l).sRemoteName, ByVal uNetApi(l).pRemoteName, lLength
          End If
          
          'Get any comment associated with it
          If uNetApi(l).pComment Then
            lLength = lstrlen(uNetApi(l).pComment)
            uNet(lMin + l).sComment = Space$(lLength)
            CopyMem ByVal uNet(lMin + l).sComment, ByVal uNetApi(l).pComment, lLength
          End If
          
          'Get the provider information
          If uNetApi(l).pProvider Then
            lLength = lstrlen(uNetApi(l).pProvider)
            uNet(lMin + l).sProvider = Space$(lLength)
            CopyMem ByVal uNet(lMin + l).sProvider, ByVal uNetApi(l).pProvider, lLength
          End If
        Next l
      End If
      lMin = lMin + lCount
    Loop While lReturn = ERROR_MORE_DATA
  End If
  
  'Check if we have a successfully opened Enumeration
  If hEnum Then
      l = WNetCloseEnum(hEnum)
  End If
  
Loop While lLastIndex < lMin

'Decode the results
Call DecodeEnum

End Sub

Private Sub DecodeLocalInfo()

On Error Resume Next

'Create a buffer
sUserName = String(255, Chr(0))

'Get the username
Call GetUserName(sUserName, 255)

'Strip the rest of the buffer
sUserName = Left(sUserName, InStr(sUserName, Chr(0)) - 1)

'Create a buffer
sMachineName = String(255, Chr(0))
Call GetComputerName(sMachineName, 255)

'Remove the unnecessary chr(0)'s
sMachineName = Left$(sMachineName, InStr(1, sMachineName, Chr(0)) - 1)

End Sub

'Sets the resource type to be enumerated
'(restricting this greatly improves performance)
'0 RESOURCETYPE_ANY = Everything
'1 RESOURCETYPE_DISK = Machines and shares only
'2 RESOURCETYPE_PRINT = Printers only
'Default is RESOURCETYPE_DISK
Public Sub SetResourceType(lResourceType As Long)

lResType = lResourceType

End Sub

'Returns the current resource type that is being scanned for
Public Function GetResourceType() As Long

GetResourceType = lResType

End Function

'Decodes the network array into a useful set of values
Private Sub DecodeEnum()
  
Dim l As Long

If UBound(uNet) > 0 Then
  'Get some local information
  Call DecodeLocalInfo
  
  'Parse the network enumeration
  For l = 0 To UBound(uNet)
    'TODO: Include comments? uNet(l).sComment
    Select Case uNet(l).dwDisplayType
    Case RESOURCEDISPLAYTYPE_DIRECTORY&
      sDirectoryList = sDirectoryList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_DOMAIN
      sDomainList = sDomainList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_FILE
      sFileList = sFileList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_GENERIC
      sGenericList = sGenericList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_GROUP
      sGroupList = sGroupList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_NETWORK&
      sNetworkList = sNetworkList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_ROOT&
      sRootList = sRootList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_SERVER
      sServerList = sServerList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_SHARE
      sShareList = sShareList + uNet(l).sRemoteName + "|"
    Case RESOURCEDISPLAYTYPE_SHAREADMIN&
      sShareAdminList = sShareAdminList + uNet(l).sRemoteName + "|"
    End Select
  Next l
End If

End Sub

'Provides the current user name
Public Function GetLocalUserName() As String

GetLocalUserName = sUserName

End Function

'Provides the current machine name
Public Function GetLocalMachineName() As String

GetLocalMachineName = sMachineName

End Function

'Returns a | delimited list of all computers on the network
Public Function GetServerList() As String

GetServerList = sServerList

End Function

'Returns a | delimited list of all shares on the network
Public Function GetShareList() As String

GetShareList = sShareList

End Function

'Returns a | delimited list of all printers on the network
Public Function GetPrinterList() As String

GetPrinterList = sPrinterList

End Function

'Returns a | delimited list of all directories on the network
Public Function GetDirectoryList() As String

GetDirectoryList = sDirectoryList

End Function

'Returns a | delimited list of all files on the network
Public Function GetFileList() As String

GetFileList = sFileList

End Function

'Returns a | delimited list of all generic items on the network
Public Function GetGenericList() As String

GetGenericList = sGenericList

End Function

'Returns a | delimited list of all groups on the network
Public Function GetGroupList() As String

GetGroupList = sGroupList

End Function

'Returns a | delimited list of all sub networks on the network
Public Function GetNetworkList() As String

GetNetworkList = sNetworkList

End Function

'Returns a | delimited list of all roots on the network
Public Function GetRootList() As String

GetRootList = sRootList

End Function

'Returns a | delimited list of all admin shares on the network
Public Function GetShareAdminList() As String

GetShareAdminList = sShareAdminList

End Function

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

Private Function WinPointerToString(ByVal StringPtr As Long) As String

Dim labyteBuffer() As Byte     ' Holds the data as a Byte Form
Dim llLength As Long           ' Holds the amount of data

'*** Errors passed up to calling routine

' If it ain't Zero
If StringPtr Then
    ' Get the amount of data
    llLength = lstrlenW(StringPtr) * 2
    ' If it ain't Zero
    If llLength Then
        ' Set the Buffer
        ReDim labyteBuffer(0 To (llLength - 1)) As Byte
        ' Fill it
        RtlMoveMemory labyteBuffer(0), ByVal StringPtr, llLength
        ' Set the return value
        WinPointerToString = labyteBuffer
    End If
End If

End Function


Public Function GetPrimaryDomainController(pstrDomainName As String) As String

Dim lsServer As String   ' Holds the name of the server
Dim llBuffer As Long     ' Holds pointer to name
Dim laByte() As Byte     ' Holds name
Dim llGotNameOK As Long  ' Reply from the API call
Dim llBufferOK As Long   ' Reply from clear buffer call
Dim DNArray() As Byte
   
On Error GoTo Catch
   
Const SUCCESS = 0

' This will return nothing if the machine is not in a domain
DNArray = pstrDomainName & vbNullChar
llGotNameOK = NetGetDCName(vbNullString, DNArray(0), llBuffer)

If llGotNameOK = SUCCESS Then ' success
   
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
   
Finally:
    ' Clean-up
    Exit Function

Catch:

    ' Reports back to developer if in debugmode
    Resume Finally

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

