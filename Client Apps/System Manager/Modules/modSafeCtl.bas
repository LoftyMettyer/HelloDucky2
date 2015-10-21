Attribute VB_Name = "modSafeCtl"
Option Explicit

Public Const IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
Public Const IID_IPersistStorage = _
  "{0000010A-0000-0000-C000-000000000046}"
Public Const IID_IPersistStream = _
  "{00000109-0000-0000-C000-000000000046}"
Public Const IID_IPersistPropertyBag = _
  "{37D84F60-42CB-11CE-8135-00AA004BB851}"

Public Const INTERFACESAFE_FOR_UNTRUSTED_CALLER = &H1
Public Const INTERFACESAFE_FOR_UNTRUSTED_DATA = &H2
Public Const E_NOINTERFACE = &H80004002
Public Const E_FAIL = &H80004005
Public Const MAX_GUIDLEN = 40

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (tGUIDStructure As udtGUID) As Long

Public Type udtGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public m_fSafeForScripting As Boolean
Public m_fSafeForInitializing As Boolean

Public Function CreateGUID() As String
  Dim sGUID   As String
  Dim tGUID   As udtGUID
  Dim bGuid() As Byte
  Dim lRtn    As Long
  Const clLen As Long = 50
  
  If CoCreateGuid(tGUID) = 0 Then
    bGuid = String(clLen, 0)
    lRtn = StringFromGUID2(tGUID, VarPtr(bGuid(0)), clLen)
    If lRtn > 0 Then
      sGUID = Mid$(bGuid, 1, lRtn - 1)
    End If
    CreateGUID = sGUID
  End If
End Function

Private Function GetCLSIDFromString(ByVal psGuid As String) As udtGUID
    Dim sCLSID As String
    Dim lBuf As Long
    sCLSID = psGuid & vbNullChar ' create null-terminated OLESTR
    lBuf = CLSIDFromString(StrPtr(sCLSID), GetCLSIDFromString)
End Function

Public Function GuidToString(ByVal psGuid As String) As String
    
  Dim sGUID   As String
  Dim tGUID   As udtGUID
  Dim bGuid() As Byte
  Dim lRtn    As Long
  Const clLen As Long = 50
  
  tGUID = GetCLSIDFromString(psGuid)

  bGuid = String(clLen, 0)
  lRtn = StringFromGUID2(tGUID, VarPtr(bGuid(0)), clLen)
  If lRtn > 0 Then
    sGUID = Mid$(bGuid, 1, lRtn - 1)
  End If
  GuidToString = sGUID
    
End Function
