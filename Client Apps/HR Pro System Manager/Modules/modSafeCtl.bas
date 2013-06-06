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

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As udtGUID, pSource As Long, ByVal ByteLen As Long)
Private Declare Function CoCreateGuid Lib "ole32.dll" (tGUIDStructure As udtGUID) As Long
Public Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As udtGUID, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Type udtGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public m_fSafeForScripting As Boolean
Public m_fSafeForInitializing As Boolean

'use Sub Main from modSysMgr.bas
'Sub Main()
'    m_fSafeForScripting = True
'    m_fSafeForInitializing = True
'End Sub

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
