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

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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
