VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mstrSupportString As String
Dim lngArray(4) As Long


Public Function GetSupportString() As String

  Dim datGeneral As clsGeneral
  Dim lngCount As Long

  Randomize Timer

  lngArray(0) = Int(Rnd * 999) + 1    'Only go up to 999 as cust nums start at 1000!
  lngArray(1) = Int(Rnd * 32678) + 1
  lngArray(2) = Int(Rnd * 32678) + 1

  Set datGeneral = New clsGeneral
  mstrSupportString = datGeneral.CreateKey(lngArray(0), lngArray(1), lngArray(2))
  Set datGeneral = Nothing
  
  GetSupportString = mstrSupportString

End Function


Public Function CheckSupportInputString(strInput As String) As Boolean

  Dim A As clsLicence
  Dim B As clsLicence
  Dim lngCount As Long

  CheckSupportInputString = False

  'Check Input is in the right format
  If Not (strInput Like "????-????-????-????") Then
    Exit Function
  End If


  'Now validate that the key breakdown matches!
  If strInput <> mstrSupportString Then

    Set A = New clsLicence
    Set B = New clsLicence

    B.LicenceKey = strInput
    
    If B.CustomerNo > 0 Then
      A.LicenceKey = mstrSupportString
      CheckSupportInputString = _
          (A.CustomerNo = B.CustomerNo And A.NumUsers = B.NumUsers And A.Modules = B.Modules)
    End If

    Set A = Nothing
    Set B = Nothing
  
  End If

End Function


Public Function GetSupportString2() As String

  Dim datGeneral As clsGeneral
  Dim lngCount As Long

  Randomize Timer

  lngArray(0) = Int(Rnd * 999) + 1    'Only go up to 999 as cust nums start at 1000!
  lngArray(1) = Int(Rnd * 32678) + 1
  lngArray(2) = Int(Rnd * 32678) + 1
  lngArray(3) = Int(Rnd * 32678) + 1
  lngArray(4) = Int(Rnd * 32678) + 1

  Set datGeneral = New clsGeneral
  mstrSupportString = datGeneral.CreateKey2(lngArray(0), lngArray(1), lngArray(2), lngArray(3), lngArray(4))
  Set datGeneral = Nothing
  
  GetSupportString2 = mstrSupportString

End Function

Public Function CheckSupportInputString2(strInput As String) As Boolean

  Dim A As clsLicence
  Dim B As clsLicence
  Dim lngCount As Long

  CheckSupportInputString2 = False

  'Check Input is in the right format
  If Not (strInput Like "????-????-????-????-????") Then
    Exit Function
  End If


  'Now validate that the key breakdown matches!
  If strInput <> mstrSupportString Then

    Set A = New clsLicence
    Set B = New clsLicence

    B.LicenceKey2 = strInput
    
    If B.CustomerNo > 0 Then
      A.LicenceKey2 = mstrSupportString
      CheckSupportInputString2 = _
          (A.CustomerNo = B.CustomerNo And _
           A.NumUsers = B.NumUsers And _
           A.IntUsers = B.IntUsers And _
           A.SSUsers = B.SSUsers And _
           A.Modules = B.Modules)
    End If

    Set A = Nothing
    Set B = Nothing
  
  End If

End Function

