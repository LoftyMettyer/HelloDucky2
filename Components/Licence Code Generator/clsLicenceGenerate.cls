VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLicenceWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngCustomerNo As Long
Private mlngDatUsers As Long
Private mlngIntUsers As Long
Private mlngSSUsers As Long
Private mlngModules As Long

Public Property Let CustomerNo(ByVal lngNewValue As Long)
  mlngCustomerNo = lngNewValue
End Property

Public Property Let NumUsers(ByVal lngNewValue As Long)
  mlngDatUsers = lngNewValue
End Property

Public Property Let IntUsers(ByVal lngNewValue As Long)
  mlngIntUsers = lngNewValue
End Property

Public Property Let SSUsers(ByVal lngNewValue As Long)
  mlngSSUsers = lngNewValue
End Property

Public Property Let Modules(ByVal lngNewValue As Long)
  mlngModules = lngNewValue
End Property

Private Function Valid() As Boolean
  Valid = (mlngCustomerNo > 0 And mlngModules > 0 And (mlngDatUsers > 0 Or mlngIntUsers > 0))
End Function


Private Function GenerateAlphaString(lngGap As Long) As String

  Dim strOutput As String
  Dim lngCount As Long
  Dim lngLoop As Long

  strOutput = vbNullString
  For lngLoop = 0 To lngGap - 1

    For lngCount = Asc("A") + lngLoop To Asc("Z") Step lngGap
      If InStr("IOQ", Chr(lngCount)) = 0 Then
        strOutput = Chr(lngCount) & strOutput
      End If
    Next

    For lngCount = Asc("1") + lngLoop To Asc("9") Step lngGap
      strOutput = Chr(lngCount) & strOutput
    Next

  Next

  GenerateAlphaString = strOutput

End Function


Private Function ConvertNumberToString(lngMode As Long, lngInput As Long) As String

  Dim lngRandomDigit As Long
  Dim strAlphaString As String

  lngRandomDigit = Int(Rnd * 26) + 1
  strAlphaString = GenerateAlphaString(lngRandomDigit)

  ConvertNumberToString = _
      Mid(strAlphaString, (lngInput Mod 31) + 1, 1) & _
      Chr(lngRandomDigit + 64)

  If lngMode = 3 Then
    ConvertNumberToString = ConvertNumberToString + _
      Mid(strAlphaString, ((lngInput \ 32768) And 31) + 1, 1)
  End If

  ConvertNumberToString = ConvertNumberToString + _
      Mid(strAlphaString, ((lngInput \ 1024) And 31) + 1, 1) & _
      Mid(strAlphaString, ((lngInput \ 32) And 31) + 1, 1) & _
      Mid(strAlphaString, (lngInput And 31) + 1, 1)


End Function


Private Function ConvertNumberToString2(lngSize As Long, lngInput As Long) As String

  Dim lngRandomDigit As Long
  Dim strAlphaString As String
  Dim lngFactor As Double
  Dim lngCount As Long

  lngRandomDigit = Int(Rnd * 26) + 1
  strAlphaString = GenerateAlphaString(lngRandomDigit)

  ConvertNumberToString2 = Mid(strAlphaString, (lngInput And 31) + 1, 1)
    
  lngFactor = 32
  For lngCount = lngSize - 1 To 2 Step -1
    ConvertNumberToString2 = _
        Mid(strAlphaString, ((lngInput \ lngFactor) And 31) + 1, 1) & _
        ConvertNumberToString2
    lngFactor = lngFactor * 32
  Next
  
  ConvertNumberToString2 = _
      Chr(lngRandomDigit + 64) & ConvertNumberToString2

End Function


Public Property Get LicenceKey()
    
  Dim strCustNo As String
  Dim strUsers As String
  Dim strModules As String
  Dim lngCount As Long
  
  If Valid Then
    Randomize Timer
    
    strCustNo = ConvertNumberToString(1, mlngCustomerNo)
    strUsers = ConvertNumberToString(2, mlngDatUsers)
    strModules = ConvertNumberToString(3, mlngModules)
  
    LicenceKey = ""
    For lngCount = 1 To 5
      LicenceKey = LicenceKey & _
        Mid(strCustNo, lngCount, 1) & _
        Mid(strUsers, lngCount, 1) & _
        Mid(strModules, lngCount, 1)
    Next

    LicenceKey = _
        Mid(LicenceKey, 1, 4) & "-" & _
        Mid(LicenceKey, 5, 4) & "-" & _
        Mid(LicenceKey, 9, 4) & "-" & _
        Mid(LicenceKey, 13, 3) & Mid(strModules, lngCount, 6)

  End If
End Property


Public Property Get LicenceKey2()
    
  Dim strCustNo As String
  Dim strDatUsers As String
  Dim strIntUsers As String
  Dim strSSUsers As String
  Dim strModules As String
  Dim lngCount As Long
  Dim strOutput As String

  LicenceKey2 = vbNullString
  If Valid Then
    Randomize Timer

    strCustNo = ConvertNumberToString2(4, mlngCustomerNo)
    strDatUsers = ConvertNumberToString2(4, mlngDatUsers)
    strIntUsers = ConvertNumberToString2(4, mlngIntUsers)
    strSSUsers = ConvertNumberToString2(4, mlngSSUsers)
    strModules = ConvertNumberToString2(4, mlngModules)

    strOutput = ""
    For lngCount = 1 To 4
      strOutput = strOutput & _
        Mid(strCustNo, lngCount, 1) & _
        Mid(strDatUsers, lngCount, 1) & _
        Mid(strIntUsers, lngCount, 1) & _
        Mid(strSSUsers, lngCount, 1) & _
        Mid(strModules, lngCount, 1)
    Next

    For lngCount = 1 To Len(strOutput) Step 4
      LicenceKey2 = LicenceKey2 & _
        IIf(LicenceKey2 <> vbNullString, "-", "") & _
        Mid(strOutput, lngCount, 4)
    Next

  End If

End Property


