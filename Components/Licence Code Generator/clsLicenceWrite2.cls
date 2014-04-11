VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLicenceWrite2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngCustomerNo As Long
Private mlngDAT As Long
Private mlngDMIM As Long
Private mlngDMIS As Long
Private mlngSSI As Long
Private mlngModules As Long

Public Property Let CustomerNo(ByVal lngNewValue As Long)
  mlngCustomerNo = lngNewValue
End Property

Public Property Let DATUsers(ByVal lngNewValue As Long)
  mlngDAT = lngNewValue
End Property

Public Property Let DMIMUsers(ByVal lngNewValue As Long)
  mlngDMIM = lngNewValue
End Property

Public Property Let DMISUsers(ByVal lngNewValue As Long)
  mlngDMIS = lngNewValue
End Property

Public Property Let SSIUsers(ByVal lngNewValue As Long)
  mlngSSI = lngNewValue
End Property

Public Property Let Modules(ByVal lngNewValue As Long)
  mlngModules = lngNewValue
End Property

Private Function Valid() As Boolean
  Valid = (mlngCustomerNo > 0 And mlngModules > 0 And (mlngDAT > 0 Or mlngDMIM > 0))
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


Private Function ConvertNumberToString2(lngSize As Long, lngInput As Long, lngDigit As Long) As String

  Dim lngRandomDigit As Long
  Dim strAlphaString As String
  Dim lngFactor As Double
  Dim lngCount As Long

  lngRandomDigit = IIf(lngDigit = 0, Int(Rnd * 26) + 1, lngDigit)
  strAlphaString = GenerateAlphaString(lngRandomDigit)

  ConvertNumberToString2 = Mid(strAlphaString, (lngInput And 31) + 1, 1)
    
  lngFactor = 32
  For lngCount = 2 To lngSize - IIf(lngDigit = 0, 1, 0)
    ConvertNumberToString2 = _
        Mid(strAlphaString, ((lngInput \ lngFactor) And 31) + 1, 1) & _
        ConvertNumberToString2
    lngFactor = lngFactor * 32
  Next
  
  If lngDigit = 0 Then
    ConvertNumberToString2 = _
        Chr(lngRandomDigit + 64) & ConvertNumberToString2
  End If

End Function


Public Property Get LicenceKey2()
    
  Dim strCustNo As String
  Dim strDAT As String
  Dim strDMIM As String
  Dim strDMIS As String
  Dim strSSI As String
  Dim strModules As String
  Dim strVersion As String
  Dim lngCount As Long
  Dim strOutput As String
  Dim lngRandomDigit As Long
  
 
  LicenceKey2 = vbNullString
  If Valid Then
    Randomize Timer
    lngRandomDigit = Int(Rnd * 26) + 1
    
    '******************************************************
    '* WHEN THE LICENCE NUMBER CHANGES PLEASE CHANGE THIS *
    '* INDICATOR TO THE NEXT LETTER IN THE ALPHABET       *
    '* (WE CAN KEEP THE FORMAT: ?????-?????-?????-?????   *
    '******************************************************
    strVersion = "A"      'Licence Version Indicator

    strCustNo = ConvertNumberToString2(4, mlngCustomerNo, 0)
    strDAT = ConvertNumberToString2(2, mlngDAT, lngRandomDigit)
    strDMIM = ConvertNumberToString2(2, mlngDMIM, lngRandomDigit)
    strDMIS = ConvertNumberToString2(2, mlngDMIS, lngRandomDigit)
    strSSI = ConvertNumberToString2(2, mlngSSI, lngRandomDigit)
    strModules = ConvertNumberToString2(6, mlngModules, 0)

    strOutput = _
        strVersion & strCustNo & strDAT & strDMIM & _
        strDMIS & strSSI & strModules & Chr(lngRandomDigit + 64)

    'Jumble it up!
    For lngCount = 1 To 4
      LicenceKey2 = LicenceKey2 & _
        IIf(LicenceKey2 <> vbNullString, "-", "") & _
          Mid(strOutput, lngCount, 1) & _
          Mid(strOutput, lngCount + 12, 1) & _
          Mid(strOutput, lngCount + 8, 1) & _
          Mid(strOutput, lngCount + 4, 1) & _
          Mid(strOutput, lngCount + 16, 1)
    Next
    
  End If

End Property
