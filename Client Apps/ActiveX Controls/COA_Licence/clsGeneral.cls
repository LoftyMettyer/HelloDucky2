VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


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


Public Function ConvertStringToNumber(lngMode As Long, strInput As String) As Long

  Dim lngRandomDigit As Long
  Dim strAlphaString As String
  Dim lngOutput As Long

  On Error GoTo exitf
  
  lngRandomDigit = Asc(Mid(strInput, 2, 1)) - 64
  strAlphaString = GenerateAlphaString(lngRandomDigit)

  If lngMode = 3 Then
    lngOutput = _
        ((InStr(strAlphaString, Mid(strInput, 3, 1)) - 1) * 32768) + _
        ((InStr(strAlphaString, Mid(strInput, 4, 1)) - 1) * 1024) + _
        ((InStr(strAlphaString, Mid(strInput, 5, 1)) - 1) * 32) + _
        (InStr(strAlphaString, Mid(strInput, 6, 1)) - 1)
  Else
    lngOutput = _
        ((InStr(strAlphaString, Mid(strInput, 3, 1)) - 1) * 1024) + _
        ((InStr(strAlphaString, Mid(strInput, 4, 1)) - 1) * 32) + _
        (InStr(strAlphaString, Mid(strInput, 5, 1)) - 1)
  End If
  
  If Mid(strAlphaString, (lngOutput Mod 31) + 1, 1) = Mid(strInput, 1, 1) Then
    ConvertStringToNumber = lngOutput
  End If

exitf:

End Function


Public Function ConvertStringToNumber2(strInput As String) As Long

  Dim lngRandomDigit As Long
  Dim strAlphaString As String
  Dim lngOutput As Long
  Dim lngFactor As Double
  Dim lngCount As Long

  On Error GoTo exitf
  
  lngRandomDigit = Asc(Mid(strInput, 1, 1)) - 64
  strAlphaString = GenerateAlphaString(lngRandomDigit)

  lngOutput = (InStr(strAlphaString, Mid(strInput, Len(strInput), 1)) - 1)
  
  lngFactor = 32
  For lngCount = Len(strInput) - 1 To 2 Step -1
    lngOutput = lngOutput + _
      ((InStr(strAlphaString, Mid(strInput, lngCount, 1)) - 1) * lngFactor)
    lngFactor = lngFactor * 32
  Next

  ConvertStringToNumber2 = lngOutput

exitf:

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


Public Function CreateKey(lngC As Long, lngN As Long, lngM As Long) As String

  Dim strArray(2) As String
  Dim lngCount As Long

  strArray(0) = ConvertNumberToString(1, lngC)
  strArray(1) = ConvertNumberToString(2, lngN)
  strArray(2) = ConvertNumberToString(3, lngM)


  CreateKey = vbNullString
  For lngCount = 1 To 5
    CreateKey = CreateKey & _
      Mid(strArray(0), lngCount, 1) & _
      Mid(strArray(1), lngCount, 1) & _
      Mid(strArray(2), lngCount, 1)
  Next
  CreateKey = CreateKey & Mid(strArray(2), 6, 1)

  CreateKey = _
      Mid(CreateKey, 1, 4) & "-" & Mid(CreateKey, 5, 4) & "-" & _
      Mid(CreateKey, 9, 4) & "-" & Mid(CreateKey, 13, 4)

End Function


Public Function CreateKey2(lngC As Long, lngN As Long, lngI As Long, lngS As Long, lngM As Long) As String

  Dim strArray(4) As String
  Dim lngCount As Long
  Dim strOutput As String

  strArray(0) = ConvertNumberToString2(4, lngC, 0)
  strArray(1) = ConvertNumberToString2(4, lngN, 0)
  strArray(2) = ConvertNumberToString2(4, lngI, 0)
  strArray(3) = ConvertNumberToString2(4, lngS, 0)
  strArray(4) = ConvertNumberToString2(4, lngM, 0)

  strOutput = vbNullString
  For lngCount = 1 To 4
    strOutput = strOutput & _
      Mid(strArray(0), lngCount, 1) & _
      Mid(strArray(1), lngCount, 1) & _
      Mid(strArray(2), lngCount, 1) & _
      Mid(strArray(3), lngCount, 1) & _
      Mid(strArray(4), lngCount, 1)
  Next
  'CreateKey = CreateKey & Mid(strArray(2), 6, 1)

  'CreateKey = _
      Mid(CreateKey, 1, 4) & "-" & Mid(CreateKey, 5, 4) & "-" & _
      Mid(CreateKey, 9, 4) & "-" & Mid(CreateKey, 13, 4)

  For lngCount = 1 To Len(strOutput) Step 4
    CreateKey2 = CreateKey2 & _
      IIf(CreateKey2 <> vbNullString, "-", "") & _
      Mid(strOutput, lngCount, 4)
  Next

End Function


Public Function CreateKey3(mlngCustomerNo As Long, lngDAT As Long, lngDMIM As Long, lngDMIS As Long, lngSSI As Long, mlngModules As Long) As String

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
  
 
  CreateKey3 = vbNullString
  'If Valid Then
    Randomize Timer
    lngRandomDigit = Int(Rnd * 26) + 1
    
    '******************************************************
    '* WHEN THE LICENCE NUMBER CHANGES PLEASE CHANGE THIS *
    '* INDICATOR TO THE NEXT LETTER IN THE ALPHABET       *
    '* (WE CAN KEEP THE FORMAT: ?????-?????-?????-?????   *
    '******************************************************
    strVersion = "A"      'Licence Version Indicator

    strCustNo = ConvertNumberToString2(4, mlngCustomerNo, 0)
    strDAT = ConvertNumberToString2(2, lngDAT, lngRandomDigit)
    strDMIM = ConvertNumberToString2(2, lngDMIM, lngRandomDigit)
    strDMIS = ConvertNumberToString2(2, lngDMIS, lngRandomDigit)
    strSSI = ConvertNumberToString2(2, lngSSI, lngRandomDigit)
    strModules = ConvertNumberToString2(6, mlngModules, 0)

    strOutput = _
        strVersion & strCustNo & strDAT & strDMIM & _
        strDMIS & strSSI & strModules & Chr(lngRandomDigit + 64)

    'Jumble it up!
    For lngCount = 1 To 4
      CreateKey3 = CreateKey3 & _
        IIf(CreateKey3 <> vbNullString, "-", "") & _
          Mid(strOutput, lngCount, 1) & _
          Mid(strOutput, lngCount + 12, 1) & _
          Mid(strOutput, lngCount + 8, 1) & _
          Mid(strOutput, lngCount + 4, 1) & _
          Mid(strOutput, lngCount + 16, 1)
    Next
    
  'End If
  
End Function
