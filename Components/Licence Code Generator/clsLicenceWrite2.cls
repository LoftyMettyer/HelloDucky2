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
Private mlngSSI As Long
Private mlngModules As Long

Private mlngHeadcount As Long
Private mlngExpiryDate As Long
Private mlngLicenceType As Long

Public Property Let CustomerNo(ByVal lngNewValue As Long)
  mlngCustomerNo = lngNewValue
End Property

Public Property Let DATUsers(ByVal lngNewValue As Long)
  mlngDAT = lngNewValue
End Property

Public Property Let DMIMUsers(ByVal lngNewValue As Long)
  mlngDMIM = lngNewValue
End Property

Public Property Let SSIUsers(ByVal lngNewValue As Long)
  mlngSSI = lngNewValue
End Property

Public Property Let Modules(ByVal lngNewValue As Long)
  mlngModules = lngNewValue
End Property

Public Property Let Headcount(ByVal lngNewValue As Long)
  mlngHeadcount = lngNewValue
End Property

Public Property Let ExpiryDate(ByVal lngNewValue As Long)
  mlngExpiryDate = lngNewValue
End Property

Public Property Let LicenceType(ByVal lngNewValue As Long)
  mlngLicenceType = lngNewValue
End Property

Private Function Valid() As Boolean
  Valid = (mlngCustomerNo > 0 And mlngModules > 0)
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
  Dim strFiller1 As String
  Dim strModules As String
  Dim strSSIHeadcount As String
  Dim strExpiryDate As String
  Dim lngCount As Long
  Dim strOutput As String
  Dim lngRandomDigit As Long
  Dim sLicenceType As String
  Dim strCheckSum As String
  Dim strGeneratedDay As String
  Dim lngSSIHeadcount As Long
  
  LicenceKey2 = vbNullString
  If Valid Then
    Randomize Timer
    lngRandomDigit = Int(Rnd * 26) + 1
    
    lngSSIHeadcount = IIf(mlngLicenceType = 0, mlngSSI, mlngHeadcount)
    
    sLicenceType = ConvertNumberToString2(1, mlngLicenceType, lngRandomDigit)
    strCustNo = ConvertNumberToString2(3, mlngCustomerNo, lngRandomDigit)
    strDAT = ConvertNumberToString2(2, mlngDAT, lngRandomDigit)
    strDMIM = ConvertNumberToString2(2, mlngDMIM, lngRandomDigit)
    strFiller1 = ConvertNumberToString2(2, 1024, lngRandomDigit)
    strSSIHeadcount = ConvertNumberToString2(4, lngSSIHeadcount, lngRandomDigit)
    strModules = ConvertNumberToString2(7, mlngModules, 0)            ' Allows upto 35 modules
    strExpiryDate = ConvertNumberToString2(5, mlngExpiryDate, lngRandomDigit)
    
    strGeneratedDay = ConvertNumberToString2(2, DayNumber(Now), lngRandomDigit)
    strCheckSum = ConvertNumberToString2(6, mlngCustomerNo + mlngLicenceType + mlngDAT + mlngDMIM + lngSSIHeadcount + mlngModules + mlngExpiryDate, lngRandomDigit)

    strOutput = _
        sLicenceType & strCustNo & strDAT & strDMIM & strCheckSum & _
        strFiller1 & strModules & Chr(lngRandomDigit + 64) & strSSIHeadcount & strExpiryDate + strGeneratedDay + "X"

    'Jumble it up!
    LicenceKey2 = Mid(strOutput, 7, 1) & Mid(strOutput, 11, 1) & Mid(strOutput, 13, 1) & Mid(strOutput, 9, 1) & Mid(strOutput, 34, 1) & Mid(strOutput, 3, 1) & "-" & _
           Mid(strOutput, 26, 1) & Mid(strOutput, 12, 1) & Mid(strOutput, 22, 1) & Mid(strOutput, 18, 1) & Mid(strOutput, 36, 1) & Mid(strOutput, 35, 1) & "-" & _
           Mid(strOutput, 14, 1) & Mid(strOutput, 20, 1) & Mid(strOutput, 33, 1) & Mid(strOutput, 24, 1) & Mid(strOutput, 28, 1) & Mid(strOutput, 16, 1) & "-" & _
           Mid(strOutput, 30, 1) & Mid(strOutput, 19, 1) & Mid(strOutput, 5, 1) & Mid(strOutput, 27, 1) & Mid(strOutput, 21, 1) & Mid(strOutput, 32, 1) & "-" & _
           Mid(strOutput, 23, 1) & Mid(strOutput, 8, 1) & Mid(strOutput, 4, 1) & Mid(strOutput, 10, 1) & Mid(strOutput, 6, 1) & Mid(strOutput, 31, 1) & "-" & _
           Mid(strOutput, 15, 1) & Mid(strOutput, 1, 1) & Mid(strOutput, 17, 1) & Mid(strOutput, 29, 1) & Mid(strOutput, 2, 1) & Mid(strOutput, 25, 1)

  End If

End Property

Private Function PadString(ByVal length As Integer, ByVal value As String, ByVal padWith As String)
  PadString = Right(String(length, padWith) & value, length)
End Function

Private Function Date2Julian(ByVal vDate As Date) As Long

    Date2Julian = CLng(Format(Year(vDate), "0000") _
                  + Format(DateDiff("d", CDate("01/01/" _
                  + Format(Year(vDate), "0000")), vDate) _
                  + 1, "000"))
    
End Function

Private Function DayNumber(ByVal vDate As Date) As Long

    DayNumber = CLng(Format(DateDiff("d", CDate("01/01/" _
                  + Format(Year(vDate), "0000")), vDate) _
                  + 1, "000"))
    
End Function


