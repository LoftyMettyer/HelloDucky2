VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngCustomerNo As Long
Private mlngDAT As Long
Private mlngDMIM As Long
Private mlngModules As Long
Private mlngHeadcountSSI As Long
Private mdExpiryDate As Date
Private mlngDateValue As Long
Private mlngLicenceType As Long

Public IsValid As Boolean
Public ValidateCreationDate As Boolean

Public Property Get CustomerNo()
  CustomerNo = mlngCustomerNo
End Property

Public Property Get DATUsers()
  DATUsers = mlngDAT
End Property

Public Property Get DMIMUsers()
  DMIMUsers = mlngDMIM
End Property

Public Property Get SSIUsers()
  SSIUsers = IIf(mlngLicenceType = 0, mlngHeadcountSSI, 0)
End Property

Public Property Get Modules()
  Modules = mlngModules
End Property

Public Property Get Headcount()
  Headcount = IIf(mlngLicenceType > 0, mlngHeadcountSSI, 0)
End Property

Public Property Get ExpiryDate() As Date
  If mlngDateValue > 0 Then
    ExpiryDate = CDate(mlngDateValue)
  End If
End Property

Public Property Get HasExpiryDate() As Boolean
  HasExpiryDate = mlngDateValue > 0
End Property

Public Property Get LicenceType()
  LicenceType = mlngLicenceType
End Property

Public Property Let LicenceKey(ByVal strNewValue As String)

  Dim strCustNo As String
  Dim strDAT As String
  Dim strDMIM As String
  Dim strDMIS As String
  Dim strModules As String
  Dim strRandomDigit As String
  Dim strSSIHeadcount As String
  Dim strExpiryDate As String
  Dim sLicenceType As String

  Dim strTemp As String
  Dim strInput As String
  Dim lngCount As Long
  Dim strCheckSum As String
  Dim lngRandomDigit As Long
  Dim strGeneratedDay As String
  Dim lngKeyGenDayNumber As Long

  '1st Generation: AAAAAAAAAAAAAAA
  '2nd Generation: 1231-2312-3123-1233
  '3rd Generation: 1234-5123-4512-3451-2345
  '4th Generation: A5316-16426-16426-16536
  '5th Generation: 45FG63-1S7G6D-1V4R63-DDG533-PG6GX
   
  Randomize Timer
  mlngCustomerNo = 0

  If strNewValue Like "??????-??????-??????-??????-??????-??????" Then
    
    strInput = Replace(strNewValue, "-", "")
    strInput = Mid(strInput, 32, 1) & Mid(strInput, 35, 1) & Mid(strInput, 6, 1) & Mid(strInput, 27, 1) & Mid(strInput, 21, 1) & Mid(strInput, 29, 1) & _
           Mid(strInput, 1, 1) & Mid(strInput, 26, 1) & Mid(strInput, 4, 1) & Mid(strInput, 28, 1) & Mid(strInput, 2, 1) & Mid(strInput, 8, 1) & _
           Mid(strInput, 3, 1) & Mid(strInput, 13, 1) & Mid(strInput, 31, 1) & Mid(strInput, 18, 1) & Mid(strInput, 33, 1) & Mid(strInput, 10, 1) & _
           Mid(strInput, 20, 1) & Mid(strInput, 14, 1) & Mid(strInput, 23, 1) & Mid(strInput, 9, 1) & Mid(strInput, 25, 1) & Mid(strInput, 16, 1) & _
           Mid(strInput, 36, 1) & Mid(strInput, 7, 1) & Mid(strInput, 22, 1) & Mid(strInput, 17, 1) & Mid(strInput, 34, 1) & Mid(strInput, 19, 1) & _
           Mid(strInput, 30, 1) & Mid(strInput, 24, 1) & Mid(strInput, 15, 1) & Mid(strInput, 5, 1) & Mid(strInput, 12, 1) & Mid(strInput, 11, 1)

    sLicenceType = Mid(strInput, 1, 1)
    strCustNo = Mid(strInput, 2, 3)
    strDAT = Mid(strInput, 5, 2)
    strDMIM = Mid(strInput, 7, 2)
    strCheckSum = Mid(strInput, 9, 6)
    strModules = Mid(strInput, 17, 7)
    strRandomDigit = Mid(strInput, 24, 1)
    strSSIHeadcount = Mid(strInput, 25, 4)
    strExpiryDate = Mid(strInput, 29, 5)
    strGeneratedDay = Mid(strInput, 34, 2)

    mlngLicenceType = ConvertStringToNumber2(strRandomDigit & sLicenceType)
    mlngCustomerNo = ConvertStringToNumber2(strRandomDigit & strCustNo)
    mlngDAT = ConvertStringToNumber2(strRandomDigit & strDAT)
    mlngDMIM = ConvertStringToNumber2(strRandomDigit & strDMIM)
    mlngHeadcountSSI = ConvertStringToNumber2(strRandomDigit & strSSIHeadcount)
    mlngModules = ConvertStringToNumber2(strModules)
    lngRandomDigit = Asc(strRandomDigit) - 64
    mlngDateValue = ConvertStringToNumber2(strRandomDigit & strExpiryDate)
        
    ' Validate checksum
    If ValidateCreationDate Then
      lngKeyGenDayNumber = ConvertStringToNumber2(strRandomDigit & strGeneratedDay)
      IsValid = lngKeyGenDayNumber >= DayNumber(Now) - 7 And lngKeyGenDayNumber <= DayNumber(Now)
    Else
      IsValid = True
    End If
    
    IsValid = IsValid And (ConvertNumberToString2(6, mlngCustomerNo + mlngLicenceType + mlngDAT + mlngDMIM + mlngHeadcountSSI + mlngModules + mlngDateValue, lngRandomDigit) = strCheckSum)

  End If

End Property

Private Function DateFromJulian(ByVal sDate As String) As Date

  Dim dResult As Date
  Dim sYear As String
  sYear = Mid(sDate, 1, 4)

  dResult = DateSerial(sYear, 1, 1)
  DateFromJulian = dResult + Mid(sDate, 5, 3) - 1
    
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

Private Function DayNumber(ByVal vDate As Date) As Long
    DayNumber = CLng(Format(DateDiff("d", CDate("01/01/" _
                  + Format(Year(vDate), "0000")), vDate) _
                  + 1, "000"))
End Function

