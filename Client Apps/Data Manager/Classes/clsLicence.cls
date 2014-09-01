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
Private mlngDMIS As Long
Private mlngSSI As Long
Private mlngModules As Long
Private mlngHeadcount As Long
Private mdExpiryDate As Date
Private mlngLicenceType As Long

Public Property Get CustomerNo()
  CustomerNo = mlngCustomerNo
End Property

Public Property Get DATUsers()
  DATUsers = mlngDAT
End Property

Public Property Get DMIMUsers()
  DMIMUsers = mlngDMIM
End Property

Public Property Get DMISUsers()
  DMISUsers = mlngDMIS
End Property

Public Property Get SSIUsers()
  SSIUsers = mlngSSI
End Property

Public Property Get Modules()
  Modules = mlngModules
End Property

Public Property Get Headcount()
  Headcount = mlngHeadcount
End Property

Public Property Get ExpiryDate()
  ExpiryDate = mdExpiryDate
End Property

Public Property Get LicenceType()
  LicenceType = mlngLicenceType
End Property

Public Property Let LicenceKey(ByVal strNewValue As String)

  Dim strCustNo As String
  Dim strDAT As String
  Dim strDMIM As String
  Dim strDMIS As String
  Dim strSSI As String
  Dim strModules As String
  Dim strRandomDigit As String
  Dim strHeadcount As String
  Dim strExpiryDate As String
  Dim sLicenceType As String

  Dim strTemp As String
  Dim strInput As String
  Dim lngCount As Long
  Dim lngDateValue As Long

  '1st Generation: AAAAAAAAAAAAAAA
  '2nd Generation: 1231-2312-3123-1233
  '3rd Generation: 1234-5123-4512-3451-2345
  '4th Generation: A5316-16426-16426-16536
  '5th Generation: 45FG63-1S7G6D-1V4R63-DDG533-PG6GX
   
  Randomize Timer
  mlngCustomerNo = 0

  If strNewValue Like "??????-??????-??????-??????-??????" Then
    
    strInput = Replace(strNewValue, "-", "")
    strInput = Mid(strInput, 1, 1) & Mid(strInput, 7, 1) & Mid(strInput, 13, 1) & Mid(strInput, 19, 1) & Mid(strInput, 25, 1) & _
               Mid(strInput, 4, 1) & Mid(strInput, 10, 1) & Mid(strInput, 16, 1) & Mid(strInput, 22, 1) & Mid(strInput, 28, 1) & _
               Mid(strInput, 3, 1) & Mid(strInput, 9, 1) & Mid(strInput, 15, 1) & Mid(strInput, 21, 1) & Mid(strInput, 27, 1) & _
               Mid(strInput, 2, 1) & Mid(strInput, 8, 1) & Mid(strInput, 14, 1) & Mid(strInput, 20, 1) & Mid(strInput, 26, 1) & _
               Mid(strInput, 6, 1) & Mid(strInput, 12, 1) & Mid(strInput, 18, 1) & Mid(strInput, 24, 1) & Mid(strInput, 30, 1) & _
               Mid(strInput, 5, 1) & Mid(strInput, 11, 1) & Mid(strInput, 17, 1) & Mid(strInput, 23, 1) & Mid(strInput, 29, 1)

    sLicenceType = Mid(strInput, 1, 1)
    strCustNo = Mid(strInput, 2, 3)
    strDAT = Mid(strInput, 5, 2)
    strDMIM = Mid(strInput, 7, 2)
    strDMIS = Mid(strInput, 9, 2)
    strSSI = Mid(strInput, 11, 4)
    strModules = Mid(strInput, 15, 6)
    strRandomDigit = Mid(strInput, 21, 1)
    strHeadcount = Mid(strInput, 22, 4)
    strExpiryDate = Mid(strInput, 26, 5)

    mlngCustomerNo = ConvertStringToNumber2(strRandomDigit & strCustNo)
    mlngDAT = ConvertStringToNumber2(strRandomDigit & strDAT)
    mlngDMIM = ConvertStringToNumber2(strRandomDigit & strDMIM)
    mlngDMIS = ConvertStringToNumber2(strRandomDigit & strDMIS)
    mlngSSI = ConvertStringToNumber2(strRandomDigit & strSSI)
    mlngModules = ConvertStringToNumber2(strModules)
    mlngHeadcount = ConvertStringToNumber2(strRandomDigit & strHeadcount)
    
    lngDateValue = ConvertStringToNumber2(strRandomDigit & strExpiryDate)
    If lngDateValue > 0 Then
      mdExpiryDate = DateFromJulian(lngDateValue)
    End If
    mlngLicenceType = ConvertStringToNumber2(strRandomDigit & sLicenceType)

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




