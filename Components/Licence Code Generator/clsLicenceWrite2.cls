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

Private mlngHeadcount As Long
Private mdExpiryDate As Date
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

Public Property Let DMISUsers(ByVal lngNewValue As Long)
  mlngDMIS = lngNewValue
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

Public Property Let ExpiryDate(ByVal dNewValue As Date)
  mdExpiryDate = dNewValue
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
  Dim strDMIS As String
  Dim strSSI As String
  Dim strModules As String
  Dim strVersion As String
  Dim strHeadcount As String
  Dim strExpiryDate As String
  Dim lngCount As Long
  Dim strOutput As String
  Dim lngRandomDigit As Long
  Dim sLicenceType As String
  
  LicenceKey2 = vbNullString
  If Valid Then
    Randomize Timer
    lngRandomDigit = Int(Rnd * 26) + 1
    
    sLicenceType = ConvertNumberToString2(1, mlngLicenceType, lngRandomDigit)
    strCustNo = ConvertNumberToString2(3, mlngCustomerNo, lngRandomDigit)
    strDAT = ConvertNumberToString2(2, mlngDAT, lngRandomDigit)
    strDMIM = ConvertNumberToString2(2, mlngDMIM, lngRandomDigit)
    strDMIS = ConvertNumberToString2(2, mlngDMIS, lngRandomDigit)
    strSSI = ConvertNumberToString2(4, mlngSSI, lngRandomDigit)
    strModules = ConvertNumberToString2(6, mlngModules, 0)            ' Allows upto 24 modules
    strHeadcount = ConvertNumberToString2(4, mlngHeadcount, lngRandomDigit)
    strExpiryDate = ConvertNumberToString2(5, Date2Julian(mdExpiryDate), lngRandomDigit)

    strOutput = _
        strVersion & sLicenceType & strCustNo & strDAT & strDMIM & _
        strDMIS & strSSI & strModules & Chr(lngRandomDigit + 64) & strHeadcount & strExpiryDate

    'Jumble it up!
    For lngCount = 1 To 5
      LicenceKey2 = LicenceKey2 & _
        IIf(LicenceKey2 <> vbNullString, "-", "") & _
          Mid(strOutput, lngCount, 1) & _
          Mid(strOutput, lngCount + 15, 1) & _
          Mid(strOutput, lngCount + 10, 1) & _
          Mid(strOutput, lngCount + 5, 1) & _
          Mid(strOutput, lngCount + 25, 1) & _
          Mid(strOutput, lngCount + 20, 1)
    Next
    
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

