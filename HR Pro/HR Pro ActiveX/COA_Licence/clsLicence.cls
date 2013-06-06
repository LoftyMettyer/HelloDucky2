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
Attribute VB_Exposed = True
Option Explicit

Private mlngCustomerNo As Long
Private mlngDatUsers As Long
Private mlngIntUsers As Long
Private mlngSSUsers As Long
Private mlngModules As Long


Public Property Get CustomerNo()
  CustomerNo = mlngCustomerNo
End Property

Public Property Get NumUsers()
  NumUsers = mlngDatUsers
End Property

Public Property Get IntUsers()
  IntUsers = mlngIntUsers
End Property

Public Property Get SSUsers()
  SSUsers = mlngSSUsers
End Property

Public Property Get Modules()
  Modules = mlngModules
End Property


Public Property Let LicenceKey(ByVal strNewValue As String)

  Dim datGeneral As clsGeneral
  Dim strCustNo As String
  Dim strUsers As String
  Dim strModules As String
  
  Dim strTemp As String
  Dim lngCount As Long

  '1231-2312-3123-1233

  Randomize Timer
  mlngCustomerNo = 0

  If strNewValue Like "????-????-????-????" Then
    strTemp = Replace(strNewValue, "-", "")

    strCustNo = ""
    strUsers = ""
    strModules = ""

    For lngCount = 1 To 15 Step 3
      strCustNo = strCustNo & Mid(strTemp, lngCount, 1)
      strUsers = strUsers & Mid(strTemp, lngCount + 1, 1)
      strModules = strModules & Mid(strTemp, lngCount + 2, 1)
    Next
    strModules = strModules & Mid(strTemp, 16, 1)

    Set datGeneral = New clsGeneral
    mlngCustomerNo = datGeneral.ConvertStringToNumber(1, strCustNo)
    mlngDatUsers = datGeneral.ConvertStringToNumber(2, strUsers)
    mlngModules = datGeneral.ConvertStringToNumber(3, strModules)
    Set datGeneral = Nothing

  End If

  
  If mlngCustomerNo = 0 Or mlngDatUsers = 0 Or mlngModules = 0 Then
    mlngCustomerNo = 0
    mlngDatUsers = 0
    mlngModules = 0
  End If

End Property


Public Property Let LicenceKey2(ByVal strNewValue As String)

  Dim datGeneral As clsGeneral
  Dim strCustNo As String
  Dim strUsers As String
  Dim strIntUsers As String
  Dim strSSUsers As String
  Dim strModules As String
  
  Dim strTemp As String
  Dim lngCount As Long

  '1231-2312-3123-1233
  '1234-5123-4512-3451-2345

  Randomize Timer
  mlngCustomerNo = 0

  If strNewValue Like "????-????-????-????-????" Then
    strTemp = Replace(strNewValue, "-", "")

    strCustNo = ""
    strUsers = ""
    strIntUsers = ""
    strSSUsers = ""
    strModules = ""

    For lngCount = 1 To Len(strTemp) Step 5
      strCustNo = strCustNo & Mid(strTemp, lngCount, 1)
      strUsers = strUsers & Mid(strTemp, lngCount + 1, 1)
      strIntUsers = strIntUsers & Mid(strTemp, lngCount + 2, 1)
      strSSUsers = strSSUsers & Mid(strTemp, lngCount + 3, 1)
      strModules = strModules & Mid(strTemp, lngCount + 4, 1)
    Next

    Set datGeneral = New clsGeneral
    mlngCustomerNo = datGeneral.ConvertStringToNumber2(strCustNo)
    mlngDatUsers = datGeneral.ConvertStringToNumber2(strUsers)
    mlngIntUsers = datGeneral.ConvertStringToNumber2(strIntUsers)
    mlngSSUsers = datGeneral.ConvertStringToNumber2(strSSUsers)
    mlngModules = datGeneral.ConvertStringToNumber2(strModules)
    Set datGeneral = Nothing

  End If

  
  If mlngCustomerNo = 0 Or mlngDatUsers = 0 Or mlngModules = 0 Then
    mlngCustomerNo = 0
    mlngDatUsers = 0
    mlngModules = 0
  End If

End Property

