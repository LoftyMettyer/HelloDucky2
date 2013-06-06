VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLicence2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngCustomerNo As Long
Private mlngDAT As Long
Private mlngDMIM As Long
Private mlngDMIS As Long
Private mlngSSI As Long
Private mlngModules As Long

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


Public Property Let LicenceKey(ByVal strNewValue As String)

  Dim datGeneral As clsGeneral
  Dim strCustNo As String
  Dim strDAT As String
  Dim strDMIM As String
  Dim strDMIS As String
  Dim strSSI As String
  Dim strModules As String
  Dim strRandomDigit As String

  Dim strTemp As String
  Dim strInput As String
  Dim lngCount As Long

  '1st Generation: AAAAAAAAAAAAAAA
  '2nd Generation: 1231-2312-3123-1233
  '3rd Generation: 1234-5123-4512-3451-2345
  '4th Generation: A5316-16426-16426-16536
  
  
  Randomize Timer
  mlngCustomerNo = 0

  'Check format and licence version indicator...
  'DO NOT REMOVE THE VERSION INDICATOR (i.e. the first character)
  If strNewValue Like "A????-?????-?????-?????" Then
    
    strInput = Replace(strNewValue, "-", "")
    strInput = Mid(strInput, 1, 1) & Mid(strInput, 6, 1) & _
               Mid(strInput, 11, 1) & Mid(strInput, 16, 1) & _
               Mid(strInput, 4, 1) & Mid(strInput, 9, 1) & _
               Mid(strInput, 14, 1) & Mid(strInput, 19, 1) & _
               Mid(strInput, 3, 1) & Mid(strInput, 8, 1) & _
               Mid(strInput, 13, 1) & Mid(strInput, 18, 1) & _
               Mid(strInput, 2, 1) & Mid(strInput, 7, 1) & _
               Mid(strInput, 12, 1) & Mid(strInput, 17, 1) & _
               Mid(strInput, 5, 1) & Mid(strInput, 10, 1) & _
               Mid(strInput, 15, 1) & Mid(strInput, 20, 1)

    strCustNo = Mid(strInput, 2, 4)
    strDAT = Mid(strInput, 6, 2)
    strDMIM = Mid(strInput, 8, 2)
    strDMIS = Mid(strInput, 10, 2)
    strSSI = Mid(strInput, 12, 2)
    strModules = Mid(strInput, 14, 6)
    strRandomDigit = Mid(strInput, 20, 1)

    Set datGeneral = New clsGeneral
    mlngCustomerNo = datGeneral.ConvertStringToNumber2(strCustNo)
    mlngDAT = datGeneral.ConvertStringToNumber2(strRandomDigit & strDAT)
    mlngDMIM = datGeneral.ConvertStringToNumber2(strRandomDigit & strDMIM)
    mlngDMIS = datGeneral.ConvertStringToNumber2(strRandomDigit & strDMIS)
    mlngSSI = datGeneral.ConvertStringToNumber2(strRandomDigit & strSSI)
    mlngModules = datGeneral.ConvertStringToNumber2(strModules)
    Set datGeneral = Nothing

  End If

  
  If mlngCustomerNo = 0 Or mlngDAT = 0 Or mlngModules = 0 Then
    mlngCustomerNo = 0
    mlngDAT = 0
    mlngDMIM = 0
    mlngDMIS = 0
    mlngSSI = 0
    mlngModules = 0
  End If

End Property
