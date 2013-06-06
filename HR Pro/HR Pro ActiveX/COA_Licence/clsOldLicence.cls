VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOldLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum Module
    Personnel = 1
    Recruitment = 2
    Absence = 4
    Training = 8
    Skills = 16
    Web = 32
    Afd = 64
End Enum

Private Function GetUsers(lCustNo As Long, sAuthCode As String) As Long

  Dim sTemp As String
  Dim sUsers As String
  Dim lTemp As Long
  
  On Error GoTo Err_Trap
  
  'Check that the authorisation code matches the customer number
  If ValidateUsers(sAuthCode, lCustNo) Then
    
    'Get the position of the dummy character in the second section
    lTemp = Mid$(sAuthCode, 7, 1)
    
    'Get the second section of the authcode
    sTemp = Mid$(sAuthCode, 8, 4)
    
    'Extract the dummy character from the second section
    If lTemp = 1 Then
      sUsers = Mid$(sTemp, 2, 3)
    ElseIf lTemp = 4 Then
      sUsers = Mid$(sTemp, 1, 3)
    Else
      sUsers = Mid$(sTemp, 1, lTemp - 1) & Mid$(sTemp, lTemp + 1, Len(sTemp))
    End If
    
    'Decode and check it matches the no. users
    sTemp = ""
    For lTemp = 1 To 3
      sTemp = sTemp & AlphaCode(Mid$(sUsers, lTemp, 1))
    Next
    GetUsers = Val(sTemp)
  Else
    GetUsers = 0
  End If
  
  Exit Function
  
Err_Trap:
  GetUsers = 0

End Function

Private Function AlphaCode(sChar As String) As Long

  'The Alphacode key
  
  Select Case UCase(sChar)
    Case "A"
      AlphaCode = 1
    Case "B"
      AlphaCode = 2
    Case "C"
      AlphaCode = 3
    Case "D"
      AlphaCode = 4
    Case "E"
      AlphaCode = 5
    Case "F"
      AlphaCode = 6
    Case "G"
      AlphaCode = 7
    Case "H"
      AlphaCode = 8
    Case "I"
      AlphaCode = 9
    Case "J"
      AlphaCode = 0
    Case "P"
      AlphaCode = 16
    Case "Z"
      AlphaCode = 32
    Case "Y"
      AlphaCode = 64
  End Select
        
End Function

Private Function ValidateUsers(sAuthCode As String, lCustNo As Long) As Boolean

  Dim lTemp As Long
  Dim sTemp As String
  Dim sCust As String
  
  'Get the position of the dummy character for the first section
  lTemp = Mid$(sAuthCode, 1, 1)
  
  'Get the first section
  sTemp = Mid$(sAuthCode, 2, 5)
  
  'Remove the dummy character from the section
  If lTemp = 1 Then
    sCust = Mid$(sTemp, 2, 4)
  ElseIf lTemp = 5 Then
    sCust = Mid$(sTemp, 1, 4)
  Else
    sCust = Mid$(sTemp, 1, lTemp - 1) & Mid$(sTemp, lTemp + 1, Len(sTemp))
  End If
  
  'Decode letters into numbers
  sTemp = ""
  For lTemp = 1 To 4
    sTemp = sTemp & AlphaCode(Mid$(sCust, lTemp, 1))
  Next
      
  'Does it match the cust no ? if so, return true
  ValidateUsers = (lCustNo = Val(sTemp))

End Function

Private Function GetModule(Module As Module, sModuleAuthCode As String, lCustNo As Long) As Boolean

  Dim lModule As Long
  Dim lCount As Long
  Dim sModCode As String
      
  'Ensure that the authcode entered matches with the customer number. If
  'not, exit the function with GetModule = False. If so, check the
  'authcode in more detail.
  
  If ValidateCustNo(sModuleAuthCode, lCustNo) Then
      
    'Loop thru the authcode, looking at the 2nd, 4th, 6th, 8th etc
    'characters. If they are letters which correspond to modules
    'the return true, otherwise they are dummy numbers so return false
    
    For lCount = 1 To Left(sModuleAuthCode, 1)
      If AlphaCode(Mid$(sModuleAuthCode, (lCount * 2), 1)) = Module Then
        GetModule = True
        Exit Function
      End If
    Next
  End If

End Function

Private Function ValidateCustNo(sAuthCode As String, lCustNo As Long) As Boolean

  Dim lCount As Long
  Dim lPos As Long
  Dim sCust As String
  
  'Ensure the customer number is 4 digits long
  sCust = Format$(lCustNo, "0000")
  
  'Examine the 3rd, 4th and 11th character of the authcode and check if it
  'matches the 1st, 2nd and 3rd character of the customer number. If not,
  'return false.
  
  For lCount = 3 To 11 Step 4
    Select Case lCount
      Case 3
        lPos = 1
      Case 7
        lPos = 2
      Case Else
        lPos = 3
    End Select
    If Mid$(sAuthCode, lCount, 1) <> Mid$(sCust, lPos, 1) Then
      ValidateCustNo = False
      Exit Function
    End If
  Next
  
  'Check the 13th character and see if it matches the 4th character of
  'the customer number. If not, return false.
  
  If Mid$(sAuthCode, 13, 1) <> Mid$(sCust, 4, 1) Then
    ValidateCustNo = False
    Exit Function
  End If

  'If we get to here, then the authcode is a valid one for the specified
  'customer number.
  
  ValidateCustNo = True

End Function


Public Function ConvertOldLicenceToNew(lngCustomerNo As Long, strOldUserLicence As String, strOldModuleLicence As String) As String

  Dim strOutput As String
  Dim lngNumUsers As Long
  Dim lngIntUsers As Long
  Dim lngSSUsers As Long
  Dim lngModules As Long
  Dim strCustNo As String
  Dim strUsers As String
  Dim strModules As String
  Dim lngCount As Long
  Dim datGeneral As clsGeneral
  Dim objSecondLicence As clsLicence

  'Once all customers are on at least v1.28 this (along with all old licence stuff can probably be removed).


'READ OLD LICENCE
  If strOldUserLicence Like "????-????-????-????" Then
    'Second Way...
    Set objSecondLicence = New clsLicence
    objSecondLicence.LicenceKey = strOldUserLicence
    lngCustomerNo = objSecondLicence.CustomerNo
    lngNumUsers = objSecondLicence.NumUsers
    lngModules = objSecondLicence.Modules
    Set objSecondLicence = Nothing

    If (lngModules And 16) Then   'If intranet enabled
      lngIntUsers = 10000
      lngSSUsers = 10000
    Else
      lngIntUsers = 0
      lngSSUsers = 0
    End If

  Else
    'First Way...
    lngNumUsers = GetUsers(lngCustomerNo, strOldUserLicence)
  
    lngModules = _
      IIf(GetModule(Personnel, strOldModuleLicence, lngCustomerNo), 1, 0) + _
      IIf(GetModule(Recruitment, strOldModuleLicence, lngCustomerNo), 2, 0) + _
      IIf(GetModule(Absence, strOldModuleLicence, lngCustomerNo), 4, 0) + _
      IIf(GetModule(Training, strOldModuleLicence, lngCustomerNo), 8, 0) + _
      IIf(GetModule(Web, strOldModuleLicence, lngCustomerNo), 16, 0) + _
      IIf(GetModule(Afd, strOldModuleLicence, lngCustomerNo), 32, 0) + _
      64  'Give Full System Manager to existing customers.
  
    If GetModule(Web, strOldModuleLicence, lngCustomerNo) Then
      lngIntUsers = 10000
      lngSSUsers = 10000
    Else
      lngIntUsers = 0
      lngSSUsers = 0
    End If
  End If


'CREATE NEW LICENCE (Third Way!)
  strOutput = ""
  If lngCustomerNo > 0 And lngNumUsers > 0 And lngModules > 0 Then
    Randomize Timer
    Set datGeneral = New clsGeneral
    strOutput = datGeneral.CreateKey2(lngCustomerNo, lngNumUsers, lngIntUsers, lngSSUsers, lngModules)
    Set datGeneral = Nothing
  End If

  ConvertOldLicenceToNew = strOutput

End Function


Public Function ConvertOldLicenceToNew2(lngCustomerNo As Long, strOldUserLicence As String, strOldModuleLicence As String) As String

  Dim strOutput As String
  Dim lngDAT As Long
  Dim lngDMIM As Long
  Dim lngDMIS As Long
  Dim lngSSI As Long
  Dim lngModules As Long
  Dim strCustNo As String
  Dim strUsers As String
  Dim strModules As String
  Dim lngCount As Long
  Dim datGeneral As clsGeneral
  Dim objSecondLicence As clsLicence

  'Once all customers are on at least v1.28 this (along with all old licence stuff can probably be removed).


'READ OLD LICENCE
  If strOldUserLicence Like "????-????-????-????-????" Then
    'Third Way
    Set objSecondLicence = New clsLicence
    objSecondLicence.LicenceKey2 = strOldUserLicence
    lngCustomerNo = objSecondLicence.CustomerNo
    lngDAT = objSecondLicence.NumUsers
    lngDMIM = objSecondLicence.IntUsers
    lngDMIS = objSecondLicence.SSUsers
    lngModules = objSecondLicence.Modules
    Set objSecondLicence = Nothing

  ElseIf strOldUserLicence Like "????-????-????-????" Then
    'Second Way...
    Set objSecondLicence = New clsLicence
    objSecondLicence.LicenceKey = strOldUserLicence
    lngCustomerNo = objSecondLicence.CustomerNo
    lngDAT = objSecondLicence.NumUsers
    lngModules = objSecondLicence.Modules
    Set objSecondLicence = Nothing

    If (lngModules And 16) Then   'If intranet enabled
      lngDMIM = 10000
      lngDMIS = 10000
    Else
      lngDMIM = 0
      lngDMIS = 0
    End If

  Else
    'First Way...
    lngDAT = GetUsers(lngCustomerNo, strOldUserLicence)
  
    lngModules = _
      IIf(GetModule(Personnel, strOldModuleLicence, lngCustomerNo), 1, 0) + _
      IIf(GetModule(Recruitment, strOldModuleLicence, lngCustomerNo), 2, 0) + _
      IIf(GetModule(Absence, strOldModuleLicence, lngCustomerNo), 4, 0) + _
      IIf(GetModule(Training, strOldModuleLicence, lngCustomerNo), 8, 0) + _
      IIf(GetModule(Web, strOldModuleLicence, lngCustomerNo), 16, 0) + _
      IIf(GetModule(Afd, strOldModuleLicence, lngCustomerNo), 32, 0) + _
      64  'Give Full System Manager to existing customers.
  
    If GetModule(Web, strOldModuleLicence, lngCustomerNo) Then
      lngDMIM = 10000
      lngDMIS = 10000
    Else
      lngDMIM = 0
      lngDMIS = 0
    End If
  End If


'CREATE NEW LICENCE (Fourth Way!)
  strOutput = ""
  If lngCustomerNo > 0 And lngDAT > 0 And lngModules > 0 Then
    Randomize Timer
    Set datGeneral = New clsGeneral
    
    If lngDMIM >= 999 And lngDMIS >= 999 Then
      lngDMIM = 999
      lngDMIS = 999
      lngSSI = 999
    ElseIf lngDMIS < 999 Then
      lngSSI = lngDMIM + lngDMIS
    Else
      lngDMIS = lngDMIM
      lngSSI = 999
    End If

    If lngDAT > 999 Then lngDAT = 999
    If lngDMIM > 999 Then lngDMIM = 999
    If lngDMIS > 999 Then lngDMIS = 999
    If lngSSI > 999 Then lngSSI = 999
    
    strOutput = datGeneral.CreateKey3(lngCustomerNo, lngDAT, lngDMIM, lngDMIS, lngSSI, lngModules)
    Set datGeneral = Nothing
  End If

  ConvertOldLicenceToNew2 = strOutput

End Function

