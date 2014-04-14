'Imports System
'Imports System.Data
'Imports System.Data.SqlClient
'Imports System.Data.SqlTypes
'Imports Microsoft.SqlServer.Server
'Imports HRProServer.Net.General

Partial Public Class OldLicence

  'Private Enum Modules
  '    Personnel = 1
  '    Recruitment = 2
  '    Absence = 4
  '    Training = 8
  '    Skills = 16
  '    Web = 32
  '    Afd = 64
  'End Enum

  'Private Shared Function GetUsers(ByVal custNo As Int32, ByVal authCode As String) As Int32

  '    Dim tempInt As Int32 = 0
  '    Dim tempString As String = String.Empty
  '    Dim usersString As String = String.Empty

  '    Try
  '        'Check that the authorisation code matches the customer number
  '        If ValidateUsers(authCode, custNo) Then

  '            'Get the position of the dummy character in the second section
  '            tempInt = CInt(authCode.Substring(6, 1))

  '            'Get the second section of the authcode
  '            tempString = authCode.Substring(7, 4)

  '            'Extract the dummy character from the second section
  '            If tempInt = 1 Then
  '                usersString = tempString.Substring(1, 3)
  '            ElseIf tempInt = 4 Then
  '                usersString = tempString.Substring(0, 3)
  '            Else
  '                usersString = tempString.Substring(0, tempInt - 1) & _
  '                                tempString.Substring(tempInt, tempString.Length)
  '            End If

  '            'Decode and check it matches the no. users
  '            tempString = String.Empty
  '            For tempInt = 1 To 3
  '                tempString &= AlphaCode(usersString.Substring(tempInt - 1, 1))
  '            Next

  '            Return CInt(tempString)
  '        Else
  '            Return 0
  '        End If

  '    Catch ex As Exception
  '        Return 0
  '    End Try

  'End Function

  'Private Shared Function AlphaCode(ByVal sChar As String) As Int32

  '    'The Alphacode key
  '    Select Case sChar.ToUpper()
  '        Case "A"
  '            Return 1
  '        Case "B"
  '            Return 2
  '        Case "C"
  '            Return 3
  '        Case "D"
  '            Return 4
  '        Case "E"
  '            Return 5
  '        Case "F"
  '            Return 6
  '        Case "G"
  '            Return 7
  '        Case "H"
  '            Return 8
  '        Case "I"
  '            Return 9
  '        Case "J"
  '            Return 0
  '        Case "P"
  '            Return 16
  '        Case "Z"
  '            Return 32
  '        Case "Y"
  '            Return 64
  '    End Select

  'End Function

  'Private Shared Function ValidateUsers(ByVal authCode As String, ByVal custNo As Int32) As Boolean

  '    Dim tempInt As Int32 = 0
  '    Dim tempString As String = String.Empty
  '    Dim custString As String = String.Empty

  '    'Get the position of the dummy character for the first section
  '    tempInt = CInt(authCode.Substring(0, 1))

  '    'Get the first section
  '    tempString = authCode.Substring(1, 5)

  '    'Remove the dummy character from the section
  '    If tempInt = 1 Then
  '        custString = tempString.Substring(1, 4)
  '    ElseIf tempInt = 5 Then
  '        custString = tempString.Substring(0, 4)
  '    Else
  '        custString = tempString.Substring(0, tempInt - 1) & _
  '                        tempString.Substring(tempInt, tempString.Length)
  '    End If

  '    'Decode letters into numbers
  '    tempString = String.Empty
  '    For tempInt = 1 To 4
  '        tempString &= AlphaCode(custString.Substring(tempInt - 1, 1))
  '    Next

  '    'Does it match the cust no ? if so, return true
  '    Return (custNo = CInt(tempString))

  'End Function

  'Private Shared Function GetModule(ByVal modules As Modules, ByVal moduleAuthCode As String, ByVal custNo As Int32) As Boolean

  '    Dim moduleInt As Int32 = 0
  '    Dim sModCode As String = String.Empty

  '    'Ensure that the authcode entered matches with the customer number. If
  '    'not, exit the function with GetModule = False. If so, check the
  '    'authcode in more detail.

  '    If ValidateCustNo(moduleAuthCode, custNo) Then
  '        'Loop thru the authcode, looking at the 2nd, 4th, 6th, 8th etc
  '        'characters. If they are letters which correspond to modules
  '        'the return true, otherwise they are dummy numbers so return false

  '        For counter As Int32 = 0 To CInt(moduleAuthCode.Substring(0, 1)) - 1
  '            If AlphaCode(moduleAuthCode.Substring((counter * 2), 1)) = modules Then
  '                Return True
  '            End If
  '        Next
  '    End If

  'End Function

  'Private Shared Function ValidateCustNo(ByVal authCode As String, ByVal custNo As Int32) As Boolean

  '    'Ensure the customer number is 4 digits long
  '    Dim custString As String = custNo.ToString().PadLeft(4, "0"c)

  '    'Examine the 3rd, 4th and 11th character of the authcode and check if it
  '    'matches the 1st, 2nd and 3rd character of the customer number. If not,
  '    'return false.
  '    For counter As Int32 = 2 To 10 Step 4

  '        Dim lPos As Int32 = 0

  '        Select Case counter
  '            Case 3
  '                lPos = 0
  '            Case 7
  '                lPos = 1
  '            Case Else
  '                lPos = 2
  '        End Select
  '        If authCode.Substring(counter, 1) <> custString.Substring(lPos, 1) Then
  '            Return False
  '        End If
  '    Next

  '    'Check the 13th character and see if it matches the 4th character of
  '    'the customer number. If not, return false.
  '    If authCode.Substring(12, 1) <> custString.Substring(3, 1) Then
  '        Return False
  '    End If

  '    'If we get to here, then the authcode is a valid one for the specified
  '    'customer number.
  '    Return True

  'End Function

  '<Microsoft.SqlServer.Server.SqlFunction(Name:="udf_ASRNetConvertOldLicenceToNew", DataAccess:=DataAccessKind.None)> _
  'Public Shared Function ConvertOldLicenceToNew(ByVal customerNo As Int32, _
  '    ByVal strOldUserLicence As String, ByVal oldModuleLicence As String) As SqlString

  '    Dim strOutput As String = String.Empty
  '    Dim DATUsers As Int32 = 0
  '    Dim DMIMUsers As Int32 = 0
  '    Dim SSIUsers As Int32 = 0
  '    Dim userModules As Int32 = 0

  '    'Once all customers are on at least v1.28 this (along with all old licence stuff can probably be removed).

  '    'READ OLD LICENCE
  '    If strOldUserLicence Like "????-????-????-????" Then
  '        'Second Way...
  '        Dim secondLicence As Licence = New Licence()
  '        secondLicence.LicenceKey = strOldUserLicence
  '        customerNo = secondLicence.CustomerNo
  '        DATUsers = secondLicence.DATUsers
  '        userModules = secondLicence.Modules
  '        secondLicence = Nothing

  '        If CBool(userModules And 16) Then   'If intranet enabled
  '            DMIMUsers = 10000
  '            SSIUsers = 10000
  '        Else
  '            DMIMUsers = 0
  '            SSIUsers = 0
  '        End If

  '    Else
  '        'First Way...
  '        DATUsers = GetUsers(customerNo, strOldUserLicence)

  '        userModules = _
  '          CInt(IIf(GetModule(Modules.Personnel, oldModuleLicence, customerNo), 1, 0)) + _
  '          CInt(IIf(GetModule(Modules.Recruitment, oldModuleLicence, customerNo), 2, 0)) + _
  '          CInt(IIf(GetModule(Modules.Absence, oldModuleLicence, customerNo), 4, 0)) + _
  '          CInt(IIf(GetModule(Modules.Training, oldModuleLicence, customerNo), 8, 0)) + _
  '          CInt(IIf(GetModule(Modules.Web, oldModuleLicence, customerNo), 16, 0)) + _
  '          CInt(IIf(GetModule(Modules.Afd, oldModuleLicence, customerNo), 32, 0)) + _
  '          64  'Give Full System Manager to existing customers.

  '        If GetModule(Modules.Web, oldModuleLicence, customerNo) Then
  '            DMIMUsers = 10000
  '            SSIUsers = 10000
  '        Else
  '            DMIMUsers = 0
  '            SSIUsers = 0
  '        End If
  '    End If

  '    'CREATE NEW LICENCE (Third Way!)
  '    strOutput = String.Empty
  '    If customerNo > 0 And DATUsers > 0 And userModules > 0 Then
  '        Randomize(Timer)
  '        strOutput = CreateKey2(customerNo, DATUsers, DMIMUsers, SSIUsers, userModules)
  '    End If

  '    Return New SqlString(strOutput)

  'End Function

  '<Microsoft.SqlServer.Server.SqlFunction(Name:="udf_ASRNetConvertOldLicenceToNew2", DataAccess:=DataAccessKind.None)> _
  'Public Shared Function ConvertOldLicenceToNew2(ByVal customerNo As Int32, _
  '        ByVal oldUserLicence As String, ByVal oldModuleLicence As String) As SqlString

  '    Dim strOutput As String = String.Empty
  '    Dim DATUsers As Int32 = 0
  '    Dim DMIMUsers As Int32 = 0
  '    Dim DMISUsers As Int32 = 0
  '    Dim SSIUsers As Int32 = 0
  '    Dim userModules As Int32 = 0

  '    'Once all customers are on at least v1.28 this (along with all old licence stuff can probably be removed).

  '    'READ OLD LICENCE
  '    If oldUserLicence Like "????-????-????-????-????" Then
  '        'Third Way
  '        Dim secondLicence As Licence = New Licence()
  '        secondLicence.LicenceKey2 = oldUserLicence
  '        customerNo = secondLicence.CustomerNo
  '        DATUsers = secondLicence.DATUsers
  '        DMIMUsers = secondLicence.DMIMUsers
  '        DMISUsers = secondLicence.SSIUsers
  '        userModules = secondLicence.Modules
  '        secondLicence = Nothing

  '    ElseIf oldUserLicence Like "????-????-????-????" Then
  '        'Second Way...
  '        Dim secondLicence As Licence = New Licence()
  '        secondLicence.LicenceKey = oldUserLicence
  '        customerNo = secondLicence.CustomerNo
  '        DATUsers = secondLicence.DATUsers
  '        userModules = secondLicence.Modules
  '        secondLicence = Nothing

  '        If CBool(userModules And 16) Then   'If intranet enabled
  '            DMIMUsers = 10000
  '            DMISUsers = 10000
  '        Else
  '            DMIMUsers = 0
  '            DMISUsers = 0
  '        End If

  '    Else
  '        'First Way...
  '        DATUsers = GetUsers(customerNo, oldUserLicence)

  '        userModules = _
  '          CInt(IIf(GetModule(Modules.Personnel, oldModuleLicence, customerNo), 1, 0)) + _
  '          CInt(IIf(GetModule(Modules.Recruitment, oldModuleLicence, customerNo), 2, 0)) + _
  '          CInt(IIf(GetModule(Modules.Absence, oldModuleLicence, customerNo), 4, 0)) + _
  '          CInt(IIf(GetModule(Modules.Training, oldModuleLicence, customerNo), 8, 0)) + _
  '          CInt(IIf(GetModule(Modules.Web, oldModuleLicence, customerNo), 16, 0)) + _
  '          CInt(IIf(GetModule(Modules.Afd, oldModuleLicence, customerNo), 32, 0)) + _
  '          64  'Give Full System Manager to existing customers.

  '        If GetModule(Modules.Web, oldModuleLicence, customerNo) Then
  '            DMIMUsers = 10000
  '            DMISUsers = 10000
  '        Else
  '            DMIMUsers = 0
  '            DMISUsers = 0
  '        End If
  '    End If

  '    'CREATE NEW LICENCE (Fourth Way!)
  '    strOutput = String.Empty
  '    If customerNo > 0 And DATUsers > 0 And userModules > 0 Then
  '        If DMIMUsers >= 999 And DMISUsers >= 999 Then
  '            DMIMUsers = 999
  '            DMISUsers = 999
  '            SSIUsers = 999
  '        ElseIf DMISUsers < 999 Then
  '            SSIUsers = DMIMUsers + DMISUsers
  '        Else
  '            DMISUsers = DMIMUsers
  '            SSIUsers = 999
  '        End If

  '        If DATUsers > 999 Then DATUsers = 999
  '        If DMIMUsers > 999 Then DMIMUsers = 999
  '        If DMISUsers > 999 Then DMISUsers = 999
  '        If SSIUsers > 999 Then SSIUsers = 999

  '        strOutput = CreateKey3(customerNo, DATUsers, DMIMUsers, DMISUsers, SSIUsers, userModules)
  '    End If

  '    Return New SqlString(strOutput)

  'End Function

End Class
