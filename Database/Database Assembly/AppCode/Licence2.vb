'Option Strict Off

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.SqlServer.Server
Imports Assembly.General

Partial Public Class Licence2

  Private Shared _licenceKey As String = String.Empty
  Private Shared _customerNo As Int32 = 0
  Private Shared _DATUsers As Int32 = 0
  Private Shared _DMIMUsers As Int32 = 0
  Private Shared _DMISUsers As Int32 = 0
  Private Shared _SSISUsers As Int32 = 0
  Private Shared _Modules As Int32 = 0

  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetIsModuleLicensed", DataAccess:=DataAccessKind.None)> _
  Public Shared Function IsModuleLicenced(ByVal inKey As String, ByVal iModuleCode As SqlInt32) As SqlBoolean

    Dim iModules As SqlInt32

    iModules = GetLicenceKey(inKey, "Modules")

    If CBool(iModules And iModuleCode) Then
      IsModuleLicenced = True
    Else
      IsModuleLicenced = False
    End If

  End Function

  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetGetLicenceKey", DataAccess:=DataAccessKind.None)> _
  Public Shared Function GetLicenceKey(ByVal inKey As String, ByVal moduleName As String) As SqlInt32

    Dim custNo As String = String.Empty
    Dim dATUsers As String = String.Empty
    Dim dMIMUsers As String = String.Empty
    Dim dMISUsers As String = String.Empty
    Dim sSISUsers As String = String.Empty
    Dim modules As String = String.Empty
    Dim randomDigit As String = String.Empty

    Dim strTemp As String = String.Empty
    Dim strInput As String = String.Empty
    Dim lngCount As Int32 = 0

    '1st Generation: AAAAAAAAAAAAAAA
    '2nd Generation: 1231-2312-3123-1233
    '3rd Generation: 1234-5123-4512-3451-2345
    '4th Generation: A5316-16426-16426-16536

    _customerNo = 0

    Try
      'Check format and licence version indicator...
      'DO NOT REMOVE THE VERSION INDICATOR (i.e. the first character)
      If inKey Like "A????-?????-?????-?????" Then

        strInput = Replace(inKey, "-", "")
        strInput = strInput.Substring(0, 1) & strInput.Substring(5, 1) & _
                   strInput.Substring(10, 1) & strInput.Substring(15, 1) & _
                   strInput.Substring(3, 1) & strInput.Substring(8, 1) & _
                   strInput.Substring(13, 1) & strInput.Substring(18, 1) & _
                   strInput.Substring(2, 1) & strInput.Substring(7, 1) & _
                   strInput.Substring(12, 1) & strInput.Substring(17, 1) & _
                   strInput.Substring(1, 1) & strInput.Substring(6, 1) & _
                   strInput.Substring(11, 1) & strInput.Substring(16, 1) & _
                   strInput.Substring(4, 1) & strInput.Substring(9, 1) & _
                   strInput.Substring(14, 1) & strInput.Substring(19, 1)

        custNo = strInput.Substring(1, 4)
        dATUsers = strInput.Substring(5, 2)
        dMIMUsers = strInput.Substring(7, 2)
        dMISUsers = strInput.Substring(9, 2)
        sSISUsers = strInput.Substring(11, 2)
        modules = strInput.Substring(13, 6)
        randomDigit = strInput.Substring(19, 1)

        _customerNo = ConvertStringToNumber2(custNo)
        _DATUsers = ConvertStringToNumber2(randomDigit & dATUsers)
        _DMIMUsers = ConvertStringToNumber2(randomDigit & dMIMUsers)
        _DMISUsers = ConvertStringToNumber2(randomDigit & dMISUsers)
        _SSISUsers = ConvertStringToNumber2(randomDigit & sSISUsers)
        _Modules = ConvertStringToNumber2(modules)

      End If

      If _customerNo = 0 Or _DATUsers = 0 Or _Modules = 0 Then
        _customerNo = 0
        _DATUsers = 0
        _DMIMUsers = 0
        _DMISUsers = 0
        _SSISUsers = 0
        _Modules = 0
      End If

      Select Case moduleName.ToLower()
        Case ("CustomerNo").ToLower()
          Return New SqlInt32(_customerNo)
        Case ("Modules").ToLower()
          Return New SqlInt32(_Modules)
        Case ("DATUsers").ToLower()
          Return New SqlInt32(_DATUsers)
        Case ("DMIMUsers").ToLower()
          Return New SqlInt32(_DMIMUsers)
        Case ("DMISUsers").ToLower()
          Return New SqlInt32(_DMISUsers)
        Case ("SSISUsers").ToLower()
          Return New SqlInt32(_SSISUsers)
        Case Else
          Throw New ArgumentException("Invalid Module name argument")
      End Select
    Catch ex As Exception
      Return New SqlInt32(0)
    Finally
      _licenceKey = String.Empty
      _customerNo = 0
      _DATUsers = 0
      _DMIMUsers = 0
      _DMISUsers = 0
      _SSISUsers = 0
      _Modules = 0
    End Try

  End Function

  Public WriteOnly Property licenceKey() As String
    Set(ByVal value As String)
      _licenceKey = value
    End Set
  End Property

  Public ReadOnly Property CustomerNo() As Int32
    Get
      Return _customerNo
    End Get
  End Property

  Public ReadOnly Property DATUsers() As Int32
    Get
      Return _DATUsers
    End Get
  End Property

  Public ReadOnly Property DMIMUsers() As Int32
    Get
      Return _DMIMUsers
    End Get
  End Property

  Public ReadOnly Property DMISUsers() As Int32
    Get
      Return _DMISUsers
    End Get
  End Property

  Public ReadOnly Property SSIUsers() As Int32
    Get
      Return _SSISUsers
    End Get
  End Property

  Public ReadOnly Property Modules() As Int32
    Get
      Return _Modules
    End Get
  End Property
End Class
