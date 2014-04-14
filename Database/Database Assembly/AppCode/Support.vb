Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.SqlServer.Server
Imports Assembly.General

Partial Public Class Support
    <Microsoft.SqlServer.Server.SqlFunction(Name:="udf_ASRNetGetSupportString", DataAccess:=DataAccessKind.None)> _
    Public Shared Function GetSupportString() As SqlString
        Dim arrKeys As New ArrayList()

        Dim randomNum As New Random()
        'Only go up to 999 as cust nums start at 1000!
        arrKeys.Add(randomNum.Next(1, 999))
        arrKeys.Add(randomNum.Next(1, 999))
        arrKeys.Add(randomNum.Next(1, 999))

        Return New SqlString(CreateKey(CInt(arrKeys.Item(0)), CInt(arrKeys.Item(1)), CInt(arrKeys.Item(2))))
    End Function

    <Microsoft.SqlServer.Server.SqlFunction(Name:="udf_ASRNetCheckSupportInputString", DataAccess:=DataAccessKind.None)> _
    Public Shared Function CheckSupportInputString(ByVal licenceKey As String, ByVal supportString As String) As SqlBoolean

        Dim returnBoolean As Boolean = False

        Try
            'Check Input is in the right format
            If Not (licenceKey Like "????-????-????-????") Then
                Return New SqlBoolean(returnBoolean)
            End If

            'Now validate that the key breakdown matches!
            If licenceKey <> supportString Then

                Dim A As Licence = New Licence()
                Dim B As Licence = New Licence()

                A.LicenceKey = licenceKey

                If B.CustomerNo > 0 Then
                    A.LicenceKey = supportString
                    returnBoolean = (A.CustomerNo = B.CustomerNo And A.DATUsers = B.DATUsers And A.Modules = B.Modules)
                End If

                A = Nothing
                B = Nothing
            End If
        Catch
        End Try

        Return New SqlBoolean(returnBoolean)

    End Function

    <Microsoft.SqlServer.Server.SqlFunction(Name:="udf_ASRNetGetSupportString2", DataAccess:=DataAccessKind.None)> _
    Public Shared Function GetSupportString2() As SqlString
        Dim arrKeys As New ArrayList()

        Dim randomNum As New Random()
        'Only go up to 999 as cust nums start at 1000!
        arrKeys.Add(randomNum.Next(1, 999))
        arrKeys.Add(randomNum.Next(1, 32678))
        arrKeys.Add(randomNum.Next(1, 32678))
        arrKeys.Add(randomNum.Next(1, 32678))
        arrKeys.Add(randomNum.Next(1, 32678))

        Return New SqlString( _
            CreateKey2(CInt(arrKeys.Item(0)), CInt(arrKeys.Item(1)), _
                CInt(arrKeys.Item(2)), CInt(arrKeys.Item(3)), CInt(arrKeys.Item(4))))
    End Function

    <Microsoft.SqlServer.Server.SqlFunction(Name:="udf_ASRNetCheckSupportInputString2", DataAccess:=DataAccessKind.None)> _
    Public Shared Function CheckSupportInputString2(ByVal licenceKey As String, ByVal supportString As String) As SqlBoolean

        Dim returnBoolean As Boolean = False

        Try
            'Check Input is in the right format
            If Not (licenceKey Like "????-????-????-????-????") Then
                Return New SqlBoolean(returnBoolean)
            End If

            'Now validate that the key breakdown matches!
            If licenceKey <> supportString Then

                Dim A As Licence = New Licence()
                Dim B As Licence = New Licence()

                B.LicenceKey2 = licenceKey

                If B.CustomerNo > 0 Then
                    A.LicenceKey2 = supportString
                    returnBoolean = _
                        (A.CustomerNo = B.CustomerNo And _
                         A.DATUsers = B.DATUsers And _
                         A.DMIMUsers = B.DMIMUsers And _
                         A.SSIUsers = B.SSIUsers And _
                         A.Modules = B.Modules)
                End If

                A = Nothing
                B = Nothing
            End If
        Catch
            Return New SqlBoolean(returnBoolean)
        End Try

        Return New SqlBoolean(returnBoolean)

    End Function
End Class
