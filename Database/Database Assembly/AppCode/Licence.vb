Imports System
Imports System.Data
Imports Assembly.General

Partial Public Class Licence
    Private _customerNo As Int32 = 0
    Private _DATUsers As Int32 = 0
    Private _DMIMUsers As Int32 = 0
    Private _DMISUsers As Int32 = 0
    Private _SSIUsers As Int32 = 0
    Private _Modules As Int32 = 0

    Public WriteOnly Property LicenceKey() As String
        Set(ByVal value As String)
            Dim custNo As String = String.Empty
            Dim dATUsers As String = String.Empty
            Dim modules As String = String.Empty
            Dim randomDigit As String = String.Empty

            Dim strTemp As String = String.Empty

            '1231-2312-3123-1233

            _customerNo = 0

            Try
                If value Like "????-????-????-????" Then
                    strTemp = Replace(value, "-", "")

                    For counter As Int32 = 0 To 14 Step 3
                        custNo &= strTemp.Substring(counter, 1)
                        dATUsers &= strTemp.Substring(counter + 1, 1)
                        modules &= strTemp.Substring(counter + 2, 1)
                    Next
                    modules &= strTemp.Substring(15, 1)

                    _customerNo = ConvertStringToNumber(1, custNo)
                    _DATUsers = ConvertStringToNumber(2, dATUsers)
                    _Modules = ConvertStringToNumber(3, modules)

                End If

                If _customerNo = 0 Or _DATUsers = 0 Or _Modules = 0 Then
                    _customerNo = 0
                    _DATUsers = 0
                    _Modules = 0
                End If
            Catch
                _customerNo = 0
                _DATUsers = 0
                _DMIMUsers = 0
                _DMISUsers = 0
                _SSIUsers = 0
                _Modules = 0
            End Try
        End Set
    End Property

    Public WriteOnly Property LicenceKey2() As String
        Set(ByVal value As String)
            Dim custNo As String = String.Empty
            Dim dATUsers As String = String.Empty
            Dim dMIMUsers As String = String.Empty
            Dim sSIUsers As String = String.Empty
            Dim modules As String = String.Empty

            Dim strTemp As String = String.Empty

            '1231-2312-3123-1233
            '1234-5123-4512-3451-2345

            _customerNo = 0

            Try
                If value Like "????-????-????-????-????" Then
                    strTemp = Replace(value, "-", "")

                    For counter As Int32 = 0 To (strTemp.Length - 1) Step 5
                        custNo &= strTemp.Substring(counter, 1)
                        dATUsers &= strTemp.Substring(counter + 1, 1)
                        dMIMUsers &= strTemp.Substring(counter + 2, 1)
                        sSIUsers &= strTemp.Substring(counter + 3, 1)
                        modules &= strTemp.Substring(counter + 4, 1)
                    Next

                    _customerNo = ConvertStringToNumber2(custNo)
                    _DATUsers = ConvertStringToNumber2(dATUsers)
                    _DMIMUsers = ConvertStringToNumber2(dMIMUsers)
                    _SSIUsers = ConvertStringToNumber2(sSIUsers)
                    _Modules = ConvertStringToNumber2(modules)

                End If

                If _customerNo = 0 Or _DATUsers = 0 Or _Modules = 0 Then
                    _customerNo = 0
                    _DATUsers = 0
                    _Modules = 0
                End If
            Catch
                _customerNo = 0
                _DATUsers = 0
                _DMIMUsers = 0
                _DMISUsers = 0
                _SSIUsers = 0
                _Modules = 0
            End Try
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
            Return _SSIUsers
        End Get
    End Property

    Public ReadOnly Property Modules() As Int32
        Get
            Return _Modules
        End Get
    End Property
End Class
