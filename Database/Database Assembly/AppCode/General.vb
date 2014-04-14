Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Reflection
Imports Microsoft.SqlServer.Server

Partial Public Class General
    <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetAssemblyVersion", DataAccess:=DataAccessKind.None)> _
      Public Shared Function AssemblyVersion() As SqlString
        Return New SqlString(System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major.ToString _
                    & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString _
                    & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build.ToString)
    End Function

    Public Shared Function GetSystemLogon() As String

        Dim returnLogon As String = String.Empty

        Try
            Using conn As New SqlConnection("context connection=true")
                Dim cmd As New SqlCommand()
                Dim sql As String = "SELECT [ParameterValue] FROM ASRSysModuleSetup WHERE [ModuleKey] = 'MODULE_SQL'" & _
                                    " AND [ParameterKey] = 'Param_FieldsLoginDetails'"
                cmd.Connection = conn
                cmd.Connection.Open()
                cmd.CommandText = sql
                cmd.CommandType = CommandType.Text

                returnLogon = CType(cmd.ExecuteScalar(), String)

                cmd.Connection.Close()
                cmd = Nothing
            End Using
        Catch ex As SqlException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try

        Return returnLogon
    End Function

  Public Shared Function GetConnectionString(ByVal userID As String, ByVal password As String, _
      ByVal databaseName As String, ByVal serverName As String, Optional ByVal appName As String = "") As String

    Dim builder As New SqlConnectionStringBuilder()
    builder.DataSource = serverName
    builder.InitialCatalog = databaseName
    builder.PacketSize = 32767

    If String.IsNullOrEmpty(appName) Then
      builder.ApplicationName = Reflection.Assembly.GetExecutingAssembly.GetName.Name
    Else
      builder.ApplicationName = appName
    End If

    builder.Pooling = False

    If Not userID.Equals(String.Empty) Then
      builder.UserID = userID
      builder.Password = password
    Else
      builder.IntegratedSecurity = True
    End If

    Return builder.ConnectionString

  End Function

    Public Shared Function FormatDateTimeWithMS(ByVal dDate As Date) As String

        Dim tempString As String

        tempString = dDate.ToString("yyyy-MM-dd hh:mm:ss") & ":" & dDate.Millisecond

        Return tempString

    End Function

    Public Shared Sub DecryptLogonDetails(ByVal input As String, ByRef userName As String, ByRef password As String, _
        ByRef database As String, ByRef server As String)

        Dim eKey As String = String.Empty
        Dim lens As String = String.Empty
        Dim start As Int32 = 0
        Dim finish As Int32 = 0

        If input = String.Empty Then
            Return
        End If

        start = input.Length - 14
        eKey = input.Substring(start, 10)
        lens = input.Substring(input.Length - 4)
        input = XOREncript(input.Substring(0, start), eKey)

        start = 0
        finish = Asc(lens.Substring(0, 1)) - 127
        userName = input.Substring(start, finish)

        start = start + finish
        finish = Asc(lens.Substring(1, 1)) - 127
        password = input.Substring(start, finish)

        start = start + finish
        finish = Asc(lens.Substring(2, 1)) - 127
        database = input.Substring(start, finish)

        start = start + finish
        finish = Asc(lens.Substring(3, 1)) - 127
        server = input.Substring(start, finish)

    End Sub

    Public Shared Function XOREncript(ByVal input As String, ByVal key As String) As String

        Dim count As Int32 = 0
        Dim output As String = String.Empty
        Dim strChar As String = String.Empty

        For count = 1 To input.Length
            strChar = key.Substring(count Mod key.Length, 1)
            output = output & Convert.ToChar(Asc(strChar) Xor Asc(input.Substring(count - 1, 1)))
        Next

        Return output

    End Function

    Public Shared Function ByteArrayToString(ByVal arrInput() As Byte) As String
        Dim i As Int32 = 0
        Dim sOutput As New Text.StringBuilder()
        sOutput.Append("0x")
        For i = 0 To arrInput.Length - 1
            sOutput.Append(arrInput(i).ToString("X2"))
        Next
        Return sOutput.ToString()
    End Function

    Private Shared Function ConvertNumberToString(ByVal lngMode As Int32, ByVal input As Int32) As String
        Dim returnString As String = String.Empty

        Dim randomDigit As Int32 = 0
        Dim alphaString As String = String.Empty

        Dim randomNum As New Random()
        randomDigit = randomNum.Next(1, 26)

        Try
            alphaString = GenerateAlphaString(randomDigit)

            returnString = alphaString.Substring((input Mod 31), 1) & Convert.ToChar(randomDigit + 64)

            If lngMode = 3 Then
                returnString += alphaString.Substring(((input \ 32768) And 31), 1)
            End If

            returnString += _
                Mid(alphaString, ((input \ 1024) And 31) + 1, 1) & _
                Mid(alphaString, ((input \ 32) And 31) + 1, 1) & _
                Mid(alphaString, (input And 31) + 1, 1)
        Catch
            Return String.Empty
        End Try

        Return returnString
    End Function

    Public Shared Function ConvertStringToNumber(ByVal mode As Int32, ByVal input As String) As Int32
        Dim returnInt As Int32 = 0
        Dim randomDigit As Int32 = 0
        Dim alphaString As String = String.Empty
        Dim output As Int32 = 0

        Try
            randomDigit = Asc(input.Substring(1, 1)) - 64
            alphaString = GenerateAlphaString(randomDigit)

            If mode = 3 Then
                output = _
                    ((InStr(alphaString, input.Substring(2, 1)) - 1) * 32768) + _
                    ((InStr(alphaString, input.Substring(3, 1)) - 1) * 1024) + _
                    ((InStr(alphaString, input.Substring(4, 1)) - 1) * 32) + _
                    (InStr(alphaString, input.Substring(5, 1)) - 1)
            Else
                output = _
                    ((InStr(alphaString, input.Substring(2, 1)) - 1) * 1024) + _
                    ((InStr(alphaString, input.Substring(3, 1)) - 1) * 32) + _
                    (InStr(alphaString, input.Substring(4, 1)) - 1)
            End If

            If alphaString.Substring((output Mod 31), 1) = input.Substring(0, 1) Then
                returnInt = output
            End If
        Catch
            Return 0
        End Try

        Return returnInt

    End Function

    Private Shared Function ConvertNumberToString2(ByVal size As Int32, ByVal input As Int32, ByVal digit As Int32) As String
        Dim returnString As String = String.Empty

        Dim randomDigit As Int32 = 0
        Dim alphaString As String = String.Empty
        Dim factor As Int32 = 32

        Dim randomNum As New Random()
        randomDigit = CInt(IIf(digit = 0, randomNum.Next(1, 26), digit))

        Try
            alphaString = GenerateAlphaString(randomDigit)

            returnString = alphaString.Substring((input And 31), 1)

            For counter As Int32 = 2 To size - CInt(IIf(digit = 0, 1, 0))
                returnString = alphaString.Substring(((input \ factor) And 31), 1) & returnString
                factor = factor * 32
            Next

            If digit = 0 Then
                returnString = Convert.ToChar(randomDigit + 64) & returnString
            End If
        Catch
            Return String.Empty
        End Try

        Return returnString
    End Function

    Public Shared Function ConvertStringToNumber2(ByVal input As String) As Int32
        Dim randomDigit As Int32 = 0
        Dim alphaString As String = String.Empty
        Dim output As Int32 = 0
        Dim factor As Int32 = 32

        Try
            randomDigit = Asc(input.Substring(0, 1)) - 64
            alphaString = GenerateAlphaString(randomDigit)

            output = (alphaString.IndexOf(input.Substring(input.Length - 1, 1)))

            For counter As Int32 = input.Length - 1 To 2 Step -1
                output += ((alphaString.IndexOf(input.Substring(counter - 1, 1))) * factor)
                factor = factor * 32
            Next

            Return output
        Catch
            Return 0
        End Try

        Return output
    End Function

    Private Shared Function GenerateAlphaString(ByVal gap As Int32) As String

        Dim output As String = String.Empty

        Try
            For looper As Int32 = 0 To gap - 1

                For counter As Int32 = Asc("A") + looper To Asc("Z") Step gap
                    If ("IOQ").IndexOf(Convert.ToChar(counter)) = -1 Then
                        output = Convert.ToChar(counter) & output
                    End If
                Next

                For counter As Int32 = Asc("1") + looper To Asc("9") Step gap
                    output = Convert.ToChar(counter) & output
                    Debug.Print(Convert.ToString("9"))
                Next

            Next
        Catch
            Return String.Empty
        End Try

        Return output

    End Function

    Public Shared Function CreateKey(ByVal lngC As Int32, ByVal lngN As Int32, ByVal lngM As Int32) As String
        Dim returnString As String = String.Empty
        Dim arrString As New ArrayList()

        Try
            arrString.Add(ConvertNumberToString(1, lngC))
            arrString.Add(ConvertNumberToString(2, lngN))
            arrString.Add(ConvertNumberToString(3, lngM))

            For counter As Int32 = 0 To 4
                returnString &= _
                  arrString.Item(0).ToString().Substring(counter, 1) & _
                  arrString.Item(1).ToString().Substring(counter, 1) & _
                  arrString.Item(2).ToString().Substring(counter, 1)
            Next
            returnString &= arrString.Item(2).ToString().Substring(5, 1)

            returnString = _
                returnString.Substring(0, 4) & "-" & returnString.Substring(4, 4) & "-" & _
                returnString.Substring(8, 4) & "-" & returnString.Substring(12, 4)
        Catch
            Return String.Empty
        End Try

        Return returnString
    End Function

    Public Shared Function CreateKey2(ByVal lngC As Int32, ByVal lngN As Int32, ByVal lngI As Int32, _
        ByVal lngS As Int32, ByVal lngM As Int32) As String

        Dim tempString As String = String.Empty
        Dim returnString As String = String.Empty
        Dim arrString As New ArrayList()

        Try
            arrString.Add(ConvertNumberToString2(4, lngC, 0))
            arrString.Add(ConvertNumberToString2(4, lngN, 0))
            arrString.Add(ConvertNumberToString2(4, lngI, 0))
            arrString.Add(ConvertNumberToString2(4, lngS, 0))
            arrString.Add(ConvertNumberToString2(4, lngM, 0))

            For counter As Int32 = 0 To 3
                tempString &= _
                  arrString.Item(0).ToString().Substring(counter, 1) & _
                  arrString.Item(1).ToString().Substring(counter, 1) & _
                  arrString.Item(2).ToString().Substring(counter, 1) & _
                  arrString.Item(3).ToString().Substring(counter, 1) & _
                  arrString.Item(4).ToString().Substring(counter, 1)
            Next

            For counter As Int32 = 0 To tempString.Length - 1 Step 4
                returnString &= String.Concat(IIf(returnString <> String.Empty, "-", ""), returnString.Substring(counter, 4))
            Next
        Catch
            Return String.Empty
        End Try

        Return returnString
    End Function

    Public Shared Function CreateKey3(ByVal customerNo As Int32, ByVal lngDAT As Int32, _
        ByVal lngDMIM As Int32, ByVal lngDMIS As Int32, ByVal lngSSI As Int32, ByVal modules As Int32) As String

        Dim custNoString As String = String.Empty
        Dim DATString As String = String.Empty
        Dim DMIMString As String = String.Empty
        Dim DMISString As String = String.Empty
        Dim SSIString As String = String.Empty
        Dim ModulesString As String = String.Empty
        Dim VersionString As String = String.Empty

        Dim randomDigit As Int32 = 0
        Dim output As String = String.Empty
        Dim returnString As String = String.Empty

        'If Valid Then
        Dim randomNum As New Random()
        randomDigit = randomNum.Next(1, 26)

        Try
            '******************************************************
            '* WHEN THE LICENCE NUMBER CHANGES PLEASE CHANGE THIS *
            '* INDICATOR TO THE NEXT LETTER IN THE ALPHABET       *
            '* (WE CAN KEEP THE FORMAT: ?????-?????-?????-?????   *
            '******************************************************
            VersionString = "A"      'Licence Version Indicator

            custNoString = ConvertNumberToString2(4, customerNo, 0)
            DATString = ConvertNumberToString2(2, lngDAT, randomDigit)
            DMIMString = ConvertNumberToString2(2, lngDMIM, randomDigit)
            DMISString = ConvertNumberToString2(2, lngDMIS, randomDigit)
            SSIString = ConvertNumberToString2(2, lngSSI, randomDigit)
            ModulesString = ConvertNumberToString2(6, modules, 0)

            output = _
                VersionString & custNoString & DATString & DMIMString & _
                DMISString & SSIString & ModulesString & Convert.ToChar(randomDigit + 64)

            'Jumble it up!
            For counter As Int32 = 0 To 3
                returnString &= _
                  CStr(IIf(returnString <> String.Empty, "-", "")) & _
                    output.Substring(counter, 1) & _
                    output.Substring(counter + 12, 1) & _
                    output.Substring(counter + 8, 1) & _
                    output.Substring(counter + 4, 1) & _
                    output.Substring(counter + 16, 1)
            Next
        Catch
            Return String.Empty
        End Try

        Return returnString
    End Function

    Public Shared ReadOnly Property ContextServerName() As String
        Get
            Dim returnServer As String = String.Empty

            Try
                Using conn As New SqlConnection("context connection=true")
                    Dim cmd As New SqlCommand()
                    ' AJE20090114 Fault #13490
                    'Dim sql As String = "SELECT @@SERVERNAME"
                    Dim sql As String = "SELECT SERVERPROPERTY('servername')"
                    cmd.Connection = conn
                    cmd.Connection.Open()
                    cmd.CommandText = sql
                    cmd.CommandType = CommandType.Text

                    returnServer = CType(cmd.ExecuteScalar(), String)

                    cmd.Connection.Close()
                    cmd = Nothing
                End Using
            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try

            Return returnServer
        End Get
    End Property

    Public Shared ReadOnly Property ContextDatabaseName() As String
        Get
            Dim returnDB As String = String.Empty

            Try
                Using conn As New SqlConnection("context connection=true")
                    Dim cmd As New SqlCommand()
                    Dim sql As String = "SELECT DB_NAME()"
                    cmd.Connection = conn
                    cmd.Connection.Open()
                    cmd.CommandText = sql
                    cmd.CommandType = CommandType.Text

                    returnDB = CType(cmd.ExecuteScalar(), String)

                    cmd.Connection.Close()
                    cmd = Nothing
                End Using
            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try

            Return returnDB
        End Get
    End Property

    Public Shared Function ProcessEncryptedString(ByVal psString As String) As String

        Dim sOutput As String = String.Empty
        Dim lngLoop As Integer = 0
        Dim sOutputPreProcess As String = psString
        Dim sSubTemp As String = String.Empty

        Const MARKERCHAR_1 As String = "J"
        Const MARKERCHAR_2 As String = "P"
        Const MARKERCHAR_3 As String = "D"
        Const DODGYCHARACTER_INCREMENT_1 As Integer = 174
        Const DODGYCHARACTER_INCREMENT_2 As Integer = 83
        Const DODGYCHARACTER_INCREMENT_3 As Integer = 1

        ' Loop through the output replacing dodgy characters with a MARKERCHAR and a safe character offset from the dodgy character.
        ' This is to avoid the dodgy characters messing up the querystring when used in a link to the Workflow website.
        Do While lngLoop < sOutputPreProcess.Length - 1
            ' Process the next character.

            Dim sChar As String = sOutputPreProcess.Substring(lngLoop, 1)

            Dim sNextChar As String = sOutputPreProcess.Substring(lngLoop + 1, 1)

            Dim iJumpChars As Int16 = 1

            If (sChar = MARKERCHAR_1 Or sChar = MARKERCHAR_2 Or sChar = MARKERCHAR_3) Then
                If sChar <> sNextChar Then
                    Dim iAscCode As Integer = Asc(sNextChar)

                    If sChar = MARKERCHAR_1 Then
                        ' Dodgy character marker. Must remove the MARKERCHAR_1 and substract 
                        ' DODGYCHARACTER_INCREMENT_1 from the dodgy character's ASC value.
                        sOutput &= Chr(iAscCode - DODGYCHARACTER_INCREMENT_1)
                    ElseIf sChar = MARKERCHAR_2 Then
                        ' Dodgy character marker. Must remove the MARKERCHAR_2 and substract 
                        ' DODGYCHARACTER_INCREMENT_2 from the dodgy character's ASC value.
                        sOutput &= Chr(iAscCode - DODGYCHARACTER_INCREMENT_2)
                    ElseIf sChar = MARKERCHAR_3 Then
                        ' Dodgy character marker. Must remove the MARKERCHAR_3 and substract 
                        ' DODGYCHARACTER_INCREMENT_3 from the dodgy character's ASC value.
                        sOutput &= Chr(iAscCode - DODGYCHARACTER_INCREMENT_3)
                    End If
                    iJumpChars = 2
                Else
                    ' NOT a dodgy character. Put it straight in the output string with out reprocessing.
                    sOutput = sOutput & sChar
                    iJumpChars = 2
                End If
            Else
                sOutput = sOutput & sChar
                iJumpChars = 1
            End If

            lngLoop += iJumpChars
        Loop

        ' process the last character now
        ' NPG Fault 13373
        ' If lngLoop <= sOutputPreProcess.Length Then
        If lngLoop < sOutputPreProcess.Length Then
            sOutput += sOutputPreProcess.Substring(lngLoop, 1)
        End If

        Return sOutput

    End Function

    '****************************************************************
    ' NullSafeString
    '****************************************************************
    Public Shared Function NullSafeString(ByVal arg As Object, _
    Optional ByVal returnIfEmpty As String = "") As String

        Dim returnValue As String

        If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
            OrElse (arg Is String.Empty) Then
            returnValue = returnIfEmpty
        Else
            Try
                returnValue = CStr(arg)
            Catch
                returnValue = returnIfEmpty
            End Try

        End If

        Return returnValue

    End Function

    '****************************************************************
    ' NullSafeInteger
    '****************************************************************
    Public Shared Function NullSafeInteger(ByVal arg As Object, _
      Optional ByVal returnIfEmpty As Integer = 0) As Integer

        Dim returnValue As Integer

        If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
            OrElse (arg Is String.Empty) Then
            returnValue = returnIfEmpty
        Else
            Try
                returnValue = CInt(arg)
            Catch
                returnValue = returnIfEmpty
            End Try
        End If

        Return returnValue

    End Function
End Class
