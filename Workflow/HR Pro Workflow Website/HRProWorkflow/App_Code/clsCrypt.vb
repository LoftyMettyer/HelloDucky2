Option Strict On

Imports Microsoft.VisualBasic
Imports System

Public Class Crypt

    Private mfInitTrue As Boolean
    Private mabytArray() As Byte
    Private miHiByte As Int32
    Private miHiBound As Int32
    Private mabytAddTable(255, 255) As Byte
    Private mabytXTable(255, 255) As Byte

    Public Function DecryptString(ByVal psText As String, _
     Optional ByVal psKey As String = "", _
     Optional ByVal pfIsTextInHex As Boolean = True) As String

        Dim abytArray() As Byte
        Dim abytKey() As Byte
        Dim abytOut() As Byte

        Const ENCRYPTIONKEY As String = "jmltn"

        If Len(psKey) = 0 Then
            psKey = ENCRYPTIONKEY
        End If

        If pfIsTextInHex = True Then
            psText = DeHex(psText)
        End If

        DecryptString = psText
        abytArray = System.Text.Encoding.Default.GetBytes(psText)
        abytKey = System.Text.Encoding.Default.GetBytes(psKey)
        abytOut = DecryptByte(abytArray, abytKey)
        DecryptString = System.Text.Encoding.Default.GetString(abytOut)

    End Function

    Private Function DeHex(ByVal psData As String) As String

        ResetByteArray()
        For iCount As Int32 = 1 To Len(psData) Step 2
            Append(Chr(CInt("&H" & Mid(psData, iCount, 2))))
        Next

        DeHex = GData()

        ResetByteArray()

    End Function

    Private Function GData() As String
        Dim sStringData As String
        sStringData = Space(miHiByte)

        sStringData = System.Text.Encoding.Default.GetString(mabytArray, 0, miHiByte)

        GData = sStringData
    End Function

    Private Sub ResetByteArray()
        miHiByte = 0
        miHiBound = 1024
        ReDim mabytArray(miHiBound)

    End Sub

    Private Sub Append(ByRef psStringData As String, _
    Optional ByVal piLength As Int32 = 0)

        Dim iDataLength As Int32
        Dim abytTemp() As Byte

        If piLength > 0 Then
            iDataLength = piLength
        Else
            iDataLength = Len(psStringData)
        End If

        If iDataLength + miHiByte > miHiBound Then
            miHiBound = miHiBound + 1024
            ReDim Preserve mabytArray(miHiBound)
        End If

        ReDim abytTemp(iDataLength)
        System.Text.Encoding.Default.GetBytes(psStringData, 0, iDataLength, abytTemp, 0)
        abytTemp.CopyTo(mabytArray, miHiByte)

        miHiByte = miHiByte + iDataLength
    End Sub

    Public Function DecryptByte(ByVal pabytDs() As Byte, _
      ByVal pabytPass() As Byte) As Byte()

        Dim abytTemp() As Byte
        Dim iBound As Int32

        Call InitTbl()

        iBound = (UBound(pabytPass) - 1)
        Dim iTemp As Int32 = (UBound(pabytDs)) Mod (UBound(pabytPass) - 1)

        For iLoop As Int32 = (UBound(pabytDs)) To 1 Step -1
            If iTemp = 0 Then iTemp = iBound
            pabytDs(iLoop - 1) = mabytXTable(pabytDs(iLoop - 1), mabytAddTable(pabytDs(iLoop), pabytPass(iTemp)))
            pabytDs(iLoop) = mabytXTable(pabytDs(iLoop - 1), pabytDs(iLoop))
            pabytDs(iLoop - 1) = mabytXTable(pabytDs(iLoop - 1), mabytAddTable(pabytDs(iLoop), pabytPass(iTemp - 1)))
            iTemp = iTemp - 1
        Next

        abytTemp = pabytDs
        ReDim pabytDs((UBound(abytTemp)) - 4)

        System.Text.Encoding.Default.GetBytes(System.Text.Encoding.Default.GetString(abytTemp, 5, UBound(abytTemp) - 4)).CopyTo(pabytDs, 0)

        ReDim Preserve pabytDs(UBound(pabytDs) - 1)

        DecryptByte = pabytDs

    End Function

    Private Sub InitTbl()

        If mfInitTrue = True Then Exit Sub

        For i As Int32 = 0 To 255
            For j As Int32 = 0 To 255
                mabytXTable(i, j) = CByte(i Xor j)
                mabytAddTable(i, j) = CByte((i + j) Mod 255)
            Next j
        Next i
        mfInitTrue = True
    End Sub

    Public Function ProcessDecryptString(ByVal psText As String) As String
        Dim sOutput As String
        Dim sChar As String
        Dim sNextChar As String

        Const MARKERCHAR_1 As String = "J"
        Const MARKERCHAR_2 As String = "P"
        Const MARKERCHAR_3 As String = "D"
        Const DODGYCHARACTER_INCREMENT_1 As Int32 = 174
        Const DODGYCHARACTER_INCREMENT_2 As Int32 = 83
        Const DODGYCHARACTER_INCREMENT_3 As Int32 = 1

        ' Loop through the string, prcessing any occurrances of the MARKERCHAR.
        sOutput = ""

        ' Ignore the last character
        psText = psText.Substring(0, psText.Length - 1)

        For iLoop As Int32 = 1 To Len(psText)
            sChar = Mid(psText, iLoop, 1)

            If (sChar = MARKERCHAR_1) Then
                sNextChar = Mid(psText, iLoop + 1, 1)

                If sNextChar <> MARKERCHAR_1 Then
                    sChar = Chr(Asc(sNextChar) - DODGYCHARACTER_INCREMENT_1)
                End If

                iLoop = iLoop + 1

            ElseIf (sChar = MARKERCHAR_2) Then
                sNextChar = Mid(psText, iLoop + 1, 1)

                If sNextChar <> MARKERCHAR_2 Then
                    sChar = Chr(Asc(sNextChar) - DODGYCHARACTER_INCREMENT_2)
                End If

                iLoop = iLoop + 1

            ElseIf (sChar = MARKERCHAR_3) Then
                sNextChar = Mid(psText, iLoop + 1, 1)

                If sNextChar <> MARKERCHAR_3 Then
                    sChar = Chr(Asc(sNextChar) - DODGYCHARACTER_INCREMENT_3)
                End If

                iLoop = iLoop + 1
            End If

            sOutput = sOutput & sChar
        Next iLoop

        ProcessDecryptString = sOutput
    End Function

    Public Function DecompactString(ByVal psSourceString As String) As String
        Dim sModifiedSourceString As String
        Dim sDecompactedString As String
        Dim sSubString As String
        Dim sNewString As String
        Dim iTemp As Int32
        Dim iTempTotal As Int32
        Dim iAscCode As Int32
        Dim iTrailingCharsToIgnore As Int16

        sDecompactedString = ""
        iTrailingCharsToIgnore = CShort(Right(psSourceString, 1))
        sModifiedSourceString = Left(psSourceString, psSourceString.Length - 1)

        Do While Len(sModifiedSourceString) > 0
            sSubString = Left(sModifiedSourceString & "00", 2)
            sModifiedSourceString = Mid(sModifiedSourceString, 3)
            iTempTotal = 0

            iAscCode = Asc(Mid(sSubString, 2, 1))
            If iAscCode = 64 Then ' @
                iTemp = 63
            ElseIf iAscCode = 36 Then ' $
                iTemp = 62
            ElseIf iAscCode >= 97 Then ' a-z
                iTemp = iAscCode - 61
            ElseIf iAscCode >= 65 Then 'A-Z
                iTemp = iAscCode - 55
            Else ' 0-9
                iTemp = iAscCode - 48
            End If
            iTempTotal = iTempTotal + iTemp

            iAscCode = Asc(Left(sSubString, 1))
            If iAscCode = 64 Then ' @
                iTemp = 63
            ElseIf iAscCode = 36 Then ' $
                iTemp = 62
            ElseIf iAscCode >= 97 Then ' a-z
                iTemp = iAscCode - 61
            ElseIf iAscCode >= 65 Then 'A-Z
                iTemp = iAscCode - 55
            Else ' 0-9
                iTemp = iAscCode - 48
            End If
            iTempTotal = iTempTotal + (iTemp * 64)

            sNewString = Right("000" & Hex$(iTempTotal), 3)
            sDecompactedString = sDecompactedString & sNewString
        Loop

        sDecompactedString = Left(sDecompactedString, Len(sDecompactedString) - iTrailingCharsToIgnore)
        DecompactString = sDecompactedString
    End Function

    Public Function SimpleEncrypt(ByVal psStringToEncrypt As String, ByVal psKey As String) As String
        Dim sEncryptedString As String
        Dim iKeyTotal As Integer
        Dim iLoop As Integer
        Dim sTemp As String

        sEncryptedString = psStringToEncrypt
        Try
            iKeyTotal = 0
            For iLoop = 1 To psKey.Length
                iKeyTotal = iKeyTotal + Asc(Mid(psKey, iLoop, 1))
            Next iLoop

            sTemp = ""
            For iLoop = 1 To psStringToEncrypt.Length
                sTemp = sTemp & Right("000" & Asc(Mid(psStringToEncrypt, iLoop, 1)).ToString, 3)
            Next iLoop

            If sTemp.Length = 0 Then
                sTemp = "0"
            End If

            sEncryptedString = (CLng(sTemp) + iKeyTotal).ToString
        Catch ex As Exception
            sEncryptedString = psStringToEncrypt
        End Try

        Return sEncryptedString
    End Function
    Public Function SimpleDecrypt(ByVal psStringToDecrypt As String, ByVal psKey As String) As String
        Dim sDecryptedString As String
        Dim iKeyTotal As Integer
        Dim iLoop As Integer
        Dim sTemp As String

        sDecryptedString = ""
        Try
            iKeyTotal = 0
            For iLoop = 1 To psKey.Length
                iKeyTotal = iKeyTotal + Asc(Mid(psKey, iLoop, 1))
            Next iLoop

            psStringToDecrypt = (CLng(psStringToDecrypt) - iKeyTotal).ToString

            While psStringToDecrypt.Length > 0
                If psStringToDecrypt.Length < 3 Then
                    psStringToDecrypt = Right("000" & psStringToDecrypt, 3)
                End If

                sTemp = Mid(psStringToDecrypt, psStringToDecrypt.Length - 2, 3)
                sDecryptedString = Chr(CInt(sTemp)) & sDecryptedString
                psStringToDecrypt = Left(psStringToDecrypt, psStringToDecrypt.Length - 3)
            End While

        Catch ex As Exception
            sDecryptedString = psStringToDecrypt
        End Try

        Return sDecryptedString
    End Function
End Class
