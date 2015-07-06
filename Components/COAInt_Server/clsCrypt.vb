Option Strict Off

Imports System
Imports Microsoft.VisualBasic

Public Class clsCrypt

	Private mfInitTrue As Boolean
	Private mabytArray() As Byte
	Private miHiByte As Int32
	Private miHiBound As Int32
	Private mabytAddTable(255, 255) As Byte
	Private mabytXTable(255, 255) As Byte

	Public Function EncryptString(ByRef psText As String, Optional ByRef psKey As String = "", Optional ByRef pbOutputInHex As Boolean = False) As String

		Dim abytArray() As Byte
		Dim abytKey() As Byte
		Dim abytOut() As Byte

		psText = psText & " "
		abytArray = System.Text.Encoding.Default.GetBytes(psText)
		abytKey = System.Text.Encoding.Default.GetBytes(psKey)
		abytOut = EncryptByte(abytArray, abytKey)
		EncryptString = System.Text.Encoding.Default.GetString(abytOut)

		If pbOutputInHex = True Then EncryptString = EnHex(EncryptString)
	End Function

	Public Function EnHex(ByRef psData As String) As String

		Dim dblCount As Double
		Dim sTemp As String

		Reset()

		For dblCount = 1 To Len(psData)
			sTemp = Hex(Asc(Mid(psData, dblCount, 1)))
			If Len(sTemp) < 2 Then sTemp = "0" & sTemp
			Append(sTemp)
		Next

		EnHex = GData()

		Reset()
	End Function

	Public Function EncryptByte(ByRef pabytText() As Byte, ByRef pabytKey() As Byte) As Object

		Dim abytTemp() As Byte
		Dim iTemp As Short
		Dim iLoop As Integer
		Dim iBound As Short

		Call InitTbl()

		ReDim abytTemp((UBound(pabytText)) + 5)
		Randomize()
		abytTemp(0) = Int((Rnd() * 254) + 1)
		abytTemp(1) = Int((Rnd() * 254) + 1)
		abytTemp(2) = Int((Rnd() * 254) + 1)
		abytTemp(3) = Int((Rnd() * 254) + 1)
		abytTemp(4) = Int((Rnd() * 254) + 1)

		pabytText.CopyTo(abytTemp, 5)

		ReDim Preserve abytTemp(UBound(abytTemp) - 1)

		ReDim pabytText(UBound(abytTemp))
		abytTemp.CopyTo(pabytText, 0)

		ReDim abytTemp(0)
		iBound = (UBound(pabytKey) - 1)
		iTemp = 0

		For iLoop = 0 To UBound(pabytText) - 1
			If iTemp = iBound Then iTemp = 0
			pabytText(iLoop) = mabytXTable(pabytText(iLoop), mabytAddTable(pabytText(iLoop + 1), pabytKey(iTemp)))
			pabytText(iLoop + 1) = mabytXTable(pabytText(iLoop), pabytText(iLoop + 1))
			pabytText(iLoop) = mabytXTable(pabytText(iLoop), mabytAddTable(pabytText(iLoop + 1), pabytKey(iTemp + 1)))
			iTemp = iTemp + 1
		Next iLoop

		EncryptByte = pabytText.Clone
	End Function

	Public Function DecryptString(ByVal psText As String, Optional ByVal psKey As String = "", Optional ByVal pfIsTextInHex As Boolean = True) As String

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

	Private Sub Append(ByRef psStringData As String, Optional ByVal piLength As Int32 = 0)

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

	Public Function DecryptByte(ByVal pabytDs() As Byte, ByVal pabytPass() As Byte) As Byte()

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

	Public Function EncryptQueryString(ByVal plngInstanceID As Long,
																		 ByVal plngStepID As Long,
																		 ByVal psUser As String,
																		 ByVal psPassword As String,
																		 ByVal psServer As String,
																		 ByVal psDatabase As String,
																		 ByVal psLoginKey As String,
																		 ByVal psPasswordKey As String) As String

		Dim sSourceString As String
		Dim sEncryptedString As String

		Try
			sSourceString = CStr(plngInstanceID) &
											vbTab & CStr(plngStepID) &
											vbTab & psUser &
											vbTab & psPassword &
											vbTab & psServer &
											vbTab & psDatabase &
											vbTab & psLoginKey &
											vbTab & psPasswordKey

			sEncryptedString = EncryptString(sSourceString, "jmltn", True)
			sEncryptedString = CompactString(sEncryptedString)

		Catch ex As Exception
			sEncryptedString = ""

		End Try

		EncryptQueryString = sEncryptedString
	End Function

	Public Function CompactString(ByRef psSourceString As String) As String
		' Compact the encrypted string.
		' psSourceString is a string of the hexadecimal values of the Ascii codes for each character in the encrypted string.
		' In this string each character in the encrypted string is represented as 2 hex digits.
		' As it's a string of hex characters all characters are in the range 0-9, A-F
		' Valid hypertext link characters are 0-9, A-Z, a-z and some others (we'll be using $ and @).
		' Take advantage of this by implementing our own base64 encoding as follows:
		Dim sCompactedString As String
		Dim sSubString As String
		Dim sModifiedSourceString As String
		Dim iValue As Short
		Dim iTemp As Short
		Dim sNewString As String

		sCompactedString = ""
		sModifiedSourceString = psSourceString
		Do While Len(sModifiedSourceString) > 0
			' Read the hex characters in chunks of 3 (ie. possible values 0 - 4095)
			' This chunk of 3 Hex characters can then be translated into 2 base64 characters (ie. still have possible values 0 - 4095)
			' Woohoo! We've reduced the length of the encrypted string by about one third!
			sNewString = ""
			sSubString = Left(sModifiedSourceString & "000", 3)
			sModifiedSourceString = Mid(sModifiedSourceString, 4)
			iValue = Val("&H" & sSubString)

			' Use our own base64 digit set.
			' Base64 digit values 0-9 are represented as 0-9
			' Base64 digit values 10-35 are represented as A-Z
			' Base64 digit values 36-61 are represented as a-z
			' Base64 digit value 62 is represented as $
			' Base64 digit value 63 is represented as @

			iTemp = iValue Mod 64
			If iTemp = 63 Then
				sNewString = "@"
			ElseIf iTemp = 62 Then
				sNewString = "$"
			ElseIf iTemp >= 36 Then
				sNewString = Chr(iTemp + 61)
			ElseIf iTemp >= 10 Then
				sNewString = Chr(iTemp + 55)
			Else
				sNewString = Chr(iTemp + 48)
			End If

			iTemp = (iValue - iTemp) / 64

			If iTemp = 63 Then
				sNewString = "@" & sNewString
			ElseIf iTemp = 62 Then
				sNewString = "$" & sNewString
			ElseIf iTemp >= 36 Then
				sNewString = Chr(iTemp + 61) & sNewString
			ElseIf iTemp >= 10 Then
				sNewString = Chr(iTemp + 55) & sNewString
			Else
				sNewString = Chr(iTemp + 48) & sNewString
			End If

			sCompactedString = sCompactedString & sNewString
		Loop

		' Append the number of characters to ignore, to the compacted string
		CompactString = sCompactedString & CStr((3 - (Len(psSourceString) Mod 3)) Mod 3)
	End Function

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

		Try

			sDecompactedString = ""
			iTrailingCharsToIgnore = CShort(Right(psSourceString, 1))
			sModifiedSourceString = Left(psSourceString, psSourceString.Length - 1)

			Do While Len(sModifiedSourceString) > 0
				sSubString = Left(sModifiedSourceString & "00", 2)
				sModifiedSourceString = Mid(sModifiedSourceString, 3)
				iTempTotal = 0

				iAscCode = Asc(Mid(sSubString, 2, 1))
				If iAscCode = 64 Then	' @
					iTemp = 63
				ElseIf iAscCode = 36 Then	' $
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
				If iAscCode = 64 Then	' @
					iTemp = 63
				ElseIf iAscCode = 36 Then	' $
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
			Return sDecompactedString

		Catch ex As Exception
			Throw

		End Try

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
