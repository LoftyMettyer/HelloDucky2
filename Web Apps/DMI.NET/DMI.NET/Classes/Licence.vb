Option Strict On
Option Explicit On

Imports System.Diagnostics.Eventing.Reader

Namespace Classes
	Public Class Licence

		Public Shared Type As LicenceType
		Public Shared CustomerNumber As Long

		Public Shared SSIUsers As Long
		Public Shared DMIUsers As Integer
		Public Shared DMISingleUsers As Integer
		Public Shared Headcount As Long
		Public Shared P14Headcount As Integer

		Public Shared ExpiryDate As DateTime

		Public Shared Modules As Long

		Public Sub Populate(licenceKey As String)

			Dim strCustNo As String
			Dim strDAT As String
			Dim strDMIM As String
			Dim strDMIS As String
			Dim strSSI As String
			Dim strModules As String
			Dim strRandomDigit As String
			Dim strHeadcount As String
			Dim strExpiryDate As String
			Dim sLicenceType As String

			Dim strInput As String

			Randomize()

			If licenceKey Like "??????-??????-??????-??????-??????" Then

				strInput = Replace(licenceKey, "-", "")
				strInput = strInput.Substring(0, 1) & strInput.Substring(6, 1) & strInput.Substring(12, 1) & strInput.Substring(18, 1) & strInput.Substring(24, 1) & _
						strInput.Substring(3, 1) & strInput.Substring(9, 1) & strInput.Substring(15, 1) & strInput.Substring(11, 1) & strInput.Substring(27, 1) & _
						strInput.Substring(2, 1) & strInput.Substring(8, 1) & strInput.Substring(14, 1) & strInput.Substring(20, 1) & strInput.Substring(26, 1) & _
						strInput.Substring(1, 1) & strInput.Substring(7, 1) & strInput.Substring(13, 1) & strInput.Substring(19, 1) & strInput.Substring(25, 1) & _
						strInput.Substring(5, 1) & strInput.Substring(11, 1) & strInput.Substring(17, 1) & strInput.Substring(23, 1) & strInput.Substring(29, 1) & _
						strInput.Substring(4, 1) & strInput.Substring(10, 1) & strInput.Substring(16, 1) & strInput.Substring(22, 1) & strInput.Substring(28, 1)

				sLicenceType = Mid(strInput, 1, 1)
				strCustNo = Mid(strInput, 2, 3)
				strDAT = Mid(strInput, 5, 2)
				strDMIM = Mid(strInput, 7, 2)
				strDMIS = Mid(strInput, 9, 2)
				strSSI = Mid(strInput, 11, 4)
				strModules = Mid(strInput, 15, 6)
				strRandomDigit = Mid(strInput, 21, 1)
				strHeadcount = Mid(strInput, 22, 4)
				strExpiryDate = Mid(strInput, 26, 5)

				CustomerNumber = ConvertBase32ToLong(strRandomDigit & strCustNo)
				'DATUsers = ConvertBase32ToLong(strRandomDigit & strDAT)
				DMIUsers = CInt(ConvertBase32ToLong(strRandomDigit & strDMIM))
				DMISingleUsers = CInt(ConvertBase32ToLong(strRandomDigit & strDMIS))
				SSIUsers = ConvertBase32ToLong(strRandomDigit & strSSI)
				Modules = ConvertBase32ToLong(strModules)
				Headcount = ConvertBase32ToLong(strRandomDigit & strHeadcount)
				Type = CType(ConvertBase32ToLong(strRandomDigit & sLicenceType), LicenceType)

				Dim lngDate = ConvertBase32ToLong(strRandomDigit & strExpiryDate)
				If lngDate > 0 Then
					ExpiryDate = CDate(DateFromJulian(lngDate.ToString()))
				End If

			End If
		End Sub

		Private Function GenerateAlphaString(lngGap As Long) As String

			Dim strOutput As String
			Dim lngCount As Long
			Dim lngLoop As Long

			strOutput = vbNullString
			For lngLoop = 0 To lngGap - 1

				For lngCount = Asc("A") + lngLoop To Asc("Z") Step lngGap
					If InStr("IOQ", Chr(CInt(lngCount))) = 0 Then
						strOutput = Chr(CInt(lngCount)) & strOutput
					End If
				Next

				For lngCount = Asc("1") + lngLoop To Asc("9") Step lngGap
					strOutput = Chr(CInt(lngCount)) & strOutput
				Next

			Next

			Return strOutput

		End Function

		Private Function ConvertBase32ToLong(strInput As String) As Long

			Dim lngRandomDigit As Long
			Dim strAlphaString As String
			Dim lngOutput As Long
			Dim lngFactor As Double
			Dim lngCount As Long

			Try

				lngRandomDigit = Asc(Mid(strInput, 1, 1)) - 64
				strAlphaString = GenerateAlphaString(lngRandomDigit)

				lngOutput = (InStr(strAlphaString, Mid(strInput, Len(strInput), 1)) - 1)

				lngFactor = 32
				For lngCount = Len(strInput) - 1 To 2 Step -1
					lngOutput = CLng(lngOutput + ((InStr(strAlphaString, Mid(strInput, CInt(lngCount), 1)) - 1) * lngFactor))
					lngFactor = lngFactor * 32
				Next

				Return lngOutput

			Catch ex As Exception
				Return 0

			End Try

		End Function

		Private Shared Function DateFromJulian(sDate As String) As Date

			Dim sYear = CInt(sDate.Substring(0, 4))
			Dim sDayNo = CInt(sDate.Substring(4, 3))

			If sYear > 1900 Then
				Return DateSerial(sYear, 1, 1).AddDays(sDayNo)
			Else
				Return DateSerial(9999, 12, 31)
			End If

		End Function

	End Class
End Namespace