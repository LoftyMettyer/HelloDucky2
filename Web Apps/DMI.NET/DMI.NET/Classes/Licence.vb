Option Strict On
Option Explicit On

Namespace Classes
	Public Class Licence

		Public Shared Type As LicenceType
		Public Shared CustomerNumber As Long

		Public Shared DATUsers As Integer
		Public Shared SSIUsers As Long
		Public Shared DMIUsers As Integer
		Public Shared Headcount As Long
		Public Shared ExpiryDate As DateTime
		Public Shared Modules As Long

		Public Shared IsValid As Boolean

		Public Sub Populate(licenceKey As String)

			Dim strCustNo As String
			Dim strDMIM As String
			Dim strDAT As String
			Dim strDMIS As String
			Dim strModules As String
			Dim strRandomDigit As String
			Dim strSSIHeadcount As String
			Dim strExpiryDate As String
			Dim sLicenceType As String

			Dim strInput As String
			Dim strCheckSum As String
			Dim lngHeadcountSSI As Long

			Try

				Randomize()

				If licenceKey Like "??????-??????-??????-??????-??????-??????" Then

					strInput = Replace(licenceKey, "-", "")
					strInput = Mid(strInput, 32, 1) & Mid(strInput, 35, 1) & Mid(strInput, 6, 1) & Mid(strInput, 27, 1) & Mid(strInput, 21, 1) & Mid(strInput, 29, 1) & _
						 Mid(strInput, 1, 1) & Mid(strInput, 26, 1) & Mid(strInput, 4, 1) & Mid(strInput, 28, 1) & Mid(strInput, 2, 1) & Mid(strInput, 8, 1) & _
						 Mid(strInput, 3, 1) & Mid(strInput, 13, 1) & Mid(strInput, 31, 1) & Mid(strInput, 18, 1) & Mid(strInput, 33, 1) & Mid(strInput, 10, 1) & _
						 Mid(strInput, 20, 1) & Mid(strInput, 14, 1) & Mid(strInput, 23, 1) & Mid(strInput, 9, 1) & Mid(strInput, 25, 1) & Mid(strInput, 16, 1) & _
						 Mid(strInput, 36, 1) & Mid(strInput, 7, 1) & Mid(strInput, 22, 1) & Mid(strInput, 17, 1) & Mid(strInput, 34, 1) & Mid(strInput, 19, 1) & _
						 Mid(strInput, 30, 1) & Mid(strInput, 24, 1) & Mid(strInput, 15, 1) & Mid(strInput, 5, 1) & Mid(strInput, 12, 1) & Mid(strInput, 11, 1)

					sLicenceType = Mid(strInput, 1, 1)
					strCustNo = Mid(strInput, 2, 3)
					strDAT = Mid(strInput, 5, 2)
					strDMIM = Mid(strInput, 7, 2)
					strCheckSum = Mid(strInput, 9, 6)
					strDMIS = Mid(strInput, 15, 2)
					strModules = Mid(strInput, 17, 7)
					strRandomDigit = Mid(strInput, 24, 1)
					strSSIHeadcount = Mid(strInput, 25, 4)
					strExpiryDate = Mid(strInput, 29, 5)

					CustomerNumber = ConvertBase32ToLong(strRandomDigit & strCustNo)
					DATUsers = CInt(ConvertBase32ToLong(strRandomDigit & strDAT))
					DMIUsers = CInt(ConvertBase32ToLong(strRandomDigit & strDMIM))
					Modules = ConvertBase32ToLong(strModules)
					lngHeadcountSSI = ConvertBase32ToLong(strRandomDigit & strSSIHeadcount)
					Type = CType(ConvertBase32ToLong(strRandomDigit & sLicenceType), LicenceType)

					If Type = LicenceType.Concurrency Then
						SSIUsers = lngHeadcountSSI
					Else
						Headcount = lngHeadcountSSI
					End If

					Dim lngDate = ConvertBase32ToLong(strRandomDigit & strExpiryDate)
					If lngDate = 0 Then
						ExpiryDate = Date.MaxValue
					Else
						ExpiryDate = Date.FromOADate(lngDate)
					End If

					IsValid = (ConvertBase32ToLong(strRandomDigit & strCheckSum) = CustomerNumber + Type + DATUsers + DMIUsers + lngHeadcountSSI + Modules + lngDate)

				End If

			Catch ex As Exception
				IsValid = False

			End Try

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

		Public Shared Function IsModuleLicenced(softwareModule As SoftwareModule) As Boolean
			Return CBool(Modules And softwareModule)
		End Function

	End Class
End Namespace