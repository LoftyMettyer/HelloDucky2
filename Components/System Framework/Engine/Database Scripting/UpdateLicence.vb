Option Strict On
Option Explicit On

Public Class UpdateLicence

	Private Property CustomerNo As Long
	Private Property DATUsers As Long
	Private Property DMIMUsers As Long
	Private Property DMISUsers As Long
	Private Property SSIUsers As Long
	Private Property Modules As Long
	Private Property Headcount As Long
	Private Property ExpiryDate As Date
	Private Property LicenceType As Long

	Public Sub SetOldLicenceKey(keyValue As String)

		Dim strCustNo As String
		Dim strDAT As String
		Dim strDMIM As String
		Dim strDMIS As String
		Dim strSSI As String
		Dim strModules As String
		Dim strRandomDigit As String

		Dim strInput As String

		Randomize()
		CustomerNo = 0

		If keyValue Like "A????-?????-?????-?????" Then

			strInput = Replace(keyValue, "-", "")
			strInput = Mid(strInput, 1, 1) & Mid(strInput, 6, 1) & _
								 Mid(strInput, 11, 1) & Mid(strInput, 16, 1) & _
								 Mid(strInput, 4, 1) & Mid(strInput, 9, 1) & _
								 Mid(strInput, 14, 1) & Mid(strInput, 19, 1) & _
								 Mid(strInput, 3, 1) & Mid(strInput, 8, 1) & _
								 Mid(strInput, 13, 1) & Mid(strInput, 18, 1) & _
								 Mid(strInput, 2, 1) & Mid(strInput, 7, 1) & _
								 Mid(strInput, 12, 1) & Mid(strInput, 17, 1) & _
								 Mid(strInput, 5, 1) & Mid(strInput, 10, 1) & _
								 Mid(strInput, 15, 1) & Mid(strInput, 20, 1)

			strCustNo = Mid(strInput, 2, 4)
			strDAT = Mid(strInput, 6, 2)
			strDMIM = Mid(strInput, 8, 2)
			strDMIS = Mid(strInput, 10, 2)
			strSSI = Mid(strInput, 12, 2)
			strModules = Mid(strInput, 14, 6)
			strRandomDigit = Mid(strInput, 20, 1)

			CustomerNo = ConvertBase32ToLong(strCustNo)
			DATUsers = ConvertBase32ToLong(strRandomDigit & strDAT)
			DMIMUsers = ConvertBase32ToLong(strRandomDigit & strDMIM)
			DMISUsers = ConvertBase32ToLong(strRandomDigit & strDMIS)
			SSIUsers = ConvertBase32ToLong(strRandomDigit & strSSI)
			Modules = ConvertBase32ToLong(strModules)

		End If

	End Sub

	Private Sub SetNewLicenceKey(keyValue As String)

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
		CustomerNo = 0

		If keyValue Like "??????-??????-??????-??????-??????" Then

			strInput = Replace(keyValue, "-", "")
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

			CustomerNo = ConvertBase32ToLong(strRandomDigit & strCustNo)
			DATUsers = ConvertBase32ToLong(strRandomDigit & strDAT)
			DMIMUsers = ConvertBase32ToLong(strRandomDigit & strDMIM)
			DMISUsers = ConvertBase32ToLong(strRandomDigit & strDMIS)
			SSIUsers = ConvertBase32ToLong(strRandomDigit & strSSI)
			Modules = ConvertBase32ToLong(strModules)
			Headcount = ConvertBase32ToLong(strRandomDigit & strHeadcount)
			ExpiryDate = CDate(DateFromJulian(CStr(ConvertBase32ToLong(strRandomDigit & strExpiryDate))))
			LicenceType = ConvertBase32ToLong(strRandomDigit & sLicenceType)

		End If

	End Sub

	Public ReadOnly Property GenerateNewKey() As String
		Get

			Dim sReturnKey As String = ""

			Try

				Dim lngCount As Integer
				Dim strOutput As String
				Dim lngRandomDigit As Integer

				Randomize()
				lngRandomDigit = CInt(Math.Ceiling(Rnd() * 26))

				Dim sLicenceType = ConvertLongToBase32String(1, LicenceType, lngRandomDigit)
				Dim strCustNo = ConvertLongToBase32String(3, CustomerNo, lngRandomDigit)
				Dim strDAT = ConvertLongToBase32String(2, DATUsers, lngRandomDigit)
				Dim strDMIM = ConvertLongToBase32String(2, DMIMUsers, lngRandomDigit)
				Dim strDMIS = ConvertLongToBase32String(2, DMISUsers, lngRandomDigit)
				Dim strSSI = ConvertLongToBase32String(4, SSIUsers, lngRandomDigit)
				Dim strModules = ConvertLongToBase32String(6, Modules, 0)	' Allows upto 24 modules
				Dim strHeadcount = ConvertLongToBase32String(4, Headcount, lngRandomDigit)
				Dim strExpiryDate = ConvertLongToBase32String(5, DateToJulian(ExpiryDate), lngRandomDigit)

				strOutput = sLicenceType & strCustNo & strDAT & strDMIM & strDMIS & strSSI & strModules & Chr(lngRandomDigit + 64) & strHeadcount & strExpiryDate

				'Jumble it up!
				For lngCount = 1 To 5
					sReturnKey &= IIf(sReturnKey <> vbNullString, "-", "").ToString() & Mid(strOutput, lngCount, 1) & Mid(strOutput, lngCount + 15, 1) & Mid(strOutput, lngCount + 10, 1) & Mid(strOutput, lngCount + 5, 1) & Mid(strOutput, lngCount + 25, 1) & Mid(strOutput, lngCount + 20, 1)
				Next

			Catch ex As Exception
				Return ""

			End Try

			Return sReturnKey
		End Get
	End Property

	Private Function DateToJulian(dDate As DateTime) As Long

		Dim returnVal As Long

		If Year(dDate) < 1999 Then
			returnVal = 0
		Else
			returnVal = CLng(String.Format("{0}{1}", dDate.Year.ToString("0000"), dDate.DayOfYear().ToString("000")))
		End If

		Return returnVal

	End Function

	Private Function DateFromJulian(sDate As String) As Date

		Dim dResult As Date
		Dim sYear As String
		sYear = Mid(sDate, 1, 4)

		dResult = DateSerial(CInt(sYear), 1, 1)
		DateFromJulian = DateTime.FromOADate(dResult.ToOADate + CDate(Mid(sDate, 5, 3)).ToOADate - 1)

	End Function

	Private Function ConvertBase32ToLong(strInput As String) As Long

		Dim lngRandomDigit As Integer
		Dim strAlphaString As String
		Dim lngOutput As Integer
		Dim lngFactor As Double
		Dim lngCount As Integer

		Try

			lngRandomDigit = Asc(Mid(strInput, 1, 1)) - 64
			strAlphaString = GenerateAlphaString(lngRandomDigit)

			lngOutput = (InStr(strAlphaString, Mid(strInput, Len(strInput), 1)) - 1)

			lngFactor = 32
			For lngCount = Len(strInput) - 1 To 2 Step -1
				lngOutput = CInt(lngOutput + ((InStr(strAlphaString, Mid(strInput, lngCount, 1)) - 1) * lngFactor))
				lngFactor = lngFactor * 32
			Next


		Catch ex As Exception
			Return 0
		End Try

		Return lngOutput

	End Function

	Private Function GenerateAlphaString(lngGap As Integer) As String

		Dim strOutput As String
		Dim lngCount As Integer
		Dim lngLoop As Integer

		strOutput = vbNullString
		For lngLoop = 0 To lngGap - 1

			For lngCount = Asc("A") + lngLoop To Asc("Z") Step lngGap
				If InStr("IOQ", Chr(lngCount)) = 0 Then
					strOutput = Chr(lngCount) & strOutput
				End If
			Next

			For lngCount = Asc("1") + lngLoop To Asc("9") Step lngGap
				strOutput = Chr(lngCount) & strOutput
			Next

		Next

		Return strOutput

	End Function

	Private Function ConvertLongToBase32String(iSize As Integer, lngInput As Long, iDigit As Integer) As String

		Dim lngRandomDigit As Integer
		Dim strAlphaString As String
		Dim lngFactor As Double
		Dim lngCount As Integer
		Dim returnValue As String

		lngRandomDigit = CInt(IIf(iDigit = 0, Int(Rnd() * 26) + 1, iDigit))
		strAlphaString = GenerateAlphaString(lngRandomDigit)

		returnValue = Mid(strAlphaString, CShort(lngInput And 31) + 1, 1)

		lngFactor = 32
		For lngCount = 2 To iSize - DirectCast(IIf(iDigit = 0, 1, 0), Integer)
			returnValue = Mid(strAlphaString, CShort((lngInput \ CLng(lngFactor)) And 31) + 1, 1) & returnValue
			lngFactor = lngFactor * 32
		Next

		If iDigit = 0 Then
			returnValue = Chr(lngRandomDigit + 64) & returnValue
		End If

		Return returnValue

	End Function




End Class
