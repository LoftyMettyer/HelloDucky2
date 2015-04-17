Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class clsLicence
	
	Private mlngCustomerNo As Integer
	Private mlngDAT As Integer
	Private mlngDMIM As Integer
	Private mlngModules As Integer
	Private mlngHeadcountSSI As Integer
	Private mdExpiryDate As Date
	Private mlngLicenceType As Integer
	
	Public IsValid As Boolean
	Public ValidateCreationDate As Boolean
	
	Public ReadOnly Property CustomerNo() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object CustomerNo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CustomerNo = mlngCustomerNo
		End Get
	End Property
	
	Public ReadOnly Property DATUsers() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object DATUsers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DATUsers = mlngDAT
		End Get
	End Property
	
	Public ReadOnly Property DMIMUsers() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object DMIMUsers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DMIMUsers = mlngDMIM
		End Get
	End Property
	
	Public ReadOnly Property SSIUsers() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object SSIUsers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SSIUsers = IIf(mlngLicenceType = 0, mlngHeadcountSSI, 0)
		End Get
	End Property
	
	Public ReadOnly Property Modules() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object Modules. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Modules = mlngModules
		End Get
	End Property
	
	Public ReadOnly Property Headcount() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object Headcount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Headcount = IIf(mlngLicenceType > 0, mlngHeadcountSSI, 0)
		End Get
	End Property
	
	Public ReadOnly Property ExpiryDate() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object ExpiryDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ExpiryDate = mdExpiryDate
		End Get
	End Property
	
	Public ReadOnly Property LicenceType() As Object
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object LicenceType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LicenceType = mlngLicenceType
		End Get
	End Property
	
	Public WriteOnly Property LicenceKey() As String
		Set(ByVal Value As String)
			
			Dim strCustNo As String
			Dim strDAT As String
			Dim strDMIM As String
			Dim strDMIS As String
			Dim strModules As String
			Dim strRandomDigit As String
			Dim strSSIHeadcount As String
			Dim strExpiryDate As String
			Dim sLicenceType As String
			
			Dim strTemp As String
			Dim strInput As String
			Dim lngCount As Integer
			Dim lngDateValue As Integer
			Dim strCheckSum As String
			Dim lngRandomDigit As Integer
			Dim strGeneratedDay As String
			
			'1st Generation: AAAAAAAAAAAAAAA
			'2nd Generation: 1231-2312-3123-1233
			'3rd Generation: 1234-5123-4512-3451-2345
			'4th Generation: A5316-16426-16426-16536
			'5th Generation: 45FG63-1S7G6D-1V4R63-DDG533-PG6GX
			
			Randomize(VB.Timer())
			mlngCustomerNo = 0
			
			If Value Like "??????-??????-??????-??????-??????-??????" Then
				
				strInput = Replace(Value, "-", "")
				strInput = Mid(strInput, 32, 1) & Mid(strInput, 35, 1) & Mid(strInput, 6, 1) & Mid(strInput, 27, 1) & Mid(strInput, 21, 1) & Mid(strInput, 29, 1) & Mid(strInput, 1, 1) & Mid(strInput, 26, 1) & Mid(strInput, 4, 1) & Mid(strInput, 28, 1) & Mid(strInput, 2, 1) & Mid(strInput, 8, 1) & Mid(strInput, 3, 1) & Mid(strInput, 13, 1) & Mid(strInput, 31, 1) & Mid(strInput, 18, 1) & Mid(strInput, 33, 1) & Mid(strInput, 10, 1) & Mid(strInput, 20, 1) & Mid(strInput, 14, 1) & Mid(strInput, 23, 1) & Mid(strInput, 9, 1) & Mid(strInput, 25, 1) & Mid(strInput, 16, 1) & Mid(strInput, 36, 1) & Mid(strInput, 7, 1) & Mid(strInput, 22, 1) & Mid(strInput, 17, 1) & Mid(strInput, 34, 1) & Mid(strInput, 19, 1) & Mid(strInput, 30, 1) & Mid(strInput, 24, 1) & Mid(strInput, 15, 1) & Mid(strInput, 5, 1) & Mid(strInput, 12, 1) & Mid(strInput, 11, 1)
				
				sLicenceType = Mid(strInput, 1, 1)
				strCustNo = Mid(strInput, 2, 3)
				strDAT = Mid(strInput, 5, 2)
				strDMIM = Mid(strInput, 7, 2)
				strCheckSum = Mid(strInput, 9, 6)
				strModules = Mid(strInput, 17, 7)
				strRandomDigit = Mid(strInput, 24, 1)
				strSSIHeadcount = Mid(strInput, 25, 4)
				strExpiryDate = Mid(strInput, 29, 5)
				strGeneratedDay = Mid(strInput, 34, 2)
				
				mlngLicenceType = ConvertStringToNumber2(strRandomDigit & sLicenceType)
				mlngCustomerNo = ConvertStringToNumber2(strRandomDigit & strCustNo)
				mlngDAT = ConvertStringToNumber2(strRandomDigit & strDAT)
				mlngDMIM = ConvertStringToNumber2(strRandomDigit & strDMIM)
				mlngHeadcountSSI = ConvertStringToNumber2(strRandomDigit & strSSIHeadcount)
				mlngModules = ConvertStringToNumber2(strModules)
				lngRandomDigit = Asc(strRandomDigit) - 64
				
				lngDateValue = ConvertStringToNumber2(strRandomDigit & strExpiryDate)
				If lngDateValue > 0 Then
					mdExpiryDate = System.Date.FromOADate(lngDateValue)
				End If
				
				' Validate checksum
				If ValidateCreationDate Then
					IsValid = (ConvertStringToNumber2(strRandomDigit & strGeneratedDay) = DayNumber(Now))
				Else
					IsValid = True
				End If
				
				IsValid = IsValid And (ConvertNumberToString2(6, mlngCustomerNo + mlngLicenceType + mlngDAT + mlngDMIM + mlngHeadcountSSI + mlngModules + lngDateValue, lngRandomDigit) = strCheckSum)
				
			End If
			
		End Set
	End Property
	
	Private Function DateFromJulian(ByVal sDate As String) As Date
		
		Dim dResult As Date
		Dim sYear As String
		sYear = Mid(sDate, 1, 4)
		
		dResult = DateSerial(CInt(sYear), 1, 1)
		DateFromJulian = System.Date.FromOADate(dResult.ToOADate + CDate(Mid(sDate, 5, 3)).ToOADate - 1)
		
	End Function
	
	Public Function ConvertStringToNumber2(ByRef strInput As String) As Integer
		
		Dim lngRandomDigit As Integer
		Dim strAlphaString As String
		Dim lngOutput As Integer
		Dim lngFactor As Double
		Dim lngCount As Integer
		
		On Error GoTo exitf
		
		lngRandomDigit = Asc(Mid(strInput, 1, 1)) - 64
		strAlphaString = GenerateAlphaString(lngRandomDigit)
		
		lngOutput = (InStr(strAlphaString, Mid(strInput, Len(strInput), 1)) - 1)
		
		lngFactor = 32
		For lngCount = Len(strInput) - 1 To 2 Step -1
			lngOutput = lngOutput + ((InStr(strAlphaString, Mid(strInput, lngCount, 1)) - 1) * lngFactor)
			lngFactor = lngFactor * 32
		Next 
		
		ConvertStringToNumber2 = lngOutput
		
exitf: 
		
	End Function
	
	Private Function GenerateAlphaString(ByRef lngGap As Integer) As String
		
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
		
		GenerateAlphaString = strOutput
		
	End Function
	
	Private Function ConvertNumberToString2(ByRef lngSize As Integer, ByRef lngInput As Integer, ByRef lngDigit As Integer) As String
		
		Dim lngRandomDigit As Integer
		Dim strAlphaString As String
		Dim lngFactor As Double
		Dim lngCount As Integer
		
		lngRandomDigit = IIf(lngDigit = 0, Int(Rnd() * 26) + 1, lngDigit)
		strAlphaString = GenerateAlphaString(lngRandomDigit)
		
		ConvertNumberToString2 = Mid(strAlphaString, CShort(lngInput And 31) + 1, 1)
		
		lngFactor = 32
		For lngCount = 2 To lngSize - IIf(lngDigit = 0, 1, 0)
			ConvertNumberToString2 = Mid(strAlphaString, CShort((lngInput \ lngFactor) And 31) + 1, 1) & ConvertNumberToString2
			lngFactor = lngFactor * 32
		Next 
		
		If lngDigit = 0 Then
			ConvertNumberToString2 = Chr(lngRandomDigit + 64) & ConvertNumberToString2
		End If
		
	End Function
	
	Private Function DayNumber(ByVal vDate As Date) As Integer
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		DayNumber = CInt(VB6.Format(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate("01/01/" & VB6.Format(Year(vDate), "0000")), vDate) + 1, "000"))
	End Function
End Class