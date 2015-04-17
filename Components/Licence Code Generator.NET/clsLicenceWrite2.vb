Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class clsLicenceWrite2
	
	Private mlngCustomerNo As Integer
	Private mlngDAT As Integer
	Private mlngDMIM As Integer
	Private mlngSSI As Integer
	Private mlngModules As Integer
	
	Private mlngHeadcount As Integer
	Private mlngExpiryDate As Integer
	Private mlngLicenceType As Integer
	
	Public WriteOnly Property CustomerNo() As Integer
		Set(ByVal Value As Integer)
			mlngCustomerNo = Value
		End Set
	End Property
	
	Public WriteOnly Property DATUsers() As Integer
		Set(ByVal Value As Integer)
			mlngDAT = Value
		End Set
	End Property
	
	Public WriteOnly Property DMIMUsers() As Integer
		Set(ByVal Value As Integer)
			mlngDMIM = Value
		End Set
	End Property
	
	Public WriteOnly Property SSIUsers() As Integer
		Set(ByVal Value As Integer)
			mlngSSI = Value
		End Set
	End Property
	
	Public WriteOnly Property Modules() As Integer
		Set(ByVal Value As Integer)
			mlngModules = Value
		End Set
	End Property
	
	Public WriteOnly Property Headcount() As Integer
		Set(ByVal Value As Integer)
			mlngHeadcount = Value
		End Set
	End Property
	
	Public WriteOnly Property ExpiryDate() As Integer
		Set(ByVal Value As Integer)
			mlngExpiryDate = Value
		End Set
	End Property
	
	Public WriteOnly Property LicenceType() As Integer
		Set(ByVal Value As Integer)
			mlngLicenceType = Value
		End Set
	End Property
	
	Public ReadOnly Property LicenceKey2() As Object
		Get
			
			Dim strCustNo As String
			Dim strDAT As String
			Dim strDMIM As String
			Dim strFiller1 As String
			Dim strModules As String
			Dim strSSIHeadcount As String
			Dim strExpiryDate As String
			Dim lngCount As Integer
			Dim strOutput As String
			Dim lngRandomDigit As Integer
			Dim sLicenceType As String
			Dim strCheckSum As String
			Dim strGeneratedDay As String
			Dim lngSSIHeadcount As Integer
			
			'UPGRADE_WARNING: Couldn't resolve default property of object LicenceKey2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LicenceKey2 = vbNullString
			If Valid Then
				Randomize(VB.Timer())
				lngRandomDigit = Int(Rnd() * 26) + 1
				
				lngSSIHeadcount = IIf(mlngLicenceType = 0, mlngSSI, mlngHeadcount)
				
				sLicenceType = ConvertNumberToString2(1, mlngLicenceType, lngRandomDigit)
				strCustNo = ConvertNumberToString2(3, mlngCustomerNo, lngRandomDigit)
				strDAT = ConvertNumberToString2(2, mlngDAT, lngRandomDigit)
				strDMIM = ConvertNumberToString2(2, mlngDMIM, lngRandomDigit)
				strFiller1 = ConvertNumberToString2(2, 1024, lngRandomDigit)
				strSSIHeadcount = ConvertNumberToString2(4, lngSSIHeadcount, lngRandomDigit)
				strModules = ConvertNumberToString2(7, mlngModules, 0) ' Allows upto 35 modules
				strExpiryDate = ConvertNumberToString2(5, mlngExpiryDate, lngRandomDigit)
				
				strGeneratedDay = ConvertNumberToString2(2, DayNumber(Now), lngRandomDigit)
				strCheckSum = ConvertNumberToString2(6, mlngCustomerNo + mlngLicenceType + mlngDAT + mlngDMIM + lngSSIHeadcount + mlngModules + mlngExpiryDate, lngRandomDigit)
				
				strOutput = sLicenceType & strCustNo & strDAT & strDMIM & strCheckSum & strFiller1 & strModules & Chr(lngRandomDigit + 64) & strSSIHeadcount & strExpiryDate & strGeneratedDay & "X"
				
				'Jumble it up!
				'UPGRADE_WARNING: Couldn't resolve default property of object LicenceKey2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				LicenceKey2 = Mid(strOutput, 7, 1) & Mid(strOutput, 11, 1) & Mid(strOutput, 13, 1) & Mid(strOutput, 9, 1) & Mid(strOutput, 34, 1) & Mid(strOutput, 3, 1) & "-" & Mid(strOutput, 26, 1) & Mid(strOutput, 12, 1) & Mid(strOutput, 22, 1) & Mid(strOutput, 18, 1) & Mid(strOutput, 36, 1) & Mid(strOutput, 35, 1) & "-" & Mid(strOutput, 14, 1) & Mid(strOutput, 20, 1) & Mid(strOutput, 33, 1) & Mid(strOutput, 24, 1) & Mid(strOutput, 28, 1) & Mid(strOutput, 16, 1) & "-" & Mid(strOutput, 30, 1) & Mid(strOutput, 19, 1) & Mid(strOutput, 5, 1) & Mid(strOutput, 27, 1) & Mid(strOutput, 21, 1) & Mid(strOutput, 32, 1) & "-" & Mid(strOutput, 23, 1) & Mid(strOutput, 8, 1) & Mid(strOutput, 4, 1) & Mid(strOutput, 10, 1) & Mid(strOutput, 6, 1) & Mid(strOutput, 31, 1) & "-" & Mid(strOutput, 15, 1) & Mid(strOutput, 1, 1) & Mid(strOutput, 17, 1) & Mid(strOutput, 29, 1) & Mid(strOutput, 2, 1) & Mid(strOutput, 25, 1)
				
			End If
			
		End Get
	End Property
	
	Private Function Valid() As Boolean
		Valid = (mlngCustomerNo > 0 And mlngModules > 0)
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
	
	Private Function PadString(ByVal length As Short, ByVal value As String, ByVal padWith As String) As Object
		PadString = Right(New String(padWith, length) & value, length)
	End Function
	
	Private Function Date2Julian(ByVal vDate As Date) As Integer
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		Date2Julian = CInt(VB6.Format(Year(vDate), "0000") & VB6.Format(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate("01/01/" & VB6.Format(Year(vDate), "0000")), vDate) + 1, "000"))
		
	End Function
	
	Private Function DayNumber(ByVal vDate As Date) As Integer
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		DayNumber = CInt(VB6.Format(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate("01/01/" & VB6.Format(Year(vDate), "0000")), vDate) + 1, "000"))
		
	End Function
End Class