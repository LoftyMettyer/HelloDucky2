Option Strict Off
Option Explicit On

Imports System.Globalization

Module modIntranet

	Friend Function ConvertNumberForSQL(strInput As String) As String
		'Get a number in the correct format for a SQL string
		'(e.g. on french systems replace decimal comma for a decimal point)
		Return Replace(strInput, CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".")
	End Function

	Friend Function ConvertNumberForDisplay(strInput As String) As String
		'Get a number in the correct format for display
		'(e.g. on french systems replace decimal point for a decimal comma)
		Return Replace(strInput, ".", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
	End Function



	Friend Function DecToBin(DeciValue As Integer, Optional NoOfBits As Short = 8) As String

		Dim i As Short 'make sure there are enough bits to contain the number
		Do While DeciValue > (2 ^ NoOfBits) - 1
			NoOfBits = NoOfBits + 8
		Loop
		DecToBin = vbNullString
		'build the string
		For i = 0 To (NoOfBits - 1)
			DecToBin = CStr(CShort(DeciValue And 2 ^ i) / 2 ^ i) & DecToBin
		Next i
	End Function


End Module