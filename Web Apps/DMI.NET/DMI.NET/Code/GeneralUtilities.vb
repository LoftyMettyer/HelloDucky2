﻿Imports System.Data.SqlClient

Public Module GeneralUtilities
	Public Function IsDataColumnDecimal(col As DataColumn) As Boolean
		If col Is Nothing Then
			Return False
		End If

		Dim numericTypes As New ArrayList

		With numericTypes
			.Add(GetType([Decimal]))
			.Add(GetType([Double]))
			.Add(GetType([Single]))
		End With

		Return numericTypes.Contains(col.DataType)

	End Function

	' Returns a simplified description of the error (SQL message contains a whole lot more unnecessary gumpff
	Public Function GetPasswordChangeFailReason(ex As SqlException) As String

		Select Case ex.Number

			Case 18456
				Return "Old password incorrect."
			Case 18463
				Return "The password does not meet policy requirements because it has been used in the recent past."
			Case 18464
				Return "The password does not meet policy requirements because it is too short."
			Case 18465
				Return "The password does not meet policy requirements because it is too long."
			Case 18466
				Return "The password does not meet policy requirements because it is not complex enough."
			Case 18467
				Return "The password does not meet the requirements of the password filter DLL."
			Case Else
				Return ex.Message

		End Select

	End Function

	'Checks if a string is a valid email address
	Public Function IsValidEmailAddress(EmailAddress As String) As Boolean
		Return Regex.IsMatch(EmailAddress, "^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$", RegexOptions.IgnoreCase)
	End Function

	'Look at https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Values,_variables,_and_literals#Literals
	'\xXX The character with the Latin-1 encoding specified by the two hexadecimal digits XX between 00 and FF
	Public Function EncodeStringToJavascriptSpecialCharacters(s As String) As String
	Dim retVal As String = ""

	For i = 0 To s.Length - 1
		retVal = String.Concat(retVal, "\x", AscW(s.Chars(i)).ToString("X"))
	Next

	Return retVal
	End Function
End Module
