Imports System.Data.SqlClient

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

End Module
