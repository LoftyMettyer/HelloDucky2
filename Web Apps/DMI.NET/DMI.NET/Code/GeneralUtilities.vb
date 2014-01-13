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
End Module
