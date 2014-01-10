Public Module GeneralUtilities
	Public Function IsDataColumnNumeric(col As DataColumn) As Boolean
		If col Is Nothing Then
			Return False
		End If

		Dim numericTypes As New ArrayList

		With numericTypes
			.Add(GetType([Byte]))
			.Add(GetType([Decimal]))
			.Add(GetType([Double]))
			.Add(GetType(Int16))
			.Add(GetType(Int32))
			.Add(GetType(Int64))
			.Add(GetType([SByte]))
			.Add(GetType([Single]))
			.Add(GetType(UInt16))
			.Add(GetType(UInt32))
			.Add(GetType(UInt64))
		End With

		Return numericTypes.Contains(col.DataType)

	End Function
End Module
