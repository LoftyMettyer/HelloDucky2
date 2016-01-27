
Namespace Classes

	Public Class JqGridColModel
		' Pass in a datatable and a jqGrid colmodel pops out the other end.
		' USAGE:	Dim colModel As List(Of Object) = jqGridColModel.CreateColModel(rstLookup, sThousandColumns, sBlankIfZeroColumns)
		'					Return "{""total"":1,""page"":1,""records"":" & rstLookup.Rows.Count & ",""rows"":" JsonConvert.SerializeObject(rstLookup) & ", ""colmodel"":" & JsonConvert.SerializeObject(colModel) & "}"
		'	Javascript: colModel: jsondata.colmodel,
		'							datatype: 'local',
		'							data: jsondata.coldata

		Public Shared Function CreateColModel(dataTable As DataTable, Optional thousandSeparators As String = "", Optional blankIfZeroColumns As String = "") As List(Of Object)

			Dim arrCellProps(dataTable.Columns.Count)

			If dataTable.Rows.Count > 0 Then
				' read the first row to deduce certain unknown properties
				For iloop = 0 To (dataTable.Columns.Count - 1)

					If dataTable.Columns(iloop).DataType = GetType(Decimal) Then

						Dim cellValue = dataTable.Rows(0)(iloop)
						If Not IsDBNull(cellValue) Then

							Dim numberAsString As String = cellValue.ToString()
							Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(LocaleDecimalSeparator(), StringComparison.Ordinal)
							Dim numberOfDecimals As Integer = 0
							If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length

							arrCellProps(iloop) = numberOfDecimals

						End If
					End If

				Next
			End If

			Dim arrThousandSeparators = thousandSeparators.ToCharArray()
			Dim arrBlankIfZeroColumns = blankIfZeroColumns.ToCharArray()
			Dim colCounter As Integer = -1

			Dim colModel = New List(Of Object)()

			For Each col As DataColumn In dataTable.Columns
				If Not (col.ColumnName = "ID" Or String.Concat(col.ColumnName, "xxx").Substring(0, 3) = "ID_") Then
					colCounter += 1
				End If

				Select Case col.DataType

					Case GetType(DateTime)

						Dim localeDateFormat As String = HttpContext.Current.Session("LocaleDateFormat").replace("dd", "d").replace("MM", "m").replace("M", "m").replace("yyyy", "Y")

						colModel.Add(New With {
							Key .name = col.ColumnName,
							Key .index = col.ColumnName,
							Key .sortable = True,
							Key .hidden = col.ColumnName = "ID" Or String.Concat(col.ColumnName, "xxx").Substring(0, 3) = "ID_",
							Key .label = col.ColumnName.Replace("_", " "),
							Key .formatter = "date",
							Key .sorttype = "date",
							Key .formatoptions = New formatoptions() With {.srcformat = "ISO8601Long", .newformat = localeDateFormat}
						})

					Case GetType(Boolean)
						colModel.Add(New With {
							Key .name = col.ColumnName,
							Key .index = col.ColumnName,
							Key .sortable = True,
							Key .hidden = col.ColumnName = "ID" Or String.Concat(col.ColumnName, "xxx").Substring(0, 3) = "ID_",
							Key .label = col.ColumnName.Replace("_", " "),
							Key .formatter = "checkbox",
							Key .align = "center"
						})

					Case GetType(Int32)
						Dim sThousandSeparator As String = ""
						Try
							If arrThousandSeparators(colCounter) = "1" Then sThousandSeparator = LocaleThousandSeparator()
						Catch ex As Exception
						End Try

						colModel.Add(New With {
							Key .name = col.ColumnName,
							Key .index = col.ColumnName,
							Key .sortable = True,
							Key .hidden = col.ColumnName = "ID" Or String.Concat(col.ColumnName, "xxx").Substring(0, 3) = "ID_",
							Key .label = col.ColumnName.Replace("_", " "),
							Key .formatter = "integer",
							Key .align = "right",
							Key .sorttype = "int",
							Key .formatoptions = New formatoptions() With {
								.thousandsSeparator = sThousandSeparator,
								.decimalSeparator = "",
								.decimalPlaces = 0,
								.defaultValue = IIf(arrBlankIfZeroColumns(colCounter) = "1", "", "0")
									}
						})


					Case GetType(Decimal)
						Dim sThousandSeparator As String = ""
						Dim defaultValue As String = "0"

						Try
							If arrThousandSeparators.Length >= colCounter AndAlso arrThousandSeparators(colCounter) = "1" Then sThousandSeparator = LocaleThousandSeparator()

							If arrCellProps(col.Ordinal) > 0 Then
								defaultValue &= LocaleDecimalSeparator()
								For iloop = 1 To arrCellProps(col.Ordinal)
									defaultValue &= "0"
								Next
							End If

							If arrBlankIfZeroColumns.Length >= colCounter AndAlso arrBlankIfZeroColumns(colCounter) = "1" Then
								defaultValue = ""
							End If

						Catch ex As Exception
						End Try

						colModel.Add(New With {
							Key .name = col.ColumnName,
							Key .index = col.ColumnName,
							Key .sortable = True,
							Key .hidden = col.ColumnName = "ID" Or String.Concat(col.ColumnName, "xxx").Substring(0, 3) = "ID_",
							Key .label = col.ColumnName.Replace("_", " "),
							Key .formatter = "number",
							Key .align = "right",
							Key .sorttype = "number",
							Key .formatoptions = New formatoptions() With {
								.thousandsSeparator = sThousandSeparator,
								.decimalSeparator = LocaleDecimalSeparator(),
								.decimalPlaces = arrCellProps(col.Ordinal),
								.defaultValue = defaultValue
									}
						})

					Case Else
						colModel.Add(New With {
							Key .name = col.ColumnName,
							Key .index = col.ColumnName,
							Key .sortable = True,
							Key .hidden = col.ColumnName = "ID" Or String.Concat(col.ColumnName, "xxx").Substring(0, 3) = "ID_",
							Key .label = col.ColumnName.Replace("_", " "),
							Key .align = "left"
						})


				End Select

			Next

			Return colModel

		End Function
	End Class

	Public Class formatoptions

		Property srcformat() As String
		Property newformat() As String
		Property thousandsSeparator() As String
		Property decimalSeparator() As String
		Property decimalPlaces() As Integer
		Property defaultValue() As String

	End Class
End Namespace