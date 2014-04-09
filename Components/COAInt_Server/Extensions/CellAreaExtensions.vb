Imports System.Runtime.CompilerServices
Imports Aspose.Cells

Namespace Extensions

	<HideModuleName()> _
	Friend Module CellAreaExtensions

		<Extension> _
		Public Function ToRange(obj As CellArea, objWorksheet As Worksheet) As Range

			Try
				With obj
					Return objWorksheet.Cells.CreateRange(.StartRow, .StartColumn, .EndRow - .StartRow + 1, .EndColumn - .StartColumn + 1)
				End With

			Catch ex As Exception
				Throw

			End Try
		End Function

		'Public Function ToCells(obj As CellArea, objWorksheet As Worksheet) As Cells

		'	Try
		'		With obj
		'			Return objWorksheet.Cells().Cells(.StartRow, .StartColumn, .EndRow - .StartRow + 1, .EndColumn - .StartColumn + 1)
		'		End With
		'	Catch ex As Exception
		'		Throw

		'	End Try

		'End Function

	End Module

End Namespace