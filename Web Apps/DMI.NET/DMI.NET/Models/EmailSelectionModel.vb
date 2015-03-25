Option Strict On
Option Explicit On
Imports System.ComponentModel.DataAnnotations

Namespace Models
	Public Class EmailSelectionModel

		Public Property SelectedEventIDs As String
		Public Property IsFromMain As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailOrderColumn As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailOrderOrder As String

		Public Property IsBatchy As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property BatchInfo As String

	End Class

End Namespace