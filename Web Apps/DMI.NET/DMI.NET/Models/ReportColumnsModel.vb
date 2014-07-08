Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports DMI.NET.Classes
Imports DMI.NET.Enums

Namespace Models

	Public Class ReportColumnsModel

		Public Property BaseTableID As Integer
		Public Property Selected As New Collection(Of ReportColumnItem)
		Public Property SelectedTableID As Integer

		Public Property AvailableTables As New List(Of ReportTableItem)
		Public Property SelectionType As ColumnSelectionType = ColumnSelectionType.Columns

	End Class

End Namespace