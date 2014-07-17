Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports DMI.NET.Classes
Imports DMI.NET.Enums
Imports HR.Intranet.Server.Enums

Namespace ViewModels.Reports

	Public Class ColumnsViewModel

		Public Property ReportID As Integer

		Public Property DisplayTableSelection As Boolean

		Public Property BaseTableID As Integer
		Public Property Selected As New Collection(Of ReportColumnItem)
		Public Property SelectedTableID As Integer

		Public Property AvailableTables As New List(Of ReportTableItem)
		Public Property SelectionType As ColumnSelectionType = ColumnSelectionType.Columns

	End Class

End Namespace