Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports HR.Intranet.Server.Enums
Imports DMI.NET.Classes

Namespace ViewModels.Reports
	Public Class SortOrderViewModel
		Implements IJsonSerialize
		Implements IReportDetail

		Public Property ReportID As Integer Implements IReportDetail.ReportID
		Public Property ReportType As UtilityType Implements IReportDetail.ReportType

		Public Property TableID As Integer
		Public Property ID As Integer Implements IJsonSerialize.ID

		<DisplayName("Column :")>
		Public Property ColumnID As Integer

		<DisplayName("Column :")>
		Public Property Name As String

		<DisplayName("Order :")>
		Public Property Order As OrderType

		Public Property Sequence As Integer

		<DisplayName("Break on Change :")>
		Public Property BreakOnChange As Boolean

		<DisplayName("Page on Change :")>
		Public Property PageOnChange As Boolean

		<DisplayName("Value on Change :")>
		Public Property ValueOnChange As Boolean

		<DisplayName("Suppress Repeated Values :")>
		Public Property SuppressRepeated As Boolean

		Public Property AvailableColumns As IEnumerable(Of ReportColumnItem)

	End Class
End Namespace