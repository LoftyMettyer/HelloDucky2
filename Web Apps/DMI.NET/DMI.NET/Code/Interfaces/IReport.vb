Option Explicit On
Option Strict On

Imports HR.Intranet.Server
Imports System.Collections.ObjectModel
Imports DMI.NET.ViewModels.Reports
Imports DMI.NET.Classes

Namespace Code.Interfaces
	Public Interface IReport

		Property ID As Integer
		ReadOnly Property ReportType As UtilityType
		Property Owner As String

		Property SessionInfo As SessionInfo

		Property BaseTableID As Integer

      Sub SetBaseTable(TableID As Integer)

      Property BaseViewID As Integer

      Function GetAvailableSortColumns(Self As SortOrderViewModel) As IEnumerable(Of ReportColumnItem)
		Function GetAvailableTables() As IEnumerable(Of ReportTableItem)
		Property Columns() As List(Of ReportColumnItem)

		Property SortOrders As List(Of SortOrderViewModel)
		ReadOnly Property SortOrdersAvailable As Integer

      Property Dependencies() As ReportDependencies

   End Interface
End Namespace