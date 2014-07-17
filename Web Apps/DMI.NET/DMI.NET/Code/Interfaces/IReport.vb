Option Explicit On
Option Strict On

Imports HR.Intranet.Server
Imports System.Collections.ObjectModel
Imports DMI.NET.ViewModels.Reports
Imports DMI.NET.Classes

Namespace Code.Interfaces
	Public Interface IReport

		Property SessionInfo As SessionInfo

		Property BaseTableID As Integer

		Sub SetBaseTable(BaseTableID As Integer)
		Function GetAvailableSortColumns() As IEnumerable(Of ReportColumnItem)

		Property SortOrders As Collection(Of SortOrderViewModel)
	End Interface
End Namespace