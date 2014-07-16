Option Strict On
Option Explicit On

Imports DMI.NET.Classes
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Namespace ViewModels

	Public Class ReportChildTableViewModel

		Public Property AvailableTables As Collection(Of ReportTableItem)

		Public Property ReportID As Integer
		Public Property TableID As Integer
		Public Property FilterID As Integer
		Public Property OrderID As Integer

		<DisplayName("Records :")>
		Public Property Records As Integer

		<DisplayName("Table :")>
		Public Property TableName As String

		<DisplayName("Filter :")>
		Public Property FilterName As String

		<DisplayName("Order :")>
		Public Property OrderName As String


	End Class

End Namespace