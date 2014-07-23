﻿Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports DMI.NET.Classes

Namespace ViewModels.Reports

	Public Class ChildTableViewModel
		Implements IJsonSerialize
		Implements IReportDetail

		<HiddenInput>
		Public Property ID As Integer Implements IJsonSerialize.ID

		<HiddenInput>
		Public Property ReportID As Integer Implements IReportDetail.ReportID
		Public Property ReportType As HR.Intranet.Server.Enums.UtilityType Implements IReportDetail.ReportType

		<DisplayName("Table :")>
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

		Public Property AvailableTables As New List(Of ReportTableItem)

	End Class

End Namespace