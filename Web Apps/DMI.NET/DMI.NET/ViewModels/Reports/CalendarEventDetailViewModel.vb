Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums

Namespace ViewModels
	Public Class CalendarEventDetailViewModel
		Implements IJsonSerialize

		<Browsable(False)>
		Public Property ID As Integer Implements IJsonSerialize.ID
		Public Property EventKey As String

		Public Property CalendarReportID As Integer

		<DisplayName("Name :")>
		Public Property Name As String

		<DisplayName("Event Table :")>
		Public Property TableID As Integer

		<DisplayName("Filter :")>
		Public Property FilterID As Integer

		<DisplayName("Start Date :")>
		Public Property EventStartDateID As Integer

		<DisplayName("End Date :")>
		Public Property EventStartSessionID As Integer

		Public Property EventEndType As CalendarEventEndType

		Public Property EventEndDateID As Integer
		Public Property EventEndSessionID As Integer
		Public Property EventDurationID As Integer

		Public Property LegendType As CalendarLegendType
		Public Property LegendCharacter As String
		Public Property LegendLookupTableID As Integer
		Public Property LegendLookupColumnID As Integer
		Public Property LegendLookupCodeID As Integer
		Public Property LegendEventColumnID As Integer

		<DisplayName("Description 1 :")>
		Public Property EventDesc1ColumnID As Integer

		<DisplayName("Description 2 :")>
		Public Property EventDesc2ColumnID As Integer

		Public Property FilterHidden As String

		' For display purposes in grids
		Public Property FilterName As String
		Public Property EventStartSessionName As String
		Public Property EventEndDateName As String
		Public Property EventEndSessionName As String
		Public Property EventDurationName As String
		Public Property LegendTypeName As String
		Public Property EventDesc1ColumnName As String
		Public Property EventDesc2ColumnName As String

		Public Property AvailableTables As List(Of ReportTableItem)


	End Class
End Namespace


