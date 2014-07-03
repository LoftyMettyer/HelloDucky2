Option Strict On
Option Explicit On

Imports System.ComponentModel

Namespace Classses
	Public Class CalendarEventDetail

		<Browsable(False)>
		Public Property ID As Integer
		Public Property EventKey As String

		Public Property CalendarReportID As Integer
		Public Property Name As String
		Public Property TableID As Integer
		Public Property FilterID As Integer
		Public Property EventStartDateID As Integer
		Public Property EventStartSessionID As Integer
		Public Property EventEndDateID As Integer
		Public Property EventEndSessionID As Integer
		Public Property EventDurationID As Integer
		Public Property LegendType As String
		Public Property LegendCharacter As String
		Public Property LegendLookupTableID As Integer
		Public Property LegendLookupColumnID As Integer
		Public Property LegendLookupCodeID As Integer
		Public Property LegendEventColumnID As Integer

		<Browsable(False)>
		Public Property EventDesc1ColumnID As Integer

		<Browsable(False)>
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

	End Class
End Namespace


