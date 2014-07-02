Option Strict On
Option Explicit On

Imports System.ComponentModel

Namespace Classses
	Public Class CalendarEventDetail

		'Public Property EventId As Integer
		'Public Property TableId As Integer
		'Public Property FilterId As Integer
		'Public Property StartDateId As Integer
		'Public Property StartSessionId As Integer
		'Public Property EndDateId As Integer
		'Public Property EndSessionId As Integer
		'Public Property DurationId As Integer
		'Public Property KeyId As Integer
		'Public Property Description1Id As Integer
		'Public Property Description2Id As Integer

		'Public Property EventName As String
		'Public Property TableName As String
		'Public Property FilterName As String
		'Public Property StartDateName As String
		'Public Property StartSessionName As String
		'Public Property EndDateName As String
		'Public Property EndSessionName As String
		'Public Property DurationName As String
		'Public Property KeyName As String
		'Public Property Description1Name As String
		'Public Property Description2Name As String

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
		Public Property EventDesc1ColumnID As Integer
		Public Property EventDesc2ColumnID As Integer

	End Class
End Namespace