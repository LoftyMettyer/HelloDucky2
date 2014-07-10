Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports System.ComponentModel
Imports DMI.NET.Classses

Namespace Models

	Public Class CalendarReportModel
		Inherits ReportBaseModel

		Public Overrides ReadOnly Property ReportType As UtilityType
			Get
				Return UtilityType.utlCalendarReport
			End Get
		End Property

		Public Property Description1Id As Integer
		Public Property Description2Id As Integer
		Public Property Description3Id As Integer
		Public Property Description3Name As String
		Public Property RegionID As Integer
		Public Property GroupByDescription As Boolean
		Public Property Separator As String

		Public Property Events As New Collection(Of CalendarEventDetail)

		Public Property StartType As CalendarDataType
		Public Property StartFixedDate As DateTime
		Public Property StartOffset As Integer
		Public Property StartOffsetPeriod As DatePeriod
		Public Property StartCustomId As Integer
		Public Property StartCustomName As String

		Public Property EndType As CalendarDataType
		Public Property EndFixedDate As DateTime
		Public Property EndOffset As Integer
		Public Property EndOffsetPeriod As DatePeriod
		Public Property EndCustomId As Integer
		Public Property EndCustomName As String

		<DisplayName("Include Bank Holidays")> _
		Public Property IncludeBankHolidays As Boolean

		<DisplayName("Working Days Only")> _
		Public Property WorkingDaysOnly As Boolean

		<DisplayName("Show Bank Holidays")> _
		Public Property ShowBankHolidays As Boolean

		<DisplayName("Show Calendar Captions")> _
		Public Property ShowCaptions As Boolean

		<DisplayName("Show Weekends")> _
		Public Property ShowWeekends As Boolean

		<DisplayName("Start on Current Month")> _
		Public Property StartOnCurrentMonth As Boolean

		Public Property Output As New ReportOutputModel

	End Class
End Namespace
