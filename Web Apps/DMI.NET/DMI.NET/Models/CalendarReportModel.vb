Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports System.ComponentModel
Imports DMI.NET.ViewModels
Imports DMI.NET.AttributeExtensions
Imports DMI.NET.ViewModels.Reports

Namespace Models

	Public Class CalendarReportModel

		Inherits ReportBaseModel

		Public Overrides ReadOnly Property ReportType As UtilityType
			Get
				Return UtilityType.utlCalendarReport
			End Get
		End Property

		Public Property Description3Name As String

		Public Property Events As New Collection(Of CalendarEventDetailViewModel)
		Public Property EventsString As String

		Public Property StartType As CalendarDataType
		Public Property StartFixedDate As DateTime?
		Public Property StartOffset As Integer
		Public Property StartOffsetPeriod As DatePeriod

		<NonZeroIf("StartType", CalendarDataType.Custom, ErrorMessage:="No custom start date selected.")> _
		Public Property StartCustomId As Integer = 0
		Public Property StartCustomName As String

		Public Property EndType As CalendarDataType
		Public Property EndFixedDate As DateTime?
		Public Property EndOffset As Integer
		Public Property EndOffsetPeriod As DatePeriod

		<NonZeroIf("EndType", CalendarDataType.Custom, ErrorMessage:="No custom end date selected.")> _
		Public Property EndCustomId As Integer = 0
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

		Public Overrides Sub SetBaseTable(TableID As Integer)
		End Sub

		Public Overrides Function GetAvailableSortColumns(Self As SortOrderViewModel) As IEnumerable(Of ReportColumnItem)

			Dim objItems As New Collection(Of ReportColumnItem)

			For Each objColumn In SessionInfo.Columns.Where(Function(m) m.TableID = BaseTableID And m.IsVisible).OrderBy(Function(m) m.Name)
				objItems.Add(New ReportColumnItem With {.ID = objColumn.ID, .Name = objColumn.Name})
			Next

			Return objItems

		End Function

	End Class
End Namespace
