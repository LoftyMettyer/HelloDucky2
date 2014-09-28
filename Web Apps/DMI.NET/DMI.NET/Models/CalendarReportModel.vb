Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports System.ComponentModel.DataAnnotations
Imports DMI.NET.Code.Attributes
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports System.ComponentModel
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

		<MinLength(3, ErrorMessage:="You must select at least one event to report on.")> _
		Public Property EventsString As String

		Public Property StartType As CalendarDataType

		<DisplayName("Start Date")>
		<DisplayFormat(ApplyFormatInEditMode:=True, DataFormatString:="{0:dd/MM/yyyy}")>
		<RequiredIf("StartType", CalendarDataType.Fixed, ErrorMessage:="You must select a fixed start date.")>
		Public Property StartFixedDate As DateTime?

		Public Property StartOffset As Integer
		Public Property StartOffsetPeriod As DatePeriod

		<NonZeroIf("StartType", CalendarDataType.Custom, ErrorMessage:="No custom start date selected.")> _
		Public Property StartCustomId As Integer = 0
		Public Property StartCustomName As String

		Public Property EndType As CalendarDataType

		<DisplayName("End Date")>
		<DisplayFormat(ApplyFormatInEditMode:=True, DataFormatString:="{0:dd/MM/yyyy}")>
		<RequiredIf("EndType", CalendarDataType.Fixed, ErrorMessage:="You must select a fixed end date.")>
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

			SelectionType = RecordSelectionType.AllRecords
			Events = New Collection(Of CalendarEventDetailViewModel)()
			SortOrders = New List(Of SortOrderViewModel)

		End Sub

		Public Overrides Function GetAvailableSortColumns(Self As SortOrderViewModel) As IEnumerable(Of ReportColumnItem)

			Dim objItems As New Collection(Of ReportColumnItem)

			For Each objColumn In SessionInfo.Columns.Where(Function(m) m.TableID = BaseTableID AndAlso m.IsVisible).OrderBy(Function(m) m.Name)
				Dim objForEachSafety = objColumn
				If Not SortOrders.Any((Function(m) m.ColumnID = objForEachSafety.ID)) Then
					objItems.Add(New ReportColumnItem With {.ID = objColumn.ID, .Name = objColumn.Name})
				End If
			Next

			If Self.ColumnID > 0 Then
				Dim objColumn = SessionInfo.Columns.FirstOrDefault(Function(m) m.ID = Self.ColumnID)
				objItems.Add(New ReportColumnItem With {.ID = objColumn.ID, .Name = objColumn.Name})
			End If

			Return objItems.OrderBy(Function(m) m.Name)

		End Function

		Public Overrides ReadOnly Property SortOrdersAvailable As Integer
			Get
				If SessionInfo IsNot Nothing Then
					Return SessionInfo.Columns.AsEnumerable.Count(Function(m) m.TableID = BaseTableID AndAlso m.IsVisible) - SortOrders.Count()
				Else
					Return 0
				End If
			End Get
		End Property

		Public Property Description3ViewAccess As String
		Public Property StartCustomViewAccess As String
		Public Property EndCustomViewAccess As String

	End Class
End Namespace
