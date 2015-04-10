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

		<NonZeroIf("StartType", CalendarDataType.Custom, ErrorMessage:="You must select a calculation for the report start date.")> _
		Public Property StartCustomId As Integer = 0
		Public Property StartCustomName As String

		Public Property EndType As CalendarDataType

		<DisplayName("End Date")>
		<DisplayFormat(ApplyFormatInEditMode:=True, DataFormatString:="{0:dd/MM/yyyy}")>
		<RequiredIf("EndType", CalendarDataType.Fixed, ErrorMessage:="You must select a fixed end date.")>
		Public Property EndFixedDate As DateTime?

		Public Property EndOffset As Integer
		Public Property EndOffsetPeriod As DatePeriod

		<NonZeroIf("EndType", CalendarDataType.Custom, ErrorMessage:="You must select a calculation for the report end date.")> _
		Public Property EndCustomId As Integer = 0
		Public Property EndCustomName As String

		<DisplayName("Include Bank Holidays *")> _
		Public Property IncludeBankHolidays As Boolean

		<DisplayName("Working Days Only *")> _
		Public Property WorkingDaysOnly As Boolean

		<DisplayName("Show Bank Holidays *")> _
		Public Property ShowBankHolidays As Boolean

		<DisplayName("Show Calendar Captions *")> _
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
				Dim skipMe As Boolean = (objForEachSafety.DataType = ColumnDataType.sqlOle OrElse objForEachSafety.DataType = ColumnDataType.sqlVarBinary)
				If skipMe = False Then
					If Not SortOrders.Any((Function(m) m.ColumnID = objForEachSafety.ID)) Then
						objItems.Add(New ReportColumnItem With {.ID = objColumn.ID, .Name = objColumn.Name})
					End If
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

		<RegularExpression("True", ErrorMessage:="You must select at least one base description column or calculation for the report.")>
	 Public ReadOnly Property IsDescriptionOK As Boolean
			Get
				Return (Description1ID > 0 OrElse Description2ID > 0 OrElse Description3ID > 0)
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="You must select a fixed end date later than or equal to the fixed start date.")>
	 Public ReadOnly Property IsFixedDatesOK As Boolean
			Get
				If StartType = CalendarDataType.Fixed AndAlso EndType = CalendarDataType.Fixed Then

					If StartFixedDate Is Nothing Then Return True
					If EndFixedDate Is Nothing Then Return True

					Return (CDate(EndFixedDate) >= CDate(StartFixedDate))
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="You must select an end date offset greater than or equal to zero.")>
	 Public ReadOnly Property IsEndOffsetDateOK As Boolean
			Get
				If (StartType = CalendarDataType.Fixed OrElse StartType = CalendarDataType.CurrentDate) AndAlso EndType = CalendarDataType.Offset Then
					Return EndOffset >= 0
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="You must select a start date offset less than or equal to zero.")>
		Public ReadOnly Property IsStartOffsetDateOK As Boolean
			Get
				If (EndType = CalendarDataType.Fixed OrElse EndType = CalendarDataType.CurrentDate) AndAlso StartType = CalendarDataType.Offset Then
					Return StartOffset <= 0
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="The end offset period must be the same as the start date offset period.")>
	 Public ReadOnly Property IsOffsetPeriodOK1 As Boolean
			Get
				If (EndType = CalendarDataType.Offset AndAlso StartType = CalendarDataType.Offset) Then
					Return StartOffsetPeriod = EndOffsetPeriod
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="The start offset period must be before the end date offset period.")>
	 Public ReadOnly Property IsOffsetPeriodOK2 As Boolean
			Get
				If (EndType = CalendarDataType.Offset AndAlso StartType = CalendarDataType.Offset) Then
					Return StartOffset <= EndOffset
				Else
					Return True
				End If
			End Get
		End Property


	End Class
End Namespace
