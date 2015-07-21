Option Explicit On
Option Strict On

Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel

Namespace Models

	Public Class CrossTabModel
		Inherits ReportBaseModel

		Public Overrides ReadOnly Property ReportType As UtilityType
			Get
				Return UtilityType.utlCrossTab
			End Get
		End Property

		<Range(1, Integer.MaxValue, ErrorMessage:="Horizontal column not selected")>
		Public Property HorizontalID As Integer
		Public Property HorizontalStart As Double
		Public Property HorizontalStop As Double
		Public Property HorizontalIncrement As Double

		<HiddenInput>
		Public Property HorizontalDataType As ColumnDataType

		<Range(1, Integer.MaxValue, ErrorMessage:="Vertical column not selected")>
		Public Property VerticalID As Integer
		Public Property VerticalStart As Double
		Public Property VerticalStop As Double
		Public Property VerticalIncrement As Double

		<HiddenInput>
		Public Property VerticalDataType As ColumnDataType

		Public Property PageBreakID As Integer
		Public Property PageBreakStart As Double
		Public Property PageBreakStop As Double
		Public Property PageBreakIncrement As Double

		<HiddenInput>
		Public Property PageBreakDataType As ColumnDataType

		Public Property IntersectionID As Integer

		<Required>
		<DisplayName("Type :")>
		Public Property IntersectionType As IntersectionType = IntersectionType.Count

		<DisplayName("Percentage of Type")>
		Public Property PercentageOfType As Boolean

		<DisplayName("Percentage of Page")>
		Public Property PercentageOfPage As Boolean

		<DisplayName("Suppress zeroes")>
		Public Property SuppressZeros As Boolean

		<DisplayName("Use 1000 separators")>
		Public Property UseThousandSeparators As Boolean

		Public Property AvailableColumns As New List(Of ReportColumnItem)

		Public Overrides Property SortOrdersString As String

		Public Property Output As New ReportOutputModel

		Public Overrides Sub SetBaseTable(TableID As Integer)
		End Sub

		<RegularExpression("True", ErrorMessage:="Vertical stop value must be greater than its start value")>
		Public ReadOnly Property IsVerticalStopOK As Boolean
			Get
				Return (VerticalStop > VerticalStart OrElse VerticalStart = 0)
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Horizontal stop value must be greater than its start value")>
		Public ReadOnly Property IsHorizontalStopOK As Boolean
			Get
				Return (HorizontalStop > HorizontalStart OrElse HorizontalStart = 0)
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Page Break stop value must be greater than its start value")>
		Public ReadOnly Property IsPageBreakStopOK As Boolean
			Get
				Return (PageBreakStop > PageBreakStart OrElse PageBreakStart = 0)
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Horizontal increment must be greater than zero")>
		Public ReadOnly Property IsHorizontalIncrementOK1 As Boolean
			Get
				If (HorizontalStart > 0 OrElse HorizontalStop > 0) Then
					Return HorizontalIncrement > 0
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Maximum number of steps between start, stop and increment value for the horizontal range has been exceeded")>
		Public ReadOnly Property IsHorizontalIncrementOK2 As Boolean
			Get
				If HorizontalIncrement > 0 Then
					Return (HorizontalStop - HorizontalStart) / HorizontalIncrement <= 32768
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Vertical increment must be greater than zero")>
	 Public ReadOnly Property IsVerticalIncrementOK1 As Boolean
			Get
				If (VerticalStart > 0 OrElse VerticalStop > 0) Then
					Return VerticalIncrement > 0
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Maximum number of steps between start, stop and increment value for the vertical range has been exceeded")>
		Public ReadOnly Property IsVerticalIncrementOK2 As Boolean
			Get
				If VerticalIncrement > 0 Then
					Return (VerticalStop - VerticalStart) / VerticalIncrement <= 32768
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Page break increment must be greater than zero")>
	 Public ReadOnly Property IsPageBreakIncrementOK1 As Boolean
			Get
				If (PageBreakStart > 0 OrElse PageBreakStop > 0) Then
					Return PageBreakIncrement > 0
				Else
					Return True
				End If
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Maximum number of steps between start, stop and increment value for the page break range has been exceeded")>
		Public ReadOnly Property IsPageBreakIncrementOK2 As Boolean
			Get
				If PageBreakIncrement > 0 Then
					Return (PageBreakStop - PageBreakStart) / PageBreakIncrement <= 32768
				Else
					Return True
				End If
			End Get
		End Property


	End Class

End Namespace