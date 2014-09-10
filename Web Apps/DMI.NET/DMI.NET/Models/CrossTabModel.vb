Option Explicit On
Option Strict On

Imports DMI.NET.Classes
Imports Foolproof
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

		<RegularExpression("True", ErrorMessage:="Page Break stop value must be greater its start value")>
		Public ReadOnly Property IsPageBreakStopOK As Boolean
			Get
				Return (PageBreakStop > PageBreakStart OrElse PageBreakStart = 0)
			End Get
		End Property

	End Class

End Namespace