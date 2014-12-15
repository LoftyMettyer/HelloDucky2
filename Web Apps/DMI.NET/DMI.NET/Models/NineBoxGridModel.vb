Option Explicit On
Option Strict On

Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel
Imports Microsoft.Ajax.Utilities

Namespace Models
	Public Class NineBoxGridModel
		Inherits ReportBaseModel

		Public Overrides ReadOnly Property ReportType As UtilityType
			Get
				Return UtilityType.utlNineBoxGrid
			End Get
		End Property

		<Range(1, Integer.MaxValue, ErrorMessage:="Horizontal column not selected")>
		Public Property HorizontalID As Integer
		Public Property HorizontalStart As Double
		Public Property HorizontalStop As Double

		<HiddenInput>
		Public Property HorizontalDataType As ColumnDataType

		<Range(1, Integer.MaxValue, ErrorMessage:="Vertical column not selected")>
		Public Property VerticalID As Integer
		Public Property VerticalStart As Double
		Public Property VerticalStop As Double

		<HiddenInput>
		Public Property VerticalDataType As ColumnDataType

		Public Property PageBreakID As Integer

		<HiddenInput>
		Public ReadOnly Property PageBreakStart As Double
			Get
				Return 0
			End Get
		End Property

		<HiddenInput>
		Public ReadOnly Property PageBreakStop As Double
			Get
				Return 0
			End Get
		End Property

		<HiddenInput>
		Public Property PageBreakDataType As ColumnDataType

		<HiddenInput>
		Public ReadOnly Property IntersectionID As Integer
			Get
				Return 0 ' "None"
			End Get
		End Property

		<HiddenInput>
		Public ReadOnly Property IntersectionType As IntersectionType
			Get
				Return IntersectionType.Count
			End Get
		End Property

		<DisplayName("Percentage of all data")>
		Public Property PercentageOfType As Boolean

		<DisplayName("Percentage of page")>
		Public Property PercentageOfPage As Boolean

		<DisplayName("Suppress zeroes")>
		Public Property SuppressZeros As Boolean

		<DisplayName("Use 1000 separators")>
		Public Property UseThousandSeparators As Boolean

		Public Property AvailableColumns As New List(Of ReportColumnItem)

		Public Overrides Property SortOrdersString As String

		Public Property Output As New ReportOutputModel

		'The default values for the following fields come from the FD document (Colors are in Hex)
		Public Property XAxisLabel As String = "Performance"
		Public Property XAxisSubLabel1 As String = "Low"
		Public Property XAxisSubLabel2 As String = "Medium"
		Public Property XAxisSubLabel3 As String = "High"
		Public Property YAxisLabel As String = "Potential"
		Public Property YAxisSubLabel1 As String = "High"
		Public Property YAxisSubLabel2 As String = "Medium"
		Public Property YAxisSubLabel3 As String = "Low"
		Public Property Description1 As String = "Enigma"
		Public Property ColorDesc1 As String = "FFFF00"
		Public Property Description2 As String = "Future Star"
		Public Property ColorDesc2 As String = "33CC33"
		Public Property Description3 As String = "Consistent Star"
		Public Property ColorDesc3 As String = "006600"
		Public Property Description4 As String = "Inconsistent Player"
		Public Property ColorDesc4 As String = "FF9900"
		Public Property Description5 As String = "Trusted Professional"
		Public Property ColorDesc5 As String = "FFFF00"
		Public Property Description6 As String = "Current Star"
		Public Property ColorDesc6 As String = "33CC33"
		Public Property Description7 As String = "Talent Risk"
		Public Property ColorDesc7 As String = "CC3300"
		Public Property Description8 As String = "Solid Professional"
		Public Property ColorDesc8 As String = "FF9900"
		Public Property Description9 As String = "Key Player"
		Public Property ColorDesc9 As String = "FFFF00"

		Public Overrides Sub SetBaseTable(TableID As Integer)
		End Sub

		<RegularExpression("True", ErrorMessage:="Vertical Maximum Value must be greater than its Minimum Value")>
		Public ReadOnly Property IsVerticalStopOK As Boolean
			Get
				Return (VerticalStop > VerticalStart)
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="Horizontal Maximum Value must be greater than its Minimum Value")>
		Public ReadOnly Property IsHorizontalStopOK As Boolean
			Get
				Return (HorizontalStop > HorizontalStart)
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="All axis labels must be completed")>
		Public ReadOnly Property AreAllAxisLabelsOk As Boolean
			Get
				Return Not ( _
						String.IsNullOrEmpty(XAxisLabel) Or _
						String.IsNullOrEmpty(XAxisSubLabel1) Or _
						String.IsNullOrEmpty(XAxisSubLabel2) Or _
						String.IsNullOrEmpty(XAxisSubLabel3) Or _
						String.IsNullOrEmpty(YAxisLabel) Or _
						String.IsNullOrEmpty(YAxisSubLabel1) Or _
						String.IsNullOrEmpty(YAxisSubLabel2) Or _
						String.IsNullOrEmpty(YAxisSubLabel3) _
						)
			End Get
		End Property
	End Class

End Namespace