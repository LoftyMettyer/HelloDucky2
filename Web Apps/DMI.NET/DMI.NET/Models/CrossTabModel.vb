Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel
Imports DMI.NET.AttributeExtensions

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
		Public Property HorizontalStart As Integer
		Public Property HorizontalStop As Integer
		Public Property HorizontalIncrement As Integer

		<HiddenInput>
		Public Property HorizontalDataType As SQLDataType

		<Range(1, Integer.MaxValue, ErrorMessage:="Vertical column not selected")>
		Public Property VerticalID As Integer

		Public Property VerticalStart As Integer
		Public Property VerticalStop As Integer
		Public Property VerticalIncrement As Integer

		<HiddenInput>
		Public Property VerticalDataType As SQLDataType

		Public Property PageBreakID As Integer
		Public Property PageBreakStart As Integer
		Public Property PageBreakStop As Integer
		Public Property PageBreakIncrement As Integer

		<HiddenInput>
		Public Property PageBreakDataType As SQLDataType

		Public Property IntersectionID As Integer

		<Required>
		<DisplayName("Type")>
		Public Property IntersectionType As IntersectionType

		<DisplayName("Percentage of Type")>
		Public Property PercentageOfType As Boolean

		<DisplayName("Percentage of Page")>
		Public Property PercentageOfPage As Boolean

		<DisplayName("Suppress zeroes")>
		Public Property SuppressZeros As Boolean

		<DisplayName("Use 1000 separators")>
		Public Property UseThousandSeparators As Boolean

		Public Property AvailableColumns As New List(Of ReportColumnItem)

		Public Property Output As New ReportOutputModel

		Public Overrides Sub SetBaseTable(TableID As Integer)
		End Sub

	End Class

End Namespace