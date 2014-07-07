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

		<Required>
		Public Property HorizontalID As Integer

		Public Property HorizontalStart As Integer

		Public Property HorizontalStop As Integer
		Public Property HorizontalIncrement As Integer

		<Required>
		Public Property VerticalID As Integer
		Public Property VerticalStart As Integer
		Public Property VerticalStop As Integer
		Public Property VerticalIncrement As Integer
		Public Property PageBreakID As Integer
		Public Property PageBreakStart As Integer
		Public Property PageBreakStop As Integer
		Public Property PageBreakIncrement As Integer

		Public Property IntersectionID As Integer
		Public Property IntersectionType As IntersectionType
		Public Property PercentageOfType As Boolean
		Public Property PercentageOfPage As Boolean
		Public Property SuppressZeros As Boolean
		Public Property UseThousandSeparators As Boolean

		Public Property AvailableColumns As New List(Of ReportColumnItem)

		Public Property Output As New ReportOutputModel

	End Class

End Namespace