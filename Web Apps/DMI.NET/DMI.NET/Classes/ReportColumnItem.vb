Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports DMI.NET.AttributeExtensions

Namespace Classes
	Public Class ReportColumnItem

		Public Property id As Integer
		Public Property IsExpression As Boolean
		Public Property Name As String
		Public Property CustomReportId As Integer
		Public Property Sequence As Integer

		<ExcludeChar("/,.!@#$%")>
		Public Property Heading As String

		Public Property DataType As SQLDataType
		Public Property Size As Long
		Public Property Decimals As Integer
		Public Property IsAverage As Boolean
		Public Property IsCount As Boolean
		Public Property IsTotal As Boolean
		Public Property IsHidden As Boolean
		Public Property IsGroupWithNext As Boolean

	End Class
End Namespace