Option Strict On
Option Explicit On

Imports DMI.NET.Code.Attributes

Namespace Classes
	Public Class ReportColumnItem
		Implements IJsonSerialize
		Implements IReportDetail

		Public Property ReportID As Integer Implements IReportDetail.ReportID
		Public Property ReportType As UtilityType Implements IReportDetail.ReportType

		Public Property ID As Integer Implements IJsonSerialize.ID
		Public Property TableID As Integer
		Public Property IsExpression As Boolean
		Public Property Name As String
		Public Property Sequence As Integer

		<ExcludeChar("/,.!@#$%")>
		Public Property Heading As String

		Public Property DataType As ColumnDataType
		Public Property Size As Long
		Public Property Decimals As Integer
		Public Property IsAverage As Boolean
		Public Property IsCount As Boolean
		Public Property IsTotal As Boolean
		Public Property IsHidden As Boolean
		Public Property IsGroupWithNext As Boolean
		Public Property IsRepeated As Boolean

		''' <summary>
		''' Gets/Sets the access rights for the column (E.g. HD/RW/RO)
		''' </summary>
		''' <value>The column access value</value>
		Public Property Access As String


		Public ReadOnly Property IsNumeric As Boolean
			Get
				Return DataType = ColumnDataType.sqlInteger OrElse DataType = ColumnDataType.sqlNumeric
			End Get
		End Property

	End Class
End Namespace