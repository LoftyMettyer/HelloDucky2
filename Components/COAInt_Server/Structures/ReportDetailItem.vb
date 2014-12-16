Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Enums

Namespace Structures
	Public Class ReportDetailItem
		Inherits Base

		Public IDColumnName As String
		Public DataType As ColumnDataType
		Public Size As Integer
		Public Decimals As Integer
		Public IsNumeric As Boolean
		Public IsAverage As Boolean
		Public IsCount As Boolean
		Public IsTotal As Boolean
		Public IsBreakOnChange As Boolean
		Public IsPageOnChange As Boolean
		Public IsValueOnChange As Boolean
		Public SuppressRepeated As Boolean
		Public LastValue As String
		Public Type As String
		Public TableID As Integer
		Public TableName As String
		Public ColumnName As String
		Public IsDateColumn As Boolean
		Public IsBitColumn As Boolean
		Public IsHidden As Boolean
		Public IsReportChildTable As Boolean
		Public Repetition As Boolean
		Public Use1000Separator As Boolean
		Public Mask As String
		Public GroupWithNextColumn As Boolean

	End Class

End Namespace
