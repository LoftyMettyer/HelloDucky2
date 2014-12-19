Option Strict On
Option Explicit On

Namespace Classes
	Public Class ColumnFilter
		Public TableID As Integer
		Public ColumnType As ColumnType = ColumnType.Unknown
		Public DataType As ColumnDataType = ColumnDataType.sqlUnknown
		Public Size As Integer = 0
		Public AddNone As Boolean = False
		Public AddDefault As Boolean = False
		Public ShowFullName As Boolean = False
		Public IncludeParents As Boolean = False
		Public IsNumeric As Boolean
	End Class
End Namespace