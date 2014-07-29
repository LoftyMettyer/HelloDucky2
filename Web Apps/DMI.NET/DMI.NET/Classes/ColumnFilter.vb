Imports HR.Intranet.Server.Enums

Namespace Classes
	Public Class ColumnFilter
		Public TableID As Integer
		Public DataType As SQLDataType = SQLDataType.sqlUnknown
		Public ColumnType As ColumnType = ColumnType.Unknown
		Public Size As Integer = 0
		Public AddNone As Boolean = False
		Public ShowFullName As Boolean = False
		Public IncludeParents As Boolean = False
	End Class
End Namespace