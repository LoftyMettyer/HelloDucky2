Option Strict On
Option Explicit On

Namespace Enums
	Public Enum ColumnType
		Unknown = -1
		Data = 0
		Lookup = 1
		Calculated = 2
		Relation = 3
		Link = 4
	End Enum
End Namespace