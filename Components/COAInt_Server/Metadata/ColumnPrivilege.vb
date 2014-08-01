Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums

Namespace Metadata

	Friend Class ColumnPrivilege

		Friend Property AllowSelect As Boolean
		Friend Property AllowUpdate As Boolean
		Friend Property ColumnName As String
		Friend Property DataType As ColumnDataType
		Friend Property ColumnID As Integer

	End Class

End Namespace
