Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums

Namespace Metadata
	Public Class TablePrivilege
		Friend Property AllowSelect() As Boolean
		Friend Property AllowUpdate() As Boolean
		Friend Property AllowInsert() As Boolean
		Friend Property AllowDelete() As Boolean
		Friend Property TableID() As Integer
		Friend Property ViewID() As Integer
		Friend Property RealSource() As String
		Friend Property IsTable() As Boolean
		Friend Property TableName() As String
      Friend Property OriginalViewName() As String
      Friend Property ViewName() As String
      Friend Property TableType() As TableTypes
		Friend Property DefaultOrderID() As Integer
		Friend Property RecordDescriptionID() As Integer
	End Class
End Namespace