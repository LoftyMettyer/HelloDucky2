Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums

Namespace Metadata
	Public Class Column
		Inherits Base

		Public TableID As Integer
		Public TableName As String
		Public DataType As ColumnDataType = ColumnDataType.sqlUnknown
		Public Size As Long									' Needs to be long to handle ole embedded ole types.
		Public ColumnSize As Long
		Public Decimals As Integer
		Public Use1000Separator As Boolean
		Public ColumnType As ColumnType
		Public LookupTableID As Integer
		Public LookupColumnID As Integer

		Public ReadOnly Property IsVisible As Boolean
			Get
				Return ColumnType <> ColumnType.Relation
			End Get
		End Property

		Public ReadOnly Property IsNumeric As Boolean
			Get
				Return DataType = ColumnDataType.sqlInteger OrElse DataType = ColumnDataType.sqlNumeric
			End Get
		End Property

	End Class
End Namespace