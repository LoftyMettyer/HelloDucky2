Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums

Namespace Metadata
	Public Class Column
		Inherits Base

		Public TableID As Integer
		Public TableName As String
		Public DataType As SQLDataType = SQLDataType.sqlUnknown
		Public Size As Long									' Needs to be long to handle ole embedded ole types.
		Public Decimals As Integer
		Public Use1000Separator As Boolean

		Public ReadOnly Property IsVisible As Boolean
			Get
				Return Name <> "ID" And Not Name.StartsWith("ID_")
			End Get
		End Property

	End Class
End Namespace