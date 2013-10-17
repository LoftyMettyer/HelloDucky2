Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums

Namespace Metadata
	Public Class Column
		Inherits Base
			Public TableID As Integer
			Public TableName As String
			Public DataType As SQLDataType
			Public Size As Long									' Needs to be long to handle ole embedded ole types.
			Public Decimals As Short
			Public Use1000Separator As Boolean
	End Class
End Namespace