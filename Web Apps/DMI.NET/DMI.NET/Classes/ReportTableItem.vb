Option Strict On
Option Explicit On

Namespace Classes
	Public Class ReportTableItem
		Implements IJsonSerialize

		Public Property [id] As Integer Implements IJsonSerialize.ID
		Public Property Name As String

	End Class
End Namespace