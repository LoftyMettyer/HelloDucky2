Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports DMI.NET.AttributeExtensions

Namespace Classes
	Public Class ReportTableItem
		Implements IJsonSerialize

		Public Property [id] As Integer Implements IJsonSerialize.ID
		Public Property Name As String

	End Class
End Namespace