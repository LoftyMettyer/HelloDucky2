Option Explicit On
Option Strict On

Namespace Classes
	Public Class ReportRepetition
		Implements IJsonSerialize

		Public Property ID As Integer Implements IJsonSerialize.ID
		Public Property Name As String
		Public Property IsRepeated As Boolean
		Public Property IsExpression As Boolean
	End Class
End Namespace
