Option Strict On
Option Explicit On

Namespace Classes
	Public Class ExpressionSelectionItem
		Implements IJsonSerialize

		Public Property [ID] As Integer Implements IJsonSerialize.ID
		Public Property Name As String
		Public Property Description As String
		Public Property UserName As String
		Public Property Access As String

	End Class
End Namespace