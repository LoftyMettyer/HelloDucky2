Option Strict On
Option Explicit On

Namespace Classes
	Public Class PromptedValue
		Public Property Key As String
		Public Property Type As ExpressionValueTypes

		<AllowHtml>
		Public Property Value As String
	End Class
End Namespace