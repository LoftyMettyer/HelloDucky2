Option Strict On
Option Explicit On

Namespace Models.ObjectRequests
	Public Class TestExpressionModel
		Public Property type As UtilityType
		Public Property components1 As String

		Public Property TableID As Integer

		<AllowHtml>
		Public Property prompts As String
		Public Property filtersAndCalcs As String
	End Class
End Namespace