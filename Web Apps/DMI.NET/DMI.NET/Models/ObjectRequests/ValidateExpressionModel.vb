Option Strict On
Option Explicit On

Namespace Models.ObjectRequests
	Public Class ValidateExpressionModel

		Public Property Action As String
		Public Property validatePass As Integer

		<AllowHtml>
		Public Property validateName As String
		Public Property validateOwner As String
		Public Property validateTimestamp As Integer
		Public Property validateUtilID As Integer
		Public Property validateUtilType As UtilityType
		Public Property validateAccess As String

		<AllowHtml>
		Public Property components1 As String

		Public Property validateBaseTableID As Integer
		Public Property validateOriginalAccess As String

	End Class
End Namespace