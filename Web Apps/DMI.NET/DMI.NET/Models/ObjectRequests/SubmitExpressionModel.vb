Option Strict On
Option Explicit On

Namespace Models.ObjectRequests

	Public Class SubmitExpressionModel
		Public Property txtSend_ID As Integer
		Public Property txtSend_type As UtilityType

		<AllowHtml>
		Public Property txtSend_name As String

		<AllowHtml>
		Public Property txtSend_description As String

		Public Property txtSend_access As String
		Public Property txtSend_userName As String

		<AllowHtml>
		Public Property txtSend_components1 As String
		Public Property txtSend_reaction As String
		Public Property txtSend_tableID As Integer

		<AllowHtml>
		Public Property txtSend_names As String
	End Class
End Namespace
