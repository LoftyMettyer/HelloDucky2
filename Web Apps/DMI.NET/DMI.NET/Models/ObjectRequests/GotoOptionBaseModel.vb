Option Strict On
Option Explicit On

Namespace Models.ObjectRequests
	Public Class GotoOptionBaseModel

		Public Property Action() As OptionActionType
		Public Property __RequestVerificationToken As String
		Public Property TableID As Integer

	End Class
End Namespace