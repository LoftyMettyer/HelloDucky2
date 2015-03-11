Option Strict On
Option Explicit On

Namespace Models.ObjectRequests
	Public Class SendEmailModel

		Public Property __RequestVerificationToken As String

		Public Property [To] As String
		Public Property CC As String
		Public Property BCC As String
		Public Property Subject As String
		Public Property Body As String

	End Class
End Namespace
