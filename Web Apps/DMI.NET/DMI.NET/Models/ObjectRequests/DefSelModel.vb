﻿Option Strict On
Option Explicit On

Namespace Models.ObjectRequests
	Public Class DefSelModel

		Public Property __RequestVerificationToken As String

		Public Property txtTableID As Integer

		Public Property utiltype As UtilityType
		Public Property utilID As Integer
		Public Property utilName As String

		Public Property Action As String

	End Class
End Namespace
