Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Models.ObjectRequests
	Public Class DelegateBookingModel
		Inherits GotoOptionBaseModel

		Public Property Key1 As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property Key2 As String
		Public Property BookingStatus As Char

	End Class
End Namespace