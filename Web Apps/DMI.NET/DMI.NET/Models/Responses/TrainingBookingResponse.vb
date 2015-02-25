Option Strict On
Option Explicit On

Namespace Models.Responses
	Public Class TrainingBookingResponse
		Inherits PostResponse

		Public NumberOfBookings As Integer
		Public CourseTitle As String

	End Class
End Namespace