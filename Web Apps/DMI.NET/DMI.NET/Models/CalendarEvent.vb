Namespace Models

	Public Class CalendarEvent
		Public Property BaseID As String
		Public Property EventName As String
		Public Property Description As String
		Public Property StartDate() As DateTime
		Public Property StartSession() As String
		Public Property EndDate() As DateTime
		Public Property EndSession() As String
		Public Property Duration As Decimal
		Public Property Reason As String
		Public Property CalendarCode As String
		Public Property Region As String
		Public Property WorkingPattern As String
	End Class

End Namespace