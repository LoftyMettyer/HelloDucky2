Imports System.ComponentModel
Namespace Models

	Public Class CalendarEvent
		Public Property BaseID As String
		<DisplayName("Event Name :")>
		Public Property EventName As String
		<DisplayName("Description :")>
		Public Property Description As String

		<DisplayName("Start Date :")>
		Public Property StartDate() As DateTime
		Public Property StartSession() As String
		<DisplayName("End Date :")>
		Public Property EndDate() As DateTime
		Public Property EndSession() As String
		<DisplayName("Duration (Actual) :")>
		Public Property Duration As Decimal

		<DisplayName("Reason :")>
		Public Property Reason As String

		<DisplayName("Event Name :")>
		Public Property Description1 As String
		<DisplayName("Event Name :")>
		Public Property Description2 As String
		Public Property Description1Column As String
		Public Property Description2Column As String

		Public Property CalendarCode As String

		<DisplayName("Region :")>
		Public Property Region As String

		<DisplayName("Working Pattern :")>
		Public Property WorkingPattern As String
	End Class

End Namespace