Option Strict On
Option Explicit On

Namespace Classes
	Public Class AbsenceBreakdownDate
		Public ContainsData As Boolean
		Public IsWeekend As Boolean
		Public Caption As String
		Public IsBankHoliday As Boolean
		Public IsWorkingDay As Boolean
		Public DisplayColor As String
		Public Type As String
		Public Reason As String
		Public WorkingPattern As String
		Public Duration As Double
		Public StartDate As DateTime
		Public StartSession As String
		Public EndDate As DateTime
		Public EndSession As String
		Public Region As String

	End Class

End Namespace