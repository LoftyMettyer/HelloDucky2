Option Strict On
Option Explicit On

Namespace Models
	Public Class EventDetailModel

		Public Property ID As Integer
		Public Property Mode As String 'EventLogMode?
		Public Property BatchRunID As Integer

		Public ReadOnly Property IsBatch As Boolean
			Get
				Return (Mode = "Batch" Or Mode = "Pack")
			End Get
		End Property

	End Class

End Namespace