Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Models.ObjectRequests
	Public Class PicklistSelectionModel
		Public Property TableID As Integer
		Public Property Action As String
		Public Property Type As String = "ALL"

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property IDs1 As String
	End Class
End Namespace