Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Models.ObjectRequests
	Public Class PromptedValuesModel

		Public Property UtilType As UtilityType
		Public Property ID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property Name As String

	End Class
End Namespace