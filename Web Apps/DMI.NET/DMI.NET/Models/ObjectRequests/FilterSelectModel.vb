Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Models.ObjectRequests
	Public Class FilterSelectModel
		Inherits GotoOptionBaseModel

		Public Property ScreenID As Integer
		Public Property ViewID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property FilterSQL As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property FilterDef As String

	End Class
End Namespace