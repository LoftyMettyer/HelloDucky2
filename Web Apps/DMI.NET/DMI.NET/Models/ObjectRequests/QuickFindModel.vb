Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Models.ObjectRequests
	Public Class QuickFindModel
		Inherits GotoOptionBaseModel

		Public Property ViewID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property FilterSQL As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property FilterDef As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property Value As String

		Public Property ColumnID As Integer

	End Class
End Namespace