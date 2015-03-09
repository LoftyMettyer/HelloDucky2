Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Models.ObjectRequests
	Public Class LookupFindModel
		Inherits GotoOptionBaseModel

		Public Property ColumnID As Integer
		Public Property LookupColumnID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property LookupValue As String

	End Class
End Namespace