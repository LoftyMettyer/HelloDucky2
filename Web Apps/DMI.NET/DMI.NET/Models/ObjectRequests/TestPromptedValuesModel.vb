Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations
Imports DMI.NET.Classes

Namespace Models.ObjectRequests
	Public Class TestPromptedValuesModel

		Public Property UtilType As UtilityType
		Public Property TableID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<AllowHtml>
		Public Property components1 As String

		Public Property PromptValues As IList(Of PromptedValue)

	End Class
End Namespace