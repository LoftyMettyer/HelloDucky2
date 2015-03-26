Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations
Imports DMI.NET.Classes

Namespace Models.ObjectRequests
	Public Class PromptedValuesModel

		Public Property UtilType As UtilityType
		Public Property ID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<AllowHtml>
		Public Property Name As String

		Public Property PromptValues As IList(Of PromptedValue)

		Public Property IsBulkBooking As Boolean

	End Class
End Namespace