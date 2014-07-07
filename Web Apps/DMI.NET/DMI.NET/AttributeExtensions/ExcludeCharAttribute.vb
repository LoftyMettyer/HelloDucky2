Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace AttributeExtensions

	Public Class ExcludeChar
		Inherits ValidationAttribute
		Private ReadOnly _chars As String

		Public Sub New(chars As String)
			MyBase.New("{0} contains invalid character.")
			_chars = chars
		End Sub

		Protected Overrides Function IsValid(value As Object, validationContext As ValidationContext) As ValidationResult
			If value IsNot Nothing Then
				For i As Integer = 0 To _chars.Length - 1
					Dim valueAsString = value.ToString()
					If valueAsString.Contains(_chars(i)) Then
						Dim errorMessage = FormatErrorMessage(validationContext.DisplayName)
						Return New ValidationResult(errorMessage)
					End If
				Next
			End If
			Return ValidationResult.Success
		End Function
	End Class

End Namespace