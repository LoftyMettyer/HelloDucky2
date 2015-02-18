Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Code.Attributes

	Public Class NoneAttribute
		Inherits ValidationAttribute

		Public Property PropertyName() As [String]
		Private Property DesiredValue() As [Object]

		Public Sub New(propertyName As [String], desiredValue As [String])
			Me.PropertyName = propertyName
			Me.DesiredValue = desiredValue
		End Sub

		Protected Overrides Function IsValid(value As Object, context As ValidationContext) As ValidationResult
			Dim instance As [Object] = context.ObjectInstance
			Dim type As Type = instance.[GetType]()
			Dim proprtyvalue As [Object] = type.GetProperty(PropertyName).GetValue(instance, Nothing)
			If proprtyvalue.ToString() = DesiredValue.ToString() Then
				If CInt(value) < 1 Then
					Return New ValidationResult(String.Format(ErrorMessageString, context.DisplayName))
				End If
			End If
			Return ValidationResult.Success

		End Function
	End Class

End Namespace

