Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Code.Attributes

	Public Class RequiredIfAttribute
		Inherits RequiredAttribute

		Private Property PropertyName() As [String]
			Get
				Return m_PropertyName
			End Get
			Set(value As [String])
				m_PropertyName = value
			End Set
		End Property

		Private m_PropertyName As [String]

		Private Property DesiredValue() As [Object]
			Get
				Return m_DesiredValue
			End Get
			Set(value As [Object])
				m_DesiredValue = value
			End Set
		End Property

		Private m_DesiredValue As [Object]

		Public Sub New(propertyName__1 As [String], desiredvalue__2 As [Object])
			PropertyName = propertyName__1
			DesiredValue = desiredvalue__2
		End Sub

		Protected Overrides Function IsValid(value As Object, context As ValidationContext) As ValidationResult
			Dim instance As [Object] = context.ObjectInstance
			Dim type As Type = instance.[GetType]()
			Dim proprtyvalue As [Object] = type.GetProperty(PropertyName).GetValue(instance, Nothing)
			If proprtyvalue.ToString() = DesiredValue.ToString() Then
				Dim result As ValidationResult = MyBase.IsValid(value, context)
				Return result
			End If
			Return ValidationResult.Success
		End Function
	End Class

End Namespace