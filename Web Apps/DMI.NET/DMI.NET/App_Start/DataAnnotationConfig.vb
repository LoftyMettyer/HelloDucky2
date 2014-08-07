Option Strict On
Option Explicit On

Imports DMI.NET.Code.Attributes

Namespace App_Start

	Public Class DataAnnotationConfig

		' Used so that client side validation can use our extra attributes
		Public Shared Sub RegisterDataAnnotations()

			DataAnnotationsModelValidatorProvider.RegisterAdapter(GetType(RequiredIfAttribute), GetType(RequiredAttributeAdapter))
			DataAnnotationsModelValidatorProvider.RegisterAdapter(GetType(NonZeroIfAttribute), GetType(DataAnnotationsModelValidator))

		End Sub

	End Class
End Namespace
