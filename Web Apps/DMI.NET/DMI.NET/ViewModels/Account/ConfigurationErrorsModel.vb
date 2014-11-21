Option Strict On
Option Explicit On

Imports DMI.NET.Classes

Namespace ViewModels.Account
	Public Class ConfigurationErrorsModel
		Public Property Errors() As New List(Of ConfigurationError)
	End Class
End Namespace
