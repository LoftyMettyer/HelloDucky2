Option Strict On
Option Explicit On

Namespace Models
	Public Class UserModel

		Public Property UserName() As String
		Public Property DeviceBrowser As String
		Public Property WebArea As WebArea

		Public ReadOnly Property WebAreaName As String
			Get

				Select Case WebArea
					Case WebArea.DMI
						Return "OpenHR Web"

					Case Else
						Return "Self-service"

				End Select

			End Get
		End Property

	End Class
End Namespace