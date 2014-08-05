Option Strict On
Option Explicit On

Namespace Classes
	Public Class GroupAccess
		Implements IJsonSerialize

		<HiddenInput>
		Public Property ID As Integer Implements IJsonSerialize.ID

		Public Property Name As String
		Public Property Access As String
		Public Property IsReadOnly As Boolean

	End Class
End Namespace