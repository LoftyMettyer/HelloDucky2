Option Strict On
Option Explicit On

Namespace Models

	Public Class DefinitionSearchResultModel

		Public Property __RequestVerificationToken As String
		Public Property IsRunAllowed As Boolean
		Public Property Id As Integer
		Public Property Access As String

		<AllowHtml>
		Public Property SearchText As String
		Public Property ReportType() As UtilityType
		Public Property Name As String
		Public Property TextToDisplay As String
		Public Property Description As String

	End Class

End Namespace