Imports System.Runtime.CompilerServices

Namespace Extensions

	<HideModuleName()> _
	Friend Module StringExtensions

		<Extension> _
		Public Function ReplaceMultiple(s As String, separators As Char(), newVal As String) As String
			Dim temp As String()

			temp = s.Split(separators, StringSplitOptions.RemoveEmptyEntries)
			Return [String].Join(newVal, temp)
		End Function
	End Module

End Namespace