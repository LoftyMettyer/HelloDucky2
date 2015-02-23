Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Security

Namespace Extensions

	<HideModuleName()> _
	Public Module StringExtensions

		<Extension> _
		Public Function ReplaceMultiple(s As String, separators As Char(), newVal As String) As String
			Dim temp As String()

			temp = s.Split(separators, StringSplitOptions.RemoveEmptyEntries)
			Return [String].Join(newVal, temp)
		End Function

		<Extension> _
		Public Function ToSecureString(Source As String) As SecureString
			If String.IsNullOrWhiteSpace(Source) Then
				Return New SecureString()
			Else
				Dim Result As New SecureString()
				For Each c As Char In Source.ToCharArray()
					Result.AppendChar(c)
				Next
				Return Result
			End If
		End Function

		<Extension> _
		Public Function RemoveSensitive(Source As String) As String
			Return Source
		End Function

		<Extension> _
		Public Function GetDirectoryNameOnly(filename As String) As String

			Dim sPath = Path.GetDirectoryName(filename)
			Return Mid(sPath, Len(Path.GetPathRoot(filename)))

		End Function

	End Module

End Namespace