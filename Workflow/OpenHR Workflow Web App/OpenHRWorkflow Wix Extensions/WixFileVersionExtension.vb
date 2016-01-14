Imports Microsoft.Tools.WindowsInstallerXml
Public Class WixFileVersionExtension
	Inherits WixExtension

	Private _versionPreprocessorExtension As WixFileVersionPreprocessorExtension

	Public Overrides ReadOnly Property PreprocessorExtension As PreprocessorExtension
		Get
			'If we haven't create the preprocessor then do it now
			If _versionPreprocessorExtension Is Nothing Then
				_versionPreprocessorExtension = New WixFileVersionPreprocessorExtension
			End If
			Return _versionPreprocessorExtension
		End Get
	End Property
End Class