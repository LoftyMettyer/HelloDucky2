Imports System.IO
Imports Microsoft.Tools.WindowsInstallerXml
Imports System.Reflection

Public Class WixFileVersionPreprocessorExtension
	Inherits PreprocessorExtension

	Private Shared ReadOnly _prefixes() As String = {"fileVersion"}

	Public Overrides ReadOnly Property Prefixes As String()
		Get
			Return _prefixes
		End Get
	End Property

	Public Overrides Function EvaluateFunction(prefix As String, [function] As String, args As String()) As String
		Dim actualFunction As String = [function]

		If [function] = "MajorAndMinorProductVersion" Or [function] = "FullProductVersion" Then
			actualFunction = "ProductVersion"
		End If

		Select Case prefix
			Case "fileVersion"
				'Make sure there actually is a file name
				If args.Length = 0 OrElse args(0).Length = 0 Then
					Throw New ArgumentException("File name not specified")
				End If

				'Make sure the file exists
				If Not File.Exists(args(0)) Then
					Throw New ArgumentException(String.Format("File name {0} does not exist", args(0)))
				End If

				'Get the file version information for the given file
				Dim fileVersionInformation As FileVersionInfo = FileVersionInfo.GetVersionInfo(args(0))

				'Get the property that matches the name of the function
				Dim propertyInfo As PropertyInfo = fileVersionInformation.[GetType]().GetProperty(actualFunction)

				'Make sure the property exists
				If propertyInfo Is Nothing Then
					Throw New ArgumentException(String.Format("Unable to find property {0} in FileVersionInfo", actualFunction))
				End If

				'Return the value of the property as a string
				Select Case [function]
					Case "MajorAndMinorProductVersion"
						Return fileVersionInformation.FileMajorPart & "." & fileVersionInformation.FileMinorPart
					Case "FullProductVersion"
						Return propertyInfo.GetValue(fileVersionInformation, Nothing).ToString()
					Case Else
						Return Nothing
				End Select
		End Select
		Return Nothing
	End Function
End Class
