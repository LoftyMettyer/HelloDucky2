Imports System
Imports System.Xml
Imports System.Data.SqlClient

Public Class Config

	Private Property ThemeFile As String
	Private Property CustomConfigFile As String
	Public Property Server As String
	Public Property Database As String
	Public Property Login As String
	Public Property Password As String
	Public Property WorkflowUrl As String
	Public Property LookupRowsRange As Integer
	Public Property MessageFontSize As Integer
	Public Property ValidationMessageFontSize As Integer
	Public Property OleFolderServer As String
	Public Property OleFolderLocal As String
	Public Property PhotographFolder As String
	Public Property ColourThemeFolder As String
	Public Property ColourThemeForeColour As String
	Public Property ColourThemeHex As String
	Public Property TabletBackColour As String
	Public Property DefaultActiveDirectoryServer As String
	Public Property ConnectionString As String
	''' <summary>
	''' Timeout measured in milliseconds
	''' </summary>
	Public Property SubmissionTimeout As Integer
	Public Property SubmissionTimeoutInSeconds As Integer

	Public Sub New(customConfigFile As String, themeFile As String)
		Me.CustomConfigFile = customConfigFile
		Me.ThemeFile = themeFile
		Load()
	End Sub

	Private Sub Load()

		ColourThemeFolder = GetSetting("Theme", "Blanco").Trim
		MessageFontSize = GetSetting("MessageFontSize", 10)
		ValidationMessageFontSize = GetSetting("ValidationMessageFontSize", 8)
		OleFolderServer = GetSetting("OLEFolder_Server", "").Trim
		OleFolderLocal = GetSetting("OLEFolder_Local", "").Trim
		PhotographFolder = GetSetting("PhotographFolder", "").Trim
		SubmissionTimeoutInSeconds = GetSetting("SubmissionTimeout", 120)
		SubmissionTimeout = SubmissionTimeoutInSeconds * 1000
		LookupRowsRange = GetSetting("LookupRowsRange", 100)
		TabletBackColour = GetSetting("TabletBackColour", "lightgray")
		DefaultActiveDirectoryServer = GetSetting("DefaultActiveDirectoryServer", "")

		' Retrieve connection string from web config.
		Try
			Dim builder As New SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

			Login = builder.UserID
			Password = builder.Password
			Server = builder.DataSource
			Database = builder.InitialCatalog
			builder.ApplicationName = "OpenHR Mobile"

			ConnectionString = builder.ToString()

		Catch ex As Exception
			ConnectionString = ""
		End Try

		'Read the Hex and Foreground values for the defined theme.
		Try
			Dim xmlReader As New XmlTextReader(ThemeFile)

			Do While (xmlReader.ReadToFollowing("theme"))
				If xmlReader.ReadToFollowing("name") Then
					If (xmlReader.Read()) Then
						If (xmlReader.Value.Trim.ToUpper = ColourThemeFolder.Trim.ToUpper) Then
							If xmlReader.ReadToFollowing("hex") Then
								If (xmlReader.Read()) Then
									ColourThemeHex = "#" & xmlReader.Value.Trim.ToUpper

									If xmlReader.ReadToFollowing("forecolour") Then
										If (xmlReader.Read()) Then
											ColourThemeForeColour = xmlReader.Value.Trim
										End If
									End If
								End If
							End If

							Exit Do
						End If
					End If
				End If
			Loop
			xmlReader.Close()
		Catch ex As Exception
			ColourThemeHex = "#FFF"
			ColourThemeForeColour = "black"
		End Try

		'Insert some fake value into the cache with a dependency on the theme & web.custom.config files
		'when they change we'll get a callback to reload the settings
		HttpRuntime.Cache.Insert("filesTheSame", True,
					New CacheDependency(New String() {CustomConfigFile, ThemeFile}),
					Cache.NoAbsoluteExpiration, Cache.NoSlidingExpiration, CacheItemPriority.Default, Sub() Load()
		)

	End Sub

    Public Shared Function GetSetting(Of T)(name As String, defaultValue As T) As T

        Dim value As String = ConfigurationManager.AppSettings(name)
        If value Is Nothing Then
            Return defaultValue
        End If
        Try
            Return CType(Convert.ChangeType(value, GetType(T)), T)
        Catch ex As Exception
            Return defaultValue
        End Try
    End Function

End Class
