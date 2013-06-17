Imports System
Imports System.Xml

Public Class Config

   Private Property ThemeFile As String
   Private Property CustomConfigFile As String
   Public Property MobileKey As String
   Public Property Server As String
   Public Property Database As String
   Public Property Login As String
   Public Property Password As String
   Public Property WorkflowUrl As String
   Public Property LookupRowsRange As Integer
   Public Property MessageFontSize As Integer
   Public Property ValidationMessageFontSize As Integer
   Public Property OLEFolderServer As String
   Public Property OLEFolderLocal As String
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

      MobileKey = GetSetting("MobileKey", "")
      WorkflowUrl = GetSetting("WorkflowURL", "")
      ColourThemeFolder = GetSetting("Theme", "Blanco").Trim
      MessageFontSize = GetSetting("MessageFontSize", 10)
      ValidationMessageFontSize = GetSetting("ValidationMessageFontSize", 8)
      OLEFolderServer = GetSetting("OLEFolder_Server", "").Trim
      OLEFolderLocal = GetSetting("OLEFolder_Local", "").Trim
      PhotographFolder = GetSetting("PhotographFolder", "").Trim
      SubmissionTimeoutInSeconds = GetSetting("SubmissionTimeout", 120)
      SubmissionTimeout = SubmissionTimeoutInSeconds * 1000
      LookupRowsRange = GetSetting("LookupRowsRange", 100)
      TabletBackColour = GetSetting("TabletBackColour", "lightgray")
      DefaultActiveDirectoryServer = GetSetting("DefaultActiveDirectoryServer", "")

      'Split the mobile key down
      Try
         Dim crypt As New Crypt
         Dim value = crypt.DecompactString(MobileKey)
         value = crypt.DecryptString(value, "", True)

         Dim values As String() = value.Split(ControlChars.Tab)

         Login = values(2)
         Password = values(3)
         Server = values(4)
         Database = values(5)

      Catch ex As Exception
         Login = ""
         Password = ""
         Server = ""
         Database = ""
      End Try
      ConnectionString = String.Format("Application Name=OpenHR Mobile;Data Source={0};Initial Catalog={1};Integrated Security=false;User ID={2};Password={3};Pooling=true", Server, Database, Login, Password)

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
      'TODO PG change timeouts
      HttpRuntime.Cache.Insert("filesTheSame", True,
                               New CacheDependency(New String() {CustomConfigFile, ThemeFile}),
                               DateTime.UtcNow.AddMinutes(1),
                               TimeSpan.Zero,
                               CacheItemPriority.Default,
                               Sub() Load()
      )

   End Sub

   Private Function GetSetting(Of T)(name As String, defaultValue As T) As T

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

   'TODO PG NOW cleanup all connection string creation (take care with application name & pooling)
   Public Function ConnectionStringFor(user As String, password As String) As String
      Return String.Format("Application Name=OpenHR Mobile;Data Source={0};Initial Catalog={1};Integrated Security=false;User ID={2};Password={3};Pooling=false", Server, Database, user, password)
   End Function

End Class
