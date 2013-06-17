Imports System.Web.Configuration
Imports System.Configuration
Imports System.DirectoryServices

Public Class WebSiteInstaller

   Public Sub New()
      MyBase.New()

      'This call is required by the Component Designer.
      InitializeComponent()

      'Add initialization code after the call to InitializeComponent

   End Sub

   Public Overrides Sub Install(ByVal stateSaver As IDictionary)

      Dim targetVDir As String, targetSite As String, path As String, iis As Object

      targetVDir = Me.Context.Parameters("VDir")
      targetSite = Me.Context.Parameters("Site").Replace("/LM/", "/")

      path = "IIS://" & Environment.MachineName & targetSite & "/ROOT/" & targetVDir

      iis = GetIISObject(path)

      'Setup correct authentication
      iis.AuthAnonymous = True
      iis.AuthBasic = False
      iis.AuthMD5 = False
      iis.AuthPassport = False
      iis.AuthNTLM = False
      iis.SetInfo()

      'Retrieve the "Friendly Site Name" from IIS for TargetSite
      Dim entry As DirectoryEntry = New DirectoryEntry("IIS://" & Environment.MachineName & targetSite)
      Dim friendlySiteName As String = entry.Properties("ServerComment").Value.ToString()

      'Open the web.Config            
      Dim config As Configuration = WebConfigurationManager.OpenWebConfiguration("/" + targetVDir, friendlySiteName)

      'Switch of debug
      Dim compilation As CompilationSection = config.GetSection("system.web/compilation")

      If compilation IsNot Nothing Then
         compilation.Debug = False
      End If

      config.Save()

   End Sub

   Private Function GetIISObject(ByVal strFullObjectPath As String) As Object

      Dim iisObject As Object = Nothing

      Try
         iisObject = GetObject(strFullObjectPath)
      Catch exp As Exception
         Err.Raise(9999, "GetIISObject", "Error opening: " & strFullObjectPath & ". " & exp.Message)
      End Try
      Return iisObject
   End Function

End Class
