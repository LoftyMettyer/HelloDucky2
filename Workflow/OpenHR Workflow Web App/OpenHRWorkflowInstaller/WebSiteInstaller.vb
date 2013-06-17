Imports System.Web.Configuration
Imports System.Configuration
Imports System.DirectoryServices
Imports System.Configuration.Install

Public Class WebSiteInstaller

	Const AppPoolName = "OpenHR v4.0"

   Public Sub New()
      MyBase.New()

      'This call is required by the Component Designer.
      InitializeComponent()

      'Add initialization code after the call to InitializeComponent

   End Sub

   Public Overrides Sub Install(ByVal stateSaver As IDictionary)

		Dim targetVDir As String, targetSite As String, metaSitePath As String

      targetVDir = Me.Context.Parameters("VDir")
      targetSite = Me.Context.Parameters("Site").Replace("/LM/", "/")

      metaSitePath = "IIS://" & Environment.MachineName & targetSite & "/ROOT/" & targetVDir

		'Setup correct authentication
		Dim iis As Object
		iis = GetIISObject(metaSitePath)
      iis.AuthAnonymous = True
      iis.AuthBasic = False
      iis.AuthMD5 = False
      iis.AuthPassport = False
      iis.AuthNTLM = False
      iis.SetInfo()

		Dim metaAppPoolPath As String = "IIS://" & System.Environment.MachineName & "/W3SVC/AppPools"

		If Not AppPoolExists(metaAppPoolPath, AppPoolName) Then
			CreateAppPool(metaAppPoolPath, AppPoolName)
		End If

		If AppPoolExists(metaAppPoolPath, AppPoolName) Then
			AssignVDirToAppPool(metaSitePath, AppPoolName)
		End If

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

	Private Function AppPoolExists(ByVal strMetabasePath As String, ByVal strAppPoolName As String) As Boolean
		strMetabasePath &= "/" + strAppPoolName
		Return DirectoryEntry.Exists(strMetabasePath)
	End Function

	Private Sub CreateAppPool(ByVal strMetabasePath As String, ByVal strAppPoolName As String)
		' strMetabasePath is of the form "IIS://<servername>/W3SVC/AppPools"
		'   For example: "IIS://localhost/W3SVC/AppPools" 
		' strAppPoolName is of the form "<name>", for example, "MyAppPool"
		Try
			If strMetabasePath.EndsWith("/W3SVC/AppPools") Then
				Dim appPools As New DirectoryEntry(strMetabasePath)
				Dim newPool = appPools.Children.Add(strAppPoolName, "IIsApplicationPool")

				newPool.Properties("ManagedRuntimeVersion")(0) = "v4.0"
				newPool.Properties("ManagedPipelineMode")(0) = 0
				newPool.Properties("IdleTimeout")(0) = 60

				newPool.CommitChanges()
			Else
				Throw New InstallException("Failed in CreateAppPool; application pools can only be created in the */W3SVC/AppPools node.")
			End If
		Catch exError As Exception
			Throw New InstallException(String.Format("Failed in CreateAppPool with the following exception: {0}", exError.Message))
		End Try

	End Sub

	Private Sub AssignVDirToAppPool(ByVal strMetabasePath As String, ByVal strAppPoolName As String)
		' strMetabasePath is of the form "IIS://<servername>/W3SVC/<siteID>/Root[/<vDir>]"
		'   For example: "IIS://localhost/W3SVC/1/Root/MyVDir" 
		' strAppPoolName is of the form "<name>", for example, "MyAppPool"
		Try
			Dim objVdir As New DirectoryEntry(strMetabasePath)
			Dim strClassName As String = objVdir.SchemaClassName.ToString()
			If strClassName.EndsWith("VirtualDir") Then
				Dim objParam As Object() = {0, strAppPoolName, True}
				objVdir.Invoke("AppCreate3", objParam)
				objVdir.Properties("AppIsolated")(0) = "2"
			Else
				Throw New InstallException("Failed in AssignVDirToAppPool; only virtual directories can be assigned to application pools")
			End If
		Catch exError As Exception
			Throw New InstallException(String.Format("Failed in AssignVDirToAppPool with the following exception: " + vbLf + "{0}", exError.Message))
		End Try

	End Sub

End Class
