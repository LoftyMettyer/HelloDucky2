Imports System.ComponentModel
Imports System.Configuration.Install

Public Class WebSiteInstaller

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub
	Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
		Dim strVDir As String
		Dim strObjectPath As String
		Dim IISVdir As Object
		Dim strSite As String

		strSite = Me.Context.Parameters.Item("Site").ToString.Replace("/LM/", "/")
		strVDir = Me.Context.Parameters.Item("VDir").ToString()
		strObjectPath = "IIS://" & System.Environment.MachineName & strSite & "/ROOT/" & strVDir

		' Gets the IIS VDir Object.
		IISVdir = GetIISObject(strObjectPath)

		' Set the AuthAnonymous property here.
		IISVdir.AuthAnonymous = True
		IISVdir.AuthBasic = True
		'IISVdir.AuthMDS = True
		IISVdir.AuthNTLM = True
		'IISVdir.AuthPassport = True

		'IISVdir.AuthFlags = 5	' AuthNTLM + AuthAnonymous 

		' Uses SetInfo to save the settings to the Server.
		IISVdir.SetInfo()
	End Sub
	Private Function GetIISObject(ByVal strFullObjectPath As String) As Object
    Dim IISObject As Object = Nothing

		Try
      IISObject = GetObject(strFullObjectPath)
    Catch exp As Exception
      Err.Raise(9999, "GetIISObject", "Error opening: " & strFullObjectPath & ". " & exp.Message)
    End Try
    Return IISObject
  End Function

End Class
