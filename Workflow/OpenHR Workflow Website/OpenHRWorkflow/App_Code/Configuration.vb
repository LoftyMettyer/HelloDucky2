﻿Imports Microsoft.VisualBasic
Imports System

Public Class Configuration

  Shared Sub New()

    MobileKey = GetAppSetting("MobileKey", "")
    WorkflowUrl = GetAppSetting("WorkflowURL", "")
    TabletBackColour = GetAppSetting("TabletBackColour", "lightgray")
    DefaultActiveDirectoryServer = GetAppSetting("DefaultActiveDirectoryServer", "")
    SubmissionTimeoutInSeconds = 120

    ConnectionString = String.Format( _
        "Application Name=OpenHR Mobile;Data Source={0};Initial Catalog={1};Integrated Security=false;User ID={2};Password={3}", _
        Server, Database, Login, Password)
  End Sub

  Public Shared Server As String
  Public Shared Database As String
  Public Shared Login As String
  Public Shared Password As String
  Public Shared ConnectionString As String
  Public Shared WorkflowUrl As String
  Public Shared TabletBackColour As String
  Public Shared DefaultActiveDirectoryServer As String
  Public Shared SubmissionTimeoutInSeconds As Integer

  Public Shared Function ConnectionStringFor(user As String, password As String) As String
    Return String.Format( _
        "Application Name=OpenHR Mobile;Data Source={0};Initial Catalog={1};Integrated Security=false;User ID={2};Password={3};Pooling=false", _
        Server, Database, user, password)
  End Function

  Private Shared WriteOnly Property MobileKey() As String
    Set(value As String)

      Try
        Dim crypt As New Crypt
        value = crypt.DecompactString(value)
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

    End Set
  End Property

  Private Shared Function GetAppSetting(name As String, defaultValue As String) As String

    Dim value As String = ConfigurationManager.AppSettings(name)
    If value Is Nothing Then
      Return defaultValue
    End If
    Return value
  End Function

End Class

Public Structure CheckLoginResult
  Public Valid As Boolean
  Public InvalidReason As String
  Public UserGroupID As Integer
End Structure