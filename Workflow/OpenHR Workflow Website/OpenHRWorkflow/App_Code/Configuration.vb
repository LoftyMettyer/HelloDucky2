
Public Class Configuration

  Shared Sub New()

    Server = "COA14270"
    Database = "OpenHR50"
    Login = "sa"
    Password = "asr"
    WorkflowUrl = ConfigurationManager.AppSettings("WorkflowURL").Trim
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
  Public Shared SubmissionTimeoutInSeconds As Integer

  Public Shared Function ConnectionStringFor(user As String, password As String) As String
    Return String.Format( _
        "Application Name=OpenHR Mobile;Data Source={0};Initial Catalog={1};Integrated Security=false;User ID={2};Password={3};Pooling=false", _
        Server, Database, user, password)
  End Function

End Class
