Imports System.Data.SqlClient
Imports System.DirectoryServices
Imports OpenHRWorkflow.Code.Classes

Public Class Security

  Public Shared Function GetStepDictionary() As Dictionary(Of Integer, StepAuthorization)

    If HttpContext.Current.Session("AuthenticationStepDictionary") Is Nothing Then
      HttpContext.Current.Session("AuthenticationStepDictionary") = New Dictionary(Of Integer, StepAuthorization)
    End If

    Return CType(HttpContext.Current.Session("AuthenticationStepDictionary"), Dictionary(Of Integer, StepAuthorization))

  End Function

    Public Shared Function ValidateUser(userName As String, password As String, authenticateOnly As Boolean) As String

        Const invalidLoginDetails As String = "The system could not log you on. Make sure your details are correct, then retype your password."

        If userName.IndexOf("\") > 0 Then
            If Not ValidateActiveDirectoryUser(userName.Split("\"c)(0), userName.Split("\"c)(1), password) Then Return invalidLoginDetails
        Else
            If Not ValidateSqlServerUser(userName, password) Then Return invalidLoginDetails
        End If

        If Not authenticateOnly Then
            Dim db As New Database(App.Config.ConnectionString)
            Dim result As CheckLoginResult = db.CheckLoginDetails(userName)
            If Not result.Valid Then
                If result.InvalidReason.ToLower() Like "*incorrect*e-mail*password*" Then Return invalidLoginDetails
                Return result.InvalidReason
            End If
        End If

        Return String.Empty

    End Function

    ''' <summary>
    ''' code from http://msdn.microsoft.com/en-us/library/ms180890%28v=vs.90%29.aspx
    ''' </summary>
    Public Shared Function ValidateActiveDirectoryUser(domainName As String, userName As String, password As String) As Boolean

      Dim path As String = "LDAP://" & App.Config.DefaultActiveDirectoryServer

      Dim domainAndUsername As String = domainName & "\" & userName

      Dim entry As New DirectoryEntry(path, domainAndUsername, password)

      Try
         ' Bind to the native AdsObject to force authentication.
         Dim obj As Object = entry.NativeObject

         Dim search As New DirectorySearcher(entry)
         search.Filter = "(SAMAccountName=" & userName & ")"
         search.PropertiesToLoad.Add("cn")
         Dim result As SearchResult = search.FindOne()

         If result Is Nothing Then
            Return False
         End If

      Catch ex As System.Runtime.InteropServices.COMException

         If ex.ErrorCode = -2147023570 Then
            Return False
         Else
            Throw
         End If
      End Try

      Return True

   End Function

   Public Shared Function ValidateSqlServerUser(userName As String, password As String) As Boolean

      Try
         Dim connString = String.Format("Application Name=OpenHR Workflow;Data Source={0};Initial Catalog={1};Integrated Security=false;User ID={2};Password={3};Pooling=false", App.Config.Server, App.Config.Database, userName, password)

         Using conn As New SqlConnection(connString)
            conn.Open()
         End Using
         Return True
      Catch ex As Exception
         Return False
      End Try

   End Function

End Class
