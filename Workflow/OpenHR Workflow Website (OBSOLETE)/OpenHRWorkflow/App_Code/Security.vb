Imports System.Data.SqlClient
Imports System.DirectoryServices

Public Class Security

  Public Shared Function ValidateUser(userName As String, password As String) As String

    Const invalidLoginDetails As String = "The system could not log you on. Make sure your details are correct, then retype your password."

    If userName.IndexOf("\") > 0 Then
      If Not ValidateActiveDirectoryUser(userName.Split("\"c)(0), userName.Split("\"c)(1), password) Then Return invalidLoginDetails
    Else
      If Not ValidateSqlServerUser(userName, password) Then Return invalidLoginDetails
    End If

    Dim db As New Database
    Dim result As CheckLoginResult = db.CheckLoginDetails(userName)
    If Not result.Valid Then
      If result.InvalidReason.ToLower() Like "*incorrect*e-mail*password*" Then Return invalidLoginDetails
      Return result.InvalidReason
    End If

    Return String.Empty

  End Function

  ''' <summary>
  ''' code from http://msdn.microsoft.com/en-us/library/ms180890%28v=vs.90%29.aspx
  ''' </summary>
  Public Shared Function ValidateActiveDirectoryUser(domainName As String, userName As String, password As String) As Boolean

    Dim path As String = "LDAP://" & Configuration.DefaultActiveDirectoryServer

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
      Using conn As New SqlConnection(Configuration.ConnectionStringFor(userName, password))
        conn.Open()
      End Using
      Return True
    Catch ex As Exception
      Return False
    End Try

  End Function

End Class
