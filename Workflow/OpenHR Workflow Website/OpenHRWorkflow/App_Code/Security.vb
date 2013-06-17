Imports System.Data.SqlClient
Imports System.DirectoryServices

Public Class Security

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
