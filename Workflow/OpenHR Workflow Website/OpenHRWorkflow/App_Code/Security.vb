Imports System.Data.SqlClient

Public Class Security

  Public Shared Function ValidateActiveDirectoryUser(domainName As String, userName As String, password As String) As Boolean

    ' Path to youR LDAP directory server.
    ' Contact your network administrator to obtain a valid path.

    Dim adPath As String = "LDAP://" & Configuration.DefaultActiveDirectoryServer

    Dim adAuth As New ActiveDirectoryValidator(adPath)

    Return adAuth.IsAuthenticated(domainName, userName, password)

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
