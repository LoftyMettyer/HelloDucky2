Imports System
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports Microsoft.SqlServer.Server
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices

Public Class Security

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spadmin_createsystemlogin")> _
  Public Shared Sub CreateSystemLogin()

    Dim command As SqlCommand
    Dim sSQL As String
    Dim sUsername As String = "OpenHR2IIS"
    Dim sPassword As String = "H@Rp3Nd3N"

    Try

      Using connection As New SqlConnection("context connection=true")
        connection.Open()

        ' Create the login
        sSQL = String.Format("IF NOT EXISTS (SELECT * FROM sys.server_principals WHERE name = N'{0}')" & _
          "CREATE LOGIN [{0}] WITH PASSWORD = '{1}';" & _
          "EXEC sys.sp_addsrvrolemember @loginame = N'{0}', @rolename = N'securityadmin';" & _
          "EXEC sys.sp_addsrvrolemember @loginame = N'{0}', @rolename = N'serveradmin';" _
          , sUsername, sPassword)
        command = New SqlCommand(sSQL, connection)
        command.ExecuteNonQuery()

        ' Create the user on this db
        sSQL = String.Format("IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'{0}')" & _
          "CREATE USER [{0}] FOR LOGIN [{0}] WITH DEFAULT_SCHEMA=[dbo]", sUsername)
        command = New SqlCommand(sSQL, connection)
        command.ExecuteNonQuery()

      End Using

    Catch ex As Exception
      Throw ex
    End Try

  End Sub

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spadmin_commitresetpassword")> _
  Public Shared Sub CommitResetPassword(ByVal code As String, ByVal NewPassword As String, ByRef ErrorMessage As String)

    Dim command As SqlCommand
    Dim sSQL As String
    Dim sUsername As String
    Dim sUnused As String
    Dim values As String()
    Dim iLoggedIn As Integer
    Dim cmd As New SqlCommand()

    Dim crypt As New Crypt
    code = crypt.DecompactString(code)
    code = crypt.DecryptString(code, "", True)

    'Extract the required parameters from the decrypted queryString.
    values = code.Split(vbTab(0))

    sUsername = values(0)
    sUnused = values(1)     ' Timestamp
    sUnused = values(2)      ' Server 
    sUnused = values(3)    ' Database

    Try

      Using connection As New SqlConnection("context connection=true")

        connection.Open()
        cmd.Connection = connection

        cmd.CommandType = CommandType.Text
        cmd.CommandTimeout = 5

        sSQL = "SELECT COUNT(*) FROM master..sysprocesses p" & _
            " WHERE    p.program_name LIKE 'OpenHR%'" & _
                    " AND p.program_name NOT LIKE 'OpenHR Workflow%'" & _
                    " AND p.program_name NOT LIKE 'OpenHR Outlook%'" & _
                    " AND p.program_name NOT LIKE 'OpenHR Server.Net%'" & _
                    " AND p.program_name NOT LIKE 'OpenHR Intranet Embedding%'" & _
                    " AND p.loginame = @loginToCheck"
        cmd.Parameters.AddWithValue("@loginToCheck", sUsername)
        cmd.CommandText = sSQL

        iLoggedIn = CType(cmd.ExecuteScalar(), Integer)

        If iLoggedIn = 0 Then
          sSQL = String.Format("IF EXISTS (SELECT * FROM sys.server_principals WHERE name = N'{0}')" & _
          "ALTER LOGIN [{0}] WITH PASSWORD = '{1}'", sUsername, NewPassword)

          cmd.CommandText = sSQL
          cmd.ExecuteNonQuery()
          ErrorMessage = "Password changed successfully"
        Else
          ErrorMessage = "User is currently logged in"
        End If

      End Using

    Catch ex As Exception
      ErrorMessage = ex.Message

    Finally

      cmd.Parameters.Clear()
      cmd.Connection.Close()
      cmd.Dispose()
      cmd = Nothing


    End Try

  End Sub

End Class

