Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.SqlServer.Server
Imports System.Transactions
Imports Assembly.General

Partial Public Class SQLFunctions
  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetCountCurrentUsersInApp", DataAccess:=DataAccessKind.Read)> _
  Public Shared Function CountCurrentUsersInApp(ByVal appName As String) As SqlInt32
    Dim userName As String = String.Empty
    Dim password As String = String.Empty
    Dim databaseName As String = String.Empty
    Dim serverName As String = String.Empty
    Dim connectString As String = String.Empty

    Dim usrCount As Integer = 0

    Dim systemLogon As String = GetSystemLogon()

    If systemLogon = String.Empty Then
      connectString = GetConnectionString("", "", ContextDatabaseName, ContextServerName)
    Else
      ' NPG20081120 Fault 13422
      systemLogon = ProcessEncryptedString(systemLogon)
      DecryptLogonDetails(systemLogon, userName, password, databaseName, serverName)
      connectString = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName)
    End If

    'AE20081002 Fault #13387
    'If systemLogon = String.Empty Then
    '  Throw New ArgumentNullException("System logon details cannot be null")
    'End If

    'DecryptLogonDetails(systemLogon, userName, password, databaseName, serverName)

    'Dim connectString As String = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName)

    Try
      Using conn As New SqlConnection(connectString)
        Dim sql As String = String.Empty

        Dim cmd As New SqlCommand()
        cmd.Connection = conn
        cmd.Connection.Open()
        cmd.CommandType = CommandType.Text

        sql = "SELECT COUNT(p.program_name) " & _
              "FROM     master..sysprocesses p " & _
              "JOIN     master..sysdatabases d " & _
              "         ON d.dbid = p.dbid " & _
              "WHERE    p.program_name = @appName " & _
              " AND     d.name = db_name() " & _
              "GROUP BY p.program_name"
        ' AE20080422 Fault #13118
        cmd.Parameters.AddWithValue("@appName", appName)
        cmd.CommandText = sql

        usrCount = CType(cmd.ExecuteScalar(), Integer)

        cmd.Parameters.Clear()
        cmd.Connection.Close()
        cmd.Dispose()
        cmd = Nothing
      End Using
    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try

    Return New SqlInt32(usrCount)

  End Function

  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetCountCurrentLogins", DataAccess:=DataAccessKind.Read)> _
  Public Shared Function CountCurrentLogins(ByVal loginToCheck As String) As SqlInt32
    Dim userName As String = String.Empty
    Dim password As String = String.Empty
    Dim databaseName As String = String.Empty
    Dim serverName As String = String.Empty
    Dim connectString As String = String.Empty

    Dim loginCount As Integer = 0

    Dim systemLogon As String = GetSystemLogon()

    If systemLogon = String.Empty Then
      connectString = GetConnectionString("", "", ContextDatabaseName, ContextServerName)
    Else
      ' NPG20081120 Fault 13422
      systemLogon = ProcessEncryptedString(systemLogon)
      DecryptLogonDetails(systemLogon, userName, password, databaseName, serverName)
      connectString = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName)
    End If

    loginCount = internal_CountCurrentLogins(connectString, loginToCheck)
    Return New SqlInt32(loginCount)

  End Function

  Public Shared Function internal_CountCurrentLogins(ByVal connectString As String, ByVal loginToCheck As String) As Integer

    Try
      Using conn As New SqlConnection(connectString)
        Dim sql As String = String.Empty

        Dim cmd As New SqlCommand()
        cmd.Connection = conn
        cmd.Connection.Open()

        cmd.CommandType = CommandType.Text
        cmd.CommandTimeout = 5

        ' AE20071119 Fault #12613
        ' Exclue OpenHR Outlook Calendar from running apps
        sql = "SELECT COUNT(*) FROM master..sysprocesses p" & _
            " WHERE    p.program_name LIKE 'OpenHR%'" & _
                    " AND p.program_name NOT LIKE 'OpenHR Workflow%'" & _
                    " AND p.program_name NOT LIKE 'OpenHR Outlook%'" & _
                    " AND p.program_name NOT LIKE 'OpenHR Server.Net%'" & _
                    " AND p.program_name NOT LIKE 'OpenHR Intranet Embedding%'" & _
                    " AND p.loginame = @loginToCheck"
        ' AE20080422 Fault #13118
        cmd.Parameters.AddWithValue("@loginToCheck", loginToCheck)
        cmd.CommandText = sql

        internal_CountCurrentLogins = CType(cmd.ExecuteScalar(), Integer)

        cmd.Parameters.Clear()
        cmd.Connection.Close()
        cmd.Dispose()
        cmd = Nothing
      End Using
    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try

  End Function



  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetIsProcessValid", DataAccess:=DataAccessKind.Read)> _
  Public Shared Function IsProcessValid(ByVal Logon As String) As SqlBoolean
    Dim userName As String = String.Empty
    Dim password As String = String.Empty
    Dim databaseName As String = String.Empty
    Dim serverName As String = String.Empty
    Dim permissionOK As Boolean = True
    Dim ProcessedLogon As String = String.Empty

    'Dim systemLogon As String = GetSystemLogon()

    'If systemLogon = String.Empty Then
    '  Throw New ArgumentNullException("System logon details cannot be null")
    'End If

    ' NPG20081111 Fault 13422
    ProcessedLogon = ProcessEncryptedString(Logon)

    ' NPG20081111 Fault 13373
    ' DecryptLogonDetails(Logon, userName, password, databaseName, serverName)
    DecryptLogonDetails(ProcessedLogon, userName, password, databaseName, serverName)

    Dim connectString As String = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName)

    Try
      Using conn As New SqlConnection(connectString)
        Dim cmd As New SqlCommand("SELECT IS_SRVROLEMEMBER('processadmin') AS Permission", conn)
        cmd.Connection.Open()
        cmd.CommandType = CommandType.Text

        Dim result As Integer = CType(cmd.ExecuteScalar(), Integer)

        cmd.Connection.Close()
        cmd = Nothing

        permissionOK = CType(IIf(result = 1, True, False), Boolean)
      End Using

    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try

    Return New SqlBoolean(permissionOK)

  End Function

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRGetCurrentUsersFromAssembly")> _
  Public Shared Sub GetCurrentUsers()

    Dim userName As String = String.Empty
    Dim password As String = String.Empty
    Dim databaseName As String = String.Empty
    Dim serverName As String = String.Empty
    Dim connectString As String = String.Empty

    Dim sqlP As SqlPipe = SqlContext.Pipe()

    Dim systemLogon As String = GetSystemLogon()

    If systemLogon = String.Empty Then
      connectString = GetConnectionString("", "", ContextDatabaseName, ContextServerName)
    Else
      ' NPG20081120 Fault 13422
      systemLogon = ProcessEncryptedString(systemLogon)
      DecryptLogonDetails(systemLogon, userName, password, databaseName, serverName)
      connectString = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName)
    End If

    Try
      Using scope As New TransactionScope(TransactionScopeOption.Suppress)
        Using conn As New SqlConnection(connectString)

          Dim cmd As New SqlCommand("spASRGetCurrentUsersFromMaster", conn)
          cmd.CommandType = CommandType.StoredProcedure
          Try
            cmd.Connection.Open()
          Catch ex As SqlException
            Throw New Exception(String.Format("Cannot connect to database {0} on server {1} ", ContextDatabaseName, ContextServerName))
          End Try

          sqlP.Send(cmd.ExecuteReader())

        End Using
      End Using
    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try

  End Sub

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRIntGetAvailableLoginsFromAssembly")> _
  Public Shared Sub GetAvailableLogins()

    Dim userName As String = String.Empty
    Dim password As String = String.Empty
    Dim databaseName As String = String.Empty
    Dim serverName As String = String.Empty
    Dim connectString As String = String.Empty

    Dim sqlP As SqlPipe = SqlContext.Pipe()

    Dim systemLogon As String = GetSystemLogon()

    If systemLogon = String.Empty Then
      connectString = GetConnectionString("", "", ContextDatabaseName, ContextServerName)
    Else
      ' NPG20081120 Fault 13422
      systemLogon = ProcessEncryptedString(systemLogon)
      DecryptLogonDetails(systemLogon, userName, password, databaseName, serverName)
      connectString = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName)
    End If

    Try
      Using scope As New TransactionScope(TransactionScopeOption.Suppress)
        Using conn As New SqlConnection(connectString)

          Dim cmd As New SqlCommand("sp_ASRIntGetAvailableLogins", conn)
          cmd.CommandType = CommandType.StoredProcedure
          Try
            cmd.Connection.Open()
          Catch ex As SqlException
            Throw New Exception(String.Format("Cannot connect to database {0} on server {1} ", ContextDatabaseName, ContextServerName))
          End Try

          sqlP.Send(cmd.ExecuteReader())

        End Using
      End Using
    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try

  End Sub

End Class