Imports System.Data
Imports Utilities

Public Class Database

  Public Shared Function IsSystemLocked() As Boolean

    Using conn As New SqlClient.SqlConnection(Configuration.ConnectionString)

      conn.Open()
      ' Check if the database is locked.
      Dim cmd = New SqlClient.SqlCommand
      cmd.CommandText = "sp_ASRLockCheck"
      cmd.Connection = conn
      cmd.CommandType = CommandType.StoredProcedure
      cmd.CommandTimeout = Configuration.SubmissionTimeoutInSeconds

      Dim dr = cmd.ExecuteReader()

      While dr.Read
        ' Not a read-only lock.
        If NullSafeInteger(dr("priority")) <> 3 Then Return True
      End While

      Return False
    End Using

  End Function

  Public Shared Function CheckLoginDetails(userName As String) As CheckLoginResult

    Dim result As CheckLoginResult

    Using conn As New SqlClient.SqlConnection(Configuration.ConnectionString)

      conn.Open()

      Dim cmd As New SqlClient.SqlCommand
      cmd.CommandText = "spASRSysMobileCheckLogin"
      cmd.Connection = conn
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psKeyParameter").Value = userName

      cmd.Parameters.Add("@piUserGroupID", SqlDbType.Int).Direction = ParameterDirection.Output

      cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

      cmd.ExecuteNonQuery()

      result.InvalidReason = NullSafeString(cmd.Parameters("@psMessage").Value())
      result.UserGroupID = NullSafeInteger(cmd.Parameters("@piUserGroupID").Value())
      result.Valid = (result.InvalidReason = Nothing)
    End Using

    Return result

  End Function

End Class
