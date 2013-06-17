Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Public Class Database

  Public Shared Function IsSystemLocked() As Boolean

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()
      ' Check if the database is locked.
      Dim cmd = New SqlCommand
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

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()

      Dim cmd As New SqlCommand
      cmd.CommandText = "spASRSysMobileCheckLogin"
      cmd.Connection = conn
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psKeyParameter").Value = userName

      cmd.Parameters.Add("@piUserGroupID", SqlDbType.Int).Direction = ParameterDirection.Output

      cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

      cmd.ExecuteNonQuery()

      Dim result As CheckLoginResult
      result.InvalidReason = NullSafeString(cmd.Parameters("@psMessage").Value())
      result.UserGroupID = NullSafeInteger(cmd.Parameters("@piUserGroupID").Value())
      result.Valid = (result.InvalidReason = Nothing)
      Return result
    End Using

  End Function

End Class

Public Structure CheckLoginResult
  Public Valid As Boolean
  Public InvalidReason As String
  Public UserGroupID As Integer
End Structure