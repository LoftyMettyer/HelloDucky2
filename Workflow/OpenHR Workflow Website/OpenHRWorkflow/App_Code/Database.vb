Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Public Class Database

  Public Shared Function IsSystemLocked() As Boolean

    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()
      ' Check if the database is locked.
      Dim cmd = New SqlCommand("sp_ASRLockCheck", conn)
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

      Dim cmd As New SqlCommand("spASRSysMobileCheckLogin", conn)
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

  Public Shared Function GetWorkflowPendingStepCount(userName As String) As Integer

    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlClient.SqlCommand
      cmd.CommandText = "spASRSysMobileCheckPendingWorkflowSteps"
      cmd.Connection = conn
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psKeyParameter").Value = userName

      Dim dr As SqlDataReader = cmd.ExecuteReader

      Dim count As Integer
      While dr.Read
        count += 1
      End While
      Return count
    End Using

  End Function

  Public Shared Function CanUserGroupRunWorkflows(userGroupID As Integer) As Boolean

    Using conn As New SqlConnection(Configuration.ConnectionString)

      ' get the run permissions for workflow for this user group.
      Dim sql As String = "SELECT  [i].[itemKey], [p].[permitted]" & _
                           " FROM [ASRSysGroupPermissions] p" & _
                           " JOIN [ASRSysPermissionItems] i ON [p].[itemID] = [i].[itemID]" & _
                           " WHERE [p].[itemID] IN (" & _
                               " SELECT [itemID] FROM [ASRSysPermissionItems]	" & _
                                " WHERE [categoryID] = (SELECT [categoryID] FROM [ASRSysPermissionCategories] WHERE [categoryKey] = 'WORKFLOW')) " & _
                                " AND [groupName] = (SELECT [Name] FROM [ASRSysGroups] WHERE [ID] = " & userGroupID.ToString & ")"

      conn.Open()
      Dim cmd As New SqlCommand(sql, conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()

      While dr.Read()
        Select Case CStr(dr("itemKey"))
          Case "RUN"
            Return CBool(dr("permitted"))
        End Select
      End While

      Return False
    End Using

  End Function

End Class

Public Structure CheckLoginResult
  Public Valid As Boolean
  Public InvalidReason As String
  Public UserGroupID As Integer
End Structure