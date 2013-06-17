Imports System.Data
Imports System.Linq
Imports System.Data.SqlClient
Imports System.Collections.Generic
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

  Public Shared Function IsMobileLicensed() As Boolean

    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlCommand("spASRIntActivateModule", conn)
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@sModule", SqlDbType.VarChar, 50).Direction = ParameterDirection.Input
      cmd.Parameters("@sModule").Value = "MOBILE"

      cmd.Parameters.Add("@bLicensed", SqlDbType.Bit).Direction = ParameterDirection.Output

      cmd.ExecuteNonQuery()

      Return NullSafeBoolean(cmd.Parameters("@bLicensed").Value())
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

  Public Shared Function GetPendingStepCount(userName As String) As Integer

    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlCommand
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

  Public Shared Function GetUserID(email As String) As Integer

    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlCommand("spASRSysMobileGetUserIDFromEmail", conn)
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@psEmail", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psEmail").Value = email

      cmd.Parameters.Add("@piUserID", SqlDbType.Int).Direction = ParameterDirection.Output

      cmd.ExecuteNonQuery()

      Return NullSafeInteger(cmd.Parameters("@piUserID").Value())

    End Using

  End Function

  Public Shared Function Register(email As String) As String

    'Check the email address relates to a user
    Dim userID = GetUserID(email)

    If userID = 0 Then
      Return "No records exist with the given email address."
    End If

    Dim crypt As New Crypt
    Dim encryptedString As String = crypt.EncryptQueryString((userID), -2, _
        Configuration.Login, _
        Configuration.Password, _
        Configuration.Server, _
        Configuration.Database, _
        "", _
        "")

    Dim activationUrl As String = Configuration.WorkflowUrl & "?" & encryptedString

    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlCommand("spASRSysMobileRegistration", conn)
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psEmailAddress").Value = email

      cmd.Parameters.Add("@psActivationURL", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psActivationURL").Value = activationUrl

      cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

      cmd.ExecuteNonQuery()

      Return CStr(cmd.Parameters("@psMessage").Value())

    End Using

  End Function

  Public Shared Function ForgotLogin(email As String) As String

    'Check the email address relates to a user
    Dim userID = GetUserID(email)

    If userID = 0 Then
      Return "No records exist with the given email address."
    End If

    'Send it all to sql to validate and email out
    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlCommand("spASRSysMobileForgotLogin", conn)
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psEmailAddress").Value = email

      cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

      cmd.ExecuteNonQuery()

      Return CStr(cmd.Parameters("@psMessage").Value())

    End Using

  End Function

  Public Shared Function GetLoginCount(userName As String) As Integer

    'Does not include being logged into the mobile site
    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlCommand("spASRGetCurrentUsersCountOnServer", conn)
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@iLoginCount", SqlDbType.Int).Direction = ParameterDirection.Output

      cmd.Parameters.Add("@psLoginName", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psLoginName").Value = userName

      cmd.ExecuteNonQuery()

      Return CInt(cmd.Parameters("@iLoginCount").Value)

    End Using

  End Function

  Public Shared Function ChangePassword(userName As String, currentPassword As String, newPassword As String) As String

    If GetLoginCount(userName) > 0 Then
      Return "Could not change your password. You are logged into the system using another application."
    End If

    ' Attempt to change the password on the SQL Server.
    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlCommand("sp_password", conn)
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@old", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@old").Value = currentPassword

      cmd.Parameters.Add("@new", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@new").Value = newPassword

      cmd.Parameters.Add("@loginame", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@loginame").Value = userName

      Try
        cmd.ExecuteNonQuery()
      Catch ex As SqlException
        If ex.Number = 15151 Then
          Return "Current password is incorrect."
        Else
          Return ex.Message
        End If
      End Try
    End Using

      ' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
      Using conn As New SqlConnection(Configuration.ConnectionString)
        conn.Open()

        Dim cmd As New SqlCommand("spASRSysMobilePasswordOK", conn)
        cmd.CommandType = CommandType.StoredProcedure

        cmd.Parameters.Add("@sCurrentUser", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
        cmd.Parameters("@sCurrentUser").Value = userName

        cmd.ExecuteNonQuery()
      End Using

      Return String.Empty

  End Function

  Public Shared Function CanRunWorkflows(userGroupID As Integer) As Boolean

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

  Public Shared Function GetWorkflowList(userGroupID As Integer) As List(Of WorkflowLink)

    Dim list As New List(Of WorkflowLink)
  
    Using conn As New SqlConnection(Configuration.ConnectionString)

      Dim sql As String = "SELECT w.Id, w.Name, w.PictureID" & _
            " FROM tbsys_mobilegroupworkflows gw" & _
            " INNER JOIN tbsys_workflows w on gw.WorkflowID = w.ID" & _
            " WHERE gw.UserGroupID = " & userGroupID & " AND w.enabled = 1 ORDER BY gw.Pos ASC"

      conn.Open()
      Dim cmd As New SqlCommand(sql, conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()

      While dr.Read()
        list.Add(New WorkflowLink() With _
                 {.ID = NullSafeInteger(dr("ID")), _
                  .Name = NullSafeString(dr("Name")), _
                  .PictureID = NullSafeInteger(dr("PictureID")) _
                 })
      End While

      Return list
    End Using

  End Function

  Public Shared Function GetPendingStepList(userName As String) As List(Of WorkflowStepLink)

    Dim list As New List(Of WorkflowStepLink)

    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlCommand("spASRSysMobileCheckPendingWorkflowSteps", conn)
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psKeyParameter").Value = userName

      Dim dr As SqlDataReader = cmd.ExecuteReader

      While dr.Read()

        Dim desc As String = CStr(dr("description"))

        If desc.StartsWith(CStr(dr("name")).Trim() & " -") Then
          desc = desc.Substring(CStr(dr("name")).Trim().Length + 2).Trim()
        End If

        list.Add(New WorkflowStepLink() With _
              {.Url = NullSafeString(dr("Url")), _
               .Name = NullSafeString(dr("Name")), _
               .Desc = desc, _
               .PictureID = NullSafeInteger(dr("PictureID")) _
              })
      End While

      Return (From x In list Order By x.Name, x.Desc).ToList()
    End Using

  End Function

End Class

Public Structure CheckLoginResult
  Public Valid As Boolean
  Public InvalidReason As String
  Public UserGroupID As Integer
End Structure

Public Class WorkflowLink
  Public ID As Integer
  Public Name As String
  Public PictureID As Integer
End Class

Public Class WorkflowStepLink
  Public Url As String
  Public Name As String
  Public Desc As String
  Public PictureID As Integer
End Class