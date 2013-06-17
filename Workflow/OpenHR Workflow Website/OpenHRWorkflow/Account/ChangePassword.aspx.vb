Imports System.Data
Imports System.Data.SqlClient

Partial Class ChangePassword
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 4)
  End Sub

  Protected Sub BtnSubmitClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click

    Dim conn As SqlConnection
    Dim cmdCheckUserSessions As SqlCommand
    Dim cmdChangePassword As SqlCommand
    Dim cmdPasswordOk As SqlCommand
    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""

    Try
      If sMessage.Length = 0 Then
        conn = New SqlConnection(Configuration.ConnectionString)
        conn.Open()

        ' Force password change only if there are no other Security logged in with the same name.
        cmdCheckUserSessions = New SqlCommand
        cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
        cmdCheckUserSessions.Connection = conn
        cmdCheckUserSessions.CommandType = CommandType.StoredProcedure

        cmdCheckUserSessions.Parameters.Add("@iLoginCount", SqlDbType.Int).Direction = ParameterDirection.Output

        cmdCheckUserSessions.Parameters.Add("@psLoginName", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
        cmdCheckUserSessions.Parameters("@psLoginName").Value = User.Identity.Name.ToString()

        cmdCheckUserSessions.ExecuteNonQuery()

        Dim iUserSessionCount As Integer = CInt(cmdCheckUserSessions.Parameters("@iLoginCount").Value)

        cmdCheckUserSessions.Dispose()

        ' is OK?
        If iUserSessionCount < 2 Then
          ' Read the Password details from the Password form.
          Dim sCurrentPassword As String = txtCurrPassword.Value
          Dim sNewPassword As String = txtNewPassword.Value

          ' Attempt to change the password on the SQL Server.
          cmdChangePassword = New SqlCommand
          cmdChangePassword.CommandText = "sp_password"
          cmdChangePassword.Connection = conn
          cmdChangePassword.CommandType = CommandType.StoredProcedure

          cmdChangePassword.Parameters.Add("@old", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          If Len(sCurrentPassword) > 0 Then
            cmdChangePassword.Parameters("@old").Value = sCurrentPassword
          Else
            cmdChangePassword.Parameters("@old").Value = vbNullString
          End If

          cmdChangePassword.Parameters.Add("@new", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          If Len(sNewPassword) > 0 Then
            cmdChangePassword.Parameters("@new").Value = sNewPassword
          Else
            cmdChangePassword.Parameters("@new").Value = vbNullString
          End If

          cmdChangePassword.Parameters.Add("@loginame", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          cmdChangePassword.Parameters("@loginame").Value = User.Identity.Name.ToString()

          cmdChangePassword.ExecuteNonQuery()

          cmdChangePassword.Dispose()
        Else
          sMessage = "You could not change your password. The account is currently being used by "
          If iUserSessionCount > 2 Then
            sMessage &= iUserSessionCount.ToString & " Security"
          Else
            sMessage &= " another user"
          End If
          sMessage &= " in the system."
        End If

        If sMessage.Length = 0 Then
          ' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
          cmdPasswordOk = New SqlCommand
          cmdPasswordOk.CommandText = "spASRSysMobilePasswordOK"
          cmdPasswordOk.Connection = conn
          cmdPasswordOk.CommandType = CommandType.StoredProcedure

          cmdPasswordOk.Parameters.Add("@sCurrentUser", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          cmdPasswordOk.Parameters("@sCurrentUser").Value = User.Identity.Name.ToString()

          cmdPasswordOk.ExecuteNonQuery()

          cmdPasswordOk.Dispose()

          ' Tell the user that the password was changed okay.
          sMessage = "Password changed successfully."
        End If
      End If

      sHeader = "Change Password Submitted"

    Catch ex As Exception
      sHeader = "Change Password Failed"
      sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message.ToString & vbCrLf & vbCrLf & "Contact your system administrator."
    End Try


    If sMessage.Length = 0 Then
      sMessage = "Password changed successfully."
      sRedirectTo = "../Home.aspx"
    End If

    CType(Master, Site).ShowMessage(sHeader, sMessage, sRedirectTo)

  End Sub

  Protected Sub BtnCancelClick(sender As Object, e As ImageClickEventArgs) Handles btnCancel.Click
    Response.Redirect("~/Home.aspx")
  End Sub

End Class
