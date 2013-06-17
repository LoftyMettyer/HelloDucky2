Imports System.Data
Imports System.Data.SqlClient

Partial Class ChangePassword
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 4)
  End Sub

  Protected Sub BtnSubmitClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click

    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""
    Dim userSessionCount As Integer

    Try
      If sMessage.Length = 0 Then
        Using conn As New SqlConnection(Configuration.ConnectionString)
          conn.Open()

          ' Force password change only if there are no other Security logged in with the same name.
          Dim cmd As New SqlCommand("spASRGetCurrentUsersCountOnServer", conn)
          cmd.CommandType = CommandType.StoredProcedure

          cmd.Parameters.Add("@iLoginCount", SqlDbType.Int).Direction = ParameterDirection.Output

          cmd.Parameters.Add("@psLoginName", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          cmd.Parameters("@psLoginName").Value = User.Identity.Name.ToString()

          cmd.ExecuteNonQuery()

          userSessionCount = CInt(cmd.Parameters("@iLoginCount").Value)

          cmd.Dispose()
        End Using

        ' is OK?
        If userSessionCount < 2 Then
          ' Read the Password details from the Password form.
          Dim sCurrentPassword As String = txtCurrPassword.Value
          Dim sNewPassword As String = txtNewPassword.Value

          ' Attempt to change the password on the SQL Server.
          Using conn As New SqlConnection(Configuration.ConnectionString)
            conn.Open()

            Dim cmd As New SqlCommand("sp_password", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.Add("@old", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
            cmd.Parameters("@old").Value = If(sCurrentPassword.Length > 0, sCurrentPassword, vbNullString)

            cmd.Parameters.Add("@new", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
            cmd.Parameters("@new").Value = If(sNewPassword.Length > 0, sNewPassword, vbNullString)

            cmd.Parameters.Add("@loginame", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
            cmd.Parameters("@loginame").Value = User.Identity.Name.ToString()

            cmd.ExecuteNonQuery()
          End Using

        Else
          sMessage = "You could not change your password. The account is currently being used by "
          If userSessionCount > 2 Then
            sMessage &= userSessionCount.ToString & " Security"
          Else
            sMessage &= " another user"
          End If
          sMessage &= " in the system."
        End If

        If sMessage.Length = 0 Then
          ' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
          Using conn As New SqlConnection(Configuration.ConnectionString)
            conn.Open()

            Dim cmd As New SqlCommand("spASRSysMobilePasswordOK", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.Add("@sCurrentUser", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
            cmd.Parameters("@sCurrentUser").Value = User.Identity.Name.ToString()

            cmd.ExecuteNonQuery()
          End Using
          
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

    CType(Master, Site).ShowDialog(sHeader, sMessage, sRedirectTo)

  End Sub

  Protected Sub BtnCancelClick(sender As Object, e As ImageClickEventArgs) Handles btnCancel.Click
    Response.Redirect("~/Home.aspx")
  End Sub

End Class
