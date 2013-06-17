Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class ChangePassword
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 4)

    Title = WebSiteName("Change Password")
    Page.Form.DefaultButton = btnSubmitButton.UniqueID
    Page.Form.DefaultFocus = txtCurrPassword.ClientID
  End Sub

  Protected Sub BtnSubmitClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmitButton.Click

    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""
  
    Try
      Dim userSessionCount As Integer

      Using conn As New SqlConnection(Configuration.ConnectionString)
        conn.Open()

        ' Force password change only if there are no other security logged in with the same name.
        Dim cmd As New SqlCommand("spASRGetCurrentUsersCountOnServer", conn)
        cmd.CommandType = CommandType.StoredProcedure

        cmd.Parameters.Add("@iLoginCount", SqlDbType.Int).Direction = ParameterDirection.Output

        cmd.Parameters.Add("@psLoginName", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
        cmd.Parameters("@psLoginName").Value = User.Identity.Name.ToString()

        cmd.ExecuteNonQuery()

        userSessionCount = CInt(cmd.Parameters("@iLoginCount").Value)

        cmd.Dispose()
      End Using

      If userSessionCount > 1 Then
        sMessage = String.Format("You could not change your password. The account is currently being used by {0} in the system.", _
                    If(userSessionCount > 2, userSessionCount.ToString & " security", "another user"))
      End If

      If sMessage.Length = 0 Then
        ' Attempt to change the password on the SQL Server.
        Using conn As New SqlConnection(Configuration.ConnectionString)
          conn.Open()

          Dim cmd As New SqlCommand("sp_password", conn)
          cmd.CommandType = CommandType.StoredProcedure

          cmd.Parameters.Add("@old", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          cmd.Parameters("@old").Value = txtCurrPassword.Text

          cmd.Parameters.Add("@new", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          cmd.Parameters("@new").Value = txtNewPassword.Text

          cmd.Parameters.Add("@loginame", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          cmd.Parameters("@loginame").Value = User.Identity.Name.ToString()

          Try
            cmd.ExecuteNonQuery()
          Catch ex As SqlException
            If ex.Number = 15151 Then
              sMessage = "Current password is incorrect."
            Else
              sMessage = ex.Message
            End If
          End Try
        End Using
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
      End If

    Catch ex As Exception
      sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message.ToString & vbCrLf & vbCrLf & "Contact your system administrator."
    End Try

    If sMessage.Length > 0 Then
      sHeader = "Change Password Failed"
    Else
      sHeader = "Change Password Submitted"
      sMessage = "Password changed successfully."
      sRedirectTo = "../Home.aspx"
    End If

    CType(Master, Site).ShowDialog(sHeader, sMessage, sRedirectTo)

  End Sub

End Class
