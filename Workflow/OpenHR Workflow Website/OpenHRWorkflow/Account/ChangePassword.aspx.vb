
Partial Class ChangePassword
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
    Title = Utilities.WebSiteName("Change Password")
    Forms.LoadControlData(Me, 4)
    Form.DefaultButton = btnSubmit.UniqueID
    Form.DefaultFocus = txtCurrPassword.ClientID
  End Sub

  Protected Sub BtnSubmitClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click

    Dim message As String = ""

    ' Force password change only if there are no other security logged in with the same name.
    Dim userSessionCount As Integer = Database.GetUserCountOnServer(User.Identity.Name)

    If userSessionCount > 1 Then
      message = String.Format("You could not change your password. The account is currently being used by {0} in the system.", _
                      If(userSessionCount > 2, userSessionCount.ToString & " security", "another user"))
    Else
      ' Change users password
      message = Database.ChangePassword(User.Identity.Name, txtCurrPassword.Text, txtNewPassword.Text)
    End If

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Change Password Failed", message)
    Else
      CType(Master, Site).ShowDialog("Change Password Submitted", "Password changed successfully.", "../Home.aspx")
    End If

  End Sub

End Class
