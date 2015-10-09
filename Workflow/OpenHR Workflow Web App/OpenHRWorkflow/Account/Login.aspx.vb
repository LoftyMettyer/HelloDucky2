
Partial Class Login
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    Forms.RedirectToHomeIfAuthentcated()

    Title = Utilities.GetPageTitle("Login")
    Forms.LoadControlData(Me, 1)
    Form.DefaultButton = btnLogin2.UniqueID
    Form.DefaultFocus = txtUserName.ClientID
  End Sub

  Protected Sub BtnLoginClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnLogin.Click, btnLogin2.Click

        Dim authenticateOnly As Boolean = (Request.QueryString.Count > 0)

        Dim message As String = Security.ValidateUser(txtUserName.Text.Trim, txtPassword.Text, authenticateOnly)

        If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Login Failed", message)
    Else
      FormsAuthentication.RedirectFromLoginPage(txtUserName.Text.Trim, chkRememberPwd.Checked)
    End If

  End Sub

End Class
