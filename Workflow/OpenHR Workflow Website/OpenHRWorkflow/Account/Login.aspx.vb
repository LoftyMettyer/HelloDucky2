
Partial Class Login
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    'Go to the home page if already logged in
    If Request.IsAuthenticated Then
      Response.Redirect("~/Home.aspx")
      Return
    End If

    Title = Utilities.WebSiteName("Login")
    Forms.LoadControlData(Me, 1)
    Form.DefaultButton = btnLogin.UniqueID
    Form.DefaultFocus = txtUserName.ClientID
  End Sub

  Protected Sub BtnLoginClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnLogin.Click

    ' Check if the system is locked
    If Database.IsSystemLocked() Then
      'TODO
      'message = "Database locked." & vbCrLf & "Contact your system administrator."
    End If

    Dim message As String = Security.ValidateUser(txtUserName.Text.Trim, txtPassword.Text)

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Login Failed", message)
    Else
      FormsAuthentication.RedirectFromLoginPage(txtUserName.Text.Trim, chkRememberPwd.Checked)
    End If

  End Sub

End Class
