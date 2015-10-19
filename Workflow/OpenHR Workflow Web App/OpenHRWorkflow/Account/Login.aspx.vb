
Partial Class Login
    Inherits Page

    Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

        Forms.RedirectToHomeIfAuthentcated()

        Title = Utilities.GetPageTitle("Login")
        Forms.LoadControlData(Me, 1)
        Form.DefaultButton = btnLogin2.UniqueID
        Form.DefaultFocus = txtUserName.ClientID

        If Session("ValidLogins") IsNot Nothing Then
            btnRegister.Visible = False
            btnForgotPwd.Visible = False
        End If

    End Sub

    Protected Sub BtnLoginClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnLogin.Click, btnLogin2.Click

        Dim authenticateOnly As Boolean = (Request.QueryString.Count > 0)
        Dim validLogins = CType(Session("ValidLogins"), List(Of String))
        Dim message As String

        If validLogins IsNot Nothing AndAlso validLogins.Count > 0 AndAlso Not validLogins.Contains(txtUserName.Text.Trim.ToLower) Then
            message = "You are not authorised for this step"
        Else
            message = Security.ValidateUser(txtUserName.Text.Trim, txtPassword.Text, authenticateOnly)
        End If

        If message.Length > 0 Then
            CType(Master, Site).ShowDialog("Login Failed", message)
        Else
            FormsAuthentication.RedirectFromLoginPage(txtUserName.Text.Trim, chkRememberPwd.Checked)
        End If

    End Sub

End Class
