
Imports OpenHRWorkflow.Code.Classes

Partial Class Login
    Inherits Page

    Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

        Forms.RedirectToHomeIfAuthentcated()

        Title = Utilities.GetPageTitle("Login")
        Forms.LoadControlData(Me, 1)
        Form.DefaultButton = btnLogin2.UniqueID
        Form.DefaultFocus = txtUserName.ClientID

        If Request.QueryString.Count = 0 Then
          Session("CurrentStep") = Nothing
        End If

        If Session("CurrentStep") IsNot Nothing Then
            btnRegister.Visible = False
            btnForgotPwd.Visible = False
        End If

    End Sub

    Protected Sub BtnLoginClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnLogin.Click, btnLogin2.Click

        Dim authenticateOnly As Boolean = (Request.QueryString.Count > 0)
        Dim AuthenticationStep = CType(Session("CurrentStep"), StepAuthorization)
        Dim message As String

        If AuthenticationStep IsNot Nothing AndAlso AuthenticationStep.AuthorizedUsers.Count > 0 _
          AndAlso Not AuthenticationStep.AuthorizedUsers.Contains(txtUserName.Text.Trim.ToLower) Then
            message = "You are not authorised for this step"
        Else
            message = Security.ValidateUser(txtUserName.Text.Trim, txtPassword.Text, authenticateOnly)
        End If

        If message.Length > 0 Then
            CType(Master, Site).ShowDialog("Login Failed", message)
        Else
            dim allSteps = Security.GetStepDictionary()
            

            If Not Session("CurrentStep") Is Nothing Then
              CType(Session("CurrentStep"), StepAuthorization).HasBeenAuthenticated = True
            End If
            FormsAuthentication.RedirectFromLoginPage(txtUserName.Text.Trim, chkRememberPwd.Checked)
        End If

    End Sub

End Class
