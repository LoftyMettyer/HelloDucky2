
Partial Class Registration
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
    Title = Utilities.WebSiteName("Registration")
    Forms.LoadControlData(Me, 3)
    Form.DefaultButton = btnRegister.UniqueID
    Form.DefaultFocus = txtEmail.ClientID
  End Sub

  Protected Sub BtnRegisterClick(sender As Object, e As EventArgs) Handles btnRegister.Click

    Dim message As String = ""

    If Configuration.WorkflowUrl.Length = 0 Then
      message = "Unable to determine Workflow URL."
    ElseIf Configuration.Login.Length = 0 Then
      message = "Unable to connect to server."
    Else
      message = Database.Register(txtEmail.Text)
    End If

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Registration Failed", message)
    Else
      CType(Master, Site).ShowDialog("Registration Submitted", "An email has been sent to the entered address. " & _
                                     "To complete your registration, click the activation link in the email.", "Login.aspx")
    End If

  End Sub

End Class
