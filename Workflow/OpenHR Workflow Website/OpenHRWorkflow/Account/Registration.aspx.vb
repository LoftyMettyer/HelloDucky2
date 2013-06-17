
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
      ' Fetch the record ID for the specified e-mail.
      Dim userID = Database.GetUserID(txtEmail.Text)

      Dim objCrypt As New Crypt
      Dim strEncryptedString As String = objCrypt.EncryptQueryString((userID), -2, _
          Configuration.Login, _
          Configuration.Password, _
          Configuration.Server, _
          Configuration.Database, _
          User.Identity.Name, _
          "")

      message = Database.Register(txtEmail.Text, Configuration.WorkflowUrl & "?" & strEncryptedString)
    End If

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Registration Failed", message)
    Else
      CType(Master, Site).ShowDialog("Registration Submitted", "An email has been sent to the entered address. " & _
                                     "To complete your registration, click the activation link in the email.", "Login.aspx")
    End If

  End Sub

End Class
