
Partial Class Registration
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
    Title = Utilities.WebSiteName("Registration")
    Forms.LoadControlData(Me, 3)
    Form.DefaultButton = btnSubmit2.UniqueID
    Form.DefaultFocus = txtEmail.ClientID
  End Sub

  Protected Sub BtnRegisterClick(sender As Object, e As EventArgs) Handles btnSubmit.Click, btnSubmit2.Click

    Dim db As New Database
    Dim message As String = db.Register(txtEmail.Text)

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Registration Failed", message)
    Else
      CType(Master, Site).ShowDialog("Registration Submitted", "An email has been sent to the entered address. " & _
                                     "To complete your registration, click the activation link in the email.", "Login.aspx")
    End If

  End Sub

End Class
