
Partial Class ForgottenLogin
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Title = Utilities.WebSiteName("Forgotten Login")
    Forms.LoadControlData(Me, 6)
    Form.DefaultButton = btnSubmit2.UniqueID
    Form.DefaultFocus = txtEmail.ClientID
  End Sub

  Protected Sub BtnSubmitClick(sender As Object, e As EventArgs) Handles btnSubmit.Click, btnSubmit2.Click

    Dim db As New Database
    Dim message As String = db.ForgotLogin(txtEmail.Text)

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Request Failed", message)
    Else
      CType(Master, Site).ShowDialog("Request Submitted", "An email has been sent to the entered address with your login details.", "Login.aspx")
    End If

  End Sub

End Class
