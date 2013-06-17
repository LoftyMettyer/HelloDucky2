
Partial Class ForgottenLogin
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Title = Utilities.WebSiteName("Forgotten Login")
    Forms.LoadControlData(Me, 6)
    Form.DefaultButton = btnSubmit.UniqueID
    Form.DefaultFocus = txtEmail.ClientID
  End Sub

  Protected Sub BtnSubmitClick(sender As Object, e As EventArgs) Handles btnSubmit.Click

    Dim message As String = ""

    'Check the email address relates to a user
    Dim userID = Database.GetUserID(txtEmail.Text)

    If userID = 0 Then
      message = "No records exist with the given email address."
    Else
      'Send it all to sql to validate and email out
      message = Database.ForgotLogin(txtEmail.Text)
    End If

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Request Failed", message)
    Else
      CType(Master, Site).ShowDialog("Request Submitted", "An email has been sent to the entered address with your login details.", "Login.aspx")
    End If

  End Sub

End Class
