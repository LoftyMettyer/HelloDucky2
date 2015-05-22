
Partial Class ForgottenLogin
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Title = Utilities.GetPageTitle("Forgotten Login")
    Forms.LoadControlData(Me, 6)
    Form.DefaultButton = btnSubmit2.UniqueID
    Form.DefaultFocus = txtEmail.ClientID
  End Sub

  Protected Sub BtnSubmitClick(sender As Object, e As EventArgs) Handles btnSubmit.Click, btnSubmit2.Click

      Dim db As New Database(App.Config.ConnectionString)
		db.ForgotLogin(txtEmail.Text)

		CType(Master, Site).ShowDialog("Request Submitted", "An email has been sent to the entered address with your login details.", "Login.aspx")

  End Sub

	Public Function CanChangeInputTypeToEmail() As Boolean
		Return IsMobileBrowser() Or IsTablet()
	End Function

End Class
