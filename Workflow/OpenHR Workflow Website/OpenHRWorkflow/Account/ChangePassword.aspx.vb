
Partial Class ChangePassword
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    Dim result As CheckLoginResult = Database.CheckLoginDetails(User.Identity.Name)

    If Not result.Valid Then
      Session("message") = result.InvalidReason
      Response.Redirect("~/Message.aspx")
    End If

    Title = Utilities.WebSiteName("Change Password")
    Forms.LoadControlData(Me, 4)
    Form.DefaultButton = btnSubmit.UniqueID
    Form.DefaultFocus = txtCurrPassword.ClientID
  End Sub

  Protected Sub BtnSubmitClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click

    Dim message As String = Database.ChangePassword(User.Identity.Name, txtCurrPassword.Text, txtNewPassword.Text)

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Change Password Failed", message)
    Else
      CType(Master, Site).ShowDialog("Change Password Submitted", "Password changed successfully.", "../Home.aspx")
    End If

  End Sub

End Class
