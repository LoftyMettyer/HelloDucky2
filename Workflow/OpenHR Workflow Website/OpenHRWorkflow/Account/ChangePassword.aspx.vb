﻿
Partial Class ChangePassword
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
    Title = Utilities.WebSiteName("Change Password")
    Forms.LoadControlData(Me, 4)
    Form.DefaultButton = btnSubmit.UniqueID
    Form.DefaultFocus = txtCurrPassword.ClientID
  End Sub

  Protected Sub BtnSubmitClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click

    ' Change users password
    Dim message As String = Database.ChangePassword(User.Identity.Name, txtCurrPassword.Text, txtNewPassword.Text)

    If message.Length > 0 Then
      CType(Master, Site).ShowDialog("Change Password Failed", message)
    Else
      CType(Master, Site).ShowDialog("Change Password Submitted", "Password changed successfully.", "../Home.aspx")
    End If

  End Sub

End Class
