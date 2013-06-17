Imports Utilities

Partial Class Login
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    'Go to the home page if already logged in
    If Request.IsAuthenticated Then
      Response.Redirect("~/Home.aspx")
      Return
    End If

    Title = WebSiteName("Login")
    Forms.LoadControlData(Me, 1)
    Form.DefaultButton = btnLogin.UniqueID
    Form.DefaultFocus = txtUserName.ClientID
  End Sub

  Protected Sub BtnLoginClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnLogin.Click

    Dim sMessage As String = ""

    Try
      ' Check if the system is locked
      If Database.IsSystemLocked() Then
        sMessage = "Database locked." & vbCrLf & "Contact your system administrator."

      ElseIf Not Security.ValidateUser(txtUserName.Text.Trim, txtPassword.Text) Then
        sMessage = "The system could not log you on. Make sure your details are correct, then retype your password."

      Else
        Dim result As CheckLoginResult = Database.CheckLoginDetails(txtUserName.Text.Trim)
        If Not result.Valid Then
          sMessage = result.InvalidReason
        End If
      End If

    Catch ex As Exception
      sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Contact your system administrator."
    End Try

    If sMessage.Length > 0 Then
      CType(Master, Site).ShowDialog("Login Failed", sMessage, "")
    Else
      FormsAuthentication.RedirectFromLoginPage(txtUserName.Text.Trim, chkRememberPwd.Checked)
    End If

  End Sub

End Class
