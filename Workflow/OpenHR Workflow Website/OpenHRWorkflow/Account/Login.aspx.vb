Imports System.Data.SqlClient
Imports Utilities

Partial Class Login
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 1)

    Title = WebSiteName("Login")

  End Sub

  Protected Sub BtnLoginClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnLoginButton.Click

    Dim sMessage As String = ""

    Try
      ' Check if the system is locked
      If Database.IsSystemLocked() Then
        sMessage = "Database locked." & vbCrLf & "Contact your system administrator."
      End If
    Catch ex As Exception
      sMessage = "Unable to perform system lock check."
    End Try

    ' Continue with authentication
    If sMessage.Length = 0 Then
      Try
        If Not Security.ValidateUser(txtUserName.Text.Trim, txtPassword.Text) Then
          sMessage = "The user name or password provided is incorrect."
        End If
      Catch ex As Exception
        sMessage = ex.Message
      End Try

    End If

    If sMessage.Length = 0 Then
      Try
        Dim result As CheckLoginResult = Database.CheckLoginDetails(txtUserName.Text.Trim)
        If Not result.Valid Then
          sMessage = result.InvalidReason
        End If
      Catch ex As Exception
        sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Contact your system administrator."
      End Try
    End If

    If sMessage.Length > 0 Then
      CType(Master, Site).ShowDialog("Login Failed", sMessage, "")
    Else
      FormsAuthentication.RedirectFromLoginPage(txtUserName.Text.Trim, chkRememberPwd.Checked)
    End If

  End Sub

End Class
