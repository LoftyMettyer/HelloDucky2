Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class ChangePassword
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Title = WebSiteName("Change Password")
    Forms.LoadControlData(Me, 4)
    Form.DefaultButton = btnSubmit.UniqueID
    Form.DefaultFocus = txtCurrPassword.ClientID
  End Sub

  Protected Sub BtnSubmitClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click

    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""

    Try
      ' Force password change only if there are no other security logged in with the same name.
      Dim userSessionCount As Integer = Database.GetUserCountOnServer(User.Identity.Name)

      If userSessionCount > 1 Then
        sMessage = String.Format("You could not change your password. The account is currently being used by {0} in the system.", _
                    If(userSessionCount > 2, userSessionCount.ToString & " security", "another user"))
      Else
        ' Change users password
        sMessage = Database.ChangePassword(User.Identity.Name, txtCurrPassword.Text, txtNewPassword.Text)
      End If

    Catch ex As Exception
      sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message.ToString & vbCrLf & vbCrLf & "Contact your system administrator."
    End Try

    If sMessage.Length > 0 Then
      sHeader = "Change Password Failed"
    Else
      sHeader = "Change Password Submitted"
      sMessage = "Password changed successfully."
      sRedirectTo = "../Home.aspx"
    End If

    CType(Master, Site).ShowDialog(sHeader, sMessage, sRedirectTo)

  End Sub

End Class
