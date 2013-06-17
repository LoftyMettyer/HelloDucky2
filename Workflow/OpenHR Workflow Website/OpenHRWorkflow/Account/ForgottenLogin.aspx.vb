Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class ForgottenLogin
  Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Title = WebSiteName("Forgotten Login")
    Forms.LoadControlData(Me, 6)
    Form.DefaultButton = btnSubmit.UniqueID
    Form.DefaultFocus = txtEmail.ClientID
  End Sub

  Protected Sub BtnSubmitClick(sender As Object, e As EventArgs) Handles btnSubmit.Click

    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""

    Try
      'Check the email address relates to a user
      Dim userID = Database.GetUserID(txtEmail.Text)

      If userID = 0 Then
        sMessage = "No records exist with the given email address."
      Else
        'Send it all to sql to validate and email out
        sMessage = Database.ForgotLogin(txtEmail.Text)
      End If

    Catch ex As Exception
      sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Contact your system administrator."
    End Try

    If sMessage.Length > 0 Then
      sHeader = "Request Failed"
    Else
      sHeader = "Request Submitted"
      sMessage = "An email has been sent to the entered address with your login details."
      sRedirectTo = "Login.aspx"
    End If

    CType(Master, Site).ShowDialog(sHeader, sMessage, sRedirectTo)

  End Sub

End Class
