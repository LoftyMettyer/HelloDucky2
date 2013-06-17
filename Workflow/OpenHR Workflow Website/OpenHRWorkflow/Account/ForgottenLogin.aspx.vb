Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class ForgottenLogin
  Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 6)
    ' Set the e-mail input field to type=email (html5 only) ASP.NET requires this to be added thus:
    txtEmail.Attributes.Add("type", "email")
  End Sub

  Protected Sub BtnSubmitClick(sender As Object, e As ImageClickEventArgs) Handles btnSubmit.Click

    Dim conn As SqlConnection
    Dim cmdForgotLogin As SqlCommand
    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""
    Dim lngUserID As Long

    Try
      conn = New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      ' Done in three parts. First get the ID for this e-mail (SQL). Second retrieve and decrypt password (VB), third send a reminder e-mail (SQL).
      ' Scratch that! First get the username from the db for this email address, then send the e-mail.

      cmdForgotLogin = New SqlCommand
      cmdForgotLogin.CommandText = "spASRSysMobileGetUserIDFromEmail"
      cmdForgotLogin.Connection = conn
      cmdForgotLogin.CommandType = CommandType.StoredProcedure

      cmdForgotLogin.Parameters.Add("@psEmail", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmdForgotLogin.Parameters("@psEmail").Value = txtEmail.Value

      cmdForgotLogin.Parameters.Add("@piUserID", SqlDbType.Int).Direction = ParameterDirection.Output

      cmdForgotLogin.ExecuteNonQuery()

      lngUserID = CLng(NullSafeInteger(cmdForgotLogin.Parameters("@piUserID").Value()))

      cmdForgotLogin.Dispose()

      If lngUserID = 0 Then sMessage = "No records exist with the given email address."

      If sMessage.Length = 0 Then
        ' ------------- Part two, send it all to sql to validate and email out -----------------
        cmdForgotLogin = New SqlCommand
        cmdForgotLogin.CommandText = "spASRSysMobileForgotLogin"
        cmdForgotLogin.Connection = conn
        cmdForgotLogin.CommandType = CommandType.StoredProcedure

        cmdForgotLogin.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        cmdForgotLogin.Parameters("@psEmailAddress").Value = txtEmail.Value

        cmdForgotLogin.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

        cmdForgotLogin.ExecuteNonQuery()

        sMessage = CStr(cmdForgotLogin.Parameters("@psMessage").Value())

        cmdForgotLogin.Dispose()
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

    CType(Master, Site).ShowMessage(sHeader, sMessage, sRedirectTo)

  End Sub

  Protected Sub BtnCancelClick(sender As Object, e As ImageClickEventArgs) Handles btnCancel.Click
    Response.Redirect("~/Account/Login.aspx")
  End Sub

End Class
