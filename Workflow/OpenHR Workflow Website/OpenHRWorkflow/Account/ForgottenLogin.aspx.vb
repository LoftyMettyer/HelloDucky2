Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class ForgottenLogin
  Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 6)

    Title = WebSiteName("Forgotten Login")
    Page.Form.DefaultButton = btnSubmitButton.UniqueID
    Page.Form.DefaultFocus = txtEmail.UniqueID
  End Sub

  Protected Sub BtnSubmitClick(sender As Object, e As EventArgs) Handles btnSubmitButton.Click

    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""

    Try
      ' Done in three parts. First get the ID for this e-mail (SQL). Second retrieve and decrypt password (VB), third send a reminder e-mail (SQL).
      ' Scratch that! First get the username from the db for this email address, then send the e-mail.
      Using conn As New SqlConnection(Configuration.ConnectionString)
        conn.Open()

        Dim cmd = New SqlCommand("spASRSysMobileGetUserIDFromEmail", conn)
        cmd.CommandType = CommandType.StoredProcedure

        cmd.Parameters.Add("@psEmail", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        cmd.Parameters("@psEmail").Value = txtEmail.Text

        cmd.Parameters.Add("@piUserID", SqlDbType.Int).Direction = ParameterDirection.Output

        cmd.ExecuteNonQuery()

        Dim userID = NullSafeInteger(cmd.Parameters("@piUserID").Value())

        If userID = 0 Then sMessage = "No records exist with the given email address."
      End Using

      If sMessage.Length = 0 Then
        ' ------------- Part two, send it all to sql to validate and email out -----------------
        Using conn As New SqlConnection(Configuration.ConnectionString)
          conn.Open()

          Dim cmd As New SqlCommand("spASRSysMobileForgotLogin", conn)
          cmd.CommandType = CommandType.StoredProcedure

          cmd.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
          cmd.Parameters("@psEmailAddress").Value = txtEmail.Text

          cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

          cmd.ExecuteNonQuery()

          sMessage = CStr(cmd.Parameters("@psMessage").Value())

        End Using
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
