Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class Registration
  Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 3)
    ' Set the e-mail input field to type=email (html5 only) ASP.NET requires this to be added thus:
    txtEmail.Attributes.Add("type", "email")

    Title = WebSiteName("Registration")
  End Sub

  Protected Sub BtnRegisterClick(sender As Object, e As EventArgs) Handles btnRegisterButton.Click

    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""
    Dim objCrypt As New Crypt
    Dim strEncryptedString As String
    Dim userID As Integer

    If txtEmail.Value.Length = 0 Then
      sMessage = "No email address entered."
    End If

    If sMessage.Length = 0 Then

      Try
        ' Fetch the record ID for the specified e-mail. 
        ' Needs to be done first (and separately) so it can be encrypted prior to sending back to SQL
        Using conn As New SqlConnection(Configuration.ConnectionString)
          conn.Open()

          Dim cmd As New SqlCommand("spASRSysMobileGetUserIDFromEmail", conn)
          cmd.CommandType = CommandType.StoredProcedure

          cmd.Parameters.Add("@psEmail", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
          cmd.Parameters("@psEmail").Value = txtEmail.Value

          cmd.Parameters.Add("@piUserID", SqlDbType.Int).Direction = ParameterDirection.Output

          cmd.ExecuteNonQuery()

          userID = NullSafeInteger(cmd.Parameters("@piUserID").Value())

        End Using

        If Configuration.WorkflowUrl.Length = 0 Then
          sMessage = "Unable to determine Workflow URL."
        End If

        If Configuration.Login.Length = 0 Then
          sMessage = "Unable to connect to server."
        End If

        If sMessage.Length = 0 Then

          strEncryptedString = objCrypt.EncryptQueryString((userID), -2, _
              Configuration.Login, _
              Configuration.Password, _
              Configuration.Server, _
              Configuration.Database, _
              User.Identity.Name, _
              "")

          Using conn As New SqlConnection(Configuration.ConnectionString)
            conn.Open()

            Dim cmd As New SqlCommand("spASRSysMobileRegistration", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
            cmd.Parameters("@psEmailAddress").Value = txtEmail.Value

            cmd.Parameters.Add("@psActivationURL", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
            cmd.Parameters("@psActivationURL").Value = Configuration.WorkflowUrl & "?" & strEncryptedString

            cmd.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

            cmd.ExecuteNonQuery()

            sMessage = CStr(cmd.Parameters("@psMessage").Value())

          End Using
        End If
      Catch ex As Exception
        sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message & vbCrLf & "Contact your system administrator."
      End Try
    End If

    If sMessage.Length > 0 Then
      sHeader = "Registration Failed"
    Else
      sHeader = "Registration Submitted"
      sMessage = "An email has been sent to the entered address. To complete your registration, click the activation link in the email."
      sRedirectTo = "Login.aspx"
    End If

    CType(Master, Site).ShowDialog(sHeader, sMessage, sRedirectTo)

  End Sub

End Class
