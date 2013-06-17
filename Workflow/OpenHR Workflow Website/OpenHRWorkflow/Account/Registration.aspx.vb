Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class Registration
  Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 3)
    ' Set the e-mail input field to type=email (html5 only) ASP.NET requires this to be added thus:
    txtEmail.Attributes.Add("type", "email")
  End Sub

  Protected Sub BtnRegisterClick(sender As Object, e As ImageClickEventArgs) Handles btnRegister.Click

    Dim conn As SqlConnection
    Dim cmdRegistration As SqlCommand
    Dim cmdUserID As SqlCommand
    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""
    Dim objCrypt As New Crypt
    Dim strEncryptedString As String
    Dim userID As Long

    If txtEmail.Value.Length = 0 Then
      sMessage = "No email address entered."
    End If

    If sMessage.Length = 0 Then

      Try
        ' Fetch the record ID for the specified e-mail. 
        ' Needs to be done first (and separately) so it can be encrypted prior to sending back to SQL
        conn = New SqlConnection(Configuration.ConnectionString)
        conn.Open()

        cmdUserID = New SqlCommand
        cmdUserID.CommandText = "spASRSysMobileGetUserIDFromEmail"
        cmdUserID.Connection = conn
        cmdUserID.CommandType = CommandType.StoredProcedure

        cmdUserID.Parameters.Add("@psEmail", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        cmdUserID.Parameters("@psEmail").Value = txtEmail.Value

        cmdUserID.Parameters.Add("@piUserID", SqlDbType.Int).Direction = ParameterDirection.Output

        cmdUserID.ExecuteNonQuery()

        userID = CLng(NullSafeInteger(cmdUserID.Parameters("@piUserID").Value()))

        cmdUserID.Dispose()

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

          conn = New SqlClient.SqlConnection(Configuration.ConnectionString)
          conn.Open()

          cmdRegistration = New SqlClient.SqlCommand
          cmdRegistration.CommandText = "spASRSysMobileRegistration"
          cmdRegistration.Connection = conn
          cmdRegistration.CommandType = CommandType.StoredProcedure

          cmdRegistration.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
          cmdRegistration.Parameters("@psEmailAddress").Value = txtEmail.Value

          cmdRegistration.Parameters.Add("@psActivationURL", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
          cmdRegistration.Parameters("@psActivationURL").Value = Configuration.WorkflowUrl & "?" & strEncryptedString

          cmdRegistration.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

          cmdRegistration.ExecuteNonQuery()

          sMessage = CStr(cmdRegistration.Parameters("@psMessage").Value())

          cmdRegistration.Dispose()
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

  Protected Sub BtnHomeClick(sender As Object, e As ImageClickEventArgs) Handles btnHome.Click
    Response.Redirect("~/Account/Login.aspx")
  End Sub

End Class
