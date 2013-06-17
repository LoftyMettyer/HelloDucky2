Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class Registration
  Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Title = WebSiteName("Registration")
    Forms.LoadControlData(Me, 3)
    Form.DefaultButton = btnRegister.UniqueID
    Form.DefaultFocus = txtEmail.ClientID
  End Sub

  Protected Sub BtnRegisterClick(sender As Object, e As EventArgs) Handles btnRegister.Click

    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""

    If Configuration.WorkflowUrl.Length = 0 Then
      sMessage = "Unable to determine Workflow URL."
    ElseIf Configuration.Login.Length = 0 Then
      sMessage = "Unable to connect to server."
    End If

    If sMessage.Length = 0 Then
      Try
        ' Fetch the record ID for the specified e-mail. 
        Dim userID = Database.GetUserID(txtEmail.Text)

        Dim objCrypt As New Crypt
        Dim strEncryptedString As String = objCrypt.EncryptQueryString((userID), -2, _
            Configuration.Login, _
            Configuration.Password, _
            Configuration.Server, _
            Configuration.Database, _
            User.Identity.Name, _
            "")

        sMessage = Database.Register(txtEmail.Text, Configuration.WorkflowUrl & "?" & strEncryptedString)

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
