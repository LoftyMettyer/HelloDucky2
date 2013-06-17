﻿Imports System.Data.SqlClient
Imports Utilities

Partial Class Login
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    Forms.LoadControlData(Me, 1)
  End Sub

  Protected Sub BtnLoginClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnLogin.Click

    Dim sMessage As String = ""
    Dim userName As String = txtUserName.Value.Trim

    ' Check if the system is locked
    Try
      If Database.IsSystemLocked() Then
        sMessage = "Database locked." & vbCrLf & "Contact your system administrator."
      End If
    Catch ex As Exception
      sMessage = "Unable to perform system lock check."
    End Try

    ' Continue with authentication
    If sMessage.Length = 0 Then

      Try
        Dim valid As Boolean

        If userName.IndexOf("\") > 0 Then
          'Active dirctory authentication
          valid = Security.ValidateActiveDirectoryUser(userName.Split("\"c)(0), userName.Split("\"c)(1), txtPassword.Value)
        Else
          'Sql server authentication
          valid = Security.ValidateSqlServerUser(userName, txtPassword.Value)
        End If

        If Not valid Then sMessage = "The user name or password provided is incorrect."
      Catch ex As Exception
        sMessage = ex.Message
      End Try

    End If

    If sMessage.Length = 0 Then
      Try
        Dim result As CheckLoginResult = Database.CheckLoginDetails(userName)
        If Not result.Valid Then
          sMessage = result.InvalidReason
        End If
      Catch ex As Exception
        sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Contact your system administrator."
      End Try
    End If

    If sMessage.Length > 0 Then
      CType(Master, Site).ShowMessage("Login Failed", sMessage, "")
    Else
      FormsAuthentication.RedirectFromLoginPage(userName, chkRememberPwd.Checked)
    End If

  End Sub

  Protected Sub BtnRegisterClick(sender As Object, e As ImageClickEventArgs) Handles btnRegister.Click
    Response.Redirect("~/Account/Registration.aspx")
  End Sub

  Protected Sub BtnForgotPwdClick(sender As Object, e As ImageClickEventArgs) Handles btnForgotPwd.Click
    Response.Redirect("~/Account/ForgottenLogin.aspx")
  End Sub

End Class
