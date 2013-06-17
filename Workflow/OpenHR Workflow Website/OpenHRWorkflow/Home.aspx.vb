﻿Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Utilities

Partial Class Home
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
    Title = WebSiteName("Home")
    Forms.LoadControlData(Me, 2)

    Dim result As CheckLoginResult = Database.CheckLoginDetails(User.Identity.Name)
    Dim userGroupID As Integer

    If result.Valid Then
      userGroupID = result.UserGroupID
    Else
      FormsAuthentication.SignOut()
      FormsAuthentication.RedirectToLoginPage()
    End If

    Dim homeItemFontInfo As New FontSetting
    Dim homeItemForeColor As Integer

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()
      Dim cmd As New SqlCommand("SELECT * FROM tbsys_mobileformlayout WHERE ID = 1", conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()
      dr.Read()

      homeItemForeColor = NullSafeInteger(dr("HomeItemForeColor"))
      homeItemFontInfo.Name = NullSafeString(dr("HomeItemFontName"))
      homeItemFontInfo.Size = NullSafeSingle(dr("HomeItemFontSize"))
      homeItemFontInfo.Bold = NullSafeBoolean(dr("HomeItemFontBold"))
      homeItemFontInfo.Italic = NullSafeBoolean(dr("HomeItemFontItalic"))
      homeItemFontInfo.Underline = NullSafeBoolean(dr("HomeItemFontUnderline"))
      homeItemFontInfo.Strikeout = NullSafeBoolean(dr("HomeItemFontStrikeout"))
    End Using

    Dim canRun As Boolean = Database.CanRunWorkflows(userGroupID)
    Dim workflows As New List(Of WorkflowLink)

    If canRun Then
      workflows = Database.GetWorkflowList(userGroupID)
    End If

    For Each item In workflows

      Dim li As New HtmlGenericControl("li")

      Dim link As New HyperLink
      link.NavigateUrl = WorkflowLink(item.ID)
      link.Target = "_blank"

      Dim image As New Image
      image.ImageUrl = If(item.PictureID = 0, "~/Images/Connected48.png", Picture.GetUrl(item.PictureID))

      Dim label = New Label
      label.Text = item.Name
      label.Font.Name = homeItemFontInfo.Name
      label.Font.Size = New FontUnit(homeItemFontInfo.Size)
      label.Font.Bold = homeItemFontInfo.Bold
      label.Font.Italic = homeItemFontInfo.Italic
      label.Font.Underline = homeItemFontInfo.Underline
      label.Font.Strikeout = homeItemFontInfo.Strikeout
      label.Style.Add("color", General.GetHtmlColour(homeItemForeColor))

      workflowList.Controls.Add(li)
      li.Controls.Add(link)
      link.Controls.Add(image)
      link.Controls.Add(label)
    Next

    If workflows.Count > 0 Then
      lblNothingTodo.Visible = False
    Else
      lblWelcome.Visible = False
    End If

    ' Update the workflow step count indicator
    If canRun Then
      Dim count As Integer = Database.GetPendingStepCount(User.Identity.Name)
      lblWFCount.Text = CStr(count)
      lblWFCount.Visible = (count > 0)
    Else
      lblWFCount.Visible = False
    End If

    ' Disable the Change Password button for windows authenticated Security
    If User.Identity.Name.Contains("\") Then
      btnChangePwd.Visible = False
    End If

    'ListView1.DataSource = workflows
    'ListView1.DataBind()

  End Sub

  Public Function WorkflowLink(ByVal workflowID As Integer) As String

    ' For externally initiated workflows:
    '      plngInstance = -1 * workflowID
    '      plngStepID = -

    Dim objCrypt As New Crypt
    Dim sEncryptedString As String = objCrypt.EncryptQueryString((-1 * workflowID), -1, _
        Configuration.Login, _
        Configuration.Password, _
        Configuration.Server, _
        Configuration.Database, _
        User.Identity.Name, _
        "")

    Return Configuration.WorkflowUrl & "?" & sEncryptedString

  End Function

  Protected Sub BtnLogoutClick(sender As Object, e As EventArgs) Handles btnLogout.Click
    FormsAuthentication.SignOut()
    FormsAuthentication.RedirectToLoginPage()
  End Sub

End Class

