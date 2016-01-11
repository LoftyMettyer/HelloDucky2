Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Utilities

Partial Class PendingSteps
  Inherits Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
    Title = WebSiteName("To Do...")
    Forms.LoadControlData(Me, 5)

    Dim db As New Database
    Dim result As CheckLoginResult = db.CheckLoginDetails(User.Identity.Name)
    Dim userGroupID As Integer

    If result.Valid Then
      userGroupID = result.UserGroupID
    Else
      FormsAuthentication.SignOut()
      FormsAuthentication.RedirectToLoginPage()
    End If

    Dim todoTitleForeColor As Integer
    Dim todoTitleFontInfo As New FontSetting
    Dim todoDescForeColor As Integer
    Dim todoDescFontInfo As New FontSetting

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()

      Dim cmd As New SqlCommand("select * from tbsys_mobileformlayout where ID = 1", conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()
      dr.Read()

      todoTitleForeColor = NullSafeInteger(dr("TodoTitleForeColor"))
      todoTitleFontInfo.Name = NullSafeString(dr("TodoTitleFontName"))
      todoTitleFontInfo.Size = NullSafeSingle(dr("TodoTitleFontSize"))
      todoTitleFontInfo.Bold = NullSafeBoolean(dr("TodoTitleFontBold"))
      todoTitleFontInfo.Italic = NullSafeBoolean(dr("TodoTitleFontItalic"))
      todoTitleFontInfo.Underline = NullSafeBoolean(dr("TodoTitleFontUnderline"))
      todoTitleFontInfo.Strikeout = NullSafeBoolean(dr("TodoTitleFontStrikeout"))

      todoDescForeColor = NullSafeInteger(dr("TodoDescForeColor"))
      todoDescFontInfo.Name = NullSafeString(dr("TodoDescFontName"))
      todoDescFontInfo.Size = NullSafeSingle(dr("TodoDescFontSize"))
      todoDescFontInfo.Bold = NullSafeBoolean(dr("TodoDescFontBold"))
      todoDescFontInfo.Italic = NullSafeBoolean(dr("TodoDescFontItalic"))
      todoDescFontInfo.Underline = NullSafeBoolean(dr("TodoDescFontUnderline"))
      todoDescFontInfo.Strikeout = NullSafeBoolean(dr("TodoDescFontStrikeout"))
    End Using

    Dim canRun As Boolean = db.CanRunWorkflows(userGroupID)
    Dim workflows As New List(Of WorkflowStepLink)

    If canRun Then
      workflows = db.GetPendingStepList(User.Identity.Name)
    End If

    For Each item In workflows

      Dim li As New HtmlGenericControl("li")

      Dim link As New HyperLink
      link.NavigateUrl = item.Url
      link.Target = "_blank"

      Dim imageContainer As New HtmlGenericControl("span")
      imageContainer.Attributes.Add("class", "image")

      Dim image As New Image
      image.ImageUrl = If(item.PictureID = 0, "~/Images/Connected48.png", Picture.GetUrl(item.PictureID))

      Dim detailContainer As New HtmlGenericControl("span")
      detailContainer.Attributes.Add("class", "detail")

      Dim label = New Label
      label.Text = item.Name
      label.Font.Name = todoTitleFontInfo.Name
      label.Font.Size = New FontUnit(todoTitleFontInfo.Size)
      label.Font.Bold = todoTitleFontInfo.Bold
      label.Font.Italic = todoTitleFontInfo.Italic
      label.Font.Underline = todoTitleFontInfo.Underline
      label.Font.Strikeout = todoTitleFontInfo.Strikeout
      label.Style.Add("color", General.GetHtmlColour(todoTitleForeColor))

      Dim labelDesc As New Label
      labelDesc.Text = item.Desc
      labelDesc.Font.Name = todoDescFontInfo.Name
      labelDesc.Font.Size = New FontUnit(todoDescFontInfo.Size)
      labelDesc.Font.Bold = todoDescFontInfo.Bold
      labelDesc.Font.Italic = todoDescFontInfo.Italic
      labelDesc.Font.Underline = todoDescFontInfo.Underline
      labelDesc.Font.Strikeout = todoDescFontInfo.Strikeout
      labelDesc.Style.Add("color", General.GetHtmlColour(todoDescForeColor))

      workflowList.Controls.Add(li)
      li.Controls.Add(link)
      link.Controls.Add(imageContainer)
      imageContainer.Controls.Add(image)
      link.Controls.Add(detailContainer)
      detailContainer.Controls.Add(label)
      detailContainer.Controls.Add(labelDesc)
    Next

    If workflows.Count > 0 Then
      lblNothingTodo.Visible = False
    Else
      lblInstruction.Visible = False
    End If

  End Sub

End Class
