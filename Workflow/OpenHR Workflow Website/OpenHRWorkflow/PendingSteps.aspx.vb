Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Utilities

Partial Class PendingSteps
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
    Title = WebSiteName("To Do...")
    Forms.LoadControlData(Me, 5)

    Dim result As CheckLoginResult = Database.CheckLoginDetails(User.Identity.Name)
    Dim userGroupID As Integer

    If result.Valid Then
      userGroupID = result.UserGroupID
    Else
      Session("message") = result.InvalidReason
      Response.Redirect("~/Message.aspx")
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

    Dim canRun As Boolean = Database.CanUserGroupRunWorkflows(userGroupID)

    Dim stepCount As Integer

    If canRun Then

      ' Get the pending steps.
      Using conn As New SqlConnection(Configuration.ConnectionString)
        conn.Open()

        Dim cmd As New SqlCommand("spASRSysMobileCheckPendingWorkflowSteps", conn)
        cmd.CommandType = CommandType.StoredProcedure

        cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        cmd.Parameters("@psKeyParameter").Value = User.Identity.Name

        Dim dr As SqlDataReader = cmd.ExecuteReader

        Dim table As Table, row As TableRow, cell As TableCell, label As Label, image As Image

        ' Create the holding table
        table = New Table

        While (dr.Read)
          ' Create a row to contain this pending step...
          row = New TableRow
          row.Attributes.Add("onclick", "window.open('" & dr("url").ToString & "');")
          row.Style.Add("cursor", "pointer")

          ' Create a cell to contain the workflow icon
          cell = New TableCell  ' Image cell
          image = New Image

          Dim fileName As String
          If NullSafeInteger(dr("PictureID")) = 0 Then
            fileName = "~/Images/Connected48.png"
          Else
            fileName = Picture.GetUrl(CInt(dr("PictureID")))
          End If
          image.ImageUrl = fileName
          image.Height() = Unit.Pixel(57)
          image.Width() = Unit.Pixel(57)
          image.Style.Add("cursor", "pointer")

          ' add ImageButton to cell
          cell.Controls.Add(image)

          ' Add cell to row
          row.Cells.Add(cell)

          ' Create a cell to contain the workflow name and description
          cell = New TableCell
          label = New Label ' Workflow name text
          label.Text = CStr(dr("name"))
          label.Font.Name = todoTitleFontInfo.Name
          label.Font.Size = New FontUnit(todoTitleFontInfo.Size)
          label.Font.Bold = todoTitleFontInfo.Bold
          label.Font.Italic = todoTitleFontInfo.Italic
          label.Font.Underline = todoTitleFontInfo.Underline
          label.Font.Strikeout = todoTitleFontInfo.Strikeout
          label.Style.Add("color", General.GetHtmlColour(todoTitleForeColor))
          label.Style.Add("cursor", "pointer")
          cell.Controls.Add(label)

          ' Line Break
          cell.Controls.Add(New LiteralControl("<br>"))

          label = New Label ' Workflow step description text

          Dim desc As String
          If Left(CStr(dr("description")), Len(dr("name")) + 2) = (Trim(CStr(dr("name"))) & " -") Then
            desc = dr("description").ToString.Remove(0, (dr("name").ToString.Length) + 2)
          Else
            desc = dr("description").ToString
          End If
          label.Text = desc
          label.Font.Name = todoDescFontInfo.Name
          label.Font.Size = New FontUnit(todoDescFontInfo.Size)
          label.Font.Bold = todoDescFontInfo.Bold
          label.Font.Italic = todoDescFontInfo.Italic
          label.Font.Underline = todoDescFontInfo.Underline
          label.Font.Strikeout = todoDescFontInfo.Strikeout
          label.Style.Add("color", General.GetHtmlColour(todoDescForeColor))
          label.Style.Add("cursor", "pointer")
          cell.Controls.Add(label)

          ' Add cell to row, and row to table.
          row.Cells.Add(cell)
          table.Rows.Add(row)

          stepCount += 1
        End While

        pnlWFList.Controls.Add(table)

      End Using

    End If

    If stepCount > 0 Then
      lblNothingTodo.Visible = False
    Else
      lblInstruction.Visible = False
    End If

  End Sub

End Class
