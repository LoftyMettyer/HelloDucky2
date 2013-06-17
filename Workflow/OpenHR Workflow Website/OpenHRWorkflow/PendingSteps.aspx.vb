Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Utilities

Partial Class PendingSteps
    Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    Dim result As CheckLoginResult = Database.CheckLoginDetails(User.Identity.Name)
    Dim userGroupID As Integer

    If result.Valid Then
      userGroupID = result.UserGroupID
    Else
      Session("message") = result.InvalidReason
      Response.Redirect("~/Message.aspx")
    End If

    Forms.LoadControlData(Me, 5)

    Title = WebSiteName("To Do...")

    Dim todoTitleStyles = New Dictionary(Of String, String)
    Dim todoTitleForeColor As Integer
    Dim todoDescStyles = New Dictionary(Of String, String)
    Dim todoDescForeColor As Integer

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()

      Dim cmd As New SqlCommand("select * from tbsys_mobileformlayout where ID = 1", conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()
      dr.Read()

      todoTitleStyles.Add("font-family", NullSafeString(dr("TodoTitleFontName")))
      todoTitleStyles.Add("font-size", NullSafeString(dr("TodoTitleFontSize")) & "pt")
      todoTitleStyles.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(dr("TodoTitleFontBold"))), "bold", "normal"))
      todoTitleStyles.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(dr("TodoTitleFontItalic"))), "italic", "normal"))

      todoTitleForeColor = NullSafeInteger(dr("TodoTitleForeColor"))

      todoDescStyles.Add("font-family", NullSafeString(dr("TodoDescFontName")))
      todoDescStyles.Add("font-size", NullSafeString(dr("TodoDescFontSize")) & "pt")
      todoDescStyles.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(dr("TodoDescFontBold"))), "bold", "normal"))
      todoDescStyles.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(dr("TodoDescFontItalic"))), "italic", "normal"))

      todoDescForeColor = NullSafeInteger(dr("TodoDescForeColor"))
    End Using

    Dim userGroupHasPermission As Boolean

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()

    ' get the run permissions for workflow for this user group.
      Dim sql As String = "SELECT  [i].[itemKey], [p].[permitted]" & _
                          " FROM [ASRSysGroupPermissions] p" & _
                          " JOIN [ASRSysPermissionItems] i ON [p].[itemID] = [i].[itemID]" & _
                          " WHERE [p].[itemID] IN (" & _
                              " SELECT [itemID] FROM [ASRSysPermissionItems]	" & _
                              " WHERE [categoryID] = (SELECT [categoryID] FROM [ASRSysPermissionCategories] WHERE [categoryKey] = 'WORKFLOW')) " & _
                          " AND [groupName] = (SELECT [Name] FROM [ASRSysGroups] WHERE [ID] = " & userGroupID.ToString & ")"

      Dim cmd As New SqlCommand(sql, conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()

      While dr.Read()
        Select Case CStr(dr("itemKey"))
          Case "RUN"
            userGroupHasPermission = (dr("permitted") = True)
        End Select
      End While

    End Using

    Dim stepCount As Integer

    If userGroupHasPermission Then

      ' Get the pending steps.
      Using conn As New SqlConnection(Configuration.ConnectionString)
        conn.Open()

        Dim cmd As New SqlCommand("spASRSysMobileCheckPendingWorkflowSteps", conn)
        cmd.CommandType = CommandType.StoredProcedure

        cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        cmd.Parameters("@psKeyParameter").Value = User.Identity.Name

        Dim dr As SqlDataReader = cmd.ExecuteReader

        Dim table As Table, row As TableRow, cell As TableCell, label As Label, image As Image

        Dim general As New General

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
          label.Font.Underline = True
          label.Text = CStr(dr("name"))
          For Each item In todoTitleStyles
            label.Style.Add(item.Key, item.Value)
          Next
          label.Style.Add("color", general.GetHtmlColour(todoTitleForeColor))
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

          For Each item In todoDescStyles
            label.Style.Add(item.Key, item.Value)
          Next
          label.Style.Add("color", general.GetHtmlColour(todoDescForeColor))
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
