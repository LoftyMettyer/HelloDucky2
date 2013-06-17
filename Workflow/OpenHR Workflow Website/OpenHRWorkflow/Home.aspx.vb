Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Utilities

Partial Class Home
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

    Forms.LoadControlData(Me, 2)

    Title = WebSiteName("Home")

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

    Dim userGroupHasPermission As Boolean

    Using conn As New SqlConnection(Configuration.ConnectionString)

      ' get the run permissions for workflow for this user group.
      Dim sql As String = "SELECT  [i].[itemKey], [p].[permitted]" & _
                           " FROM [ASRSysGroupPermissions] p" & _
                           " JOIN [ASRSysPermissionItems] i ON [p].[itemID] = [i].[itemID]" & _
                           " WHERE [p].[itemID] IN (" & _
                               " SELECT [itemID] FROM [ASRSysPermissionItems]	" & _
                                " WHERE [categoryID] = (SELECT [categoryID] FROM [ASRSysPermissionCategories] WHERE [categoryKey] = 'WORKFLOW')) " & _
                                " AND [groupName] = (SELECT [Name] FROM [ASRSysGroups] WHERE [ID] = " & userGroupID.ToString & ")"
      conn.Open()
      Dim cmd As New SqlCommand(sql, conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()

      While dr.Read()
        Select Case CStr(dr("itemKey"))
          Case "RUN"
            userGroupHasPermission = (dr("permitted") = True)
        End Select
      End While

    End Using

    Dim itemCount As Integer

    If userGroupHasPermission Then

      Using conn As New SqlConnection(Configuration.ConnectionString)

        Dim sql As String = "SELECT w.Id, w.Name, w.PictureID" & _
              " FROM tbsys_mobilegroupworkflows gw" & _
              " INNER JOIN tbsys_workflows w on gw.WorkflowID = w.ID" & _
              " WHERE gw.UserGroupID = " & userGroupID & " AND w.enabled = 1 ORDER BY gw.Pos ASC"

        conn.Open()
        Dim cmd As New SqlCommand(sql, conn)
        Dim dr As SqlDataReader = cmd.ExecuteReader()

        ' Create the holding table for the list of workflows.
        Dim table = New Table
        table.Style.Add("width", "100%")

        'Iterate through the results
        While dr.Read()

          ' Create a row to contain this pending step...
          Dim row = New TableRow
          row.Style.Add("width", "100%")
          row.Attributes.Add("onclick", "window.open('" & WorkflowLink(CInt(dr("ID"))) & "');")
          row.Style.Add("cursor", "pointer")

          ' Create a cell to contain the workflow icon
          Dim cell = New TableCell  ' Image cell
          cell.Style.Add("width", "57px")

          Dim image As New Image, sImageFileName As String
          If NullSafeInteger(dr("PictureID")) = 0 Then
            sImageFileName = "~/Images/Connected48.png"
          Else
            sImageFileName = Picture.GetUrl(CInt(dr("PictureID")))
          End If
          image.ImageUrl = sImageFileName
          image.Height() = Unit.Pixel(57)
          image.Width() = Unit.Pixel(57)
          image.Style.Add("cursor", "pointer")

          ' add ImageButton to cell
          cell.Controls.Add(image)

          ' Add cell to row
          row.Cells.Add(cell)

          ' Create a cell to contain the workflow name and description
          cell = New TableCell
          Dim label = New Label ' Workflow name text
          label.Text = CStr(dr("Name"))
          label.Font.Name = homeItemFontInfo.Name
          label.Font.Size = New FontUnit(homeItemFontInfo.Size)
          label.Font.Bold = homeItemFontInfo.Bold
          label.Font.Italic = homeItemFontInfo.Italic
          label.Font.Underline = homeItemFontInfo.Underline
          label.Font.Strikeout = homeItemFontInfo.Strikeout
          label.Style.Add("color", General.GetHtmlColour(homeItemForeColor))
          label.Style.Add("cursor", "pointer")

          cell.Controls.Add(label)

          ' Add cell to row, and row to table.
          row.Cells.Add(cell)

          table.Rows.Add(row)

          itemCount += 1
        End While
        pnlWFList.Controls.Add(table)

      End Using

    End If

    If itemCount > 0 Then
      lblNothingTodo.Visible = False
    Else
      lblWelcome.Visible = False
    End If

    ' Update the wf steps count
    If userGroupHasPermission Then

      Dim count As Integer = GetPendingStepsCount()

      lblWFCount.Text = CStr(count)
      lblWFCount.Visible = (count > 0)
    End If

    ' Disable the Change Password button for windows authenticated Security
    If User.Identity.Name.Contains("\") Then
      btnChangePwd.Visible = False
    End If

  End Sub

  Public Function WorkflowLink(ByVal workflowID As Integer) As String

    If Configuration.WorkflowUrl.Length = 0 Then
      Return ""
    End If

    If Configuration.Login.Length = 0 Then
      Return ""
    End If

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

    ' clear authentication cookie
    Dim cookie As HttpCookie = New HttpCookie(FormsAuthentication.FormsCookieName, "")
    cookie.Expires = DateTime.Now.AddYears(-1)
    Response.Cookies.Add(cookie)

    FormsAuthentication.RedirectToLoginPage()

  End Sub

  Private Function GetPendingStepsCount() As Integer

    Using conn As New SqlConnection(Configuration.ConnectionString)
      conn.Open()

      Dim cmd As New SqlClient.SqlCommand
      cmd.CommandText = "spASRSysMobileCheckPendingWorkflowSteps"
      cmd.Connection = conn
      cmd.CommandType = CommandType.StoredProcedure

      cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmd.Parameters("@psKeyParameter").Value = User.Identity.Name

      Dim dr As SqlClient.SqlDataReader = cmd.ExecuteReader

      Dim count As Integer
      While dr.Read
        count += 1
      End While
      Return count
    End Using

  End Function

End Class
