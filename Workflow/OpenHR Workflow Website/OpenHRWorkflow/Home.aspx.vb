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

    Dim sMessage As String = ""
    Dim sImageFileName As String = ""
    Dim homeItemStyles = New Dictionary(Of String, String)

    Using conn As New SqlConnection(Configuration.ConnectionString)

      conn.Open()
      Dim cmd As New SqlCommand("SELECT * FROM tbsys_mobileformlayout WHERE ID = 1", conn)
      Dim dr As SqlDataReader = cmd.ExecuteReader()
      dr.Read()

      homeItemStyles.Add("font-family", NullSafeString(dr("HomeItemFontName")))
      homeItemStyles.Add("font-size", NullSafeString(dr("HomeItemFontSize")) & "pt")
      homeItemStyles.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(dr("HomeItemFontBold"))), "bold", "normal"))
      homeItemStyles.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(dr("HomeItemFontItalic"))), "italic", "normal"))

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
      Dim cmd As New SqlClient.SqlCommand(sql, conn)
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

          Dim image = New Image
          If NullSafeInteger(dr("pictureID")) = 0 Then
            sImageFileName = "~/Images/Connected48.png"
          Else
            sImageFileName = "~/" & Picture.LoadPicture(CInt(dr("pictureID")))
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
          For Each item In homeItemStyles
            label.Style.Add(item.Key, item.Value)
          Next
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
      If count > 0 Then
        lblWFCount.InnerText = CStr(count)
        pnlWFCount.Style.Add("visibility", "visible")
      Else
        pnlWFCount.Style.Add("visibility", "hidden")
      End If
    End If

    ' Disable the Change Password button for windows authenticated Security
    If User.Identity.Name.Contains("\") Then
      btnChangePwd.Visible = False
      btnChangePwd_label.Visible = False
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

    'TODO workflow link
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

  Protected Sub BtnToDoListClick(sender As Object, e As ImageClickEventArgs) Handles btnToDoList.Click
    Response.Redirect("~/PendingSteps.aspx")
  End Sub

  Protected Sub BtnChangePwdClick(sender As Object, e As ImageClickEventArgs) Handles btnChangePwd.Click
    Response.Redirect("~/Account/ChangePassword.aspx")
  End Sub

  Protected Sub BtnLogoutClick(sender As Object, e As ImageClickEventArgs) Handles btnLogout.Click

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
