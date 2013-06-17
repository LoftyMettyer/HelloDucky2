Imports System
Imports System.Data
Imports System.Collections.Generic
Imports Utilities

Partial Class Home
  Inherits Page

  Private _imageCount As Int16

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    Dim ctlFormHtmlGenericControl As HtmlGenericControl
    Dim ctlFormHtmlInputText As HtmlInputText
    Dim ctlFormImageButton As ImageButton
    Dim objGeneral As New General
    Dim sMessage As String = ""
    Dim drLayouts As SqlClient.SqlDataReader
    Dim drElements As SqlClient.SqlDataReader
    Dim sImageFileName As String = ""
    Dim sql As String
    Dim command As SqlClient.SqlCommand
    Dim reader As IDataReader

    _imageCount = 0

    ' Establish Connection
    Dim myConnection As New SqlClient.SqlConnection(Configuration.ConnectionString)
    myConnection.Open()
    Dim myCommand As New SqlClient.SqlCommand("select * from tbsys_mobileformlayout where ID = 1", myConnection)

    ' Create a DataReader to ferry information back from the database
    drLayouts = myCommand.ExecuteReader()
    drLayouts.Read()

    For i As Integer = 1 To 3

      Dim prefix As String = String.Empty
      Dim control As HtmlGenericControl = Nothing

      Select Case i
        Case 1
          prefix = "Header"
          control = pnlHeader
        Case 2
          prefix = "Main"
          control = ScrollerFrame
        Case 3
          prefix = "Footer"
          control = pnlFooter
      End Select

      If Not IsDBNull(drLayouts(prefix & "BackColor")) Then
        control.Style("Background-color") = objGeneral.GetHTMLColour(CInt(drLayouts(prefix & "BackColor")))
      End If

      If Not IsDBNull(drLayouts(prefix & "PictureID")) Then
        control.Style("Background-image") = LoadPicture(CInt(drLayouts(prefix & "PictureID")), sMessage)
        control.Style("background-repeat") = objGeneral.BackgroundRepeat(CShort(drLayouts(prefix & "PictureLocation")))
        control.Style("background-position") = objGeneral.BackgroundPosition(CShort(drLayouts(prefix & "PictureLocation")))
      End If

      'Header Image
      If i = 1 AndAlso Not IsDBNull(drLayouts("HeaderLogoID")) Then

        Dim imageControl As New Image

        With imageControl
          .Style("position") = "absolute"

          If NullSafeInteger(drLayouts("HeaderLogoVerticalOffsetBehaviour")) = 0 Then
            .Style("top") = Unit.Pixel(NullSafeInteger(drLayouts("HeaderLogoVerticalOffset"))).ToString
          Else
            .Style("bottom") = Unit.Pixel(NullSafeInteger(drLayouts("HeaderLogoVerticalOffset"))).ToString
          End If

          If NullSafeInteger(drLayouts("HeaderLogoHorizontalOffsetBehaviour")) = 0 Then
            .Style("left") = Unit.Pixel(NullSafeInteger(drLayouts("HeaderLogoHorizontalOffset"))).ToString
          Else
            .Style("right") = Unit.Pixel(NullSafeInteger(drLayouts("HeaderLogoHorizontalOffset"))).ToString
          End If

          .BackColor = System.Drawing.Color.Transparent
          .ImageUrl = LoadPicture(NullSafeInteger(drLayouts("HeaderLogoID")), sMessage)
          .Height() = Unit.Pixel(NullSafeInteger(drLayouts("HeaderLogoHeight")))
          .Width() = Unit.Pixel(NullSafeInteger(drLayouts("HeaderLogoWidth")))
          .Style.Add("z-index", "1")
        End With

        pnlHeader.Controls.Add(imageControl)
      End If

    Next

    Dim homeItemStyles = New Dictionary(Of String, String)
    homeItemStyles.Add("font-family", NullSafeString(drLayouts("HomeItemFontName")))
    homeItemStyles.Add("font-size", NullSafeString(drLayouts("HomeItemFontSize")) & "pt")
    homeItemStyles.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(drLayouts("HomeItemFontBold"))), "bold", "normal"))
    homeItemStyles.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(drLayouts("HomeItemFontItalic"))), "italic", "normal"))

    ' Close the connection (will automatically close the reader)
    myConnection.Close()
    drLayouts.Close()

    ' ======================== NOW FOR THE INDIVIDUAL ELEMENTS  ====================================

    ' Establish Connection
    myConnection = New SqlClient.SqlConnection(Configuration.ConnectionString)
    myConnection.Open()

    ' Create command
    myCommand = New SqlClient.SqlCommand("select * from tbsys_mobileformelements where form = 2", myConnection)

    ' Create a DataReader to ferry information back from the database
    drElements = myCommand.ExecuteReader()

    'Iterate through the results
    While drElements.Read()
      Select Case CInt(drElements("Type"))

        Case 0 ' Button

          If NullSafeString(drElements("Name")).Length > 0 Then
            ctlFormImageButton = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), ImageButton)

            With ctlFormImageButton
              sImageFileName = LoadPicture(NullSafeInteger(drElements("pictureID")), sMessage)
              .ImageUrl = sImageFileName
              .Font.Name = NullSafeString(drElements("FontName"))
              .Font.Size = FontUnit.Parse(NullSafeString(drElements("FontSize")))
              .Font.Bold = NullSafeBoolean(NullSafeBoolean(drElements("FontBold")))
              .Font.Italic = NullSafeBoolean(NullSafeBoolean(drElements("FontItalic")))
            End With

            ' Footer text
            If NullSafeString(drElements("Caption")).Length > 0 Then
              ctlFormHtmlGenericControl = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name")) & "_label"), HtmlGenericControl)
              With ctlFormHtmlGenericControl
                .Style("word-wrap") = "break-word"
                .Style("overflow") = "auto"
                .Style.Add("z-index", "1")
                .InnerText = NullSafeString(drElements("caption"))
                .Style.Add("background-color", "Transparent")
                .Style.Add("font-family", "Verdana")
                .Style.Add("font-size", "6pt")
                .Style.Add("font-weight", "normal")
                .Style.Add("font-style", "normal")
              End With
            End If
          End If

        Case 2 ' Label
          If NullSafeString(drElements("Name")).Length > 0 Then
            ctlFormHtmlGenericControl = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), HtmlGenericControl)
            With ctlFormHtmlGenericControl
              .Style("word-wrap") = "break-word"
              .Style("overflow") = "auto"
              .Style("text-align") = "left"
              .Style.Add("z-index", "1")
              .InnerText = NullSafeString(drElements("caption"))
              .Style.Add("color", objGeneral.GetHTMLColour(NullSafeInteger(drElements("ForeColor"))))
              .Style.Add("font-family", NullSafeString(drElements("FontName")))
              .Style.Add("font-size", NullSafeString(drElements("FontSize")) & "pt")
              .Style.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(drElements("FontBold"))), "bold", "normal"))
              .Style.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(drElements("FontItalic"))), "italic", "normal"))
            End With

          End If


        Case 3 ' Input value - character
          If NullSafeString(drElements("Name")).Length > 0 Then

            ctlFormHtmlInputText = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), HtmlInputText)
            ctlFormHtmlInputText.Style("resize") = "none"
            ctlFormHtmlInputText.Style.Add("border-style", "solid")
            ctlFormHtmlInputText.Style.Add("border-width", "1")
            ctlFormHtmlInputText.Style.Add("border-color", objGeneral.GetHTMLColour(5730458))
            ctlFormHtmlInputText.Style.Add("color", objGeneral.GetHTMLColour(NullSafeInteger(drElements("ForeColor"))))
            ctlFormHtmlInputText.Style.Add("font-family", NullSafeString(drElements("FontName")))
            ctlFormHtmlInputText.Style.Add("font-size", NullSafeString(drElements("FontSize")) & "pt")
            ctlFormHtmlInputText.Style.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(drElements("FontBold"))), "bold", "normal"))
            ctlFormHtmlInputText.Style.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(drElements("FontItalic"))), "italic", "normal"))
          End If

      End Select

    End While
    drElements.Close()

    ' Disable the Change Password button for windows authenticated users
    If User.Identity.Name.Contains("\") Then
      btnChangePwd.Visible = False
      btnChangePwd_label.Visible = False
    End If

    Dim groupId As Integer
    Dim fUserHasRunPermission As Boolean

    'TODO close your session come straight to this page and this value is not populated!!!!
    If Session("UserGroupID") <> "0" Then groupId = CInt(Session("UserGroupID"))

    If groupId <> 0 Then

      ' get the run permissions for workflow for this user group.
      sql = "SELECT  [i].[itemKey], [p].[permitted]" & _
                           " FROM [ASRSysGroupPermissions] p" & _
                           " JOIN [ASRSysPermissionItems] i ON [p].[itemID] = [i].[itemID]" & _
                           " WHERE [p].[itemID] IN (" & _
                               " SELECT [itemID] FROM [ASRSysPermissionItems]	" & _
                                " WHERE [categoryID] = (SELECT [categoryID] FROM [ASRSysPermissionCategories] WHERE [categoryKey] = 'WORKFLOW')) " & _
                                " AND [groupName] = (SELECT [Name] FROM [ASRSysGroups] WHERE [ID] = " & groupId.ToString & ")"
      Try
        command = New SqlClient.SqlCommand(sql, myConnection)
        reader = command.ExecuteReader()

        While reader.Read()
          Select Case reader("itemKey")
            Case "RUN"
              fUserHasRunPermission = (reader("permitted") = True)
          End Select
        End While

        reader.Close()
      Catch ex As Exception

      End Try

    End If

    If fUserHasRunPermission Then

      sql = "select w.Id, w.Name, w.PictureID from tbsys_mobilegroupworkflows gw inner join tbsys_workflows w on gw.WorkflowID = w.ID where gw.UserGroupID = " & groupId & " and w.enabled = 1 order by gw.Pos ASC"
      command = New SqlClient.SqlCommand(sql, myConnection)

      reader = command.ExecuteReader()

      ' Create the holding table for the list of workflows.
      Dim table = New Table
      table.Style.Add("width", "100%")

      'Iterate through the results
      Dim itemCount As Integer
      While reader.Read()

        ' Create a row to contain this pending step...
        Dim row = New TableRow
        row.Style.Add("width", "100%")
        row.Attributes.Add("onclick", "window.open('" & WorkflowLink(CInt(reader("ID"))) & "');")

        ' Create a cell to contain the workflow icon
        Dim cell = New TableCell  ' Image cell
        cell.Style.Add("width", "57px")

        Dim image = New Image
        If NullSafeInteger(reader("pictureID")) = 0 Then
          sImageFileName = "~/Images/Connected48.png"
        Else
          sImageFileName = LoadPicture(NullSafeInteger(reader("pictureID")), sMessage)
        End If
        image.ImageUrl = sImageFileName
        image.Height() = Unit.Pixel(57)
        image.Width() = Unit.Pixel(57)

        ' add ImageButton to cell
        cell.Controls.Add(image)

        ' Add cell to row
        row.Cells.Add(cell)

        ' Create a cell to contain the workflow name and description
        cell = New TableCell
        Dim label = New Label ' Workflow name text
        label.Text = CStr(reader("Name"))
        For Each item In homeItemStyles
          label.Style.Add(item.Key, item.Value)
        Next

        cell.Controls.Add(label)

        ' Add cell to row, and row to table.
        row.Cells.Add(cell)

        table.Rows.Add(row)

        itemCount += 1
      End While
      reader.Close()
      pnlWFList.Controls.Add(table)

      hdnItemCount.Value = CStr(itemCount)
    End If

    ' close sql connection
    myConnection.Close()

    ' Update the wf steps count
    If fUserHasRunPermission Then CountPendingWfSteps()

  End Sub


  Private Sub CountPendingWfSteps()
    ' Update number of OS workflows
    Dim count As Integer = CheckPendingSteps()
    If count > 0 Then
      lblWFCount.InnerText = CStr(count)
      pnlWFCount.Style.Add("visibility", "visible")
    Else
      pnlWFCount.Style.Add("visibility", "hidden")
    End If
  End Sub

  Private Function LoadPicture(ByVal piPictureID As Int32, ByRef psErrorMessage As String) As String

    Dim conn As SqlClient.SqlConnection
    Dim cmdSelect As SqlClient.SqlCommand
    Dim dr As SqlClient.SqlDataReader
    Dim sImageFileName As String
    Dim sImageFilePath As String
    Dim sImageWebPath As String
    Dim sTempName As String
    Dim fs As IO.FileStream
    Dim bw As IO.BinaryWriter
    Dim iBufferSize As Integer = 100
    Dim outByte(iBufferSize - 1) As Byte
    Dim retVal As Long
    Dim startIndex As Long
    Dim sExtension As String = ""
    Dim iIndex As Integer
    Dim sName As String

    Try
      _imageCount = CShort(_imageCount + 1)

      psErrorMessage = ""
      LoadPicture = ""
      sImageFileName = ""
      sImageWebPath = "~/pictures"
      sImageFilePath = Server.MapPath(sImageWebPath)

      conn = New SqlClient.SqlConnection(Configuration.ConnectionString)
      conn.Open()

      cmdSelect = New SqlClient.SqlCommand
      cmdSelect.CommandText = "spASRGetPicture"
      cmdSelect.Connection = conn
      cmdSelect.CommandType = CommandType.StoredProcedure
      cmdSelect.CommandTimeout = 30 ' miSubmissionTimeoutInSeconds

      cmdSelect.Parameters.Add("@piPictureID", SqlDbType.Int).Direction = ParameterDirection.Input
      cmdSelect.Parameters("@piPictureID").Value = piPictureID

      Try
        dr = cmdSelect.ExecuteReader(CommandBehavior.SequentialAccess)

        Do While dr.Read
          sName = NullSafeString(dr("name"))
          iIndex = sName.LastIndexOf(".")
          If iIndex >= 0 Then
            sExtension = sName.Substring(iIndex)
          End If

          sImageFileName = Session.SessionID().ToString & _
           "_" & _imageCount.ToString & _
           "_" & Date.Now.Ticks.ToString & _
           sExtension
          sTempName = sImageFilePath & "\" & sImageFileName

          ' Create a file to hold the output.
          fs = New IO.FileStream(sTempName, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
          bw = New IO.BinaryWriter(fs)

          ' Reset the starting byte for a new BLOB.
          startIndex = 0

          ' Read bytes into outbyte() and retain the number of bytes returned.
          retVal = dr.GetBytes(1, startIndex, outByte, 0, iBufferSize)

          ' Continue reading and writing while there are bytes beyond the size of the buffer.
          Do While retVal = iBufferSize
            bw.Write(outByte)
            bw.Flush()

            ' Reposition the start index to the end of the last buffer and fill the buffer.
            startIndex += iBufferSize
            retVal = dr.GetBytes(1, startIndex, outByte, 0, iBufferSize)
          Loop

          ' Write the remaining buffer.
          bw.Write(outByte)
          bw.Flush()

          ' Close the output file.
          bw.Close()
          fs.Close()
        Loop

        dr.Close()
        cmdSelect.Dispose()

        ' Ensure URL encoding doesn't stuff up the picture name, so encode the % character as %25.
        LoadPicture = sImageWebPath & "/" & sImageFileName

      Catch ex As Exception
        LoadPicture = ""
        psErrorMessage = ex.Message

      Finally
        conn.Close()
        conn.Dispose()
      End Try
    Catch ex As Exception
      LoadPicture = ""
      psErrorMessage = ex.Message
    End Try
  End Function

  Public Function WorkflowLink(ByVal workflowID As Integer) As String

    Dim objCrypt As New Crypt

    If Configuration.WorkflowUrl.Length = 0 Then
      Return ""
    End If

    If Configuration.Login.Length = 0 Then
      Return ""
    End If

    ' For externally initiated workflows:
    '      plngInstance = -1 * workflowID
    '      plngStepID = -

    'TODO check
    Dim sEncryptedString As String = objCrypt.EncryptQueryString((-1 * workflowID), -1, _
        Configuration.Login, _
        Configuration.Password, _
        Configuration.Server, _
        Configuration.Database, _
        User.Identity.Name, _
        "")

    Return Configuration.WorkflowUrl & "?" & sEncryptedString

  End Function

  Private Function CheckPendingSteps() As Integer

    Dim conn As SqlClient.SqlConnection
    Dim cmd As SqlClient.SqlCommand
    Dim dr As SqlClient.SqlDataReader

    ' Open a connection to the database.
    conn = New SqlClient.SqlConnection(Configuration.ConnectionString)
    conn.Open()

    cmd = New SqlClient.SqlCommand
    cmd.CommandText = "spASRSysMobileCheckPendingWorkflowSteps"
    cmd.Connection = conn
    cmd.CommandType = CommandType.StoredProcedure

    cmd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
    cmd.Parameters("@psKeyParameter").Value = User.Identity.Name

    dr = cmd.ExecuteReader

    Dim count As Integer
    While dr.Read
      count += 1
    End While
    dr.Close()
    cmd.Dispose()
    'TODO clean up
    Return count

  End Function

  Protected Sub BtnToDoListClick(sender As Object, e As ImageClickEventArgs) Handles btnToDoList.Click
    Response.Redirect("~/Mobile/MobilePendingSteps.aspx")
  End Sub

  Protected Sub BtnChangePwdClick(sender As Object, e As ImageClickEventArgs) Handles btnChangePwd.Click
    Response.Redirect("~/Mobile/MobileChangePassword.aspx")
  End Sub

  Protected Sub BtnLogoutClick(sender As Object, e As ImageClickEventArgs) Handles btnLogout.Click
    LogoutAuthenticatedUser()
  End Sub

  Private Sub LogoutAuthenticatedUser()
    ' Remove the cookie from cookies collection.

    FormsAuthentication.SignOut()
    Session.Abandon()

    ' clear authentication cookie
    Dim cookie As HttpCookie = New HttpCookie(FormsAuthentication.FormsCookieName, "")
    cookie.Expires = DateTime.Now.AddYears(-1)
    Response.Cookies.Add(cookie)

    ' clear session cookie (not necessary for your current problem but i would recommend you do it anyway)
    Dim cookie2 As HttpCookie = New HttpCookie("ASP.NET_SessionId", "")
    cookie2.Expires = DateTime.Now.AddYears(-1)
    Response.Cookies.Add(cookie2)

    Response.Redirect("~/MobileLogin.aspx")
  End Sub


End Class
