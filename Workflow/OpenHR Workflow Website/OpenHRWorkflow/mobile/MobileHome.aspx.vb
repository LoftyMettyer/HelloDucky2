
Imports System
Imports System.Data
Imports System.Collections.Generic
Imports Utilities

Partial Class Home
  Inherits System.Web.UI.Page

  Private miImageCount As Int16
  Private miStepCount As Integer
  Private mobjConfig As New Config
  Const wfCategoryKey As String = "WORKFLOW"

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ctlFormHtmlGenericControl As HtmlGenericControl
    Dim ctlFormHtmlInputText As HtmlInputText
    Dim ctlFormImageButton As ImageButton   ' Button
    Dim strConn As String
    Dim objGeneral As New General
    Dim sMessage As String = ""
    Dim drLayouts As System.Data.SqlClient.SqlDataReader
    Dim drElements As System.Data.SqlClient.SqlDataReader
    Dim sImageFileName As String = ""
    Dim sql As String
    Dim command As SqlClient.SqlCommand
    Dim reader As IDataReader
    miImageCount = 0

    Try
      mobjConfig.Mob_Initialise()
      Session("Server") = mobjConfig.Server
      Session("Database") = mobjConfig.Database
      Session("Login") = mobjConfig.Login
      Session("Password") = mobjConfig.Password
      Session("WorkflowURL") = mobjConfig.WorkflowURL

    Catch ex As Exception

    End Try

    ' Establish Connection
    strConn = CType(("Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
                     ";Initial Catalog=" & Session("Database") & _
                     ";Integrated Security=false;User ID=" & Session("Login") & _
                     ";Password=" & Session("Password") & _
                     ";Pooling=false"), String)
    'strConn = "Application Name=OpenHR Workflow;Data Source=.\sqlexpress;Initial Catalog=hrprostd43;Integrated Security=false;User ID=sa;Password=asr;Pooling=false"

    Dim myConnection As New SqlClient.SqlConnection(strConn)
    myConnection.Open()

    ' Create command
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

        Dim imageControl As New System.Web.UI.WebControls.Image

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
    strConn = CType(("Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
                     ";Initial Catalog=" & Session("Database") & _
                     ";Integrated Security=false;User ID=" & Session("Login") & _
                     ";Password=" & Session("Password") & _
                     ";Pooling=false"), String)
    'strConn = "Application Name=OpenHR Workflow;Data Source=.\sqlexpress;Initial Catalog=hrprostd43;Integrated Security=false;User ID=sa;Password=asr;Pooling=false"

    myConnection = New SqlClient.SqlConnection(strConn)
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
            ctlFormHtmlGenericControl = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), HtmlGenericControl)  'New Label
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
    If NullSafeString(Session("LoginKey")).IndexOf("\") >= 0 Then
      btnChangePwd.Visible = False
      btnChangePwd_label.Visible = False
    End If

    Dim groupId As Integer
    Dim fUserHasRunPermission As Boolean

    If Session("UserGroupID") <> "0" Then groupId = CInt(Session("UserGroupID"))

    If groupId <> 0 Then

      ' get the run permissions for workflow for this user group.
      sql = "SELECT  [i].[itemKey], [p].[permitted]" & _
                           " FROM [ASRSysGroupPermissions] p" & _
                           " JOIN [ASRSysPermissionItems] i ON [p].[itemID] = [i].[itemID]" & _
                           " WHERE [p].[itemID] IN (" & _
                               " SELECT [itemID] FROM [ASRSysPermissionItems]	" & _
                                " WHERE [categoryID] = (SELECT [categoryID] FROM [ASRSysPermissionCategories] WHERE [categoryKey] = '" & wfCategoryKey & "')) " & _
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
        sImageFileName = LoadPicture(NullSafeInteger(reader("PictureID")), sMessage)
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
    If fUserHasRunPermission Then CountPendingWFSteps()

  End Sub


  Private Sub CountPendingWFSteps()
    ' Update number of OS workflows
    Dim iWFCount As Integer = CheckPendingSteps()
    If iWFCount > 0 Then
      lblWFCount.InnerText = CStr(iWFCount)
      pnlWFCount.Style.Add("visibility", "visible")
    Else
      pnlWFCount.Style.Add("visibility", "hidden")
    End If
  End Sub

  Private Function LoadPicture(ByVal piPictureID As Int32, _
    ByRef psErrorMessage As String) As String

    Dim strConn As String
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim cmdSelect As System.Data.SqlClient.SqlCommand
    Dim dr As System.Data.SqlClient.SqlDataReader
    Dim sImageFileName As String
    Dim sImageFilePath As String
    Dim sImageWebPath As String
    Dim sTempName As String
    Dim fs As System.IO.FileStream
    Dim bw As System.IO.BinaryWriter
    Dim iBufferSize As Integer = 100
    Dim outByte(iBufferSize - 1) As Byte
    Dim retVal As Long
    Dim startIndex As Long = 0
    Dim sExtension As String = ""
    Dim iIndex As Integer
    Dim sName As String

    Try
      miImageCount = CShort(miImageCount + 1)

      psErrorMessage = ""
      LoadPicture = ""
      sImageFileName = ""
      sImageWebPath = "../pictures"
      sImageFilePath = Server.MapPath(sImageWebPath)

      strConn = CType(("Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
                       ";Initial Catalog=" & Session("Database") & _
                       ";Integrated Security=false;User ID=" & Session("Login") & _
                       ";Password=" & Session("Password") & _
                       ";Pooling=false"), String)
      'strConn = "Application Name=OpenHR Workflow;Data Source=.\sqlexpress;Initial Catalog=hrprostd43;Integrated Security=false;User ID=sa;Password=asr;Pooling=false"
      'strConn = "Application Name=OpenHR Workflow;Data Source=" & msServer & ";Initial Catalog=" & msDatabase & ";Integrated Security=false;User ID=" & msUser & ";Password=" & msPwd & ";Pooling=false"
      conn = New SqlClient.SqlConnection(strConn)
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
           "_" & miImageCount.ToString & _
           "_" & Date.Now.Ticks.ToString & _
           sExtension
          sTempName = sImageFilePath & "\" & sImageFileName

          ' Create a file to hold the output.
          fs = New System.IO.FileStream(sTempName, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
          bw = New System.IO.BinaryWriter(fs)

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

  Public Function WorkflowLink(ByVal pintWorkflowID As Integer) As String
    Dim sURL As String
    Dim sUser As String
    Dim sEncryptedString As String
    Dim objCrypt As New Crypt

    WorkflowLink = ""

    sURL = Session("WorkflowURL")
    If Len(sURL) = 0 Then
      Exit Function
    End If

    sUser = Session("Login")
    If Len(sUser) = 0 Then
      Exit Function
    End If

    ' For externally initiated workflows:
    '      plngInstance = -1 * workflowID
    '      plngStepID = -
    sEncryptedString = objCrypt.EncryptQueryString((-1 * pintWorkflowID), -1, sUser, _
        Session("Password"), _
        Session("Server"), _
        Session("Database"), _
        Session("LoginKey"), _
        Session("LoginPWD"))

    WorkflowLink = sURL & "?" & sEncryptedString

  End Function


  Public Function StepCount() As String
    StepCount = miStepCount
  End Function

  Private Function CheckPendingSteps() As Integer
    Dim iLoop As String
    Dim strConn As String
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim cmdSteps As System.Data.SqlClient.SqlCommand
    Dim rstSteps As System.Data.SqlClient.SqlDataReader

    Session("In") = "True"

    ' Open a connection to the database.
    strConn = "Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
        ";Initial Catalog=" & Session("Database") & _
        ";Integrated Security=false;User ID=" & Session("Login") & _
        ";Password=" & Session("Password") & _
        ";Pooling=false"

    conn = New SqlClient.SqlConnection(strConn)
    conn.Open()

    cmdSteps = New SqlClient.SqlCommand
    cmdSteps.CommandText = "spASRSysMobileCheckPendingWorkflowSteps"
    cmdSteps.Connection = conn
    cmdSteps.CommandType = CommandType.StoredProcedure

    cmdSteps.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
    cmdSteps.Parameters("@psKeyParameter").Value = Session("LoginKey")

    rstSteps = cmdSteps.ExecuteReader

    iLoop = 0

    While (rstSteps.Read)
      iLoop = iLoop + 1
    End While

    miStepCount = iLoop

    rstSteps.Close()
    cmdSteps.Dispose()

    CheckPendingSteps = miStepCount

  End Function

  Protected Sub btnChangePwd_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnChangePwd.Click
    Response.Redirect("MobileChangePassword.aspx")
  End Sub

  Protected Sub btnLogout_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnLogout.Click
    LogoutAuthenticatedUser()
  End Sub

  Protected Sub btnToDoList_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnToDoList.Click
    Response.Redirect("MobilePendingSteps.aspx")
  End Sub

  Private Sub LogoutAuthenticatedUser()
    ' Remove the cookie from cookies collection.

    FormsAuthentication.SignOut()
    Session.Abandon()

    ' clear authentication cookie
    Dim cookie1 As HttpCookie = New HttpCookie(FormsAuthentication.FormsCookieName, "")
    cookie1.Expires = DateTime.Now.AddYears(-1)
    Response.Cookies.Add(cookie1)

    ' clear session cookie (not necessary for your current problem but i would recommend you do it anyway)
    Dim cookie2 As HttpCookie = New HttpCookie("ASP.NET_SessionId", "")
    cookie2.Expires = DateTime.Now.AddYears(-1)
    Response.Cookies.Add(cookie2)

    'FormsAuthentication.RedirectToLoginPage()
    Response.Redirect("~/MobileLogin.aspx")
  End Sub


End Class
