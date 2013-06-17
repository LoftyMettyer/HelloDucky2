Imports System.Data
Imports Utilities
Imports System.Web.Security

Partial Class MobileLogin
  Inherits Page

  Private _miImageCount As Int16
  Private ReadOnly _mobjConfig As New Config
  Private _miSubmissionTimeoutInSeconds As Int32

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    Dim fAuthenticated As Boolean
    ' is there is a remember me cookie, get the user name out of it and bypass the login screen.
    Dim cookieUser As String = HttpContext.Current.User.Identity.Name

    If Not (cookieUser = "" Or cookieUser = "sqluser") Then
      Session("LoginKey") = cookieUser
      fAuthenticated = True
    Else
      Session("LoginKey") = Nothing
      fAuthenticated = False
    End If

    If fAuthenticated Then
      ' forward on to requested page
      FormsAuthentication.RedirectFromLoginPage("", True)
    End If

  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

    Dim ctlFormHtmlGenericControl As HtmlGenericControl
    Dim ctlFormHtmlInputText As HtmlInputText
    Dim ctlFormImageButton As ImageButton   ' Button
    Dim strConn As String
    Dim objGeneral As New General
    Dim sMessage As String = ""
    Dim drLayouts As SqlClient.SqlDataReader
    Dim drElements As SqlClient.SqlDataReader
    Dim sImageFileName As String

    _miImageCount = 0
    If Not IsPostBack Then

      Try
        _mobjConfig.Mob_Initialise()
        Session("Server") = _mobjConfig.Server
        Session("Database") = _mobjConfig.Database
        Session("Login") = _mobjConfig.Login
        Session("Password") = _mobjConfig.Password
        Session("WorkflowURL") = _mobjConfig.WorkflowURL

      Catch ex As Exception
        sMessage = "Error" & vbCrLf & ex.Message
        lblMsgBox.InnerText = sMessage
        pnlGreyOut.Style.Add("visibility", "visible")
        pnlMsgBox.Style.Add("visibility", "visible")
      End Try

      If User.Identity.IsAuthenticated Then
        ' skip this page!
        submitLoginDetails()
      End If

      Try
        ' Establish Connection
        strConn = CType(("Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
                         ";Initial Catalog=" & Session("Database") & _
                         ";Integrated Security=false;User ID=" & Session("Login") & _
                         ";Password=" & Session("Password") & _
                         ";Pooling=false"), String)

        Dim myConnection As New SqlClient.SqlConnection(strConn)
        myConnection.Open()

        ' Create command
        Dim myCommand As New SqlClient.SqlCommand("select * from tbsys_mobileformlayout where ID = 1", myConnection)

        ' Create a DataReader to ferry information back from the database
        drLayouts = myCommand.ExecuteReader()

        If drLayouts.Read() Then
          For iPanelID As Integer = 1 To 3
            Dim prefix As String = String.Empty
            Dim control As HtmlGenericControl = Nothing

            Select Case iPanelID
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
            If iPanelID = 1 AndAlso Not IsDBNull(drLayouts("HeaderLogoID")) Then
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

                .BackColor = Drawing.Color.Transparent
                .ImageUrl = LoadPicture(NullSafeInteger(drLayouts("HeaderLogoID")), sMessage)
                .Height() = Unit.Pixel(NullSafeInteger(drLayouts("HeaderLogoHeight")))
                .Width() = Unit.Pixel(NullSafeInteger(drLayouts("HeaderLogoWidth")))
                .Style.Add("z-index", "1")
              End With

              pnlHeader.Controls.Add(imageControl)
            End If
          Next
        End If

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

        myConnection = New SqlClient.SqlConnection(strConn)
        myConnection.Open()

        ' Create command
        myCommand = New SqlClient.SqlCommand("select * from tbsys_mobileformelements where form = 1", myConnection)

        ' Create a DataReader to ferry information back from the database
        drElements = myCommand.ExecuteReader()

        'Iterate through the results
        While drElements.Read()
          Select Case CInt(drElements("Type"))

            Case 0 ' Button
              If NullSafeString(drElements("Name")).Length > 0 Then
                ctlFormImageButton = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), ImageButton)

                With ctlFormImageButton
                  sImageFileName = LoadPicture(NullSafeInteger(drElements("PictureID")), sMessage)
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
                ctlFormHtmlInputText.Style.Add("font-family", NullSafeString(drElements("FontName")))
                ctlFormHtmlInputText.Style.Add("font-size", NullSafeString(drElements("FontSize")) & "pt")
                ctlFormHtmlInputText.Style.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(drElements("FontBold"))), "bold", "normal"))
                ctlFormHtmlInputText.Style.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(drElements("FontItalic"))), "italic", "normal"))
              End If

          End Select
        End While

        ' Close the connection (will automatically close the reader)
        myConnection.Close()
        drElements.Close()

      Catch ex As Exception   ' conn creation 
        sMessage = "Error creating SQL connection:" & vbCrLf & vbCrLf & _
        ex.Message.ToString & vbCrLf & vbCrLf & _
        "Contact your system administrator."
      End Try

      If sMessage.Length > 0 Then
        Session("message") = sMessage
        Response.Redirect("Message.aspx")
      End If
    End If

  End Sub

  Private Function loadPicture(ByVal piPictureID As Int32, _
    ByRef psErrorMessage As String) As String

    Dim strConn As String
    Dim conn As SqlClient.SqlConnection
    Dim cmdSelect As SqlClient.SqlCommand
    Dim dr As SqlClient.SqlDataReader
    Dim sImageFileName As String
    Dim sImageFilePath As String
    Dim sImageWebPath As String
    Dim sTempName As String
    Dim fs As IO.FileStream
    Dim bw As IO.BinaryWriter
    Const iBufferSize As Integer = 100
    Dim outByte(iBufferSize - 1) As Byte
    Dim retVal As Long
    Dim startIndex As Long
    Dim sExtension As String = ""
    Dim iIndex As Integer
    Dim sName As String

    Try
      _miImageCount = CShort(_miImageCount + 1)

      psErrorMessage = ""
      sImageFileName = ""
      sImageWebPath = "pictures"
      sImageFilePath = Server.MapPath(sImageWebPath)

      strConn = CType(("Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
                       ";Initial Catalog=" & Session("Database") & _
                       ";Integrated Security=false;User ID=" & Session("Login") & _
                       ";Password=" & Session("Password") & _
                       ";Pooling=false"), String)

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
           "_" & _miImageCount.ToString & _
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

  Protected Sub btnLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLogin.Click
    submitLoginDetails()
  End Sub

  Private Sub submitLoginDetails()

    Dim strConn As String
    Dim conn As SqlClient.SqlConnection
    Dim cmdCheck As SqlClient.SqlCommand
    Dim dr As SqlClient.SqlDataReader
    Dim sMessage As String = ""
    Dim aUserDetails() As String
    Dim fAuthenticated As Boolean = False
    Dim sAuthUser As String = ""

    _miSubmissionTimeoutInSeconds = _mobjConfig.SubmissionTimeoutInSeconds

    If NullSafeString(Session("loginKey")) <> vbNullString Then fAuthenticated = True

    If Not fAuthenticated Then
      Dim fPersistCookie As Boolean = chkRememberPwd.Checked

      ' Get the password stored for this user, decrypt it, then compare it here.
      ' do a few quick checks before validating the user
      If txtUserName.Value.Trim.Length = 0 Then
        sMessage = "No Login entered."
      End If

      ' Perform system Lock Check first.
      strConn = CType(("Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
                             ";Initial Catalog=" & Session("Database") & _
                             ";Integrated Security=false;User ID=" & Session("Login") & _
                             ";Password=" & Session("Password") & _
                             ";Pooling=false"), String)

      Try
        conn = New SqlClient.SqlConnection(strConn)
        conn.Open()

        ' Check if the database is locked.
        cmdCheck = New SqlClient.SqlCommand
        cmdCheck.CommandText = "sp_ASRLockCheck"
        cmdCheck.Connection = conn
        cmdCheck.CommandType = CommandType.StoredProcedure
        cmdCheck.CommandTimeout = _miSubmissionTimeoutInSeconds

        dr = cmdCheck.ExecuteReader()

        While dr.Read
          If NullSafeInteger(dr("priority")) <> 3 Then
            ' Not a read-only lock.
            sMessage = "Database locked." & vbCrLf & "Contact your system administrator."
            Exit While
          End If
        End While

        dr.Close()
        cmdCheck.Dispose()
        conn.Close()
      Catch ex As Exception
        sMessage = "Unable to perform system lock check"
      End Try

      ' Continue with authentication
      If sMessage.Length = 0 Then

        If txtUserName.Value.IndexOf("\") > 0 Then
          ' ============ WINDOWS AUTHENTICATION =============
          ' The presence of a backslash indicates this is a windows user. 
          ' validate the user against AD
          aUserDetails = txtUserName.Value.Split(CType("\", Char))
          Try
            authenticateUser(aUserDetails(0), aUserDetails(1), txtPassword.Value)
            sAuthUser = txtUserName.Value
            fAuthenticated = True
            Session("LoginKey") = txtUserName.Value
            Session("RememberPWD") = chkRememberPwd.Value
          Catch ex As Exception
            'lblError.Text = ex.Message
            'lblError.Visible = True
            fAuthenticated = False
            sMessage = ex.Message
            Session("LoginKey") = Nothing
            Session("RememberPWD") = Nothing
          End Try
        Else
          ' ============ SQL AUTHENTICATION =============
          ' SQL User - validate against the DB
          If sMessage.Length = 0 Then
            strConn = CType(("Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
                 ";Initial Catalog=" & Session("Database") & _
                 ";Integrated Security=false;User ID=" & txtUserName.Value & _
                 ";Password=" & txtPassword.Value & _
                 ";Pooling=false"), String)
            conn = New SqlClient.SqlConnection(strConn)

            Try
              conn.Open()
              conn.Close()
              FormsAuthentication.SetAuthCookie(txtUserName.Value, fPersistCookie)
              sAuthUser = txtUserName.Value
              fAuthenticated = True
              Session("LoginKey") = txtUserName.Value
              Session("RememberPWD") = chkRememberPwd.Value
            Catch ex As Exception
              sMessage = ex.Message
              fAuthenticated = False
              Session("LoginKey") = Nothing
              Session("RememberPWD") = Nothing
            End Try
          End If
        End If
      End If
    End If

    If fAuthenticated And sMessage.Length = 0 Then
      Try
        strConn = CType(("Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
                         ";Initial Catalog=" & Session("Database") & _
                         ";Integrated Security=false;User ID=" & Session("Login") & _
                         ";Password=" & Session("Password") & _
                         ";Pooling=false"), String)

        conn = New SqlClient.SqlConnection(strConn)
        conn.Open()

        cmdCheck = New SqlClient.SqlCommand
        cmdCheck.CommandText = "spASRSysMobileCheckLogin"
        cmdCheck.Connection = conn
        cmdCheck.CommandType = CommandType.StoredProcedure

        cmdCheck.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        cmdCheck.Parameters("@psKeyParameter").Value = Session("LoginKey")

        cmdCheck.Parameters.Add("@piUserGroupID", SqlDbType.Int).Direction = ParameterDirection.Output

        cmdCheck.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

        cmdCheck.ExecuteNonQuery()

        sMessage = NullSafeString(cmdCheck.Parameters("@psMessage").Value())
        Session("UserGroupID") = NullSafeInteger(cmdCheck.Parameters("@piUserGroupID").Value())

        cmdCheck.Dispose()

      Catch ex As Exception
        sMessage = "Error :" & vbCrLf & _
        ex.Message.ToString() & vbCrLf & _
        "Contact your system administrator."
      End Try
    End If

    If sMessage.Length > 0 Then
      Session("message") = sMessage
      Session("nextPage") = "~/MobileLogin"
      lblMsgBox.InnerText = sMessage
      pnlGreyOut.Style.Add("visibility", "visible")
      pnlMsgBox.Style.Add("visibility", "visible")

      FormsAuthentication.SignOut()
      Session("LoginKey") = Nothing
      Session("nextPage") = "~/MobileLogin"

    Else
      If fAuthenticated Then
        ' Where were we heading before being asked for authentication?
        If FormsAuthentication.GetRedirectUrl(sAuthUser, False) = FormsAuthentication.DefaultUrl Then
          ' Nowhere! Go to the home page.
          Response.Redirect("~/mobile/MobileHome.aspx")
        Else
          'Go to the page originally specified by the client.
          FormsAuthentication.SignOut()
          HttpContext.Current.Response.Redirect(FormsAuthentication.GetRedirectUrl(sAuthUser, False))
        End If
      End If
    End If

  End Sub
  Protected Sub btnRegister_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnRegister.Click
    ' set a temporary user name so that we can move past the authentication page.
    FormsAuthentication.SetAuthCookie("sqluser", False)
    Response.Redirect("~/mobile/MobileRegistration.aspx")
  End Sub

  Protected Sub btnForgotPwd_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnForgotPwd.Click
    ' set a temporary user name so that we can move past the authentication page.
    FormsAuthentication.SetAuthCookie("sqluser", False)
    Response.Redirect("~/mobile/MobileForgottenLogin.aspx")
  End Sub

  Private Sub authenticateUser(domainName As String, userName As String, password As String)
    ' Path to youR LDAP directory server.
    ' Contact your network administrator to obtain a valid path.

    Dim adPath As String = "LDAP://" & ConfigurationManager.AppSettings("DefaultActiveDirectoryServer")

    Dim adAuth As New clsActiveDirectoryValidator(adPath)
    If True = adAuth.IsAuthenticated(domainName, userName, password) Then
      ' Create the authetication ticket
      Dim authTicket As New FormsAuthenticationTicket(1, userName, DateTime.Now, DateTime.Now.AddMinutes(20), False, "")
      ' Now encrypt the ticket.
      Dim encryptedTicket As String = FormsAuthentication.Encrypt(authTicket)
      ' Create a cookie and add the encrypted ticket to the
      ' cookie as data.
      Dim authCookie As New HttpCookie(FormsAuthentication.FormsCookieName, encryptedTicket)
      ' Add the cookie to the outgoing cookies collection.
      HttpContext.Current.Response.Cookies.Add(authCookie)
      FormsAuthentication.SetAuthCookie(txtUserName.Value, chkRememberPwd.Checked)
      ' Redirect the user to the originally requested page
      ' HttpContext.Current.Response.Redirect(FormsAuthentication.GetRedirectUrl(userName, False))
    End If
  End Sub


End Class



