Imports System.Data
Imports Utilities
Imports System.Web.Security
Imports System.Globalization
Imports System.Threading

Partial Class MobileLogin
  Inherits System.Web.UI.Page

  Private miImageCount As Int16
  Private mobjConfig As New Config
  Private miSubmissionTimeoutInSeconds As Int32

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init

    Dim miInstanceID As Integer
    Dim miElementID As Integer
    Dim msServer As String
    Dim msUser As String
    Dim msPwd As String
    Dim sRedirectUrl As String
    Dim sRedirectWebPageName As String
    Dim objCrypt As New Crypt
    Dim fAuthenticated As Boolean = False

    ' Forms authentication can be overruled, and we do this for any existing workflow link.

    ' is there is a remember me cookie, get the user name out of it and bypass the login screen.
    Dim CustomerGUID As String = HttpContext.Current.User.Identity.Name

    If Not (CustomerGUID = "" Or CustomerGUID = "sqluser") Then
      Session("LoginKey") = CustomerGUID
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

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ctlForm_HTMLGenericControl As HtmlGenericControl
    Dim ctlForm_HtmlInputText As HtmlInputText
    Dim ctlForm_ImageButton As ImageButton   ' Button
    Dim strConn As String
    Dim objGeneral As New General
    Dim lngPanel_1_Height As Long = 57
    Dim lngPanel_2_Height As Long = 57
    Dim sMessage As String = ""
    Dim drLayouts As System.Data.SqlClient.SqlDataReader
    Dim drElements As System.Data.SqlClient.SqlDataReader
    Dim sImageFileName As String = ""
    Dim iTempHeight As Integer
    Dim iTempWidth As Integer

    miImageCount = 0
    If Not IsPostBack Then

      Try
        mobjConfig.Mob_Initialise()
        Session("Server") = mobjConfig.Server
        Session("Database") = mobjConfig.Database
        Session("Login") = mobjConfig.Login
        Session("Password") = mobjConfig.Password
        Session("WorkflowURL") = mobjConfig.WorkflowURL

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
        'strConn = "Application Name=OpenHR Workflow;Data Source=.\sqlexpress;Initial Catalog=hrprostd43;Integrated Security=false;User ID=sa;Password=asr;Pooling=false"

        Dim myConnection As New SqlClient.SqlConnection(strConn)
        myConnection.Open()

        ' Create command
        Dim myCommand As New SqlClient.SqlCommand("select * from tbsys_mobileformlayout where ID = 1", myConnection)

        ' Create a DataReader to ferry information back from the database
        drLayouts = myCommand.ExecuteReader()

        If drLayouts.Read() Then

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

        End If

        ' Close the connection (will automatically close the reader)
        myConnection.Close()
        drLayouts.Close()

        ' ======================== NOW FOR THE INDIVIDUAL ELEMENTS  ====================================

        ' Establish Connection
        strConn = "Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
          ";Initial Catalog=" & Session("Database") & _
          ";Integrated Security=false;User ID=" & Session("Login") & _
          ";Password=" & Session("Password") & _
          ";Pooling=false"
        'strConn = "Application Name=OpenHR Workflow;Data Source=.\sqlexpress;Initial Catalog=hrprostd43;Integrated Security=false;User ID=sa;Password=asr;Pooling=false"

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
                ctlForm_ImageButton = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), ImageButton)

                'ctlForm_ImageButton = New ImageButton
                With ctlForm_ImageButton
                  .Width = Unit.Pixel(40)
                  .Height = Unit.Pixel(40)
                  sImageFileName = LoadPicture(NullSafeInteger(drElements("PictureID")), sMessage)
                  .ImageUrl = sImageFileName
                  .Font.Name = NullSafeString(drElements("FontName"))
                  .Font.Size = FontUnit.Parse(NullSafeString(drElements("FontSize")))
                  .Font.Bold = NullSafeBoolean(NullSafeBoolean(drElements("FontBold")))
                  .Font.Italic = NullSafeBoolean(NullSafeBoolean(drElements("FontItalic")))
                End With

                ' Footer text
                If NullSafeString(drElements("Caption")).Length > 0 Then
                  ctlForm_HTMLGenericControl = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name")) & "_label"), HtmlGenericControl)
                  With ctlForm_HTMLGenericControl
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
                ctlForm_HTMLGenericControl = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), HtmlGenericControl)  'New Label
                With ctlForm_HTMLGenericControl
                  '.Style("position") = "absolute"

                  '' Vertical Offset
                  'If NullSafeInteger(drElements("VerticalOffsetBehaviour")) = 0 Then
                  '  .Style("top") = Unit.Pixel(NullSafeInteger(drElements("VerticalOffset"))).ToString
                  'Else
                  '  .Style("bottom") = Unit.Pixel(NullSafeInteger(drElements("VerticalOffset"))).ToString
                  'End If

                  ' horizontal position
                  If NullSafeInteger(drElements("HorizontalOffsetBehaviour")) = 0 Then
                    .Style("left") = Unit.Pixel(NullSafeInteger(drElements("HorizontalOffset"))).ToString
                  Else
                    .Style("right") = Unit.Pixel(NullSafeInteger(drElements("HorizontalOffset"))).ToString
                  End If

                  .Style("word-wrap") = "break-word"
                  .Style("overflow") = "auto"
                  .Style("text-align") = "left"
                  .Style.Add("z-index", "1")
                  .InnerText = NullSafeString(drElements("caption"))

                  If NullSafeInteger(drElements("BackStyle")) = 0 Then
                    .Style.Add("background-color", "Transparent")
                  Else
                    .Style.Add("background-color", objGeneral.GetHTMLColour(NullSafeInteger(drElements("BackColor"))))
                  End If

                  .Style.Add("color", objGeneral.GetHTMLColour(NullSafeInteger(drElements("ForeColor"))))

                  .Style.Add("font-family", NullSafeString(drElements("FontName")))
                  .Style.Add("font-size", NullSafeString(drElements("FontSize")) & "pt")
                  .Style.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(drElements("FontBold"))), "bold", "normal"))
                  .Style.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(drElements("FontItalic"))), "italic", "normal"))

                  iTempHeight = NullSafeInteger(drElements("Height"))
                  iTempWidth = NullSafeInteger(drElements("Width"))

                  '.Height() = Unit.Pixel(iTempHeight)
                  .Style.Add("width", CStr(iTempWidth))
                End With

              End If


            Case 3 ' Input value - character
              If NullSafeString(drElements("Name")).Length > 0 Then

                ctlForm_HtmlInputText = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), HtmlInputText)

                If NullSafeInteger(drElements("HorizontalOffsetBehaviour")) = 0 Then
                  ctlForm_HtmlInputText.Style("left") = Unit.Pixel(NullSafeInteger(drElements("HorizontalOffset"))).ToString
                Else
                  ctlForm_HtmlInputText.Style("right") = Unit.Pixel(NullSafeInteger(drElements("HorizontalOffset"))).ToString
                End If

                ctlForm_HtmlInputText.Style("resize") = "none"
                ctlForm_HtmlInputText.Style.Add("border-style", "solid")
                ctlForm_HtmlInputText.Style.Add("border-width", "1")
                ctlForm_HtmlInputText.Style.Add("border-color", objGeneral.GetHTMLColour(5730458))
                ctlForm_HtmlInputText.Style.Add("background-color", objGeneral.GetHTMLColour(NullSafeInteger(drElements("BackColor"))))
                ctlForm_HtmlInputText.Style.Add("color", objGeneral.GetHTMLColour(NullSafeInteger(drElements("ForeColor"))))
                ctlForm_HtmlInputText.Style.Add("font-family", NullSafeString(drElements("FontName")))
                ctlForm_HtmlInputText.Style.Add("font-size", NullSafeString(drElements("FontSize")) & "pt")
                ctlForm_HtmlInputText.Style.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(drElements("FontBold"))), "bold", "normal"))
                ctlForm_HtmlInputText.Style.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(drElements("FontItalic"))), "italic", "normal"))
                ctlForm_HtmlInputText.Style.Add("width", CStr(NullSafeInteger(drElements("Width"))) & "px")

              End If

          End Select

        End While

        ' Close the connection (will automatically close the reader)
        myConnection.Close()
        drElements.Close()

      Catch ex As Exception   ' conn creation 
        sMessage = "Error creating SQL connection:<BR><BR>" & _
        ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
        "Contact your system administrator."
      End Try


      If sMessage.Length > 0 Then
        Session("message") = sMessage
        Response.Redirect("Message.aspx")
      End If
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
      sImageWebPath = "~/pictures"
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

  Protected Sub btnLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLogin.Click
    submitLoginDetails()

  End Sub

  Private Sub submitLoginDetails()
    
    Dim strConn As String
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim cmdCheck As System.Data.SqlClient.SqlCommand
    Dim dr As System.Data.SqlClient.SqlDataReader

    Dim objCrypt As New Crypt

    Dim sMessage As String = ""
    Dim sEncryptedPwd As String = ""
    Dim sDecryptedPwd As String = ""
    Dim aUserDetails() As String
    Dim fAuthenticated As Boolean = False
    Dim sAuthUser As String = ""

    miSubmissionTimeoutInSeconds = mobjConfig.SubmissionTimeoutInSeconds

    If NullSafeString(Session("loginKey")) <> vbNullString Then fAuthenticated = True

    If Not fAuthenticated Then
      Dim fPersistCookie As Boolean = chkRememberPwd.Checked

      ' Get the password stored for this user, decrypt it, then compare it here.
      ' do a few quick checks before validating the user
      If txtUserName.Value.Trim.Length = 0 Then
        sMessage = "No Login entered."
      End If


      ' Lock Check.
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
        cmdCheck.CommandTimeout = miSubmissionTimeoutInSeconds

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
          ' The presence of a backslash indicates this is a windows user. 
          ' validate the user against AD
          aUserDetails = txtUserName.Value.Split("\")
          Try
            Me.AuthenticateUser(aUserDetails(0), aUserDetails(1), txtPassword.Value)
            sAuthUser = txtUserName.Value
            fAuthenticated = True
          Catch ex As Exception
            'lblError.Text = ex.Message
            'lblError.Visible = True
            fAuthenticated = False
            sMessage = ex.Message
          End Try
        Else
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
              'lblError.Text = ex.Message
              'lblError.Visible = True
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

        'cmdCheck.Parameters.Add("@psPWDParameter", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Output

        cmdCheck.Parameters.Add("@piUserGroupID", SqlDbType.Int).Direction = ParameterDirection.Output

        cmdCheck.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

        cmdCheck.ExecuteNonQuery()

        sMessage = NullSafeString(cmdCheck.Parameters("@psMessage").Value())
        ' sEncryptedPwd = NullSafeString(cmdCheck.Parameters("@psPWDParameter").Value())

        Session("UserGroupID") = NullSafeInteger(cmdCheck.Parameters("@piUserGroupID").Value())

        cmdCheck.Dispose()

      Catch ex As Exception
        sMessage = "Error :<BR><BR>" & _
        ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
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

  Private Sub AuthenticateUser(domainName As String, userName As String, password As String)
    ' Path to you LDAP directory server.
    ' Contact your network administrator to obtain a valid path.


    Dim adPath As String = "LDAP://" & System.Configuration.ConfigurationManager.AppSettings("DefaultActiveDirectoryServer")

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

  Public Shared Function isMobileBrowser() As Boolean
    'GETS THE CURRENT USER CONTEXT
    Dim context As HttpContext = HttpContext.Current

    'FIRST TRY BUILT IN ASP.NT CHECK
    If context.Request.Browser.IsMobileDevice Then
      Return True
    End If
    'THEN TRY CHECKING FOR THE HTTP_X_WAP_PROFILE HEADER
    If context.Request.ServerVariables("HTTP_X_WAP_PROFILE") IsNot Nothing Then
      Return True
    End If
    'THEN TRY CHECKING THAT HTTP_ACCEPT EXISTS AND CONTAINS WAP
    If context.Request.ServerVariables("HTTP_ACCEPT") IsNot Nothing AndAlso context.Request.ServerVariables("HTTP_ACCEPT").ToLower().Contains("wap") Then
      Return True
    End If
    'AND FINALLY CHECK THE HTTP_USER_AGENT 
    'HEADER VARIABLE FOR ANY ONE OF THE FOLLOWING
    If context.Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
      'Create a list of all mobile types
      Dim mobiles As String() = New String() {"midp", "j2me", "avant", "docomo", "novarra", "palmos", _
"palmsource", "240x320", "opwv", "chtml", "pda", "windows ce", _
"mmp/", "blackberry", "mib/", "symbian", "wireless", "nokia", _
"hand", "mobi", "phone", "cdm", "up.b", "audio", _
"SIE-", "SEC-", "samsung", "HTC", "mot-", "mitsu", _
"sagem", "sony", "alcatel", "lg", "eric", "vx", _
"philips", "mmm", "xx", "panasonic", "sharp", _
"wap", "sch", "rover", "pocket", "benq", "java", _
"pt", "pg", "vox", "amoi", "bird", "compal", _
"kg", "voda", "sany", "kdd", "dbt", "sendo", _
"sgh", "gradi", "jb", "dddi", "moto", "iphone"}

      'Loop through each item in the list created above 
      'and check if the header contains that text
      For Each s As String In mobiles
        If context.Request.ServerVariables("HTTP_USER_AGENT").ToLower().Contains(s.ToLower()) Then
          Return True
        End If
      Next

    End If

    Return False
  End Function





End Class



