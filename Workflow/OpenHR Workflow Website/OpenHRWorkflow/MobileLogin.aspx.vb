﻿Imports System.Data
Imports Utilities

Partial Class MobileLogin
  Inherits System.Web.UI.Page

  Private miImageCount As Int16
  Private mobjConfig As New Config

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ctlForm_HTMLGenericControl As HtmlGenericControl
    Dim ctlForm_HtmlInputText As HtmlInputText
    Dim ctlForm_Image As Image
    Dim ctlForm_ImageButton As ImageButton   ' Button
    Dim strConn As String
    Dim objGeneral As New General
    Dim lngPanel_1_Height As Long = 57
    Dim lngPanel_2_Height As Long = 57
    Dim sBackgroundImage As String
    Dim sBackgroundRepeat As String
    Dim sBackgroundPosition As String
    Dim iBackgroundColour As Integer
    Dim sBackgroundColourHex As String
    Dim iBackgroundImagePosition As Integer
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

      End Try

      ' Establish Connection
      strConn = "Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
        ";Initial Catalog=" & Session("Database") & _
        ";Integrated Security=false;User ID=" & Session("Login") & _
        ";Password=" & Session("Password") & _
        ";Pooling=false"
      'strConn = "Application Name=OpenHR Workflow;Data Source=.\sqlexpress;Initial Catalog=hrprostd43;Integrated Security=false;User ID=sa;Password=asr;Pooling=false"

      Dim myConnection As New SqlClient.SqlConnection(strConn)
      myConnection.Open()

      ' Create command
      Dim myCommand As New SqlClient.SqlCommand("select * from tbsys_mobileformlayout where FormTypeID = 1", myConnection)

      ' Create a DataReader to ferry information back from the database
      drLayouts = myCommand.ExecuteReader()

      ctlForm_HTMLGenericControl = New HtmlGenericControl

      'Iterate through the results
      While drLayouts.Read()
        sBackgroundImage = ""
        sBackgroundRepeat = ""
        sBackgroundPosition = ""

        '... Work with the current record ...
        Select Case CInt(drLayouts("PanelID"))
          Case 1  ' ================= Banner Div ================
            'Dim strHeader As String = "pnlHeader"
            'ctlForm_HTMLGenericControl = TryCast(pnlContainer.FindControl(strHeader), HtmlGenericControl)

            ctlForm_HTMLGenericControl = TryCast(pnlHeader, HtmlGenericControl)

          Case 2    ' ================= Main Body Div ================
            ctlForm_HTMLGenericControl = TryCast(ScrollerFrame, HtmlGenericControl)

          Case 3   ' ================= Footer Div ================
            ctlForm_HTMLGenericControl = TryCast(pnlFooter, HtmlGenericControl)
        End Select

        ' Background Image
        If CInt(drLayouts("PictureID")) > 0 Then
          sBackgroundImage = LoadPicture(CInt(drLayouts("PictureID")), sMessage)
          ctlForm_HTMLGenericControl.Style("Background-image") = sBackgroundImage
          sBackgroundImage = "url('" & sBackgroundImage & "')"
          iBackgroundImagePosition = CInt(drLayouts("PictureLocation"))
          sBackgroundRepeat = objGeneral.BackgroundRepeat(CShort(iBackgroundImagePosition))
          sBackgroundPosition = objGeneral.BackgroundPosition(CShort(iBackgroundImagePosition))
          ctlForm_HTMLGenericControl.Style("background-repeat") = sBackgroundRepeat
          ctlForm_HTMLGenericControl.Style("background-position") = sBackgroundPosition
        End If

        ' background colour
        sBackgroundColourHex = ""
        If Not IsDBNull(drLayouts("BackColor")) Then
          iBackgroundColour = CInt(drLayouts("BackColor"))
          sBackgroundColourHex = objGeneral.GetHTMLColour(iBackgroundColour).ToString()
          ctlForm_HTMLGenericControl.Style("Background-color") = objGeneral.GetHTMLColour(NullSafeInteger(iBackgroundColour))
        End If

      End While

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
      myCommand = New SqlClient.SqlCommand("select * from tbsys_mobileformelements where layoutid in (select layoutid from tbsys_mobileformlayout where formtypeid = 1)", myConnection)

      ' Create a DataReader to ferry information back from the database
      drElements = myCommand.ExecuteReader()

      'Iterate through the results
      While drElements.Read()
        Select Case CInt(drElements("ElementType"))

          Case 0 ' Button

            If CInt(drElements("ButtonStyle")) = 1 Then   ' Footer Icon

              If NullSafeString(drElements("HTMLElementID")).Length > 0 Then
                ctlForm_ImageButton = TryCast(pnlContainer.FindControl(NullSafeString(drElements("HTMLElementID"))), ImageButton)

                'ctlForm_ImageButton = New ImageButton
                With ctlForm_ImageButton
                  .Width = Unit.Pixel(40)
                  .Height = Unit.Pixel(40)
                  sImageFileName = LoadPicture(NullSafeInteger(drElements("pictureID")), sMessage)
                  .ImageUrl = sImageFileName
                  .Font.Name = NullSafeString(drElements("FontName"))
                  .Font.Size = FontUnit.Parse(NullSafeString(drElements("FontSize")))
                  .Font.Bold = NullSafeBoolean(NullSafeBoolean(drElements("FontBold")))
                  .Font.Italic = NullSafeBoolean(NullSafeBoolean(drElements("FontItalic")))
                  .Font.Strikeout = NullSafeBoolean(NullSafeBoolean(drElements("FontStrikeThru")))
                End With

                ' Footer text
                If NullSafeString(drElements("Caption")).Length > 0 Then
                  ctlForm_HTMLGenericControl = TryCast(pnlContainer.FindControl(NullSafeString(drElements("HTMLElementID")) & "_label"), HtmlGenericControl)
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
            End If


          Case 2 ' Label
            If NullSafeString(drElements("HTMLElementID")).Length > 0 Then
              ctlForm_HTMLGenericControl = TryCast(pnlContainer.FindControl(NullSafeString(drElements("HTMLElementID"))), HtmlGenericControl)  'New Label
              With ctlForm_HTMLGenericControl
                '.Style("position") = "absolute"

                '' Vertical Offset
                'If Not IsDBNull(drElements("TopCoord")) Or NullSafeInteger(drElements("VerticalOffsetBehaviour")) = 0 Then
                '  .Style("top") = Unit.Pixel(NullSafeInteger(drElements("VerticalOffset"))).ToString
                'Else
                '  .Style("bottom") = Unit.Pixel(NullSafeInteger(drElements("VerticalOffset"))).ToString
                'End If

                ' horizontal position
                If Not IsDBNull(drElements("LeftCoord")) Or NullSafeInteger(drElements("HorizontalOffsetBehaviour")) = 0 Then
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
                .Style.Add("font-weight", IIf(NullSafeBoolean(NullSafeBoolean(drElements("FontBold"))), "bold", "normal"))
                .Style.Add("font-style", IIf(NullSafeBoolean(NullSafeBoolean(drElements("FontItalic"))), "italic", "normal"))
                '.Font.Strikeout = NullSafeBoolean(NullSafeBoolean(drElements("FontStrikeThru")))

                iTempHeight = NullSafeInteger(drElements("Height"))
                iTempWidth = NullSafeInteger(drElements("Width"))

                '.Height() = Unit.Pixel(iTempHeight)
                .Style.Add("width", CStr(iTempWidth))
              End With

            End If


          Case 3 ' Input value - character
            If NullSafeString(drElements("HTMLElementID")).Length > 0 Then

              ctlForm_HtmlInputText = TryCast(pnlContainer.FindControl(NullSafeString(drElements("HTMLElementID"))), HtmlInputText)

              If Not IsDBNull(drElements("LeftCoord")) Or NullSafeInteger(drElements("HorizontalOffsetBehaviour")) = 0 Then
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
              ctlForm_HtmlInputText.Style.Add("font-size", NullSafeString(drElements("FontSize")))
              ctlForm_HtmlInputText.Style.Add("font-weight", IIf(NullSafeBoolean(NullSafeBoolean(drElements("FontBold"))), "bold", "normal"))
              ctlForm_HtmlInputText.Style.Add("font-style", IIf(NullSafeBoolean(NullSafeBoolean(drElements("FontItalic"))), "italic", "normal"))
              ctlForm_HtmlInputText.Style.Add("width", CStr(NullSafeInteger(drElements("Width"))) & "px")

            End If


          Case 10 ' Image
            ctlForm_Image = New System.Web.UI.WebControls.Image

            With ctlForm_Image
              .Style("position") = "absolute"

              ' Vertical Offset
              If Not IsDBNull(drElements("TopCoord")) Or NullSafeInteger(drElements("VerticalOffsetBehaviour")) = 0 Then
                .Style("top") = Unit.Pixel(NullSafeInteger(drElements("VerticalOffset"))).ToString
              Else
                .Style("bottom") = Unit.Pixel(NullSafeInteger(drElements("VerticalOffset"))).ToString
              End If

              ' horizontal position
              If Not IsDBNull(drElements("LeftCoord")) Or NullSafeInteger(drElements("HorizontalOffsetBehaviour")) = 0 Then
                .Style("left") = Unit.Pixel(NullSafeInteger(drElements("HorizontalOffset"))).ToString
              Else
                .Style("right") = Unit.Pixel(NullSafeInteger(drElements("HorizontalOffset"))).ToString
              End If

              If NullSafeInteger(drElements("BackStyle")) = 0 Then
                .BackColor = System.Drawing.Color.Transparent
              Else
                .BackColor = objGeneral.GetColour(NullSafeInteger(drElements("BackColor")))
              End If

              sImageFileName = LoadPicture(NullSafeInteger(drElements("pictureID")), sMessage)
              If sMessage.Length > 0 Then
                Exit While
              End If
              .ImageUrl = sImageFileName

              iTempHeight = NullSafeInteger(drElements("Height"))
              iTempWidth = NullSafeInteger(drElements("Width"))

              .Height() = Unit.Pixel(iTempHeight)
              .Width() = Unit.Pixel(iTempWidth)
              .Style.Add("z-index", "1")
            End With

            pnlContainer.Controls.Add(ctlForm_Image)
        End Select

      End While

      ' Close the connection (will automatically close the reader)
      myConnection.Close()
      drElements.Close()
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
      sImageFilePath = Server.MapPath("pictures")

      strConn = "Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
        ";Initial Catalog=" & Session("Database") & _
        ";Integrated Security=false;User ID=" & Session("Login") & _
        ";Password=" & Session("Password") & _
        ";Pooling=false"
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
        LoadPicture = "pictures/" & sImageFileName

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
    Dim strConn As String
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim cmdCheck As System.Data.SqlClient.SqlCommand
    Dim objCrypt As New Crypt

    Dim sMessage As String = ""
    Dim sEncryptedPwd As String = ""
    Dim sDecryptedPwd As String = ""

    ' Get the password stored for this user, decrypt it, then compare it here.

    Try
      If txtUserName.Value.Trim.Length = 0 Then
        sMessage = "No Login entered."
      Else
        Session("LoginKey") = txtUserName.Value        
        Session("RememberPWD") = chkRememberPwd.Value

        strConn = "Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
        ";Initial Catalog=" & Session("Database") & _
        ";Integrated Security=false;User ID=" & Session("Login") & _
        ";Password=" & Session("Password") & _
        ";Pooling=false"
        conn = New SqlClient.SqlConnection(strConn)
        conn.Open()

        cmdCheck = New SqlClient.SqlCommand
        cmdCheck.CommandText = "spASRSysMobileCheckLogin"
        cmdCheck.Connection = conn
        cmdCheck.CommandType = CommandType.StoredProcedure

        cmdCheck.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        cmdCheck.Parameters("@psKeyParameter").Value = Session("LoginKey")

        cmdCheck.Parameters.Add("@psPWDParameter", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Output

        cmdCheck.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

        cmdCheck.ExecuteNonQuery()

        sMessage = NullSafeString(cmdCheck.Parameters("@psMessage").Value())
        sEncryptedPwd = NullSafeString(cmdCheck.Parameters("@psPWDParameter").Value())

        cmdCheck.Dispose()
      End If

    Catch ex As Exception
      sMessage = "Error :<BR><BR>" & _
      ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
      "Contact your system administrator."
    End Try

    If sMessage.Length > 0 Then
      Session("message") = sMessage
      Session("nextPage") = "MobileLogin"
      lblMsgBox.InnerText = sMessage
      pnlGreyOut.Style.Add("visibility", "visible")
      pnlMsgBox.Style.Add("visibility", "visible")
    Else
      ' Validate the password
      sDecryptedPwd = objCrypt.DecryptString(sEncryptedPwd, "jmltn", False)

      If Trim(txtPassword.Value) = Trim(sDecryptedPwd) Then
        Session("LoginPWD") = sDecryptedPwd
        Response.Redirect("MobileHome.aspx")
      Else
        sMessage = "Invalid password."
        lblMsgBox.InnerText = sMessage
        pnlGreyOut.Style.Add("visibility", "visible")
        pnlMsgBox.Style.Add("visibility", "visible")
      End If

    End If

  End Sub


  Protected Sub btnRegister_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnRegister.Click
    Response.Redirect("MobileRegistration.aspx")
  End Sub

  Protected Sub btnForgotPwd_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnForgotPwd.Click
    Response.Redirect("MobileForgottenLogin.aspx")
  End Sub
End Class



