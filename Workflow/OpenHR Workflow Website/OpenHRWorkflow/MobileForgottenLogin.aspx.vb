Imports System.Data
Imports Utilities

Partial Class ForgottenLogin
  Inherits Page

  Private _imageCount As Int16

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init

    Dim ctlFormHtmlGenericControl As HtmlGenericControl
    Dim ctlFormHtmlInputText As HtmlInputText
    Dim ctlFormImageButton As ImageButton
    Dim objGeneral As New General
    Dim sMessage As String = ""
    Dim drLayouts As SqlClient.SqlDataReader
    Dim drElements As SqlClient.SqlDataReader
    Dim sImageFileName As String = ""

    _ImageCount = 0

    ' Establish Connection
    Dim myConnection As New SqlClient.SqlConnection(Configuration.ConnectionString)
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
          control.Style("Background-image") = LoadPicture(CInt(drLayouts(prefix & "PictureID")), sMessage, True)
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

            .BackColor = Drawing.Color.Transparent
            .ImageUrl = LoadPicture(NullSafeInteger(drLayouts("HeaderLogoID")), sMessage, False)
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
    myConnection = New SqlClient.SqlConnection(Configuration.ConnectionString)
    myConnection.Open()

    ' Create command
    myCommand = New SqlClient.SqlCommand("select * from tbsys_mobileformelements where form = 6", myConnection)

    ' Create a DataReader to ferry information back from the database
    drElements = myCommand.ExecuteReader()

    'Iterate through the results
    While drElements.Read()
      Select Case CInt(drElements("Type"))

        Case 0 ' Button

          If NullSafeString(drElements("Name")).Length > 0 Then
            ctlFormImageButton = TryCast(pnlContainer.FindControl(NullSafeString(drElements("Name"))), ImageButton)

            With ctlFormImageButton
              sImageFileName = LoadPicture(NullSafeInteger(drElements("pictureID")), sMessage, False)
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

    ' Close the connection (will automatically close the reader)
    myConnection.Close()
    drElements.Close()

    SetCustomControlSettings()

  End Sub

  Private Sub SetCustomControlSettings()
    ' Set the e-mail input field to type=email (html5 only) ASP.NET requires this to be added thus:
    txtEmail.Attributes.Add("type", "email")
  End Sub

  Private Function LoadPicture(ByVal piPictureID As Int32, ByRef psErrorMessage As String, ByVal pfServerSide As Boolean) As String

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
    Dim iBufferSize As Integer = 100
    Dim outByte(iBufferSize - 1) As Byte
    Dim retVal As Long
    Dim startIndex As Long
    Dim sExtension As String = ""
    Dim iIndex As Integer
    Dim sName As String

    Try
      _ImageCount = CShort(_ImageCount + 1)

      psErrorMessage = ""
      LoadPicture = ""
      sImageFileName = ""
      sImageWebPath = "pictures"
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
           "_" & _ImageCount.ToString & _
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

  Protected Sub BtnSubmitClick(sender As Object, e As ImageClickEventArgs) Handles btnSubmit.Click

    Dim conn As SqlClient.SqlConnection
    Dim cmdForgotLogin As SqlClient.SqlCommand
    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""
    Dim lngUserID As Long

    Try
      conn = New SqlClient.SqlConnection(Configuration.ConnectionString)
      conn.Open()

      ' Done in three parts. First get the ID for this e-mail (SQL). Second retrieve and decrypt password (VB), third send a reminder e-mail (SQL).
      ' Scratch that! First get the username from the db for this email address, then send the e-mail.

      cmdForgotLogin = New SqlClient.SqlCommand
      cmdForgotLogin.CommandText = "spASRSysMobileGetUserIDFromEmail"
      cmdForgotLogin.Connection = conn
      cmdForgotLogin.CommandType = CommandType.StoredProcedure

      cmdForgotLogin.Parameters.Add("@psEmail", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmdForgotLogin.Parameters("@psEmail").Value = txtEmail.Value

      cmdForgotLogin.Parameters.Add("@piUserID", SqlDbType.Int).Direction = ParameterDirection.Output

      cmdForgotLogin.ExecuteNonQuery()

      lngUserID = CLng(NullSafeInteger(cmdForgotLogin.Parameters("@piUserID").Value()))

      cmdForgotLogin.Dispose()

      If lngUserID = 0 Then sMessage = "No records exist with the given email address."

      If sMessage.Length = 0 Then
        ' ------------- Part two, send it all to sql to validate and email out -----------------
        cmdForgotLogin = New SqlClient.SqlCommand
        cmdForgotLogin.CommandText = "spASRSysMobileForgotLogin"
        cmdForgotLogin.Connection = conn
        cmdForgotLogin.CommandType = CommandType.StoredProcedure

        cmdForgotLogin.Parameters.Add("@psEmailAddress", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        cmdForgotLogin.Parameters("@psEmailAddress").Value = txtEmail.Value

        cmdForgotLogin.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

        cmdForgotLogin.ExecuteNonQuery()

        sMessage = CStr(cmdForgotLogin.Parameters("@psMessage").Value())

        cmdForgotLogin.Dispose()
      End If

    Catch ex As Exception
      sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Contact your system administrator."
    End Try

    If sMessage.Length > 0 Then
      sHeader = "Request Failed"
    Else
      sHeader = "Request Submitted"
      sMessage = "An email has been sent to the entered address with your login details."
      sRedirectTo = "MobileLogin.aspx"
    End If

    ShowMessage(sHeader, sMessage, sRedirectTo)

  End Sub

  Private Sub ShowMessage(headerText As String, messageText As String, redirectTo As String)

    lblMsgHeader.InnerText = headerText
    lblMsgBox.InnerText = messageText
    hdnRedirectTo.Value = redirectTo
    pnlGreyOut.Style.Add("visibility", "visible")
    pnlMsgBox.Style.Add("visibility", "visible")

  End Sub

  Protected Sub BtnCancelClick(sender As Object, e As ImageClickEventArgs) Handles btnCancel.Click
    Response.Redirect("~/MobileLogin.aspx")
  End Sub

End Class
