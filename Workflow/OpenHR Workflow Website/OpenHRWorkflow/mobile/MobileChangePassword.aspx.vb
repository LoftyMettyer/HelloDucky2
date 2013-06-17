Imports System.Data
Imports Utilities

Partial Class ChangePassword

  Inherits System.Web.UI.Page
  Private mobjConfig As New Config
  Private miImageCount As Int16

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

    Try
      mobjConfig.Mob_Initialise()
      Session("Server") = mobjConfig.Server
      Session("Database") = mobjConfig.Database
      Session("Login") = mobjConfig.Login
      Session("Password") = mobjConfig.Password
      Session("WorkflowURL") = mobjConfig.WorkflowURL

    Catch ex As Exception
      sMessage = "Unable to initialise screen" & vbCrLf & ex.Message
    End Try

    If sMessage.Length = 0 Then

      ' Establish Connection
      strConn = "Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
              ";Initial Catalog=" & Session("Database") & _
              ";Integrated Security=false;User ID=" & Session("Login") & _
              ";Password=" & Session("Password") & _
              ";Pooling=false"
      'strConn = "Application Name=OpenHR Mobile;Data Source=.\sqlexpress;Initial Catalog=hrprostd43;Integrated Security=false;User ID=sa;Password=asr;Pooling=false"

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

      myConnection = New SqlClient.SqlConnection(strConn)
      myConnection.Open()

      ' Create command
      myCommand = New SqlClient.SqlCommand("select * from tbsys_mobileformelements where form = 4", myConnection)

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
                sImageFileName = LoadPicture(NullSafeInteger(drElements("pictureID")), sMessage)
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

    End If


    If sMessage.Length > 0 Then
      ' Display message box.
      lblMsgBox.InnerText = sMessage
      pnlGreyOut.Style.Add("visibility", "visible")
      pnlMsgBox.Style.Add("visibility", "visible")
      Session("nextPage") = "MobileHome"
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


  Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
    Dim strConn As String
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim sMessage As String = ""
    Dim objCrypt As New Crypt
    Dim strEncryptedPwd As String
    Dim cmdChgPwd As System.Data.SqlClient.SqlCommand

    Try ' conn creation 

      ' Encrypt the password so that it can be transmitted in clear text
      strEncryptedPwd = objCrypt.EncryptString(txtNewPassword.Value, "jmltn", False)

      strConn = "Application Name=OpenHR Mobile;Data Source=" & Session("Server") & _
       ";Initial Catalog=" & Session("Database") & _
       ";Integrated Security=false;User ID=" & Session("Login") & _
       ";Password=" & Session("Password") & _
       ";Pooling=false"
      conn = New SqlClient.SqlConnection(strConn)
      conn.Open()

      cmdChgPwd = New SqlClient.SqlCommand
      cmdChgPwd.CommandText = "spASRSysMobileChangePassword"
      cmdChgPwd.Connection = conn
      cmdChgPwd.CommandType = CommandType.StoredProcedure

      cmdChgPwd.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmdChgPwd.Parameters("@psKeyParameter").Value = Session("LoginKey")

      cmdChgPwd.Parameters.Add("@psPWDParameterNew", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
      cmdChgPwd.Parameters("@psPWDParameterNew").Value = strEncryptedPwd

      cmdChgPwd.ExecuteNonQuery()
      cmdChgPwd.Dispose()

    Catch ex As Exception

      sMessage = "Error :<BR><BR>" & _
     ex.Message.Replace(vbCrLf, "<BR>") & "<BR><BR>" & _
     "Contact your system administrator."

    End Try

    If sMessage.Length = 0 Then
      Session("LoginPWD") = txtNewPassword.Value
      sMessage = "Password changed successfully."
    End If

    ' Display message box.
    lblMsgBox.InnerText = sMessage
    pnlGreyOut.Style.Add("visibility", "visible")
    pnlMsgBox.Style.Add("visibility", "visible")
    Session("nextPage") = "MobileHome"
  End Sub

  Protected Sub btnCancel_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnCancel.Click
    Response.Redirect("MobileHome.aspx")
  End Sub
End Class
