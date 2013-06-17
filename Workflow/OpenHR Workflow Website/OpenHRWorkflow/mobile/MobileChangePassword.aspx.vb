Imports System
Imports System.Data
Imports Utilities

Partial Class ChangePassword
  Inherits Page

  Private _imageCount As Int16

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    Dim ctlFormHtmlGenericControl As HtmlGenericControl
    Dim ctlFormImageButton As ImageButton
    Dim ctlFormHtmlInputText As HtmlInputText
    Dim objGeneral As New General
    Dim sMessage As String = ""
    Dim drLayouts As SqlClient.SqlDataReader
    Dim drElements As SqlClient.SqlDataReader
    Dim sImageFileName As String = ""

    _imageCount = 0

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

    End If

    ' Close the connection (will automatically close the reader)
    myConnection.Close()
    drLayouts.Close()

    ' ======================== NOW FOR THE INDIVIDUAL ELEMENTS  ====================================

    ' Establish Connection
    myConnection = New SqlClient.SqlConnection(Configuration.ConnectionString)
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

    ' Close the connection (will automatically close the reader)
    myConnection.Close()
    drElements.Close()

  End Sub

  Private Function LoadPicture(ByVal piPictureID As Int32, ByRef psErrorMessage As String) As String

    Dim conn As SqlClient.SqlConnection
    Dim cmdSelect As SqlClient.SqlCommand
    Dim dr As SqlClient.SqlDataReader
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

  Protected Sub BtnSubmitClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click

    Dim conn As SqlClient.SqlConnection
    Dim cmdCheckUserSessions As SqlClient.SqlCommand
    Dim cmdChangePassword As SqlClient.SqlCommand
    Dim cmdPasswordOk As SqlClient.SqlCommand
    Dim sHeader As String = ""
    Dim sMessage As String = ""
    Dim sRedirectTo As String = ""

    Try
      'TODO how does nick want change password thing to work???
      ' Like this :P

      If Not User.Identity.IsAuthenticated Then sMessage = "Unable to identify you as a user"

      If sMessage.Length = 0 Then
        conn = New SqlClient.SqlConnection(Configuration.ConnectionString)
        conn.Open()

        ' Force password change only if there are no other users logged in with the same name.
        cmdCheckUserSessions = New SqlClient.SqlCommand
        cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
        cmdCheckUserSessions.Connection = conn
        cmdCheckUserSessions.CommandType = CommandType.StoredProcedure

        cmdCheckUserSessions.Parameters.Add("@iLoginCount", SqlDbType.Int).Direction = ParameterDirection.Output

        cmdCheckUserSessions.Parameters.Add("@psLoginName", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
        cmdCheckUserSessions.Parameters("@psLoginName").Value = User.Identity.Name.ToString()

        cmdCheckUserSessions.ExecuteNonQuery()

        Dim iUserSessionCount As Integer = CInt(cmdCheckUserSessions.Parameters("@iLoginCount").Value)

        cmdCheckUserSessions.Dispose()


        ' is OK?
        If iUserSessionCount < 2 Then
          ' Read the Password details from the Password form.
          Dim sCurrentPassword As String = txtCurrPassword.Value
          Dim sNewPassword As String = txtNewPassword.Value

          ' Attempt to change the password on the SQL Server.
          cmdChangePassword = New SqlClient.SqlCommand
          cmdChangePassword.CommandText = "sp_password"
          cmdChangePassword.Connection = conn
          cmdChangePassword.CommandType = CommandType.StoredProcedure

          cmdChangePassword.Parameters.Add("@old", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          If Len(sCurrentPassword) > 0 Then
            cmdChangePassword.Parameters("@old").Value = sCurrentPassword
          Else
            cmdChangePassword.Parameters("@old").Value = vbNullString
          End If

          cmdChangePassword.Parameters.Add("@new", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          If Len(sNewPassword) > 0 Then
            cmdChangePassword.Parameters("@new").Value = sNewPassword
          Else
            cmdChangePassword.Parameters("@new").Value = vbNullString
          End If

          cmdChangePassword.Parameters.Add("@loginame", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          cmdChangePassword.Parameters("@loginame").Value = User.Identity.Name.ToString()

          cmdChangePassword.ExecuteNonQuery()

          cmdChangePassword.Dispose()
        Else
          sMessage = "You could not change your password. The account is currently being used by "
          If iUserSessionCount > 2 Then
            sMessage &= iUserSessionCount.ToString & " users"
          Else
            sMessage &= " another user"
          End If
          sMessage &= " in the system."
        End If

        If sMessage.Length = 0 Then
          ' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
          cmdPasswordOK = New SqlClient.SqlCommand
          cmdPasswordOK.CommandText = "spASRSysMobilePasswordOK"
          cmdPasswordOK.Connection = conn
          cmdPasswordOK.CommandType = CommandType.StoredProcedure

          cmdPasswordOK.Parameters.Add("@sCurrentUser", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
          cmdPasswordOK.Parameters("@sCurrentUser").Value = User.Identity.Name.ToString()

          cmdPasswordOK.ExecuteNonQuery()

          cmdPasswordOK.Dispose()

          ' Tell the user that the password was changed okay.
          sMessage = "Password changed successfully."
        End If
      End If

      sHeader = "Change Password Submitted"
      
    Catch ex As Exception
      sHeader = "Change Password Failed"
      sMessage = "Error :" & vbCrLf & vbCrLf & ex.Message.ToString & vbCrLf & vbCrLf & "Contact your system administrator."
    End Try


    If sMessage.Length = 0 Then
      sMessage = "Password changed successfully."
      sRedirectTo = "/Mobile/MobileHome.aspx"
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
    Response.Redirect("~/Mobile/MobileHome.aspx")
  End Sub
End Class
