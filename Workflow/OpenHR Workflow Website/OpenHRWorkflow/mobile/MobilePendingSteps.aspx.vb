﻿Imports System
Imports System.Data
Imports System.Collections.Generic
Imports Utilities

Partial Class PendingSteps
  Inherits Page

  Private _imageCount As Int16

  Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

    Dim conn As SqlClient.SqlConnection
    Dim cmdSteps As SqlClient.SqlCommand
    Dim rstSteps As SqlClient.SqlDataReader
    Dim ctlFormHtmlGenericControl As HtmlGenericControl
    Dim ctlFormHtmlInputText As HtmlInputText
    Dim ctlFormImage As Image
    Dim ctlFormImageButton As ImageButton   ' Button
    Dim objGeneral As New General
    Dim sMessage As String = ""
    Dim drLayouts As SqlClient.SqlDataReader
    Dim drElements As SqlClient.SqlDataReader
    Dim sImageFileName As String = ""
    Dim ctlFormTable As Table
    Dim ctlFormRow As TableRow
    Dim ctlFormCell As TableCell
    Dim ctlFormLabel As Label
    Dim sql As String
    Dim command As SqlClient.SqlCommand
    Dim reader As IDataReader

    Dim strWFStepText As String

    _ImageCount = 0

    ' Establish Connection
    Dim myConnection As New SqlClient.SqlConnection(Configuration.ConnectionString)
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

    Dim todoTitleStyles = New Dictionary(Of String, String)
    todoTitleStyles.Add("font-family", NullSafeString(drLayouts("TodoTitleFontName")))
    todoTitleStyles.Add("font-size", NullSafeString(drLayouts("TodoTitleFontSize")) & "pt")
    todoTitleStyles.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(drLayouts("TodoTitleFontBold"))), "bold", "normal"))
    todoTitleStyles.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(drLayouts("TodoTitleFontItalic"))), "italic", "normal"))

    Dim todoDescStyles = New Dictionary(Of String, String)
    todoDescStyles.Add("font-family", NullSafeString(drLayouts("TodoDescFontName")))
    todoDescStyles.Add("font-size", NullSafeString(drLayouts("TodoDescFontSize")) & "pt")
    todoDescStyles.Add("font-weight", If(NullSafeBoolean(NullSafeBoolean(drLayouts("TodoDescFontBold"))), "bold", "normal"))
    todoDescStyles.Add("font-style", If(NullSafeBoolean(NullSafeBoolean(drLayouts("TodoDescFontItalic"))), "italic", "normal"))

    ' Close the connection (will automatically close the reader)
    myConnection.Close()
    drLayouts.Close()

    ' ======================== NOW FOR THE INDIVIDUAL ELEMENTS  ====================================

    ' Establish Connection
    myConnection = New SqlClient.SqlConnection(Configuration.ConnectionString)
    myConnection.Open()

    ' Create command
    myCommand = New SqlClient.SqlCommand("select * from tbsys_mobileformelements where form = 5", myConnection)

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

    ' Close the connection (will automatically close the reader)
    myConnection.Close()
    drElements.Close()

    ' -----------------------------------------------------------------------------------------------------------------------------------

    conn = New SqlClient.SqlConnection(Configuration.ConnectionString)
    conn.Open()

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
                                " WHERE [categoryID] = (SELECT [categoryID] FROM [ASRSysPermissionCategories] WHERE [categoryKey] = 'WORKFLOW')) " & _
                           " AND [groupName] = (SELECT [Name] FROM [ASRSysGroups] WHERE [ID] = " & groupId.ToString & ")"
      Try
        command = New SqlClient.SqlCommand(sql, conn)
        reader = command.ExecuteReader()

        While reader.Read()
          Select Case CStr(reader("itemKey"))
            Case "RUN"
              fUserHasRunPermission = (reader("permitted") = True)

          End Select
        End While

        reader.Close()
      Catch ex As Exception

      End Try

    End If

    If fUserHasRunPermission Then

      ' Get the pending steps.
      ' Open a connection to the database.
      cmdSteps = New SqlClient.SqlCommand
      cmdSteps.CommandText = "spASRSysMobileCheckPendingWorkflowSteps"
      cmdSteps.Connection = conn
      cmdSteps.CommandType = CommandType.StoredProcedure

      cmdSteps.Parameters.Add("@psKeyParameter", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmdSteps.Parameters("@psKeyParameter").Value = User.Identity.Name

      rstSteps = cmdSteps.ExecuteReader

      ' Create the holding table
      ctlFormTable = New Table

      Dim iLoop As Integer
      While (rstSteps.Read)
        ' Create a row to contain this pending step...
        ctlFormRow = New TableRow
        ctlFormRow.Attributes.Add("onclick", "window.open('" & rstSteps("url").ToString & "');")

        ' Create a cell to contain the workflow icon
        ctlFormCell = New TableCell  ' Image cell
        ctlFormImage = New Image
        If NullSafeInteger(rstSteps("pictureID")) = 0 Then
          sImageFileName = "~/Images/Connected48.png"
        Else
          sImageFileName = LoadPicture(NullSafeInteger(rstSteps("pictureID")), sMessage)
        End If
        ctlFormImage.ImageUrl = sImageFileName
        ctlFormImage.Height() = Unit.Pixel(57)
        ctlFormImage.Width() = Unit.Pixel(57)

        ' add ImageButton to cell
        ctlFormCell.Controls.Add(ctlFormImage)

        ' Add cell to row
        ctlFormRow.Cells.Add(ctlFormCell)

        ' Create a cell to contain the workflow name and description
        ctlFormCell = New TableCell
        ctlFormLabel = New Label ' Workflow name text
        ctlFormLabel.Font.Underline = True
        ctlFormLabel.Text = CStr(rstSteps("name"))
        For Each item In todoTitleStyles
          ctlFormLabel.Style.Add(item.Key, item.Value)
        Next
        ctlFormCell.Controls.Add(ctlFormLabel)

        ' Line Break
        ctlFormCell.Controls.Add(New LiteralControl("<br>"))

        ctlFormLabel = New Label ' Workflow step description text

        If Left(rstSteps("description"), Len(rstSteps("name")) + 2) = (Trim(rstSteps("name")) & " -") Then
          strWFStepText = rstSteps("description").ToString.Remove(0, (rstSteps("name").ToString.Length) + 2)
        Else
          strWFStepText = rstSteps("description").ToString
        End If
        ctlFormLabel.Text = strWFStepText
        For Each item In todoDescStyles
          ctlFormLabel.Style.Add(item.Key, item.Value)
        Next
        ctlFormCell.Controls.Add(ctlFormLabel)

        ' Add cell to row, and row to table.
        ctlFormRow.Cells.Add(ctlFormCell)
        ctlFormTable.Rows.Add(ctlFormRow)

        iLoop += 1
      End While

      pnlWFList.Controls.Add(ctlFormTable)

      hdnStepCount.Value = CStr(iLoop)

      rstSteps.Close()
      cmdSteps.Dispose()
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
      _ImageCount = CShort(_ImageCount + 1)

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
           "_" & _ImageCount.ToString & _
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

  Protected Sub BtnRefreshClick(sender As Object, e As ImageClickEventArgs) Handles btnRefresh.Click
    Response.Redirect("~/Mobile/MobilePendingSteps.aspx")
  End Sub

  Protected Sub BtnCancelClick(sender As Object, e As ImageClickEventArgs) Handles btnCancel.Click
    Response.Redirect("~/Mobile/MobileHome.aspx")
  End Sub
End Class
