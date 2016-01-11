Imports System.IO
Imports System.Data
Imports Utilities

Partial Class FileUpload
   Inherits Page

   'TODO PG NOW expose Config and remove other functions
   Private ReadOnly mobjConfig As New Config
   Private miSubmissionTimeoutInSeconds As Int32

   Public Function ColourThemeHex() As String
      ColourThemeHex = mobjConfig.ColourThemeHex
   End Function

   Public Function ColourThemeFolder() As String
      ColourThemeFolder = mobjConfig.ColourThemeFolder
   End Function

   Private Function SaveImage(ByVal abtImage As Byte(), ByVal psContentType As String, ByVal psFileName As String, ByVal pfClear As Boolean) As Integer
      Dim iRowsAffected As Integer
      Dim strConn As String
      Dim conn As SqlClient.SqlConnection
      Dim cmdSave As SqlClient.SqlCommand

      If abtImage.GetUpperBound(0) = 0 Then
         psContentType = ""
         psFileName = ""
      End If

      strConn = CStr("Application Name=OpenHR Workflow;Data Source=" & Session("Server") & ";Initial Catalog=" & Session("Database") & ";Integrated Security=false;User ID=" & Session("User") & ";Password=" & Session("Pwd") & ";Pooling=false")
      conn = New SqlClient.SqlConnection(strConn)
      conn.Open()

      cmdSave = New SqlClient.SqlCommand("spASRWorkflowFileUpload", conn)
      cmdSave.CommandType = CommandType.StoredProcedure
      cmdSave.CommandTimeout = miSubmissionTimeoutInSeconds

      cmdSave.Parameters.AddWithValue("@piElementItemID", ViewState("ElementItemID"))
      cmdSave.Parameters.AddWithValue("@piInstanceID", Session("InstanceID"))
      cmdSave.Parameters.AddWithValue("@pimgFile", abtImage)
      cmdSave.Parameters.AddWithValue("@psContentType", psContentType)
      cmdSave.Parameters.AddWithValue("@psFileName", psFileName)
      cmdSave.Parameters.AddWithValue("@pfClear", pfClear)

      iRowsAffected = cmdSave.ExecuteNonQuery()

      cmdSave.Dispose()

      conn.Close()
      conn.Dispose()

      SaveImage = iRowsAffected
   End Function

   Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

      Dim sTemp As String
      Dim iTemp As Integer
      Dim sQueryString As String
      Dim iElementItemID As Integer
      Dim strConn As String
      Dim drDetails As SqlClient.SqlDataReader
      Dim cmdDetails As SqlClient.SqlCommand
      Dim conn As SqlClient.SqlConnection
      Dim iMaxFileSize As Integer
      Dim sFileExtensions As String
      Dim sFileName As String
      Dim sFileNameWithoutPath As String
      Dim objCrypt As New Crypt
      Dim fAlreadyUploaded As Boolean

      Try
         mobjConfig.Initialise(Server.MapPath("themes/ThemeHex.xml"))
         miSubmissionTimeoutInSeconds = mobjConfig.SubmissionTimeoutInSeconds

         Response.CacheControl = "no-cache"
         Response.AddHeader("Pragma", "no-cache")
         Response.Expires = -1

         If Not IsPostBack Then
            sFileExtensions = ""
            iMaxFileSize = 0
            sFileName = ""
            fAlreadyUploaded = False

            sTemp = Request.RawUrl.ToString
            iTemp = sTemp.IndexOf("?")
            sQueryString = ""
            If iTemp >= 0 Then
               sQueryString = sTemp.Substring(iTemp + 1)
               fAlreadyUploaded = (sQueryString.Substring(0, 1) = "1")

               sQueryString = sQueryString.Substring(1)
               sQueryString = objCrypt.SimpleDecrypt(sQueryString, Session.SessionID)
            End If

            hdnElementID.Value = sQueryString
            iElementItemID = CInt(sQueryString)

            Dim workflowUrl = CType(Session("workflowUrl"), WorkflowUrl)

            strConn = CStr("Application Name=OpenHR Workflow;Data Source=" & workflowUrl.Server & ";Initial Catalog=" & workflowUrl.Database & ";Integrated Security=false;User ID=" & workflowUrl.User & ";Password=" & workflowUrl.Password & ";Pooling=false")
            conn = New SqlClient.SqlConnection(strConn)
            conn.Open()
            Try
               cmdDetails = New SqlClient.SqlCommand
               cmdDetails.CommandText = "spASRGetWorkflowFileUploadDetails"
               cmdDetails.Connection = conn
               cmdDetails.CommandType = CommandType.StoredProcedure
               cmdDetails.CommandTimeout = miSubmissionTimeoutInSeconds

               cmdDetails.Parameters.Add("@piElementItemID", SqlDbType.Int).Direction = ParameterDirection.Input
               cmdDetails.Parameters("@piElementItemID").Value = iElementItemID

               cmdDetails.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
               cmdDetails.Parameters("@piInstanceID").Value = workflowUrl.InstanceID

               cmdDetails.Parameters.Add("@piSize", SqlDbType.Int).Direction = ParameterDirection.Output
               cmdDetails.Parameters.Add("@psFileName", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

               drDetails = cmdDetails.ExecuteReader

               While drDetails.Read
                  sFileExtensions = sFileExtensions & vbTab & "." & drDetails(0).ToString.ToLower
               End While
               If sFileExtensions.Length > 0 Then
                  sFileExtensions = sFileExtensions & vbTab
               End If

               drDetails.Close()

               iMaxFileSize = NullSafeInteger(cmdDetails.Parameters("@piSize").Value)
               If iMaxFileSize <= 0 Then
                  iMaxFileSize = 8000
               End If

               If fAlreadyUploaded Then
                  sFileName = NullSafeString(cmdDetails.Parameters("@psFileName").Value)
               End If

               cmdDetails.Dispose()

               ViewState("MaxFileSize") = iMaxFileSize
               ViewState("FileExtensions") = sFileExtensions
               ViewState("FileName") = sFileName
               ViewState("ElementItemID") = iElementItemID

               lblFileUploadPrompt.Font.Size = mobjConfig.MessageFontSize
               lblFileUploadPrompt.ForeColor = General.GetColour(6697779)

               lblErrors.Font.Size = mobjConfig.ValidationMessageFontSize
               lblErrors.ForeColor = General.GetColour(6697779)
               bulletErrors.Font.Size = mobjConfig.ValidationMessageFontSize
               bulletErrors.ForeColor = General.GetColour(6697779)

               btnCancel.Attributes.Add("onclick", "try{exitFileUpload(0);}catch(e){};")

               btnClear.Visible = (sFileName.Length > 0)

               With lblFileUploadPrompt
                  If (sFileName.Length > 0) Then
                     sFileNameWithoutPath = sFileName
                     iTemp = sFileNameWithoutPath.LastIndexOf("\")
                     If iTemp > 0 Then
                        sFileNameWithoutPath = sFileNameWithoutPath.Substring(iTemp + 1)
                     End If

                     .Text = "You have already uploaded '" & sFileNameWithoutPath & "'.<BR>" &
                             "Click 'Clear' to remove the uploaded file, or select a different file to upload in its place:"
                  Else
                     .Text = "Select the file you wish to upload:"
                  End If
               End With
            Finally
               conn.Close()
               conn.Dispose()

               ViewState("MaxFileSize") = iMaxFileSize
               ViewState("FileExtensions") = sFileExtensions
               ViewState("FileName") = sFileName
               ViewState("ElementItemID") = iElementItemID
            End Try
         End If
      Catch ex As Exception
         '' jpd handle the exception
      End Try
   End Sub

   Protected Sub btnFileUpload_ServerClick(sender As Object, e As EventArgs) Handles btnFileUpload.ServerClick

      Dim reader As BinaryReader
      Dim abtImage As Byte()
      Dim asErrorMessages() As String
      Dim sFileExt As String
      Dim iIndex As Integer
      Dim iMaxFileSize As Integer
      Dim sFileExtensions As String

      ReDim asErrorMessages(0)
      bulletErrors.Items.Clear()
      hdnCount_Errors.Value = "0"

      Try
         If FileUpload1.Value.Length > 0 Then
            iMaxFileSize = CInt(ViewState("MaxFileSize"))
            sFileExtensions = CStr(ViewState("FileExtensions"))

            'Check if the file has a valid size.
            If (iMaxFileSize > 0) _
               And (FileUpload1.PostedFile.ContentLength > (iMaxFileSize * 1024)) Then

               ReDim Preserve asErrorMessages(asErrorMessages.GetUpperBound(0) + 1)
               asErrorMessages(asErrorMessages.GetUpperBound(0)) = "The selected file exceeds the size limit of " & iMaxFileSize.ToString & " KB."
            End If

            'Check if the file is of a valid type.
            If sFileExtensions.Length > 0 Then
               sFileExt = vbTab & Path.GetExtension(FileUpload1.Value).ToLower & vbTab

               If (InStr(sFileExtensions, sFileExt) = 0) Then
                  ' Report lack of file selected.
                  ReDim Preserve asErrorMessages(asErrorMessages.GetUpperBound(0) + 1)
                  asErrorMessages(asErrorMessages.GetUpperBound(0)) = "Only files of the following type are permitted - " & Replace(Mid(sFileExtensions, 2, sFileExtensions.Length - 2), vbTab, ", ") & "."
               End If
            End If

            ' Check content type.
            If asErrorMessages.GetUpperBound(0) = 0 Then
               reader = New BinaryReader(FileUpload1.PostedFile.InputStream)
               abtImage = reader.ReadBytes(FileUpload1.PostedFile.ContentLength)
               SaveImage(abtImage, FileUpload1.PostedFile.ContentType, FileUpload1.PostedFile.FileName, False)
               hdnExitMode.Value = "2"
            End If
         Else
            ' Report lack of file selected.
            ReDim Preserve asErrorMessages(asErrorMessages.GetUpperBound(0) + 1)
            asErrorMessages(asErrorMessages.GetUpperBound(0)) = "No file selected."
         End If

      Catch ex As Exception
         ' Report exceptional failure.
         ReDim Preserve asErrorMessages(asErrorMessages.GetUpperBound(0) + 1)
         asErrorMessages(asErrorMessages.GetUpperBound(0)) = ex.Message
      End Try

      If asErrorMessages.GetUpperBound(0) > 0 Then
         hdnCount_Errors.Value = CStr(asErrorMessages.GetUpperBound(0))
         lblErrors.Text = "Unable to upload the file due to the following error" &
                          If(asErrorMessages.GetUpperBound(0) > 1, "s", "") & ":"

         For iIndex = 1 To asErrorMessages.GetUpperBound(0)
            bulletErrors.Items.Add(asErrorMessages(iIndex))
         Next
      End If
   End Sub

   Protected Sub btnClear_ServerClick(sender As Object, e As EventArgs) Handles btnClear.ServerClick

      Dim abtImage As Byte()
      ReDim abtImage(0)

      bulletErrors.Items.Clear()
      hdnCount_Errors.Value = "0"

      SaveImage(abtImage, "", "", True)
      hdnExitMode.Value = "1"
   End Sub
End Class
