Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Utilities

Partial Class FileUpload
	Inherits System.Web.UI.Page

	Private mobjConfig As New Config
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
		Dim conn As System.Data.SqlClient.SqlConnection
		Dim cmdSave As System.Data.SqlClient.SqlCommand

		If abtImage.GetUpperBound(0) = 0 Then
			psContentType = ""
			psFileName = ""
		End If

		strConn = "Application Name=HR Pro Workflow;Data Source=" & Session("Server") & ";Initial Catalog=" & Session("Database") & ";Integrated Security=false;User ID=" & Session("User") & ";Password=" & Session("Pwd") & ";Pooling=false"
		conn = New SqlClient.SqlConnection(strConn)
		conn.Open()

		cmdSave = New SqlClient.SqlCommand("spASRWorkflowFileUpload", conn)
		cmdSave.CommandType = CommandType.StoredProcedure
		cmdSave.CommandTimeout = miSubmissionTimeoutInSeconds

		cmdSave.Parameters.AddWithValue("@piElementItemID", Me.ViewState("ElementItemID"))
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

	Protected Sub btnFileUpload_Click(ByVal sender As Object, ByVal e As Infragistics.WebUI.WebDataInput.ButtonEventArgs) Handles btnFileUpload.Click
		Dim reader As BinaryReader
		Dim abtImage As Byte()
		Dim asErrorMessages() As String
		Dim sFileExt As String
		Dim iIndex As Integer
		Dim iMaxFileSize As Integer
		Dim sFileExtensions As String

		ReDim asErrorMessages(0)
		bulletErrors.Items.Clear()
		hdnCount_Errors.Value = 0

		Try
			If FileUpload1.HasFile Then
				iMaxFileSize = Me.ViewState("MaxFileSize")
				sFileExtensions = Me.ViewState("FileExtensions")

				'Check if the file has a valid size.
				If (iMaxFileSize > 0) _
					And (FileUpload1.PostedFile.ContentLength > (iMaxFileSize * 1024)) Then

					ReDim Preserve asErrorMessages(asErrorMessages.GetUpperBound(0) + 1)
					asErrorMessages(asErrorMessages.GetUpperBound(0)) = "The selected file exceeds the size limit of " & iMaxFileSize.ToString & " KB."
				End If

				'Check if the file is of a valid type.
				If sFileExtensions.Length > 0 Then
					sFileExt = vbTab & System.IO.Path.GetExtension(FileUpload1.FileName).ToLower & vbTab

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
					hdnExitMode.Value = 2
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
			hdnCount_Errors.Value = asErrorMessages.GetUpperBound(0)
			lblErrors.Text = "Unable to upload the file due to the following error" & _
					IIf(asErrorMessages.GetUpperBound(0) > 1, "s", "") & ":"

			For iIndex = 1 To asErrorMessages.GetUpperBound(0)
				bulletErrors.Items.Add(asErrorMessages(iIndex))
			Next
		End If
	End Sub

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		Dim sTemp As String
		Dim iTemp As Integer
		Dim sQueryString As String
		Dim iElementItemID As Integer
		Dim strConn As String
		Dim drDetails As System.Data.SqlClient.SqlDataReader
		Dim cmdDetails As System.Data.SqlClient.SqlCommand
		Dim conn As System.Data.SqlClient.SqlConnection
		Dim iMaxFileSize As Integer
		Dim sFileExtensions As String
		Dim sFileName As String
		Dim sFileNameWithoutPath As String
		Dim objCrypt As New Crypt
		Dim fAlreadyUploaded As Boolean
		Dim objGeneral As New General

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
				iElementItemID = 0
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

				strConn = "Application Name=HR Pro Workflow;Data Source=" & Session("Server") & ";Initial Catalog=" & Session("Database") & ";Integrated Security=false;User ID=" & Session("User") & ";Password=" & Session("Pwd") & ";Pooling=false"
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
					cmdDetails.Parameters("@piInstanceID").Value = Session("InstanceID")

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
					drDetails = Nothing

					iMaxFileSize = NullSafeInteger(cmdDetails.Parameters("@piSize").Value)
					If iMaxFileSize <= 0 Then
						iMaxFileSize = 8000
					End If

					If fAlreadyUploaded Then
						sFileName = NullSafeString(cmdDetails.Parameters("@psFileName").Value)
					End If

					cmdDetails.Dispose()
					cmdDetails = Nothing

					Me.ViewState("MaxFileSize") = iMaxFileSize
					Me.ViewState("FileExtensions") = sFileExtensions
					Me.ViewState("FileName") = sFileName
					Me.ViewState("ElementItemID") = iElementItemID

					lblFileUploadPrompt.Font.Size = mobjConfig.MessageFontSize
					lblFileUploadPrompt.ForeColor = objGeneral.GetColour(6697779)

					FileUpload1.Font.Size = mobjConfig.MessageFontSize
					FileUpload1.Font.Name = "Verdana"
					FileUpload1.ForeColor = objGeneral.GetColour(6697779)
					FileUpload1.BackColor = objGeneral.GetColour(15988214)
					FileUpload1.BorderStyle = BorderStyle.Solid
					FileUpload1.BorderWidth = 1
					FileUpload1.BorderColor = objGeneral.GetColour(5730458)

					lblErrors.Font.Size = mobjConfig.ValidationMessageFontSize
					lblErrors.ForeColor = objGeneral.GetColour(6697779)
					bulletErrors.Font.Size = mobjConfig.ValidationMessageFontSize
					bulletErrors.ForeColor = objGeneral.GetColour(6697779)

					With btnCancel
						.Appearance.Style.BackColor = objGeneral.GetColour(16249587)
						.Appearance.Style.BorderStyle = BorderStyle.Solid
						.Appearance.Style.BorderWidth = 1
						.Appearance.InnerBorder.StyleTop = BorderStyle.None
						.Appearance.Style.BorderColor = objGeneral.GetColour(10720408)
						.Appearance.Style.ForeColor = objGeneral.GetColour(6697779)
						.FocusAppearance.Style.BorderColor = objGeneral.GetColour(562943)
						.FocusAppearance.Style.BackColor = objGeneral.GetColour(12775933)
						.HoverAppearance.Style.BorderColor = objGeneral.GetColour(562943)

						.Font.Name = "Verdana"
						.ClientSideEvents.Click = "try{exitFileUpload(0);}catch(e){};"
					End With

					With btnClear
						.Appearance.Style.BackColor = objGeneral.GetColour(16249587)
						.Appearance.Style.BorderStyle = BorderStyle.Solid
						.Appearance.Style.BorderWidth = 1
						.Appearance.InnerBorder.StyleTop = BorderStyle.None
						.Appearance.Style.BorderColor = objGeneral.GetColour(10720408)
						.Appearance.Style.ForeColor = objGeneral.GetColour(6697779)
						.FocusAppearance.Style.BorderColor = objGeneral.GetColour(562943)
						.FocusAppearance.Style.BackColor = objGeneral.GetColour(12775933)
						.HoverAppearance.Style.BorderColor = objGeneral.GetColour(562943)

						.Font.Name = "Verdana"
						.Visible = (sFileName.Length > 0)
					End With

					With btnFileUpload
						.Appearance.Style.BackColor = objGeneral.GetColour(16249587)
						.Appearance.Style.BorderStyle = BorderStyle.Solid
						.Appearance.Style.BorderWidth = 1
						.Appearance.InnerBorder.StyleTop = BorderStyle.None
						.Appearance.Style.BorderColor = objGeneral.GetColour(10720408)
						.Appearance.Style.ForeColor = objGeneral.GetColour(6697779)
						.FocusAppearance.Style.BorderColor = objGeneral.GetColour(562943)
						.FocusAppearance.Style.BackColor = objGeneral.GetColour(12775933)
						.HoverAppearance.Style.BorderColor = objGeneral.GetColour(562943)

						.Font.Name = "Verdana"
					End With

					With lblFileUploadPrompt
						If (sFileName.Length > 0) Then
							sFileNameWithoutPath = sFileName
							iTemp = sFileNameWithoutPath.LastIndexOf("\")
							If iTemp > 0 Then
								sFileNameWithoutPath = sFileNameWithoutPath.Substring(iTemp + 1)
							End If

							.Text = "You have already uploaded '" & sFileNameWithoutPath & "'.<BR>" & _
									"Click 'Clear' to remove the uploaded file, or select a different file to upload in its place:"
						Else
							.Text = "Select the file you wish to upload:"
						End If
					End With
				Finally
					conn.Close()
					conn.Dispose()

					Me.ViewState("MaxFileSize") = iMaxFileSize
					Me.ViewState("FileExtensions") = sFileExtensions
					Me.ViewState("FileName") = sFileName
					Me.ViewState("ElementItemID") = iElementItemID
				End Try
			End If
		Catch ex As Exception
			'' jpd handle the exception
		End Try
	End Sub

	Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As Infragistics.WebUI.WebDataInput.ButtonEventArgs) Handles btnClear.Click
		Dim abtImage As Byte()
		ReDim abtImage(0)

		bulletErrors.Items.Clear()
		hdnCount_Errors.Value = 0

		SaveImage(abtImage, "", "", True)
		hdnExitMode.Value = 1

	End Sub


End Class
