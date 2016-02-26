Imports System.Data

Partial Class FileDownload
	Inherits Page

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
		Dim sTemp As String
		Dim iTemp As Integer
		Dim sQueryString As String
		Dim iElementItemID As Long
		Dim objCrypt As New Crypt
		Dim sErrorMessage As String

		sErrorMessage = ""

		Try
			sTemp = Request.RawUrl.ToString
			iTemp = sTemp.IndexOf("?")
			sQueryString = ""
			If iTemp >= 0 Then
				sQueryString = sTemp.Substring(iTemp + 1)
				sQueryString = objCrypt.SimpleDecrypt(sQueryString, Session.SessionID)
			End If
			iElementItemID = CLng(sQueryString)

			sErrorMessage = ShowTheFile(iElementItemID)
		Catch ex As Exception
			sErrorMessage = ex.Message.Replace(vbCrLf, "<BR>")
		End Try

		If sErrorMessage.Length > 0 Then
			If sErrorMessage.Length > 0 Then
				sErrorMessage = "An error occurred."	' Quick fix to fix an XSS issue.
				sErrorMessage = "Unable to download file.<BR>" & sErrorMessage & "<BR><BR>"
			End If

			Session("message") = sErrorMessage
			Server.Transfer("Message.aspx")
		End If
	End Sub

	Private Function ShowTheFile(ByVal piElementItemID As Long) As String

		Dim conn As System.Data.SqlClient.SqlConnection
		Dim cmdRead As System.Data.SqlClient.SqlCommand
		Dim drFile As System.Data.SqlClient.SqlDataReader
		Dim abtImage As Byte()
		Dim sContentType As String
		Dim sFileName As String
		Dim iIndex As Integer
		Dim iItemType As Integer
		Dim sTemp As String
		Dim abtTemp As Byte()
		Dim iOffset As Integer
		Dim iOLEType As Integer
		Dim iColumnOLEType As Integer
		Dim iColumnDataType As Integer
		Dim sUNC As String
		Dim sFilePath As String
		Dim sFullFilePath As String
		Dim sOLEFolder_Server As String
		Dim sOLEFolder_Local As String
		Dim sPhotographFolder As String
		Dim sErrorMessage As String
		Dim fFileOK As Boolean

		sContentType = ""
		sFileName = ""
		sUNC = ""
		sFilePath = ""
		iOffset = 0
		sErrorMessage = ""
		fFileOK = False
		ReDim abtImage(0)

		Dim workflowUrl = CType(Session("workflowUrl"), WorkflowUrl)

		Try
			conn = New SqlClient.SqlConnection(App.Config.ConnectionString)
			conn.Open()

			cmdRead = New SqlClient.SqlCommand("spASRWorkflowFileDownload", conn)
			cmdRead.CommandType = CommandType.StoredProcedure
			cmdRead.CommandTimeout = App.Config.SubmissionTimeoutInSeconds

			cmdRead.Parameters.AddWithValue("@piElementItemID", piElementItemID)
			cmdRead.Parameters.AddWithValue("@piInstanceID", workflowUrl.InstanceID)

			cmdRead.Parameters.Add("@piItemType", SqlDbType.Int).Direction = ParameterDirection.Output
			cmdRead.Parameters.Add("@psFileName", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
			cmdRead.Parameters.Add("@psContentType", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
			cmdRead.Parameters.Add("@psErrorMessage", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
			cmdRead.Parameters.Add("@piOLEType", SqlDbType.Int).Direction = ParameterDirection.Output
			cmdRead.Parameters.Add("@piDBColumnType", SqlDbType.Int).Direction = ParameterDirection.Output

			drFile = cmdRead.ExecuteReader

			drFile.Read()

			If drFile.HasRows Then
				' Only interested in this for linked/embedded.
				If (Not drFile("file") Is DBNull.Value) Then
					abtImage = CType(drFile("file"), Byte())
					fFileOK = (abtImage.GetLength(0) > 0)
				End If
			End If

			drFile.Close()

			iItemType = NullSafeInteger(cmdRead.Parameters("@piItemType").Value)
			If iItemType = 19 Then
				' 19 = DB File
				' HR Pro OLE FORMAT...
				'   8 characters denoting format version (eg "<<V002>>")
				'   2 characters denoting OLE type (02 = Embedded document, 03 = UNC link)
				'   70 characters denoting filename
				'   210 character denoting file path
				'   60 characters for UNC
				'   10 characters for file size
				'   20 characters for file create date (in format dd/MM/yyyy HH:MM:SS)
				'   20 characters for file last modified date (in format dd/MM/yyyy HH:MM:SS)
				'   Remainder of the data is the contents of the embedded document.
				iColumnOLEType = NullSafeInteger(cmdRead.Parameters("@piOLEType").Value)
				iColumnDataType = NullSafeInteger(cmdRead.Parameters("@piDBColumnType").Value)

				If (iColumnOLEType = 0) Or (iColumnOLEType = 1) Then
					sFileName = NullSafeString(cmdRead.Parameters("@psFileName").Value)

					If sFileName.Length = 0 Then
						sErrorMessage = "No file to download."
					Else
						If iColumnDataType = -3 Then
							' Photograph
							sPhotographFolder = App.Config.PhotographFolder().Trim

							If sPhotographFolder.Length = 0 Then
								sErrorMessage = "The Photograph path has not been defined, or is invalid."
							Else
								sFullFilePath = sPhotographFolder _
																& If(Right(sPhotographFolder, 1) = "\", "", "\") _
																& sFileName

								Try
									abtImage = My.Computer.FileSystem.ReadAllBytes(sFullFilePath)
									fFileOK = (abtImage.GetLength(0) > 0)
								Catch ex As Exception
									sErrorMessage = "The specified file no longer exists, or access is denied."
								End Try
							End If
						Else
							Select Case iColumnOLEType
								Case 0
									' Local
									sOLEFolder_Local = App.Config.OLEFolderLocal()

									If sOLEFolder_Local.Length = 0 Then
										sErrorMessage = "The Local OLE path has not been defined, or is invalid."
									Else
										sFullFilePath = sOLEFolder_Local _
																		& If(Right(sOLEFolder_Local, 1) = "\", "", "\") _
																		& sFileName

										Try
											abtImage = My.Computer.FileSystem.ReadAllBytes(sFullFilePath)
											fFileOK = (abtImage.GetLength(0) > 0)
										Catch ex As Exception
											sErrorMessage = "The specified file no longer exists, or access is denied."
										End Try
									End If

								Case 1
									'Server
									sOLEFolder_Server = App.Config.OLEFolderServer()

									If sOLEFolder_Server.Length = 0 Then
										sErrorMessage = "The " & If(iColumnDataType = -3, "Photograph", "Server OLE") &
																		" path has not been defined, or is invalid."
									Else

										sFullFilePath = sOLEFolder_Server _
																		& If(Right(sOLEFolder_Server, 1) = "\", "", "\") _
																		& sFileName

										Try
											abtImage = My.Computer.FileSystem.ReadAllBytes(sFullFilePath)
											fFileOK = (abtImage.GetLength(0) > 0)
										Catch ex As Exception
											sErrorMessage = "The specified file no longer exists, or access is denied."
										End Try
									End If
							End Select
						End If
					End If
				ElseIf fFileOK Then
					ReDim abtTemp(2)
					Array.ConstrainedCopy(abtImage, 8, abtTemp, 0, 2)
					sTemp = Encoding.ASCII.GetString(abtTemp).Trim
					iOLEType = CInt(Left(sTemp, sTemp.Length - 1))

					ReDim abtTemp(70)
					Array.ConstrainedCopy(abtImage, 10, abtTemp, 0, 70)
					sTemp = Encoding.ASCII.GetString(abtTemp).Trim
					sFileName = Left(sTemp, sTemp.Length - 1).Trim

					ReDim abtTemp(210)
					Array.ConstrainedCopy(abtImage, 80, abtTemp, 0, 210)
					sTemp = Encoding.ASCII.GetString(abtTemp).Trim
					sFilePath = Left(sTemp, sTemp.Length - 1).Trim

					ReDim abtTemp(60)
					Array.ConstrainedCopy(abtImage, 290, abtTemp, 0, 60)
					sTemp = Encoding.ASCII.GetString(abtTemp).Trim
					sUNC = Left(sTemp, sTemp.Length - 1).Trim

					'ReDim abtTemp(10)
					'Array.ConstrainedCopy(abtImage, 350, abtTemp, 0, 10)
					'sTemp = System.Text.Encoding.ASCII.GetString(abtTemp).Trim

					'ReDim abtTemp(20)
					'Array.ConstrainedCopy(abtImage, 360, abtTemp, 0, 20)
					'sTemp = System.Text.Encoding.ASCII.GetString(abtTemp).Trim

					'ReDim abtTemp(20)
					'Array.ConstrainedCopy(abtImage, 380, abtTemp, 0, 20)
					'sTemp = System.Text.Encoding.ASCII.GetString(abtTemp).Trim

					Select Case iOLEType
						Case 2
							' Embedded
							iOffset = 400
						Case 3
							' UNC
							If (sUNC.Length = 0) _
								 Or (sFileName.Length = 0) Then
								sErrorMessage = "No file to download."
							Else
								sFullFilePath = sUNC & sFilePath & "\" & sFileName

								Try
									abtImage = My.Computer.FileSystem.ReadAllBytes(sFullFilePath)
									fFileOK = (abtImage.GetLength(0) > 0)
								Catch ex As Exception
									sErrorMessage = "The specified file no longer exists, or access is denied."
								End Try
							End If
					End Select
				End If
			Else
				' Else (20) = WF File
				sFileName = NullSafeString(cmdRead.Parameters("@psFileName").Value)
				sContentType = NullSafeString(cmdRead.Parameters("@psContentType").Value)
			End If

			cmdRead.Dispose()
			conn.Close()

			If (sErrorMessage.Length = 0) _
				 And (Not fFileOK) Then
				sErrorMessage = "No file to download."
			End If

			If sErrorMessage.Length = 0 Then
				If sContentType.Length = 0 Then
					sContentType = General.ContentTypeFromExtension(sFileName)
				End If

				iIndex = InStrRev(sFileName, "\")
				If iIndex > 0 Then
					sFileName = Mid(sFileName, iIndex + 1)
				End If

				Response.Clear()
				Response.ClearHeaders()
				Response.AddHeader("content-disposition", "attachment; filename='" + sFileName + "'")
				Response.ClearContent()
				Response.ContentEncoding = Encoding.UTF8
				Response.ContentType = sContentType
				Response.OutputStream.Write(abtImage, iOffset, abtImage.Length - iOffset)
				Response.Flush()
				Response.Close()
			End If
		Catch ex As Exception
			sErrorMessage = ex.Message.Replace(vbCrLf, "<BR>")
		End Try

		Return sErrorMessage
	End Function
End Class
