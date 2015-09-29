Option Strict On
Option Explicit On

Imports Aspose.Words
Imports System.IO
Imports System.Net.Mail
Imports HR.Intranet.Server
Imports Aspose.Words.Reporting
Imports HR.Intranet.Server.Metadata

Namespace Code

	Public Class MailMergeRun
		Implements IFieldMergingCallback

		Public TemplateName As String
		Private _outputFileName As String
		Public EmailSubject As String
		Public EmailCalculationID As Integer
		Public IsAttachment As Boolean
		Public AttachmentName As String
		Public Name As String
		Public PrinterName As String

		Public MergeData As DataTable
		Public MergeDocument As MemoryStream

		Public Errors As New List(Of String)
		Public Columns As List(Of MergeColumn)

		Public Property OutputFileName As String
			Get
				If _outputFileName = "" Then
					Return String.Format("{0}.docx", Name)
				Else
					Return _outputFileName
				End If
			End Get
			Set(value As String)
				_outputFileName = value
			End Set
		End Property

#Region "Mail Merge Callback"

		Public Sub FieldMerging(args As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging

			If TypeOf (args.FieldValue) Is DateTime Then
				Dim sLocaleFormat = HttpContext.Current.Session("LocaleDateFormat").ToString()
				args.Text = String.Format("{0}", CDate(args.FieldValue).ToString(sLocaleFormat))
			ElseIf TypeOf (args.FieldValue) Is Boolean Then
				args.Text = IIf(CBool(args.FieldValue), "Y", "N").ToString()


			End If

		End Sub

		Public Sub ImageFieldMerging(args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
			Throw New NotImplementedException()
		End Sub

#End Region

		Public Function ExecuteToEmail() As Boolean
			Dim doc As Document

			Dim mailClient As SmtpClient
			Dim message As MailMessage
			Dim attachment As Attachment
			Dim strToEmail As String
			Dim objStream As MemoryStream

			Dim context As HttpContext = HttpContext.Current

			Dim objErrorLog As New clsEventLog(CType(context.Session("SessionContext"), SessionInfo).LoginInfo)
			Dim objDatabase As Database = CType(context.Session("DatabaseFunctions"), Database)
			Try
				Dim objWordLicense As New License
				objWordLicense.SetLicense("Aspose.Words.lic")

				Dim objTemplate = CType(HttpContext.Current.Session("MailMerge_Template"), Stream)

				If objTemplate Is Nothing Then
					Errors.Add("No template file selected")
					Return False
				End If

				'Check that we have a From field defined in IIS
				message = New MailMessage
				If message.From Is Nothing Then
					Errors.Add("No 'From' Email address has been defined in your configuration file.")
					Return False
				End If

				'Check that the From address is a valid email address
				If Not GeneralUtilities.IsValidEmailAddress(message.From.Address) Then
					Errors.Add("The 'From' Email address defined in your configuration file is not a valid email address")
					Return False
				End If

				mailClient = New SmtpClient	'Take SMTP settings from Web.config (i.e. SMTP settings defined for the website in IIS)

				For Each objRow As DataRow In MergeData.Rows
					objTemplate.Position = 0

					doc = New Document(objTemplate)
					doc.MailMerge.Execute(objRow)
					objStream = New MemoryStream()
					message = New MailMessage

					If IsAttachment Then
						' TODO - Support different output formats
						'Select Case Path.GetExtension(AttachmentName).ToLower()
						'	Case "pdf"
						'		doc.Save(objStream, SaveFormat.Pdf)
						'	Case Else
						doc.Save(objStream, SaveFormat.Docx)
						'End Select

						objStream.Seek(0, SeekOrigin.Begin) '"Rewind" the stream so it can be properly attached to the message; if it's no "rewinded" then the attachment is empty
						attachment = New Attachment(objStream, Path.GetFileName(AttachmentName))
						message.Attachments.Add(attachment)
					Else
						' TODO - Check that this is the correct format to handle images
						doc.Save(objStream, SaveFormat.Html)

						message.Body = doc.ToString(Aspose.Words.SaveFormat.Html)
						message.IsBodyHtml = True
						message.Attachments.Clear()
					End If

					message.Subject = EmailSubject

					' TODO - Alter this to read with initial dataset - would speed up performance
					strToEmail = objDatabase.GetEmailAddress(CInt(objRow("ID")), EmailCalculationID)

					If strToEmail.Length > 0 Then
						message.To.Add(strToEmail)
						mailClient.Send(message)

						' TODO - send emails async - means passing the async flag through multiple previous pages - needs some work!
						'mailClient.SendAsync(message, "OpenHR message")
					Else
						If objErrorLog.EventLogID = 0 Then
							objErrorLog = New clsEventLog(CType(context.Session("SessionContext"), SessionInfo).LoginInfo)
							objErrorLog.AddHeader(HR.Intranet.Server.Enums.EventLog_Type.eltMailMerge, Name)
						End If

						objErrorLog.AddDetailEntry("No email address found")
					End If
				Next
			Catch ex As Exception

				Dim errMessage As String
				If ex.InnerException Is Nothing Then
					errMessage = ""
				Else
					errMessage = ex.InnerException.Message
				End If

				Errors.Add(String.Format("The following error occured when emailing your document:" _
					& "{0}{0}{1}{0}{0}{2}{0}{0}Please check with your administrator for further details.", "<br/>", _
					ex.Message, errMessage))
				Return False
			End Try

			Return True

		End Function

		Public Function ExecuteMailMerge(DirectToPrinter As Boolean) As Boolean

			Try

				Dim objWordLicense As New License
				objWordLicense.SetLicense("Aspose.Words.lic")

				Dim objTemplate = CType(HttpContext.Current.Session("MailMerge_Template"), Stream)

				If objTemplate Is Nothing Then
					Errors.Add("No template file selected")
					Return False
				End If

				objTemplate.Position = 0

				Dim doc As New Document(objTemplate)
				doc.MailMerge.FieldMergingCallback = Me
				doc.MailMerge.Execute(MergeData)
				MergeDocument = New MemoryStream

				If DirectToPrinter And PrinterName.Length > 0 Then
					doc.Print(PrinterName)
				Else
					doc.Save(MergeDocument, SaveFormat.Docx)
				End If

				MergeDocument.Position = 0

			Catch ex As Exception
				Errors.Add(ex.Message)
				Return False

			End Try

			Return True
		End Function

		Public Function ValidateDefinition() As Boolean

			Dim duplicates = Columns.GroupBy(Function(i) i.MergeName.ToLower()) _
													.Where(Function(g) g.Count() > 1) _
													.[Select](Function(g) g.Key)

			' Check for dupliacte column names
			If duplicates.Count > 0 Then

				Errors.Add(String.Format("The following merge fields are duplicated in your definition:" _
							& "{0}{0}{1}{0}{0}Please edit your definition.", "<br/>", Join(duplicates.ToArray(), "<br/>")))
				Return False
			End If


			Return True

		End Function

		Public Function ValidateTemplate() As Boolean

			' Check for file access
			'If Not File.Exists(TemplateName) Then
			'	Errors.Add(String.Format("The file {0} cannot be found. {1}{1} Please ensure that the template file is a valid UNC path" _
			'				& " that is accessible from the OpenHR Web server.", TemplateName, "<br/>"))
			'	Return False
			'End If
			Try

				Dim objTemplate = CType(HttpContext.Current.Session("MailMerge_Template"), Stream)
				If objTemplate Is Nothing Then
					Errors.Add("No template file selected")
					Return False
				End If

				' Verify template integrity
				Dim doc As New Document(objTemplate)
				Dim templateFields = doc.MailMerge.GetFieldNames().Distinct().ToList()
				
				' If no template fields no point running the merge (Also stops corrupt files being used (e.g pdf))
				If templateFields.Count = 0 Then
					Errors.Add(String.Format("The uploaded template has no merge fields defined or is an invalid template.{0}" _
												, "<br/>"))
					Return False
				End If

				For Each objColumn In Columns
					templateFields.Remove(objColumn.MergeName)
				Next

				If templateFields.Count > 0 Then
					Errors.Add(String.Format("The uploaded template has the following merge fields which are missing from your definition:{0}{0}{1}{0}{0}Please edit the template or the definition." _
												, "<br/>", Join(templateFields.ToArray(), "<br/>")))
					Return False
				End If

			Catch ex As Exception
				Errors.Add(ex.Message)
				Return False

			End Try

			Return True

		End Function

	End Class
End Namespace