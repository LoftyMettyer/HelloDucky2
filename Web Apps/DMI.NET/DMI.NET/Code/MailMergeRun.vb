Option Strict On
Option Explicit On

Imports Aspose.Words
Imports System.IO
Imports Aspose.Email.Mail
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
			End If

		End Sub

		Public Sub ImageFieldMerging(args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
			Throw New NotImplementedException()
		End Sub

#End Region

		Public Function ExecuteToEmail() As Boolean

			Dim doc As Document

			Dim mailClient As Aspose.Email.Mail.SmtpClient
			Dim message As Aspose.Email.Mail.MailMessage
			Dim attachment As Aspose.Email.Mail.Attachment
			Dim strToEmail As String
			Dim objStream As MemoryStream
			Dim objOptions As New MailMessageLoadOptions

			Dim context As HttpContext = HttpContext.Current

			Dim objErrorLog As New clsEventLog(CType(context.Session("SessionContext"), SessionInfo).LoginInfo)
			Dim objDatabase As Database = CType(context.Session("DatabaseFunctions"), Database)

			Try

				mailClient = New Aspose.Email.Mail.SmtpClient(ApplicationSettings.SMTP_Host, ApplicationSettings.SMTP_Port)

				objOptions.MessageFormat = MessageFormat.Mht

				For Each objRow As DataRow In MergeData.Rows
					doc = New Document(TemplateName)
					doc.MailMerge.Execute(objRow)
					objStream = New MemoryStream()

					If IsAttachment Then

						' TODO - Support different output formats
						'Select Case Path.GetExtension(AttachmentName).ToLower()
						'	Case "pdf"
						'		doc.Save(objStream, SaveFormat.Pdf)
						'	Case Else
						doc.Save(objStream, SaveFormat.Docx)
						'End Select

						attachment = New Aspose.Email.Mail.Attachment(objStream, Path.GetFileName(AttachmentName))
						message = New Aspose.Email.Mail.MailMessage
						message.Attachments.Add(attachment)
						message.Body = ""

					Else

						' TODO - Check that this is the correct format to handle images
						doc.Save(objStream, SaveFormat.Mhtml)
						objStream.Position = 0
						message = MailMessage.Load(objStream, objOptions)
						message.Attachments.Clear()

					End If

					message.Subject = EmailSubject
					message.From = New Aspose.Email.Mail.MailAddress(ApplicationSettings.MailMerge_From, "OpenHR")

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

				Errors.Add(String.Format("The following error occured when emailing your document" _
							& "{0}{0}{1}{0}{0}Please check with your administrator for further details", "<br/>", ex.Message))

				Return False


			End Try

			Return True

		End Function

		Public Function ExecuteMailMerge(DirectToPrinter As Boolean) As Boolean

			Try

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

				If DirectToPrinter Then
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