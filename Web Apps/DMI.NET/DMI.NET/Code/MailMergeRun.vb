Option Strict On
Option Explicit On

Imports Aspose.Words
Imports System.IO
Imports Aspose.Email.Mail
Imports System.ComponentModel
Imports HR.Intranet.Server

Namespace Code

	Public Class MailMergeRun

		Public TemplateName As String
		Public OutputFileName As String
		Public EmailSubject As String
		Public EmailCalculationID As Integer
		Public IsAttachment As Boolean
		Public AttachmentName As String
		Public Name As String

		Public MergeData As DataTable
		Public MergeDocument As MemoryStream

		Public Errors As New List(Of String)

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

			Return False

		End Function

		Private Shared Sub SendCompletedCallback(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)

				'Get the unique identifier for this asynchronous operation.
				Dim token As String = DirectCast(e.UserState, String)

				If e.Cancelled Then
						Console.WriteLine("[{0}] Send canceled.", token)
				End If

				If e.[Error] IsNot Nothing Then
						Console.WriteLine("[{0}] {1}", token, e.[Error].ToString())

				Else
						Console.WriteLine("Message Sent.")
				End If
		End Sub

		Public Function ExecuteMailMerge() As Boolean

			Try

				Dim doc As New Document(TemplateName)
				doc.MailMerge.Execute(MergeData)
				MergeDocument = New MemoryStream
				doc.Save(MergeDocument, SaveFormat.Docx)
				MergeDocument.Position = 0

			Catch ex As Exception
				Trace.WriteLine(ex.ToString())
				Throw

			End Try

			Return True
		End Function


		Public Function ValidateTemplate() As Boolean

			Dim bFileExits As Boolean
			Dim bFieldsOK As Boolean = True

			bFileExits = File.Exists(TemplateName)
			If Not bFileExits Then
				Errors.Add(String.Format("The file {0} cannot be found. {1}{1} Please ensure that the template file is a valid UNC path" _
							& " that is accessible from the OpenHR Web server.", TemplateName, "<br />"))
			End If

			If Not bFieldsOK Then
				Errors.Add("The template file does not match the mail merge definition. Please edit the template or the definition.")
			End If

			Return bFileExits And bFieldsOK

		End Function

	End Class
End Namespace