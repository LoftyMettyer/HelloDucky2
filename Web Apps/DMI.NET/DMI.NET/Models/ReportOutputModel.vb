Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports DMI.NET.AttributeExtensions
Imports System.ComponentModel

Namespace Models

	Public Class ReportOutputModel

		<Required>
		<DisplayName("Output Format :")>
		Public Property Format As OutputFormats

		<DisplayName("Preview on screen")>
		Public Property IsPreview As Boolean

		<DisplayName("Display output on screen")>
		Public Property ToScreen As Boolean

		<DisplayName("Send to printer")>
		Public Property ToPrinter As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<DisplayName("Printer name :")>
		Public Property PrinterName As String

		<DisplayName("Save to file")>
		Public Property SaveToFile As Boolean

		<RequiredIf("SaveToFile", True, ErrorMessage:="No filename entered.")>
		<MaxLength(255, ErrorMessage:="File Name cannot be longer than 255 characters.")>
		<DisplayName("File Name :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<ExcludeChar("\/:*?""<>|")>
		Public Property Filename As String

		<DisplayName("If file exists:")>
		Public Property SaveExisting As ExistingFile

		<DisplayName("Send As email")>
		Public Property SendToEmail As Boolean

		<NonZeroIf("SendToEmail", True, ErrorMessage:="No email group selected.")>
		Public Property EmailGroupID As Integer

		<DisplayName("Email Group :")>
		Public Property EmailGroupName As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailAddress As String

		<DisplayName("Subject :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailSubject As String

		<RequiredIf("SendToEmail", True, ErrorMessage:="No email attachment name entered.")>
		<DisplayName("Attach As :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailAttachmentName As String

	End Class

End Namespace