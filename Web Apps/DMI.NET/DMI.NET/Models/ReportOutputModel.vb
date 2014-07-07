Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports DMI.NET.Enums
Imports DMI.NET.Classes
Imports System.ComponentModel.DataAnnotations
Imports DMI.NET.AttributeExtensions
Imports System.ComponentModel

Namespace Models

	Public Class ReportOutputModel

		<Required>
		<DisplayName("Output Format :")>
		Public Property Format As OutputFormats

		Public Property IsPreview As Boolean
		Public Property ToScreen As Boolean
		Public Property ToPrinter As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property PrinterName As String

		Public Property SaveToFile As Boolean

		<RequiredIf("SaveToFile", True, ErrorMessage:="No filename entered.")>
		<MaxLength(255, ErrorMessage:="File Name cannot be longer than 255 characters.")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property Filename As String

		Public Property SaveExisting As ExistingFile

		<DisplayName("Send As email")>
		Public Property SendToEmail As Boolean

		<NonZeroIf("SendToEmail", True, ErrorMessage:="No email group selected.")> _
		Public Property EmailGroupID As Integer

		<DisplayName("Email Group :")>
		Public Property EmailGroupName As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property EmailAddress As String

		<DisplayName("Email Subject :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailSubject As String

		<RequiredIf("SendToEmail", True, ErrorMessage:="No email attachment name entered.")>
		<DisplayName("Attach As :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailAttachmentName As String

	End Class

End Namespace