Option Strict On
Option Explicit On

Imports DMI.NET.Code.Attributes
Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel

Namespace Models

	Public Class ReportOutputModel

		Public Property ReportType() As UtilityType

		<Required>
		<DisplayName("Output Format :")>
		Public Property Format As OutputFormats

		<DisplayName("Preview on screen")>
		Public Property IsPreview As Boolean

		<DisplayName("Display output on screen")>
		Public Property ToScreen As Boolean

		<DisplayName("Send to printer")>
		Public Property ToPrinter As Boolean

		<AllowHtml>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<DisplayName("Printer name :")>
		Public Property PrinterName As String

		<DisplayName("Save to file")>
		Public Property SaveToFile As Boolean

		<RequiredIf("SaveToFile", True, ErrorMessage:="No filename entered.")>
		<MaxLength(255, ErrorMessage:="File Name cannot be longer than 255 characters.")>
		<DisplayName("File Name :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<ExcludeChar("/*?""<>|")>
		<AllowHtml>
		Public Property Filename As String

		<DisplayName("If file exists :")>
		Public Property SaveExisting As ExistingFile

		<DisplayName("Send as email")>
		Public Property SendToEmail As Boolean

		<NoneAttribute("EmailGroupName", "None", ErrorMessage:="No email group selected.")>
		Public Property EmailGroupID As Integer

		<DisplayName("Email Group :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailGroupName As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailAddress As String

		<DisplayName("Subject :")>
		<RequiredIf("SendToEmail", True, ErrorMessage:="No email subject name entered.")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<ExcludeChar("/*?""<>|")>
		<AllowHtml>
		Public Property EmailSubject As String

		<DisplayName("Attach As :")>
		<RequiredIf("SendToEmail", True, ErrorMessage:="No email attachment name entered.")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<ExcludeChar("/*?""<>|")>
		<AllowHtml>
		Public Property EmailAttachmentName As String

		<RegularExpression("True", ErrorMessage:="No output destination selected.")>
		Public ReadOnly Property IsDestinationOK As Boolean
			Get
				Return (IsPreview OrElse ToScreen OrElse SaveToFile OrElse ToPrinter OrElse SendToEmail)
			End Get
		End Property

		<RegularExpression("True", ErrorMessage:="You must select a destination in addition to preview.")>
	 Public ReadOnly Property IsOtherThanPreviewOK As Boolean
			Get
				If IsPreview Then
					Return (ToScreen OrElse SaveToFile OrElse ToPrinter)
				Else
					Return True
				End If

			End Get
		End Property



	End Class

End Namespace