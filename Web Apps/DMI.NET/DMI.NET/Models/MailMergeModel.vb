Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata
Imports DMI.NET.Classes
Imports DMI.NET.ViewModels
Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports HR.Intranet.Server.Enums
Imports DMI.NET.ViewModels.Reports

Namespace Models

	Public Class MailMergeModel
		Inherits ReportBaseModel

		Public Overrides ReadOnly Property ReportType As UtilityType
			Get
				Return UtilityType.utlMailMerge
			End Get
		End Property

		<DisplayName("Template"), DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property TemplateFileName As String

		<DisplayName("Pause before merge")>
		Public Property PauseBeforeMerge As Boolean

		<DisplayName("Suppress blank lines")>
		Public Property SuppressBlankLines As Boolean
		Public Property OutputFormat As MailMergeOutputTypes

		<DisplayName("Display output on screen")>
		Public Property DisplayOutputOnScreen As Boolean

		<DisplayName("Send to printa")>
		Public Property SendToPrinter As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property PrinterName As String
		Public Property SaveTofile As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property Filename As String

		<DisplayName("Email Address")>
		Public Property EmailGroupID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property EmailSubject As String

		<DisplayName("Send As Attachment")>
		Public Property EmailAsAttachment As Boolean

		<Required(ErrorMessage:="Name is required.")> _
		<MaxLength(5, ErrorMessage:="Name cannot be longer than 5 characters.")> _
		<DisplayName("Attach As")>
		Public Property EmailAttachmentName As String

		Public Overrides Sub SetBaseTable(TableID As Integer)
		End Sub

	End Class

End Namespace