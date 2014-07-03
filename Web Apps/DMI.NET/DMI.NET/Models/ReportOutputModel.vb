Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports DMI.NET.Enums
Imports DMI.NET.Classes
Imports System.ComponentModel.DataAnnotations

Namespace Models

	Public Class ReportOutputModel

		Public Property Format As OutputFormats
		Public Property IsPreview As Boolean
		Public Property ToScreen As Boolean
		Public Property ToPrinter As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property PrinterName As String
		Public Property SaveToFile As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property Filename As String
		Public Property SaveExisting As ExistingFile
		Public Property EmailGroupID As Integer
		Public Property SendToEmail As Boolean
		Public Property EmailGroupName As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property EmailAddress As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property EmailSubject As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property EmailAttachmentName As String

	End Class

End Namespace