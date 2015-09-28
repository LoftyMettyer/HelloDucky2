Option Strict On
Option Explicit On

Imports System.Collections.ObjectModel
Imports DMI.NET.ViewModels

Namespace Models.ObjectRequests
	Public Class DefSelModel

		Public Property __RequestVerificationToken As String

		Public Property txtTableID As Integer

		Public Property utiltype As UtilityType
		Public Property utilID As Integer

		<AllowHtml>
		Public Property utilName As String

		Public Property Action As String

		Public Property txtGotoFromMenu As Boolean
		Public Property RecordID As Integer
		Public Property OnlyMine As Boolean

		Public Property Usage As Collection(Of DefinitionPropertiesViewModel)
		Public Property Status As String

		Public Property MultipleRecordIDs As String

	End Class
End Namespace
