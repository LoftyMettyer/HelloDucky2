Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel
Imports DMI.NET.Classes
Imports DMI.NET.Enums
Imports DMI.NET.AttributeExtensions
Imports System.Collections.ObjectModel

Namespace ViewModels.Reports

	Public Class DefinitionViewModel

		Public Property IsReadOnly As Boolean
		Public Property ReportType As UtilityType

		Public Property ID As Integer
		Public Property Owner As String

		<Required(ErrorMessage:="A base table must be selected.")>
		Public Property BaseTableID As Integer

		<Required(ErrorMessage:="Name is required.")>
		<MaxLength(50, ErrorMessage:="Name cannot be longer than 50 characters.")>
		<DisplayName("Name :")>
		Public Property Name As String

		<MaxLength(255, ErrorMessage:="Description cannot be longer than 255 characters.")>
		<DisplayName("Description :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property Description As String

		Public Property GroupAccess As New Collection(Of GroupAccess)
		Public Property SelectionType As RecordSelectionType

		<NonZeroIf("SelectionType", RecordSelectionType.Filter, ErrorMessage:="No filter selected for base table.")> _
		Public Property FilterID As Integer

		<NonZeroIf("SelectionType", RecordSelectionType.Picklist, ErrorMessage:="No picklist selected for base table.")>
		Public Property PicklistID As Integer

		Public Property FilterName As String
		Public Property PicklistName As String

		Public Property BaseTables As New List(Of ReportTableItem)

	End Class

End Namespace