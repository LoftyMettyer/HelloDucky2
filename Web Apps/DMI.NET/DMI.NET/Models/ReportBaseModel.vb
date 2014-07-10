Option Explicit On
Option Strict On

Imports DMI.NET.Classes
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports DMI.NET.Enums
Imports System.ComponentModel.DataAnnotations
Imports DMI.NET.AttributeExtensions
Imports HR.Intranet.Server.Enums

Namespace Models
	Public MustInherit Class ReportBaseModel

		Public Property IsReadOnly As Boolean
		Public MustOverride ReadOnly Property ReportType As UtilityType

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

		<DisplayName("Display Title In Report Header")>
		Public Property DisplayTitleInReportHeader As Boolean

		Public Property SortOrderColumns As New Collection(Of ReportSortItem)
		Public Property Repetition As New Collection(Of ReportRepetition)

		Public Property JobsToHide As New Collection(Of Integer)

	End Class
End Namespace