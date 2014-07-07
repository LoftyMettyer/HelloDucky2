Option Explicit On
Option Strict On

Imports DMI.NET.Classes
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports DMI.NET.Enums
Imports System.ComponentModel.DataAnnotations

Namespace Models
	Public Class ReportBaseModel

		Public Property ID As Integer
		Public Property BaseTableID As Integer
		Public Property Owner As String

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
		Public Property FilterID As Integer
		Public Property PicklistID As Integer

		Public Property FilterName As String
		Public Property PicklistName As String

		Public Property BaseTables As New Collection(Of SelectListItem)

		Public Property DisplayTitleInReportHeader As Boolean

		Public Property SortOrderColumns As New Collection(Of ReportSortItem)

		Public Property JobsToHide As New Collection(Of Integer)

	End Class
End Namespace