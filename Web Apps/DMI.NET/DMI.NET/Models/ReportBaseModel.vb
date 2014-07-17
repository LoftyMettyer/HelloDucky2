Option Explicit On
Option Strict On

Imports DMI.NET.Classes
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports DMI.NET.Enums
Imports System.ComponentModel.DataAnnotations
Imports DMI.NET.AttributeExtensions
Imports HR.Intranet.Server.Enums
Imports System.Web.Script.Serialization
Imports HR.Intranet.Server
Imports DMI.NET.ViewModels.Reports

Namespace Models
	Public MustInherit Class ReportBaseModel
		Implements IDataAccess
		Implements IReport

		Public Property IsReadOnly As Boolean
		Public MustOverride ReadOnly Property ReportType As UtilityType

		Public Property ID As Integer
		Public Property Owner As String

		<Required(ErrorMessage:="A base table must be selected.")>
		Public Property BaseTableID As Integer Implements IReport.BaseTableID

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

		Public Property SortOrders As New Collection(Of SortOrderViewModel) Implements IReport.SortOrders
		Public Property Repetition As New Collection(Of ReportRepetition)

		Public Property SortOrdersString As String

		Public Property JobsToHide As New Collection(Of Integer)

		Public Property SessionContext As SessionInfo Implements IDataAccess.SessionContext

		Public Property SessionInfo As SessionInfo Implements IReport.SessionInfo
		Public MustOverride Sub SetBaseTable(BaseTableID As Integer) Implements IReport.SetBaseTable
		Public MustOverride Function GetAvailableSortColumns() As IEnumerable(Of ReportColumnItem) Implements IReport.GetAvailableSortColumns

		<DisplayName("Description 1: ")>
		Public Property Description1ID As Integer

		<DisplayName("Description 2: ")>
		Public Property Description2ID As Integer

		<DisplayName("Description 3: ")>
		Public Property Description3ID As Integer

		Public Property RegionID As Integer
		Public Property GroupByDescription As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property Separator As String


	End Class
End Namespace