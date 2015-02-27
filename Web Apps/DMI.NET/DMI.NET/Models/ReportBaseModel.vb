Option Explicit On
Option Strict On

Imports DMI.NET.Code.Attributes
Imports DMI.NET.Classes
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports HR.Intranet.Server
Imports DMI.NET.ViewModels.Reports

Namespace Models
	Public MustInherit Class ReportBaseModel
		Implements IDataAccess
		Implements IReport

		Public Property IsReadOnly As Boolean
		Public MustOverride ReadOnly Property ReportType As UtilityType Implements IReport.ReportType

		Public Property ID As Integer Implements IReport.ID
		Public Property Owner As String Implements IReport.Owner

		Public Property ActionType As UtilityActionType
		Public Property Timestamp As Long
		Public Property ValidityStatus As ReportValidationStatus = ReportValidationStatus.InvalidOnClient

		<Required(ErrorMessage:="A base table must be selected.")>
		Public Property BaseTableID As Integer Implements IReport.BaseTableID
		Public Property BaseViewAccess As String

		<Required(ErrorMessage:="Definition name is required.")>
		<MaxLength(50, ErrorMessage:="Definition name cannot be longer than 50 characters.")>
		<DisplayName("Name :")>
		Public Property Name As String

		<MaxLength(255, ErrorMessage:="Description cannot be longer than 255 characters.")>
		<DisplayName("Description :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property Description As String

		<DisplayName("Ignore zeros when calculating aggregates")>
		Public Property IgnoreZerosForAggregates As Boolean

		<DisplayName("Summary Report")>
		Public Property IsSummary As Boolean

		Public Property GroupAccess As New Collection(Of GroupAccess)
		Public Property SelectionType As RecordSelectionType

		<NonZeroIf("SelectionType", RecordSelectionType.Filter, ErrorMessage:="No filter selected for base table.")> _
		Public Property FilterID As Integer

		<NonZeroIf("SelectionType", RecordSelectionType.Picklist, ErrorMessage:="No picklist selected for base table.")>
		Public Property PicklistID As Integer

		Public Property FilterName As String
		Public Property PicklistName As String

		<DisplayName("Display filter or picklist title in the report header")>
		Public Property DisplayTitleInReportHeader As Boolean

		Public Property Columns As New List(Of ReportColumnItem) Implements IReport.Columns
		Public Overridable Property ColumnsAsString As String
		Public Property SortOrders As New List(Of SortOrderViewModel) Implements IReport.SortOrders

		Public Property DefinitionAccessBasedOnSelectedCalculationColumns As String

		Public Overridable ReadOnly Property SortOrdersAvailable As Integer Implements IReport.SortOrdersAvailable
			Get
				If Columns IsNot Nothing Then
					Return Columns.Where(Function(m) m.IsExpression = False).Count - SortOrders.Count
				Else
					Return 0
				End If
			End Get
		End Property

		<MinLength(3, ErrorMessage:="You must select at least one column to order the definition by.")> _
		Public Overridable Property SortOrdersString As String

		Public Property JobsToHide As New Collection(Of Integer)

		Public Property SessionContext As SessionInfo Implements IDataAccess.SessionContext

		Public Property SessionInfo As SessionInfo Implements IReport.SessionInfo
		Public MustOverride Sub SetBaseTable(BaseTableID As Integer) Implements IReport.SetBaseTable

		Public Overridable Function GetAvailableSortColumns(Self As SortOrderViewModel) As IEnumerable(Of ReportColumnItem) Implements IReport.GetAvailableSortColumns

			Dim objItems As New Collection(Of ReportColumnItem)

			' Add all columns that aren't already included in the sort collection
			For Each objColumn In Columns.Where(Function(m) m.IsExpression = False)
				If SortOrders.Where(Function(m) m.ColumnID = objColumn.ID).Count = 0 Then
					objItems.Add(objColumn)
				End If
			Next

			' Add self to collection if not already there
			If Self.ColumnID > 0 And objItems.Where(Function(m) m.ID = Self.ColumnID).Count = 0 Then
				Dim objItem = Columns.Where(Function(m) m.ID = Self.ColumnID).FirstOrDefault
				objItems.Add(objItem)
			End If

			Return objItems.OrderBy(Function(m) m.Sequence)

		End Function

		Public Overridable Function GetAvailableTables() As IEnumerable(Of ReportTableItem) Implements IReport.GetAvailableTables

			Dim objItems As New Collection(Of ReportTableItem)
			Dim objBaseTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
			objItems.Add(New ReportTableItem With {.id = objBaseTable.ID, .Name = objBaseTable.Name, .Relation = ReportRelationType.Base})
			Return objItems

		End Function

		<DisplayName("Description 1 : ")>
		Public Property Description1ID As Integer

		<DisplayName("Description 2 : ")>
		Public Property Description2ID As Integer

		<DisplayName("Description 3 : ")>
		Public Property Description3ID As Integer

		<DisplayName("Region : ")>
		Public Property RegionID As Integer

		<DisplayName("Group by Description")>
		Public Property GroupByDescription As Boolean

		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<DisplayName("Separator : ")>
		Public Property Separator As String

		Public Sub Attach(ByRef session As SessionInfo)
			SessionInfo = session
		End Sub

		Public Property Dependencies As New ReportDependencies Implements IReport.Dependencies

		Public ReadOnly Property CanEditSecurityGroups As Boolean
			Get
				If SessionInfo IsNot Nothing Then
					Return SessionInfo.LoginInfo.IsSystemOrSecurityAdmin
				End If
			End Get
		End Property

	End Class
End Namespace