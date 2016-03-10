Option Explicit On
Option Strict On

Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports DMI.NET.Classes
Imports DMI.NET.Code.Attributes

Namespace Models
	Public Class TalentReportModel
		Inherits ReportBaseModel

		Public Overrides ReadOnly Property ReportType As UtilityType
			Get
				Return UtilityType.TalentReport
			End Get
		End Property

		Public Overrides Sub SetBaseTable(TableID As Integer)
		End Sub

		<NonZeroIf("SelectionType", RecordSelectionType.Filter, ErrorMessage:="No filter selected for role table.")> _
		Public Overloads Property FilterID As Integer

		<NonZeroIf("SelectionType", RecordSelectionType.Picklist, ErrorMessage:="No picklist selected for role table.")>
		Public Overloads Property PicklistID As Integer

		Public Property BaseSelection As RecordSelectionType
		Public Property BasePicklistID As Integer
		Public Property BaseFilterID As Integer

        <Range(1, Integer.MaxValue, ErrorMessage:="Role match table not selected.")>
        Public Property BaseChildTableID As Integer

        <Range(1, Integer.MaxValue, ErrorMessage:="Role match table column not selected.")>
        Public Property BaseChildColumnID As Integer
        Public Property BaseChildColumnDataType As Integer

		Public Property BaseMinimumRatingColumnID As Integer
		Public Property BasePreferredRatingColumnID As Integer
		Public Property MatchTableID As Integer
		Public Property MatchSelectionType As RecordSelectionType

		<NonZeroIf("MatchSelectionType", RecordSelectionType.Picklist, ErrorMessage:="No picklist selected for person table.")> _
		Public Property MatchPicklistID As Integer

		<NonZeroIf("MatchSelectionType", RecordSelectionType.Filter, ErrorMessage:="No filter selected for person table.")> _
		Public Property MatchFilterID As Integer

		<Range(1, Integer.MaxValue, ErrorMessage:="Person match table not selected.")>
		Public Property MatchChildTableID As Integer

		<Range(1, Integer.MaxValue, ErrorMessage:="Person match table column not selected.")>
		Public Property MatchChildColumnID As Integer
		Public Property MatchChildColumnDataType As Integer

		Public Property MatchChildRatingColumnID As Integer
		Public Property MatchAgainstType As MatchAgainstType
		Public Property Output As New ReportOutputModel
		Public Property MatchViewAccess As String
		<AllowHtml>
		Public Property MatchFilterName As String
		<AllowHtml>
		Public Property MatchPicklistName As String

		<DisplayName("Include Unmatched Records")>
		Public Property IncludeUnmatched As Boolean

		<Range(0, 100, ErrorMessage:="Minimum Match Score not defined")>
		<DisplayName("Minimum Match Score : ")>
		Public Property MinimumScore As Integer

		<MinLength(0)>
		Public Overrides Property SortOrdersString As String

		<MinLength(3, ErrorMessage:="You must select at least one column for your report.")> _
		 Public Overrides Property ColumnsAsString As String

		Public Overrides Function GetAvailableTables() As IEnumerable(Of ReportTableItem)

			Dim objItems As New Collection(Of ReportTableItem)

			' Add base table
			Dim objTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
			objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Base})

			Return objItems.OrderBy(Function(m) m.Name)

		End Function

	End Class
End Namespace