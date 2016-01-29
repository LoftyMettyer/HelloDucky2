Option Explicit On
Option Strict On

Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports System.Collections.ObjectModel
Imports DMI.NET.Classes

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

		Public Property BaseSelection As RecordSelectionType
		Public Property BasePicklistID As Integer
		Public Property BaseFilterID As Integer

		<Range(1, Integer.MaxValue, ErrorMessage:="Role match table not selected.")>
		Public Property BaseChildTableID As Integer

		<Range(1, Integer.MaxValue, ErrorMessage:="Role match table column not selected.")>
		Public Property BaseChildColumnID As Integer

		Public Property BaseMinimumRatingColumnID As Integer
		Public Property BasePreferredRatingColumnID As Integer
		Public Property MatchTableID As Integer
		Public Property MatchSelectionType As RecordSelectionType
		Public Property MatchPicklistID As Integer
		Public Property MatchFilterID As Integer

		<Range(1, Integer.MaxValue, ErrorMessage:="Person match table not selected.")>
		Public Property MatchChildTableID As Integer

		<Range(1, Integer.MaxValue, ErrorMessage:="Person match table column not selected.")>
		Public Property MatchChildColumnID As Integer

		Public Property MatchChildRatingColumnID As Integer
		Public Property MatchAgainstType As MatchAgainstType
		Public Property Output As New ReportOutputModel
		Public Property MatchViewAccess As String
		<AllowHtml>
		Public Property MatchFilterName As String
		<AllowHtml>
		Public Property MatchPicklistName As String


		Public Overrides Function GetAvailableTables() As IEnumerable(Of ReportTableItem)

			Dim objItems As New Collection(Of ReportTableItem)

			' Add base table
			Dim objTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
			objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Base})

			Return objItems.OrderBy(Function(m) m.Name)

		End Function

	End Class
End Namespace