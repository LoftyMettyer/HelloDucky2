Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports HR.Intranet.Server.Metadata
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports DMI.NET.ViewModels.Reports

Namespace Models

	Public Class CustomReportModel
		Inherits ReportBaseModel

		Public Overrides ReadOnly Property ReportType As UtilityType
			Get
				Return UtilityType.utlCustomReport
			End Get
		End Property

		<MinLength(3, ErrorMessage:="You must select at least one column for your report.")> _
		Public Overrides Property ColumnsAsString As String

		Public Property ChildTables As New List(Of ChildTableViewModel)

		<DisplayFormat(ConvertEmptyStringToNull:=False, NullDisplayText:="")>
		Public Property ChildTablesString As String
		Public Property ChildTablesAvailable As Boolean

		Public Property Parent1 As New ReportRelatedTable
		Public Property Parent2 As New ReportRelatedTable

		<DisplayName("Summary Report")>
		Public Property IsSummary As Boolean

		<DisplayName("Ignore zeros when calculating aggregates")>
		Public Property IgnoreZerosForAggregates As Boolean

		Public Property Output As New ReportOutputModel

		' Flags to detect if thius definition needs to be marked as hidden
		Public Property Parent1ViewAccess As String
		Public Property Parent2ViewAccess As String

		Public Overrides Sub SetBaseTable(TableID As Integer)

			ChildTables = New List(Of ChildTableViewModel)
			BaseTableID = TableID
			SelectionType = RecordSelectionType.AllRecords
			Columns = New List(Of ReportColumnItem)
			SortOrders = New Collection(Of SortOrderViewModel)

			Dim objParents = SessionInfo.Relations.Where(Function(m) m.ChildID = TableID)

			Parent1.ID = 0
			Parent1.Name = ""
			Parent1.SelectionType = RecordSelectionType.AllRecords
			Parent1.PicklistID = 0
			Parent1.PicklistName = ""
			Parent1.FilterID = 0
			Parent1.FilterName = ""

			Parent2.ID = 0
			Parent2.Name = ""
			Parent2.SelectionType = RecordSelectionType.AllRecords
			Parent2.PicklistID = 0
			Parent2.PicklistName = ""
			Parent2.FilterID = 0
			Parent2.FilterName = ""

			If objParents.Count > 0 Then
				With objParents(0)
					Parent1.ID = .ParentID
					Parent1.Name = SessionInfo.Tables.Where(Function(m) m.ID = .ParentID).FirstOrDefault.Name
				End With
			End If

			If objParents.Count > 1 Then
				With objParents(1)
					Parent2.ID = .ParentID
					Parent2.Name = SessionInfo.Tables.Where(Function(m) m.ID = .ParentID).FirstOrDefault.Name
				End With
			End If

		End Sub

		Public Overrides Function GetAvailableTables() As IEnumerable(Of ReportTableItem)

			Dim objItems As New Collection(Of ReportTableItem)
			Dim objTable As Table

			' Add base table
			objTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
			objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Base})

			' Add child tables
			For Each objChild In ChildTables
				objItems.Add(New ReportTableItem() With {.id = objChild.TableID, .Name = objChild.TableName, .Relation = ReportRelationType.Child})
			Next

			' Add parent tables
			If Parent1.ID > 0 Then
				objTable = SessionInfo.Tables.Where(Function(m) m.ID = Parent1.ID).FirstOrDefault
				objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Parent1})
			End If

			If Parent2.ID > 0 Then
				objTable = SessionInfo.Tables.Where(Function(m) m.ID = Parent2.ID).FirstOrDefault
				objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Parent2})
			End If

			Return objItems.OrderBy(Function(m) m.Name)

		End Function

	End Class

End Namespace