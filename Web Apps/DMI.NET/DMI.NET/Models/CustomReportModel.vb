﻿Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata
Imports DMI.NET.Classes
Imports System.Runtime.CompilerServices
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

		Private _baseTable As Integer

		Public Property Columns As New ColumnsViewModel

		Public Property AvailableTables As New Collection(Of ChildTableViewModel)

		Public Property ChildTables As New Collection(Of ChildTableViewModel)
		Public Property ChildTablesString As String

		Public Property Parent1 As New ReportRelatedTable
		Public Property Parent2 As New ReportRelatedTable

		Public Property IsSummary As Boolean
		Public Property IgnoreZerosForAggregates As Boolean

		Public Property Output As New ReportOutputModel

		' Flags to detect if thius definition needs to be marked as hidden
		Public Property p1Hidden As Boolean
		Public Property p2Hidden As Boolean
		Public Property childHidden As Boolean

		Public Overrides Sub SetBaseTable(TableID As Integer)

			ChildTables = New Collection(Of ChildTableViewModel)
			BaseTableID = TableID
			Columns.DisplayTableSelection = True
			SelectionType = Enums.RecordSelectionType.AllRecords
			Columns.Selected = New Collection(Of ReportColumnItem)
			SortOrders = New Collection(Of SortOrderViewModel)
			Repetition = New Collection(Of ReportRepetition)

			Dim objParents = SessionInfo.Relations.Where(Function(m) m.ChildID = TableID)

			Parent1.ID = 0
			Parent1.Name = ""
			Parent1.SelectionType = Enums.RecordSelectionType.AllRecords
			Parent1.PicklistID = 0
			Parent1.PicklistName = ""
			Parent1.FilterID = 0
			Parent1.FilterName = ""

			Parent2.ID = 0
			Parent2.Name = ""
			Parent2.SelectionType = Enums.RecordSelectionType.AllRecords
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

		Public Overrides Function GetAvailableSortColumns() As IEnumerable(Of ReportColumnItem)

			Dim objItems As New Collection(Of ReportColumnItem)

			For Each objColumn In Columns.Selected
				objItems.Add(objColumn)
			Next

			Return objItems

		End Function

	End Class

End Namespace