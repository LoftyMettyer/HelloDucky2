Option Explicit On
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

		Public Property Columns As New ReportColumnsModel

		Public Property AvailableTables As New Collection(Of ReportChildTables)

		Public Property ChildTables As New Collection(Of ReportChildTables)
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


	End Class

End Namespace