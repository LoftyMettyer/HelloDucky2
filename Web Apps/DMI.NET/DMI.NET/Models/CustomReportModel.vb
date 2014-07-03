Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata
Imports DMI.NET.Classes

Namespace Models

	Public Class CustomReportModel
		Inherits ReportBaseModel

		Private _baseTable As Integer

		Public Property Columns As New ReportColumnsModel

		Public Property ChildTables As New Collection(Of ReportChildTables)

		Public Property Parent1 As New ReportRelatedTable
		Public Property Parent2 As New ReportRelatedTable

		Public Property Repetition As New Collection(Of ReportRepetition)

		Public Property IsSummary As Boolean
		Public Property IgnoreZerosForAggregates As Boolean

		Public Property Output As New ReportOutputModel

	End Class
End Namespace