Option Explicit On
Option Strict On

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

	  Public Property BaseSelection As Integer
	  Public Property BasePicklistID As Integer
	  Public Property BaseFilterID As Integer
	  Public Property BaseChildTableID As Integer
	  Public Property BaseChildColumnID As Integer
	  Public Property BaseMinimumRatingColumnID As Integer
	  Public Property BasePreferredRatingColumnID As Integer
    Public Property MatchTableID As Integer
	  Public Property MatchSelection  As Integer
	  Public Property MatchPicklistID  As Integer
	  Public Property MatchFilterID  As Integer
	  Public Property MatchChildTableID As Integer
	  Public Property MatchChildColumnID As Integer
	  Public Property MatchChildRatingColumnID  As Integer
	  Public Property MatchAgainstType  As Integer
 		Public Property Output As New ReportOutputModel

    Public Overrides Function GetAvailableTables() As IEnumerable(Of ReportTableItem)

			Dim objItems As New Collection(Of ReportTableItem)

			Dim objBaseTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
			objItems.Add(New ReportTableItem With {.id = objBaseTable.ID, .Name = objBaseTable.Name, .Relation = ReportRelationType.Base})

			' Add base table
			Dim objTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
			objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Base})

      ' Add match table
      objTable = SessionInfo.Tables.Where(Function(m) m.ID = MatchTableID).FirstOrDefault
			objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Base})
      
			Return objItems.OrderBy(Function(m) m.Name)

		End Function

  End Class
End Namespace