Option Explicit On
Option Strict On

Imports System.ComponentModel.DataAnnotations
Imports System.Collections.ObjectModel
Imports DMI.NET.Classes
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata


Namespace Models
   Public Class OrganisationReportModel
      Inherits ReportBaseModel

      Public Overrides ReadOnly Property ReportType As UtilityType
         Get
            Return UtilityType.OrgReporting
         End Get
      End Property

      Public Property BaseViewFilterID As Integer

      Public Property FiltersFieldList As New List(Of OrganisationReportFilterItem)

      'Property FiltersFieldList() As List(Of OrganisationReportFilterItem)


      <MinLength(3, ErrorMessage:="You must select at least one column for your report.")>
      Public Overrides Property ColumnsAsString As String

      Public Property FilterColoumnList As New Collection(Of SelectListItem)

      Public Property BaseViewList As New List(Of ReportTableItem)

      Public Overrides Sub SetBaseTable(TableID As Integer)
      End Sub

      Public Overrides Function GetAvailableTables() As IEnumerable(Of ReportTableItem)

         Dim objItems As New Collection(Of ReportTableItem)

         ' Add base table
         Dim objTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
         objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Base})

         Return objItems.OrderBy(Function(m) m.Name)

      End Function

      Friend Function GetAvailableTableViews(tableId As Integer) As List(Of ReportTableItem)

         Dim objView = SessionInfo.GetTableAssociatedViews(tableId)
         Dim objItems As New Collection(Of ReportTableItem)
         For Each item As View In objView
            objItems.Add(New ReportTableItem With {.id = item.ViewId, .Name = item.ViewName, .Relation = ReportRelationType.Base})
         Next

         Return objItems.OrderBy(Function(m) m.Name).ToList

      End Function

   End Class


End Namespace