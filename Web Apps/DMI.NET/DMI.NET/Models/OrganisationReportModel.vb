Option Explicit On
Option Strict On

Imports HR.Intranet.Server.Enums
Imports System.ComponentModel.DataAnnotations
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports DMI.NET.Classes
Imports DMI.NET.Code.Attributes
Imports HR.Intranet.Server.Structures

Namespace Models
   Public Class OrganisationReportModel
      Inherits ReportBaseModel

      Public Overrides ReadOnly Property ReportType As UtilityType
         Get
            Return UtilityType.OrgReporting
         End Get
      End Property

      Public Property BaseViewFilterID As Integer

      Public Property Filters As New List(Of OrganisationReportFilterItem)

      <MinLength(3, ErrorMessage:="You must select at least one column for your report.")>
      Public Overrides Property ColumnsAsString As String

      Public Overrides Sub SetBaseTable(TableID As Integer)
      End Sub

      Public Overrides Function GetAvailableTables() As IEnumerable(Of ReportTableItem)

         Dim objItems As New Collection(Of ReportTableItem)

         ' Add base table
         Dim objTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
         objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Base})

         Return objItems.OrderBy(Function(m) m.Name)

      End Function

      Public Function GetAvailableTableViews() As IEnumerable(Of ReportTableItem)

         'Dim objMenu As HR.Intranet.Server.Menu
         Dim objItems As New Collection(Of ReportTableItem)

         Dim objMenu = New HR.Intranet.Server.Menu()
         Dim avPrimaryMenuInfo As List(Of MenuInfo)
         objMenu.SessionInfo = SessionContext
         avPrimaryMenuInfo = objMenu.GetPrimaryTableMenu

         ' Add base table
         Dim objView = objMenu.GetPrimaryTableMenu.Where(Function(m) m.TableID = BaseTableID).FirstOrDefault
         objItems.Add(New ReportTableItem With {.id = objView.ViewID, .Name = objView.ViewName, .Relation = ReportRelationType.Base})

         Return objItems.OrderBy(Function(m) m.Name)

      End Function



   End Class


End Namespace