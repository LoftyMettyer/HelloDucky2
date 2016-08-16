Option Explicit On
Option Strict On

Imports System.ComponentModel.DataAnnotations
Imports System.Collections.ObjectModel
Imports DMI.NET.Classes
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata
Imports System.Web.Script.Serialization

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
      <MinLength(3, ErrorMessage:="You must select at least one column for your report.")>
      Public Overrides Property ColumnsAsString As String
      Public Property FilterColumnsAsString As String
      Public Property BaseViewList As New List(Of ReportTableItem)

      Public Property PreviewColumnList As New List(Of ReportColumnItem)

      Public Property AllAvailableViewList As New List(Of ReportTableItem)

      Public Property PostBasedTableName As String

      Public Property PostBasedTableId As Integer

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

         Friend Function ProcessColumnsForPreview(Columns As List(Of ReportColumnItem)) As List(Of ReportColumnItem)
         Dim space As String = " "
         Dim count As Integer
         Dim openBracket As String = "< "
         Dim closeBracket As String = " >"
         Dim ignoreNextItem As Boolean = False
         PreviewColumnList.Clear()
         Dim BaseViewColumns As New List(Of ReportColumnItem)

         'Sort the data
         If PostBasedTableId > 0 Then
            BaseViewColumns = Columns.FindAll(Function(x) x.ViewID = BaseViewID)
            For Each item As ReportColumnItem In BaseViewColumns
               Columns.RemoveAll(Function(m) m.ID = item.ID)
            Next
            BaseViewColumns.AddRange(Columns)
         Else
            BaseViewColumns = Columns
         End If

         While (count < BaseViewColumns.Count)

            Dim item = BaseViewColumns(count)
            item.Heading = (openBracket + item.Prefix + space + item.Heading).Trim

            If item.IsGroupWithNext Then

               'set row height
               item.DefaultHeight = item.Height
               item.Height = Convert.ToInt32(item.Height * Math.Round(item.FontSize * 1.5))
               Dim suffix = String.Empty
               While (count < BaseViewColumns.Count)
                  count = count + 1
                  Dim nextItem = BaseViewColumns(count)

                  'set group name with next
                  item.Heading = item.Heading + space + nextItem.Heading
                  suffix = nextItem.Suffix

                  If nextItem.IsGroupWithNext = False Then
                     Exit While
                  End If
               End While

               item.Heading = (item.Heading + suffix + closeBracket).Trim

            Else
               'set row height
               item.DefaultHeight = item.Height
               item.Height = Convert.ToInt32(item.Height * Math.Round(item.FontSize * 1.5))
               item.Heading = (item.Heading + space + item.Suffix + closeBracket).Trim
               count = count + 1

            End If

            PreviewColumnList.Add(item)

         End While

         Return PreviewColumnList
      End Function

   End Class


End Namespace