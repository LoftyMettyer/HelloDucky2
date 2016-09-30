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

      <MinLength(3, ErrorMessage:="You must select at least one column for your report.")>
      Public Overrides Property ColumnsAsString As String

      Public Property FilterColumnsAsString As String

      Public Property BaseViewList As New List(Of ReportTableItem)

      Public Property PreviewColumnList As New List(Of ReportColumnItem)

      Public Property InvalidColumnList As New List(Of ReportColumnItem)

      Public Property AllAvailableViewList As New List(Of ReportTableItem)

      Public Property PostBasedTableName As String

      Public Property PostBasedTableId As Integer

      Public Property IsBaseViewAccessDenied As Boolean

      Public Property SelectViewOnColumnsTab As Integer

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

      Sub RemoveInvalidColumns()

         InvalidColumnList.Clear()

         For Each column As ReportColumnItem In Columns
            Dim delimiter As Char = "."c
            Dim substrings() As String = column.Name.Split(delimiter)
            Dim isValidColumn = SessionInfo.ValidateColumnPermissions(substrings(0), column.Heading)
            If (isValidColumn = False) Then
               InvalidColumnList.Add(column)
            End If
         Next

         Dim baseViewName = String.Empty
         If BaseViewList.Exists(Function(x) x.id = BaseViewID) Then
            baseViewName = BaseViewList.Find(Function(x) x.id = BaseViewID).Name
         End If

         If baseViewName <> String.Empty Then
            For Each column As OrganisationReportFilterItem In FiltersFieldList
               Dim delimiter As Char = "."c
               Dim substrings() As String = column.FieldName.Split(delimiter)
               Dim isValidColumn = SessionInfo.ValidateColumnPermissions(baseViewName, column.FieldName)
               If (isValidColumn = False) Then
                  If (InvalidColumnList.Exists(Function(x) x.Heading = column.FieldName AndAlso x.ID = column.FieldID) = False) Then
                     Dim reportColumn As New ReportColumnItem
                     reportColumn.Heading = column.FieldName
                     reportColumn.ID = column.FieldID
                     reportColumn.ColumnValue = column.FilterValue
                     InvalidColumnList.Add(reportColumn)
                  End If
               End If
            Next
         End If

         If InvalidColumnList.Count > 0 Then

            'Remove from selected column and filters for that column
            For Each column As ReportColumnItem In InvalidColumnList
               Columns.Remove(column)
               FiltersFieldList.RemoveAll(Function(f) f.FieldID = column.ID AndAlso f.FieldName = column.Heading)
            Next

            'Remove group with next for all columns
            For Each column As ReportColumnItem In Columns
               column.IsGroupWithNext = False
            Next
         End If

      End Sub

      Friend Function ProcessColumnsForPreview(Columns As List(Of ReportColumnItem)) As List(Of ReportColumnItem)
         Dim space As String = " "
         Dim recordCount As Integer
         Dim openBracket As String = "<" + space
         Dim closeBracket As String = space + ">"
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

         While (recordCount < BaseViewColumns.Count)

            Dim item = BaseViewColumns(recordCount)
            item.Heading = (openBracket + item.Prefix + item.Heading + item.Suffix).Trim

            If item.IsGroupWithNext Then

               'Set row height
               item.DefaultHeight = item.Height
               item.Height = Convert.ToInt32(item.Height * Math.Round(item.FontSize * 1.5))

               While (recordCount < BaseViewColumns.Count)
                  recordCount = recordCount + 1
                  Dim nextItem = BaseViewColumns(recordCount)

                  'Set group name with next
                  item.Heading = item.Heading + space + nextItem.Prefix + nextItem.Heading + nextItem.Suffix

                  If nextItem.IsGroupWithNext = False Then
                     Exit While
                  End If
               End While

               item.Heading = (item.Heading + closeBracket).Trim
               recordCount = recordCount + 1

            Else
               'Set row height
               item.DefaultHeight = item.Height
               item.Height = Convert.ToInt32(item.Height * Math.Round(item.FontSize * 1.5))
               item.Heading = (item.Heading + closeBracket).Trim
               recordCount = recordCount + 1

            End If

            PreviewColumnList.Add(item)

         End While

         Return PreviewColumnList
      End Function

   End Class


End Namespace