Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices

Namespace Helpers
  <HideModuleName> _
  Public Module SortOrderGridExtension

    <Extension()> _
    Public Function SortOrderGrid(helper As HtmlHelper, name As String, items As IList, attributes As IDictionary(Of String, Object)) As MvcHtmlString
      If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
        Return MvcHtmlString.Empty
      End If

      Return MvcHtmlString.Create(BuildTable(name, items, attributes))
    End Function

    Private Function BuildTable(name As String, items As IList, attributes As IDictionary(Of String, Object)) As String

      Dim sb As New StringBuilder()
      Dim iRow As Integer = 0
      BuildTableHeader(sb, items(0).[GetType]())

      For Each item In items
        BuildTableRow(sb, item, name, iRow)
        iRow += 1
      Next

      Dim builder As New TagBuilder("table")
      builder.MergeAttributes(attributes)
      builder.MergeAttribute("id", name)
      builder.MergeAttribute("name", name)
      builder.InnerHtml = sb.ToString()
      Return builder.ToString(TagRenderMode.Normal)
    End Function

    Private Sub BuildTableRow(sb As StringBuilder, obj As Object, name As String, rownumber As Integer)
      Dim objType As Type = obj.[GetType]()
      sb.AppendLine(vbTab & "<tr>")

      For Each [property] In objType.GetProperties()

        Dim sName As String = String.Format("{0}[{1}].{2}", name, rownumber, [property].Name)
        Dim sID As String = String.Format("{0}_{1}__{2}", name, rownumber, [property].Name)

        Select Case [property].Name.ToLower
          Case "columnname", "order", "columnid"
            sb.AppendFormat(vbTab & vbTab & "<td><input name='{0}' id='{1}' value='{2}'/></td>" & vbLf, sName, sID, [property].GetValue(obj, Nothing))

          Case Else
            If [property].GetValue(obj, Nothing) = True Then
              sb.AppendFormat(vbTab & vbTab & "<td><input type='checkbox' checked name='{0}' id='{1}' /></td>" & vbLf, sName, sID)
            Else
              sb.AppendFormat(vbTab & vbTab & "<td><input type='checkbox' name='{0}' id='{1}' /></td>" & vbLf, sName, sID, [property].GetValue(obj, Nothing))
            End If

        End Select

      Next
      sb.AppendLine(vbTab & "</tr>")
    End Sub

    Private Sub BuildTableHeader(sb As StringBuilder, p As Type)
      sb.AppendLine(vbTab & "<tr>")
      For Each [property] In p.GetProperties()
        sb.AppendFormat(vbTab & vbTab & "<th>{0}</th>" & vbLf, [property].Name)
      Next
      sb.AppendLine(vbTab & "</tr>")
    End Sub

  End Module

End Namespace