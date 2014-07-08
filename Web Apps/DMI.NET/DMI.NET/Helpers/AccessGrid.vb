Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices

'Namespace Helpersdm
Public Module MVCExtensions

  <Extension()> _
  Public Function AccessGrid(helper As HtmlHelper, name As String, items As IList, attributes As IDictionary(Of String, Object)) As String
    If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
      Return String.Empty
    End If

    Return BuildTable(name, items, attributes)
  End Function

  Private Function BuildTable(name As String, items As IList, attributes As IDictionary(Of String, Object)) As String
    Dim sb As New StringBuilder()
    BuildTableHeader(sb, items(0).[GetType]())

    Dim iRow As Integer = 0
    For Each item In items
      BuildTableRow(sb, item, name, iRow)
      iRow += 1
    Next

    Dim builder As New TagBuilder("table")
    builder.MergeAttributes(attributes)
		builder.MergeAttribute("name", name)
    builder.InnerHtml = sb.ToString()
    Return builder.ToString(TagRenderMode.Normal)
  End Function

  Private Sub BuildTableRow(sb As StringBuilder, obj As Object, name As String, rownumber As Integer)
    Dim objType As Type = obj.[GetType]()
    sb.AppendLine(vbTab & "<tr>")
    For Each [property] In objType.GetProperties()
      '      sb.AppendFormat(vbTab & vbTab & "<td>{0}</td>" & vbLf, [property].GetValue(obj, Nothing))

      Dim sName As String = String.Format("{0}[{1}].{2}", name, rownumber, [property].Name)
      Dim sID As String = String.Format("{0}_{1}__{2}", name, rownumber, [property].Name)
      Dim iSelected As Integer

      Select Case [property].Name.ToLower
        Case "access"

          Select Case [property].GetValue(obj, Nothing).ToString.ToUpper
            Case "RW"
              iSelected = 0
            Case "RO"
              iSelected = 1
            Case Else
              iSelected = 2

          End Select

          Dim strSelected As String = "selected='selected'"

          Dim dropDown = String.Format("<td><select name='{0}'><option {1} value='RW'>Read / Write</option><option {2} value='RO'>Read Only</option><option {3} value='HD'>Hidden</option></select></td>" _
                  , sName, IIf(iSelected = 0, strSelected, ""), IIf(iSelected = 1, strSelected, ""), IIf(iSelected = 2, strSelected, ""))

          sb.AppendFormat(vbTab & vbTab & dropDown & vbLf)

        Case Else
					sb.AppendFormat(vbTab & vbTab & "<td><input name='{0}' id='{1}' value='{2}' readonly='true'/></td>" & vbLf, sName, sID, [property].GetValue(obj, Nothing))

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

'End Namespace