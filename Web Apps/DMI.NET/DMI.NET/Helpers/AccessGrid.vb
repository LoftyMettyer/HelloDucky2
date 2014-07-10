Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices
Imports DMI.NET.Classes

Public Module MVCExtensions

	<Extension()> _
	Public Function AccessGrid(helper As HtmlHelper, name As String, items As IList(Of GroupAccess), attributes As IDictionary(Of String, Object)) As String
		If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
			Return String.Empty
		End If

		Return BuildTable(name, items, attributes)
	End Function

	Private Function BuildTable(name As String, items As IList(Of GroupAccess), attributes As IDictionary(Of String, Object)) As String
		Dim sb As New StringBuilder()
		BuildTableHeader(sb)

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

	Private Sub BuildTableRow(sb As StringBuilder, obj As GroupAccess, name As String, rownumber As Integer)

		Dim iSelected As Integer
		Dim strSelected As String = "selected='selected'"
		Dim sName As String = String.Format("{0}[{1}]", name, rownumber)

		sb.AppendLine("<tr>")
		sb.AppendFormat("<td><input name='{0}.Name' value='{1}' readonly='true'/></td>", sName, obj.Name)

		Select Case obj.Access.ToUpper
			Case "RW"
				iSelected = 0
			Case "RO"
				iSelected = 1
			Case Else
				iSelected = 2
		End Select

		sb.AppendFormat("<td><select name='{0}.Access'><option {1} value='RW'>Read / Write</option><option {2} value='RO'>Read Only</option><option {3} value='HD'>Hidden</option></select></td>" _
				, sName, IIf(iSelected = 0, strSelected, ""), IIf(iSelected = 1, strSelected, ""), IIf(iSelected = 2, strSelected, ""))

		sb.AppendLine("</tr>")

	End Sub

	Private Sub BuildTableHeader(sb As StringBuilder)
		sb.AppendLine(vbTab & "<tr>")
		sb.AppendFormat("<th>Name</th><th>Access</th>")
		sb.AppendLine(vbTab & "</tr>")
	End Sub

End Module
