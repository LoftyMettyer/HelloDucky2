Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices
Imports DMI.NET.Classes

Namespace Helpers

	<HideModuleName>
	Public Module SelectedReportColumns

		<Extension()> _
		Public Function SelectedReportColumns(helper As HtmlHelper, name As String, items As IList(Of ReportColumnItem), attributes As IDictionary(Of String, Object)) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Dim sbPropertyBoxes As New StringBuilder()
			Dim sbColumnGrid As New StringBuilder()

			'Start table
			BuildTableHeader(sbColumnGrid)

			Dim iRow As Integer = 0
			For Each objItem In items

				' Append line line table
				BuildTableRow(sbColumnGrid, objItem, name, iRow)
				BuildPropertyGrid(sbPropertyBoxes, objItem, name, iRow)

				iRow += 1
			Next

			'builder.ToString(TagRenderMode.Normal) &
			'Dim builder As New TagBuilder("ul")
			'builder.InnerHtml = sbListbox.ToString()

			Dim tgColumnGrid As New TagBuilder("table")
			tgColumnGrid.MergeAttribute("id", "ColumnsSelected")
			tgColumnGrid.InnerHtml = sbColumnGrid.ToString

			Dim builder = New TagBuilder("div")
			builder.MergeAttributes(attributes)
			builder.MergeAttribute("name", name)
			builder.InnerHtml = tgColumnGrid.ToString & sbPropertyBoxes.ToString()
			Return MvcHtmlString.Create(builder.ToString)

		End Function

		Private Sub BuildPropertyGrid(tgPropertyGrid As StringBuilder, objItem As ReportColumnItem, name As String, rownumber As Integer)

			Dim builder = New TagBuilder("div")
			builder.MergeAttribute("id", String.Format("columnproperty{0}Breakdown", rownumber))
			Dim sb As New StringBuilder

			sb.AppendLine(String.Format("<input type='hidden' name='{0}[{1}].id' value='{2}' /></br/>", name, rownumber, objItem.id))
			sb.AppendLine(String.Format("<input type='hidden' name='{0}[{1}].DataType' value='{2}' /></br/>", name, rownumber, objItem.DataType))

			sb.AppendLine(String.Format("Heading: <input type='text' name='{0}[{1}].Heading' value='{2}' /></br/>", name, rownumber, objItem.Heading))
			sb.AppendLine(String.Format("Size: <input type='text' name='{0}[{1}].Size' value='{2}' /></br/>", name, rownumber, objItem.Size))
			sb.AppendLine(String.Format("Decimals: <input type='text' name='{0}[{1}].Decimals' value='{2}' /></br/>", name, rownumber, objItem.Decimals))

			sb.AppendLine(String.Format("Average: <input type='checkbox' name='{0}[{1}].IsAverage' {2} />", name, rownumber, IIf(objItem.IsAverage, "checked", "")))
			sb.AppendLine(String.Format("Count: <input type='checkbox' name='{0}[{1}].IsCount' {2} />", name, rownumber, IIf(objItem.IsCount, "checked", "")))
			sb.AppendLine(String.Format("Total: <input type='checkbox' name='{0}[{1}].IsTotal' {2} /></br/>", name, rownumber, IIf(objItem.IsTotal, "checked", "")))

			sb.AppendLine(String.Format("Hidden: <input type='checkbox' name='{0}[{1}].IsHidden' {2} />", name, rownumber, IIf(objItem.IsHidden, "checked", "")))
			sb.AppendLine(String.Format("Group With Next: <input type='checkbox' name='{0}[{1}].IsGroupWithNext' {2} /></br/></br>", name, rownumber, IIf(objItem.IsGroupWithNext, "checked", "")))

			builder.InnerHtml = sb.ToString
			tgPropertyGrid.Append(builder.ToString(TagRenderMode.Normal))

		End Sub


		Private Sub BuildTableRow(sb As StringBuilder, objItem As ReportColumnItem, name As String, rownumber As Integer)

			sb.AppendLine(vbTab & "<tr>")

			'Dim sName As String = String.Format("{0}[{1}].{2}", name, rownumber, [property].Name)
			'Dim sID As String = String.Format("{0}_{1}__{2}", name, rownumber, [property].Name)

			'	Select Case [property].Name.ToLower
			'	Case "columnname", "order", "columnid"
			sb.AppendFormat("<td name='{0}[{1}].id'>{2}</td>", name, rownumber, objItem.id)
			sb.AppendFormat("<td name='{0}[{1}].Name'>{2}</td>", name, rownumber, objItem.Name)

			'Case Else
			'	If [property].GetValue(obj, Nothing) = True Then
			'		sb.AppendFormat(vbTab & vbTab & "<td><input type='checkbox' checked name='{0}' id='{1}' /></td>" & vbLf, sName, sID)
			'	Else
			'		sb.AppendFormat(vbTab & vbTab & "<td><input type='checkbox' name='{0}' id='{1}' /></td>" & vbLf, sName, sID, [property].GetValue(obj, Nothing))
			'	End If

			'	End Select

			sb.AppendLine(vbTab & "</tr>")
		End Sub

		Private Sub BuildTableHeader(sb As StringBuilder)
			sb.AppendLine(vbTab & "<tr>")
			sb.Append("<th>id</th><th>Name</th>")
			sb.AppendLine(vbTab & "</tr>")
		End Sub

	End Module

End Namespace