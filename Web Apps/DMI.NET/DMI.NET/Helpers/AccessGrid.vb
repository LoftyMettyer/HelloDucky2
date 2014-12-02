Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Web.Mvc
Imports DMI.NET.Classes

Namespace Helpers

	<HideModuleName()>
 Public Module MVCExtensions

		<Extension()> _
		Public Function AccessGrid(helper As HtmlHelper, name As String, items As IList(Of GroupAccess), htmlAttributes As Object) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Return MvcHtmlString.Create(BuildTable(name, items, HtmlHelper.AnonymousObjectToHtmlAttributes(htmlAttributes)).ToString)

		End Function

		Private Function BuildTable(name As String, items As IEnumerable(Of GroupAccess), attributes As IDictionary(Of String, Object)) As String
			Dim sb As New StringBuilder()

			sb.AppendLine("<tr><th>User Group</th><th>Access</th></tr>")

			' Add <apply to all groups> row.
			BuildTableRow(sb, New GroupAccess() With {.Access = "", .IsReadOnly = False, .Name = "(All Groups)"}, "drpSetAllSecurityGroups", 0)

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
			Dim sNiceText As String
			Const strSelected As String = "selected='selected'"
			Dim sName As String = String.Format("{0}[{1}]", name, rownumber)

			sb.AppendLine("<tr>")
			If obj.IsReadOnly Then
				sb.AppendFormat("<td><input type='hidden' style='width:100%'  name='{0}.Name' value='{1}' readonly='true'/><label class='ui-state-disabled'>{1}</label></td>", sName, obj.Name)
			Else
				sb.AppendFormat("<td><input type='hidden' style='width:100%'  name='{0}.Name' value='{1}' readonly='true'/><label>{1}</label></td>", sName, obj.Name)
			End If

			Select Case obj.Access.ToUpper
				Case "RW"
					iSelected = 0
					sNiceText = "Read / Write"
				Case "RO"
					iSelected = 1
					sNiceText = "Read Only"
				Case Else
					iSelected = 2
					sNiceText = "Hidden"
			End Select

			If obj.IsReadOnly Then
				sb.AppendFormat("<td><select style='padding: 0!important;width:95%' class='floatright ui-state-disabled' name='{0}.Access'><option value='{1}'>{2}</option></select></td>" _
												, sName, obj.Access.ToUpper, sNiceText)
			Else
				If name = "drpSetAllSecurityGroups" Then
					' cater for unselected option
					sb.AppendFormat("<td><select style='width:95%' onchange='setAllSecurityGroups();' id='drpSetAllSecurityGroups' class='floatright enableSaveButtonOnComboChange' name='drpSetAllSecurityGroups'><option selected value=''></option><option value='RW'>Read / Write</option><option value='RO'>Read Only</option><option value='HD'>Hidden</option></select></td>")
				Else
					sb.AppendFormat("<td><select style='width:95%'  class='reportViewAccessGroup floatright enableSaveButtonOnComboChange' name='{0}.Access'><option {1} value='RW'>Read / Write</option><option {2} value='RO'>Read Only</option><option {3} value='HD'>Hidden</option></select></td>" _
						, sName, IIf(iSelected = 0, strSelected, ""), IIf(iSelected = 1, strSelected, ""), IIf(iSelected = 2, strSelected, ""))

				End If
			End If

			sb.AppendLine("</tr>")

		End Sub

	End Module
End Namespace