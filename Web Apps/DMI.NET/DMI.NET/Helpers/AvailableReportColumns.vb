Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices
Imports DMI.NET.Classes

Namespace Helpers
	<HideModuleName> _
	Public Module AvailableReportColumns

		<Extension()> _
	 Public Function AvailableReportColumns(helper As HtmlHelper, name As String, items As IList(Of ReportColumnItem), attributes As IDictionary(Of String, Object)) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Dim sb As New StringBuilder
			Dim builder As New TagBuilder("table")
			builder.MergeAttribute("id", name)

			' Table header
			sb.Append("<tr><th>id</th><th>Name</th></tr>")

			For Each objItem In items
				sb.AppendFormat("<tr><td data-datatype={4} data-size={2} data-decimals={3}>{0}</td><td>{1}</td></tr>" _
								, objItem.id, objItem.Name, objItem.Size, objItem.Decimals, objItem.DataType)

			Next

			builder.InnerHtml = sb.ToString
			Return MvcHtmlString.Create(builder.ToString)

		End Function


	End Module

End Namespace
