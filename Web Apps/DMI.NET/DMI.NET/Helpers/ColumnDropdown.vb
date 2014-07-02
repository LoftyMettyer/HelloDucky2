Option Explicit On
Option Strict On

Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices

Namespace Helpers
	<HideModuleName> _
	Public Module ColumnDropdownExtension

		<Extension()> _
		Public Function ColumnDropdown(helper As HtmlHelper, name As String, bindValue As Integer, items As IList, onChangeEvent As String) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("onchange", onChangeEvent)

			For Each item In items
				Dim objType As Type = item.[GetType]()
				Dim iID As Integer = CInt(objType.GetProperty("id").GetValue(item, Nothing))

				content.AppendFormat(String.Format("<option value={0} data-datatype={4} data-size={2} data-decimals={3} {5}>{1}</option>" _
																, iID.ToString() _
																, objType.GetProperty("Name").GetValue(item, Nothing).ToString() _
																, objType.GetProperty("Size").GetValue(item, Nothing).ToString() _
																, objType.GetProperty("Decimals").GetValue(item, Nothing).ToString() _
																, CInt(objType.GetProperty("DataType").GetValue(item, Nothing)) _
																, IIf(bindValue = iID, "selected", "")))
			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function

	End Module

End Namespace