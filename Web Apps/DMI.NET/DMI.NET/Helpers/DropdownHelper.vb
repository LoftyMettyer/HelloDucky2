Option Explicit On
Option Strict On

Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices
Imports DMI.NET.Classes
Imports HR.Intranet.Server

Namespace Helpers
	<HideModuleName> _
	Public Module DropdownHelper

		<Extension()> _
		Public Function ColumnDropdown(helper As HtmlHelper, name As String, bindValue As Integer, items As List(Of ReportColumnItem), onChangeEvent As String) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("id", name)
			builder.MergeAttribute("onchange", onChangeEvent)

			For Each item In items
				Dim objType As Type = item.[GetType]()
				Dim iID As Integer = CInt(objType.GetProperty("ID").GetValue(item, Nothing))

				content.AppendFormat("<option value={0} data-datatype={4} data-size={2} data-decimals={3} {5}>{1}</option>" _
																, iID.ToString(), item.Name, item.Size.ToString, item.Decimals.ToString _
																, CInt(item.DataType), IIf(bindValue = iID, "selected", ""))

			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function

		<Extension()> _
		Public Function TableDropdown(helper As HtmlHelper, name As String, bindValue As Integer, items As IEnumerable(Of ReportTableItem), onChangeEvent As String) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("id", name)
			builder.MergeAttribute("onchange", onChangeEvent)

			For Each item In items
				Dim objType As Type = item.[GetType]()
				Dim iID As Integer = CInt(objType.GetProperty("id").GetValue(item, Nothing))

				content.AppendFormat("<option value={0} {2}>{1}</option>", iID.ToString(), item.Name, IIf(bindValue = iID, "selected", ""))

			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function


		<Extension()> _
		Public Function ColumnDropdown2(helper As HtmlHelper, name As String, bindValue As Integer) As MvcHtmlString

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("id", name)

			For Each item In objSessionInfo.Columns

				content.AppendFormat("<option value={0} data-datatype={4} data-size={2} data-decimals={3} {5}>{1}</option>" _
																, item.ID, item.Name, item.Size.ToString, item.Decimals.ToString _
																, CInt(item.DataType), IIf(bindValue = item.ID, "selected", ""))

			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function



	End Module

End Namespace