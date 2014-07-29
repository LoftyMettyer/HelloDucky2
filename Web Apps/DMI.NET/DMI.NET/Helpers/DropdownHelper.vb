Option Explicit On
Option Strict On

Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices
Imports DMI.NET.Classes
Imports System.Linq.Expressions
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Namespace Helpers
	<HideModuleName> _
	Public Module DropdownHelper

		Private _objSessionInfo As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)

		<Extension()> _
		Public Function ColumnDropdown(helper As HtmlHelper, name As String, id As String, bindValue As Integer, items As IEnumerable(Of ReportColumnItem), onChangeEvent As String) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("id", id)
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
		Public Function TableDropdown(helper As HtmlHelper, name As String, id As String, bindValue As Integer, items As IEnumerable(Of ReportTableItem), onChangeEvent As String) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("id", id)
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
		Public Function LookupTableDropdown(helper As HtmlHelper, name As String, id As String, bindValue As Integer) As MvcHtmlString

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("id", id)

			For Each item In _objSessionInfo.Tables.Where(Function(m) m.TableType = TableTypes.tabLookup)
				content.AppendFormat("<option value={0} {2}>{1}</option>", item.ID, item.Name, IIf(bindValue = item.ID, "selected", ""))
			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function

		<Extension()> _
		Public Function ColumnDropdown2(helper As HtmlHelper, bindValue As Integer, TableID As Integer, DataType As SQLDataType, AddNone As Boolean, LimitToLookups As Boolean, htmlAttributes As Object) As MvcHtmlString

			Dim objAttributes = HtmlHelper.AnonymousObjectToHtmlAttributes(htmlAttributes)

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttributes(objAttributes)

			If AddNone Then
				content.AppendFormat("<option value=0 data-datatype={1} {0}>None</option>", IIf(bindValue = 0, "selected", ""), CInt(SQLDataType.sqlUnknown))
			End If

			Dim objColumns As IEnumerable(Of Column)
			If DataType = Nothing Then
				objColumns = _objSessionInfo.Columns.Where(Function(m) m.TableID = TableID AndAlso m.IsVisible).OrderBy(Function(m) m.Name)
			Else
				objColumns = _objSessionInfo.Columns.Where(Function(m) m.TableID = TableID AndAlso m.IsVisible AndAlso m.DataType = DataType).OrderBy(Function(m) m.Name)
			End If

			For Each item In objColumns

				content.AppendFormat("<option value={0} data-datatype={4} data-size={2} data-decimals={3} {5}>{1}</option>" _
																, item.ID, item.Name, item.Size.ToString, item.Decimals.ToString _
																, CInt(item.DataType), IIf(bindValue = item.ID, "selected", ""))

			Next


			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function

		<Extension()> _
		Public Function EmailGroupDropdown(helper As HtmlHelper, name As String, bindValue As Integer, items As IEnumerable(Of ReportTableItem)) As MvcHtmlString
			If items Is Nothing OrElse items.Count = 0 OrElse String.IsNullOrEmpty(name) Then
				Return MvcHtmlString.Empty
			End If

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)

			For Each item In items
				Dim iID As Integer = item.id

				content.AppendFormat("<option value={0} {2}>{1}</option>" _
																, iID.ToString(), item.Name, IIf(bindValue = iID, "selected", ""))

			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function

	End Module

End Namespace