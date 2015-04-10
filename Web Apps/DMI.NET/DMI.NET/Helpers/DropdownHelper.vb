Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices
Imports DMI.NET.Classes
Imports System.Linq.Expressions
Imports HR.Intranet.Server
Imports HR.Intranet.Server.Metadata

Namespace Helpers
	<HideModuleName> _
	Public Module DropdownHelper

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

			For Each columnItem In items
				Dim objType As Type = columnItem.[GetType]()
				Dim iID As Integer = CInt(objType.GetProperty("ID").GetValue(columnItem, Nothing))

				content.AppendFormat("<option value={0} data-ishidden={6} data-datatype={4} data-size={2} data-decimals={3} {5}>{1}</option>" _
						, iID.ToString(), columnItem.Name, columnItem.Size.ToString, columnItem.Decimals.ToString, CInt(columnItem.DataType), IIf(bindValue = iID, "selected", ""), columnItem.IsHidden)

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

			For Each tableItem In items
				Dim objType As Type = tableItem.[GetType]()
				Dim iID As Integer = CInt(objType.GetProperty("id").GetValue(tableItem, Nothing))

				content.AppendFormat("<option value={0} {2}>{1}</option>", iID.ToString(), tableItem.Name, IIf(bindValue = iID, "selected", ""))

			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function

		<Extension()> _
		Public Function LookupTableDropdown(helper As HtmlHelper, name As String, id As String, bindValue As Integer, onChangeEvent As String, htmlAttributes As Object) As MvcHtmlString

			Dim objSessionInfo As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objAttributes = HtmlHelper.AnonymousObjectToHtmlAttributes(htmlAttributes)

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("id", id)
			builder.MergeAttribute("onchange", onChangeEvent)
			builder.MergeAttributes(objAttributes)

			For Each table In objSessionInfo.Tables.Where(Function(m) m.TableType = TableTypes.tabLookup).OrderBy(Function(m) m.Name)
				content.AppendFormat("<option value={0} {2}>{1}</option>", table.ID, table.Name, IIf(bindValue = table.ID, "selected", ""))
			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function

		<Extension> _
		Public Function ColumnDropdownFor(Of TModel, TValue)(html As HtmlHelper(Of TModel), expression As Expression(Of Func(Of TModel, TValue)), filter As ColumnFilter, htmlAttributes As Object) As MvcHtmlString

			Dim objSessionInfo As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)

			Dim htmlFieldName = ExpressionHelper.GetExpressionText(expression)
			Dim fullHtmlFieldName = TagBuilder.CreateSanitizedId(html.ViewContext.ViewData.TemplateInfo.GetFullHtmlFieldName(htmlFieldName))
			Dim bindValue = CInt(ModelMetadata.FromLambdaExpression(expression, html.ViewData).Model)
			Dim objAttributes = HtmlHelper.AnonymousObjectToHtmlAttributes(htmlAttributes)
			Dim objColumns As IEnumerable(Of Column)

			Dim content As New StringBuilder
			Dim builder As New TagBuilder("select")
			builder.MergeAttribute("name", fullHtmlFieldName)
			builder.MergeAttribute("id", fullHtmlFieldName)
			builder.MergeAttributes(objAttributes)

			'In calendar report region table should display 'Default' for base table
			If filter.AddNone Then
				content.AppendFormat("<option value=0 data-datatype={1} data-decimals=0 data-lookuptableID=0 {0}>None</option>", IIf(bindValue = 0, "selected", ""), CInt(ColumnDataType.sqlUnknown))
			ElseIf filter.AddDefault Then
				content.AppendFormat("<option value=0 data-datatype={1} data-decimals=0 data-lookuptableID=0 {0}>Default</option>", IIf(bindValue = 0, "selected", ""), CInt(ColumnDataType.sqlUnknown))
			End If

			Dim iParent1 As Integer
			Dim iParent2 As Integer

			If filter.IncludeParents Then
				Dim objParent1 = objSessionInfo.Relations.FirstOrDefault(Function(m) m.ChildID = filter.TableID)
				If objParent1 IsNot Nothing Then
					iParent1 = objParent1.ParentID
				End If

				Dim objParent2 = objSessionInfo.Relations.LastOrDefault(Function(m) m.ChildID = filter.TableID)
				If objParent2 IsNot Nothing Then
					iParent2 = objParent2.ParentID
				End If

			End If

			If filter.IsNumeric Then
				objColumns = objSessionInfo.Columns.Where(Function(m) (m.TableID = filter.TableID OrElse m.TableID = iParent1 OrElse m.TableID = iParent2) AndAlso
																		m.IsVisible AndAlso m.IsNumeric
																		).OrderBy(Function(m) m.TableID).ThenBy(Function(m) m.Name)
			Else
				objColumns = objSessionInfo.Columns.Where(Function(m) (m.TableID = filter.TableID OrElse m.TableID = iParent1 OrElse m.TableID = iParent2) AndAlso
																		m.IsVisible AndAlso
																		(m.Size = filter.Size OrElse filter.Size = 0) AndAlso
																		(m.DataType = filter.DataType OrElse filter.DataType = ColumnDataType.sqlUnknown) AndAlso
																		(m.ColumnType = ColumnType.Lookup OrElse filter.ColumnType = ColumnType.Unknown)
																		).OrderBy(Function(m) m.TableID).ThenBy(Function(m) m.Name)
			End If

			If filter.ExcludeOleAndPhoto Then
				For Each column In objColumns
					Dim skipMe As Boolean = (column.DataType = ColumnDataType.sqlOle OrElse column.DataType = ColumnDataType.sqlVarBinary)
					If skipMe = False Then
						content.AppendFormat("<option value={0} data-datatype={4} data-size={2} data-decimals={3} data-lookuptableID={6} {5}>{1}</option>" _
																				 , column.ID _
																				 , IIf(filter.ShowFullName, column.TableName & "." & column.Name, column.Name) _
																				 , column.Size.ToString, column.Decimals.ToString _
																				 , CInt(column.DataType), IIf(bindValue = column.ID, "selected", "") _
																				 , column.LookupTableID)
					End If
				Next
			Else
				For Each column In objColumns

					content.AppendFormat("<option value={0} data-datatype={4} data-size={2} data-decimals={3} data-lookuptableID={6} {5}>{1}</option>" _
																			 , column.ID _
																			 , IIf(filter.ShowFullName, column.TableName & "." & column.Name, column.Name) _
																			 , column.Size.ToString, column.Decimals.ToString _
																			 , CInt(column.DataType), IIf(bindValue = column.ID, "selected", "") _
																			 , column.LookupTableID)
				Next
			End If

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

			For Each tableItem In items
				Dim iID As Integer = tableItem.id

				content.AppendFormat("<option value={0} {2}>{1}</option>" _
																, iID.ToString(), tableItem.Name, IIf(bindValue = iID, "selected", ""))

			Next

			builder.InnerHtml = content.ToString
			Return MvcHtmlString.Create(builder.ToString())

		End Function

	End Module

End Namespace