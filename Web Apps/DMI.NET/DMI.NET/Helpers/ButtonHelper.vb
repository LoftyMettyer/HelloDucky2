Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Web.Mvc
Imports System.Runtime.CompilerServices

Namespace Helpers

	<HideModuleName>
	Public Module ButtonHelper

		<Extension()> _
		Public Function EllipseButton(helper As HtmlHelper, name As String, clickEvent As String, enabled As Boolean) As MvcHtmlString

			Dim builder = New TagBuilder("input")
			builder.MergeAttribute("name", name)
			builder.MergeAttribute("id", name)
			builder.MergeAttribute("type", "button")
			builder.MergeAttribute("value", "...")
			builder.MergeAttribute("onclick", clickEvent)

			If Not enabled Then
				builder.MergeAttribute("disabled", "disabled")
			End If

			Return MvcHtmlString.Create(builder.ToString)

		End Function

	End Module

End Namespace