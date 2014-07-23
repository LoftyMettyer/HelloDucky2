Imports System.Runtime.CompilerServices
Imports System.Linq.Expressions

Namespace Helpers

	Public Module LabelExtensions

		<Extension> _
		Public Function LabelFor(Of TModel, TValue)(html As HtmlHelper(Of TModel), expression As Expression(Of Func(Of TModel, TValue)), htmlAttributes As Object) As MvcHtmlString
			Return html.LabelFor(expression, Nothing, htmlAttributes)
		End Function

		<Extension> _
		Public Function LabelFor(Of TModel, TValue)(html As HtmlHelper(Of TModel), expression As Expression(Of Func(Of TModel, TValue)), labelText As String, htmlAttributes As Object) As MvcHtmlString
			Return html.LabelHelper(ModelMetadata.FromLambdaExpression(expression, html.ViewData), ExpressionHelper.GetExpressionText(expression), HtmlHelper.AnonymousObjectToHtmlAttributes(htmlAttributes), labelText)
		End Function

		<Extension> _
		Private Function LabelHelper(html As HtmlHelper, metadata As ModelMetadata, htmlFieldName As String, htmlAttributes As IDictionary(Of String, Object), Optional labelText As String = Nothing) As MvcHtmlString

			Dim str = If(labelText, (If(metadata.DisplayName, (If(metadata.PropertyName, htmlFieldName.Split(New Char() {"."c}).Last())))))

			If String.IsNullOrEmpty(str) Then
				Return MvcHtmlString.Empty
			End If

			Dim tagBuilder__1 = New TagBuilder("label")
			tagBuilder__1.MergeAttributes(htmlAttributes)
			tagBuilder__1.Attributes.Add("for", TagBuilder.CreateSanitizedId(html.ViewContext.ViewData.TemplateInfo.GetFullHtmlFieldName(htmlFieldName)))
			tagBuilder__1.SetInnerText(str)

			Return tagBuilder__1.ToMvcHtmlString(TagRenderMode.Normal)
		End Function

		<Extension> _
		Private Function ToMvcHtmlString(tagBuilder As TagBuilder, renderMode As TagRenderMode) As MvcHtmlString
			Return New MvcHtmlString(tagBuilder.ToString(renderMode))
		End Function

	End Module

End Namespace