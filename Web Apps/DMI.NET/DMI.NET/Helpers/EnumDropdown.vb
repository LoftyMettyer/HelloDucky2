Imports System.Linq.Expressions
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports System.ComponentModel

Namespace Helpers
	<HideModuleName> _
	Public Module EnumDropdownExtensions

		Public Function GetInputName(Of TModel, TProperty)(expression As Expression(Of Func(Of TModel, TProperty))) As String
			If expression.Body.NodeType = ExpressionType.[Call] Then
				Dim methodCallExpression As MethodCallExpression = DirectCast(expression.Body, MethodCallExpression)
				Dim name As String = GetInputName(methodCallExpression)

				Return name.Substring(expression.Parameters(0).Name.Length + 1)
			End If
			Return expression.Body.ToString().Substring(expression.Parameters(0).Name.Length + 1)
		End Function

		Private Function GetInputName(expression As MethodCallExpression) As String
			' p => p.Foo.Bar().Baz.ToString() => p.Foo OR throw...
			Dim methodCallExpression As MethodCallExpression = TryCast(expression.[Object], MethodCallExpression)
			If methodCallExpression IsNot Nothing Then
				Return GetInputName(methodCallExpression)
			End If
			Return expression.[Object].ToString()
		End Function

		<Extension> _
		Public Function EnumDropDownListFor(Of TModel As Class, TProperty)(htmlHelper As HtmlHelper(Of TModel), expression As Expression(Of Func(Of TModel, TProperty))) As MvcHtmlString
			Dim inputName As String = GetInputName(expression)
			Dim value = If(htmlHelper.ViewData.Model Is Nothing, Nothing, expression.Compile()(htmlHelper.ViewData.Model))

			Return htmlHelper.DropDownList(inputName, ToSelectList(GetType(TProperty), value.ToString()))
		End Function

		Public Function ToSelectList(enumType As Type, selectedItem As String) As SelectList
			Dim items As New List(Of SelectListItem)()
			For Each item In [Enum].GetValues(enumType)
				Dim fi As FieldInfo = enumType.GetField(item.ToString())
				Dim attribute = fi.GetCustomAttributes(GetType(DescriptionAttribute), True).FirstOrDefault()
				Dim title = If(attribute Is Nothing, item.ToString(), DirectCast(attribute, DescriptionAttribute).Description)
				Dim listItem = New SelectListItem() With { _
					 .Value = CInt(item).ToString(), _
					 .Text = title, _
					 .Selected = selectedItem = CInt(item).ToString() _
				}
				items.Add(listItem)
			Next

			Return New SelectList(items, "Value", "Text", selectedItem)
		End Function

		<Extension> _
	 Public Function CustomEnumDropDownListFor(Of TModel, TEnum)(htmlHelper__1 As HtmlHelper(Of TModel), expression As Expression(Of Func(Of TModel, TEnum)), htmlAttributes As Object) As MvcHtmlString
			Dim metadata = ModelMetadata.FromLambdaExpression(expression, htmlHelper__1.ViewData)
			Dim values = [Enum].GetValues(GetType(TEnum)).Cast(Of TEnum)()

			Dim items = values.[Select](Function(value) New SelectListItem() With { _
				.Text = GetEnumDescription(value), _
				.Value = value.ToString(), _
				.Selected = value.Equals(metadata.Model)})

			Dim attributes = HtmlHelper.AnonymousObjectToHtmlAttributes(htmlAttributes)
			Return htmlHelper__1.DropDownListFor(expression, items, attributes)
		End Function

		Public Function GetEnumDescription(Of TEnum)(value As TEnum) As String
			Dim field = value.[GetType]().GetField(value.ToString())
			Dim attributes = DirectCast(field.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
			Return If(attributes.Length > 0, attributes(0).Description, value.ToString())
		End Function
		
	End Module

End Namespace