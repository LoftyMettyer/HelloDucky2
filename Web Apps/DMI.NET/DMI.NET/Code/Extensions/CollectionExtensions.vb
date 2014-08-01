Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports System.Web.Script.Serialization

Namespace Code.Extensions

	<HideModuleName>
	Public Module CollectionExtensions

		<Extension()>
		Public Function ToJsonResult(Of T As IJsonSerialize)(ByVal items As ICollection(Of T)) As MvcHtmlString

			Dim results = New With {.total = 1, .page = 1, .records = 1, .rows = items}
			Dim jsonSerialiser = New JavaScriptSerializer()
			Dim json = jsonSerialiser.Serialize(results)
			Return MvcHtmlString.Create(json)

		End Function

	End Module

End Namespace
