Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports System.Web.Script.Serialization
Imports DMI.NET.Classes

Namespace Code.Extensions

	<HideModuleName>
	Public Module CollectionExtensions

		<Extension()>
		Public Function ToJsonResult(Of T As IJsonSerialize)(items As ICollection(Of T)) As MvcHtmlString

			Dim results = New With {.total = 1, .page = 1, .records = 1, .rows = items}
			Dim jsonSerialiser = New JavaScriptSerializer()
			Dim json = HttpUtility.JavaScriptStringEncode(jsonSerialiser.Serialize(results))
			Return MvcHtmlString.Create(json)

		End Function

		<Extension()>
		Public Function HiddenGroups(Of T As GroupAccess)(items As ICollection(Of T)) As String

			Dim aryGroups As New ArrayList

			For Each objGroup In items
				If objGroup.Access = "HD" Then
					aryGroups.Add(objGroup.Name)
				End If
			Next

			If aryGroups.Count = 0 Then
				Return ""
			Else
				Return vbTab + String.Join(vbTab, aryGroups.ToArray()) + vbTab
			End If

		End Function


	End Module

End Namespace
