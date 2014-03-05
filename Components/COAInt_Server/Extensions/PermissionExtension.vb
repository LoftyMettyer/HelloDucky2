Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports System.Linq

Namespace Extensions

	<HideModuleName()> _
	Public Module PermissionExtension

		<Extension()>
		Friend Function GetByKey(Of T As Permission)(ByVal items As ICollection(Of T), ByVal key As String) As Boolean

			Dim objPermission = items.FirstOrDefault(Function(baseItem) (baseItem.Key = key))
			If objPermission Is Nothing Then
				Return False
			End If

			Return objPermission.IsPermitted
		End Function

		<Extension()>
		Friend Function IsPermitted(Of T As Permission)(ByVal items As ICollection(Of T), category As String, key As String) As Boolean

			Dim objPermission = items.FirstOrDefault(Function(baseItem) (baseItem.CategoryKey = category And baseItem.Key = key))
			If objPermission Is Nothing Then
				Return False
			End If

			Return objPermission.IsPermitted
		End Function




	End Module

End Namespace
