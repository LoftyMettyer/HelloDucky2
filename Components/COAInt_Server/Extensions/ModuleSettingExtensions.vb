Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports System.Linq

Namespace Extensions

	<HideModuleName()> _
	Friend Module ModuleSettingExtensions

		<Extension()>
	 Public Function GetSetting(Of T As ModuleSetting)(ByVal items As ICollection(Of T), ByVal moduleKey As String, ByVal parameterKey As String) As T

			Dim objSetting As ModuleSetting = items.FirstOrDefault(Function(item) item.ModuleKey = moduleKey And item.ParameterKey = parameterKey)
			If objSetting Is Nothing Then Return New ModuleSetting
			Return objSetting

		End Function

		<Extension()>
		Public Function GetByKey(Of T As ModuleSetting)(ByVal items As List(Of T), ByVal key As String) As T
			Return items.FirstOrDefault(Function(item) item.ModuleKey = key)
		End Function

	End Module
End Namespace
