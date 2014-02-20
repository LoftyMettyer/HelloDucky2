Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports System.Linq

Namespace Extensions

	<HideModuleName()> _
	Friend Module UserSettingExtensions

		<Extension()>
	 Public Function GetUserSetting(Of T As UserSetting)(ByVal items As ICollection(Of T), ByVal section As String, ByVal Key As String) As T
			Return items.FirstOrDefault(Function(item) item.Section = section And item.Key = Key)
		End Function

		<Extension()>
		Public Function GetSetting(Of T As UserSetting)(ByVal items As ICollection(Of T), ByVal section As String, ByVal Key As String, ByVal [Default] As Object) As T

			'TODO If not found return one with default default set

			Return items.FirstOrDefault(Function(item) item.Section = section And item.Key = Key)
		End Function

	End Module

End Namespace