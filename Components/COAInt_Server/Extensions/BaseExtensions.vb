Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports System.Linq

Namespace Extensions

	<HideModuleName()> _
	Public Module BaseExtensions

		<Extension()>
		Public Function GetById(Of T As Base)(ByVal items As ICollection(Of T), ByVal id As Integer) As T
			Return items.FirstOrDefault(Function(item) item.ID = id)
		End Function

		<Extension()>
		Public Function GetByIndex(Of T As Base)(ByVal items As ICollection(Of T), ByVal index As Integer) As T
			Return items.ElementAt(index)
		End Function

	End Module

End Namespace