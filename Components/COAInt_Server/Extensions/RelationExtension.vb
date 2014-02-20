Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports System.Linq

Namespace Extensions

	<HideModuleName()> _
	Public Module RelationExtension

		<Extension()>
	 Public Function IsRelation(Of T As Relation)(ByVal items As ICollection(Of T), ByVal parentid As Integer, childID As Integer) As Boolean
			Return items.Any(Function(item) item.ChildID = childID And item.ParentID = parentid)
		End Function

	End Module

End Namespace