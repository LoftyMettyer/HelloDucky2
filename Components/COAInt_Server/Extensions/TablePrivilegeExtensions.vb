Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports System.Linq

Namespace Extensions

	<HideModuleName()> _
	Friend Module TablePrivilegeExtensions

		<Extension()>
	 Public Function Item(Of T As TablePrivilege)(ByVal items As ICollection(Of T), ByVal name As String) As T
			Return items.FirstOrDefault(Function(baseItem) (baseItem.TableName = name.ToUpper() And baseItem.IsTable = True) Or (baseItem.ViewName = name And baseItem.IsTable = False))
		End Function

		<Extension()>
		Public Function GetItemByTableId(Of T As TablePrivilege)(ByVal items As ICollection(Of T), ByVal id As Long) As T
			Return items.FirstOrDefault(Function(baseItem) baseItem.TableID = id)
		End Function

		<Extension()>
		Public Function Collection(Of T As TablePrivilege)(ByVal items As ICollection(Of T)) As ICollection(Of T)
			Return items
		End Function

		<Extension()>
		Public Function FindTableID(Of T As TablePrivilege)(ByVal items As ICollection(Of T), ByVal id As Integer) As T
			Return items.FirstOrDefault(Function(baseItem) baseItem.TableID = id)
		End Function

		<Extension()>
		Public Function FindViewID(Of T As TablePrivilege)(ByVal items As ICollection(Of T), ByVal id As Integer) As T
			Return items.FirstOrDefault(Function(baseItem) baseItem.ViewID = id)
		End Function

		<Extension()>
		Public Function FindRealSource(Of T As TablePrivilege)(ByVal items As ICollection(Of T), ByVal name As String) As T
			Return items.FirstOrDefault(Function(baseItem) baseItem.RealSource = name)
		End Function

	End Module

End Namespace
