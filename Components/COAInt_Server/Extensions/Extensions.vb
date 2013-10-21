Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports System.Linq
Imports HR.Intranet.Server.Metadata

Friend Module Extensions

	<Extension()>
	Public Function GetById(Of T As Base)(ByVal items As ICollection(Of T), ByVal id As Integer) As T
		Return items.FirstOrDefault(Function(item) item.ID = id)
	End Function

	<Extension()>
	Public Function IsRelation(Of T As Relation)(ByVal items As ICollection(Of T), ByVal parentid As Integer, childID As Integer) As Boolean
		Return items.Any(Function(item) item.ChildID = childID And item.ParentID = parentid)
	End Function

	<Extension()>
	Public Function GetSetting(Of T As ModuleSetting)(ByVal items As ICollection(Of T), ByVal moduleKey As String, ByVal parameterKey As String) As T

		Dim objSetting As ModuleSetting = items.FirstOrDefault(Function(item) item.ModuleKey = moduleKey And item.ParameterKey = parameterKey)
		If objSetting Is Nothing Then Return New ModuleSetting
		Return objSetting

	End Function

	<Extension()>
	Public Function GetUserSetting(Of T As UserSetting)(ByVal items As ICollection(Of T), ByVal section As String, ByVal Key As String) As T
		Return items.FirstOrDefault(Function(item) item.Section = section And item.Key = Key)
	End Function


	<Extension()>
	Public Function Item(Of T As CTablePrivilege)(ByVal items As ICollection(Of T), ByVal name As String) As T
		Return items.FirstOrDefault(Function(baseItem) (baseItem.TableName = name And baseItem.IsTable = True) Or (baseItem.ViewName = name And baseItem.IsTable = False))
	End Function

	<Extension()>
	Public Function GetItemByTableId(Of T As CTablePrivilege)(ByVal items As ICollection(Of T), ByVal id As Long) As T
		Return items.FirstOrDefault(Function(baseItem) baseItem.TableID = id)
	End Function


	<Extension()>
	Public Function Collection(Of T As CTablePrivilege)(ByVal items As ICollection(Of T)) As ICollection(Of T)
		Return items
	End Function


	' Don't know exactly what these function do or if they are necessary yet. This is just a proof of concept for speeding up the login process

	<Extension()>
	Public Function FindTableID(Of T As CTablePrivilege)(ByVal items As ICollection(Of T), ByVal id As String) As T
		Return items.FirstOrDefault(Function(baseItem) baseItem.TableID = id)
	End Function

	<Extension()>
	Public Function FindViewID(Of T As CTablePrivilege)(ByVal items As ICollection(Of T), ByVal id As Integer) As T
		Return items.FirstOrDefault(Function(baseItem) baseItem.ViewID = id)
	End Function

	<Extension()>
	Public Function FindRealSource(Of T As CTablePrivilege)(ByVal items As ICollection(Of T), ByVal name As String) As T
		Return items.FirstOrDefault(Function(baseItem) baseItem.RealSource = name)
	End Function

End Module
