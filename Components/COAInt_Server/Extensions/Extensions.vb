Option Explicit On
Option Strict Off

Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports System.Linq
Imports System.ComponentModel
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures

Friend Module Extensions

	<Extension()>
	Public Function GetById(Of T As Base)(ByVal items As ICollection(Of T), ByVal id As Integer) As T
		Return items.FirstOrDefault(Function(item) item.ID = id)
	End Function

	<Extension()>
	Public Function GetByIndex(Of T As Base)(ByVal items As ICollection(Of T), ByVal index As Integer) As T
		Return items.ElementAt(index)
	End Function

	<Extension()>
	Public Function GetLegend(Of T As CalendarLegend)(ByVal items As ICollection(Of T), ByVal key As String) As T
		Return items.FirstOrDefault(Function(item) item.LegendKey = key)
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

	'<Extension()>
	'Public Function WithRealSource(Of T As TablePrivilege)(ByVal items As ICollection(Of T)) As ICollection(Of T)

	'	Dim obj = items.Select(Function(baseitem) (Not baseitem.RealSource Is Nothing))
	'	Dim objColl As New Collection




	'	objColl = obj.ToList()
	'	Return Collection(Of T)()

	'	'Return items.en.Select(Function(baseitem) baseitem)

	'	'Return items.All(Function(baseitem) Not baseitem.RealSource Is Nothing)

	'	'Return items.Any(Function(baseitem) Not baseitem.RealSource Is Nothing)

	'	'Return CType(items.Select(Function(baseitem) (Not baseitem.RealSource Is Nothing)).ToList())

	'	Dim c As ICollection(Of T) = DirectCast([Get]("pk"), ICollection(Of T))

	'End Function





	' Don't know exactly what these function do or if they are necessary yet. This is just a proof of concept for speeding up the login process

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

' end of don't know what its doing!


		<Extension()>
	Friend Function GetByKey(Of T As Permission)(ByVal items As ICollection(Of T), ByVal key As String) As Boolean

		Dim objPermission = Permissions.FirstOrDefault(Function(baseItem) (baseItem.Key = key))
		If objPermission Is Nothing Then
			Return False
		End If

		Return objPermission.IsPermitted

	End Function

	<Extension()>
	Public Function GetByKey(Of T As ModuleSetting)(ByVal items As List(Of T), ByVal key As String) As T
		Return items.FirstOrDefault(Function(item) item.ModuleKey = key)
	End Function

	<Extension()>
	Public Function ToDataTable(Of T)(list As ICollection(Of T)) As DataTable
		Dim table As DataTable = clsDataAccess.CreateTable(Of T)()
		Dim entityType As Type = GetType(T)
		Dim properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(entityType)

		For Each item As T In list
			Dim row As DataRow = table.NewRow()

			For Each prop As PropertyDescriptor In properties
				row(prop.Name) = prop.GetValue(item)
			Next

			table.Rows.Add(row)
		Next

		Return table
	End Function



End Module
