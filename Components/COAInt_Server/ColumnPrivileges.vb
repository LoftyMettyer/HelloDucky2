Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.Metadata

Friend Class CColumnPrivileges
	Implements IEnumerable

	'local variable to hold collection
	Private mCol As Collection
	Private msTag As String

	Public Function IsValid(ByVal pvIndexKey As String) As Boolean
		Return mCol.Contains(pvIndexKey)
	End Function

	Public Function Add(ByVal pfSelect As Boolean, ByVal pfUpdate As Boolean, ByVal psColumnName As String, ByVal piColumnType As Short, ByVal piDataType As Short, ByVal plngColumnID As Integer, ByVal pfUniqueCheck As Boolean) As ColumnPrivilege
		'create a new object
		Dim objNewMember As ColumnPrivilege
		objNewMember = New ColumnPrivilege

		With objNewMember
			.ColumnName = psColumnName
			.AllowSelect = pfSelect
			.AllowUpdate = pfUpdate
			.DataType = piDataType
			.ColumnID = plngColumnID
		End With

		mCol.Add(objNewMember, psColumnName)

		'return the object created
		Add = objNewMember
		'UPGRADE_NOTE: Object objNewMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objNewMember = Nothing

	End Function

	Public Function FindColumnID(ByRef plngColumnID As Integer) As ColumnPrivilege
		' Return the column privilege object with the given column ID.
		Dim objRequiredColumn As ColumnPrivilege = Nothing

		For Each objColumn As ColumnPrivilege In mCol
			If objColumn.ColumnID = plngColumnID Then
				objRequiredColumn = objColumn
				Exit For
			End If
		Next objColumn
		
		FindColumnID = objRequiredColumn
	End Function

	Default Public ReadOnly Property Item(ByVal vntIndexKey As Object) As ColumnPrivilege
		Get
			Return mCol.Item(vntIndexKey)

		End Get
	End Property

	Public ReadOnly Property Count() As Integer
		Get
			Count = mCol.Count()

		End Get
	End Property

	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
	'Get
	'NewEnum = mCol._NewEnum
	'
	'End Get
	'End Property

	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		GetEnumerator = mCol.GetEnumerator
	End Function



	Public Property Tag() As String
		Get
			' Return the object's tag.
			Tag = msTag

		End Get
		Set(ByVal Value As String)
			' Set the object's tag property.
			msTag = Value

		End Set
	End Property


	Public Sub Remove(ByRef vntIndexKey As Object)
		mCol.Remove(vntIndexKey)

	End Sub

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mCol = New Collection

	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing

	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class