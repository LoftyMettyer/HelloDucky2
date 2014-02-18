Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Friend Class CColumnPrivileges
	Implements IEnumerable

	'local variable to hold collection
	Private mCol As Collection
	Private msTag As String

	Friend Function IsValid(pvIndexKey As String) As Boolean
		Return mCol.Contains(pvIndexKey)
	End Function

	Friend Function Add(pfSelect As Boolean, pfUpdate As Boolean, psColumnName As String, piColumnType As Short, piDataType As SQLDataType, plngColumnID As Integer, pfUniqueCheck As Boolean) As ColumnPrivilege
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

		Return objNewMember

	End Function

	Friend Function FindColumnID(plngColumnID As Integer) As ColumnPrivilege
		' Return the column privilege object with the given column ID.
		Dim objRequiredColumn As ColumnPrivilege = Nothing

		For Each objColumn As ColumnPrivilege In mCol
			If objColumn.ColumnID = plngColumnID Then
				objRequiredColumn = objColumn
				Exit For
			End If
		Next objColumn

		Return objRequiredColumn
	End Function

	Default Friend ReadOnly Property Item(vntIndexKey As String) As ColumnPrivilege
		Get
			Return CType(mCol.Item(vntIndexKey), ColumnPrivilege)

		End Get
	End Property

	Friend ReadOnly Property Count() As Integer
		Get
			Count = mCol.Count()

		End Get
	End Property

	Friend Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
		Return mCol.GetEnumerator
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

	Public Sub New()
		MyBase.New()
		mCol = New Collection
	End Sub

	Protected Overrides Sub Finalize()
		mCol = Nothing

		MyBase.Finalize()
	End Sub
End Class