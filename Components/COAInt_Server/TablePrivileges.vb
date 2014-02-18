Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Friend Class CTablePrivileges

	Private mCol As Collection

	Friend Property Collection() As Collection
		Get
			Collection = mCol
		End Get

		Set(ByVal Value As Collection)
			mCol = Value
		End Set
	End Property

	Friend ReadOnly Property Item(vntIndexKey As String) As TablePrivilege
		Get
			Return CType(mCol.Item(vntIndexKey), TablePrivilege)
		End Get
	End Property

	Friend ReadOnly Property Count() As Integer
		Get
			Count = mCol.Count()
		End Get
	End Property

	Friend Function Add(psTableName As String, plngTableID As Integer, piTableType As TableTypes, plngDfltOrderID As Integer, plngRecDescID As Integer, pfIsTable As Boolean, plngViewID As Integer, psViewName As String) As TablePrivilege
		' Add a new member to the collection of table privileges.

		Dim skey As String
		Dim objNewMember As New TablePrivilege

		' Initialise the privileges.
		With objNewMember
			.TableID = plngTableID
			.TableName = psTableName
			.TableType = piTableType
			.DefaultOrderID = plngDfltOrderID
			.RecordDescriptionID = plngRecDescID

			.IsTable = pfIsTable

			.ViewID = plngViewID
			.ViewName = psViewName

			If (Not pfIsTable) Then
				skey = psViewName
			Else
				skey = psTableName
			End If

			.AllowSelect = False
			.AllowUpdate = False
			.AllowDelete = False
			.AllowInsert = False
		End With

		mCol.Add(objNewMember, skey)

		Add = objNewMember
		'UPGRADE_NOTE: Object objNewMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objNewMember = Nothing

	End Function


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

	Friend Function FindRealSource(psRealSource As String) As TablePrivilege
		' Return the table/view privilege object with the given real source.
		Dim objTable As TablePrivilege
		Dim objRequiredTable As TablePrivilege

		For Each objTable In mCol
			If objTable.RealSource = psRealSource Then
				objRequiredTable = objTable
				Exit For
			End If
		Next objTable
		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTable = Nothing

		FindRealSource = objRequiredTable

	End Function

	Friend Function FindTableID(plngTableID As Integer) As TablePrivilege
		' Return the table/view privilege object with the given table ID.
		Dim objTable As TablePrivilege
		Dim objRequiredTable As TablePrivilege

		For Each objTable In mCol
			' JPD 6/9/00 This function has been modified to ensure that the object returned is for the
			' given table, and not just a view on the given table.
			'    If objTable.TableID = plngTableID Then
			If (objTable.TableID = plngTableID) And (objTable.IsTable) Then
				objRequiredTable = objTable
				Exit For
			End If
		Next objTable
		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTable = Nothing

		Return objRequiredTable

	End Function

	Friend Function FindViewID(plngViewID As Integer) As TablePrivilege
		' Return the table/view privilege object with the given table ID.
		Dim objView As TablePrivilege
		Dim objRequiredView As TablePrivilege

		For Each objView In mCol
			If objView.ViewID = plngViewID Then
				objRequiredView = objView
				Exit For
			End If
		Next objView
		'UPGRADE_NOTE: Object objView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objView = Nothing

		Return objRequiredView

	End Function
End Class