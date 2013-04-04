Option Strict Off
Option Explicit On
Friend Class CTablePrivileges
	
	Private mCol As Collection
	
	
	Public Property Collection() As Collection
		Get
			Collection = mCol
			
		End Get
		Set(ByVal Value As Collection)
			mCol = Value
			
		End Set
	End Property
	
	Public ReadOnly Property Item(ByVal vntIndexKey As Object) As CTablePrivilege
		Get
			Item = mCol.Item(vntIndexKey)
			
		End Get
	End Property
	
	Public ReadOnly Property Count() As Integer
		Get
			
			Count = mCol.Count()
			
		End Get
	End Property
	
	Public Function Add(ByRef psTableName As String, ByRef plngTableID As Integer, ByRef piTableType As Short, ByRef plngDfltOrderID As Integer, ByRef plngRecDescID As Integer, ByRef pfIsTable As Boolean, ByRef plngViewID As Integer, ByRef psViewName As String) As CTablePrivilege
		' Add a new member to the collection of table privileges.
		Dim lngChildViewID As Integer
		Dim skey As String
		Dim objNewMember As New CTablePrivilege
		
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
	
	
	
	
	
	Public Function FindRealSource(ByRef psRealSource As String) As CTablePrivilege
		' Return the table/view privilege object with the given real source.
		Dim objTable As CTablePrivilege
		Dim objRequiredTable As CTablePrivilege
		
		For	Each objTable In mCol
			If objTable.RealSource = psRealSource Then
				objRequiredTable = objTable
				Exit For
			End If
		Next objTable
		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTable = Nothing
		
		FindRealSource = objRequiredTable
		
	End Function
	Public Function FindTableID(ByRef plngTableID As Integer) As CTablePrivilege
		' Return the table/view privilege object with the given table ID.
		Dim objTable As CTablePrivilege
		Dim objRequiredTable As CTablePrivilege
		
		For	Each objTable In mCol
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
		
		FindTableID = objRequiredTable
		
	End Function
	
	Public Function FindViewID(ByRef plngViewID As Integer) As CTablePrivilege
		' Return the table/view privilege object with the given table ID.
		Dim objView As CTablePrivilege
		Dim objRequiredView As CTablePrivilege
		
		For	Each objView In mCol
			If objView.ViewID = plngViewID Then
				objRequiredView = objView
				Exit For
			End If
		Next objView
		'UPGRADE_NOTE: Object objView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objView = Nothing
		
		FindViewID = objRequiredView
		
	End Function
End Class