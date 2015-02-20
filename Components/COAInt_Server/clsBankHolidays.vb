Option Strict Off
Option Explicit On
Friend Class clsBankHolidays
	Implements System.Collections.IEnumerable
	
	
	' Hold local collection
	Private mCol As Collection
	
	Public Function Add(ByRef pstrRegion As String, ByRef pstrDescription As String, ByRef pdtHolidayDate As Date) As clsBankHoliday
		
		' Add a new object to the collection
		
		Dim objNewMember As clsBankHoliday
		objNewMember = New clsBankHoliday
		
		objNewMember.Region = pstrRegion
		objNewMember.Description = pstrDescription
		objNewMember.HolidayDate = pdtHolidayDate
		
		mCol.Add(objNewMember)
		
		Add = objNewMember
		
		'UPGRADE_NOTE: Object objNewMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objNewMember = Nothing
		
	End Function
	
	
	Public Property Collection() As Collection
		Get
			Collection = mCol
		End Get
		Set(ByVal Value As Collection)
			mCol = Value
		End Set
	End Property
	
	Public ReadOnly Property Item(sColTypeAndID As String) As clsBankHoliday
		Get
			' Provide a reference to a specific item in the collection
			Return mCol.Item(sColTypeAndID)
		End Get
	End Property
	
	Public ReadOnly Property Count() As Integer
		Get
			
			' Provide number of objects in the collection
			
			Count = mCol.Count()
			
		End Get
	End Property
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
		'Get
			'
			'NewEnum = mCol._NewEnum
			'
		'End Get
	'End Property
	
	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
		'GetEnumerator = mCol.GetEnumerator
	End Function
	
	Public Sub Remove(ByRef sColTypeAndID As String)
		
		' Remove a specific object from the collection
		
		mCol.Remove(sColTypeAndID)
		
	End Sub
	
	Public Sub RemoveAll()
		
		' Remove all object from the collection
		
		Do While mCol.Count() > 1
			mCol.Remove(mCol.Count())
		Loop 
		
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