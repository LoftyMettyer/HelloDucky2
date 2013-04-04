Option Strict Off
Option Explicit On
Friend Class CColumnPrivilege
	
	Private msColumnName As String
	Private mfSelect As Boolean
	Private mfUpdate As Boolean
	Private miColumnType As Short
	Private miDataType As Short
	Private mlngColumnID As Integer
	Private mfUniqueCheck As Boolean
	
	
	Public Property AllowSelect() As Boolean
		Get
			AllowSelect = mfSelect
			
		End Get
		Set(ByVal Value As Boolean)
			mfSelect = Value
			
		End Set
	End Property
	
	
	Public Property AllowUpdate() As Boolean
		Get
			AllowUpdate = mfUpdate
			
		End Get
		Set(ByVal Value As Boolean)
			mfUpdate = Value
			
		End Set
	End Property
	
	
	Public Property ColumnName() As String
		Get
			ColumnName = msColumnName
			
		End Get
		Set(ByVal Value As String)
			msColumnName = Value
			
		End Set
	End Property
	
	
	
	Public Property ColumnType() As Short
		Get
			ColumnType = miColumnType
			
		End Get
		Set(ByVal Value As Short)
			miColumnType = Value
			
		End Set
	End Property
	
	
	Public Property DataType() As Short
		Get
			DataType = miDataType
			
		End Get
		Set(ByVal Value As Short)
			miDataType = Value
			
		End Set
	End Property
	
	
	Public Property ColumnID() As Integer
		Get
			ColumnID = mlngColumnID
			
		End Get
		Set(ByVal Value As Integer)
			mlngColumnID = Value
			
		End Set
	End Property
	
	
	Public Property UniqueCheck() As Boolean
		Get
			UniqueCheck = mfUniqueCheck
			
		End Get
		Set(ByVal Value As Boolean)
			mfUniqueCheck = Value
			
		End Set
	End Property
End Class