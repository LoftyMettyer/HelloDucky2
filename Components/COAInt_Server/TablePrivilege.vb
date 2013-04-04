Option Strict Off
Option Explicit On
Friend Class CTablePrivilege
	
	Private mlngTableID As Integer
	Private msTableName As String
	Private miTableType As Declarations.TableTypes
	Private mlngDfltOrderID As Integer
	Private mlngRecDescID As Integer
	
	Private mlngViewID As Integer
	Private msViewName As String
	
	Private mfIsTable As Boolean
	
	Private msRealSource As String
	
	Private mfSelect As Boolean
	Private mfUpdate As Boolean
	Private mfInsert As Boolean
	Private mfDelete As Boolean
	
	
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
	
	
	Public Property AllowInsert() As Boolean
		Get
			AllowInsert = mfInsert
			
		End Get
		Set(ByVal Value As Boolean)
			mfInsert = Value
			
		End Set
	End Property
	
	
	Public Property AllowDelete() As Boolean
		Get
			AllowDelete = mfDelete
			
		End Get
		Set(ByVal Value As Boolean)
			mfDelete = Value
			
		End Set
	End Property
	
	Public Property TableID() As Integer
		Get
			TableID = mlngTableID
			
		End Get
		Set(ByVal Value As Integer)
			mlngTableID = Value
			
		End Set
	End Property
	Public Property ViewID() As Integer
		Get
			ViewID = mlngViewID
			
		End Get
		Set(ByVal Value As Integer)
			mlngViewID = Value
			
		End Set
	End Property
	
	
	
	
	Public Property RealSource() As String
		Get
			RealSource = msRealSource
			
		End Get
		Set(ByVal Value As String)
			msRealSource = Value
			
		End Set
	End Property
	
	
	
	Public Property IsTable() As Boolean
		Get
			IsTable = mfIsTable
			
		End Get
		Set(ByVal Value As Boolean)
			mfIsTable = Value
			
		End Set
	End Property
	
	
	
	
	Public Property TableName() As String
		Get
			TableName = msTableName
			
		End Get
		Set(ByVal Value As String)
			msTableName = Value
			
		End Set
	End Property
	Public Property ViewName() As String
		Get
			ViewName = msViewName
			
		End Get
		Set(ByVal Value As String)
			msViewName = Value
			
		End Set
	End Property
	
	
	
	Public Property TableType() As Declarations.TableTypes
		Get
			TableType = miTableType
			
		End Get
		Set(ByVal Value As Declarations.TableTypes)
			miTableType = Value
			
		End Set
	End Property
	
	
	Public Property DefaultOrderID() As Integer
		Get
			DefaultOrderID = mlngDfltOrderID
			
		End Get
		Set(ByVal Value As Integer)
			mlngDfltOrderID = Value
			
		End Set
	End Property
	
	
	Public Property RecordDescriptionID() As Integer
		Get
			RecordDescriptionID = mlngRecDescID
			
		End Get
		Set(ByVal Value As Integer)
			mlngRecDescID = Value
			
		End Set
	End Property
End Class