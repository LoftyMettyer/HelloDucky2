Option Strict Off
Option Explicit On
Friend Class clsBankHoliday
	
	Private mstrRegion As String
	Private mstrDescription As String
	Private mdtHolidayDate As Date
	Private mlngBaseRecordID As String
	
	Public Property Region() As String
		Get
			Region = mstrRegion
		End Get
		Set(ByVal Value As String)
			mstrRegion = Value
		End Set
	End Property
	
	
	Public Property Description() As String
		Get
			Description = mstrDescription
		End Get
		Set(ByVal Value As String)
			mstrDescription = Value
		End Set
	End Property
	
	
	Public Property HolidayDate() As Date
		Get
			HolidayDate = mdtHolidayDate
		End Get
		Set(ByVal Value As Date)
			mdtHolidayDate = Value
		End Set
	End Property
	
	
	Public Property BaseRecordID() As Integer
		Get
			BaseRecordID = CInt(mlngBaseRecordID)
		End Get
		Set(ByVal Value As Integer)
			mlngBaseRecordID = CStr(Value)
		End Set
	End Property
End Class