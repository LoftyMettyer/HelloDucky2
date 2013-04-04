Option Strict Off
Option Explicit On
Friend Class clsCalendarEvent
	
	Private mstrEventKey As String
	Private mstrName As String 'also used for Calendar Breakdown
	Private mlngTableID As Integer
	Private mstrTableName As String
	Private mlngFilterID As Integer
	Private mlngStartDateID As Integer
	Private mstrStartDateName As String 'also used for Calendar Breakdown
	Private mlngStartSessionID As Integer
	Private mstrStartSessionName As String 'also used for Calendar Breakdown
	Private mlngEndDateID As Integer
	Private mstrEndDateName As String 'also used for Calendar Breakdown
	Private mlngEndSessionID As Integer
	Private mstrEndSessionName As String 'also used for Calendar Breakdown
	Private mlngDurationID As Integer
	Private mstrDurationName As String 'also used for Calendar Breakdown
	Private mintLegendType As Short
	Private mstrLegendCharacter As String 'also used for Calendar Breakdown
	Private mlngLegendTableID As Integer
	Private mstrLegendTableName As String
	Private mlngLegendColumnID As Integer
	Private mstrLegendColumnName As String
	Private mlngLegendCodeID As Integer
	Private mstrLegendCodeName As String
	Private mlngLegendEventTypeID As Integer
	Private mstrLegendEventTypeName As String
	Private mlngDesc1ID As Integer
	Private mstrDesc1Name As String 'also used for Calendar Breakdown
	Private mlngDesc2ID As Integer
	Private mstrDesc2Name As String 'also used for Calendar Breakdown
	
	'extra properties added for Calendar Breakdown
	Private mstrBaseDescription As String
	Private mstrRegion As String
	Private mstrWorkingPattern As String
	Private mstrDesc1Value As String
	Private mstrDesc2Value As String
	Public Property Key() As String
		Get
			Key = mstrEventKey
		End Get
		Set(ByVal Value As String)
			mstrEventKey = Value
		End Set
	End Property
	Public Property Name() As String
		Get
			Name = mstrName
		End Get
		Set(ByVal Value As String)
			mstrName = Value
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
	Public Property TableName() As String
		Get
			TableName = mstrTableName
		End Get
		Set(ByVal Value As String)
			mstrTableName = Value
		End Set
	End Property
	Public Property FilterID() As Integer
		Get
			FilterID = mlngFilterID
		End Get
		Set(ByVal Value As Integer)
			mlngFilterID = Value
		End Set
	End Property
	Public Property StartDateID() As Integer
		Get
			StartDateID = mlngStartDateID
		End Get
		Set(ByVal Value As Integer)
			mlngStartDateID = Value
		End Set
	End Property
	Public Property StartDateName() As String
		Get
			StartDateName = mstrStartDateName
		End Get
		Set(ByVal Value As String)
			mstrStartDateName = Value
		End Set
	End Property
	Public Property StartSessionID() As Integer
		Get
			StartSessionID = mlngStartSessionID
		End Get
		Set(ByVal Value As Integer)
			mlngStartSessionID = Value
		End Set
	End Property
	Public Property StartSessionName() As String
		Get
			StartSessionName = mstrStartSessionName
		End Get
		Set(ByVal Value As String)
			mstrStartSessionName = Value
		End Set
	End Property
	Public Property EndDateID() As Integer
		Get
			EndDateID = mlngEndDateID
		End Get
		Set(ByVal Value As Integer)
			mlngEndDateID = Value
		End Set
	End Property
	Public Property EndDateName() As String
		Get
			EndDateName = mstrEndDateName
		End Get
		Set(ByVal Value As String)
			mstrEndDateName = Value
		End Set
	End Property
	Public Property EndSessionID() As Integer
		Get
			EndSessionID = mlngEndSessionID
		End Get
		Set(ByVal Value As Integer)
			mlngEndSessionID = Value
		End Set
	End Property
	Public Property EndSessionName() As String
		Get
			EndSessionName = mstrEndSessionName
		End Get
		Set(ByVal Value As String)
			mstrEndSessionName = Value
		End Set
	End Property
	Public Property DurationID() As Integer
		Get
			DurationID = mlngDurationID
		End Get
		Set(ByVal Value As Integer)
			mlngDurationID = Value
		End Set
	End Property
	Public Property DurationName() As String
		Get
			DurationName = mstrDurationName
		End Get
		Set(ByVal Value As String)
			mstrDurationName = Value
		End Set
	End Property
	Public Property LegendType() As Short
		Get
			LegendType = mintLegendType
		End Get
		Set(ByVal Value As Short)
			mintLegendType = Value
		End Set
	End Property
	Public Property LegendCharacter() As String
		Get
			LegendCharacter = mstrLegendCharacter
		End Get
		Set(ByVal Value As String)
			mstrLegendCharacter = Value
		End Set
	End Property
	Public Property LegendTableID() As Integer
		Get
			LegendTableID = mlngLegendTableID
		End Get
		Set(ByVal Value As Integer)
			mlngLegendTableID = Value
		End Set
	End Property
	Public Property LegendTableName() As String
		Get
			LegendTableName = mstrLegendTableName
		End Get
		Set(ByVal Value As String)
			mstrLegendTableName = Value
		End Set
	End Property
	Public Property LegendColumnID() As Integer
		Get
			LegendColumnID = mlngLegendColumnID
		End Get
		Set(ByVal Value As Integer)
			mlngLegendColumnID = Value
		End Set
	End Property
	Public Property LegendColumnName() As String
		Get
			LegendColumnName = mstrLegendColumnName
		End Get
		Set(ByVal Value As String)
			mstrLegendColumnName = Value
		End Set
	End Property
	Public Property LegendCodeID() As Integer
		Get
			LegendCodeID = mlngLegendCodeID
		End Get
		Set(ByVal Value As Integer)
			mlngLegendCodeID = Value
		End Set
	End Property
	Public Property LegendCodeName() As String
		Get
			LegendCodeName = mstrLegendCodeName
		End Get
		Set(ByVal Value As String)
			mstrLegendCodeName = Value
		End Set
	End Property
	Public Property LegendEventTypeID() As Integer
		Get
			LegendEventTypeID = mlngLegendEventTypeID
		End Get
		Set(ByVal Value As Integer)
			mlngLegendEventTypeID = Value
		End Set
	End Property
	Public Property LegendEventTypeName() As String
		Get
			LegendEventTypeName = mstrLegendEventTypeName
		End Get
		Set(ByVal Value As String)
			mstrLegendEventTypeName = Value
		End Set
	End Property
	Public Property Description1ID() As Integer
		Get
			Description1ID = mlngDesc1ID
		End Get
		Set(ByVal Value As Integer)
			mlngDesc1ID = Value
		End Set
	End Property
	Public Property Description1Name() As String
		Get
			Description1Name = mstrDesc1Name
		End Get
		Set(ByVal Value As String)
			mstrDesc1Name = Value
		End Set
	End Property
	Public Property Description2ID() As Integer
		Get
			Description2ID = mlngDesc2ID
		End Get
		Set(ByVal Value As Integer)
			mlngDesc2ID = Value
		End Set
	End Property
	Public Property Description2Name() As String
		Get
			Description2Name = mstrDesc2Name
		End Get
		Set(ByVal Value As String)
			mstrDesc2Name = Value
		End Set
	End Property
	Public Property BaseDescription() As String
		Get
			BaseDescription = mstrBaseDescription
		End Get
		Set(ByVal Value As String)
			mstrBaseDescription = Value
		End Set
	End Property
	Public Property Region() As String
		Get
			Region = mstrRegion
		End Get
		Set(ByVal Value As String)
			mstrRegion = Value
		End Set
	End Property
	Public Property WorkingPattern() As String
		Get
			WorkingPattern = mstrWorkingPattern
		End Get
		Set(ByVal Value As String)
			mstrWorkingPattern = Value
		End Set
	End Property
	Public Property Desc1Value() As String
		Get
			Desc1Value = mstrDesc1Value
		End Get
		Set(ByVal Value As String)
			mstrDesc1Value = Value
		End Set
	End Property
	Public Property Desc2Value() As String
		Get
			Desc2Value = mstrDesc2Value
		End Get
		Set(ByVal Value As String)
			mstrDesc2Value = Value
		End Set
	End Property
End Class