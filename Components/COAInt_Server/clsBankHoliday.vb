Option Strict On
Option Explicit On
Friend Class clsBankHoliday
	
  Private _mstrRegion As String
  Private _mstrDescription As String
  Private _mdtHolidayDate As Date
  Private _mlngBaseRecordId As String

  Public Property Region() As String
    Get
      Region = _mstrRegion
    End Get
    Set(ByVal value As String)
      _mstrRegion = value
    End Set
  End Property


  Public Property Description() As String
    Get
      Description = _mstrDescription
    End Get
    Set(ByVal value As String)
      _mstrDescription = value
    End Set
  End Property


  Public Property HolidayDate() As Date
    Get
      HolidayDate = _mdtHolidayDate
    End Get
    Set(ByVal value As Date)
      _mdtHolidayDate = value
    End Set
  End Property


  Public Property BaseRecordID() As Integer
    Get
      BaseRecordID = CInt(_mlngBaseRecordId)
    End Get
    Set(ByVal value As Integer)
      _mlngBaseRecordId = CStr(value)
    End Set
  End Property
End Class