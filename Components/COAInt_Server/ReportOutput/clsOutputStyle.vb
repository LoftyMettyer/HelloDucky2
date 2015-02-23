Option Strict On
Option Explicit On

Namespace ReportOutput
	Friend Class clsOutputStyle

		Private mlngStartCol As Integer
		Private mlngStartRow As Integer
		Private mlngEndCol As Integer
		Private mlngEndRow As Integer
		Private mblnGridlines As Boolean
		Private mblnBold As Boolean
		Private mblnUnderLine As Boolean
		Private mlngBackCol As Integer
		Private mlngForeCol As Integer
		Private mblnCenterText As Boolean
		Private mstrName As String
		Private mlngBackCol97 As Integer 'Colour Index for Word 97
		Private mlngForeCol97 As Integer 'Colour Index for Word 97

		Public ReadOnly Property Font As Font
			Get
				Return New Font("Calibri", 11)
			End Get
		End Property

		Public Property StartCol() As Integer
			Get
				StartCol = mlngStartCol
			End Get
			Set(ByVal Value As Integer)
				mlngStartCol = Value
			End Set
		End Property

		Public Property StartRow() As Integer
			Get
				StartRow = mlngStartRow
			End Get
			Set(ByVal Value As Integer)
				mlngStartRow = Value
			End Set
		End Property

		Public Property EndCol() As Integer
			Get
				EndCol = mlngEndCol
			End Get
			Set(ByVal Value As Integer)
				mlngEndCol = Value
			End Set
		End Property

		Public Property EndRow() As Integer
			Get
				EndRow = mlngEndRow
			End Get
			Set(ByVal Value As Integer)
				mlngEndRow = Value
			End Set
		End Property

		Public Property Gridlines() As Boolean
			Get
				Gridlines = mblnGridlines
			End Get
			Set(ByVal Value As Boolean)
				mblnGridlines = Value
			End Set
		End Property

		Public Property Bold() As Boolean
			Get
				Bold = mblnBold
			End Get
			Set(ByVal Value As Boolean)
				mblnBold = Value
			End Set
		End Property

		Public Property Underline() As Boolean
			Get
				Underline = mblnUnderLine
			End Get
			Set(ByVal Value As Boolean)
				mblnUnderLine = Value
			End Set
		End Property

		Public Property ForeCol() As Integer
			Get
				ForeCol = mlngForeCol
			End Get
			Set(ByVal Value As Integer)
				mlngForeCol = Value
			End Set
		End Property

		Public Property BackCol() As Integer
			Get
				BackCol = mlngBackCol
			End Get
			Set(ByVal Value As Integer)
				mlngBackCol = Value
			End Set
		End Property

		Public Property CenterText() As Boolean
			Get
				CenterText = mblnCenterText
			End Get
			Set(ByVal Value As Boolean)
				mblnCenterText = Value
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

		Public Property ForeCol97() As Integer
			Get
				ForeCol97 = mlngForeCol97
			End Get
			Set(ByVal Value As Integer)
				mlngForeCol97 = Value
			End Set
		End Property

		Public Property BackCol97() As Integer
			Get
				BackCol97 = mlngBackCol97
			End Get
			Set(ByVal Value As Integer)
				mlngBackCol97 = Value
			End Set
		End Property
	End Class
End Namespace