Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.Enums

Friend Class clsColumn
	' Note : ColType and ID are combined to provide a key, ie C24, E107 etc
	'        But are stored in separate fields to allow checking when
	'        deleting fields etc.


	Private msColType As String	' "C" for Column or "E" for expression
	Private mlID As Integer	' ID of the Column or Expression
	Private msHeading As String	' Report Col Heading
	Private mlSize As Integer	' Number of chars to display
	Private miDecPlaces As Short ' Display to number of d.p. (numerics only)
	Private mbAverage As Boolean ' Avg this col/expression (numerics only)
	Private mbCount As Boolean ' Count this col/expression
	Private mbTotal As Boolean ' Total this col/expression (numerics only)
	Private mbHidden As Boolean	' Hidden this col/expression is hidden in the report
	Private mbGroupWithNext As Boolean 'Group With Next this col/expression is groups with next column
	Private mbIsNumeric As Boolean ' Is this col/exp numeric ?
	Private miSequence As Short
	Private mlSortSeq As Integer
	Private msSortDir As String	'Sort Direction
	Private mlTableID As Integer
	Private msTableName As String
	Private msColumnName As String
	Private mlDataType As Integer
	Private mblnThousandSeparator As Boolean
	Private mstrSQL As String

	Public Property GroupWithNext() As Boolean
		Get
			GroupWithNext = mbGroupWithNext
		End Get
		Set(ByVal Value As Boolean)
			mbGroupWithNext = Value
		End Set
	End Property

	Public Property Hidden() As Boolean
		Get
			Hidden = mbHidden
		End Get
		Set(ByVal Value As Boolean)
			mbHidden = Value
		End Set
	End Property


	Public Property ColType() As String
		Get
			ColType = msColType
		End Get
		Set(ByVal Value As String)
			msColType = Value
		End Set
	End Property


	Public Property ID() As Integer
		Get
			ID = mlID
		End Get
		Set(ByVal Value As Integer)
			mlID = Value
		End Set
	End Property


	Public Property Heading() As String
		Get
			Heading = msHeading
		End Get
		Set(ByVal Value As String)
			msHeading = Value
		End Set
	End Property


	Public Property Size() As Integer
		Get
			Size = mlSize
		End Get
		Set(ByVal Value As Integer)
			mlSize = Value
		End Set
	End Property


	Public Property DecPlaces() As Short
		Get
			DecPlaces = miDecPlaces
		End Get
		Set(ByVal Value As Short)
			miDecPlaces = Value
		End Set
	End Property


	Public Property Average() As Boolean
		Get
			Average = mbAverage
		End Get
		Set(ByVal Value As Boolean)
			mbAverage = Value
		End Set
	End Property


	Public Property Count() As Boolean
		Get
			Count = mbCount
		End Get
		Set(ByVal Value As Boolean)
			mbCount = Value
		End Set
	End Property


	Public Property Total() As Boolean
		Get
			Total = mbTotal
		End Get
		Set(ByVal Value As Boolean)
			mbTotal = Value
		End Set
	End Property


	'UPGRADE_NOTE: IsNumeric was upgraded to IsNumeric_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Property IsNumeric_Renamed() As Boolean
		Get
			IsNumeric_Renamed = mbIsNumeric
		End Get
		Set(ByVal Value As Boolean)
			mbIsNumeric = Value
		End Set
	End Property


	Public Property Sequence() As Short
		Get
			Sequence = miSequence
		End Get
		Set(ByVal Value As Short)
			miSequence = Value
		End Set
	End Property


	Public Property SortSeq() As Integer
		Get
			SortSeq = mlSortSeq
		End Get
		Set(ByVal Value As Integer)
			mlSortSeq = Value
		End Set
	End Property


	Public Property SortDir() As String
		Get
			SortDir = msSortDir
		End Get
		Set(ByVal Value As String)
			msSortDir = Value
		End Set
	End Property


	Public Property TableID() As Integer
		Get
			TableID = mlTableID
		End Get
		Set(ByVal Value As Integer)
			mlTableID = Value
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


	Public Property ColumnName() As String
		Get
			ColumnName = msColumnName
		End Get
		Set(ByVal Value As String)
			msColumnName = Value
		End Set
	End Property


	Public Property DataType() As SQLDataType
		Get
			DataType = mlDataType
		End Get
		Set(ByVal Value As SQLDataType)
			mlDataType = Value
		End Set
	End Property


	Public Property SQL() As String
		Get
			SQL = mstrSQL
		End Get
		Set(ByVal Value As String)
			mstrSQL = Value
		End Set
	End Property


	Public Property ThousandSeparator() As Boolean
		Get
			ThousandSeparator = mblnThousandSeparator
		End Get
		Set(ByVal Value As Boolean)
			mblnThousandSeparator = Value
		End Set
	End Property
End Class