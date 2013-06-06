VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Note : ColType and ID are combined to provide a key, ie C24, E107 etc
'        But are stored in separate fields to allow checking when
'        deleting fields etc.

Option Explicit

Private msColType As String    ' "C" for Column or "E" for expression
Private mlID As Long           ' ID of the Column or Expression
Private msHeading As String    ' Report Col Heading
Private mlSize As Long         ' Number of chars to display
Private miDecPlaces As Integer ' Display to number of d.p. (numerics only)
Private mbAverage As Boolean   ' Avg this col/expression (numerics only)
Private mbCount As Boolean     ' Count this col/expression
Private mbTotal As Boolean     ' Total this col/expression (numerics only)
Private mbHidden As Boolean    ' Hidden this col/expression is hidden in the report
Private mbGroupWithNext As Boolean  'Group With Next this col/expression is groups with next column
Private mbIsNumeric As Boolean ' Is this col/exp numeric ?
Private miSequence As Integer
Private mlSortSeq As Long
Private msSortDir As String      'Sort Direction
Private mlTableID As Long
Private msTableName As String
Private msColumnName As String
Private mlDataType As Long
Private mblnThousandSeparator As Boolean
Private mstrSQL As String

Private mbBreakOnChange As Boolean
Private mbPageOnChange As Boolean
Private mbValueOnChange As Boolean
Private mbSurpressRepeatedValues As Boolean
Private mbRepetition As Boolean

Public Property Get BreakOnChange() As Boolean
  BreakOnChange = mbBreakOnChange
End Property
Public Property Let BreakOnChange(ByVal Value As Boolean)
  mbBreakOnChange = Value
End Property

Public Property Get PageOnChange() As Boolean
  PageOnChange = mbPageOnChange
End Property
Public Property Let PageOnChange(ByVal Value As Boolean)
  mbPageOnChange = Value
End Property

Public Property Get ValueOnChange() As Boolean
  ValueOnChange = mbValueOnChange
End Property
Public Property Let ValueOnChange(ByVal Value As Boolean)
  mbValueOnChange = Value
End Property

Public Property Get SurpressRepeatedValues() As Boolean
  SurpressRepeatedValues = mbSurpressRepeatedValues
End Property
Public Property Let SurpressRepeatedValues(ByVal Value As Boolean)
  mbSurpressRepeatedValues = Value
End Property

Public Property Get Repetition() As Boolean
  Repetition = mbRepetition
End Property
Public Property Let Repetition(ByVal Value As Boolean)
  mbRepetition = Value
End Property

Public Property Get GroupWithNext() As Boolean
  GroupWithNext = mbGroupWithNext
End Property
Public Property Let GroupWithNext(ByVal bGroupWithNext As Boolean)
  mbGroupWithNext = bGroupWithNext
End Property

Public Property Get Hidden() As Boolean
  Hidden = mbHidden
End Property
Public Property Let Hidden(ByVal Value As Boolean)
  mbHidden = Value
End Property

Public Property Get ColType() As String
  ColType = msColType
End Property

Public Property Let ColType(ByVal sColType As String)
  msColType = sColType
End Property

Public Property Get ID() As Long
  ID = mlID
End Property

Public Property Let ID(ByVal lID As Long)
  mlID = lID
End Property

Public Property Get Heading() As String
  Heading = msHeading
End Property

Public Property Let Heading(ByVal sHeading As String)
  msHeading = sHeading
End Property

Public Property Get Size() As Long
  Size = mlSize
End Property

Public Property Let Size(ByVal lSize As Long)
  mlSize = lSize
End Property

Public Property Get DecPlaces() As Integer
  DecPlaces = miDecPlaces
End Property

Public Property Let DecPlaces(ByVal iDecPlaces As Integer)
  miDecPlaces = iDecPlaces
End Property

Public Property Get Average() As Boolean
  Average = mbAverage
End Property

Public Property Let Average(ByVal bAverage As Boolean)
  mbAverage = bAverage
End Property

Public Property Get Count() As Boolean
  Count = mbCount
End Property

Public Property Let Count(ByVal bCount As Boolean)
  mbCount = bCount
End Property

Public Property Get Total() As Boolean
  Total = mbTotal
End Property

Public Property Let Total(ByVal bTotal As Boolean)
  mbTotal = bTotal
End Property

Public Property Get IsNumeric() As Boolean
  IsNumeric = mbIsNumeric
End Property

Public Property Let IsNumeric(ByVal bIsNumeric As Boolean)
  mbIsNumeric = bIsNumeric
End Property

Public Property Get Sequence() As Integer
  Sequence = miSequence
End Property

Public Property Let Sequence(ByVal iSequence As Integer)
  miSequence = iSequence
End Property

Public Property Get SortSeq() As Long
  SortSeq = mlSortSeq
End Property

Public Property Let SortSeq(ByVal lNewValue As Long)
  mlSortSeq = lNewValue
End Property

Public Property Get SortDir() As String
  SortDir = msSortDir
End Property

Public Property Let SortDir(ByVal sNewValue As String)
  msSortDir = sNewValue
End Property

Public Property Get TableID() As Long
  TableID = mlTableID
End Property

Public Property Let TableID(ByVal lNewValue As Long)
  mlTableID = lNewValue
End Property

Public Property Get TableName() As String
  TableName = msTableName
End Property

Public Property Let TableName(ByVal sNewValue As String)
  msTableName = sNewValue
End Property

Public Property Get ColumnName() As String
  ColumnName = msColumnName
End Property

Public Property Let ColumnName(ByVal sNewValue As String)
  msColumnName = sNewValue
End Property

Public Property Get DataType() As SQLDataType
  DataType = mlDataType
End Property

Public Property Let DataType(ByVal vNewValue As SQLDataType)
  mlDataType = vNewValue
End Property

Public Property Get SQL() As String
  SQL = mstrSQL
End Property

Public Property Let SQL(ByVal mstrNewValue As String)
  mstrSQL = mstrNewValue
End Property

Public Property Get ThousandSeparator() As Boolean
  ThousandSeparator = mblnThousandSeparator
End Property

Public Property Let ThousandSeparator(ByVal blnNewValue As Boolean)
  mblnThousandSeparator = blnNewValue
End Property

