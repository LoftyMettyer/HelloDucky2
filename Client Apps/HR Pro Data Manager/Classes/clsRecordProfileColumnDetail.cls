VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordProfileColDtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msType As String       ' "C" for Column / "H" for heading / "S" for separator
Private mlngID As Long         ' ID of the Column
Private msHeading As String    ' Report Col Heading
Private mlngSize As Long       ' Number of chars to display
Private miDecPlaces As Integer ' Display to number of d.p. (numerics only)
Private miSequence As Integer
Private mlngTableID As Long
Private msTableName As String
Private msColumnName As String
Private miDataType As Integer
'Private msDigitSeparator As String
'Private miDigitGrouping As Integer
Private mlngGridColumnWidth As Long
Private mblnThousandSeparator As Boolean
Private mbBlankIfZero As Boolean
Private miOLEType As HRProDataMgr.OLEType

Private miPositionInRecordset As Integer
Private mfDisplayed As Boolean

Private msKey As String

Private Const sTYPECODE_HEADING = "H"
Private Const sTYPECODE_SEPARATOR = "S"
Private Const sTYPECODE_COLUMN = "C"
Private Const sTYPECODE_ID = "I"
Public Property Get DecPlaces() As Integer
  DecPlaces = miDecPlaces
End Property


Public Property Let DecPlaces(ByVal piDecPlaces As Integer)
  miDecPlaces = piDecPlaces
End Property


Public Property Let ColumnName(ByVal psNewValue As String)
  msColumnName = psNewValue
  
End Property


Public Property Get DataType() As SQLDataType
  DataType = miDataType
  
End Property


Public Property Let DataType(ByVal piNewValue As SQLDataType)
  miDataType = piNewValue
  
End Property

Public Property Get Heading() As String
  Heading = msHeading
  
End Property


Public Property Let Heading(ByVal psHeading As String)
  msHeading = psHeading
  
End Property


Public Property Get ID() As Long
  ID = mlngID
  
End Property


Public Property Let ID(ByVal plngID As Long)
  mlngID = plngID
  
End Property


Public Function IsNumeric() As Boolean
  IsNumeric = ((miDataType = sqlNumeric) Or _
    (miDataType = sqlInteger)) And _
    (msType = sTYPECODE_COLUMN)
  
End Function

Public Function IsPhoto() As Boolean
  IsPhoto = (miDataType = sqlVarBinary) And _
    (msType = sTYPECODE_COLUMN)
  
End Function
Public Function IsDate() As Boolean
  IsDate = (miDataType = sqlDate) And _
    (msType = sTYPECODE_COLUMN)
  
End Function



Public Function IsLogic() As Boolean
  IsLogic = (miDataType = sqlBoolean) And _
    (msType = sTYPECODE_COLUMN)
  
End Function





Public Property Get Sequence() As Integer
  Sequence = miSequence
  
End Property


Public Property Let Sequence(ByVal piSequence As Integer)
  miSequence = piSequence
  
End Property


Public Property Get Size() As Long
  Size = mlngSize
  
End Property


Public Property Let Size(ByVal plngSize As Long)
  mlngSize = plngSize
  
End Property


Public Property Get TableID() As Long
  TableID = mlngTableID
  
End Property


Public Property Let TableID(ByVal plngNewValue As Long)
  mlngTableID = plngNewValue
  
End Property


Public Property Get TableName() As String
  TableName = msTableName
  
End Property


Public Property Let TableName(ByVal psNewValue As String)
  msTableName = psNewValue
  
End Property


Public Property Get ColumnName() As String
  ColumnName = msColumnName
  
End Property





Public Property Let Key(ByVal psKey As String)
  msKey = psKey
  
End Property


Public Property Let ColType(ByVal psColType As String)
  msType = psColType
  
End Property


Public Property Get Key() As String
  Key = msKey
  
End Property



Public Property Get ColType() As String
  ColType = msType
  
End Property
Public Property Get PositionInRecordset() As Integer
  PositionInRecordset = miPositionInRecordset
  
End Property

Public Property Let PositionInRecordset(ByVal piNewValue As Integer)
  miPositionInRecordset = piNewValue
  
End Property

Public Property Get Displayed() As Boolean
  Displayed = mfDisplayed
  
End Property

Public Property Let Displayed(ByVal pfNewValue As Boolean)
  mfDisplayed = pfNewValue
  
End Property

Public Property Get ExcelFormat() As String

  If msType = sTYPECODE_COLUMN Then
    Select Case miDataType
      Case sqlNumeric, sqlInteger
        If miDecPlaces > 0 Then
          If miDecPlaces > 127 Then
            ExcelFormat = "0." & String(127, "0")
          Else
            ExcelFormat = "0." & String(miDecPlaces, "0")
          End If
        Else
          If mlngSize > 0 Then
            ExcelFormat = "0"
          Else
            ExcelFormat = "General"
          End If
        End If

      Case sqlDate
        ExcelFormat = DateFormat
        
      Case Else
        ExcelFormat = "@"
    End Select
  Else
    ExcelFormat = "General"
  End If
  
End Property

'Public Property Get DigitSeparator() As String
'  DigitSeparator = msDigitSeparator
'
'End Property
'
'Public Property Let DigitSeparator(ByVal psNewValue As String)
'  msDigitSeparator = psNewValue
'
'End Property
'
'Public Property Get DigitGrouping() As Integer
'  DigitGrouping = miDigitGrouping
'
'End Property
'
'Public Property Let DigitGrouping(ByVal piNewValue As Integer)
'  miDigitGrouping = piNewValue
'
'End Property

Public Property Get GridColumnWidth() As Long
  GridColumnWidth = mlngGridColumnWidth
  
End Property

Public Property Let GridColumnWidth(ByVal plngNewValue As Long)
  mlngGridColumnWidth = plngNewValue
  
End Property

Public Property Get ThousandSeparator() As Boolean
  ThousandSeparator = mblnThousandSeparator
End Property

Public Property Let ThousandSeparator(ByVal blnNewValue As Boolean)
  mblnThousandSeparator = blnNewValue
End Property

Public Property Get BlankIfZero() As Boolean
  BlankIfZero = mbBlankIfZero
End Property

Public Property Let BlankIfZero(ByVal pfNewValue As Boolean)
  mbBlankIfZero = pfNewValue
End Property

Public Property Get OLEType() As HRProDataMgr.OLEType
  OLEType = miOLEType
End Property

Public Property Let OLEType(ByVal piNewValue As HRProDataMgr.OLEType)
  miOLEType = piNewValue
End Property

Public Function IsOLE() As Boolean
  IsOLE = (miDataType = sqlOle) And (msType = sTYPECODE_COLUMN)
End Function

