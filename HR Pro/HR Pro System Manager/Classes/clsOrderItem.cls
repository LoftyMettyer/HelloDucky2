VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Order Item definition variables.
Private mlngColumnID As Long
Private msColumnName As String
Private msType As String
Private miSequence As Integer
Private mfAscending As Boolean


Public Property Get Ascending() As Boolean
  Ascending = mfAscending
  
End Property

Public Property Let Ascending(pfIsAscending As Boolean)
  mfAscending = pfIsAscending
  
End Property

Public Property Get ColumnID() As Long
  ' Return the column ID.
  ColumnID = mlngColumnID
  
End Property

Public Property Let ColumnID(pLngNewID As Long)
  ' Set the column ID property.
  mlngColumnID = pLngNewID
  
End Property

Public Property Get ColumnName() As String
  ColumnName = msColumnName
  
End Property

Public Function FullColumnName() As String
  ' Return the the full column name.
  ' ie. <table name>.<column name>
  Dim sSQL As String
  Dim sFullColumnName As String
  Dim rsInfo As dao.Recordset
  
  sFullColumnName = msColumnName
  
  ' Get the column's table name.
  sSQL = "SELECT tmpTables.tableName" & _
    " FROM tmpColumns, tmpTables" & _
    " WHERE tmpColumns.columnID = " & Trim(Str(mlngColumnID)) & _
    " AND tmpTables.tableID = tmpColumns.tableID"
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  If Not (rsInfo.BOF And rsInfo.EOF) Then
    sFullColumnName = rsInfo!TableName & "." & msColumnName
  End If
  rsInfo.Close
  Set rsInfo = Nothing
  
  FullColumnName = sFullColumnName
  
End Function


Public Property Let ColumnName(psName As String)
  msColumnName = psName
  
End Property

Public Property Get ItemType() As String
  ItemType = msType
  
End Property

Public Property Let ItemType(psNewType As String)
  msType = psNewType
  
End Property

Public Property Get Sequence() As Integer
  ' Return the seqeunce property.
  Sequence = miSequence
  
End Property

Public Property Let Sequence(piNewSequence As Integer)
  ' Set the sequence property.
  miSequence = piNewSequence
  
End Property


Public Function PrintOrderItem() As Boolean
  ' Print the orderitem definition.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim sOutput As String
  
  fOK = True
  
  sOutput = FullColumnName
  If ItemType = "O" Then
    sOutput = sOutput & " (" & IIf(mfAscending, "ascending", "descending") & ")"
  End If
    
  Printer.CurrentX = giPRINT_XINDENT + giPRINT_XSPACE
  Printer.CurrentY = Printer.CurrentY + giPRINT_YSPACE
  Printer.Print sOutput

TidyUpAndExit:
  PrintOrderItem = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

