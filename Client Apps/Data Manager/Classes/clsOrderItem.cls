VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Order Item definition variables.
Private mlngColumnID As Long
Private msColumnName As String
Private msTableName As String
Private msType As String
Private miSequence As Integer
Private mfAscending As Boolean
Private mbIsEditable As Boolean

Public Function PrintOrderItem() As Boolean
  ' Print the Order Item definition.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim sOutput As String
  
  fOK = True
  
  sOutput = FullColumnName
  If ItemType = "O" Then
    sOutput = sOutput & " (" & IIf(Ascending, "ascending", "descending") & ")"
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

Public Property Get ColumnName() As String
  ColumnName = msColumnName
  
End Property

Public Property Let ColumnName(psName As String)
  msColumnName = psName
  
End Property

Public Property Get Ascending() As Boolean
  Ascending = mfAscending
  
End Property

Public Function FullColumnName() As String
  ' Return the the full column name.
  ' ie. <table name>.<column name>
 
  FullColumnName = msTableName & "." & msColumnName
  
End Function

Public Property Let Ascending(pfNewValue As Boolean)
  mfAscending = pfNewValue
  
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
  
End Property

Public Property Let ColumnID(pLngNewID As Long)
  mlngColumnID = pLngNewID
  
End Property

Public Property Get Sequence() As Integer
  Sequence = miSequence
  
End Property

Public Property Let Sequence(piNewSequence As Integer)
  miSequence = piNewSequence
  
End Property

Public Property Get ItemType() As String
  ItemType = msType
  
End Property

Public Property Let ItemType(psNewType As String)
  msType = psNewType
  
End Property

Public Property Get TableName() As String
  ' Return the order column's table name.
  TableName = msTableName
  
End Property

Public Property Let TableName(ByVal psNewValue As String)
  ' Set the order column's table name.
  msTableName = psNewValue

End Property

Public Property Let IsEditable(ByVal NewValue As Boolean)
  mbIsEditable = NewValue
End Property

Public Property Get IsEditable() As Boolean
  IsEditable = mbIsEditable
End Property
