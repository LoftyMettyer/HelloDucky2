VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSummaryField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Properties.
Private gLngID As Long
Private gLngHistoryTableID As Long
Private gLngSummaryColumnID As Long
Private gLngSequence As Long
Private gfStartOfGroup As Boolean
Private mbStartOfColumn As Boolean

Public Function CloneSummaryField() As cSummaryField
  ' Copy the Summary Field.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim objNewSummary As cSummaryField
  
  ' Read the current column definition.
  fOK = ReadSummaryField
  If fOK Then
    ' Instantiate a Summary Field object to be the clone.
    Set objNewSummary = New cSummaryField
    ' Copy properties from the current column object to the clone.
    With objNewSummary
      .ID = 0
      .HistoryTableID = gLngHistoryTableID
      .SummaryColumnID = gLngSummaryColumnID
      .Sequence = gLngSequence
      .StartOfGroup = gfStartOfGroup
      .StartOfColumn = mbStartOfColumn
    End With
  End If
  
TidyUpAndExit:
  If fOK Then
    Set CloneSummaryField = objNewSummary
  Else
    Set CloneSummaryField = Nothing
  End If
  ' Disassociate object variables.
  Set objNewSummary = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function FindSummaryField() As Boolean
  ' Locate the current SummaryField in the recordset.
  On Error GoTo ErrorTrap
  
  With recSummaryEdit
    .Index = "idxID"
    .Seek "=", gLngID
    FindSummaryField = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindSummaryField = False

End Function


Public Property Get StartOfGroup() As Boolean
  ' Return the StartOfGroup property.
  StartOfGroup = gfStartOfGroup

End Property

Public Property Let StartOfGroup(ByVal pfNewValue As Boolean)
  ' Set the StartOfGroup property.
  gfStartOfGroup = pfNewValue

End Property

Public Property Get HistoryTableID() As Long
  ' Return the HistoryTableID property.
  HistoryTableID = gLngHistoryTableID
  
End Property

Public Property Let HistoryTableID(ByVal pLngNewValue As Long)
  ' Set the HistoryTableID property.
  gLngHistoryTableID = pLngNewValue

End Property
Public Property Get ID() As Long
  ' Return the ID property.
  ID = gLngID
  
End Property

Public Property Let ID(ByVal pLngNewValue As Long)
  ' Set the ID property.
  gLngID = pLngNewValue

End Property

Public Function ReadSummaryField() As Boolean
  ' Read the SummaryField definition from the Summary Field recordset.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean
  
  fReadOk = False
  
  ' Read the basic component details.
  If FindSummaryField Then
    
    With recSummaryEdit
      
      gLngID = IIf(IsNull(.Fields("ID")), 0, .Fields("ID"))
      gLngHistoryTableID = IIf(IsNull(.Fields("HistoryTableID")), 0, .Fields("HistoryTableID"))
      gLngSummaryColumnID = IIf(IsNull(.Fields("ParentColumnID")), 0, .Fields("ParentColumnID"))
      gLngSequence = IIf(IsNull(.Fields("Sequence")), 0, .Fields("Sequence"))
      gfStartOfGroup = IIf(IIf(IsNull(.Fields("StartOfGroup")), False, .Fields("StartOfGroup")), True, False)
      mbStartOfColumn = IIf(IIf(IsNull(.Fields("StartOfColumn")), False, .Fields("StartOfColumn")), True, False)
      
      fReadOk = True
  
    End With
  End If
   
TidyUpAndExit:
  ReadSummaryField = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function
Public Function WriteSummaryField() As Boolean
  ' Write the SummaryField to the recordset.
  On Error GoTo ErrorTrap
  
  Dim lngNewID As Long
  Dim fOK As Boolean
  
  fOK = True
  
  lngNewID = Database.UniqueColumnValue("tmpSummary", "ID")
  gLngID = lngNewID
  
  With recSummaryEdit
  
    'Add a record for new expression
    .AddNew
    
    !ID = gLngID
    !HistoryTableID = gLngHistoryTableID
    !parentColumnID = gLngSummaryColumnID
    !Sequence = gLngSequence
    !StartOfGroup = gfStartOfGroup
    !StartOfColumn = mbStartOfColumn
    
    .Update
    
  End With
  
TidyUpAndExit:
  WriteSummaryField = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function



Public Property Get SummaryColumnID() As Long
  ' Return the SummaryColumnID property.
  SummaryColumnID = gLngSummaryColumnID

End Property

Public Property Let SummaryColumnID(ByVal pLngNewValue As Long)
  ' Set the SummaryColumnID property.
  gLngSummaryColumnID = pLngNewValue

End Property

Public Property Get Sequence() As Long
  ' Return the Sequence property.
  Sequence = gLngSequence
  
End Property

Public Property Let Sequence(ByVal pLngNewValue As Long)
  ' Set the Sequence property.
  gLngSequence = pLngNewValue

End Property

Public Property Get StartOfColumn() As Boolean
  ' Return the StartOfColumn property.
  StartOfColumn = mbStartOfColumn

End Property

Public Property Let StartOfColumn(ByVal pbNewValue As Boolean)
  ' Set the StartOfColumn property.
  mbStartOfColumn = pbNewValue

End Property
