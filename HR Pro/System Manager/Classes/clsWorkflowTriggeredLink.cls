VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkflowTriggeredLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngLinkID As Long
Private mlngWorkflowID As Long
Private mlngTableID As Long
Private mlngFilterID As Long
Private mdtEffectiveDate As Date
Private miType As WorkflowTriggerLinkType
Private mfRecordInsert As Boolean
Private mfRecordUpdate As Boolean
Private mfRecordDelete As Boolean
Private mlngDateColumnID As Long
Private mlngDateOffset As Long
Private miDateOffsetPeriod As WorkflowTriggerOffsetPeriod

Private mcolLinkColumns As Collection

Private mfNew As Boolean
Private mfChanged As Boolean
Private mfDeleted As Boolean

Public Enum WorkflowTriggerLinkType
  WORKFLOWTRIGGERLINKTYPE_COLUMN = 0
  WORKFLOWTRIGGERLINKTYPE_RECORD = 1
  WORKFLOWTRIGGERLINKTYPE_DATE = 2
End Enum

Public Enum WorkflowTriggerOffsetPeriod
  WORKFLOWTRIGGERLINKOFFESTPERIOD_DAY = 0
  WORKFLOWTRIGGERLINKOFFESTPERIOD_WEEK = 1
  WORKFLOWTRIGGERLINKOFFESTPERIOD_MONTH = 2
  WORKFLOWTRIGGERLINKOFFESTPERIOD_YEAR = 3
End Enum

Public Property Get Deleted() As Boolean
  Deleted = mfDeleted
  
End Property


Public Property Let Deleted(ByVal blnNewValue As Boolean)
  mfDeleted = blnNewValue
  
End Property


Public Property Get LinkColumns() As Collection
  Set LinkColumns = mcolLinkColumns
  
End Property


Public Property Let LinkColumns(ByVal colNewValue As Collection)
  Set mcolLinkColumns = colNewValue
  
End Property


Public Property Get FilterID() As Long
  FilterID = mlngFilterID
  
End Property


Public Property Get EffectiveDate() As Date
  EffectiveDate = mdtEffectiveDate
  
End Property



Public Property Let FilterID(ByVal lngNewValue As Long)
  mlngFilterID = lngNewValue
  
End Property


Public Property Let LinkType(ByVal piNewValue As WorkflowTriggerLinkType)
  miType = piNewValue
  
End Property



Public Property Let EffectiveDate(ByVal pdtNewValue As Date)
  mdtEffectiveDate = pdtNewValue
  
End Property




Public Property Get TableID() As Long
  TableID = mlngTableID
  
End Property

Public Property Get WorkflowID() As Long
  WorkflowID = mlngWorkflowID
  
End Property
Public Property Let WorkflowID(ByVal lngNewValue As Long)
  mlngWorkflowID = lngNewValue
  
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
  
End Property


Public Property Get LinkID() As Long
  LinkID = mlngLinkID
  
End Property


Public Property Get LinkType() As WorkflowTriggerLinkType
  LinkType = miType
  
End Property



Public Property Let LinkID(ByVal lngNewValue As Long)
  mlngLinkID = lngNewValue
  
End Property


Public Sub AddColumn(lngColumnID As Long)

  Dim objLinkColumn As clsWFTriggeredLinkColumn

  Set objLinkColumn = New clsWFTriggeredLinkColumn
  objLinkColumn.ColumnID = lngColumnID

  mcolLinkColumns.Add objLinkColumn, "C" & CStr(lngColumnID)

  Set objLinkColumn = Nothing

End Sub
Private Sub Class_Initialize()
  ClearColumns

End Sub


Public Sub ClearColumns()

  Set mcolLinkColumns = Nothing
  Set mcolLinkColumns = New Collection

End Sub



Public Function CloneLink(pavCloneRegister As Variant) As clsWorkflowTriggeredLink
  
  ' Clone the current Workflow Triggered Link.
  ' Used when copying table definitions.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim objClonedLink As clsWorkflowTriggeredLink

  ' Read the Link definition.
  fOK = FindLink
  If fOK Then
    fOK = ReadLink
  End If

  If fOK Then
    ' Instantiate the clone link.
    Set objClonedLink = New clsWorkflowTriggeredLink
    ' Clone the current Link's properties.
    With objClonedLink
      '.IsNew = True
      '.IsDeleted = False
      '.IsChanged = False

      .LinkID = 0
      .TableID = GetCloneID(pavCloneRegister, "TABLE", Me.TableID)
      .WorkflowID = GetCloneID(pavCloneRegister, "WORKFLOW", Me.WorkflowID)
      .FilterID = GetCloneID(pavCloneRegister, "EXPRESSION", Me.FilterID)
      .EffectiveDate = Me.EffectiveDate
      .LinkType = Me.LinkType
      .RecordInsert = Me.RecordInsert
      .RecordUpdate = Me.RecordUpdate
      .RecordDelete = Me.RecordDelete
      .DateColumnID = GetCloneID(pavCloneRegister, "COLUMN", Me.DateColumnID)
      .DateOffset = Me.DateOffset
      .DateOffsetPeriod = Me.DateOffsetPeriod

      With recWorkflowTriggeredLinkColumns
        .Index = "idxLinkID"
        .Seek "=", Me.LinkID

        If Not .NoMatch Then
          Do While Not .EOF
            If !LinkID <> Me.LinkID Then
              Exit Do
            End If
            objClonedLink.AddColumn GetCloneID(pavCloneRegister, "COLUMN", !ColumnID)
            .MoveNext
          Loop
        End If
      End With
    End With
  End If

TidyUpAndExit:
  If fOK Then
    Set CloneLink = objClonedLink
  Else
    Set CloneLink = Nothing
  End If
  ' Disassociate object variables.
  Set objClonedLink = Nothing
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function GetCloneID(pavCloneRegister As Variant, strType As String, lngOriginalID As Long) As Long
      
  Dim lngCloneID As Long
  Dim iIndex As Integer
      
  lngCloneID = 0
  If lngOriginalID > 0 Then
    For iIndex = 1 To UBound(pavCloneRegister, 2)
      If pavCloneRegister(1, iIndex) = strType And _
        pavCloneRegister(2, iIndex) = lngOriginalID Then
        lngCloneID = pavCloneRegister(3, iIndex)
        Exit For
      End If
    Next iIndex
  End If

  GetCloneID = lngCloneID

End Function

Public Function ReadLink() As Boolean
  ' Read the WorkflowTriggeredLink definition from the WorkflowTriggeredLink recordset.
  On Error GoTo ErrorTrap

  Dim fReadOk As Boolean

  fReadOk = False

  With recWorkflowTriggeredLinks
    mlngLinkID = !LinkID
    mlngTableID = !TableID
    mlngWorkflowID = !WorkflowID
    mlngFilterID = !FilterID
    mdtEffectiveDate = !EffectiveDate
    miType = !Type
    mfRecordInsert = !RecordInsert
    mfRecordUpdate = !RecordUpdate
    mfRecordDelete = !RecordDelete
    mlngDateColumnID = !DateColumn
    mlngDateOffset = !DateOffset
    miDateOffsetPeriod = !DateOffsetPeriod
  End With

  With recWorkflowTriggeredLinkColumns
    '.MoveFirst
    .Index = "idxLinkID"
    .Seek "=", mlngLinkID

    ClearColumns

    If Not .NoMatch Then
      Do While Not .EOF
        If !LinkID <> mlngLinkID Then
          Exit Do
        End If

        AddColumn !ColumnID
        .MoveNext
      Loop
    End If
  End With

  fReadOk = True
  
TidyUpAndExit:
  ReadLink = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function


Public Function WriteLink() As Boolean
  
  On Error GoTo 0 'ErrorTrap
  
  Dim objLinkColumn As clsWFTriggeredLinkColumn
  Dim fOK As Boolean

  fOK = True

  With recWorkflowTriggeredLinks
    If mlngLinkID = 0 Then
      mlngLinkID = Database.UniqueColumnValue("tmpWorkflowTriggeredLinks", "linkID")
    End If

    .AddNew

    !LinkID = mlngLinkID
    !TableID = mlngTableID
    !WorkflowID = mlngWorkflowID
    !FilterID = mlngFilterID
    !EffectiveDate = mdtEffectiveDate
    !Type = miType
    !RecordInsert = mfRecordInsert
    !RecordUpdate = mfRecordUpdate
    !RecordDelete = mfRecordDelete
    !DateColumn = mlngDateColumnID
    !DateOffset = mlngDateOffset
    !DateOffsetPeriod = miDateOffsetPeriod

    !New = mfNew
    !Changed = mfChanged
    !Deleted = mfDeleted

    .Update

  End With

  With recWorkflowTriggeredLinkColumns
    '.MoveFirst
    .Index = "idxLinkID"
    .Seek "=", mlngLinkID

    If Not .NoMatch Then
      Do While Not .EOF
        If !LinkID <> mlngLinkID Then
          Exit Do
        End If
        .Delete
        .MoveNext
      Loop
    End If

    For Each objLinkColumn In mcolLinkColumns
      .AddNew
      !LinkID = mlngLinkID
      !ColumnID = objLinkColumn.ColumnID

      '!New = True
      '!Changed = True
      '!Deleted = False

      .Update
    Next
  End With

TidyUpAndExit:
  WriteLink = fOK
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
    Stop
  End If
  
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function



Public Function FindLink() As Boolean
  On Error GoTo ErrorTrap
  
  With recWorkflowTriggeredLinks
    .MoveFirst
    Do While Not .EOF
      If !LinkID = mlngLinkID Then
        Exit Do
      End If
      .MoveNext
    Loop
    FindLink = (Not .EOF)
  End With
  
  Exit Function
  
ErrorTrap:
  FindLink = False

End Function



Public Property Get RecordInsert() As Boolean
  RecordInsert = mfRecordInsert
  
End Property

Public Property Let RecordInsert(ByVal pfNewValue As Boolean)
  mfRecordInsert = pfNewValue
  
End Property

Public Property Get RecordUpdate() As Boolean
  RecordUpdate = mfRecordUpdate
  
End Property

Public Property Let RecordUpdate(ByVal pfNewValue As Boolean)
  mfRecordUpdate = pfNewValue
  
End Property

Public Property Get RecordDelete() As Boolean
  RecordDelete = mfRecordDelete
  
End Property

Public Property Let RecordDelete(ByVal pfNewValue As Boolean)
  mfRecordDelete = pfNewValue
  
End Property

Public Property Get DateColumnID() As Long
  DateColumnID = mlngDateColumnID
  
End Property

Public Property Let DateColumnID(ByVal plngNewValue As Long)
  mlngDateColumnID = plngNewValue
  
End Property

Public Property Get DateOffset() As Long
  DateOffset = mlngDateOffset
  
End Property

Public Property Let DateOffset(ByVal plngNewValue As Long)
  mlngDateOffset = plngNewValue
  
End Property

Public Property Get DateOffsetPeriod() As WorkflowTriggerOffsetPeriod
  DateOffsetPeriod = miDateOffsetPeriod
  
End Property

Public Property Let DateOffsetPeriod(ByVal piNewValue As WorkflowTriggerOffsetPeriod)
  miDateOffsetPeriod = piNewValue
  
End Property
