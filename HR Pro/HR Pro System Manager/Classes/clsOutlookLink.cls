VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutlookLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngLinkID As Long
Private mlngTableID As Long
Private mstrTitle As String
Private mlngFilterID As Long
Private mlngSubject As Long
Private mlngBusyStatus As Long
Private mlngStartDate As Long
Private mlngEndDate As Long
Private mlngTimeRange As Long
Private mstrFixedStartTime As String
Private mstrFixedEndTime As String
Private mlngColumnStartTime As Long
Private mlngColumnEndTime As Long
Private mstrContent As String
Private mblnReminder As Boolean
Private mlngReminderOffset As Long
Private mlngReminderPeriod As Long

Private mcolLinkColumns As Collection
Private mlngDestinations() As Long

Private mfNew As Boolean
Private mfChanged As Boolean
Private mfDeleted As Boolean


Public Sub ClearColumns()

  Set mcolLinkColumns = Nothing
  Set mcolLinkColumns = New Collection

End Sub


Public Sub AddColumn(lngColumnID As Long, strHeading As String, lngSequence As Long)

  Dim objLinkColumn As clsOutlookLinkColumn

  Set objLinkColumn = New clsOutlookLinkColumn
  objLinkColumn.ColumnID = lngColumnID
  objLinkColumn.Heading = strHeading
  objLinkColumn.Sequence = lngSequence

  mcolLinkColumns.Add objLinkColumn, "C" & CStr(lngColumnID)

  Set objLinkColumn = Nothing

End Sub


Public Sub ClearDestinations()

  ReDim mlngDestinations(0) As Long

End Sub

Public Sub AddDestination(lngDestination As Long)

  Dim lngIndex As Long

  lngIndex = UBound(mlngDestinations) + 1
  ReDim Preserve mlngDestinations(lngIndex)
  mlngDestinations(lngIndex) = lngDestination

End Sub


Public Property Get LinkID() As Long
  LinkID = mlngLinkID
End Property

Public Property Let LinkID(ByVal lngNewValue As Long)
  mlngLinkID = lngNewValue
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
End Property

Public Property Get Title() As String
  Title = mstrTitle
End Property

Public Property Let Title(ByVal strNewValue As String)
  mstrTitle = strNewValue
End Property

Public Property Get FilterID() As Long
  FilterID = mlngFilterID
End Property

Public Property Let FilterID(ByVal lngNewValue As Long)
  mlngFilterID = lngNewValue
End Property

Public Property Get Subject() As Long
  Subject = mlngSubject
End Property

Public Property Let Subject(ByVal lngNewValue As Long)
  mlngSubject = lngNewValue
End Property

Public Property Get BusyStatus() As Long
  BusyStatus = mlngBusyStatus
End Property

Public Property Let BusyStatus(ByVal lngNewValue As Long)
  mlngBusyStatus = lngNewValue
End Property

Public Property Get StartDate() As Long
  StartDate = mlngStartDate
End Property

Public Property Let StartDate(ByVal lngNewValue As Long)
  mlngStartDate = lngNewValue
End Property

Public Property Get EndDate() As Long
  EndDate = mlngEndDate
End Property

Public Property Let EndDate(ByVal lngNewValue As Long)
  mlngEndDate = lngNewValue
End Property

Public Property Get TimeRange() As Long
  TimeRange = mlngTimeRange
End Property

Public Property Let TimeRange(ByVal lngNewValue As Long)
  mlngTimeRange = lngNewValue
End Property

Public Property Get FixedStartTime() As String
  FixedStartTime = mstrFixedStartTime
End Property

Public Property Let FixedStartTime(ByVal strNewValue As String)
  mstrFixedStartTime = strNewValue
End Property

Public Property Get FixedEndTime() As String
  FixedEndTime = mstrFixedEndTime
End Property

Public Property Let FixedEndTime(ByVal strNewValue As String)
  mstrFixedEndTime = strNewValue
End Property

Public Property Get ColumnStartTime() As Long
  ColumnStartTime = mlngColumnStartTime
End Property

Public Property Let ColumnStartTime(ByVal lngNewValue As Long)
  mlngColumnStartTime = lngNewValue
End Property

Public Property Get ColumnEndTime() As Long
  ColumnEndTime = mlngColumnEndTime
End Property

Public Property Let ColumnEndTime(ByVal lngNewValue As Long)
  mlngColumnEndTime = lngNewValue
End Property

Public Property Get Content() As String
  Content = mstrContent
End Property

Public Property Let Content(ByVal strNewValue As String)
  mstrContent = strNewValue
End Property

Public Property Get Reminder() As Boolean
  Reminder = mblnReminder
End Property

Public Property Let Reminder(ByVal blnNewValue As Boolean)
  mblnReminder = blnNewValue
End Property

Public Property Get ReminderOffset() As Long
  ReminderOffset = mlngReminderOffset
End Property

Public Property Let ReminderOffset(ByVal lngNewValue As Long)
  mlngReminderOffset = lngNewValue
End Property

Public Property Get ReminderPeriod() As Long
  ReminderPeriod = mlngReminderPeriod
End Property

Public Property Let ReminderPeriod(ByVal lngNewValue As Long)
  mlngReminderPeriod = lngNewValue
End Property



Public Function WriteLink() As Boolean
  
  On Error GoTo 0 'ErrorTrap
  
  Dim objLinkColumn As clsOutlookLinkColumn
  Dim lngNewID As Long
  Dim fOK As Boolean
  Dim lngCount As Long
  
  fOK = True
  
  With recOutlookLinks

    If mlngLinkID = 0 Then
      mlngLinkID = Database.UniqueColumnValue("tmpOutlookLinks", "linkID")
    End If

    .AddNew

    !LinkID = mlngLinkID
    !TableID = mlngTableID
    !Title = mstrTitle
    !FilterID = mlngFilterID
    !BusyStatus = mlngBusyStatus
    !StartDate = mlngStartDate
    !EndDate = mlngEndDate
    !TimeRange = mlngTimeRange
    !FixedStartTime = mstrFixedStartTime
    !FixedEndTime = mstrFixedEndTime
    !ColumnStartTime = mlngColumnStartTime
    !ColumnEndTime = mlngColumnEndTime
    !Reminder = mblnReminder
    !ReminderOffset = mlngReminderOffset
    !ReminderPeriod = mlngReminderPeriod

    !Subject = mlngSubject
    !Content = mstrContent

    !New = mfNew
    !Changed = mfChanged
    !Deleted = mfDeleted

    .Update
    
  End With

  
  With recOutlookLinksDestinations
    '.MoveFirst
    .Index = "idxLinkID"
    .Seek ">=", mlngLinkID

    If Not .NoMatch Then
      Do While Not .EOF
        If !LinkID <> mlngLinkID Then
          Exit Do
        End If
        .Delete
        .MoveNext
      Loop
    End If

    For lngCount = 1 To UBound(mlngDestinations)
      .AddNew
      !LinkID = mlngLinkID
      !FolderID = mlngDestinations(lngCount)
      .Update
    Next

  End With


  With recOutlookLinksColumns
    '.MoveFirst
    .Index = "idxLinkSeqID"
    .Seek ">=", mlngLinkID, 0

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
      !Heading = objLinkColumn.Heading
      !Sequence = objLinkColumn.Sequence
      
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
  
  With recOutlookLinks
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

Public Function ReadLink() As Boolean
  ' Read the Email Link definition from the Email Link recordset.
  On Error GoTo ErrorTrap

  Dim fReadOk As Boolean

  fReadOk = False

  With recOutlookLinks

    mlngLinkID = !LinkID
    mlngTableID = !TableID
    mstrTitle = !Title
    mlngFilterID = !FilterID
    mlngSubject = !Subject
    mlngBusyStatus = !BusyStatus
    mlngStartDate = !StartDate
    mlngEndDate = !EndDate
    mlngTimeRange = !TimeRange
    mstrFixedStartTime = !FixedStartTime
    mstrFixedEndTime = !FixedEndTime
    mlngColumnStartTime = !ColumnStartTime
    mlngColumnEndTime = !ColumnEndTime
    mstrContent = IIf(IsNull(!Content), vbNullString, !Content)
    mblnReminder = !Reminder
    mlngReminderOffset = !ReminderOffset
    mlngReminderPeriod = !ReminderPeriod

  End With


  With recOutlookLinksDestinations
    '.MoveFirst
    .Index = "idxLinkID"
    .Seek ">=", mlngLinkID

    ClearDestinations

    If Not .NoMatch Then
      Do While Not .EOF

        If !LinkID <> mlngLinkID Then
          Exit Do
        End If

        AddDestination !FolderID
        .MoveNext
      Loop
    End If

  End With


  With recOutlookLinksColumns
    '.MoveFirst
    .Index = "idxLinkSeqID"
    .Seek ">=", mlngLinkID, 0

    ClearColumns

    If Not .NoMatch Then
      Do While Not .EOF

        If !LinkID <> mlngLinkID Then
          Exit Do
        End If

        AddColumn !ColumnID, !Heading, !Sequence
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

Public Property Get Destinations() As Variant
  Destinations = mlngDestinations
End Property

Public Property Let Destinations(ByVal vNewValue As Variant)
  mlngDestinations = vNewValue
End Property

Public Property Get LinkColumns() As Collection
  Set LinkColumns = mcolLinkColumns
End Property

Public Property Let LinkColumns(ByVal colNewValue As Collection)
  Set mcolLinkColumns = colNewValue
End Property

Private Sub Class_Initialize()
  ClearColumns
  ClearDestinations
End Sub

Public Property Get Deleted() As Boolean
  Deleted = mfDeleted
End Property

Public Property Let Deleted(ByVal blnNewValue As Boolean)
  mfDeleted = blnNewValue
End Property

Public Function CloneOutlookLink(pavCloneRegister As Variant) As clsOutlookLink
  
  ' Clone the current Outlook.
  ' Used when copying table definitions.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim lngNewColumnID As Long
  Dim objClonedOutlook As clsOutlookLink
  Dim lngNewExprID As Long
  Dim lngFolderID As Long
  
  ' Read the Outlook definition.
  fOK = FindLink
  If fOK Then
    fOK = ReadLink
  End If

  If fOK Then
    ' Instantiate the clone Outlook.
    Set objClonedOutlook = New clsOutlookLink
    ' Clone the current Outlook's properties.
    With objClonedOutlook
      '.IsNew = True
      '.IsDeleted = False
      '.IsChanged = False

      .LinkID = 0
      .Title = Me.Title
      .BusyStatus = Me.BusyStatus
      .TimeRange = Me.TimeRange
      .FixedStartTime = Me.FixedStartTime
      .FixedEndTime = Me.FixedEndTime
      .Content = Me.Content
      .Reminder = Me.Reminder
      .ReminderOffset = Me.ReminderOffset
      .ReminderPeriod = Me.ReminderPeriod

      .TableID = GetCloneID(pavCloneRegister, "TABLE", Me.TableID)
      .FilterID = GetCloneID(pavCloneRegister, "EXPRESSION", Me.FilterID)
      .Subject = GetCloneID(pavCloneRegister, "EXPRESSION", Me.Subject)
      .StartDate = GetCloneID(pavCloneRegister, "COLUMN", Me.StartDate)
      .EndDate = GetCloneID(pavCloneRegister, "COLUMN", Me.EndDate)
      .ColumnStartTime = GetCloneID(pavCloneRegister, "COLUMN", Me.ColumnStartTime)
      .ColumnEndTime = GetCloneID(pavCloneRegister, "COLUMN", Me.ColumnEndTime)


      With recOutlookLinksDestinations
        .Index = "idxLinkID"
        .Seek ">=", Me.LinkID
    
        If Not .NoMatch Then
          Do While Not .EOF
            If !LinkID <> Me.LinkID Then
              Exit Do
            End If
            lngFolderID = GetCloneID(pavCloneRegister, "OUTLOOKFOLDER", !FolderID)
            If lngFolderID = 0 Then
              'Must be a fixed folder so leave it as it is
              lngFolderID = !FolderID
            End If
            objClonedOutlook.AddDestination lngFolderID
            .MoveNext
          Loop
        End If

      End With


      With recOutlookLinksColumns
        .Index = "idxLinkSeqID"
        .Seek ">=", Me.LinkID, 0
    
        If Not .NoMatch Then
          Do While Not .EOF
            If !LinkID <> Me.LinkID Then
              Exit Do
            End If
            objClonedOutlook.AddColumn GetCloneID(pavCloneRegister, "COLUMN", !ColumnID), !Heading, !Sequence
            .MoveNext
          Loop
        End If

      End With

    End With
  End If



TidyUpAndExit:
  If fOK Then
    Set CloneOutlookLink = objClonedOutlook
  Else
    Set CloneOutlookLink = Nothing
  End If
  ' Disassociate object variables.
  Set objClonedOutlook = Nothing
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
