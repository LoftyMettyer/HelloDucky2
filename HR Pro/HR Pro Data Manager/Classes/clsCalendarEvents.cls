VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalendarEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Hold local collection
Private mCol As Collection
Public Function Add(pstrKey As String, pstrName As String, _
                    Optional plngTableID As Long, Optional pstrTableName As String, _
                    Optional plngFilterID As Long, _
                    Optional plngStartDateID As Long, Optional pstrStartDateName As String, _
                    Optional plngStartSessionID As Long, Optional pstrStartSessionName As String, _
                    Optional plngEndDateID As Long, Optional pstrEndDateName As String, _
                    Optional plngEndSessionID As Long, Optional pstrEndSessionName As String, _
                    Optional plngDurationID As Long, Optional pstrDurationName As String, _
                    Optional pintLegendType As Integer, Optional pstrLegendCharacter As String, _
                    Optional plngLegendTableID As Long, Optional pstrLegendTableName As String, _
                    Optional plngLegendColumnID As Long, Optional pstrLegendColumnName As String, _
                    Optional plngLegendCodeID As Long, Optional pstrLegendCodeName As String, _
                    Optional plngLegendEventTypeID As Long, Optional pstrLegendEventTypeName As String, _
                    Optional plngDesc1ID As Long, Optional pstrDesc1Name As String, _
                    Optional plngDesc2ID As Long, Optional pstrDesc2Name As String, _
                    Optional pstrBaseDescription As String, _
                    Optional pstrRegion As String, Optional pstrWorkingPattern As String, _
                    Optional pstrDesc1Value As String, Optional pstrDesc2Value As String) As clsCalendarEvent
                    
  ' Add a new object to the collection
  Dim objNewMember As clsCalendarEvent
  Set objNewMember = New clsCalendarEvent
  
  With objNewMember
    .Key = pstrKey
    .Name = pstrName
    
    If Not IsMissing(plngTableID) Then .TableID = plngTableID
    If Not IsMissing(pstrTableName) Then .TableName = pstrTableName
    If Not IsMissing(plngFilterID) Then .FilterID = plngFilterID
    If Not IsMissing(plngStartDateID) Then .StartDateID = plngStartDateID
    If Not IsMissing(pstrStartDateName) Then .StartDateName = pstrStartDateName
    If Not IsMissing(plngStartSessionID) Then .StartSessionID = plngStartSessionID
    If Not IsMissing(pstrStartSessionName) Then .StartSessionName = pstrStartSessionName
    If Not IsMissing(plngEndDateID) Then .EndDateID = plngEndDateID
    If Not IsMissing(pstrEndDateName) Then .EndDateName = pstrEndDateName
    If Not IsMissing(plngEndSessionID) Then .EndSessionID = plngEndSessionID
    If Not IsMissing(pstrEndSessionName) Then .EndSessionName = pstrEndSessionName
    If Not IsMissing(plngDurationID) Then .DurationID = plngDurationID
    If Not IsMissing(pstrDurationName) Then .DurationName = pstrDurationName
    If Not IsMissing(pintLegendType) Then .LegendType = pintLegendType
    If Not IsMissing(pstrLegendCharacter) Then .LegendCharacter = pstrLegendCharacter
    If Not IsMissing(plngLegendTableID) Then .LegendTableID = plngLegendTableID
    If Not IsMissing(pstrLegendTableName) Then .LegendTableName = pstrLegendTableName
    If Not IsMissing(plngLegendColumnID) Then .LegendColumnID = plngLegendColumnID
    If Not IsMissing(pstrLegendColumnName) Then .LegendColumnName = pstrLegendColumnName
    If Not IsMissing(plngLegendCodeID) Then .LegendCodeID = plngLegendCodeID
    If Not IsMissing(pstrLegendCodeName) Then .LegendCodeName = pstrLegendCodeName
    If Not IsMissing(plngLegendEventTypeID) Then .LegendEventTypeID = plngLegendEventTypeID
    If Not IsMissing(pstrLegendEventTypeName) Then .LegendEventTypeName = pstrLegendEventTypeName
    If Not IsMissing(plngDesc1ID) Then .Description1ID = plngDesc1ID
    If Not IsMissing(pstrDesc1Name) Then .Description1Name = pstrDesc1Name
    If Not IsMissing(plngDesc2ID) Then .Description2ID = plngDesc2ID
    If Not IsMissing(pstrDesc2Name) Then .Description2Name = pstrDesc2Name
    
    '*************************************
    'optional event data used for calendar report breakdown
    If Not IsMissing(pstrBaseDescription) Then .BaseDescription = pstrBaseDescription
    If Not IsMissing(pstrRegion) Then .Region = pstrRegion
    If Not IsMissing(pstrWorkingPattern) Then .WorkingPattern = pstrWorkingPattern
    If Not IsMissing(pstrDesc1Value) Then .Desc1Value = pstrDesc1Value
    If Not IsMissing(pstrDesc2Value) Then .Desc2Value = pstrDesc2Value
    '*************************************
    
  End With
  
  mCol.Add objNewMember, pstrKey
  
  Set Add = objNewMember
  
  Set objNewMember = Nothing
                    
End Function
Public Property Get Item(pstrEventName As String) As clsCalendarEvent
  
  ' Provide a reference to a specific item in the collection
  
  On Error Resume Next
  
  Set Item = mCol(pstrEventName)
  
End Property
Public Property Get Collection() As Collection
  Set Collection = mCol
End Property
Public Property Let Collection(ByVal pcolNewValue As Collection)
  Set mCol = pcolNewValue
End Property
Public Property Get Count() As Long
  ' Provide number of objects in the collection
  Count = mCol.Count
End Property
Public Sub Remove(pstrName As String)
  ' Remove a specific object from the collection
  mCol.Remove pstrName
End Sub
Public Sub RemoveAll()
  ' Remove all object from the collection
  Do While mCol.Count > 1
    mCol.Remove mCol.Count
  Loop
End Sub
Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub
Public Property Get NewEnum() As IUnknown
  Set NewEnum = mCol.[_NewEnum]
End Property
