VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTableValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngValidationID As Long
Private mlngTableID As Long
Private miValidationType As ValidationType
Private mlngEventStartDateColumnID  As Long
Private mlngEventStartSessionColumnID  As Long
Private mlngEventEndDateColumnID  As Long
Private mlngEventEndSessionColumnID  As Long
Private mlngEventTypeColumnID As Long
Private mlngFilterID As Long
Private miSeverity As Severity
Private mstrMessage As String

Private mbIsNew As Boolean
Private mbIsChanged As Boolean
Private mbIsDeleted As Boolean

Public Property Get ValidationID() As Long
  ValidationID = mlngValidationID
End Property

Public Property Let ValidationID(ByVal NewValue As Long)
  mlngValidationID = NewValue
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let TableID(ByVal NewValue As Long)
  mlngTableID = NewValue
End Property

Public Property Get ValidationType() As ValidationType
  ValidationType = miValidationType
End Property

Public Property Let ValidationType(ByVal NewValue As ValidationType)
  miValidationType = NewValue
End Property

Public Property Get EventStartdateColumnID() As Long
  EventStartdateColumnID = mlngEventStartDateColumnID
End Property

Public Property Let EventStartdateColumnID(ByVal NewValue As Long)
  mlngEventStartDateColumnID = NewValue
End Property

Public Property Get EventStartSessionColumnID() As Long
  EventStartSessionColumnID = mlngEventStartSessionColumnID
End Property

Public Property Let EventStartSessionColumnID(ByVal NewValue As Long)
  mlngEventStartSessionColumnID = NewValue
End Property

Public Property Get EventEnddateColumnID() As Long
  EventEnddateColumnID = mlngEventEndDateColumnID
End Property

Public Property Let EventEnddateColumnID(ByVal NewValue As Long)
  mlngEventEndDateColumnID = NewValue
End Property

Public Property Get EventEndSessionColumnID() As Long
  EventEndSessionColumnID = mlngEventEndSessionColumnID
End Property

Public Property Let EventEndSessionColumnID(ByVal NewValue As Long)
  mlngEventEndSessionColumnID = NewValue
End Property

Public Property Get FilterID() As Long
  FilterID = mlngFilterID
End Property

Public Property Let FilterID(ByVal NewValue As Long)
  mlngFilterID = NewValue
End Property

Public Property Get Severity() As Severity
  Severity = miSeverity
End Property

Public Property Let Severity(ByVal NewValue As Severity)
  miSeverity = NewValue
End Property

Public Property Get EventTypeColumnID() As Long
  EventTypeColumnID = mlngEventTypeColumnID
End Property

Public Property Let EventTypeColumnID(ByVal NewValue As Long)
  mlngEventTypeColumnID = NewValue
End Property

Public Property Get Message() As String
  'If mstrMessage = vbNullString Then
   ' Message = GetColumnName(mlngEventStartDateColumnID, True) & " to " & GetColumnName(mlngEventEndDateColumnID, True) & " overlaps with another record."
  'Else
    Message = mstrMessage
  'End If
End Property

Public Property Let Message(ByVal NewValue As String)
  mstrMessage = NewValue
End Property

Public Property Get Deleted() As Boolean
  Deleted = mbIsDeleted
End Property

Public Property Let Deleted(ByVal NewValue As Boolean)
  mbIsDeleted = NewValue
End Property


Public Function ReadValidation() As Boolean
  On Error GoTo ErrorTrap

  Dim fReadOk As Boolean

  fReadOk = True

  With recTableValidationEdit
  
    .Index = "idxValidationID"
    .Seek "=", mlngValidationID
        
    If Not .NoMatch Then
      mlngValidationID = .Fields("ValidationID").value
      mlngTableID = .Fields("TableID").value
      miValidationType = .Fields("Type").value
      mlngEventStartDateColumnID = .Fields("EventStartDateColumnID").value
      mlngEventStartSessionColumnID = .Fields("EventStartSessionColumnID").value
      mlngEventEndDateColumnID = .Fields("EventEndDateColumnID").value
      mlngEventEndSessionColumnID = .Fields("EventEndSessionColumnID").value
      mlngEventTypeColumnID = .Fields("EventTypeColumnID").value
      mlngFilterID = .Fields("FilterID").value
      miSeverity = .Fields("Severity").value
      mstrMessage = .Fields("Message").value
    Else
      fReadOk = False
    End If
  End With
 
TidyUpAndExit:
  ReadValidation = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function

Public Function WriteValidation() As Boolean

  On Error GoTo ErrorTrap

  Dim fReadOk As Boolean

  fReadOk = True

  With recTableValidationEdit
    If mlngValidationID = 0 Then
      mlngValidationID = Database.UniqueColumnValue("tmpTableValidations", "ValidationID")
    End If
  
    .AddNew
        
    .Fields("ValidationID").value = mlngValidationID
    .Fields("TableID").value = mlngTableID
    .Fields("Type").value = miValidationType
    .Fields("EventStartDateColumnID").value = mlngEventStartDateColumnID
    .Fields("EventStartSessionColumnID").value = mlngEventStartSessionColumnID
    .Fields("EventEndDateColumnID").value = mlngEventEndDateColumnID
    .Fields("EventEndSessionColumnID").value = mlngEventEndSessionColumnID
    .Fields("EventTypeColumnID").value = mlngEventTypeColumnID
    .Fields("FilterID").value = mlngFilterID
    .Fields("Severity").value = miSeverity
    .Fields("Message").value = mstrMessage
  
    .Fields("New").value = mbIsNew
    .Fields("Changed").value = mbIsChanged
    .Fields("Deleted").value = mbIsDeleted
  
    .Update
  
  End With
 
TidyUpAndExit:
  WriteValidation = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function

