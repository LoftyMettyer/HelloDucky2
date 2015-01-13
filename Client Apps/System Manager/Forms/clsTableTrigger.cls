VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTableTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TriggerID As Long
Public TableID As Long
Public Name As String
Public IsSystem As Boolean
Public CodePosition As TriggerCodePosition
Public content As String

Private mbIsNew As Boolean
Private mbIsChanged As Boolean
Private mbIsDeleted As Boolean

Public Property Get Deleted() As Boolean
  Deleted = mbIsDeleted
End Property

Public Property Let Deleted(ByVal NewValue As Boolean)
  mbIsDeleted = NewValue
End Property

Public Function ReadTrigger() As Boolean
  On Error GoTo ErrorTrap

  Dim fReadOk As Boolean

  fReadOk = True

  With recTableTriggerEdit
  
    .Index = "idxTriggerID"
    .Seek "=", TriggerID
        
    If Not .NoMatch Then
      TriggerID = .Fields("TriggerID").value
      TableID = .Fields("TableID").value
      Name = .Fields("Name").value
      content = .Fields("Content").value
      IsSystem = .Fields("IsSystem").value
      CodePosition = .Fields("CodePosition").value
    Else
      fReadOk = False
    End If
  End With
 
TidyUpAndExit:
  ReadTrigger = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function


Public Function WriteTrigger() As Boolean

  On Error GoTo ErrorTrap

  Dim fReadOk As Boolean

  fReadOk = True

  With recTableTriggerEdit
    If TriggerID = 0 Then
      TriggerID = UniqueColumnValue("tmpTableTriggers", "TriggerID")
    End If
  
    .AddNew
        
    .Fields("TriggerID").value = TriggerID
    .Fields("TableID").value = TableID

    .Fields("Name").value = Name
    .Fields("Content").value = content
    .Fields("IsSystem").value = IsSystem
    .Fields("CodePosition").value = CodePosition
  
    .Fields("New").value = mbIsNew
    .Fields("Changed").value = mbIsChanged
    .Fields("Deleted").value = mbIsDeleted
  
    .Update
  
  End With
 
TidyUpAndExit:
  WriteTrigger = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function

