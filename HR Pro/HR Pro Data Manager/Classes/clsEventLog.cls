VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mclsData As New clsDataAccess
Private mlngEventLogID As Long
Private mstrBatchName As String
Private mlngBatchRunID As Long
Private mlngBatchJobID As Long
'Private gblnBatchMode As Boolean

'Public Property Let BatchMode(pblnNewValue As Boolean)
'  gblnBatchMode = pblnNewValue
'  mlngBatchRunID = GetBatchRunID
'End Property

Public Property Let BatchJobName(ByVal pstrBatchJobName As String)
  ' When the batch job name is passed into the class module, the BatchRunID is
  ' allocated automatically. Process is complete hidden from the utilities.
  mstrBatchName = pstrBatchJobName
End Property

Public Property Let BatchJobID(plngBatchJobID As Long)
  mlngBatchJobID = plngBatchJobID
End Property

Private Sub Class_Initialize()
  Set mclsData = New clsDataAccess
End Sub

Private Sub Class_Terminate()
  Set mclsData = Nothing
End Sub

Public Sub TidyUp()

  ' Must be called from batchjobs when the batch has finished, otherwise all
  ' subsequent functions in the users current session will be logged as being
  ' in the batch job
  mstrBatchName = vbNullString
  mlngBatchRunID = 0
  mlngBatchJobID = 0
  
End Sub

Public Function AddHeader(udtType As EventLog_Type, _
                          strName As String) As Boolean

  On Local Error GoTo ErrTrap

  Dim fInTransaction As Boolean
  Dim lngTimeOut As Long
  
  lngTimeOut = Timer + 5
  
  mlngEventLogID = 0
  fInTransaction = False
  
  ' If we are in batchmode, is everything ok ?
  If (mlngBatchRunID = -1) And (mstrBatchName <> vbNullString) And (mlngBatchJobID > 0) Then
    AddHeader = False
    Exit Function
  End If

  ' Temp recordset to hold the returned id of the record just added
  Dim prstRowAdded As ADODB.Recordset

  ' Start a transaction
  gADOCon.BeginTrans
  fInTransaction = True

  ' Do the insert
  gADOCon.Execute "INSERT INTO AsrSysEventLog (" & _
                  "DateTime," & _
                  "Type," & _
                  "Name," & _
                  "Status," & _
                  "Username," & _
                  "Mode," & _
                  "BatchName," & _
                  "SuccessCount," & _
                  "FailCount," & _
                  "BatchRunID," & _
                  "BatchJobID) " & _
                  "VALUES(" & _
                  "GETDATE()," & _
                   udtType & "," & _
                  "'" & Replace(strName, "'", "''") & "'," & _
                  elsPending & "," & _
                  "'" & datGeneral.UserNameForSQL & "'," & _
                  IIf(mstrBatchName = vbNullString, 0, 1) & "," & _
                  "'" & Replace(mstrBatchName, "'", "''") & "'," & _
                  "NULL," & _
                  "NULL," & _
                  IIf(mlngBatchRunID > 0, mlngBatchRunID, "NULL") & "," & _
                  IIf(mlngBatchJobID > 0, mlngBatchJobID, "NULL") & ")"

  ' Retrieve the id just added
  Set prstRowAdded = gADOCon.Execute("SELECT MAX(ID) FROM AsrSysEventLog")

  ' Set function return to the id just added
  mlngEventLogID = prstRowAdded.Fields(0)

  ' Commit the transaction
  gADOCon.CommitTrans
  fInTransaction = False

  ' Tidy up
  Set prstRowAdded = Nothing

  AddHeader = True

Exit Function

ErrTrap:
  If Timer < lngTimeOut Then
    Resume 0
  End If
  
  If fInTransaction Then
    gADOCon.RollbackTrans
  End If

  AddHeader = False
  
End Function

Public Function ChangeHeaderStatus(udtStatus As EventLog_Status, _
                                   Optional lngSuccess As Variant, _
                                   Optional lngFailed As Variant) As Boolean

  'NOTE: lngSuccess and lngFailed need to be variants in order to
  'use the ISMISSING function ?
  
  Dim strSQL As String
  
  On Local Error GoTo ErrTrap
  
  If Not IsMissing(lngSuccess) And Not IsMissing(lngFailed) Then
    strSQL = " [AsrSysEventLog].[SuccessCount] = " & CStr(lngSuccess) & ", " & _
             " [AsrSysEventLog].[FailCount] = " & CStr(lngFailed) & ", "
  Else
    strSQL = vbNullString
  End If

  'add the EndTime of the event to the log.
  strSQL = strSQL & " [AsrSysEventLog].[EndTime] = GETDATE(), "

  strSQL = "UPDATE [AsrSysEventLog] SET " & _
           strSQL & _
           " [AsrSysEventLog].[Status] = " & udtStatus & " " & _
           " WHERE [AsrSysEventLog].[ID] = " & mlngEventLogID
  gADOCon.Execute strSQL

  
  'now that the EndTime  is populated calculate the duration.
  strSQL = "UPDATE [AsrSysEventLog] SET " & _
           " [AsrSysEventLog].[Duration] = DATEDIFF(second, [AsrSysEventLog].[DateTime], [AsrSysEventLog].[EndTime]) " & _
           " WHERE [AsrSysEventLog].[ID] = " & mlngEventLogID
  gADOCon.Execute strSQL
  
  ChangeHeaderStatus = True
  Exit Function
  
ErrTrap:
  ChangeHeaderStatus = False
  
End Function

Public Function AddDetailEntry(pstrNotes As String) As Boolean

  On Local Error GoTo ErrTrap

  Dim strNotes As String

  If mlngEventLogID > 0 Then
    
    If InStrB(1, pstrNotes, "'") Then
      strNotes = Replace(pstrNotes, "'", "''")
    Else
      strNotes = pstrNotes
    End If
    
    gADOCon.Execute "INSERT INTO AsrSysEventLogDetails (" & _
                    "EventLogID," & _
                    "Notes) " & _
                    "VALUES(" & _
                    mlngEventLogID & "," & _
                    "'" & strNotes & "')", , adCmdText + adExecuteNoRecords
    AddDetailEntry = True
    Exit Function
  End If
  
ErrTrap:
  AddDetailEntry = False
  
End Function

Public Function GetBatchRunID() As Long

  On Local Error GoTo ErrTrap
  
  Dim prstRowAdded As ADODB.Recordset

  If Not gblnBatchMode Then
    mlngBatchRunID = 0
    Exit Function
  End If


  ' Start a transaction
  gADOCon.BeginTrans

  ' Retrieve the previous max id
  Set prstRowAdded = gADOCon.Execute("SELECT MAX(BatchRunID) FROM AsrSysEventLog")

  ' Set function return to the id just added
  mlngBatchRunID = IIf(IsNull(prstRowAdded.Fields(0)), 0, prstRowAdded.Fields(0)) + 1

  ' Commit the transaction
  gADOCon.CommitTrans

  ' Tidy up
  Set prstRowAdded = Nothing

  GetBatchRunID = mlngBatchRunID

  Exit Function
  
ErrTrap:
  
  mlngBatchRunID = -1
  GetBatchRunID = mlngBatchRunID
  MsgBox "Warning : Error whilst retrieving the maximum BatchRunID." & vbCrLf & "(" & Err.Description & ")", vbCritical + vbOKOnly, "Event Log Error"
  
End Function

Public Property Get EventLogID() As Long
  EventLogID = mlngEventLogID
End Property
