VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents conEmail As ADODB.Connection
Attribute conEmail.VB_VarHelpID = -1

Public Sub SendImmediateEmails()

  Dim strConn As String
  Dim strSQL As String

  On Local Error GoTo LocalError

  Set conEmail = New ADODB.Connection
  With conEmail
    .ConnectionString = GetConnectionString("APP=", "Immediate Emails") & IIf(gbUseWindowsAuthentication, ";Integrated Security=SSPI;", "")
    .Provider = "SQLOLEDB"
    .CommandTimeout = 60
    .ConnectionTimeout = 60
    .CursorLocation = adUseServer
    .Mode = adModeReadWrite
    .Properties("Application Name") = "Immediate Emails"
    .Properties("Packet Size") = 32767
    .Open
  End With
  
  strSQL = "EXEC spASREmailImmediate '" & datGeneral.UserNameForSQL & "'"
  conEmail.Execute strSQL, , (adAsyncExecute Or adExecuteNoRecords)
  'conEmail.Execute strSQL

LocalError_Handler:
  
  Exit Sub

LocalError:
  If ASRDEVELOPMENT Then
    COAMsgBox Err.Description, vbCritical, "ASR DEVELOPMENT"
    Stop
  End If
  
  Resume LocalError_Handler

End Sub

Public Property Get StillExecuting() As Boolean
  StillExecuting = False
  If Not (conEmail Is Nothing) Then
    StillExecuting = (conEmail.State And adStateExecuting)
  End If
End Property

Private Sub conEmail_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
  conEmail.Close
  Set conEmail = Nothing
End Sub
