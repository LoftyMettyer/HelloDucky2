Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsEventLog_NET.clsEventLog")> Public Class clsEventLog
	
	Public Enum EventLog_Type
		eltCrossTab = 1
		eltCustomReport = 2
		eltDataTransfer = 3
		eltExport = 4
		eltGlobalAdd = 5
		eltGlobalDelete = 6
		eltGlobalUpdate = 7
		eltImport = 8
		eltMailMerge = 9
		eltDiaryDelete = 10
		eltDiaryRebuild = 11
		eltEmailRebuild = 12
		eltStandardReport = 13 'MH20010305
		eltRecordEditing = 14
		eltSystemError = 15
		eltMatchReport = 16
		eltCalandarReport = 17
		eltLabel = 18
		eltLabelType = 19
		eltRecordProfile = 20
		eltSuccessionPlanning = 21
		eltCareerProgression = 22
	End Enum
	
	Public Enum EventLog_Status
		elsPending = 0
		elsCancelled = 1
		elsFailed = 2
		elsSuccessful = 3
		elsSkipped = 4
		elsError = 5
	End Enum
	
	Private mclsData As New clsDataAccess
	Private mlngEventLogID As Integer
	Private mstrBatchName As String
	Private mlngBatchRunID As Integer
	Private mlngBatchJobID As Integer
	Private mblnBatchMode As Boolean
	
	Public WriteOnly Property BatchMode() As Boolean
		Set(ByVal Value As Boolean)
			mblnBatchMode = Value
			mlngBatchRunID = GetBatchRunID
		End Set
	End Property
	
	Public WriteOnly Property BatchJobName() As String
		Set(ByVal Value As String)
			' When the batch job name is passed into the class module, the BatchRunID is
			' allocated automatically. Process is complete hidden from the utilities.
			mstrBatchName = Value
		End Set
	End Property
	
	Public WriteOnly Property BatchJobID() As Integer
		Set(ByVal Value As Integer)
			mlngBatchJobID = Value
		End Set
	End Property
	
	
	
	Public Property EventLogID() As Integer
		Get
			EventLogID = mlngEventLogID
		End Get
		Set(ByVal Value As Integer)
			mlngEventLogID = Value
		End Set
	End Property
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mclsData = New clsDataAccess
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsData = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Sub TidyUp()
		
		' Must be called from batchjobs when the batch has finished, otherwise all
		' subsequent functions in the users current session will be logged as being
		' in the batch job
		mstrBatchName = vbNullString
		mlngBatchRunID = 0
		mlngBatchJobID = 0
		
	End Sub
	
	Public Function AddHeader(ByRef udtType As EventLog_Type, ByRef strName As String) As Boolean
		
		On Error GoTo ErrTrap
		
		Dim cmdAddHeader As New ADODB.Command
		Dim prmNewID As New ADODB.Parameter
		Dim prmType As New ADODB.Parameter
		Dim prmName As New ADODB.Parameter
		Dim prmUserName As New ADODB.Parameter
		Dim prmBatchName As New ADODB.Parameter
		Dim prmBatchRunID As New ADODB.Parameter
		Dim prmBatchJobID As New ADODB.Parameter
		Dim prmInsertSQL As New ADODB.Parameter
		Dim strInsertSQL As String
		Dim sErrorMsg As String
		Dim iLoop As Short
		
		'JPD 20041104 Replaced the 'InsertSQL' parameter with individual parameters
		' for security tightening.
		'strInsertSQL = "INSERT INTO AsrSysEventLog (" & _
		'"DateTime," & _
		'"Type," & _
		'"Name," & _
		'"Status," & _
		'"Username," & _
		'"Mode," & _
		'"BatchName," & _
		'"SuccessCount," & _
		'"FailCount," & _
		'"BatchRunID," & _
		'"BatchJobID) " & _
		'"VALUES(" & _
		'"GETDATE()," & _
		'udtType & "," & _
		'"'" & Replace(strName, "'", "''") & "'," & _
		'elsPending & "," & _
		'"'" & StrConv(gsUsername, vbProperCase) & "'," & _
		'IIf(mstrBatchName = vbNullString, 0, 1) & "," & _
		'"'" & Replace(mstrBatchName, "'", "''") & "'," & _
		'"NULL," & _
		'"NULL," & _
		'IIf(mlngBatchRunID > 0, mlngBatchRunID, "NULL") & "," & _
		'IIf(mlngBatchJobID > 0, mlngBatchJobID, "NULL") & ")"
		
		cmdAddHeader.CommandText = "sp_ASRIntAddEventLogHeader"
		cmdAddHeader.CommandType = 4
		cmdAddHeader.let_ActiveConnection(gADOCon)
		
		prmNewID = cmdAddHeader.CreateParameter("newID", 3, 2)
		cmdAddHeader.Parameters.Append(prmNewID)
		
		prmType = cmdAddHeader.CreateParameter("type", 3, 1)
		cmdAddHeader.Parameters.Append(prmType)
		prmType.Value = udtType
		
		prmName = cmdAddHeader.CreateParameter("name", 200, 1, 150)
		cmdAddHeader.Parameters.Append(prmName)
		prmName.Value = strName
		
		prmUserName = cmdAddHeader.CreateParameter("userName", 200, 1, 50)
		cmdAddHeader.Parameters.Append(prmUserName)
		prmUserName.Value = StrConv(gsUsername, VbStrConv.ProperCase)
		
		prmBatchName = cmdAddHeader.CreateParameter("batchName", 200, 1, 50)
		cmdAddHeader.Parameters.Append(prmBatchName)
		prmBatchName.Value = mstrBatchName
		
		prmBatchRunID = cmdAddHeader.CreateParameter("batchRunID", 3, 1)
		cmdAddHeader.Parameters.Append(prmBatchRunID)
		prmBatchRunID.Value = mlngBatchRunID
		
		prmBatchJobID = cmdAddHeader.CreateParameter("batchJobID", 3, 1)
		cmdAddHeader.Parameters.Append(prmBatchJobID)
		prmBatchJobID.Value = mlngBatchJobID
		
		cmdAddHeader.ActiveConnection.Errors.Clear()
		
		cmdAddHeader.Execute()
		
		If cmdAddHeader.ActiveConnection.Errors.Count > 0 Then
			For iLoop = 1 To cmdAddHeader.ActiveConnection.Errors.Count
				sErrorMsg = sErrorMsg & vbNewLine & (cmdAddHeader.ActiveConnection.Errors.Item(iLoop - 1).Description)
			Next 
			cmdAddHeader.ActiveConnection.Errors.Clear()
			mlngEventLogID = 0
			AddHeader = False
			Exit Function
		Else
			mlngEventLogID = cmdAddHeader.Parameters("newID").Value
		End If
		
		'UPGRADE_NOTE: Object cmdAddHeader may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdAddHeader = Nothing
		
		'  Dim fInTransaction As Boolean
		'  Dim lngTimeOut As Long
		'
		'  lngTimeOut = Timer + 5
		'
		'  mlngEventLogID = 0
		'  fInTransaction = False
		'
		'  ' If we are in batchmode, is everything ok ?
		'  If (mlngBatchRunID = -1 And mstrBatchName <> vbNullString) Then
		'    AddHeader = False
		'    Exit Function
		'  End If
		'
		'  ' Temp recordset to hold the returned id of the record just added
		'  Dim prstRowAdded As ADODB.Recordset
		'
		'  ' Start a transaction
		'  gADOCon.BeginTrans
		'  fInTransaction = True
		'
		'  ' Do the insert
		'  gADOCon.Execute "INSERT INTO AsrSysEventLog (" & _
		''                  "DateTime," & _
		''                  "Type," & _
		''                  "Name," & _
		''                  "Status," & _
		''                  "Username," & _
		''                  "Mode," & _
		''                  "BatchName," & _
		''                  "SuccessCount," & _
		''                  "FailCount," & _
		''                  "BatchRunID) " & _
		''                  "VALUES(" & _
		''                  "GETDATE()," & _
		''                   udtType & "," & _
		''                  "'" & Replace(strName, "'", "''") & "'," & _
		''                  elsPending & "," & _
		''                  "'" & gsUsername & "'," & _
		''                  IIf(mstrBatchName = vbNullString, 0, 1) & "," & _
		''                  "'" & Replace(mstrBatchName, "'", "''") & "'," & _
		''                  "NULL," & _
		''                  "NULL," & _
		''                  IIf(mlngBatchRunID > 0, mlngBatchRunID, "NULL") & ")"
		'
		'  ' Retrieve the id just added
		'  Set prstRowAdded = gADOCon.Execute("SELECT MAX(ID) FROM AsrSysEventLog")
		'
		'  ' Set function return to the id just added
		'  mlngEventLogID = prstRowAdded.Fields(0)
		'
		'  ' Commit the transaction
		'  gADOCon.CommitTrans
		'  fInTransaction = False
		'
		'  ' Tidy up
		'  Set prstRowAdded = Nothing
		'
		AddHeader = True
		
		Exit Function
		
ErrTrap: 
		'  If Timer < lngTimeOut Then
		'    Resume 0
		'  End If
		'
		'  If fInTransaction Then
		'    gADOCon.RollbackTrans
		'  End If
		'
		AddHeader = False
		
	End Function
	
	Public Function ChangeHeaderStatus(ByRef udtStatus As EventLog_Status, Optional ByRef lngSuccess As Object = Nothing, Optional ByRef lngFailed As Object = Nothing) As Boolean
		
		'NOTE: lngSuccess and lngFailed need to be variants in order to
		'use the ISMISSING function ?
		
		Dim strSQL As String
		
		On Error GoTo ErrTrap
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If Not IsNothing(lngSuccess) And Not IsNothing(lngFailed) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object lngFailed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object lngSuccess. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strSQL = " [AsrSysEventLog].[SuccessCount] = " & CStr(lngSuccess) & ", " & " [AsrSysEventLog].[FailCount] = " & CStr(lngFailed) & ", "
		Else
			strSQL = vbNullString
		End If
		
		'add the EndTime of the event to the log.
		strSQL = strSQL & " [AsrSysEventLog].[EndTime] = GETDATE(), "
		
		strSQL = "UPDATE [AsrSysEventLog] SET " & strSQL & " [AsrSysEventLog].[Status] = " & udtStatus & " " & " WHERE [AsrSysEventLog].[ID] = " & mlngEventLogID
		gADOCon.Execute(strSQL)
		
		
		'now that the EndTime  is populated calculate the duration.
		strSQL = "UPDATE [AsrSysEventLog] SET " & " [AsrSysEventLog].[Duration] = DATEDIFF(second, [AsrSysEventLog].[DateTime], [AsrSysEventLog].[EndTime]) " & " WHERE [AsrSysEventLog].[ID] = " & mlngEventLogID
		gADOCon.Execute(strSQL)
		
		ChangeHeaderStatus = True
		Exit Function
		
ErrTrap: 
		ChangeHeaderStatus = False
		
	End Function
	
	Public Function AddDetailEntry(ByRef pstrNotes As String) As Boolean
		
		On Error GoTo ErrTrap
		
		If mlngEventLogID > 0 Then
			gADOCon.Execute("INSERT INTO AsrSysEventLogDetails (" & "EventLogID," & "Notes) " & "VALUES(" & mlngEventLogID & "," & "'" & Replace(pstrNotes, "'", "''") & "')")
			AddDetailEntry = True
			Exit Function
		End If
		
ErrTrap: 
		AddDetailEntry = False
		
	End Function
	
	Private Function GetBatchRunID() As Integer
		
		On Error GoTo ErrTrap
		
		Dim prstRowAdded As ADODB.Recordset
		
		' Start a transaction
		gADOCon.BeginTrans()
		
		' Retrieve the previous max id
		prstRowAdded = gADOCon.Execute("SELECT MAX(BatchRunID) FROM AsrSysEventLog")
		
		' Set function return to the id just added
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		GetBatchRunID = IIf(IsDbNull(prstRowAdded.Fields(0).Value), 0, prstRowAdded.Fields(0).Value) + 1
		
		' Commit the transaction
		gADOCon.CommitTrans()
		
		' Tidy up
		'UPGRADE_NOTE: Object prstRowAdded may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prstRowAdded = Nothing
		
		Exit Function
		
ErrTrap: 
		
		GetBatchRunID = -1
		'NO MSGBOX ON THE SERVER ! - MsgBox "Warning : Error whilst retrieving the maximum BatchRunID." & vbNewLine & "(" & Err.Description & ")", vbCritical + vbOKOnly, "Event Log Error"
		
	End Function
End Class