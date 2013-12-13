Option Strict Off
Option Explicit On

Imports ADODB
Imports HR.Intranet.Server.Enums

Public Class clsEventLog

	Private mlngEventLogID As Integer
	Private mstrBatchName As String = ""
	Private mlngBatchRunID As Integer
	Private mlngBatchJobID As Integer

	Public WriteOnly Property BatchMode() As Boolean
		Set(ByVal Value As Boolean)
			mlngBatchRunID = GetBatchRunID()
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

		Dim cmdAddHeader As New Command
		Dim prmNewID As Parameter
		Dim prmType As Parameter
		Dim prmName As Parameter
		Dim prmUserName As Parameter
		Dim prmBatchName As Parameter
		Dim prmBatchRunID As Parameter
		Dim prmBatchJobID As Parameter
		Dim sErrorMsg As String = ""
		Dim iLoop As Short

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

		Return True

ErrTrap:

		Return False

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
			strSQL = " [SuccessCount] = " & CStr(lngSuccess) & ", " & " [FailCount] = " & CStr(lngFailed) & ", "
		Else
			strSQL = vbNullString
		End If


		strSQL = String.Format("UPDATE [AsrSysEventLog] SET {0} [EndTime] = GETDATE(), [Duration] = DATEDIFF(second, [DateTime],  GETDATE()), Status = {1} WHERE [ID] = {2}", strSQL, CInt(udtStatus), mlngEventLogID)
		gADOCon.Execute(strSQL)

		Return True

ErrTrap:
		Return False

	End Function

	Public Function AddDetailEntry(ByRef pstrNotes As String) As Boolean

		On Error GoTo ErrTrap

		If mlngEventLogID > 0 Then
			gADOCon.Execute("INSERT INTO AsrSysEventLogDetails (" & "EventLogID," & "Notes) " & "VALUES(" & mlngEventLogID & "," & "'" & Replace(pstrNotes, "'", "''") & "')")
			Return True
		End If

ErrTrap:
		Return False

	End Function

	Private Function GetBatchRunID() As Integer

		On Error GoTo ErrTrap

		Dim prstRowAdded As Recordset

		' Start a transaction
		gADOCon.BeginTrans()

		' Retrieve the previous max id
		prstRowAdded = gADOCon.Execute("SELECT MAX(BatchRunID) FROM AsrSysEventLog")

		' Set function return to the id just added
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		GetBatchRunID = IIf(IsDBNull(prstRowAdded.Fields(0).Value), 0, prstRowAdded.Fields(0).Value) + 1

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