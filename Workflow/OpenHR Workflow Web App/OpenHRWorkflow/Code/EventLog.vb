Namespace Code

	Public Class EventLog

		Private Const MAXLOGENTRYLENGTH As Integer = 32766
		Private mobjEventLog As System.Diagnostics.EventLog

		Private mstrLog As String
		Private mstrSource As String

		Public Sub WriteEntry(ByVal psEventLogEntry As String, _
		ByVal pEntryType As System.Diagnostics.EventLogEntryType)
			Dim sEventLogEntry As String

			Try
				sEventLogEntry = psEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
				mobjEventLog = New System.Diagnostics.EventLog()
				mobjEventLog.Log = mstrLog
				mobjEventLog.Source = mstrSource
				mobjEventLog.WriteEntry(sEventLogEntry, pEntryType)
				mobjEventLog.Close()
				mobjEventLog = Nothing
			Catch ex As Exception

			End Try

		End Sub

		Public Function WriteEntry(ByVal psEventLogEntry As String, _
		ByVal pEntryType As System.Diagnostics.EventLogEntryType, _
		ByVal psIgnoreIfMatchesEntry As String) As Boolean
			Dim fEntryMatches As Boolean

			fEntryMatches = True

			Try
				If psEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim <> psIgnoreIfMatchesEntry.Trim Then
					WriteEntry(psEventLogEntry, pEntryType)
					fEntryMatches = False
				End If
			Catch ex As Exception

			End Try

			WriteEntry = fEntryMatches
		End Function

		Public Sub New()

		End Sub

		Public Sub New(ByVal log As String, ByVal source As String)
			'mobjEventLog = New system.diagnostics.eventlog()
			'mobjEventLog.Log = log
			'mobjEventLog.Source = source
			mstrLog = log
			mstrSource = source
		End Sub


	End Class
End Namespace