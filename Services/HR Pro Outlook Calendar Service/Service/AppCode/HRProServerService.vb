Imports System.IO
Imports System.Configuration
Imports System.Collections
Imports System.Data.SqlClient

Public Class OpenHROutlookCalendarService
    Private _svcVersion As VersionNumber
    Private _openHRSystems As List(Of OpenHRSystem) = New List(Of OpenHRSystem)
    Private _exchangeServer As String = String.Empty
    Private _serviceAccountPassword As String = String.Empty
    Private _outlookApp As OutlookCalendar
    Private _outlookWait As Int32 = 0
    Private _enableTrace As Boolean = False
    Private _CommandTimeout As Integer

    Private bIsProcessing As Boolean

    ''' <remarks>
    ''' The minimum OpenHR database version required.
    ''' </remarks>
  Private Const MINIMUMDBVERSION As Single = 8.0
    Private Const POLLINGINTERVAL As Int32 = 60000
    Private Const SERVERTAG As String = "server"
    Private Const DATABASETAG As String = "database"
    Private Const EXCHANGETAG As String = "exchange"
    Private Const DEBUGTAG As String = "debug"
    Private Const COMMANDTIMEOUTTAG As String = "commandtimeout"

    Protected Overrides Sub OnStart(ByVal args() As String)
        Dim appSettings As Specialized.NameValueCollection = ConfigurationManager.AppSettings
        Dim _tempOpenHRSystems As List(Of OpenHRSystem) = New List(Of OpenHRSystem)
        Dim defaultServer As String = "."
        Dim Tag As String = String.Empty
        Dim Value As String = String.Empty
        Dim sEventLogEntry As String = String.Empty

        Dim sUserName As String = ""
        Dim sPassword As String = ""

        Try

            bIsProcessing = False

            Dim fileMap As ExeConfigurationFileMap = New ExeConfigurationFileMap()
            fileMap.ExeConfigFilename = System.Reflection.Assembly.GetExecutingAssembly.Location.Replace(".exe", ".custom.config")

            Dim config As Configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None)

            If Not config.HasFile Then
                config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
                If Not config.HasFile Then
                    Throw New ArgumentNullException("No config file specified")
                End If
            End If

            Dim section As AppSettingsSection = config.AppSettings

            Dim currentOpenHR As OpenHRSystem

            _CommandTimeout = -1

            For Each configValue As KeyValueConfigurationElement In section.Settings
                If configValue.Key.ToLower().Trim() = "openhruser" Then sUserName = configValue.Value
                If configValue.Key.ToLower().Trim() = "openhrpassword" Then sPassword = configValue.Value
            Next

            For Each configValue As KeyValueConfigurationElement In section.Settings
                If configValue.Key = "serviceAccountPassword" Then
                    _serviceAccountPassword = configValue.Value
                    Continue For
                End If

                Tag = configValue.Key.ToLower().Trim()
                Value = configValue.Value.ToLower().Trim()

                currentOpenHR = New OpenHRSystem()

                ' User credentials - same user for all databases
                currentOpenHR.UserName = sUserName
                currentOpenHR.Password = sPassword

                ' Debug key
                If Tag.Length = DEBUGTAG.Length AndAlso Tag.Substring(0, DEBUGTAG.Length) = DEBUGTAG.ToLower() Then
                    If Value.Length > 0 AndAlso Value = "1" Then
                        _enableTrace = True
                    End If
                End If

                ' Exchange key
                If Tag.Length = EXCHANGETAG.Length AndAlso Tag.Substring(0, EXCHANGETAG.Length) = EXCHANGETAG.ToLower() Then
                    If Value.Length > 0 Then
                        _exchangeServer = Value
                    End If
                End If

                ' Command timeout key
                If Tag.Length = COMMANDTIMEOUTTAG.Length AndAlso Tag.Substring(0, COMMANDTIMEOUTTAG.Length) = COMMANDTIMEOUTTAG.ToLower() Then
                    If Value.Length > 0 Then
                        Try
                            _CommandTimeout = CInt(Value)
                        Catch ex As Exception
                        End Try
                    End If
                End If

                ' Server key
                If Tag.Length >= SERVERTAG.Length AndAlso Tag.Substring(0, SERVERTAG.Length) = SERVERTAG.ToLower() Then
                    If Value.Length > 0 Then
                        currentOpenHR.ServerName = Value
                        defaultServer = currentOpenHR.ServerName
                    End If
                ElseIf Tag.Length >= DATABASETAG.Length AndAlso Tag.Substring(0, DATABASETAG.Length) = DATABASETAG.ToLower() Then
                    ' Database key
                    currentOpenHR.DatabaseName = Value

                    If currentOpenHR.ServerName = String.Empty Then
                        currentOpenHR.ServerName = defaultServer
                        currentOpenHR.DefaultServer = True
                    End If

                    currentOpenHR.Serviced = True
                    currentOpenHR.ServiceServer = Environment.MachineName
                    currentOpenHR.VersionOK = True

                    ' We can only add a system to the list if we have a database to service .... 
                    ' so lets do it in here
                    _tempOpenHRSystems.Add(currentOpenHR)

                    ' Free the list from memory
                    currentOpenHR = Nothing
                End If
            Next
        Catch ex As ArgumentNullException
            LogEvent(ex.Message, EventLogEntryType.Warning)
        Catch ex As Exception
            LogEvent(ex.Message, ex.StackTrace, EventLogEntryType.Error)
        End Try

        LogEvent(String.Format("OpenHR Outlook Calendar Service ({0}) started successfully.", _svcVersion.ToString()))

        Try
            Dim blankSystems As Int16 = 0

            ' describe in the eventlog which systems we're servicing
            For Each currentOpenHR As OpenHRSystem In _tempOpenHRSystems
                If Not currentOpenHR.DefaultServer AndAlso currentOpenHR.DatabaseName <> String.Empty Then
                    ' User specified both server and database parameters
                    LogEvent(String.Format("OpenHR Outlook Calendar service configured for ({0}.{1}). Exchange Server = {2}", _
                                           currentOpenHR.ServerName, _
                                           currentOpenHR.DatabaseName, _
                                           _exchangeServer))
                    _openHRSystems.Add(currentOpenHR)
                ElseIf currentOpenHR.DefaultServer AndAlso currentOpenHR.DatabaseName = String.Empty Then
                    ' No server or db specified so ignore (in theory should never happen
                    blankSystems = CShort(blankSystems + 1)
                ElseIf currentOpenHR.DefaultServer Then
                    ' No server parameter was defined so default is assumed
                    LogEvent(String.Format("OpenHR Outlook Calendar service configured for ({0}.{1}). Exchange Server = {2}", _
                                           currentOpenHR.ServerName, _
                                           currentOpenHR.DatabaseName, _
                                           _exchangeServer))
                    _openHRSystems.Add(currentOpenHR)
                ElseIf currentOpenHR.DatabaseName = String.Empty Then
                    ' User specified a server but no database..... ignore
                    LogEvent(String.Format("No database parameter defined for the server ({0}). System will be ignored.", _
                        currentOpenHR.ServerName), EventLogEntryType.Warning)
                End If
            Next

            If _exchangeServer.Equals(String.Empty) Then
                LogEvent("No exchange server parameter defined.  Service stopping.", EventLogEntryType.Warning)
                Me.Stop()
            ElseIf _openHRSystems.Count <= 0 Then
                LogEvent("No server and database parameters defined.  Service stopping.", EventLogEntryType.Warning)
                Me.Stop()
            ElseIf blankSystems > 0 Then
                LogEvent(String.Format("OpenHR Outlook Calendar service configured for {0} system(s). {1} blank systems were ignored.", _
                    _openHRSystems.Count, blankSystems), EventLogEntryType.Warning)
            Else
                LogEvent(String.Format("OpenHR Outlook Calendar service configured for {0} OpenHR system(s).", _
                                       _openHRSystems.Count), EventLogEntryType.Information)
            End If

            If _CommandTimeout < 0 Then
                _CommandTimeout = 200
                sEventLogEntry = "The default command timeout of " & CStr(_CommandTimeout) & " seconds has been applied."
            ElseIf _CommandTimeout = 0 Then
                sEventLogEntry = "Indefinite command timeout configured."
            Else
                sEventLogEntry = "Command timeout configured to be " & CStr(_CommandTimeout) & " seconds."
            End If

            LogEvent(sEventLogEntry, EventLogEntryType.Information)

            If _openHRSystems.Count > 0 Then
                SrvTmr.Interval = POLLINGINTERVAL
                SrvTmr.Enabled = True
            End If
        Catch ex As Exception
            LogEvent(ex.Message, ex.StackTrace, EventLogEntryType.Error)
        End Try

        appSettings = Nothing
        _tempOpenHRSystems = Nothing

    End Sub

    Protected Overrides Sub OnStop()
        If Not _outlookApp Is Nothing Then
            _outlookApp.Quit()
            _outlookApp = Nothing
        End If

        For iOpenHRSystem As Int32 = 0 To _openHRSystems.Count - 1
            If _openHRSystems(iOpenHRSystem).Serviced AndAlso _
              _openHRSystems(iOpenHRSystem).ServiceServer = Environment.MachineName Then

                Try
                    Using conn As New SqlConnection(_openHRSystems(iOpenHRSystem).ConnectionString)
                        conn.Open()
                        SaveSystemSetting("outlook service", "running", "0", conn)
                        conn.Close()
                    End Using

                Catch ex As Exception
                    LogEvent(String.Format("Error clearing system parameters for database ({0}). {1}{2}", _
                      _openHRSystems(iOpenHRSystem).ToString(), ControlChars.NewLine, ex.Message), EventLogEntryType.Error)
                End Try
            End If
        Next

        _openHRSystems = Nothing

        LogEvent(String.Format("OpenHR Outlook Calendar Service ({0}) stopped successfully.", _svcVersion.ToString()))
    End Sub

    Private Sub SrvTmr_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles SrvTmr.Elapsed
        'SrvTmr.Enabled = False

        'If IsProcessRunningInContext("outlook") Then
        '  If _outlookWait < 5 Then
        '    ' Wait for 5 minutes and then kill it
        '    If _outlookWait <= 0 Then
        '      LogEvent("An Outlook process is already running - killing process in 5 minutes.", EventLogEntryType.Warning)
        '    End If
        '    _outlookWait += 1
        '  Else
        '    LogEvent("Killing Outlook process.", EventLogEntryType.Warning)
        '    If IsProcessRunningInContext("outlook", True) Then
        '      _outlookWait = 0
        '    End If
        '  End If
        'Else
        'OutlookBatch()
        'End If

        'SrvTmr.Enabled = True

        Dim sw As StreamWriter = Nothing

        'If _enableTrace Then
        '  sw = File.AppendText( _
        '    FolderFromFileName(Reflection.Assembly.GetExecutingAssembly.Location) & "OutlookSvc_Debug.txt")
        '  TraceLog("---------------", sw, _enableTrace)
        '  TraceLog("Timer Activated", sw, _enableTrace)
        '  TraceLog("---------------", sw, _enableTrace)
        '  sw.Close()
        '  sw.Dispose()
        '  sw = Nothing
        'End If

        If Not bIsProcessing Then
            bIsProcessing = True
            OutlookBatch()
            bIsProcessing = False
        End If

    End Sub

    ''' <summary>
    ''' Batch to add events to Outlook Calendars
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks>Replaces the stored procedure spASROutlookBatch</remarks>
    Private Function OutlookBatch() As Boolean

        Dim dateFormat As String = String.Empty
        Dim sql As String = String.Empty

        Dim linkID As Int32 = 0, folderID As Int32 = 0
        Dim startDateColumnID As Int32 = 0, endDateColumnID As Int32 = 0
        Dim startTimeColumnID As Int32 = 0, endTimeColumnID As Int32 = 0
        Dim recordID As Int32 = 0, recordDescExprID As Int32 = 0, subjectExprID As Int32 = 0
        Dim storeID As String = String.Empty, entryID As String = String.Empty
        Dim title As String = String.Empty, content As String = String.Empty
        Dim fixedStartTime As String = String.Empty, fixedEndTime As String = String.Empty
        Dim deleted As Boolean = False, reminder As Boolean = False
        Dim reminderOffset As Int32 = 0, reminderPeriod As Int32 = 0
        Dim busyStatus As Int32 = 0, timeRange As Int32 = 0
        Dim folderType As Int32 = 0, folderExprID As Int32 = 0
        Dim folderPath As String = String.Empty

        Dim createdEntry As Boolean = False
        Dim doOutlookOK As Boolean = False

        Dim allDayEvent As Boolean = False
        Dim startDate As DateTime = CDate(SqlTypes.SqlDateTime.MinValue)
        Dim endDate As DateTime = CDate(SqlTypes.SqlDateTime.MinValue)
        Dim startTime As String = String.Empty
        Dim endTime As String = String.Empty
        Dim subject As String = String.Empty
        Dim folder As String = String.Empty

        Dim sw As StreamWriter = Nothing
        Try
            If _enableTrace Then
                sw = File.AppendText( _
                  FolderFromFileName(Reflection.Assembly.GetExecutingAssembly.Location) & "OutlookSvc_Debug.txt")
                TraceLog("-------------", sw, _enableTrace)
                TraceLog("Batch Started", sw, _enableTrace)
                TraceLog("-------------", sw, _enableTrace)
            End If

            Using _outlookApp As New OutlookCalendar(sw, _enableTrace, _serviceAccountPassword)

                For iOpenHRSystem As Int32 = 0 To _openHRSystems.Count - 1
                    Dim serverDBName As String = String.Format("{0}.{1}", _
                      _openHRSystems(iOpenHRSystem).ServerName.ToUpper, _openHRSystems(iOpenHRSystem).DatabaseName.ToUpper)

                    Try
                        TraceLog(String.Format("Opening Connection {0}", _openHRSystems(iOpenHRSystem).ConnectionString), sw, _enableTrace)
                        Using conn As New SqlConnection(_openHRSystems(iOpenHRSystem).ConnectionString)
                            conn.Open()
                            TraceLog("Connection Open", sw, _enableTrace)

                            If DatabaseIsOK(_openHRSystems(iOpenHRSystem), conn, sw) Then
                                TraceLog("Database is okay", sw, _enableTrace)
                                dateFormat = GetSystemSetting("email", "date format", conn)

                                sql = GetOutlookEventsSQL()
                                Dim cmdEvents As SqlCommand = New SqlCommand(sql, conn)
                                cmdEvents.CommandType = CommandType.Text

                                Dim reader As SqlDataReader = cmdEvents.ExecuteReader()

                                If reader.HasRows Then
                                    TraceLog("Checking Exchange", sw, _enableTrace)
                                    If (Not _outlookApp.LoggedOn) AndAlso _outlookApp.Logon(_exchangeServer) Then
                                        ' Logged onto Outlook
                                    ElseIf Not _outlookApp.LoggedOn Then
                                        LogEvent("Could not logon to Exchange :- " & _outlookApp.ErrorMessage, EventLogEntryType.Warning)
                                        Exit For
                                    Else
                                        ' Already logged onto Outlook
                                    End If
                                    TraceLog("Connected to Exchange", sw, _enableTrace)

                                    While reader.Read()
                                        storeID = Utilities.NullSafeString(reader("StoreID"))
                                        entryID = Utilities.NullSafeString(reader("EntryID"))
                                        deleted = Utilities.NullSafeBoolean(reader("Deleted"))
                                        linkID = Utilities.NullSafeInteger(reader("LinkID"))
                                        folderID = Utilities.NullSafeInteger(reader("FolderID"))
                                        recordID = Utilities.NullSafeInteger(reader("RecordID"))
                                        startDateColumnID = Utilities.NullSafeInteger(reader("StartDate"))
                                        endDateColumnID = Utilities.NullSafeInteger(reader("EndDate"))
                                        fixedStartTime = Utilities.NullSafeString(reader("FixedStartTime"))
                                        fixedEndTime = Utilities.NullSafeString(reader("FixedEndTime"))
                                        startTimeColumnID = Utilities.NullSafeInteger(reader("ColumnStartTime"))
                                        endTimeColumnID = Utilities.NullSafeInteger(reader("ColumnEndTime"))
                                        timeRange = Utilities.NullSafeInteger(reader("TimeRange"))
                                        title = Utilities.NullSafeString(reader("Title"))
                                        subjectExprID = Utilities.NullSafeInteger(reader("Subject"))
                                        recordDescExprID = Utilities.NullSafeInteger(reader("RecordDescExprID"))
                                        folderType = Utilities.NullSafeInteger(reader("FolderType"))
                                        folderPath = Utilities.NullSafeString(reader("FixedPath"))
                                        folderExprID = Utilities.NullSafeInteger(reader("ExprID"))
                                        content = Utilities.NullSafeString(reader("Content"))
                                        reminder = Utilities.NullSafeBoolean(reader("Reminder"))
                                        reminderOffset = Utilities.NullSafeInteger(reader("ReminderOffset"))
                                        reminderPeriod = Utilities.NullSafeInteger(reader("ReminderPeriod"))
                                        busyStatus = Utilities.NullSafeInteger(reader("BusyStatus"))

                                        _outlookApp.ResetStoreAndEntry()
                                        doOutlookOK = True

                                        Dim emptyIDs As Boolean = (storeID = String.Empty AndAlso entryID = String.Empty)

                                        TraceLog(String.Format("Servicing {0}", serverDBName), sw, _enableTrace)
                                        TraceLog(title, sw, _enableTrace)
                                        TraceLog("Deleted : " & deleted.ToString(), sw, _enableTrace)
                                        TraceLog("StoreID : " & storeID, sw, _enableTrace)
                                        TraceLog("EntryID : " & entryID, sw, _enableTrace)

                                        If deleted OrElse (Not emptyIDs) Then
                                            _outlookApp.StoreID = storeID
                                            _outlookApp.EntryID = entryID

                                            Dim transaction As SqlTransaction

                                            Using connEvents As New SqlConnection(_openHRSystems(iOpenHRSystem).ConnectionString)
                                                connEvents.Open()

                                                Dim cmd As SqlCommand = connEvents.CreateCommand()

                                                ' Start a local transaction
                                                transaction = connEvents.BeginTransaction()

                                                TraceLog("BEGIN TRANS", sw, _enableTrace)

                                                ' Must assign both transaction object and connection
                                                ' to Command object for a pending local transaction.
                                                cmd.Connection = connEvents
                                                cmd.Transaction = transaction
                                                cmd.CommandTimeout = CInt(IIf((_CommandTimeout < 200) And (_CommandTimeout <> 0), 200, _CommandTimeout))

                                                Try
                                                    If Not deleted Then
                                                        TraceLog("Update to NULLs", sw, _enableTrace)

                                                        ' Update to NULLs
                                                        sql = "UPDATE ASRSysOutlookEvents WITH(ROWLOCK) "
                                                        sql &= "SET StoreID = '' "
                                                        sql &= ", EntryID = '' "
                                                        sql &= ", RefreshDate = GETDATE() "
                                                        sql &= "WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID"
                                                        cmd.Parameters.AddWithValue("@LinkID", linkID)
                                                        cmd.Parameters.AddWithValue("@FolderID", folderID)
                                                        cmd.Parameters.AddWithValue("@RecordID", recordID)
                                                        cmd.CommandText = sql
                                                        cmd.ExecuteNonQuery()
                                                        cmd.Parameters.Clear()
                                                    Else
                                                        TraceLog("Delete from ASRSysOutlookEvents", sw, _enableTrace)

                                                        ' Delete entry
                                                        sql = "DELETE FROM ASRSysOutlookEvents WITH(ROWLOCK) "
                                                        sql &= "WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID"
                                                        cmd.Parameters.AddWithValue("@LinkID", linkID)
                                                        cmd.Parameters.AddWithValue("@FolderID", folderID)
                                                        cmd.Parameters.AddWithValue("@RecordID", recordID)
                                                        cmd.CommandText = sql
                                                        cmd.ExecuteNonQuery()
                                                        cmd.Parameters.Clear()
                                                    End If

                                                    If _outlookApp.DeleteEntry() Then
                                                        TraceLog("Delete from Outlook", sw, _enableTrace)
                                                        TraceLog("COMMIT TRANS", sw, _enableTrace)
                                                        ' Attempt to commit the transaction.
                                                        transaction.Commit()
                                                        doOutlookOK = True
                                                    Else
                                                        Try
                                                            TraceLog("ROLLBACK TRANS : Delete from Outlook failed with '" & _
                                                              _outlookApp.ErrorMessage & "'", sw, _enableTrace)

                                                            ' Rollback
                                                            transaction.Rollback()
                                                        Catch ex2 As Exception
                                                            ' This catch block will handle any errors that may have occurred
                                                            ' on the server that would cause the rollback to fail, such as
                                                            ' a closed connection.
                                                            LogEvent("Error rolling back transaction duplicates may occur.", EventLogEntryType.Error)
                                                        End Try

                                                        doOutlookOK = False
                                                        LogEvent("Delete Error :- " & _outlookApp.ErrorMessage, EventLogEntryType.Warning)

                                                        ' Update so we don't try again
                                                        Using connUpd As New SqlConnection(_openHRSystems(iOpenHRSystem).ConnectionString)
                                                            connUpd.Open()

                                                            cmd = connEvents.CreateCommand()

                                                            TraceLog("BEGIN TRANS : Update to not retry", sw, _enableTrace)
                                                            ' Begin Trans
                                                            transaction = connUpd.BeginTransaction()
                                                            cmd.Connection = connUpd
                                                            cmd.Transaction = transaction
                                                            cmd.CommandTimeout = CInt(IIf((_CommandTimeout < 200) And (_CommandTimeout <> 0), 200, _CommandTimeout))

                                                            sql = "UPDATE ASRSysOutlookEvents WITH(ROWLOCK) "
                                                            sql &= "SET Refresh = 0 "
                                                            sql &= ", Deleted = 0 "
                                                            sql &= ", RefreshDate = GETDATE() "
                                                            sql &= "WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID"
                                                            cmd.Parameters.AddWithValue("@LinkID", linkID)
                                                            cmd.Parameters.AddWithValue("@FolderID", folderID)
                                                            cmd.Parameters.AddWithValue("@RecordID", recordID)
                                                            cmd.CommandText = sql
                                                            cmd.ExecuteNonQuery()
                                                            cmd.Parameters.Clear()

                                                            ' Attempt to commit the transaction.
                                                            transaction.Commit()
                                                            TraceLog("COMMIT TRANS : Update to not retry", sw, _enableTrace)
                                                        End Using
                                                    End If
                                                Catch sqlEx As SqlException
                                                    doOutlookOK = False
                                                    ' Rollback
                                                    Try
                                                        TraceLog("ROLLBACK TRANS : Update or Delete failed", sw, _enableTrace)
                                                        transaction.Rollback()
                                                    Catch ex2 As Exception
                                                        ' This catch block will handle any errors that may have occurred
                                                        ' on the server that would cause the rollback to fail, such as
                                                        ' a closed connection.
                                                        LogEvent("Error rolling back transaction duplicates may occur.", EventLogEntryType.Error)
                                                    End Try

                                                    If sqlEx.Number = 1205 Then
                                                        TraceLog("*** DEADLOCK ***", sw, _enableTrace)
                                                        TraceLog(sw, _enableTrace)

                                                        ' Deadlock encountered - return false an wait for a minute
                                                        transaction.Dispose()
                                                        transaction = Nothing
                                                        If Not reader.IsClosed Then
                                                            reader.Close()
                                                        End If
                                                        reader = Nothing
                                                        Return False
                                                    Else
                                                        LogEvent(String.Format("Sql Error in transaction :- {0}.", sqlEx.Message), EventLogEntryType.Error)
                                                    End If

                                                Catch ex As Exception
                                                    doOutlookOK = False
                                                    LogEvent(String.Format("Error in transaction :- {0}.", ex.Message), EventLogEntryType.Error)
                                                End Try
                                            End Using
                                        End If

                                        TraceLog("doOutlookOK : " & doOutlookOK.ToString(), sw, _enableTrace)

                                        If doOutlookOK AndAlso Not deleted Then
                                            Using connSP As New SqlConnection(_openHRSystems(iOpenHRSystem).ConnectionString)
                                                ' Lets call spASRNetOutlookBatch to get us the rest of our properties
                                                Dim cmdSP As New SqlCommand("spASRNetOutlookBatch", connSP)
                                                With cmdSP
                                                    .Connection.Open()
                                                    .CommandType = CommandType.StoredProcedure

                                                    ' Input/Output parameters
                                                    .Parameters.Add("@Content", SqlDbType.VarChar, 8000).Direction = ParameterDirection.InputOutput
                                                    .Parameters("@Content").Value = content

                                                    ' Output parameters
                                                    .Parameters.Add("@AllDayEvent", SqlDbType.Bit).Direction = ParameterDirection.Output
                                                    .Parameters.Add("@StartDate", SqlDbType.DateTime).Direction = ParameterDirection.Output
                                                    .Parameters.Add("@EndDate", SqlDbType.DateTime).Direction = ParameterDirection.Output
                                                    .Parameters.Add("@StartTime", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                                                    .Parameters.Add("@EndTime", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                                                    .Parameters.Add("@Subject", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                                                    .Parameters.Add("@Folder", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                                                    ' Input parameters
                                                    .Parameters.AddWithValue("@LinkID", linkID)
                                                    .Parameters.AddWithValue("@RecordID", recordID)
                                                    .Parameters.AddWithValue("@FolderID", folderID)
                                                    .Parameters.AddWithValue("@StartDateColumnID", startDateColumnID)
                                                    .Parameters.AddWithValue("@EndDateColumnID", endDateColumnID)
                                                    .Parameters.AddWithValue("@FixedStartTime", fixedStartTime)
                                                    .Parameters.AddWithValue("@FixedEndTime", fixedEndTime)
                                                    .Parameters.AddWithValue("@StartTimeColumnID", startTimeColumnID)
                                                    .Parameters.AddWithValue("@EndTimeColumnID", endTimeColumnID)
                                                    .Parameters.AddWithValue("@TimeRange", timeRange)
                                                    .Parameters.AddWithValue("@Title", title)
                                                    .Parameters.AddWithValue("@SubjectExprID", subjectExprID)
                                                    .Parameters.AddWithValue("@RecordDescExprID", recordDescExprID)
                                                    .Parameters.AddWithValue("@DateFormat", dateFormat)
                                                    .Parameters.AddWithValue("@FolderPath", folderPath)
                                                    .Parameters.AddWithValue("@FolderType", folderType)
                                                    .Parameters.AddWithValue("@FolderExprID", folderExprID)

                                                    .ExecuteNonQuery()

                                                    ' Retrieve output parameters
                                                    content = Utilities.NullSafeString(.Parameters("@Content").Value)
                                                    allDayEvent = Utilities.NullSafeBoolean(.Parameters("@AllDayEvent").Value)

                                                    If Not (.Parameters("@StartDate").Value Is DBNull.Value) Then
                                                        startDate = Convert.ToDateTime(.Parameters("@StartDate").Value)
                                                    End If

                                                    If Not (.Parameters("@EndDate").Value Is DBNull.Value) Then
                                                        endDate = Convert.ToDateTime(.Parameters("@EndDate").Value)
                                                    Else
                                                        endDate = startDate
                                                    End If

                                                    If Date.Compare(startDate, endDate) > 0 Then
                                                        endDate = startDate
                                                    End If

                                                    startTime = Utilities.NullSafeString(.Parameters("@StartTime").Value)
                                                    endTime = Utilities.NullSafeString(.Parameters("@EndTime").Value)
                                                    subject = Utilities.NullSafeString(.Parameters("@Subject").Value)
                                                    folder = Utilities.NullSafeString(.Parameters("@Folder").Value)

                                                    .Parameters.Clear()
                                                    .Connection.Close()
                                                End With
                                                cmdSP.Dispose()
                                                cmdSP = Nothing
                                            End Using

                                            ' Create the outlook appointment using the OutlookCalendar class
                                            With _outlookApp
                                                .Reminder = reminder
                                                .ReminderOffset = reminderOffset
                                                .ReminderPeriod = reminderPeriod
                                                .AllDayEvent = allDayEvent
                                                .StartDate = startDate
                                                .EndDate = endDate
                                                .StartTime = startTime
                                                .EndTime = endTime
                                                .Subject = subject
                                                .Content = content
                                                .BusyStatus = busyStatus
                                                .Folder = folder

                                                TraceLog("Create Entry : " & subject, sw, _enableTrace)
                                                createdEntry = .CreateEntry()
                                                If Not createdEntry Then
                                                    TraceLog("Create Entry : FAILED", sw, _enableTrace)
                                                    LogEvent(String.Format("Could not create entry {0} :- {1}.", subject, .ErrorMessage), _
                                                        EventLogEntryType.Warning)
                                                Else
                                                    TraceLog("StoreID : " & .StoreID, sw, _enableTrace)
                                                    TraceLog("EntryID : " & .EntryID, sw, _enableTrace)
                                                End If
                                            End With

                                            Try
                                                TraceLog("Update ASRSysOutlookEvents", sw, _enableTrace)

                                                ' Need to update the row in the events table
                                                Using connUpd As New SqlConnection(_openHRSystems(iOpenHRSystem).ConnectionString)
                                                    sql = "UPDATE ASRSysOutlookEvents WITH(ROWLOCK) "
                                                    sql &= "SET ErrorMessage = @ErrorMessage "
                                                    sql &= ", StoreID = @StoreID "
                                                    sql &= ", EntryID = @EntryID "
                                                    sql &= ", Refresh = 0 "
                                                    sql &= ", StartDate = @StartDate "
                                                    sql &= ", Subject = @Subject "
                                                    sql &= ", Folder = @Folder "
                                                    sql &= ", EndDate = @EndDate "
                                                    sql &= ", RefreshDate = GETDATE() "
                                                    sql &= "WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID"
                                                    connUpd.Open()
                                                    Dim cmdUpd As New SqlCommand(sql, connUpd)
                                                    With _outlookApp
                                                        cmdUpd.Parameters.AddWithValue("@ErrorMessage", .ErrorMessage)
                                                        cmdUpd.Parameters.AddWithValue("@StoreID", .StoreID)
                                                        cmdUpd.Parameters.AddWithValue("@EntryID", .EntryID)
                                                        cmdUpd.Parameters.AddWithValue("@StartDate", .StartDate)
                                                        cmdUpd.Parameters.AddWithValue("@Subject", .Subject)
                                                        cmdUpd.Parameters.AddWithValue("@Folder", .Folder)
                                                        cmdUpd.Parameters.AddWithValue("@EndDate", .EndDate)
                                                        cmdUpd.Parameters.AddWithValue("@LinkID", linkID)
                                                        cmdUpd.Parameters.AddWithValue("@FolderID", folderID)
                                                        cmdUpd.Parameters.AddWithValue("@RecordID", recordID)
                                                    End With
                                                    cmdUpd.ExecuteNonQuery()
                                                    cmdUpd.Parameters.Clear()
                                                    cmdUpd.Connection.Close()
                                                    cmdUpd.Dispose()
                                                    cmdUpd = Nothing
                                                End Using

                                            Catch sqlEx As SqlException
                                                Dim errDesc As String = CStr(IIf(sqlEx.Number = 1205, "deadlock", "error"))

                                                TraceLog("Update : FAILED", sw, _enableTrace)

                                                If _outlookApp.DeleteEntry() Then
                                                    TraceLog("Entry deleted", sw, _enableTrace)
                                                    LogEvent(String.Format("Sql {0} occured :- {1} set to retry.", errDesc, _outlookApp.Subject), EventLogEntryType.Warning)
                                                Else
                                                    TraceLog("FAILED to delete entry : " & _outlookApp.ErrorMessage, sw, _enableTrace)
                                                    LogEvent(String.Format("Sql {0} occurred in UPDATE {1}{1}{2} could not be deleted duplicates may occur.", errDesc, _
                                                                  ControlChars.NewLine, _outlookApp.Subject), EventLogEntryType.Error)
                                                End If

                                                If sqlEx.Number = 1205 Then
                                                    TraceLog("*** DEADLOCK ***", sw, _enableTrace)
                                                    If Not reader.IsClosed Then
                                                        reader.Close()
                                                    End If
                                                    reader = Nothing

                                                    TraceLog(sw, _enableTrace)
                                                    Return False
                                                End If
                                            End Try
                                        End If

                                        TraceLog(sw, _enableTrace)

                                    End While
                                Else
                                    ' No events to process
                                End If
                            End If
                            conn.Close()

                        End Using

                    Catch sqlEx As SqlException
                        TraceLog("*** UNHANDLED ERROR ***", sw, _enableTrace)
                        TraceLog("  - " & sqlEx.Message, sw, _enableTrace)
                        LogEvent(String.Format("SQL Error {1} occurred ({0}) :- {2}.", serverDBName, _
                             sqlEx.Number, sqlEx.Message), EventLogEntryType.Error)
                    Catch ex As Exception
                        Throw ex
                    End Try
                Next iOpenHRSystem

            End Using

        Catch ex As Exception
            LogEvent(ex.Message, ex.StackTrace, EventLogEntryType.Error)
            Return False
        Finally
            TraceLog("--------------", sw, _enableTrace)
            TraceLog("Batch Finished", sw, _enableTrace)
            TraceLog("--------------", sw, _enableTrace)

            If (sw IsNot Nothing) Then
                sw.Close()
                sw.Dispose()
            End If

            If (_outlookApp IsNot Nothing) Then
                _outlookApp.Quit()
                _outlookApp = Nothing
            End If
        End Try

        Return True

    End Function

    ''' <summary>Performs checks to see if the database is valid for servicing</summary>
    ''' <param name="conn">connection to the database</param>
    ''' <returns></returns>
    ''' <remarks>Boolean</remarks>
    Private Function DatabaseIsOK(ByRef openHRSystem As OpenHRSystem, ByVal conn As SqlConnection, ByVal sw As TextWriter) As Boolean
        Dim returnOK As Boolean = True

        ' Check if the given database is locked.
        TraceLog("Checking database is locked", sw, _enableTrace)
        returnOK = Not DatabaseIsLocked(openHRSystem, conn)

        If returnOK Then
            ' Check if the given database is in the middle of the overnight job update.
            TraceLog("Checking database is running overnight process", sw, _enableTrace)
            returnOK = Not DatabaseIsRunningOvernight(openHRSystem, conn)
        End If

        If returnOK Then
            ' Check if the given database is the correct version.
            TraceLog("Checking database version is okay", sw, _enableTrace)
            returnOK = DatabaseVersionIsOK(openHRSystem, conn)
        End If

        If returnOK Then
            ' Check to see if the specified database is already being service elsewhere
            TraceLog("Checking database is being serviced", sw, _enableTrace)
            returnOK = Not DatabaseIsBeingServiced(openHRSystem, conn)
        End If

        Return returnOK
    End Function

    ''' <summary>
    ''' Performs a check to see if the specified database is already being service elsewhere
    ''' </summary>
    ''' <param name="openHRSystem"></param>
    ''' <param name="conn"></param>
    ''' <returns>Boolean</returns>
    Private Function DatabaseIsBeingServiced(ByRef openHRSystem As OpenHRSystem, ByVal conn As SqlConnection) As Boolean
        Dim returnBool As Boolean = False
        Dim running As Boolean = False
        Dim lastRun As DateTime = CDate(SqlTypes.SqlDateTime.MinValue)
        Dim runServer As String = String.Empty

        Try
            running = CType(GetSystemSetting("outlook service", "running", conn), Int32) = 1

            If running Then
                Try
                    lastRun = CDate(GetSystemSetting("outlook service", "last run", conn))
                Catch ex As FormatException
                    ' Can't convert to a datetime so we'll stick with the MinValue
                End Try

                runServer = GetSystemSetting("outlook service", "server", conn)

                If (runServer = openHRSystem.ServiceServer) Then
                    SaveSystemSetting("outlook service", "running", "1", conn)
                    SaveSystemSetting("outlook service", "server", Environment.MachineName, conn)
                    SaveSystemSetting("outlook service", "last run", DateTime.Now.ToString(), conn)

                    openHRSystem.Serviced = True
                    returnBool = False

                ElseIf (lastRun <> SqlTypes.SqlDateTime.MinValue) AndAlso DateDiff(DateInterval.Minute, lastRun, DateTime.Now) >= 5 Then
                    SaveSystemSetting("outlook service", "running", "1", conn)
                    SaveSystemSetting("outlook service", "server", Environment.MachineName, conn)
                    SaveSystemSetting("outlook service", "last run", DateTime.Now.ToString(), conn)

                    LogEvent(String.Format("Database ({0}) now being serviced.", _
                      openHRSystem.ToString(), runServer), EventLogEntryType.Information)

                    openHRSystem.ServiceServer = Environment.MachineName
                    openHRSystem.Serviced = True
                    returnBool = False

                ElseIf openHRSystem.Serviced Then
                    LogEvent(String.Format("Database ({0}) already being serviced by server {1}.", openHRSystem.ToString(), runServer), _
                      EventLogEntryType.Warning)
                    openHRSystem.Serviced = False
                    returnBool = True

                End If
            Else
                If openHRSystem.Serviced = False Then
                    LogEvent(String.Format("Database ({0}) now being serviced.", _
                                openHRSystem.ToString(), runServer), EventLogEntryType.Information)
                End If

                SaveSystemSetting("outlook service", "running", "1", conn)
                SaveSystemSetting("outlook service", "server", Environment.MachineName, conn)
                SaveSystemSetting("outlook service", "last run", DateTime.Now.ToString(), conn)

                openHRSystem.ServiceServer = Environment.MachineName
                openHRSystem.Serviced = True
                returnBool = False
            End If

        Catch ex As Exception
            LogEvent(String.Format("DatabaseIsBeingServiced - {0}", ex.Message), EventLogEntryType.Error)
        End Try

        Return returnBool

    End Function

    ''' <summary>Performs a check to see if the specified database is locked</summary>
    ''' <param name="openHRSystem"></param>
    ''' <param name="conn"></param>
    ''' <returns>Boolean</returns>
    Private Function DatabaseIsLocked(ByRef openHRSystem As OpenHRSystem, ByVal conn As SqlConnection) As Boolean
        Dim lockCheck As SqlCommand = New SqlCommand
        Dim returnBool As Boolean = False

        Try
            lockCheck.CommandText = "sp_ASRLockCheck"
            lockCheck.Connection = conn
            lockCheck.CommandType = CommandType.StoredProcedure
            lockCheck.CommandTimeout = CInt(IIf((_CommandTimeout < 200) And (_CommandTimeout <> 0), 200, _CommandTimeout))

            Dim reader As SqlDataReader = lockCheck.ExecuteReader()
            Dim priority As Int32 = 0

            While reader.Read
                If Utilities.NullSafeInteger(reader("priority")) <> 3 Then
                    returnBool = True
                    Exit While
                End If
            End While

            reader.Close()
            reader = Nothing

            If Not openHRSystem.Locked AndAlso returnBool Then
                LogEvent(String.Format("Database ({0}) locked.", openHRSystem.ToString()))
            ElseIf openHRSystem.Locked AndAlso Not returnBool Then
                LogEvent(String.Format("Database ({0}) unlocked.", openHRSystem.ToString()))
            End If

            openHRSystem.Locked = returnBool

        Catch ex As Exception
            LogEvent(String.Format("Lock check ({0}) - {1}", openHRSystem.ToString(), ex.Message), _
                ex.StackTrace, EventLogEntryType.Error)
        Finally
            lockCheck.Dispose()
            lockCheck = Nothing
        End Try

        Return returnBool

    End Function

    ''' <summary>Performs a check to see if the specified database is locked</summary>
    ''' <param name="openHRSystem"></param>
    ''' <param name="conn"></param>
    ''' <returns>Boolean</returns>
    Private Function DatabaseIsRunningOvernight(ByRef openHRSystem As OpenHRSystem, ByVal conn As SqlConnection) As Boolean
        Dim overnightCheck As SqlCommand = New SqlCommand
        Dim returnBool As Boolean = False

        Try
            returnBool = CType(GetSystemSetting("database", "updatingdatedependantcolumns", conn), Int32) = 1

            If Not openHRSystem.Suspended AndAlso returnBool Then
                LogEvent(String.Format("Database ({0}) suspended.", openHRSystem.ToString()))
            ElseIf openHRSystem.Suspended AndAlso Not returnBool Then
                LogEvent(String.Format("Database ({0}) resumed.", openHRSystem.ToString()))
            End If

            openHRSystem.Suspended = returnBool

        Catch ex As Exception
            LogEvent(String.Format("Overnight Job check ({0}) - {1}", openHRSystem.ToString(), _
                ex.Message), ex.StackTrace, EventLogEntryType.Error)
        Finally
            overnightCheck.Dispose()
            overnightCheck = Nothing
        End Try

        Return returnBool

    End Function

    ''' <summary>Performs a check to see if the specified database version is correct</summary>
    ''' <param name="openHRSystem"></param>
    ''' <param name="conn"></param>
    ''' <returns>Boolean</returns>
    Private Function DatabaseVersionIsOK(ByRef openHRSystem As OpenHRSystem, ByVal conn As SqlConnection) As Boolean
        Dim returnBool As Boolean = False

        Dim minSvcVersion As VersionNumber = New VersionNumber
        Dim minSvcVersionString As String = GetSystemSetting("outlook service", "minimum version", conn)

        Try
            Dim start As Int32 = 0, length As Int32 = 0
            start = 0
            length = minSvcVersionString.IndexOf(".", start)
            minSvcVersion.Major = CInt(minSvcVersionString.Substring(start, length))

            start += length + 1
            length = (minSvcVersionString.IndexOf(".", start)) - start
            minSvcVersion.Minor = CInt(minSvcVersionString.Substring(start, length))

            start += length + 1
            If minSvcVersionString.IndexOf(".", start) = -1 Then
                minSvcVersion.Build = CInt(minSvcVersionString.Substring(start))
            Else
                length = (minSvcVersionString.IndexOf(".", start)) - start
                minSvcVersion.Build = CInt(minSvcVersionString.Substring(start, length))
            End If
        Catch ex As Exception
            minSvcVersion.Major = 0
            minSvcVersion.Minor = 0
            minSvcVersion.Build = 0
        End Try

        Dim svcVersionOK As Boolean = False
        If minSvcVersion.Major <= _svcVersion.Major _
          AndAlso minSvcVersion.Minor <= _svcVersion.Minor _
          AndAlso minSvcVersion.Build <= _svcVersion.Build Then

            svcVersionOK = True
        End If

        Dim dbVersionOK As Boolean = False
        Dim dbVersion As Single = 0
        dbVersion = CType(GetSystemSetting("database", "version", conn), Single)
        dbVersionOK = MINIMUMDBVERSION <= dbVersion

        returnBool = (svcVersionOK AndAlso dbVersionOK)

        If (openHRSystem.VersionOK AndAlso (Not returnBool)) Then
            LogEvent( _
              String.Format("Database ({0}) is incompatible with the Outlook Calendar Service ({1}). Contact your system administrator.", _
              openHRSystem.ToString(), _svcVersion.ToString()), EventLogEntryType.Warning)
        ElseIf (returnBool AndAlso (Not openHRSystem.VersionOK)) Then
            LogEvent(String.Format("Database ({0}) version incompatibility corrected.", openHRSystem.ToString()))
        End If

        openHRSystem.VersionOK = returnBool

        Return returnBool

    End Function

    ''' <summary>
    ''' SQL Query used to retrieve Outlook Events from OpenHR
    ''' </summary>
    ''' <returns>String</returns>
    Private Function GetOutlookEventsSQL() As String
        Dim sqlString As Text.StringBuilder = New Text.StringBuilder()
        Dim returnSQL As String = String.Empty

        Try
            With sqlString
                .Append("SELECT ASRSysOutlookEvents.LinkID,")
                .Append(" ASRSysOutlookEvents.FolderID,")
                .Append(" ASRSysOutlookEvents.TableID,")
                .Append(" ASRSysOutlookEvents.RecordID,")
                .Append(" ASRSysOutlookEvents.Refresh,")
                .Append(" ASRSysOutlookEvents.Deleted,")
                .Append(" ASRSysOutlookEvents.StoreID,")
                .Append(" ASRSysOutlookEvents.EntryID,")
                .Append(" ASRSysOutlookLinks.Title,")
                .Append(" ASRSysOutlookLinks.BusyStatus,")
                .Append(" ASRSysOutlookLinks.StartDate,")
                .Append(" ASRSysOutlookLinks.EndDate,")
                .Append(" ASRSysOutlookLinks.TimeRange,")
                .Append(" ASRSysOutlookLinks.FixedStartTime,")
                .Append(" ASRSysOutlookLinks.FixedEndTime,")
                .Append(" ASRSysOutlookLinks.ColumnStartTime,")
                .Append(" ASRSysOutlookLinks.ColumnEndTime,")
                .Append(" ASRSysOutlookLinks.Subject,")
                .Append(" ISNULL(ASRSysOutlookLinks.Content,'') [Content],")
                .Append(" ASRSysOutlookLinks.Reminder,")
                .Append(" ASRSysOutlookLinks.ReminderOffset,")
                .Append(" ASRSysOutlookLinks.ReminderPeriod,")
                .Append(" ASRSysOutlookFolders.FolderType,")
                .Append(" ASRSysOutlookFolders.FixedPath,")
                .Append(" ASRSysOutlookFolders.ExprID,")
                .Append(" ASRSysTables.RecordDescExprID ")
                .Append("FROM ASRSysOutlookEvents WITH(READPAST) ")
                .Append("LEFT OUTER JOIN ASRSysOutlookLinks WITH(READPAST)")
                .Append(" ON ASRSysOutlookEvents.LinkID = ASRSysOutlookLinks.LinkID ")
                .Append("LEFT OUTER JOIN ASRSysOutlookFolders WITH(READPAST)")
                .Append(" ON ASRSysOutlookEvents.FolderID = ASRSysOutlookFolders.FolderID ")
                .Append("LEFT OUTER JOIN ASRSysTables WITH(READPAST)")
                .Append(" ON ASRSysOutlookEvents.TableID = ASRSysTables.TableID ")
                .Append("WHERE(ASRSysOutlookEvents.Refresh = 1)")
                .Append(" OR ASRSysOutlookEvents.Deleted = 1 ")
                .Append("ORDER BY ASRSysOutlookEvents.RecordID")
            End With

            returnSQL = sqlString.ToString()
        Catch ex As Exception
            Throw (ex)
        Finally
            sqlString = Nothing
        End Try

        Return returnSQL

    End Function

    Private Function GetSystemSetting(ByVal section As String, ByVal key As String, ByVal conn As SqlConnection) As String
        Dim returnString As String = String.Empty
        Dim cmd As SqlCommand = New SqlCommand

        Try
            With cmd
                .CommandText = "spASRGetSetting"
                .Connection = conn
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = CInt(IIf((_CommandTimeout < 200) And (_CommandTimeout <> 0), 200, _CommandTimeout))

                .Parameters.Add("@psSection", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                .Parameters("@psSection").Value = section

                .Parameters.Add("@psKey", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                .Parameters("@psKey").Value = key

                .Parameters.Add("@psDefault", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                .Parameters("@psDefault").Value = "0"

                .Parameters.Add("@pfUserSetting", SqlDbType.Bit).Direction = ParameterDirection.Input
                .Parameters("@pfUserSetting").Value = False

                .Parameters.Add("@psResult", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                .ExecuteNonQuery()

                returnString = Utilities.NullSafeString(.Parameters("@psResult").Value)
            End With
        Catch ex As Exception
            LogEvent(String.Format("GetSystemSetting - {0}", ex.Message), ex.StackTrace, EventLogEntryType.Error)
        Finally
            cmd.Dispose()
            cmd = Nothing
        End Try

        Return returnString
    End Function

    Private Function SaveSystemSetting(ByVal section As String, ByVal key As String, _
      ByVal value As String, ByVal conn As SqlConnection) As Boolean

        Dim returnbool As Boolean = False
        Dim cmd As SqlCommand = New SqlCommand

        Try
            With cmd
                .CommandText = "spASRSaveSetting"
                .Connection = conn
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = CInt(IIf((_CommandTimeout < 200) And (_CommandTimeout <> 0), 200, _CommandTimeout))

                .Parameters.Add("@psSection", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                .Parameters("@psSection").Value = section

                .Parameters.Add("@psKey", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                .Parameters("@psKey").Value = key

                .Parameters.Add("@psValue", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
                .Parameters("@psValue").Value = value

                .ExecuteNonQuery()

                returnbool = True
            End With
        Catch ex As Exception
            LogEvent(String.Format("SaveSystemSetting - {0}", ex.Message), ex.StackTrace, EventLogEntryType.Error)
        Finally
            cmd.Dispose()
            cmd = Nothing
        End Try

        Return returnbool
    End Function

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        _svcVersion = New VersionNumber

        With _svcVersion
            .Major = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major
            .Minor = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor
            .Build = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build
        End With
    End Sub
End Class
