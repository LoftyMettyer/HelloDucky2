Imports System.ServiceProcess
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Data

Public Class OpenHRWorkflowService
  Inherits System.ServiceProcess.ServiceBase

#Region " Component Designer generated code "

  Public Sub New()
    MyBase.New()

    ' This call is required by the Component Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call
    _svcVersion = New VersionNumber

    With _svcVersion
      .Major = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major
      .Minor = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor
      .Build = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build
    End With

  End Sub

  'UserService overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  ' The main entry point for the process
  <MTAThread()> _
  Shared Sub Main()
    Dim ServicesToRun() As System.ServiceProcess.ServiceBase

    ' More than one NT Service may run within the same process. To add
    ' another service to this process, change the following line to
    ' create a second service object. For example,
    '
    '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
    '
    ServicesToRun = New System.ServiceProcess.ServiceBase() {New OpenHRWorkflowService}

    System.ServiceProcess.ServiceBase.Run(ServicesToRun)
  End Sub

  'Required by the Component Designer
  Private components As System.ComponentModel.IContainer
  Private databasesAndServers(5, 0) As String
  Private miCommandTimeout As Integer
  Private mobjEventLog As clsEventLog = New clsEventLog("Advanced Business Solutions", "OpenHR Workflow Service")
  Private _svcVersion As VersionNumber
  Private Const MAXLOGENTRYLENGTH As Integer = 32766
  Private Const MINIMUMDBVERSION As Single = 3.7

  Public Structure VersionNumber
    Dim Major As Int32
    Dim Minor As Int32
    Dim Build As Int32

    Public Overrides Function ToString() As String
      If Major = 0 AndAlso Minor = 0 AndAlso Build = 0 Then
        Return "<unknown>"
      Else
        Return String.Format("v{0}.{1}.{2}", Major, Minor, Build)
      End If
    End Function
  End Structure

  ' NOTE: The following procedure is required by the Component Designer
  ' It can be modified using the Component Designer.  
  ' Do not modify it using the code editor.
  'Friend WithEvents EventLog1 As System.Diagnostics.EventLog
  Friend WithEvents Timer1 As System.Timers.Timer
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.Timer1 = New System.Timers.Timer()
    CType(Me.Timer1, System.ComponentModel.ISupportInitialize).BeginInit()
    '
    'Timer1
    '
    Me.Timer1.Enabled = True
    Me.Timer1.Interval = 5000.0R
    '
    'OpenHRWorkflowService
    '
    Me.AutoLog = False
    Me.CanPauseAndContinue = True
    Me.ServiceName = "OpenHR Workflow"
    CType(Me.Timer1, System.ComponentModel.ISupportInitialize).EndInit()

  End Sub

#End Region

  Protected Overrides Sub OnStart(ByVal args() As String)
    ' Add code here to start your service. This method should set things
    ' in motion so your service can do its work.
    Dim sServiceVersion As String
    Dim newConfigValues As AppSettingsSection
    Dim config As System.Configuration.Configuration
    Dim sTag As String
    Dim sValue As String
    Dim sCurrentServer As String
    Dim sServer As String
    Dim sEventLogEntry As String
    Dim filemap As ExeConfigurationFileMap
    Dim keyValue As KeyValueConfigurationElement

    Const SERVERTAG As String = "SERVER"
    Const DATABASETAG As String = "DATABASE"
    Const COMMANDTIMEOUTTAG As String = "COMMANDTIMEOUT"

    Try
      sCurrentServer = ""

      sServiceVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major.ToString _
       & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString _
       & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build.ToString

      If sServiceVersion.Length = 0 Then
        sServiceVersion = "unknown version"
        ' Else
        ' sServiceVersion = "v" & sServiceVersion
      End If

      sEventLogEntry = "Service (" & sServiceVersion & ") started."
      mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Information)

      ' Index 0 = server
      ' Index 1 = database
      ' Index 2 = last entry entered into the event log
      ' Index 3 = Y if currently being serviced
      ' Index 4 = Y if service resumed message required
      ' Index 5 = name of server servicing the db
      ReDim databasesAndServers(5, 0)
      databasesAndServers(0, 0) = ""
      databasesAndServers(1, 0) = ""
      databasesAndServers(2, 0) = ""
      databasesAndServers(3, 0) = ""
      databasesAndServers(4, 0) = ""
      databasesAndServers(5, 0) = ""

      miCommandTimeout = -1

      ' Try to open the custom config file.
      filemap = New ExeConfigurationFileMap()
      filemap.ExeConfigFilename = System.Reflection.Assembly.GetExecutingAssembly.Location.Replace(".exe", ".custom.config")

      Try
        config = ConfigurationManager.OpenMappedExeConfiguration(filemap, ConfigurationUserLevel.None)

        If Not config.HasFile Then
          config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        End If

        newConfigValues = config.AppSettings

        For Each keyValue In newConfigValues.Settings
          sTag = keyValue.Key.ToUpper.Trim
          sValue = keyValue.Value.Trim

          If sValue.Length > 0 Then

            If sTag.Substring(0, SERVERTAG.Length) = SERVERTAG.ToUpper Then
              ' Server key
              sCurrentServer = sValue
            ElseIf sTag.Substring(0, DATABASETAG.Length) = DATABASETAG.ToUpper Then
              ' Database key
              If sCurrentServer.Length = 0 Then
                sServer = "."
              Else
                sServer = sCurrentServer
              End If

              ReDim Preserve databasesAndServers(5, databasesAndServers.GetUpperBound(1) + 1)
              databasesAndServers(0, databasesAndServers.GetUpperBound(1)) = sServer
              databasesAndServers(1, databasesAndServers.GetUpperBound(1)) = sValue
              databasesAndServers(2, databasesAndServers.GetUpperBound(1)) = ""
              databasesAndServers(3, databasesAndServers.GetUpperBound(1)) = ""
              databasesAndServers(4, databasesAndServers.GetUpperBound(1)) = ""
              databasesAndServers(5, databasesAndServers.GetUpperBound(1)) = Environment.MachineName

              sEventLogEntry = "Service configured for " & _
                 databasesAndServers(0, databasesAndServers.GetUpperBound(1)) & "." & databasesAndServers(1, databasesAndServers.GetUpperBound(1)) & "."
              mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Information)
            ElseIf sTag.Substring(0, COMMANDTIMEOUTTAG.Length) = COMMANDTIMEOUTTAG.ToUpper Then
              ' Command Timeout key
              Try
                miCommandTimeout = CInt(sValue)
              Catch ex As Exception
              End Try
            End If
          End If
        Next

        If miCommandTimeout < 0 Then
          miCommandTimeout = 200
          sEventLogEntry = "The default command timeout of " & CStr(miCommandTimeout) & " seconds has been applied."
        ElseIf miCommandTimeout = 0 Then
          sEventLogEntry = "Indefinite command timeout configured."
        Else
          sEventLogEntry = "Command timeout configured to be " & CStr(miCommandTimeout) & " seconds."
        End If

        mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Information)

      Catch ex As Exception
        sEventLogEntry = ex.Message
        mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Error)
      End Try

      If databasesAndServers.GetUpperBound(1) = 0 Then
        sEventLogEntry = "No server and database parameters defined."
        mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Warning)
        Me.Stop()
      End If

    Catch theError As Exception
      sEventLogEntry = "OnStart - " & theError.Message
      mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Error)
    End Try
  End Sub

  Protected Overrides Sub OnStop()
    ' Add code here to perform any tear-down necessary to stop your service.
    Dim sEventLogEntry As String = String.Empty

    For iDBServerLoop As Int32 = databasesAndServers.GetLowerBound(1) To databasesAndServers.GetUpperBound(1)
      If databasesAndServers(5, iDBServerLoop) = Environment.MachineName Then
        Try
          Dim strConn As String = _
           "Application Name=OpenHR Workflow Service;data source=" & databasesAndServers(0, iDBServerLoop) & _
           ";initial catalog=" & databasesAndServers(1, iDBServerLoop) & _
           ";Integrated Security=SSPI;Pooling=false"

          Using conn As New SqlClient.SqlConnection(strConn)
            conn.Open()
            SaveSystemSetting("workflow service", "running", "0", conn)
            conn.Close()
          End Using
        Catch theError As SqlClient.SqlException
          sEventLogEntry = String.Format("Error clearing system parameters for database ({0}.{1}). {2}{3}", _
           databasesAndServers(0, iDBServerLoop), _
           databasesAndServers(1, iDBServerLoop), _
           ControlChars.NewLine, _
           theError.Message)
          mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Error)
        End Try
      End If
    Next

    Try
      sEventLogEntry = "Service stopped."
      mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Information)
    Catch theError As Exception
    End Try
  End Sub

  Protected Overrides Sub OnContinue()
    Dim sEventLogEntry As String

    Try
      sEventLogEntry = "Service continued."
      mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Information)
    Catch theError As Exception
    End Try

  End Sub

  Protected Overrides Sub OnPause()
    Dim sEventLogEntry As String

    Try
      sEventLogEntry = "Service paused."
      mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Information)
    Catch theError As Exception
    End Try

  End Sub

  Private Sub Timer1_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed
    Dim sEventLogEntry As String

    Try
      ' Action the non-StoredData steps
      ActionWorkflowSteps_Part1()

      ' Action the StoredData steps
      ActionWorkflowSteps_Part2()

      ' Initiate any triggered workflows
      InitiateTriggeredWorkflows()

    Catch theError As Exception
      sEventLogEntry = "Timer - " & theError.Message
      mobjEventLog.WriteEntry(sEventLogEntry, EventLogEntryType.Error)
    End Try

  End Sub

  Private Sub ActionWorkflowSteps_Part1()
    ' Run the stored procedure to action all workflow steps that 
    ' can be actioned using SQL on its own (ie. all steps except Stored Data steps).
    Dim strConn As String
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim cmdActionSteps As System.Data.SqlClient.SqlCommand
    Dim iDBServerLoop As Int16
    Dim sErrorMsg As String
    Dim sEventLogEntry As String
    Dim iLoop As Integer
    Dim blnFound As Boolean
    Dim asErrors(0) As String
    Dim sqlError As SqlClient.SqlError
    Dim blnGeneralError As Boolean

    For iDBServerLoop = databasesAndServers.GetLowerBound(1) To databasesAndServers.GetUpperBound(1)

      If (databasesAndServers(0, iDBServerLoop).Length > 0) And _
      (databasesAndServers(1, iDBServerLoop).Length > 0) And _
      (databasesAndServers(3, iDBServerLoop) <> "Y") Then

        sErrorMsg = ""
        blnGeneralError = False

        Try ' General block for each configured server & database
          databasesAndServers(3, iDBServerLoop) = "Y"

          strConn = "Application Name=OpenHR Workflow Service;data source=" & databasesAndServers(0, iDBServerLoop) & ";initial catalog=" & databasesAndServers(1, iDBServerLoop) & ";Integrated Security=SSPI;Pooling=false"
          conn = New SqlClient.SqlConnection(strConn)

          ' Instantiate command objects so that they can go into the Finally block without warnings
          cmdActionSteps = New SqlClient.SqlCommand

          Try ' conn creation block
            conn.Open()

            If DatabaseIsOK(iDBServerLoop, conn) Then
              With cmdActionSteps
                .CommandText = "spASRActionActiveWorkflowSteps"
                .Connection = conn
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)
                Try
                  .ExecuteNonQuery()

                Catch ex As SqlClient.SqlException
                  ReDim asErrors(0)
                  For Each sqlError In ex.Errors
                    blnFound = False

                    For iLoop = 0 To asErrors.GetUpperBound(0)
                      If asErrors(iLoop) = sqlError.Message Then
                        blnFound = True
                        Exit For
                      End If
                    Next iLoop

                    If Not blnFound Then
                      ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                      asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                      sErrorMsg = sErrorMsg & _
                       sqlError.Message & vbNewLine
                    End If
                  Next sqlError

                Finally
                  .Dispose()
                End Try

              End With
            End If

          Catch ex As Exception ' conn creation block
            sErrorMsg = ex.Message
            blnGeneralError = True

          Finally ' conn creation block
            If Not IsNothing(cmdActionSteps) Then cmdActionSteps.Dispose()
            If Not IsNothing(conn) Then conn.Close()
          End Try ' conn creation block

        Catch ex As Exception ' General block for each configured server & database
          sErrorMsg = ex.Message
          blnGeneralError = True

        Finally ' General block for each configured server & database
          databasesAndServers(3, iDBServerLoop) = ""

          conn = Nothing

          If blnGeneralError Then
            databasesAndServers(4, iDBServerLoop) = "Y"
          ElseIf databasesAndServers(4, iDBServerLoop) = "Y" Then
            sEventLogEntry = "Database (" & databasesAndServers(0, iDBServerLoop) & "." & databasesAndServers(1, iDBServerLoop) & ") service OK. "

            mobjEventLog.WriteEntry(sEventLogEntry, _
             EventLogEntryType.Information, _
             databasesAndServers(2, iDBServerLoop))

            databasesAndServers(2, iDBServerLoop) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
            databasesAndServers(4, iDBServerLoop) = ""
          End If

          If sErrorMsg.Length > 0 Then
            sEventLogEntry = "(" & databasesAndServers(0, iDBServerLoop) & "." & databasesAndServers(1, iDBServerLoop) & ") - " & sErrorMsg

            mobjEventLog.WriteEntry("Step 1 " & sEventLogEntry, _
             EventLogEntryType.Error, _
             "Step 1 " & databasesAndServers(2, iDBServerLoop))

            databasesAndServers(2, iDBServerLoop) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
          End If
        End Try ' General block for each configured server & database
      End If
    Next iDBServerLoop

  End Sub

  Private Sub ActionWorkflowSteps_Part2()
    ' Action all workflow steps that CANNOT be actioned using SQL 
    ' on its own (ie. just the Stored Data steps).
    Dim strConn As String
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim cmdDetails As System.Data.SqlClient.SqlCommand
    Dim cmdGetSteps As System.Data.SqlClient.SqlCommand
    Dim cmdSubmit As System.Data.SqlClient.SqlCommand
    Dim cmdSubmit2 As System.Data.SqlClient.SqlCommand
    Dim cmdAction As System.Data.SqlClient.SqlCommand
    Dim cmdFailure As System.Data.SqlClient.SqlCommand
    Dim cmdRecordDesc As System.Data.SqlClient.SqlCommand
    Dim dr As System.Data.SqlClient.SqlDataReader
    Dim iInstanceID As Integer
    Dim iElementID As Integer
    Dim blnOK As Boolean
    Dim aiActionedSteps(1, 0) As Integer
    Dim iLoop As Integer
    Dim iRecordID As Integer
    Dim iDataAction As Integer
    Dim iTableID As Integer
    Dim sTableName As String
    Dim sSQL As String
    Dim iDBServerLoop As Int16
    Dim sRecordDescription As String
    Dim sErrorMsg As String
    Dim sMessage As String
    Dim sqlError As SqlClient.SqlError
    Dim sEventLogEntry As String
    Dim iLoop2 As Integer
    Dim blnFound As Boolean
    Dim asErrors(0) As String
    Dim blnGeneralError As Boolean
    Dim iRetryCount As Integer
    Dim blnDeadlock As Boolean
    Dim iErrorNumber As Integer
    Dim fSQLResult As Boolean

    Const DEADLOCKERRORNUMBER As Integer = 1205
    Const MAXRETRIES As Integer = 5
    Const PAUSE As Integer = 5000

    For iDBServerLoop = databasesAndServers.GetLowerBound(1) To databasesAndServers.GetUpperBound(1)

      If (databasesAndServers(0, iDBServerLoop).Length > 0) And _
      (databasesAndServers(1, iDBServerLoop).Length > 0) And _
      (databasesAndServers(3, iDBServerLoop) <> "Y") Then

        sErrorMsg = ""
        blnGeneralError = False

        Try ' General block for each configured server & database
          databasesAndServers(3, iDBServerLoop) = "Y"

          strConn = "Application Name=OpenHR Workflow Service;data source=" & databasesAndServers(0, iDBServerLoop) & ";initial catalog=" & databasesAndServers(1, iDBServerLoop) & ";Integrated Security=SSPI;Pooling=false"
          conn = New SqlClient.SqlConnection(strConn)

          ' Instantiate command objects so that they can go into the Finally block without warnings
          cmdGetSteps = New SqlClient.SqlCommand

          Try ' conn creation block
            conn.Open()

            If DatabaseIsOK(iDBServerLoop, conn) Then

              ReDim aiActionedSteps(1, 0)

              cmdGetSteps.CommandText = "spASRGetActiveWorkflowStoredDataSteps"
              cmdGetSteps.Connection = conn
              cmdGetSteps.CommandType = CommandType.StoredProcedure
              cmdGetSteps.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

              Try
                dr = cmdGetSteps.ExecuteReader()

                While dr.Read
                  ReDim Preserve aiActionedSteps(1, UBound(aiActionedSteps, 2) + 1)
                  aiActionedSteps(0, UBound(aiActionedSteps, 2)) = dr("instanceID")
                  aiActionedSteps(1, UBound(aiActionedSteps, 2)) = dr("elementID")
                End While

                dr.Close()

              Catch ex As SqlClient.SqlException
                ReDim asErrors(0)

                For Each sqlError In ex.Errors
                  blnFound = False

                  For iLoop2 = 0 To asErrors.GetUpperBound(0)
                    If asErrors(iLoop2) = sqlError.Message Then
                      blnFound = True
                      Exit For
                    End If
                  Next iLoop2

                  If Not blnFound Then
                    ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                    asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                    sErrorMsg = sErrorMsg & _
                     sqlError.Message & vbNewLine
                  End If
                Next sqlError

                ReDim aiActionedSteps(1, 0)

              Finally
                cmdGetSteps.Dispose()
              End Try

              For iLoop = 1 To UBound(aiActionedSteps, 2)
                blnOK = True
                iInstanceID = aiActionedSteps(0, iLoop)
                iElementID = aiActionedSteps(1, iLoop)
                iRecordID = 0
                sSQL = ""
                iTableID = 0
                sTableName = ""
                iDataAction = 0
                iRecordID = 0
                sRecordDescription = ""
                sMessage = ""

                cmdDetails = New SqlClient.SqlCommand
                cmdAction = New SqlClient.SqlCommand
                cmdAction.CommandTimeout = IIf((miCommandTimeout < 600) And (miCommandTimeout <> 0), 600, miCommandTimeout)

                cmdSubmit = New SqlClient.SqlCommand
                cmdSubmit2 = New SqlClient.SqlCommand
                cmdFailure = New SqlClient.SqlCommand
                cmdRecordDesc = New SqlClient.SqlCommand

                Try ' Action step block
                  ' Get the StoredData action details.
                  cmdDetails.CommandText = "spASRGetStoredDataActionDetails"
                  cmdDetails.Connection = conn
                  cmdDetails.CommandType = CommandType.StoredProcedure
                  cmdDetails.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

                  cmdDetails.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                  cmdDetails.Parameters("@piInstanceID").Value = iInstanceID

                  cmdDetails.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                  cmdDetails.Parameters("@piElementID").Value = iElementID

                  cmdDetails.Parameters.Add("@psSQL", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output
                  cmdDetails.Parameters.Add("@piDataTableID", SqlDbType.Int).Direction = ParameterDirection.Output
                  cmdDetails.Parameters.Add("@psTableName", SqlDbType.VarChar, 255).Direction = ParameterDirection.Output
                  cmdDetails.Parameters.Add("@piDataAction", SqlDbType.Int).Direction = ParameterDirection.Output
                  cmdDetails.Parameters.Add("@piRecordID", SqlDbType.Int).Direction = ParameterDirection.Output
									cmdDetails.Parameters.Add("@bUseAsTargetIdentifier", SqlDbType.Bit).Direction = ParameterDirection.Output
									cmdDetails.Parameters.Add("@pfResult", SqlDbType.Bit).Direction = ParameterDirection.Output

                  Try
                    cmdDetails.ExecuteNonQuery()

                    sSQL = cmdDetails.Parameters("@psSQL").Value
                    iTableID = cmdDetails.Parameters("@piDataTableID").Value
                    sTableName = cmdDetails.Parameters("@psTableName").Value
                    iDataAction = cmdDetails.Parameters("@piDataAction").Value
                    iRecordID = cmdDetails.Parameters("@piRecordID").Value
                    fSQLResult = cmdDetails.Parameters("@pfResult").Value

                  Catch ex As SqlClient.SqlException
                    ReDim asErrors(0)

                    For Each sqlError In ex.Errors
                      blnFound = False

                      For iLoop2 = 0 To asErrors.GetUpperBound(0)
                        If asErrors(iLoop2) = sqlError.Message Then
                          blnFound = True
                          Exit For
                        End If
                      Next iLoop2

                      If Not blnFound Then
                        ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                        asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                        sErrorMsg = sErrorMsg & _
                         sqlError.Message & vbNewLine
                      End If
                    Next sqlError

                  Finally
                    cmdDetails.Dispose()
                  End Try

                  ' Run the required Stored Data action
                  ' If Len(sSQL) > 0 Then
                  If fSQLResult = True Then
                    ' Execute the StoredData action
                    Select Case iDataAction
                      Case 0 ' Insert
                        cmdAction.CommandText = "sp_ASRInsertNewRecord"

                        cmdAction.Parameters.Add("@piNewRecordID", SqlDbType.Int).Direction = ParameterDirection.Output

                        cmdAction.Parameters.Add("@psInsertString", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@psInsertString").Value = sSQL

                      Case 1 ' Update
                        cmdAction.CommandText = "sp_ASRUpdateRecord"

                        cmdAction.Parameters.Add("@piResult", SqlDbType.Int).Direction = ParameterDirection.Output

                        cmdAction.Parameters.Add("@psUpdateString", SqlDbType.NVarChar, 2147483646).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@psUpdateString").Value = sSQL

                        cmdAction.Parameters.Add("@piTableID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@piTableID").Value = iTableID

                        cmdAction.Parameters.Add("@psRealSource", SqlDbType.NVarChar, 128).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@psRealSource").Value = sTableName

                        cmdAction.Parameters.Add("@piID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@piID").Value = iRecordID

                        cmdAction.Parameters.Add("@piTimestamp", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@piTimestamp").Value = DBNull.Value

                      Case 2 ' Delete
                        ' Get the record description 
                        ' Must do this before performing the DataAction as the record will be deleted.
                        cmdRecordDesc.CommandText = "spASRRecordDescription"
                        cmdRecordDesc.Connection = conn
                        cmdRecordDesc.CommandType = CommandType.StoredProcedure
                        cmdRecordDesc.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

                        cmdRecordDesc.Parameters.Add("@piTableID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdRecordDesc.Parameters("@piTableID").Value = iTableID

                        cmdRecordDesc.Parameters.Add("@piRecordID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdRecordDesc.Parameters("@piRecordID").Value = iRecordID

                        cmdRecordDesc.Parameters.Add("@psRecordDescription", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

                        Try
                          cmdRecordDesc.ExecuteNonQuery()

                          sRecordDescription = cmdRecordDesc.Parameters("@psRecordDescription").Value
                          cmdRecordDesc.Dispose()

                        Catch ex As SqlClient.SqlException
                          ReDim asErrors(0)

                          For Each sqlError In ex.Errors
                            blnFound = False

                            For iLoop2 = 0 To asErrors.GetUpperBound(0)
                              If asErrors(iLoop2) = sqlError.Message Then
                                blnFound = True
                                Exit For
                              End If
                            Next iLoop2

                            If Not blnFound Then
                              ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                              asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                              sErrorMsg = sErrorMsg & _
                               sqlError.Message & vbNewLine
                            End If
                          Next sqlError
                        End Try

                        ' Perform the DataAction - delete the record
                        cmdAction.CommandText = "sp_ASRDeleteRecord"

                        cmdAction.Parameters.Add("@piResult", SqlDbType.Int).Direction = ParameterDirection.Output

                        cmdAction.Parameters.Add("@piTableID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@piTableID").Value = iTableID

                        cmdAction.Parameters.Add("@psRealSource", SqlDbType.NVarChar, 128).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@psRealSource").Value = sTableName

                        cmdAction.Parameters.Add("@piID", SqlDbType.Int).Direction = ParameterDirection.Input
                        cmdAction.Parameters("@piID").Value = iRecordID

                      Case Else
                        blnOK = False
                        sMessage = "Unrecognised data action."
                    End Select

                    If blnOK Then
                      cmdAction.Connection = conn
                      cmdAction.CommandType = CommandType.StoredProcedure

                      iRetryCount = 0
                      blnDeadlock = True

                      Do While blnDeadlock
                        Try
                          blnDeadlock = False
                          iErrorNumber = 0

                          cmdAction.ExecuteNonQuery()

                          If iDataAction = 0 Then
                            ' Insert
                            iRecordID = cmdAction.Parameters("@piNewRecordID").Value
                          End If

                        Catch ex As SqlClient.SqlException
                          ReDim asErrors(0)

                          For Each sqlError In ex.Errors
                            ' Ignore the 'The transaction ended in the trigger. The batch has been aborted.' errors.
                            If sqlError.Number = DEADLOCKERRORNUMBER Then
                              If (iRetryCount < MAXRETRIES) And (ex.Errors.Count = 1) Then
                                iRetryCount = iRetryCount + 1
                                blnDeadlock = True
                                ' Pause before resubmitting the SQL command.
                                System.Threading.Thread.Sleep(PAUSE)
                              Else
                                blnOK = False
                                blnFound = False

                                For iLoop2 = 0 To asErrors.GetUpperBound(0)
                                  If asErrors(iLoop2) = sqlError.Message Then
                                    blnFound = True
                                    Exit For
                                  End If
                                Next iLoop2

                                If Not blnFound Then
                                  ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                                  asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                                  sMessage = sMessage & _
                                   sqlError.Message & vbNewLine

                                  sErrorMsg = sErrorMsg & _
                                   sqlError.Message & vbNewLine
                                End If
                              End If

                            ElseIf sqlError.Number <> 3609 Then
                              blnOK = False
                              blnFound = False

                              For iLoop2 = 0 To asErrors.GetUpperBound(0)
                                If asErrors(iLoop2) = sqlError.Message Then
                                  blnFound = True
                                  Exit For
                                End If
                              Next iLoop2

                              If Not blnFound Then
                                ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                                asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                                sMessage = sMessage & _
                                 sqlError.Message & vbNewLine

                                sErrorMsg = sErrorMsg & _
                                 sqlError.Message & vbNewLine
                              End If
                            End If
                          Next sqlError
                        End Try
                      Loop
                    End If

                    cmdAction.Dispose()

                    If blnOK _
                     And ((iDataAction = 0) Or (iDataAction = 1)) Then

                      cmdSubmit.CommandText = "spASRStoredDataFileActions"
                      cmdSubmit.Connection = conn
                      cmdSubmit.CommandType = CommandType.StoredProcedure
                      cmdSubmit.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

                      cmdSubmit.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                      cmdSubmit.Parameters("@piInstanceID").Value = iInstanceID

                      cmdSubmit.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                      cmdSubmit.Parameters("@piElementID").Value = iElementID

                      cmdSubmit.Parameters.Add("@piRecordID", SqlDbType.Int).Direction = ParameterDirection.Input
                      cmdSubmit.Parameters("@piRecordID").Value = iRecordID

                      Try
                        cmdSubmit.ExecuteNonQuery()

                      Catch ex As SqlClient.SqlException
                        ReDim asErrors(0)

                        For Each sqlError In ex.Errors
                          blnFound = False

                          For iLoop2 = 0 To asErrors.GetUpperBound(0)
                            If asErrors(iLoop2) = sqlError.Message Then
                              blnFound = True
                              Exit For
                            End If
                          Next iLoop2

                          If Not blnFound Then
                            ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                            asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                            sErrorMsg = sErrorMsg & _
                             sqlError.Message & vbNewLine
                          End If
                        Next sqlError
                      End Try

                      cmdSubmit.Dispose()
                    End If

                    If blnOK Then
                      ' Stored Data action succeeded
                      cmdSubmit2.CommandText = "spASRSubmitWorkflowStep"
                      cmdSubmit2.Connection = conn
                      cmdSubmit2.CommandType = CommandType.StoredProcedure
                      cmdSubmit2.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

                      cmdSubmit2.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                      cmdSubmit2.Parameters("@piInstanceID").Value = iInstanceID

                      cmdSubmit2.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                      cmdSubmit2.Parameters("@piElementID").Value = iElementID

                      cmdSubmit2.Parameters.Add("@psFormInput1", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
                      cmdSubmit2.Parameters("@psFormInput1").Value = CStr(iRecordID) & vbTab & sRecordDescription

                      cmdSubmit2.Parameters.Add("@psFormElements", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output
                      cmdSubmit2.Parameters.Add("@pfSavedForLater", SqlDbType.Bit).Direction = ParameterDirection.Output

                      cmdSubmit2.Parameters.Add("@piPageNo", SqlDbType.Int).Direction = ParameterDirection.Input
                      cmdSubmit2.Parameters("@piPageNo").Value = 0

                      Try
                        cmdSubmit2.ExecuteNonQuery()

                      Catch ex As SqlClient.SqlException
                        ReDim asErrors(0)

                        For Each sqlError In ex.Errors
                          blnFound = False

                          For iLoop2 = 0 To asErrors.GetUpperBound(0)
                            If asErrors(iLoop2) = sqlError.Message Then
                              blnFound = True
                              Exit For
                            End If
                          Next iLoop2

                          If Not blnFound Then
                            ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                            asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                            sErrorMsg = sErrorMsg & _
                             sqlError.Message & vbNewLine
                          End If
                        Next sqlError
                      End Try

                      cmdSubmit2.Dispose()
                    Else
                      ' Stored Data action failed. 
                      cmdFailure.CommandText = "spASRWorkflowActionFailed"
                      cmdFailure.Connection = conn
                      cmdFailure.CommandType = CommandType.StoredProcedure
                      cmdFailure.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

                      cmdFailure.Parameters.Add("@piInstanceID", SqlDbType.Int).Direction = ParameterDirection.Input
                      cmdFailure.Parameters("@piInstanceID").Value = iInstanceID

                      cmdFailure.Parameters.Add("@piElementID", SqlDbType.Int).Direction = ParameterDirection.Input
                      cmdFailure.Parameters("@piElementID").Value = iElementID

                      cmdFailure.Parameters.Add("@psMessage", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
                      cmdFailure.Parameters("@psMessage").Value = sMessage

                      Try
                        cmdFailure.ExecuteNonQuery()

                      Catch ex As SqlClient.SqlException
                        ReDim asErrors(0)

                        For Each sqlError In ex.Errors
                          blnFound = False

                          For iLoop2 = 0 To asErrors.GetUpperBound(0)
                            If asErrors(iLoop2) = sqlError.Message Then
                              blnFound = True
                              Exit For
                            End If
                          Next iLoop2

                          If Not blnFound Then
                            ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                            asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                            sErrorMsg = sErrorMsg & _
                             sqlError.Message & vbNewLine
                          End If
                        Next sqlError

                      Finally
                        cmdFailure.Dispose()
                      End Try
                    End If
                  End If

                Catch ex As Exception ' Action step block
                  sErrorMsg = ex.Message
                  blnGeneralError = True

                Finally ' Action step block
                  If cmdDetails IsNot Nothing Then cmdDetails.Dispose()
                  If cmdAction IsNot Nothing Then cmdAction.Dispose()
                  If cmdSubmit IsNot Nothing Then cmdSubmit.Dispose()
                  If cmdSubmit2 IsNot Nothing Then cmdSubmit.Dispose()
                  If cmdFailure IsNot Nothing Then cmdFailure.Dispose()
                  If cmdRecordDesc IsNot Nothing Then cmdRecordDesc.Dispose()
                End Try ' Action step block
              Next iLoop
            End If

          Catch ex As Exception ' conn creation block
            sErrorMsg = ex.Message
            blnGeneralError = True

          Finally ' conn creation block
            If cmdGetSteps IsNot Nothing Then cmdGetSteps.Dispose()
            If conn IsNot Nothing Then conn.Close()
          End Try ' conn creation block

        Catch ex As Exception ' General block for each configured server & database
          sErrorMsg = ex.Message
          blnGeneralError = True

        Finally ' General block for each configured server & database
          databasesAndServers(3, iDBServerLoop) = ""

          conn = Nothing

          If blnGeneralError Then
            databasesAndServers(4, iDBServerLoop) = "Y"
          ElseIf databasesAndServers(4, iDBServerLoop) = "Y" Then
            sEventLogEntry = "Database (" & databasesAndServers(0, iDBServerLoop) & "." & databasesAndServers(1, iDBServerLoop) & ") service OK. "

            mobjEventLog.WriteEntry(sEventLogEntry, _
             EventLogEntryType.Information, _
             databasesAndServers(2, iDBServerLoop))

            databasesAndServers(2, iDBServerLoop) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
            databasesAndServers(4, iDBServerLoop) = ""
          End If

          If sErrorMsg.Length > 0 Then
            sEventLogEntry = "(" & databasesAndServers(0, iDBServerLoop) & "." & databasesAndServers(1, iDBServerLoop) & ") - " & sErrorMsg

            mobjEventLog.WriteEntry("Step 2 " & sEventLogEntry, _
             EventLogEntryType.Error, _
             "Step 2 " & databasesAndServers(2, iDBServerLoop))

            databasesAndServers(2, iDBServerLoop) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
          End If
        End Try ' General block for each configured server & database
      End If
    Next iDBServerLoop
  End Sub

  Private Sub InitiateTriggeredWorkflows()
    ' Run the stored procedure to initiate any triggered workflows 
    Dim strConn As String
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim cmdInstantiateTriggered As System.Data.SqlClient.SqlCommand
    Dim iDBServerLoop As Int16
    Dim sErrorMsg As String
    Dim sEventLogEntry As String
    Dim iLoop As Integer
    Dim blnFound As Boolean
    Dim asErrors(0) As String
    Dim sqlError As SqlClient.SqlError
    Dim blnGeneralError As Boolean

    For iDBServerLoop = databasesAndServers.GetLowerBound(1) To databasesAndServers.GetUpperBound(1)

      If (databasesAndServers(0, iDBServerLoop).Length > 0) And _
      (databasesAndServers(1, iDBServerLoop).Length > 0) And _
      (databasesAndServers(3, iDBServerLoop) <> "Y") Then

        sErrorMsg = ""
        blnGeneralError = False

        Try ' General block for each configured server & database
          databasesAndServers(3, iDBServerLoop) = "Y"

          strConn = "Application Name=OpenHR Workflow Service;data source=" & databasesAndServers(0, iDBServerLoop) & ";initial catalog=" & databasesAndServers(1, iDBServerLoop) & ";Integrated Security=SSPI;Pooling=false"
          conn = New SqlClient.SqlConnection(strConn)

          ' Instantiate command objects so that they can go into the Finally block without warnings
          cmdInstantiateTriggered = New SqlClient.SqlCommand

          Try ' conn creation block
            conn.Open()

            If DatabaseIsOK(iDBServerLoop, conn) Then
              cmdInstantiateTriggered.CommandText = "spASRInstantiateTriggeredWorkflows"
              cmdInstantiateTriggered.Connection = conn
              cmdInstantiateTriggered.CommandType = CommandType.StoredProcedure
              cmdInstantiateTriggered.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

              Try
                cmdInstantiateTriggered.ExecuteNonQuery()

              Catch ex As SqlClient.SqlException
                ReDim asErrors(0)
                For Each sqlError In ex.Errors
                  blnFound = False

                  For iLoop = 0 To asErrors.GetUpperBound(0)
                    If asErrors(iLoop) = sqlError.Message Then
                      blnFound = True
                      Exit For
                    End If
                  Next iLoop

                  If Not blnFound Then
                    ReDim Preserve asErrors(asErrors.GetUpperBound(0) + 1)
                    asErrors(asErrors.GetUpperBound(0)) = sqlError.Message

                    sErrorMsg = sErrorMsg & _
                     sqlError.Message & vbNewLine
                  End If
                Next sqlError

              Finally
                cmdInstantiateTriggered.Dispose()
              End Try
            End If

          Catch ex As Exception ' conn creation block
            sErrorMsg = ex.Message
            blnGeneralError = True

          Finally ' conn creation block
            If Not IsNothing(cmdInstantiateTriggered) Then cmdInstantiateTriggered.Dispose()
            If Not IsNothing(conn) Then conn.Close()
          End Try ' conn creation block

        Catch ex As Exception ' General block for each configured server & database
          sErrorMsg = ex.Message
          blnGeneralError = True

        Finally ' General block for each configured server & database
          databasesAndServers(3, iDBServerLoop) = ""

          If blnGeneralError Then
            databasesAndServers(4, iDBServerLoop) = "Y"
          ElseIf databasesAndServers(4, iDBServerLoop) = "Y" Then
            sEventLogEntry = "Database (" & databasesAndServers(0, iDBServerLoop) & "." & databasesAndServers(1, iDBServerLoop) & ") service OK. "

            mobjEventLog.WriteEntry(sEventLogEntry, _
             EventLogEntryType.Information, _
             databasesAndServers(2, iDBServerLoop))

            databasesAndServers(2, iDBServerLoop) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
            databasesAndServers(4, iDBServerLoop) = ""
          End If

          If (sErrorMsg.Length > 0) Then
            sEventLogEntry = "(" & databasesAndServers(0, iDBServerLoop) & "." & databasesAndServers(1, iDBServerLoop) & ") - " & sErrorMsg

            mobjEventLog.WriteEntry("Initiate Triggered Workflows " & sEventLogEntry, _
             EventLogEntryType.Error, _
             "Initiate Triggered Workflows " & databasesAndServers(2, iDBServerLoop))

            databasesAndServers(2, iDBServerLoop) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
          End If
        End Try ' General block for each configured server & database
      End If
    Next iDBServerLoop

  End Sub

  Private Function DatabaseIsInOvernightJob(ByVal piDBServerIndex As Integer, _
  ByVal pConn As System.Data.SqlClient.SqlConnection) As Boolean

    ' Check if the workflow service has been suspended for the given Server/Database.
    Dim cmdCheck As System.Data.SqlClient.SqlCommand
    Dim blnOvernightJob As Boolean
    Dim sEventLogEntry As String
    Dim sLastEventLogEntry As String
    Dim sSuspendedMsg As String

    blnOvernightJob = False
    sSuspendedMsg = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") overnight job in progress."
    sLastEventLogEntry = databasesAndServers(2, piDBServerIndex)
    cmdCheck = New SqlClient.SqlCommand

    Try
      cmdCheck.CommandText = "spASRGetSetting"
      cmdCheck.Connection = pConn
      cmdCheck.CommandType = CommandType.StoredProcedure
      cmdCheck.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

      cmdCheck.Parameters.Add("@psSection", SqlDbType.VarChar, 255).Direction = ParameterDirection.Input
      cmdCheck.Parameters("@psSection").Value = "database"

      cmdCheck.Parameters.Add("@psKey", SqlDbType.VarChar, 255).Direction = ParameterDirection.Input
      cmdCheck.Parameters("@psKey").Value = "updatingdatedependantcolumns"

      cmdCheck.Parameters.Add("@psDefault", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmdCheck.Parameters("@psDefault").Value = "0"

      cmdCheck.Parameters.Add("@pfUserSetting", SqlDbType.Bit).Direction = ParameterDirection.Input
      cmdCheck.Parameters("@pfUserSetting").Value = False

      cmdCheck.Parameters.Add("@psResult", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

      cmdCheck.ExecuteNonQuery()

      blnOvernightJob = (cmdCheck.Parameters("@psResult").Value = "1")
      If blnOvernightJob Then
        sEventLogEntry = sSuspendedMsg

        mobjEventLog.WriteEntry(sEventLogEntry, _
         EventLogEntryType.Information, _
         databasesAndServers(2, piDBServerIndex))

        databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      ElseIf (sLastEventLogEntry = sSuspendedMsg) Then
        sEventLogEntry = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") overnight job completed."

        mobjEventLog.WriteEntry(sEventLogEntry, _
         EventLogEntryType.Information, _
         databasesAndServers(2, piDBServerIndex))

        databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      End If

      cmdCheck.Dispose()

    Catch ex As Exception
      sEventLogEntry = "Overnight Job Check (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") - " & ex.Message

      mobjEventLog.WriteEntry(sEventLogEntry, _
       EventLogEntryType.Error, _
       databasesAndServers(2, piDBServerIndex))

      databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      databasesAndServers(4, piDBServerIndex) = "Y"

    Finally
      If Not IsNothing(cmdCheck) Then cmdCheck.Dispose()
    End Try

    DatabaseIsInOvernightJob = blnOvernightJob

  End Function

  Private Function DatabaseIsSuspended(ByVal piDBServerIndex As Integer, _
  ByVal pConn As System.Data.SqlClient.SqlConnection) As Boolean

    ' Check if the workflow service has been suspended for the given Server/Database.
    Dim cmdSuspendedCheck As System.Data.SqlClient.SqlCommand
    Dim blnSuspended As Boolean
    Dim sEventLogEntry As String
    Dim sLastEventLogEntry As String
    Dim sSuspendedMsg As String

    blnSuspended = False
    sSuspendedMsg = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") service suspended."
    sLastEventLogEntry = databasesAndServers(2, piDBServerIndex)
    cmdSuspendedCheck = New SqlClient.SqlCommand

    Try
      cmdSuspendedCheck.CommandText = "spASRGetSetting"
      cmdSuspendedCheck.Connection = pConn
      cmdSuspendedCheck.CommandType = CommandType.StoredProcedure
      cmdSuspendedCheck.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

      cmdSuspendedCheck.Parameters.Add("@psSection", SqlDbType.VarChar, 255).Direction = ParameterDirection.Input
      cmdSuspendedCheck.Parameters("@psSection").Value = "workflow"

      cmdSuspendedCheck.Parameters.Add("@psKey", SqlDbType.VarChar, 255).Direction = ParameterDirection.Input
      cmdSuspendedCheck.Parameters("@psKey").Value = "suspended"

      cmdSuspendedCheck.Parameters.Add("@psDefault", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
      cmdSuspendedCheck.Parameters("@psDefault").Value = "0"

      cmdSuspendedCheck.Parameters.Add("@pfUserSetting", SqlDbType.Bit).Direction = ParameterDirection.Input
      cmdSuspendedCheck.Parameters("@pfUserSetting").Value = False

      cmdSuspendedCheck.Parameters.Add("@psResult", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output

      cmdSuspendedCheck.ExecuteNonQuery()

      blnSuspended = (cmdSuspendedCheck.Parameters("@psResult").Value = "1")
      If blnSuspended Then
        sEventLogEntry = sSuspendedMsg

        mobjEventLog.WriteEntry(sEventLogEntry, _
         EventLogEntryType.Information, _
         databasesAndServers(2, piDBServerIndex))

        databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      ElseIf (sLastEventLogEntry = sSuspendedMsg) Then
        sEventLogEntry = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") service resumed."

        mobjEventLog.WriteEntry(sEventLogEntry, _
         EventLogEntryType.Information, _
         databasesAndServers(2, piDBServerIndex))

        databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      End If

      cmdSuspendedCheck.Dispose()

    Catch ex As Exception
      sEventLogEntry = "Suspension Check (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") - " & ex.Message

      mobjEventLog.WriteEntry(sEventLogEntry, _
       EventLogEntryType.Error, _
       databasesAndServers(2, piDBServerIndex))

      databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      databasesAndServers(4, piDBServerIndex) = "Y"

    Finally
      If Not IsNothing(cmdSuspendedCheck) Then cmdSuspendedCheck.Dispose()
    End Try

    DatabaseIsSuspended = blnSuspended

  End Function
  'Private Function DatabaseIsWrongVersion(ByVal piDBServerIndex As Integer, _
  'ByVal pConn As System.Data.SqlClient.SqlConnection) As Boolean

  '	' Check if the given database and website versions match.
  '	Dim cmdVersionCheck As System.Data.SqlClient.SqlCommand
  '	Dim blnWrongVersion As Boolean
  '	Dim sEventLogEntry As String
  '	Dim sDBVersion As String
  '	Dim sDBMsgVersion As String
  '	Dim sServiceVersion As String
  '	Dim sServiceMsgVersion As String
  '	Dim sLastEventLogEntry As String
  '	Dim sWrongVersionMsg As String

  '	blnWrongVersion = False
  '	sLastEventLogEntry = databasesAndServers(2, piDBServerIndex)
  '	cmdVersionCheck = New SqlClient.SqlCommand

  '	Try
  '		cmdVersionCheck.CommandText = "spASRGetSetting"
  '		cmdVersionCheck.Connection = pConn
  '		cmdVersionCheck.CommandType = CommandType.StoredProcedure

  '		cmdVersionCheck.Parameters.Add("@psSection", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
  '		cmdVersionCheck.Parameters("@psSection").Value = "database"

  '		cmdVersionCheck.Parameters.Add("@psKey", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
  '		cmdVersionCheck.Parameters("@psKey").Value = "version"

  '		cmdVersionCheck.Parameters.Add("@psDefault", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input
  '		cmdVersionCheck.Parameters("@psDefault").Value = ""

  '		cmdVersionCheck.Parameters.Add("@pfUserSetting", SqlDbType.Bit).Direction = ParameterDirection.Input
  '		cmdVersionCheck.Parameters("@pfUserSetting").Value = False

  '		cmdVersionCheck.Parameters.Add("@psResult", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

  '		cmdVersionCheck.ExecuteNonQuery()

  '		sDBVersion = cmdVersionCheck.Parameters("@psResult").Value
  '		sDBVersion = sDBVersion.Trim.ToUpper
  '		If sDBVersion.Length = 0 Then
  '			sDBMsgVersion = "<unknown>"
  '		Else
  '			sDBMsgVersion = "v" & sDBVersion
  '		End If

  '		sServiceVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString
  '		If sServiceVersion.Length = 0 Then
  '			sServiceVersion = "<unknown>"
  '		Else
  '			sServiceVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major & _
  '			 "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor
  '			sServiceVersion = sServiceVersion.Trim.ToUpper
  '		End If

  '		sServiceMsgVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major.ToString _
  '		& "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString _
  '		& "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build.ToString

  '		If sServiceMsgVersion.Length = 0 Then
  '			sServiceMsgVersion = "unknown version"
  '		Else
  '			sServiceMsgVersion = "v" & sServiceMsgVersion
  '		End If

  '		cmdVersionCheck.Dispose()

  '		sWrongVersionMsg = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ", " & sDBMsgVersion & ") " & _
  '	 "is incompatible with the Workflow service (" & sServiceMsgVersion & ")." & _
  '	 " Contact your system administrator."

  '		If (sDBVersion <> sServiceVersion) _
  '		 Or (sDBVersion.Length = 0) Then

  '			' Version mismatch.
  '			blnWrongVersion = True

  '			sEventLogEntry = sWrongVersionMsg

  '			mobjEventLog.WriteEntry(sEventLogEntry, _
  '			 EventLogEntryType.Warning, _
  '			 databasesAndServers(2, piDBServerIndex))

  '			databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
  '		ElseIf (sLastEventLogEntry = sWrongVersionMsg) Then
  '			sEventLogEntry = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") version incompatibility corrected."

  '			mobjEventLog.WriteEntry(sEventLogEntry, _
  '			 EventLogEntryType.Information, _
  '			 databasesAndServers(2, piDBServerIndex))

  '			databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
  '		End If

  '	Catch ex As Exception
  '		sEventLogEntry = "Version Check (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") - " & ex.Message

  '		mobjEventLog.WriteEntry(sEventLogEntry, _
  '		 EventLogEntryType.Error, _
  '		 databasesAndServers(2, piDBServerIndex))

  '		databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
  '		databasesAndServers(4, piDBServerIndex) = "Y"

  '	Finally
  '		If Not IsNothing(cmdVersionCheck) Then cmdVersionCheck.Dispose()
  '	End Try

  '	DatabaseIsWrongVersion = blnWrongVersion

  'End Function

  Private Function DatabaseIsWrongVersion(ByVal piDBServerIndex As Integer, _
   ByVal pConn As SqlClient.SqlConnection) As Boolean

    Dim returnBool As Boolean = False

    Dim minSvcVersion As VersionNumber = New VersionNumber
    Dim minSvcVersionString As String = GetSystemSetting("workflow service", "minimum version", pConn)

    Dim lastEventLogEntry As String = databasesAndServers(2, piDBServerIndex)

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
    dbVersion = CType(GetSystemSetting("database", "version", pConn), Single)
    dbVersionOK = MINIMUMDBVERSION <= dbVersion

    returnBool = Not (svcVersionOK AndAlso dbVersionOK)

    Dim wrongVersionMsg As String = _
     String.Format("Database ({0}.{1}) is incompatible with the Workflow service ({2}). Contact your system administrator.", _
      databasesAndServers(0, piDBServerIndex), databasesAndServers(1, piDBServerIndex), _svcVersion.ToString())

    If returnBool Then
      mobjEventLog.WriteEntry(wrongVersionMsg, _
      EventLogEntryType.Warning, _
        databasesAndServers(2, piDBServerIndex))

      databasesAndServers(2, piDBServerIndex) = wrongVersionMsg.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
    ElseIf (lastEventLogEntry = wrongVersionMsg) Then
      Dim eventLogEntry As String = _
       String.Format("Database ({0}.{1}) version incompatibility corrected.", _
       databasesAndServers(0, piDBServerIndex), _
       databasesAndServers(1, piDBServerIndex))

      mobjEventLog.WriteEntry(eventLogEntry, _
         EventLogEntryType.Information, _
           databasesAndServers(2, piDBServerIndex))

      databasesAndServers(2, piDBServerIndex) = eventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
    End If

    Return returnBool

  End Function

  Private Function IsServiceAccountValid(ByRef piDBServerIndex As Integer, ByVal pConn As SqlClient.SqlConnection) As Boolean

    ' Check if the workflow service has sufficient privileges
    Dim cmdServiceCheck As SqlClient.SqlCommand
    Dim sMessage As String
    Dim bIsValid As Boolean = False

    Dim sEventLogEntry As String
    Dim sLastEventLogEntry As String

    sMessage = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") CE004 - Unable to connect."
    sLastEventLogEntry = databasesAndServers(2, piDBServerIndex)
    cmdServiceCheck = New SqlClient.SqlCommand

    Try
      cmdServiceCheck.CommandText = "spASRWorkflowValidateService"
      cmdServiceCheck.Connection = pConn
      cmdServiceCheck.CommandType = CommandType.StoredProcedure
      cmdServiceCheck.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

      cmdServiceCheck.Parameters.Add("@allow", SqlDbType.Bit).Direction = ParameterDirection.Output

      cmdServiceCheck.ExecuteNonQuery()

      bIsValid = CBool(cmdServiceCheck.Parameters("@allow").Value)
      If Not bIsValid Then
        sEventLogEntry = sMessage

        mobjEventLog.WriteEntry(sEventLogEntry,
         EventLogEntryType.Information,
         databasesAndServers(2, piDBServerIndex))

        databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim

      End If

      cmdServiceCheck.Dispose()

    Catch ex As Exception
      sEventLogEntry = "Service Check (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") - " & ex.Message

      mobjEventLog.WriteEntry(sEventLogEntry,
       EventLogEntryType.Error,
       databasesAndServers(2, piDBServerIndex))

      databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      databasesAndServers(4, piDBServerIndex) = "Y"

    Finally
      If Not IsNothing(cmdServiceCheck) Then cmdServiceCheck.Dispose()
    End Try

    Return bIsValid

  End Function

  Private Function DatabaseIsBeingServiced(ByRef piDBServerIndex As Integer, ByVal conn As SqlClient.SqlConnection) As Boolean
    Dim returnBool As Boolean = False
    Dim running As Boolean = False
    Dim lastRun As DateTime = CDate(SqlTypes.SqlDateTime.MinValue)
    Dim runServer As String = String.Empty
    Dim logEntry As String = String.Empty

    Try
      Dim lastEventLogEntry As String = databasesAndServers(2, piDBServerIndex)

      running = CType(GetSystemSetting("workflow service", "running", conn), Int32) = 1

      If running Then
        Try
          lastRun = CDate(GetSystemSetting("workflow service", "last run", conn))
        Catch ex As Exception
          ' Can't convert to a datetime so we'll stick with the MinValue
        End Try

        runServer = GetSystemSetting("workflow service", "server", conn)

        If (runServer = databasesAndServers(5, piDBServerIndex)) Then
          SaveSystemSetting("workflow service", "running", "1", conn)
          SaveSystemSetting("workflow service", "server", Environment.MachineName, conn)
          SaveSystemSetting("workflow service", "last run", DateTime.Now.ToString(), conn)

          returnBool = False

        ElseIf (lastRun <> SqlTypes.SqlDateTime.MinValue) AndAlso DateDiff(DateInterval.Minute, lastRun, DateTime.Now) >= 5 Then
          SaveSystemSetting("workflow service", "running", "1", conn)
          SaveSystemSetting("workflow service", "server", Environment.MachineName, conn)
          SaveSystemSetting("workflow service", "last run", DateTime.Now.ToString(), conn)

          logEntry = String.Format("Database ({0}.{1}) now being serviced.", _
           databasesAndServers(0, piDBServerIndex), databasesAndServers(1, piDBServerIndex))

          mobjEventLog.WriteEntry(logEntry, _
           EventLogEntryType.Information, _
           databasesAndServers(2, piDBServerIndex))

          databasesAndServers(2, piDBServerIndex) = logEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
          databasesAndServers(5, piDBServerIndex) = Environment.MachineName
          returnBool = False

        Else
          logEntry = String.Format("Database ({0}.{1}) already being serviced by server {2}.", _
           databasesAndServers(0, piDBServerIndex), databasesAndServers(1, piDBServerIndex), runServer)

          mobjEventLog.WriteEntry(logEntry, _
           EventLogEntryType.Warning, _
           databasesAndServers(2, piDBServerIndex))

          databasesAndServers(2, piDBServerIndex) = logEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
          databasesAndServers(5, piDBServerIndex) = Environment.MachineName
          returnBool = True
        End If
      Else
        SaveSystemSetting("workflow service", "running", "1", conn)
        SaveSystemSetting("workflow service", "server", Environment.MachineName, conn)
        SaveSystemSetting("workflow service", "last run", DateTime.Now.ToString(), conn)

        If lastEventLogEntry <> String.Empty Then
          logEntry = String.Format("Database ({0}.{1}) now being serviced.", _
            databasesAndServers(0, piDBServerIndex), databasesAndServers(1, piDBServerIndex))

          mobjEventLog.WriteEntry(logEntry, _
           EventLogEntryType.Information, _
           databasesAndServers(2, piDBServerIndex))
          databasesAndServers(2, piDBServerIndex) = logEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
        End If

        databasesAndServers(5, piDBServerIndex) = Environment.MachineName
        returnBool = False
      End If

    Catch ex As Exception
      logEntry = "Serviced Check (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") - " & ex.ToString

      mobjEventLog.WriteEntry(logEntry, _
       EventLogEntryType.Error, _
       databasesAndServers(2, piDBServerIndex))

      databasesAndServers(2, piDBServerIndex) = logEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      databasesAndServers(4, piDBServerIndex) = "Y"
    End Try

    Return returnBool

  End Function

  Private Function DatabaseIsLocked(ByVal piDBServerIndex As Integer, _
   ByVal pConn As System.Data.SqlClient.SqlConnection) As Boolean

    ' Check if the database is locked for the given Server/Database.
    Dim cmdLockCheck As System.Data.SqlClient.SqlCommand
    Dim dr As System.Data.SqlClient.SqlDataReader
    Dim blnLocked As Boolean
    Dim sLockedMsg As String
    Dim sEventLogEntry As String
    Dim sLastEventLogEntry As String

    blnLocked = False
    sLockedMsg = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") locked."
    sLastEventLogEntry = databasesAndServers(2, piDBServerIndex)
    cmdLockCheck = New SqlClient.SqlCommand

    Try
      cmdLockCheck.CommandText = "sp_ASRLockCheck"
      cmdLockCheck.Connection = pConn
      cmdLockCheck.CommandType = CommandType.StoredProcedure
      cmdLockCheck.CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

      dr = cmdLockCheck.ExecuteReader()

      While dr.Read
        If dr("priority") <> 3 Then
          ' Not a read-only lock.
          blnLocked = True

          sEventLogEntry = sLockedMsg

          mobjEventLog.WriteEntry(sEventLogEntry, _
           EventLogEntryType.Information, _
           databasesAndServers(2, piDBServerIndex))

          databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim

          Exit While
        End If
      End While

      dr.Close()
      cmdLockCheck.Dispose()

      If (Not blnLocked) _
       And (sLastEventLogEntry = sLockedMsg) Then
        sEventLogEntry = "Database (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") unlocked."

        mobjEventLog.WriteEntry(sEventLogEntry, _
         EventLogEntryType.Information, _
         databasesAndServers(2, piDBServerIndex))

        databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      End If

    Catch ex As Exception
      sEventLogEntry = "Lock Check (" & databasesAndServers(0, piDBServerIndex) & "." & databasesAndServers(1, piDBServerIndex) & ") - " & ex.Message

      mobjEventLog.WriteEntry(sEventLogEntry, _
       EventLogEntryType.Error, _
       databasesAndServers(2, piDBServerIndex))

      databasesAndServers(2, piDBServerIndex) = sEventLogEntry.PadRight(MAXLOGENTRYLENGTH).Substring(0, MAXLOGENTRYLENGTH).Trim
      databasesAndServers(4, piDBServerIndex) = "Y"
    Finally
      If Not IsNothing(cmdLockCheck) Then cmdLockCheck.Dispose()
    End Try

    DatabaseIsLocked = blnLocked

  End Function

  Private Function DatabaseIsOK(ByVal piDBServerIndex As Integer, _
   ByVal pConn As System.Data.SqlClient.SqlConnection) As Boolean

    ' Check if the given database is OK to service.
    Dim blnOK As Boolean

    ' Check if the given database has the service suspended.
    blnOK = Not DatabaseIsSuspended(piDBServerIndex, pConn)

    If blnOK Then
      ' Check if the given database is the correct version.
      blnOK = Not DatabaseIsWrongVersion(piDBServerIndex, pConn)
    End If

    If blnOK Then
      ' Check if the given database is locked.
      blnOK = Not DatabaseIsLocked(piDBServerIndex, pConn)
    End If

    If blnOK Then
      ' Check if the given database is in the middle of the overnight job update.
      blnOK = Not DatabaseIsInOvernightJob(piDBServerIndex, pConn)
    End If

    If blnOK Then
      ' Check if the service account is valid.
      blnOK = Not IsServiceAccountValid(piDBServerIndex, pConn)
    End If

    If blnOK Then
      ' Check to see if the specified database is already being service elsewhere
      blnOK = Not DatabaseIsBeingServiced(piDBServerIndex, pConn)
    End If

    DatabaseIsOK = blnOK

  End Function

  Private Function GetSystemSetting(ByVal section As String, ByVal key As String, ByVal conn As SqlClient.SqlConnection) As String
    Dim returnString As String = String.Empty
    Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand

    Try
      With cmd
        .CommandText = "spASRGetSetting"
        .Connection = conn
        .CommandType = CommandType.StoredProcedure
        .CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

        .Parameters.Add("@psSection", SqlDbType.VarChar, 255).Direction = ParameterDirection.Input
        .Parameters("@psSection").Value = section

        .Parameters.Add("@psKey", SqlDbType.VarChar, 255).Direction = ParameterDirection.Input
        .Parameters("@psKey").Value = key

        .Parameters.Add("@psDefault", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Input
        .Parameters("@psDefault").Value = "0"

        .Parameters.Add("@pfUserSetting", SqlDbType.Bit).Direction = ParameterDirection.Input
        .Parameters("@pfUserSetting").Value = False

        .Parameters.Add("@psResult", SqlDbType.VarChar, 2147483646).Direction = ParameterDirection.Output
        .ExecuteNonQuery()

        returnString = .Parameters("@psResult").Value
      End With
    Catch ex As Exception
      mobjEventLog.WriteEntry(String.Format("GetSystemSetting - {0}", ex.Message), _
        EventLogEntryType.Error)
    Finally
      cmd.Dispose()
      cmd = Nothing
    End Try

    Return returnString
  End Function

  Private Function SaveSystemSetting(ByVal section As String, ByVal key As String, _
   ByVal value As String, ByVal conn As SqlClient.SqlConnection) As Boolean

    Dim returnbool As Boolean = False
    Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand

    Try
      With cmd
        .CommandText = "spASRSaveSetting"
        .Connection = conn
        .CommandType = CommandType.StoredProcedure
        .CommandTimeout = IIf((miCommandTimeout < 200) And (miCommandTimeout <> 0), 200, miCommandTimeout)

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
      mobjEventLog.WriteEntry(String.Format("SaveSystemSetting - {0}", ex.Message), _
        EventLogEntryType.Error)
    Finally
      cmd.Dispose()
      cmd = Nothing
    End Try

    Return returnbool
  End Function
End Class
