Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

''' <summary>OpenHR System object</summary>
Public Structure OpenHRSystem
    Dim ServerName As String
    Dim DatabaseName As String
    Dim DefaultServer As Boolean
    Dim Locked As Boolean
    Dim Suspended As Boolean
    Dim Serviced As Boolean
    Dim ServiceServer As String
    Dim VersionOK As Boolean
    Dim UserName As String
    Dim Password As String

    ''' <summary>Gets the ConnectionString for the given OpenHR System</summary>
    Public ReadOnly Property ConnectionString() As String
        Get
            Return ""
        End Get
    End Property

    ''' <summary>String representation of OpenHR Server in format ServerName.DatabaseName</summary>
    Public Overrides Function ToString() As String
        Return String.Concat(ServerName, ".", DatabaseName)
    End Function
End Structure

Public Structure VersionNumber
  Dim Major As Int32
  Dim Minor As Int32
  Dim Build As Int32

  Public Overrides Function ToString() As String
    If Major = 0 AndAlso Minor = 0 AndAlso Build = 0 Then
      Return "<unknown>"
    Else
      '  Return String.Format("v{0}.{1}.{2}", Major, Minor, Build)
      Return String.Format("{0}.{1}.{2}", Major, Minor, Build)
    End If
  End Function
End Structure

Module General
  ''' <summary>
    ''' Logs a message to the Advanced event log, if the severity is error 
  ''' then the StackTrace is added to the message
  ''' </summary>
  ''' <param name="LogMessage">Message to log in event entry</param>
  ''' <param name="Severity">Optional severity, assumes information if not specified</param>
  ''' <remarks>Change logName and source as appropriate</remarks>
  Public Sub LogEvent(ByVal LogMessage As String, Optional ByVal Severity As EventLogEntryType = EventLogEntryType.Information)

    Dim log As New EventLog()
        Dim logName As String = "Advanced Business Solutions"
        Dim source As String = "OpenHR Outlook Calendar Service"

    log.Log = logName
    log.Source = source
    log.WriteEntry(LogMessage, Severity)
        log.Close()
        log = Nothing

    End Sub

  Public Sub LogEvent(ByVal LogMessage As String, ByVal StackTrace As String, _
    Optional ByVal Severity As EventLogEntryType = EventLogEntryType.Information)

    Dim log As New EventLog()
        Dim logName As String = "Advanced Business Solutions"
        Dim source As String = "OpenHR Outlook Calendar Service"

    log.Log = logName
    log.Source = source
    log.WriteEntry(String.Format("{0} {1}", LogMessage, StackTrace), Severity)
        log.Close()
        log = Nothing

  End Sub

  Public Sub TraceLog(ByVal w As TextWriter, ByVal enabled As Boolean)
    ' Only output if trace is enabled
    If Not enabled Then Exit Sub

    w.Write(ControlChars.CrLf)
    w.WriteLine("-------------------------------")
    w.Write(ControlChars.CrLf)
    ' Update the underlying file.
    w.Flush()
  End Sub

  Public Sub TraceLog(ByVal logMessage As String, ByVal w As TextWriter, ByVal enabled As Boolean)
    ' Only output if trace is enabled
    If Not enabled Then Exit Sub

    w.Write("{0} {1}", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString())
    w.WriteLine("  : {0}", logMessage)
    ' Update the underlying file.
    w.Flush()
  End Sub
  
  ''' <summary>
  ''' Determines whether the specified process is running in the current user context and whether to kill it.
  ''' </summary>
  ''' <param name="processName">Process to look for</param>
  ''' <param name="killProcess">True to terminate the process if running</param>
  ''' <returns>Boolean</returns>
  Public Function IsProcessRunningInContext(ByVal processName As String, Optional ByVal killProcess As Boolean = False) As Boolean
    Dim processRunning As Boolean = False

    For Each proc As Process In Process.GetProcessesByName(processName.ToUpper())
      If proc.SessionId = Process.GetCurrentProcess.SessionId Then
        processRunning = True

        If killProcess Then
          proc.Kill()
        End If

        Exit For
      End If
    Next

    Return processRunning
  End Function

  Public Function FolderFromFileName _
     (ByVal FileFullPath As String) As String

    Dim intPos As Integer
    intPos = FileFullPath.LastIndexOfAny(CType("\", Char()))
    intPos += 1

    Return FileFullPath.Substring(0, intPos)

  End Function

   
End Module
