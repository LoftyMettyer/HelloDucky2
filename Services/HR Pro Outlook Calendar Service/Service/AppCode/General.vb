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

    ''' <summary>Gets the ConnectionString for the given OpenHR System</summary>
    Public ReadOnly Property ConnectionString() As String
        Get
            Return GetConnectionString(String.Empty, String.Empty, DatabaseName, ServerName)
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
      Return String.Format("v{0}.{1}.{2}", Major, Minor, Build)
    End If
  End Function
End Structure

Module General
  ''' <summary>Returns the version of the service</summary>
  Public Function GetVersion() As String
    Return System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major.ToString _
                & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString _
                & "." & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build.ToString
  End Function

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

  ''' <summary>Creates a database connection string based on specified parameters</summary>
  ''' <param name="userID"></param>
  ''' <param name="password"></param>
  ''' <param name="databaseName"></param>
  ''' <param name="serverName"></param>
  ''' <returns>string</returns>
  Public Function GetConnectionString(ByVal userID As String, ByVal password As String, _
      ByVal databaseName As String, ByVal serverName As String) As String

    Dim builder As New SqlConnectionStringBuilder()
    builder.DataSource = serverName
    builder.InitialCatalog = databaseName
    builder.PacketSize = 32767
    builder.ApplicationName = Reflection.Assembly.GetExecutingAssembly.GetName.Name
    builder.Pooling = False

    If Not userID.Equals(String.Empty) Then
      builder.UserID = userID
      builder.Password = password
    Else
      builder.IntegratedSecurity = True
    End If

    Return builder.ConnectionString

  End Function

  ''' <summary>Returns a date formatted with milliseconds</summary>
  ''' <param name="dDate"></param>
  ''' <returns>string</returns>
  Public Function FormatDateTimeWithMS(ByVal dDate As Date) As String

    Dim tempString As String

    tempString = dDate.ToString("yyyy-MM-dd hh:mm:ss") & ":" & dDate.Millisecond

    Return tempString

  End Function

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

    ''' <summary>Decrypts the OpenHR logon details</summary>
  ''' <param name="input"></param>
  ''' <param name="userName"></param>
  ''' <param name="password"></param>
  ''' <param name="database"></param>
  ''' <param name="server"></param>
  ''' <remarks></remarks>
  Public Sub DecryptLogonDetails(ByVal input As String, ByRef userName As String, ByRef password As String, _
      ByRef database As String, ByRef server As String)

    Dim eKey As String = String.Empty
    Dim lens As String = String.Empty
    Dim start As Integer = 0
    Dim finish As Integer = 0

    If input = String.Empty Then
      Return
    End If

    start = input.Length - 14
    eKey = input.Substring(start, 10)
    lens = input.Substring(input.Length - 4)
    input = XOREncript(input.Substring(0, start), eKey)

    start = 0
    finish = Asc(lens.Substring(0, 1)) - 127
    userName = input.Substring(start, finish)

    start = start + finish
    finish = Asc(lens.Substring(1, 1)) - 127
    password = input.Substring(start, finish)

    start = start + finish
    finish = Asc(lens.Substring(2, 1)) - 127
    database = input.Substring(start, finish)

    start = start + finish
    finish = Asc(lens.Substring(3, 1)) - 127
    server = input.Substring(start, finish)

  End Sub

    ''' <summary>Used as part of the OpenHR decryption method</summary>
  ''' <param name="input"></param>
  ''' <param name="key"></param>
  ''' <returns>string</returns>
  Public Function XOREncript(ByVal input As String, ByVal key As String) As String

    Dim count As Integer = 0
    Dim output As String = String.Empty
    Dim strChar As String = String.Empty

    For count = 1 To input.Length
      strChar = key.Substring(count Mod key.Length, 1)
      output = output & Chr(Asc(strChar) Xor Asc(input.Substring(count - 1, 1)))
    Next

    Return output

  End Function

  ''' <summary>Converts a byte array to a readable hex string</summary>
  ''' <param name="arrInput"></param>
  ''' <returns>string</returns>
  Public Function ByteArrayToString(ByVal arrInput() As Byte) As String
    Dim i As Integer = 0
    Dim sOutput As New Text.StringBuilder()
    sOutput.Append("0x")
    For i = 0 To arrInput.Length - 1
      sOutput.Append(arrInput(i).ToString("X2"))
    Next
    Return sOutput.ToString()
  End Function

  Public Function FolderFromFileName _
     (ByVal FileFullPath As String) As String

    Dim intPos As Integer
    intPos = FileFullPath.LastIndexOfAny(CType("\", Char()))
    intPos += 1

    Return FileFullPath.Substring(0, intPos)

  End Function

  Public Function NameOnlyFromFullPath _
    (ByVal FileFullPath As String) As String

    Dim intPos As Integer

    intPos = FileFullPath.LastIndexOfAny(CType("\", Char()))
    intPos += 1

    Return FileFullPath.Substring(intPos, (FileFullPath.Length - intPos))

  End Function
End Module
