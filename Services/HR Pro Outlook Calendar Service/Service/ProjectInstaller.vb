Imports System.ComponentModel
Imports System.Configuration.Install
Imports Microsoft.Win32

Public Class ProjectInstaller

  Public Sub New()
    MyBase.New()

    'This call is required by the Component Designer.
    InitializeComponent()

    'Add initialization code after the call to InitializeComponent
        Dim logName As String = "Advanced Business Solutions"
        Dim source As String = "OpenHR Outlook Calendar Service"

    If Diagnostics.EventLog.SourceExists(source) Then
      Diagnostics.EventLog.DeleteEventSource(source)
    End If

    Dim SourceSettings As New EventSourceCreationData(source, logName)
    Diagnostics.EventLog.CreateEventSource(SourceSettings)

    Dim regKey As RegistryKey = Nothing
    Try
      regKey = Registry.LocalMachine.OpenSubKey( _
          String.Format("SYSTEM\CurrentControlSet\Services\EventLog\{0}", logName), True)

      Dim ver As Int32 = CInt(regKey.GetValue("Retention", 0))
      If ver <> 0 Then
        regKey.SetValue("Retention", 0)
      End If
    Finally
      regKey.Close()
      regKey = Nothing
    End Try
  End Sub

End Class
