Imports System.ComponentModel
Imports System.Configuration.Install
Imports Microsoft.Win32

<RunInstaller(True)> Public Class ProjectInstaller
  Inherits System.Configuration.Install.Installer

#Region " Component Designer generated code "

  Public Sub New()
    MyBase.New()

        Try
            MsgBox(Me.Context.Parameters.Item("TXTDATABASE").ToString, MsgBoxStyle.OkOnly, "TXTDATABASE")
        Catch ex As Exception

        End Try


        'This call is required by the Component Designer.
    InitializeComponent()

    'Add initialization code after the call to InitializeComponent
    Dim logName As String = "Advanced Business Solutions"
    Dim source As String = "OpenHR Workflow Service"

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


    'Installer overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ServiceProcessInstaller1 As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents ServiceInstaller1 As System.ServiceProcess.ServiceInstaller
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ServiceProcessInstaller1 = New System.ServiceProcess.ServiceProcessInstaller
        Me.ServiceInstaller1 = New System.ServiceProcess.ServiceInstaller
        '
        'ServiceProcessInstaller1
        '
        Me.ServiceProcessInstaller1.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.ServiceProcessInstaller1.Password = Nothing
        Me.ServiceProcessInstaller1.Username = Nothing
        '
        'ServiceInstaller1
        '
    Me.ServiceInstaller1.Description = "Processes the OpenHR Workflow queue."
    Me.ServiceInstaller1.DisplayName = "OpenHR Workflow Service"
    Me.ServiceInstaller1.ServiceName = "OpenHR Workflow"
        Me.ServiceInstaller1.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessInstaller1, Me.ServiceInstaller1})

    End Sub

#End Region

End Class
