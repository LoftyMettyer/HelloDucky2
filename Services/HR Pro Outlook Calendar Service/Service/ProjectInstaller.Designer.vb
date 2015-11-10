<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer. 
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ServiceProcessInstaller = New System.ServiceProcess.ServiceProcessInstaller
        Me.OpenHROutlookCalendarService = New System.ServiceProcess.ServiceInstaller
        '
        'ServiceProcessInstaller
        '
        Me.ServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.ServiceProcessInstaller.Password = Nothing
        Me.ServiceProcessInstaller.Username = Nothing
        '
        'OpenHROutlookCalendarService
        '
        Me.OpenHROutlookCalendarService.Description = "Provides OpenHR with the ability to perform operations in the Outlook calendar"
        Me.OpenHROutlookCalendarService.DisplayName = "OpenHR Outlook Calendar Service"
        Me.OpenHROutlookCalendarService.ServiceName = "OpenHR Outlook Calendar"
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessInstaller, Me.OpenHROutlookCalendarService})

    End Sub
    Friend WithEvents ServiceProcessInstaller As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents OpenHROutlookCalendarService As System.ServiceProcess.ServiceInstaller

End Class
