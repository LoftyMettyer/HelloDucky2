Imports System.ComponentModel

Public Class AuditLogForm

    Public ConString As String
    Private model As AuditLogModel

    Private Sub AuditLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        model = New AuditLogModel(New DBContext(ConString))

        BindControls()

        AddHandler Application.Idle,
           Sub()
               datePanel.Visible = (model.Period = DatePeriod.SpecificDates)
           End Sub

        model.Find()
    End Sub

    Private Sub BindControls()
        'usually all done automatically by ui framework
        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.ThisMonth, "This Month"))
        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.LastMonth, "Last Month"))
        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.SpecificDates, "Specific Dates"))
        userEditor.Items.Add(New Infragistics.Win.ValueListItem(Nothing, ""))
        userEditor.Items.AddRange(model.Users.Select(Function(u) New Infragistics.Win.ValueListItem(u)).ToArray)
        periodEditor.DataBindings.Add(New Binding("Value", model, "Period", True, DataSourceUpdateMode.OnPropertyChanged))
        dateFromEditor.DataBindings.Add(New Binding("Value", model, "DateFrom", True, DataSourceUpdateMode.OnPropertyChanged))
        dateToEditor.DataBindings.Add(New Binding("Value", model, "DateTo", True, DataSourceUpdateMode.OnPropertyChanged))
        userEditor.DataBindings.Add(New Binding("Text", model, "User", True, DataSourceUpdateMode.OnValidation))
        auditLogsGrid.DataSource = model.AuditLogs
        AddHandler findButton.Click,
            Sub()
                model.Find()
            End Sub
        AddHandler model.AuditLogsChanged,
            Sub()
                auditLogsGrid.DataSource = model.AuditLogs
            End Sub
    End Sub

    Private Sub auditLogsGrid_InitializeLayout1(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles auditLogsGrid.InitializeLayout
        e.Layout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.None
        e.Layout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.None
        e.Layout.Override.FilterUIType = Infragistics.Win.UltraWinGrid.FilterUIType.HeaderIcons
        e.Layout.Bands(0).Columns("OldValue").Header.Caption = "Old Value"
        e.Layout.Bands(0).Columns("NewValue").Header.Caption = "New Value"
        e.Layout.Bands(0).Columns("RecordDescription").Header.Caption = "Record Description"
    End Sub

    Private Sub butOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butOutput.Click
        Try
            gridExporter.Export(auditLogsGrid, txtFilePath.Text)
            MsgBox("The file has been created.", , Me.Text)

        Catch ex As System.IO.DirectoryNotFoundException
            MsgBox("The directory specified does not exist.", MsgBoxStyle.Exclamation, Me.Text)
        Catch ex As System.IO.IOException
            MsgBox("The file specified is being used by another application.", MsgBoxStyle.Exclamation, Me.Text)
        Catch ex As ArgumentException
            MsgBox("No filename specified.", MsgBoxStyle.Exclamation, Me.Text)
        Catch ex As Exception
            MsgBox("Export failed.", MsgBoxStyle.Exclamation, Me.Text)
        Finally

        End Try
    End Sub
End Class