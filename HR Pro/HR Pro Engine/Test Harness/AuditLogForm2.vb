Imports System.ComponentModel

Public Class AuditLogForm2

    Private model As AuditLogModel

    Private Sub AuditLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If model Is Nothing Then model = New AuditLogModel With {.Context = New DBContext}

        BindControls()

        AddHandler model.PropertyChanged,
           Sub(s As Object, ev As PropertyChangedEventArgs)
               If ev.PropertyName = "Period" AndAlso model.Period <> DatePeriod.SpecificDates Then model.Show()
               If ev.PropertyName = "Period" Then datePanel.Visible = (model.Period = DatePeriod.SpecificDates)
           End Sub

        model.Show()
    End Sub

    Private Sub BindControls()
        'usually all done automatically by ui framework
        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.ThisMonth, "This Month"))
        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.LastMonth, "Last Month"))
        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.SpecificDates, "Specific Dates"))
        periodEditor.DataBindings.Add(New Binding("Value", model, "Period", True, DataSourceUpdateMode.OnPropertyChanged))
        dateFromEditor.DataBindings.Add(New Binding("Value", model, "DateFrom", True, DataSourceUpdateMode.OnPropertyChanged))
        dateToEditor.DataBindings.Add(New Binding("Value", model, "DateTo", True, DataSourceUpdateMode.OnPropertyChanged))
        auditLogsGrid.DataSource = model.AuditLogs
        AddHandler showButton.Click,
            Sub()
                model.Show()
            End Sub
        AddHandler model.PropertyChanged,
            Sub(s As Object, ev As PropertyChangedEventArgs)
                If ev.PropertyName = "AuditLogs" Then auditLogsGrid.DataSource = model.AuditLogs
            End Sub
    End Sub
End Class