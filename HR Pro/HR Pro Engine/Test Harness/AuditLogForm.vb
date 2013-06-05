Public Class AuditLog

    'TODO: Columns?
    'TODO: Export
    'TODO: Path + Validation
    'TODO: Look
    'TODO: As Model + Binding

    Private db As DBContext
    Private loading As Boolean = True

    Private Sub AuditLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        periodEditor.SelectedIndex = 0
        loading = False

        db = New DBContext
        ShowAuditLogs()
    End Sub

    Private Function GetAuditLogs() As IList

        Dim firstDayOfThisMonth As Date = New Date(Today.Year, Today.Month, 1)

        Dim query = From a In db.ASRSysAuditTrails

        Select Case periodEditor.Value
            Case 1 'This Month
                query = query.Where(Function(a) a.DateTimeStamp >= firstDayOfThisMonth And a.DateTimeStamp < firstDayOfThisMonth.AddMonths(1))
            Case 2 'Last Month
                query = query.Where(Function(a) a.DateTimeStamp >= firstDayOfThisMonth.AddMonths(-1) And a.DateTimeStamp < firstDayOfThisMonth)
            Case 3 'Between Dates
                If IsDate(dateFromEditor.Value) Then
                    query = query.Where(Function(a) a.DateTimeStamp >= CDate(dateFromEditor.Value))
                End If

                If IsDate(dateToEditor.Value) Then
                    query = query.Where(Function(a) a.DateTimeStamp < CDate(dateToEditor.Value).AddDays(1.0))
                End If
        End Select

        Dim data = From a In query
                   Select New With {
                    .User = a.UserName,
                    .Date = a.DateTimeStamp,
                    .Table = a.Tablename,
                    .Column = a.Columnname,
                    .OldValue = a.OldValue,
                    .NewValue = a.NewValue,
                    .RecordDescription = a.RecordDesc,
                    .Id = a.id
                   }

        Return data.ToList
    End Function

    Private Sub ShowAuditLogs()
        grdAudit.DataSource = GetAuditLogs()
    End Sub

    Private Sub ShowControls()
        datePanel.Visible = (periodEditor.Value = 3)
    End Sub

    Private Sub showButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles showButton.Click
        If Not IsDate(dateFromEditor.Value) AndAlso Not IsDate(dateToEditor.Value) Then
            MsgBox("You must enter a search date.", MsgBoxStyle.Exclamation, "Audit Log")
            Exit Sub
        End If
        ShowAuditLogs()
    End Sub

    Private Sub periodEditor_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles periodEditor.ValueChanged
        ShowControls()
        If Not loading AndAlso (periodEditor.Value = 1 OrElse periodEditor.Value = 2) Then
            ShowAuditLogs()
        End If
    End Sub

    Private Sub grdAudit_InitializeLayout1(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAudit.InitializeLayout
        e.Layout.Override.FilterUIType = Infragistics.Win.UltraWinGrid.FilterUIType.HeaderIcons
    End Sub

    Private Sub butOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butOutput.Click
        UltraGridExcelExporter1.Export(grdAudit, txtFilePath.Text)
    End Sub

End Class