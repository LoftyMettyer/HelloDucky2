Public Class AuditLogForm

    Public ConString As String
    Private db As DBContext
    Private loading As Boolean = True

    Private Sub AuditLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        db = New DBContext(ConString)

        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.ThisMonth, "This Month"))
        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.LastMonth, "Last Month"))
        periodEditor.Items.Add(New Infragistics.Win.ValueListItem(DatePeriod.SpecificDates, "Specific Dates"))
        userEditor.Items.Add(New Infragistics.Win.ValueListItem(Nothing))
        userEditor.Items.AddRange(GetUsers().Select(Function(u) New Infragistics.Win.ValueListItem(u)).ToArray)

        periodEditor.SelectedIndex = 0
        loading = False
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

        If Not String.IsNullOrWhiteSpace(userEditor.Text) Then
            query = query.Where(Function(a) a.UserName = userEditor.Text)
        End If

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

    Private Function GetUsers() As IList(Of String)
        Return db.ASRSysAuditTrails.Select(Function(a) a.UserName).Distinct.ToList()
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



End Class