Imports System.ComponentModel

Public Enum DatePeriod
    ThisMonth = 1
    LastMonth
    SpecificDates
End Enum

Public Class AuditLogModel

    Private db As DBContext
    Public Property Period As DatePeriod = DatePeriod.ThisMonth
    Public Property DateFrom As Date?
    Public Property DateTo As Date?
    Public Property User As String
    Public Property Users As IList(Of String)
    Public Property AuditLogs As IList
    Public Event AuditLogsChanged()

    Public Sub New(ByVal db As DBContext)
        Me.db = db
        Users = db.ASRSysAuditTrails.Select(Function(a) a.UserName).Distinct().ToList
    End Sub

    Public Sub Find()

        Dim query = From a In db.ASRSysAuditTrails

        Dim monthStart As Date = New Date(Today.Year, Today.Month, 1)

        Select Case Period
            Case DatePeriod.ThisMonth
                query = query.Where(Function(a) a.DateTimeStamp >= monthStart And a.DateTimeStamp < monthStart.AddMonths(1))
            Case DatePeriod.LastMonth
                query = query.Where(Function(a) a.DateTimeStamp >= monthStart.AddMonths(-1) And a.DateTimeStamp < monthStart)
            Case DatePeriod.SpecificDates
                If DateFrom.HasValue Then
                    query = query.Where(Function(a) a.DateTimeStamp >= DateFrom.Value)
                End If

                If DateTo.HasValue Then
                    query = query.Where(Function(a) a.DateTimeStamp < DateTo.Value.AddDays(1.0))
                End If
        End Select

        If Not String.IsNullOrWhiteSpace(User) Then
            query = query.Where(Function(a) a.UserName = User)
        End If

        Dim data = From a In query
                   Select New With {
                    .User = a.UserName,
                    .Date = a.DateTimeStamp,
                    .Table = a.Tablename,
                    .Column = a.Columnname,
                    .OldValue = a.OldValue,
                    .NewValue = a.NewValue,
                    .RecordDescription = a.RecordDesc
                   }

        AuditLogs = data.ToList
        RaiseEvent AuditLogsChanged()
    End Sub

End Class
