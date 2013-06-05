Imports System.ComponentModel

Public Enum DatePeriod
    ThisMonth = 1
    LastMonth
    SpecificDates
End Enum

Public Class AuditLogModel
    Inherits ModelBase

    Friend Context As DBContext

#Region "    Public Property Period As DatePeriod"

    'cant use auto property for period cos havent implemented auto inotifypropertychanged code
    Private _period As DatePeriod = DatePeriod.ThisMonth
    Public Property Period() As DatePeriod
        Get
            Return _period
        End Get
        Set(ByVal value As DatePeriod)
            _period = value
            OnPropertyChanged("Period")
        End Set
    End Property
#End Region
    Public Property DateFrom As Date?
    Public Property DateTo As Date?
    Public Property AuditLogs As IList

    Public Sub Show()

        Dim monthStart As Date = New Date(Today.Year, Today.Month, 1)

        Dim query = From a In Context.ASRSysAuditTrails

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

        AuditLogs = data.ToList
        OnPropertyChanged("AuditLogs")
    End Sub

End Class

#Region "Framework"

Public Class ModelBase
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(ByVal propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
End Class

#End Region