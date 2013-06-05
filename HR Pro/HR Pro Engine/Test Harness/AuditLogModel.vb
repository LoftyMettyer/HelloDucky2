Imports System.ComponentModel

'Public Enum DatePeriod
'    ThisMonth = 1
'    LastMonth
'End Enum

'Public Class AuditLogModel

'    Private db As DBContext

'    Public Property Period As DatePeriod
'    Public Property DateFrom As Date?
'    Public Property DateTo As Date?
'    Public Property CanShow As Boolean = True
'    Public Property Periods As IList(Of Lookup)
'    Public Property AuditLogs As IList

'    Public Event Bind()

'    Public Sub New(ByVal db As DBContext)
'        Me.db = db
'        Period = DatePeriod.ThisMonth
'        Periods = New List(Of Lookup) From {New Lookup(0, ""), New Lookup(DatePeriod.ThisMonth, "This Month"), New Lookup(DatePeriod.LastMonth, "Last Month")}
'        AuditLogs = GetAuditLogs()
'    End Sub

'    Public Sub Show()
'        AuditLogs = GetAuditLogs()
'        RaiseEvent Bind()
'    End Sub

'    Private Function GetAuditLogs() As IList

'        Dim data = From a In db.ASRSysAuditTrails
'                   Select a

'        If DateFrom.HasValue Then
'            data = data.Where(Function(a) a.DateTimeStamp >= DateFrom.Value)
'        End If

'        If DateTo.HasValue Then
'            data = data.Where(Function(a) a.DateTimeStamp < DateTo.Value.AddDays(1.0))
'        End If

'        Return data.ToList
'    End Function

'End Class

'Public Class Lookup
'    Public Sub New(ByVal id As Integer, ByVal name As String)
'        Me.Id = id
'        Me.Name = name
'    End Sub
'    Public Id As Integer
'    Public Name As String
'End Class