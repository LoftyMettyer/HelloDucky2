Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Data.SqlClient
Imports HR.Intranet.Server
Imports DMI.NET.Classes

Namespace Models

   Public Class OrgReportChartNode

      Public Property EmployeeID() As Integer
      Public Property EmployeeStaffNo() As String
      Public Property LineManagerStaffNo() As String
      Public Property HierarchyLevel() As Integer
      Public Property ReportColumnItemList As New List(Of ReportColumnItem)
      Public Property NodeTypeClass() As String
      Public Property PostTitle() As String
      Public Property PostWiseNodeList As New List(Of OrgReportChartNode)

   End Class

End Namespace