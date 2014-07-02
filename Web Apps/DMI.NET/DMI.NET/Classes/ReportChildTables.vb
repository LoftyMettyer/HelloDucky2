Option Strict On
Option Explicit On

Namespace Classes

  Public Class ReportChildTables
    Public Property TableID As Integer
    Public Property FilterID As Integer
    Public Property OrderID As Integer
    Public Property Records As Integer

    ' these are for display purposes (better way?)
    Public Property TableName As String
    Public Property FilterName As String
    Public Property OrderName As String

  End Class

End Namespace