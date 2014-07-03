Option Strict On
Option Explicit On

Imports DMI.NET.Enums

Namespace Classes

  Public Class ReportRelatedTable

    Public Property ID As Integer
    Public Property Name As String
    Public Property SelectionType As RecordSelectionType
    Public Property FilterID As Integer
		Public Property PicklistID As Integer
		Public Property FilterName As String
		Public Property PicklistName As String

  End Class
End Namespace