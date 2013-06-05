Namespace Things

  Public Class WorkflowElementColumn
    Inherits Things.Base

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.WorkflowElementColumn
      End Get
    End Property

    Public Property ColumnID As Integer
    Public Property ValueType As Integer
    Public Property Value As String
    Public Property WFFormIdentifier As String
    Public Property WFValueIdentifier As String
    Public Property DBColumnID As Integer
    Public Property DBRecord As Integer
    Public Property CalcID As Integer

  End Class
End Namespace
