Namespace Things

  Public Class WorkflowElementColumn
    Inherits Things.Base

    'Public Overrides Function Commit() As Boolean
    'End Function

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.WorkflowElementColumn
      End Get
    End Property

    Public ColumnID As HCMGuid
    Public ValueType As Integer
    Public Value As String
    Public WFFormIdentifier As String
    Public WFValueIdentifier As String
    Public DBColumnID As HCMGuid
    Public DBRecord As Integer
    Public CalcID As HCMGuid

  End Class
End Namespace
