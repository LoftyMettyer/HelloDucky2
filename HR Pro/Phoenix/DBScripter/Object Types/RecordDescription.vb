Namespace Things
  <Serializable()> _
  Public Class RecordDescription
    Inherits Things.Expression

    '    Public Code As String

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.RecordDescription
      End Get
    End Property

    'Public Sub Generate()
    '  Dim objColumn As Things.Column
    '  For Each objColumn In Me.Parent.Objects(Things.Type.Column)
    '  Next
    'End Sub


  End Class
End Namespace
