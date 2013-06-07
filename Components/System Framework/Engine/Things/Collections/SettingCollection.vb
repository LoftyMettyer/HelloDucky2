Namespace Collections

  Public Class Settings
    Inherits Collection(Of Setting)

    Public Function Setting(ByVal [module] As String, ByVal parameter As String) As Setting

      Dim getItem = Items.FirstOrDefault(Function(s) s.Module.ToLower = [module].ToLower AndAlso s.Parameter.ToLower = parameter.ToLower)

      If getItem IsNot Nothing Then
        Return getItem
      Else
        Return New Setting
      End If

    End Function

  End Class

End Namespace
