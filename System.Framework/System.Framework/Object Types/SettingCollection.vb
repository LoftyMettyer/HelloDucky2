Namespace Things

  Public Class SettingCollection
    Inherits ObjectModel.Collection(Of Setting)

    Public Function Setting(ByVal [module] As String, ByVal parameter As String) As Setting

      Dim item = Items.FirstOrDefault(Function(s) s.Module.ToLower = [module].ToLower AndAlso s.Parameter.ToLower = parameter.ToLower)

      If item IsNot Nothing Then
        Return item
      Else
        Return New Setting
      End If

    End Function

  End Class

End Namespace
