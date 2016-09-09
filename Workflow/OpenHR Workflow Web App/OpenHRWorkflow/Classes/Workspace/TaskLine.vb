Namespace Classes.Workspace
   Public Class TaskLine
      Public header As String
      Public details As List(Of KeyValuePair)
      Public actions As List(Of TaskAction)

      Public Sub New()
         details = New List(Of KeyValuePair)
         actions = New List(Of TaskAction)
      End Sub
   End Class
End Namespace