Namespace Classes.Workspace
   Public Class Task
      Public id As String
      Public category As String
      Public subCategory As String
      Public created As DateTime
      Public createdBy As String
      Public actionBefore As DateTime
      Public priority As Integer
      Public status As String
      Public sourceName As String
      Public sourceId As String
      ' assigned by the DataFeed and needed for routing.
      Public historical As String

      ' explicit details of the task
      Public headers As List(Of String)
      Public details As List(Of KeyValuePair)
      Public lines As List(Of TaskLine)
      Public actions As List(Of TaskAction)
      Public images As List(Of TaskImage)

      Public Sub New()
         headers = New List(Of String)
         details = New List(Of KeyValuePair)
         lines = New List(Of TaskLine)
         actions = New List(Of TaskAction)
         images = New List(Of TaskImage)
      End Sub
   End Class
End Namespace