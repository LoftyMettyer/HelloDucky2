
Namespace Classes.Workspace
   Public Class TaskAction
      Public name As String
      Public label As String
      Public url As String
      Public urlTargetName As String
      Public style As String
      Public iconName As String
      Public showOnListScreen As String
      Public userInputs As List(Of TaskActionUserInput)

      Public Sub New
         userInputs = New List(Of TaskActionUserInput)
      End Sub
   End Class
End Namespace