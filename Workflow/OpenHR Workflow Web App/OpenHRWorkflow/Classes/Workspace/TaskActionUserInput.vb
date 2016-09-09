Namespace Classes.Workspace
   Public Class TaskActionUserInput
      Public name As String
      Public label As String
      Public dataType As String
      Public mandatory As String 'YES or NO
      Public userValue As String 'blank when received from e5, cp etc but updated in TaskManager from User Input
   End Class
End Namespace