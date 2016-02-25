Option Strict On
Option Explicit On

Namespace Code.Classes
  Public Class StepAuthorization

    Public InstanceId As Integer
    Public ElementId As Integer
    Public RequiresAuthorization As Boolean
    Public AuthorizedUsers As List(Of String)

    Public Function IsValidUserForStep (currentUser As String) As Boolean
      Return AuthorizedUsers.Contains(currentUser)
    End Function

    Public HasBeenAuthenticated As Boolean = False

  End Class
End Namespace