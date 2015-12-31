Option Strict On
Option Explicit On

Namespace Structures
  Public Class Competency
    Public Property Name As String
    Public Property Actual As Double
    Public Property Preferred As Double
    Public Property Minimum As Double

    Public Function TalentGridJson As String

      Return String.Format("{{""Competency"":""{0}"", ""MinScore"":{1}, ""PrefScore"":{2}, ""ActualScore"":{3}}}" _
        ,Name, Minimum, Preferred, Actual)

    End Function
  End Class
End Namespace
