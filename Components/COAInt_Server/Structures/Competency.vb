Option Strict On
Option Explicit On

Namespace Structures
  Public Class Competency
    Public Property Name As String
    Public Property Actual As Double
    Public Property Preferred As Double
    Public Property Minimum As Double
    Public Property Maximum As Double

    Public Function TalentGridJson As String

      Return String.Format("{{""Competency"":""{0}"", ""MinScore"":{1}, ""PrefScore"":{2}, ""ActualScore"":{3}, ""MaxScore"":{4}}}" _
        ,Name, Minimum, Preferred, Actual, Maximum)

    End Function
  End Class
End Namespace
