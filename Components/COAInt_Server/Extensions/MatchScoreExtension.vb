Imports System.Collections.Generic
Imports HR.Intranet.Server.Structures
Imports System.Linq
Imports System.Runtime.CompilerServices

Namespace Extensions

	<HideModuleName()> _
Friend Module MatchScoreExtension

    <Extension()>
    Public Function MatchScore(Of T As Competency)(items As ICollection(Of T)) As Double

      Dim score As Double = 0

      For Each competency In items
        If competency.Actual >= competency.Preferred Or competency.Minimum <= 0 Then
          score += 1
        Else 
          score += competency.Actual / competency.Preferred
        End If
      Next

      if items.Any() Then
        score = score / items.Count() * 100
      End If

      return score
		End Function

    <Extension()>
    Public Function MatchCount(Of T As Competency)(items As ICollection(Of T)) As Integer
      Return items.Where(Function(i) i.Actual >= i.Minimum Or i.Include).Count()
    End Function

  End Module


End Namespace
