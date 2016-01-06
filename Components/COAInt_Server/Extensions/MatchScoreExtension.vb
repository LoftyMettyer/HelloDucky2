﻿Imports System.Collections.Generic
Imports HR.Intranet.Server.Structures
Imports System.Linq
Imports System.Runtime.CompilerServices

Namespace Extensions

	<HideModuleName()> _
Friend Module MatchScoreExtension

    <Extension()>
    Public Function MatchScore(Of T As Competency)(items As ICollection(Of T)) As Double

      Dim score As Double = 0

      For Each competency In items.Where(Function(m) m.Actual >= m.Minimum)

        If competency.Actual >= Math.Max(competency.Minimum, competency.Preferred) Then
          score += 1
        ElseIf competency.Actual < Math.Min(competency.Minimum, competency.Preferred)
          score += 0
        Else 
          score += (competency.Actual - competency.Minimum) / (competency.Preferred - competency.Minimum)
        End If

      Next

      if items.Any() Then
        score = score / items.Count() * 100
      End If

      return score
		End Function

    <Extension()>
    Public Function MatchCount(Of T As Competency)(items As ICollection(Of T)) As Integer
      Return items.Where(Function(i) i.Actual >= i.Minimum).Count()
    End Function

    <Extension()>
    Public Function AllMatched(Of T As Competency)(items As ICollection(Of T)) As Boolean
      Return Not items.Any(Function(i) i.Actual < i.Minimum)
    End Function

  End Module


End Namespace