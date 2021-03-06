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

      For Each competency In items.Where(function(c) c.Actual >= c.Minimum And c.Actual > 0)
        If competency.Actual >= competency.Preferred Or competency.Minimum <= 0 Then
          score += 1
        Else 
          Dim range = (competency.Preferred - competency.Minimum)
          If range > 0 Then
            score += Math.Min((competency.Actual - competency.Minimum + 1) / range, 1)
           End If
        End If
      Next

      if items.Any() Then
        score = score / items.Count() * 100
      End If

      return score
		End Function

    <Extension()>
    Public Function MatchCount(Of T As Competency)(items As ICollection(Of T)) As Integer
      Return items.Where(Function(i) (i.Actual >= i.Minimum And i.Actual > 0) Or i.Include).Count()
    End Function

  End Module


End Namespace
