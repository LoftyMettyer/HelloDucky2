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

      For Each competency In items.Where(Function(m) m.Actual >= m.Minimum)
        score += Math.Min(competency.Actual, competency.Preferred)       
      Next

      if items.Any() Then
        score = score / items.Count()
      End If

      return score
'			Return items.FirstOrDefault(Function(baseItem) (baseItem.TableName = name.ToUpper() And baseItem.IsTable = True) Or (baseItem.ViewName = name And baseItem.IsTable = False))
		End Function

  End Module


End Namespace
