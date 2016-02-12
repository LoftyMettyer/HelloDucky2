Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports HR.Intranet.Server.Structures

Namespace Extensions

	<HideModuleName()> _
Friend Module TalentChartExtensions

    <Extension()>
    Public Function TalentChart(Of T As Competency)(items As ICollection(Of T)) As String

      Dim output As String = ""

      For Each competency In items.Where(Function(m) m.Include or m.actual >= m.Minimum)
        output &= String.Format("{0} - Minimum : {1}, Preferred : {2}, Actual : {3}", _
                                competency.Name, competency.Preferred, competency.Minimum, competency.Actual) & vbNewLine
      Next

      If output.Length > 0 Then
        return output.Substring(0, output.Length - 1)
      Else 
        Return ""
      End If

		End Function

    <Extension()>
    Public Function TalentChartJSON(Of T As Competency)(items As ICollection(Of T)) As String

      Dim output As String = ""

      For Each competency In items.Where(Function(m) m.Include Or m.Actual >= m.Minimum)
        output &= String.Format("{{""Competency"":""{0}"", ""MinScore"":{1}, ""PrefScore"":{2}, ""ActualScore"":{3}, ""MaxScore"":{4}}}," _
                  ,competency.Name, competency.Minimum, competency.Preferred, competency.Actual, competency.Maximum)

      Next

      If output.Length > 0 Then
        return "[" & output.Substring(0, output.Length - 1) & "]"
      Else 
        Return ""
      End If


		End Function



  End Module


End Namespace
