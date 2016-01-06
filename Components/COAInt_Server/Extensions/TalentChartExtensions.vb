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

      For Each competency In items.Where(Function(m) m.Actual >= m.Minimum)
        output &= String.Format("{0} - Minimum : {1}, Preferred : {2}, Actual : {3}", _
                                competency.Name, competency.Preferred, competency.Minimum, competency.Actual) & vbNewLine
      Next

      return output.Substring(0, output.Length - 1)

		End Function


  End Module


End Namespace
