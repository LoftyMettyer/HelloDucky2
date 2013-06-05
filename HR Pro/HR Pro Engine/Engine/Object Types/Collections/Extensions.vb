Imports System.Runtime.CompilerServices

Public Module Extensions

  'TODO: MUST
  <Extension()>
  Public Sub AddIfNew(Of T)(ByVal list As IList(Of T), ByVal item As T)

  End Sub

  <Extension()>
  Public Function GetById(Of T)(ByVal items As IList(Of T), ByVal id As HCMGuid) As T
    Return Nothing
  End Function

End Module
