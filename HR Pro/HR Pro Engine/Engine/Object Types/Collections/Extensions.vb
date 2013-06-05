Imports System.Runtime.CompilerServices
Imports SystemFramework.Things

Public Module Extensions

  <Extension()>
  Public Sub AddIfNew(Of T As Base)(ByVal list As IList(Of T), ByVal item As T)

    If list.GetById(item.id) Is Nothing Then
      list.Add(item)
    End If

  End Sub

  <Extension()>
  Public Function GetById(Of T As Base)(ByVal items As IList(Of T), ByVal id As Integer) As T

    Return items.SingleOrDefault(Function(item) item.ID = id)

  End Function

End Module
