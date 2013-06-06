﻿Imports System.Runtime.CompilerServices
Imports SystemFramework.Things

Public Module Extensions

  <Extension()>
  Public Sub AddIfNew(Of T As Base)(ByVal items As ICollection(Of T), ByVal item As T)

    If Not items.Any(Function(i) i.Id = item.Id AndAlso item.GetType = i.GetType) Then
      items.Add(item)
    End If

  End Sub

  <Extension()>
  Public Function GetById(Of T As Base)(ByVal items As ICollection(Of T), ByVal id As Integer) As T

    Return items.FirstOrDefault(Function(item) item.Id = id)

  End Function

  <Extension()>
  Public Sub Merge(Of T As Base)(ByVal first As ICollection(Of T), ByVal second As ICollection(Of T))

    For Each item As T In second
      If Not first.Contains(item) Then
        first.Add(item)
      End If
    Next

  End Sub

End Module
