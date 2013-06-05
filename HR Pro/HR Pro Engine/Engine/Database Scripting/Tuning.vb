Option Explicit On

Imports System.Runtime.InteropServices
Imports System.Text

Namespace ScriptDB

  Public Class Tuning

    Private mlngFilterCount As Long = 0
    Private mlngOrderCount As Long = 0
    Private mlngSelectCount As Long = 0
    Private mlngSelectUDFCount As Long = 0
    Private mlngSelectAsParameter As Long = 0

    Public Sub IncrementFilter()
      mlngFilterCount = mlngFilterCount + 1
    End Sub

    Public Sub IncrementOrder()
      mlngOrderCount = mlngOrderCount + 1
    End Sub

    Public Sub IncrementSelect()
      mlngSelectCount = mlngSelectCount + 1
    End Sub

    Public Sub IncrementSelectAsCalculation()
      mlngSelectUDFCount = mlngSelectUDFCount + 1
    End Sub

    Public Sub IncrementSelectAsParameter()
      mlngSelectAsParameter = mlngSelectAsParameter + 1
    End Sub


    'Public Function TuningReport() As String
    '  Dim sReport As String = vbNullString
    '  Return sReport
    'End Function

  End Class

End Namespace
