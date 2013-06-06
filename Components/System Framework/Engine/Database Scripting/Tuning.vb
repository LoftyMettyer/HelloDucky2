Namespace ScriptDB

  <Serializable()>
  Public Class Tuning

    Public Rating As Long
    Public Usage As Long

    Private mlngFilterCount As Long
    Private mlngOrderCount As Long
    Private mlngSelectCount As Long
    Private mlngSelectUDFCount As Long
    Private mlngSelectAsParameter As Long

    Private _mlngFunctions As Long

    Public Sub IncrementFunction()
      _mlngFunctions += 1
    End Sub

    Public Sub IncrementFilter()
      mlngFilterCount += 1
    End Sub

    Public Sub IncrementOrder()
      mlngOrderCount += 1
    End Sub

    Public Sub IncrementSelectAsCalculation()
      mlngSelectUDFCount += 1
    End Sub

    Public Sub IncrementSelectAsParameter()
      mlngSelectAsParameter += 1
    End Sub

    Public ReadOnly Property ExpressionComplexity As String
      Get
        Dim sComplexity As String
        sComplexity = String.Format("{0} Functions. {1} Parent Columns. {2} Child Columns. {3} Trigger Columns." _
              , _mlngFunctions, 0, 0, 0)

        Return sComplexity

      End Get
    End Property

    'Public Function TuningReport() As String
    '  Dim sReport As String = vbNullString
    '  Return sReport
    'End Function

  End Class

End Namespace
