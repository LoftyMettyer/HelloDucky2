Public Interface IErrors
  Sub OutputToFile(ByVal fileName As String)
  Sub Show()
  ReadOnly Property ErrorCount As Integer
  ReadOnly Property IsCatastrophic As Boolean
End Interface
