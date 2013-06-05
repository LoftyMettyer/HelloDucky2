Public Class HCMProgressBar
  Public Event Update1(ByVal Value As Long)
  Public Event Update2(ByVal Value As Long)
  Private miTotalSteps1 As Integer
  Private miTotalSteps2 As Integer
  Private miCurrentSteps1 As Integer = 0
  Private miCurrentSteps2 As Integer = 0

  Public Property TotalSteps1 As Integer
    Get
      Return miTotalSteps1
    End Get
    Set(ByVal value As Integer)
      miTotalSteps1 = value
      miCurrentSteps1 = 0
    End Set
  End Property

  Public Property TotalSteps2 As Integer
    Get
      Return miTotalSteps2
    End Get
    Set(ByVal value As Integer)
      miTotalSteps2 = value
      miCurrentSteps2 = 0
    End Set
  End Property

  Public Sub NextStep1()

    Dim iPercentage As Long

    If miTotalSteps1 = 0 Then
      RaiseEvent Update1(100)
    Else
      iPercentage = (miCurrentSteps1 / miTotalSteps1) * 100
      RaiseEvent Update1(iPercentage)

    End If

    ' Advance the step counter
    miCurrentSteps1 = miCurrentSteps1 + 1

  End Sub

  Public Sub NextStep2()

    Dim iPercentage As Long

    If miTotalSteps2 = 0 Then
      RaiseEvent Update2(100)
    Else
      iPercentage = (miCurrentSteps2 / miTotalSteps2) * 100
      RaiseEvent Update2(iPercentage)

    End If

    ' Advance the step counter
    miCurrentSteps2 = miCurrentSteps2 + 1

  End Sub


End Class
