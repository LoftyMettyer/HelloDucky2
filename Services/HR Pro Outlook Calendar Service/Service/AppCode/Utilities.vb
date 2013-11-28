Module Utilities

  '****************************************************************
  ' NullSafeString
  '****************************************************************
  Public Function NullSafeString(ByVal arg As Object, _
  Optional ByVal returnIfEmpty As String = "") As String

    Dim returnValue As String

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
                     OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CStr(arg).Trim
      Catch
        returnValue = returnIfEmpty
      End Try

    End If

    Return returnValue

  End Function

  '****************************************************************
  ' NullSafeInteger
  '****************************************************************
  Public Function NullSafeInteger(ByVal arg As Object, _
    Optional ByVal returnIfEmpty As Integer = 0) As Integer

    Dim returnValue As Integer

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
                     OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CInt(arg)
      Catch
        returnValue = returnIfEmpty
      End Try
    End If

    Return returnValue

  End Function

  '****************************************************************
  '   NullSafeDouble
  '****************************************************************
  Public Function NullSafeDouble(ByVal arg As Object, _
    Optional ByVal returnIfEmpty As Double = 0) As Double

    Dim returnValue As Double

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
                     OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CDbl(arg)
      Catch
        returnValue = returnIfEmpty
      End Try
    End If

    Return returnValue

  End Function

  '****************************************************************
  '   NullSafeSingle
  '****************************************************************
  Public Function NullSafeSingle(ByVal arg As Object, _
    Optional ByVal returnIfEmpty As Single = 0) As Single

    Dim returnValue As Single

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
                     OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CSng(arg)
      Catch
        returnValue = returnIfEmpty
      End Try
    End If

    Return returnValue

  End Function

  '****************************************************************
  ' NullSafeBoolean
  '****************************************************************
  Public Function NullSafeBoolean(ByVal arg As Object) As Boolean

    Dim returnValue As Boolean

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
                     OrElse (arg Is String.Empty) Then
      returnValue = False
    Else
      Try
        returnValue = CBool(arg)
      Catch
        returnValue = False
      End Try
    End If

    Return returnValue

  End Function
End Module
