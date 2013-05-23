'Imports System.Globalization

Public Class VB6

  Public Shared Function CopyArray(ByVal sourceArray) As Object
    '  Return Microsoft.VisualBasic.Compatibility.VB6.CopyArray(sourceArray)
    Return sourceArray.Clone()

  End Function

  Public Shared Function Format([value] As Object, Optional style As String = "", Optional [firstDayOfWeek] As FirstDayOfWeek = FirstDayOfWeek.Sunday) As String
    '    Return Microsoft.VisualBasic.Compatibility.VB6.Format([value], style, [firstDayOfWeek])
    Dim theDate As Date

    If IsDate(value) Then
      theDate = Convert.ToDateTime(value)
      Return theDate.ToString(style)
    End If

    Return [value].ToString()

  End Function


End Class
