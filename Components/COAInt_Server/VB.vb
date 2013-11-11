'Imports System.Globalization

Public Class VB6

  Public Shared Function CopyArray(ByVal sourceArray) As Object
    '  Return Microsoft.VisualBasic.Compatibility.VB6.CopyArray(sourceArray)
    Return sourceArray.Clone()

  End Function

	Public Shared Function Format(value As Object, Optional style As String = "", Optional firstDayOfWeek As FirstDayOfWeek = FirstDayOfWeek.Sunday) As String
		'    Return Microsoft.VisualBasic.Compatibility.VB6.Format([value], style, [firstDayOfWeek])
		Dim theDate As Date

		If value Is Nothing Or IsDBNull(value) Then
			Return vbNullString
		End If

		Try
			If IsDate(value) Then
				theDate = Convert.ToDateTime(value)
				Return theDate.ToString(style)
			ElseIf style = "ddd" And value >= 1 And value <= 7 Then	'Day of week
				Return WeekdayName(value, False, firstDayOfWeek).Substring(0, 1)
			End If

			Return value.ToString()
		Catch ex As Exception
			Return vbNullString
		End Try
	End Function
End Class
