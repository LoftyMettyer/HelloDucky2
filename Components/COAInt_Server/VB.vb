'Imports System.Globalization

Namespace Scripting

	Public Class File
		Public Size As Long
		Public DateLastModified As DateTime
		Public Attributes As String

	End Class

	Public Class FileSystemObject
		Public Function FolderExists(ByVal Folder As String) As Boolean
			Return True
		End Function
		Public Function FileExists(ByVal Folder As String) As Boolean
			Return True
		End Function
		Public Function GetFile(ByVal Filename As String) As File
			Return New File
		End Function
		Public Sub CopyFile(ByVal From As String, ByVal [To] As String, ByVal Overwrite As Boolean)
		End Sub
		Public Function GetFolder() As Object
			Return New Object()
		End Function

	End Class


End Namespace

Public Class VB6

	Public Shared TwipsPerPixelX As Integer
	Public Shared TwipsPerPixelY As Integer

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
			ElseIf style = "ddd" AndAlso value >= 1 AndAlso value <= 7 Then	'Day of week
				Return WeekdayName(value, False, firstDayOfWeek).Substring(0, 1)
			End If

			Return value.ToString()
		Catch ex As Exception
			Return vbNullString
		End Try
	End Function
End Class
