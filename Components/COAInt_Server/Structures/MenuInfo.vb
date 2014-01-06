Imports System.Collections.Generic

Namespace Structures
	Public Class MenuInfo
		Public TableID As Integer
		Public TableName As String
		Public TableScreenCount As Integer
		Public TableScreenID As Integer
		Public TableReadable As Boolean
		Public TableViewCount As Integer
		Public ViewID As Integer
		Public ViewName As String
		Public ViewScreenCount As Integer
		Public ViewScreenID As Integer
		Public TableScreenPictureID As Integer
		Public ViewScreenPictureID As Integer
		Public SubItems As ICollection(Of MenuInfo)
		Public ScreenName As String
	End Class
End Namespace
