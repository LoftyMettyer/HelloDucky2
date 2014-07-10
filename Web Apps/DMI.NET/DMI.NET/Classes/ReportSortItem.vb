Option Strict On
Option Explicit On

Namespace Classes
	Public Class ReportSortItem
		Implements IJsonSerialize

		Public Property TableID As Integer
		Public Property ID As Integer Implements IJsonSerialize.ID
		Public Property Name As String
		Public Property Order As String
		Public Property Sequence As Integer
		Public Property BreakOnChange As Boolean
		Public Property PageOnChange As Boolean
		Public Property ValueOnChange As Boolean
		Public Property SuppressRepeated As Boolean
	End Class
End Namespace