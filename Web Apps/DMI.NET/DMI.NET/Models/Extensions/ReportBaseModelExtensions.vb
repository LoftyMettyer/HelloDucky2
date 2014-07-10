Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports DMI.NET.Models

Public Module ReportBaseModelExtensions

	<Extension()>
	Public Sub Remove(Of T As ReportBaseModel)(items As ICollection(Of T), id As Integer)

		For Each objItem In items
			If objItem.ID = id Then
				items.Remove(objItem)
				Return
			End If
		Next

	End Sub

End Module