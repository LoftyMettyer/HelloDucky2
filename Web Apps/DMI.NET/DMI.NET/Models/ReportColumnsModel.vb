Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Metadata

Namespace Models

	Public Class ReportColumnsModel

		'Public Property Available As New List(Of ReportColumnItem)
		Public Property BaseTableID As Integer
		Public Property Selected As New Collection(Of ReportColumnItem)

	End Class

End Namespace