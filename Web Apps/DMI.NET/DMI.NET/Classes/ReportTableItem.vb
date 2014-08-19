﻿Option Strict On
Option Explicit On

Namespace Classes
	Public Class ReportTableItem
		Implements IJsonSerialize

		Public Property [id] As Integer Implements IJsonSerialize.ID
		Public Property Name As String
		Public Property Relation As ReportRelationType

	End Class
End Namespace