Option Strict On
Option Explicit On

Imports DMI.NET.Enums
Imports DMI.NET.AttributeExtensions
Imports HR.Intranet.Server.Enums

Namespace Classes

	Public Class ReportRelatedTable
		Implements IJsonSerialize

		Public Property ID As Integer Implements IJsonSerialize.ID
		Public Property Name As String
		Public Property SelectionType As RecordSelectionType

		<NonZeroIf("SelectionType", RecordSelectionType.Filter, ErrorMessage:="No filter on selected table.")> _
		Public Property FilterID As Integer

		<NonZeroIf("SelectionType", RecordSelectionType.Picklist, ErrorMessage:="No picklist on selected table.")> _
		Public Property PicklistID As Integer

		Public Property FilterName As String
		Public Property PicklistName As String

		Public Property RelationType As RelationType

		Public ReadOnly Property Visibility As String
			Get
				Return (If(ID < 1, "disabled", ""))
			End Get
		End Property

	End Class
End Namespace