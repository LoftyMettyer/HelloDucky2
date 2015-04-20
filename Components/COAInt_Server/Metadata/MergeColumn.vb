Option Strict On
Option Explicit On

Namespace Metadata

	Public Class MergeColumn
		Inherits Column

		Public IsExpression As Boolean

		Public ReadOnly Property MergeName() As String
			Get

				Dim returnName As String

				If IsExpression Then
					returnName = String.Format("{0}{1}", TableName, Name.Replace(" ", "_"))
				Else
					returnName = String.Format("{0}_{1}", TableName, Name.Replace(" ", "_"))
				End If

				Return If(returnName.Length <= 40, returnName, returnName.Substring(0, 40))

			End Get

		End Property

		Public SortOrder As String

	End Class

End Namespace
