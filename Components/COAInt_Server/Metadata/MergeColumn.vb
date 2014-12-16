Option Strict On
Option Explicit On

Namespace Metadata

	Public Class MergeColumn
		Inherits Column

		Public IsExpression As Boolean

		Public ReadOnly Property MergeName() As String
			Get
				If IsExpression Then
					Return String.Format("{0}{1}", TableName, Name.Replace(" ", "_"))
				Else
					Return String.Format("{0}_{1}", TableName, Name.Replace(" ", "_"))
				End If

			End Get

		End Property

		Public SortOrder As String

	End Class

End Namespace
