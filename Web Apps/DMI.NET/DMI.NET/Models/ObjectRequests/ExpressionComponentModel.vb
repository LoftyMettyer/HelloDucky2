Option Strict On
Option Explicit On

Namespace Models.ObjectRequests
	Public Class ExpressionComponentModel

		Public Property txtGotoOptionAction As OptionActionType
		Public Property txtGotoOptionTableID As Integer
		Public Property txtGotoOptionLinkRecordID As String

		<AllowHtml>
		Public Property txtGotoOptionExtension As String
		Public Property txtGotoOptionExprType As Integer
		Public Property txtGotoOptionExprID As Integer
		Public Property txtGotoOptionFunctionID As Integer
		Public Property txtGotoOptionParameterIndex As Integer

	End Class
End Namespace