Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices

Namespace Helpers
	Public Module HtmlExtensions

		<Extension> _
		Public Function DisabledIf(html As HtmlHelper, condition As Boolean) As HtmlString
			Return New HtmlString(If(condition, "disabled=""disabled""", ""))
		End Function

	End Module
End Namespace