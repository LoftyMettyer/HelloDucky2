Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices

Namespace Extensions

	<HideModuleName()> _
 Public Module DateExtensions

		<Extension> _
		Public Function ToEndOfMonth(source As Date) As Date
			Dim days = DateTime.DaysInMonth(source.Year, source.Month)
			Return New DateTime(source.Year, source.Month, days)
		End Function

		<Extension> _
		Public Function ToStartOfMonth(source As Date) As Date
			Return New DateTime(source.Year, source.Month, 1)
		End Function

	End Module
End Namespace