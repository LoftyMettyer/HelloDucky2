Option Strict On
Option Explicit On

Imports DMI.NET.Enums

Namespace Code

	<HideModuleName>
	Public Module Conversions

		Public Function ActionToUtilityAction(action As String) As UtilityActionType

			Select Case action.ToUpper
				Case "NEW"
					Return UtilityActionType.New
				Case "EDIT"
					Return UtilityActionType.Edit
				Case "COPY"
					Return UtilityActionType.Copy
				Case Else
					Return UtilityActionType.View

			End Select

		End Function

	End Module
End Namespace
