﻿Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports DMI.NET.ViewModels.Reports

Namespace Code.Extensions

	<HideModuleName()>
	Public Module ModelStateDictionaryExtensions

		<Extension()>
		Public Function ToWebMessage(Of T As ModelStateDictionary)(item As T) As SaveWarningModel

			Dim objWarning As New SaveWarningModel

			For Each objError In item.Values.SelectMany(Function(v) v.Errors)
				objWarning.ErrorCode = ReportValidationStatus.InvalidOnClient
				objWarning.ErrorMessage += String.Format("{0}{1}", objError.ErrorMessage, "<BR/>")
			Next

			Return objWarning

		End Function

	End Module
End Namespace