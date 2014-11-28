Option Strict On
Option Explicit On

Imports HR.Intranet.Server.BaseClasses

Public Class StandardReport
	Inherits BaseForDMI

	Public Function IsCalculationValid(id As Integer) As String
		Dim isValid As String
		isValid = IsCalcValid(id)
		If isValid <> vbNullString Then
			Logs.AddDetailEntry(isValid & "It has been removed from the definition.")
		End If
		Return isValid
	End Function

	Public Function IsFilterNameValid(id As Integer) As String
		Dim isValid As String
		isValid = IsFilterValid(id)
		If isValid <> vbNullString Then
			Logs.AddDetailEntry(isValid & "It has been removed from the definition.")
		End If
		Return isValid
	End Function

	Public Function GetPicklistFilterName(pstrType As String, plngId As Integer) As String

		Dim strRecSelStatus As String
		Dim strName As String

		strName = String.Empty

		Select Case pstrType
			Case "A"

			Case "F"
				strRecSelStatus = IsFilterValid(plngId)
				If (strRecSelStatus <> vbNullString) Then
					strName = "None"
				Else
					strName = General.GetFilterName(plngId)
				End If
			Case "P"
				strRecSelStatus = IsPicklistValid(plngId)
				If (strRecSelStatus <> vbNullString) Then
					strName = "None"
				Else
					strName = General.GetPicklistName(plngId)
				End If
		End Select

		Return Replace(strName, """", "")

	End Function

	Public Function GetFilterName(filterId As Integer) As String
		Return General.GetFilterName(filterId)
	End Function

End Class