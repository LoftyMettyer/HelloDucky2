Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports HR.Intranet.Server.Enums

Namespace Extensions

	<HideModuleName()> _
	Public Module UtilityTypeExtension

		<Extension> _
		Public Function ToSecurityPrefix(Source As UtilityType) As String
			Dim sSecurityID As String = ""

			Select Case Source
				Case UtilityType.utlCrossTab
					sSecurityID = "CROSSTABS"

				Case UtilityType.utlCustomReport
					sSecurityID = "CUSTOMREPORTS"

				Case UtilityType.utlMailMerge
					sSecurityID = "MAILMERGE"

				Case UtilityType.utlPicklist
					sSecurityID = "PICKLISTS"

				Case UtilityType.utlFilter
					sSecurityID = "FILTERS"

				Case UtilityType.utlCalculation
					sSecurityID = "CALCULATIONS"

				Case UtilityType.utlCalendarReport
					sSecurityID = "CALENDARREPORTS"

				Case UtilityType.utlWorkflow
					sSecurityID = "WORKFLOW"

				Case UtilityType.utlNineBoxGrid
					sSecurityID = "NINEBOXGRID"

				Case UtilityType.NewUser
					sSecurityID = "NEWUSER"

			End Select

			Return sSecurityID

		End Function


	End Module

End Namespace
