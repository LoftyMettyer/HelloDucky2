Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.Metadata

Namespace ModuleSpecifics

	Friend Class modBankHolidaySpecifics

		Public gfBankHolidaysEnabled As Boolean

		' Bank Holiday Region Table
		Public glngBHolRegionTableID As Integer
		Public gsBHolRegionTableName As String

		' Bank Holiday Region Column
		Public glngBHolRegionID As Integer
		Public gsBHolRegionColumnName As String

		' Bank Holiday Instances Table
		Public glngBHolTableID As Integer
		Public gsBHolTableName As String
		Public gsBHolTableRealSource As String

		' Bank Holiday Instances Date Column
		Public glngBHolDateID As Integer
		Public gsBHolDateColumnName As String

		' Bank Holiday Instances Description Column
		Public glngBHolDescriptionID As Integer
		Public gsBHolDescriptionColumnName As String

		Public Sub ReadBankHolidayParameters()

			Dim objTable As TablePrivilege

			On Error GoTo ReadParametersERROR

			gfBankHolidaysEnabled = True

			' Bank Holiday Region Table and Column
			glngBHolRegionTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE))
			If glngBHolRegionTableID > 0 Then
				gsBHolRegionTableName = datGeneral.GetTableName(glngBHolRegionTableID)
			Else
				gsBHolRegionTableName = ""
				gfBankHolidaysEnabled = False
			End If

			glngBHolRegionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGION))
			If glngBHolRegionID > 0 Then
				gsBHolRegionColumnName = datGeneral.GetColumnName(glngBHolRegionID)
			Else
				gsBHolRegionColumnName = ""
				gfBankHolidaysEnabled = False
			End If

			' Bank Holiday Instance Table and Columns

			glngBHolTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE))
			If glngBHolTableID > 0 Then
				gsBHolTableName = datGeneral.GetTableName(glngBHolTableID)

				' Get the realsource into a variable too
				objTable = gcoTablePrivileges.FindTableID(glngBHolTableID)
				gsBHolTableRealSource = objTable.RealSource
				'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objTable = Nothing

			Else
				gsBHolTableName = ""
				gfBankHolidaysEnabled = False
			End If

			glngBHolDateID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDATE))
			If glngBHolDateID > 0 Then
				gsBHolDateColumnName = datGeneral.GetColumnName(glngBHolDateID)
			Else
				gsBHolDateColumnName = ""
				gfBankHolidaysEnabled = False
			End If

			glngBHolDescriptionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDESCRIPTION))
			If glngBHolDescriptionID > 0 Then
				gsBHolDescriptionColumnName = datGeneral.GetColumnName(glngBHolDescriptionID)
			Else
				gsBHolDescriptionColumnName = ""
				gfBankHolidaysEnabled = False
			End If

			'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objTable = Nothing

			Exit Sub

ReadParametersERROR:

			'NO MSGBOX ON THE SERVER ! - MsgBox "Error reading the Bank Holiday parameters." & vbNewLine & _
			'Err.Description, vbExclamation + vbOKOnly, App.Title
			gfBankHolidaysEnabled = False
			'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objTable = Nothing

		End Sub


		Public Function ValidateBankHolidayParameters() As Boolean

		End Function
	End Class
End Namespace