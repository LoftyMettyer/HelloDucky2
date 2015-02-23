Option Strict On
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Metadata

Namespace ModuleSpecifics

	Friend Class modBankHolidaySpecifics
		Inherits BaseModuleSpecific

		Public Sub New(value As SessionInfo)
			MyBase.New(value)
		End Sub

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

			Try

				gfBankHolidaysEnabled = True

				' Bank Holiday Region Table and Column
				glngBHolRegionTableID = CInt(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE))
				If glngBHolRegionTableID > 0 Then
					gsBHolRegionTableName = _tables.GetById(glngBHolRegionTableID).Name
				Else
					gsBHolRegionTableName = ""
					gfBankHolidaysEnabled = False
				End If

				glngBHolRegionID = CInt(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGION))
				If glngBHolRegionID > 0 Then
					gsBHolRegionColumnName = _columns.GetById(glngBHolRegionID).Name
				Else
					gsBHolRegionColumnName = ""
					gfBankHolidaysEnabled = False
				End If

				' Bank Holiday Instance Table and Columns

				glngBHolTableID = CInt(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE))
				If glngBHolTableID > 0 Then
					gsBHolTableName = _tables.GetById(glngBHolTableID).Name

					' Get the realsource into a variable too
					objTable = _tablePrivileges.FindTableID(glngBHolTableID)
					gsBHolTableRealSource = objTable.RealSource

				Else
					gsBHolTableName = ""
					gfBankHolidaysEnabled = False
				End If

				glngBHolDateID = CInt(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDATE))
				If glngBHolDateID > 0 Then
					gsBHolDateColumnName = _columns.GetById(glngBHolDateID).Name
				Else
					gsBHolDateColumnName = ""
					gfBankHolidaysEnabled = False
				End If

				glngBHolDescriptionID = CInt(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDESCRIPTION))
				If glngBHolDescriptionID > 0 Then
					gsBHolDescriptionColumnName = _columns.GetById(glngBHolDescriptionID).Name
				Else
					gsBHolDescriptionColumnName = ""
					gfBankHolidaysEnabled = False
				End If


			Catch ex As Exception
				gfBankHolidaysEnabled = False

			End Try

		End Sub

	End Class
End Namespace