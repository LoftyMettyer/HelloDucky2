Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace ViewModels.Home

	Public Class DataViewModel

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtAction As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtReaction As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtCurrentTableID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtCurrentScreenID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtCurrentViewID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtSelectSQL As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtFromDef As String

		<AllowHtml>
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtFilterSQL As String

		<AllowHtml>
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtFilterDef As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtRealSource As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtOriginalRecordID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtRecordID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtParentTableID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtParentRecordID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtDefaultCalcCols As String

		<AllowHtml>
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtInsertUpdateDef As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtTimestamp As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtTBCourseRecordID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtTBEmployeeRecordID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtTBBookingStatusValue As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtTBOverride As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtTBCreateWLRecords As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtReportBaseTableID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtReportParent1TableID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtReportParent2TableID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtReportChildTableID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtUserChoice As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtParam1 As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtELFilterUser As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtELFilterType As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtELFilterStatus As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtELFilterMode As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtELOrderColumn As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtELOrderOrder As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtELAction As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Property txtELCurrRecCount As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtEL1stRecPos As String


	End Class
End Namespace