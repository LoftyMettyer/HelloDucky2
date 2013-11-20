Option Strict Off
Option Explicit On
Module modUserSettings

	Public gstrSettingWordTemplate As String
	Public gstrSettingExcelTemplate As String
	Public gblnSettingExcelGridlines As Boolean
	Public gblnSettingExcelHeaders As Boolean
	Public gblnSettingExcelOmitSpacerRow As Boolean
	Public gblnSettingExcelOmitSpacerCol As Boolean
	Public gblnSettingAutoFitCols As Boolean
	Public gblnSettingLandscape As Boolean

	Public glngSettingTitleCol As Integer
	Public glngSettingTitleRow As Integer
	Public glngSettingDataCol As Integer
	Public glngSettingDataRow As Integer

	Public gblnSettingTitleGridlines As Boolean
	Public gblnSettingTitleBold As Boolean
	Public gblnSettingTitleUnderline As Boolean
	Public glngSettingTitleBackcolour As Integer
	Public glngSettingTitleForecolour As Integer
	Public glngSettingTitleBackcolour97 As Integer
	Public glngSettingTitleForecolour97 As Integer

	Public gblnSettingHeadingGridlines As Boolean
	Public gblnSettingHeadingBold As Boolean
	Public gblnSettingHeadingUnderline As Boolean
	Public glngSettingHeadingBackcolour As Integer
	Public glngSettingHeadingForecolour As Integer
	Public glngSettingHeadingBackcolour97 As Integer
	Public glngSettingHeadingForecolour97 As Integer

	Public gblnSettingDataGridlines As Boolean
	Public gblnSettingDataBold As Boolean
	Public gblnSettingDataUnderline As Boolean
	Public glngSettingDataBackcolour As Integer
	Public glngSettingDataForecolour As Integer
	Public glngSettingDataBackcolour97 As Integer
	Public glngSettingDataForecolour97 As Integer

	Public gblnEmailSystemPermission As Boolean
End Module