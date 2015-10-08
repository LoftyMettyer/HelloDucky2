Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace ViewModels.Home

	Public Class EmptyOptionViewModel

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtAction As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtErrorMessage As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtFromDef As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtOrderID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtFilterSQL As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtFilterDef As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtRecordID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtLinkRecordID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtLookupColumnID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtColumnID As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtValue As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtFile As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtFileValue As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtResultCode As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtPreReqFails As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtUnAvailFails As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtOverlapFails As String
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtOverBooked As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Property txtSelectedRecordsInFindGrid As String


		Public Sub New()

			txtAction = NullSafeString(HttpContext.Current.Session("optionAction"))
			txtErrorMessage = NullSafeString(HttpContext.Current.Session("errorMessage"))
			txtFromDef = NullSafeString(HttpContext.Current.Session("fromDef"))
			txtOrderID = NullSafeString(HttpContext.Current.Session("orderID"))
			txtFilterSQL = NullSafeString(HttpContext.Current.Session("optionFilterSQL"))
			txtFilterDef = NullSafeString(HttpContext.Current.Session("optionFilterDef"))
			txtSelectedRecordsInFindGrid = NullSafeString(HttpContext.Current.Session("OptionSelectedRecordIds"))
			txtRecordID = NullSafeString(HttpContext.Current.Session("optionRecordID"))
			txtLinkRecordID = NullSafeString(HttpContext.Current.Session("optionLinkRecordID"))
			txtLookupColumnID = NullSafeString(HttpContext.Current.Session("optionLookupColumnID"))
			txtColumnID = NullSafeString(HttpContext.Current.Session("optionColumnID"))
			txtValue = NullSafeString(HttpContext.Current.Session("optionLookupValue"))
			txtFile = NullSafeString(HttpContext.Current.Session("optionFile"))
			txtFileValue = NullSafeString(HttpContext.Current.Session("optionFileValue"))
			txtResultCode = NullSafeString(HttpContext.Current.Session("TBResultCode"))
			txtPreReqFails = NullSafeString(HttpContext.Current.Session("PreReqFails"))
			txtUnAvailFails = NullSafeString(HttpContext.Current.Session("UnAvailFails"))
			txtOverlapFails = NullSafeString(HttpContext.Current.Session("OverlapFails"))
			txtOverBooked = NullSafeString(HttpContext.Current.Session("Overbooked"))
			
		End Sub


	End Class
End Namespace