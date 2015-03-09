Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Models
	Public Class GotoOptionDataModel

		Public Property txtGotoOptionAction As OptionActionType
		Public Property txtGotoOptionTableID As Integer
		Public Property txtGotoOptionViewID As Integer
		Public Property txtGotoOptionOrderID As Integer
		Public Property txtGotoOptionColumnID As Integer 'string?
		Public Property txtGotoOptionPageAction As OptionActionType
		Public Property txtGotoOptionFirstRecPos As Integer
		Public Property txtGotoOptionCurrentRecCount As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoLocateValue As String ' htmlencode?

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoOptionCourseTitle As String ' htmlencode?

		Public Property txtGotoOptionRecordID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoOptionLinkRecordID As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoOptionValue As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoOptionSQL As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoOptionPromptSQL As String

		Public Property txtOptionOnlyNumerics As Integer
		Public Property txtGotoOptionLookupColumnID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoOptionLookupFilterValue As String	' htmlencode?
		Public Property txtGotoOptionIsLookupTable As Boolean
		Public Property txtGotoOptionParentTableID As Integer
		Public Property txtGotoOptionParentRecordID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoOption1000SepCols As String

		Public Property txtStandardReportType As Integer


		Public Property txtGotoOptionScreenID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoOptionFilterDef As String

		Public Property txtGotoOptionFilterSQL As String

		Public Property txtGotoOptionLinkTableID As Integer
		Public Property txtGotoOptionLinkOrderID As Integer
		Public Property txtGotoOptionLinkViewID As Integer

		Public Property txtGotoOptionLookupMandatory As String
		Public Property txtGotoOptionLookupValue As String

		Public Property txtGotoOptionFile As String
		Public Property txtGotoOptionExtension As String

		Public Property txtGotoOptionExprType As Integer
		Public Property txtGotoOptionExprID As Integer
		Public Property txtGotoOptionFunctionID As Integer
		Public Property txtGotoOptionParameterIndex As Integer
		Public Property txtGotoOptionRealsource As String

		Public Property txtGotoOptionDefSelType As Integer
		Public Property txtGotoOptionDefSelRecordID As Integer
		Public Property txtGotoOptionOLEType As Integer
		Public Property txtGotoOptionOLEMaxEmbedSize As Integer
		Public Property txtGotoOptionOLEReadOnly As Boolean
		Public Property txtGotoOptionIsPhoto As Boolean



	End Class
End Namespace
