Option Strict On
Option Explicit On

Imports System.ComponentModel.DataAnnotations

Namespace Models
	Public Class OptionDataModel

		Public Property txtOptionAction As OptionActionType
		Public Property txtoptionTableID As Integer
		Public Property txtOptionViewID As Integer
		Public Property txtOptionOrderID As Integer
		Public Property txtOptionColumnID As Integer 'string?
		Public Property txtOptionPageAction As OptionActionType
		Public Property txtOptionFirstRecPos As Integer
		Public Property txtOptionCurrentRecCount As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtGotoLocateValue As String ' htmlencode?

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtOptionCourseTitle As String ' htmlencode?

		Public Property txtOptionRecordID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtOptionLinkRecordID As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtOptionValue As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtOptionSQL As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtOptionPromptSQL As String

		Public Property txtOptionOnlyNumerics As Integer
		Public Property txtOptionLookupColumnID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtOptionLookupFilterValue As String	' htmlencode?
		Public Property txtOptionIsLookupTable As Boolean
		Public Property txtOptionParentTableID As Integer
		Public Property txtOptionParentRecordID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property txtOption1000SepCols As String

		Public Property txtStandardReportType As Integer




	End Class
End Namespace
