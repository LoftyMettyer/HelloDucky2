Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums

Namespace Metadata

	Public Class Link
		Inherits Base

		Public LinkType As Short
		Public LinkOrder As Short
		Public Text As String
		Public Prompt As String
		Public ScreenID As Integer
		Public PageTitle As String
		Public URL As String
		Public EmailAddress As String
		Public EmailSubject As String
		Public AppFilePath As String
		Public AppParameters As String
		Public DocumentFilePath As String
		Public DisplayDocumentHyperlink As Boolean
		Public SeparatorOrientation As Short
		Public PictureID As Integer
		Public Chart_ShowLegend As Boolean
		Public Chart_Type As Short
		Public Chart_ShowGrid As Boolean
		Public Chart_StackSeries As Boolean
		Public Chart_ViewID As Integer
		Public Chart_TableID As Integer
		Public Chart_ColumnID As Integer
		Public Chart_FilterID As Integer
		Public Chart_AggregateType As Integer
		Public Element_Type As ElementType
		Public Chart_ShowValues As Boolean

		Public UseFormatting As Boolean
		Public Formatting_DecimalPlaces As Short
		Public Formatting_Use1000Separator As Boolean
		Public Formatting_Prefix As String
		Public Formatting_Suffix As String

		Public UseConditionalFormatting As Boolean
		Public ConditionalFormatting_Operator_1 As String
		Public ConditionalFormatting_Value_1 As String
		Public ConditionalFormatting_Style_1 As String
		Public ConditionalFormatting_Colour_1 As String
		Public ConditionalFormatting_Operator_2 As String
		Public ConditionalFormatting_Value_2 As String
		Public ConditionalFormatting_Style_2 As String
		Public ConditionalFormatting_Colour_2 As String
		Public ConditionalFormatting_Operator_3 As String
		Public ConditionalFormatting_Value_3 As String
		Public ConditionalFormatting_Style_3 As String
		Public ConditionalFormatting_Colour_3 As String

		Public SeparatorColour As String

		Public InitialDisplayMode As Short
		Public Chart_TableID_2 As Integer
		Public Chart_ColumnID_2 As Integer
		Public Chart_TableID_3 As Integer
		Public Chart_ColumnID_3 As Integer
		Public Chart_SortOrderID As Integer
		Public Chart_SortDirection As Short
		Public Chart_ColourID As Integer
		Public Chart_ShowPercentages As Boolean

		Public Chart_ColumnName As String
		Public Chart_ColumnName_2 As String

		Public UtilityType As UtilityType
		Public UtilityID As Integer
		Public NewWindow As Boolean
		Public BaseTableID As Long
		Public Text1 As String
		Public Text2 As String
		Public SingleRecord As Short
		Public LinkToFind As Short
		Public TableID As Integer
		Public ViewID As Integer
		Public PrimarySequence As Short
		Public SecondarySequence As Short
		Public FindPage As Short
		Public DrillDownHidden As Boolean

		Public IsSeparator As Boolean
	End Class

End Namespace