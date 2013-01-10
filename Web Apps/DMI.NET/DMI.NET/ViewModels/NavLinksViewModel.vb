Public Class NavLinksViewModel
	Public Property NumberOfLinks As Integer

	Public Property NavigationLinks As List(Of navigationLink)

End Class


Public Class navigationLink

	Sub New(p_ID As Long,
  p_DrillDownHidden As Boolean,
  p_LinkType As Integer,
  p_LinkOrder As Integer,
  p_Text As String,
  p_Text1 As String,
  p_Text2 As String,
  p_Prompt As String,
  p_ScreenID As Long,
  p_TableID As Long,
  p_ViewID As Long,
  p_PageTitle As String,
  p_URL As String,
  p_UtilityType As Integer,
  p_UtilityID As Long,
  p_NewWindow As Boolean,
  p_BaseTable As String,
  p_LinkToFind As Integer,
  p_SingleRecord As Integer,
  p_PrimarySequence As Integer,
  p_SecondarySequence As Integer,
  p_FindPage As Boolean,
  p_EmailAddress As String,
  p_EmailSubject As String,
  p_AppFilePath As String,
  p_AppParameters As String,
  p_DocumentFilePath As String,
  p_DisplayDocumentHyperlink As Boolean,
  p_IsSeparator As Boolean,
  p_Element_Type As Integer,
  p_SeparatorOrientation As Integer,
  p_PictureID As Long,
  p_Chart_ShowLegend As Boolean,
  p_Chart_Type As Integer,
  p_Chart_ShowGrid As Boolean,
  p_Chart_StackSeries As Boolean,
  p_Chart_ShowValues As Boolean,
  p_Chart_ViewID As Long,
  p_Chart_TableID As Long,
  p_Chart_ColumnID As Long,
  p_Chart_FilterID As Long,
  p_Chart_AggregateType As Long,
  p_Chart_ColumnName As String,
  p_Chart_ColumnName_2 As String,
  p_UseFormatting As Boolean,
  p_Formatting_DecimalPlaces As Integer,
  p_Formatting_Use1000Separator As Boolean,
  p_Formatting_Prefix As String,
  p_Formatting_Suffix As String,
  p_UseConditionalFormatting As Boolean,
  p_ConditionalFormatting_Operator_1 As String,
  p_ConditionalFormatting_Value_1 As String,
  p_ConditionalFormatting_Style_1 As String,
  p_ConditionalFormatting_Colour_1 As String,
  p_ConditionalFormatting_Operator_2 As String,
  p_ConditionalFormatting_Value_2 As String,
  p_ConditionalFormatting_Style_2 As String,
  p_ConditionalFormatting_Colour_2 As String,
  p_ConditionalFormatting_Operator_3 As String,
  p_ConditionalFormatting_Value_3 As String,
  p_ConditionalFormatting_Style_3 As String,
  p_ConditionalFormatting_Colour_3 As String,
  p_SeparatorColour As String,
  p_InitialDisplayMode As Integer,
  p_Chart_TableID_2 As Long,
  p_Chart_ColumnID_2 As Long,
  p_Chart_TableID_3 As Long,
  p_Chart_ColumnID_3 As Long,
  p_Chart_SortOrderID As Long,
  p_Chart_SortDirection As Integer,
  p_Chart_ColourID As Long,
  p_Chart_ShowPercentages As Boolean
  )
		ID = p_ID
		DrillDownHidden = p_DrillDownHidden
		LinkType = p_LinkType
		LinkOrder = p_LinkOrder
		Text = p_Text
		Text1 = p_Text1
		Text2 = p_Text2
		Prompt = p_Prompt
		ScreenID = p_ScreenID
		TableID = p_TableID
		ViewID = p_ViewID
		PageTitle = p_PageTitle
		URL = p_URL
		UtilityType = p_UtilityType
		UtilityID = p_UtilityID
		NewWindow = p_NewWindow
		BaseTable = p_BaseTable
		LinkToFind = p_LinkToFind
		SingleRecord = p_SingleRecord
		PrimarySequence = p_PrimarySequence
		SecondarySequence = p_SecondarySequence
		FindPage = p_FindPage
		EmailAddress = p_EmailAddress
		EmailSubject = p_EmailSubject
		AppFilePath = p_AppFilePath
		AppParameters = p_AppParameters
		DocumentFilePath = p_DocumentFilePath
		DisplayDocumentHyperlink = p_DisplayDocumentHyperlink
		IsSeparator = p_IsSeparator
		Element_Type = p_Element_Type
		SeparatorOrientation = p_SeparatorOrientation
		PictureID = p_PictureID
		Chart_ShowLegend = p_Chart_ShowLegend
		Chart_Type = p_Chart_Type
		Chart_ShowGrid = p_Chart_ShowGrid
		Chart_StackSeries = p_Chart_StackSeries
		Chart_ShowValues = p_Chart_ShowValues
		Chart_ViewID = p_Chart_ViewID
		Chart_TableID = p_Chart_TableID
		Chart_ColumnID = p_Chart_ColumnID
		Chart_FilterID = p_Chart_FilterID
		Chart_AggregateType = p_Chart_AggregateType
		Chart_ColumnName = p_Chart_ColumnName
		Chart_ColumnName_2 = p_Chart_ColumnName_2
		UseFormatting = p_UseFormatting
		Formatting_DecimalPlaces = p_Formatting_DecimalPlaces
		Formatting_Use1000Separator = p_Formatting_Use1000Separator
		Formatting_Prefix = p_Formatting_Prefix
		Formatting_Suffix = p_Formatting_Suffix
		UseConditionalFormatting = p_UseConditionalFormatting
		ConditionalFormatting_Operator_1 = p_ConditionalFormatting_Operator_1
		ConditionalFormatting_Value_1 = p_ConditionalFormatting_Value_1
		ConditionalFormatting_Style_1 = p_ConditionalFormatting_Style_1
		ConditionalFormatting_Colour_1 = p_ConditionalFormatting_Colour_1
		ConditionalFormatting_Operator_2 = p_ConditionalFormatting_Operator_2
		ConditionalFormatting_Value_2 = p_ConditionalFormatting_Value_2
		ConditionalFormatting_Style_2 = p_ConditionalFormatting_Style_2
		ConditionalFormatting_Colour_2 = p_ConditionalFormatting_Colour_2
		ConditionalFormatting_Operator_3 = p_ConditionalFormatting_Operator_3
		ConditionalFormatting_Value_3 = p_ConditionalFormatting_Value_3
		ConditionalFormatting_Style_3 = p_ConditionalFormatting_Style_3
		ConditionalFormatting_Colour_3 = p_ConditionalFormatting_Colour_3
		SeparatorColour = p_SeparatorColour
		InitialDisplayMode = p_InitialDisplayMode
		Chart_TableID_2 = p_Chart_TableID_2
		Chart_ColumnID_2 = p_Chart_ColumnID_2
		Chart_TableID_3 = p_Chart_TableID_3
		Chart_ColumnID_3 = p_Chart_ColumnID_3
		Chart_SortOrderID = p_Chart_SortOrderID
		Chart_SortDirection = p_Chart_SortDirection
		Chart_ColourID = p_Chart_ColourID
		Chart_ShowPercentages = p_Chart_ShowPercentages

	End Sub


	' mmmmm auto-implemented properties.....


	Public Property ID As Long
	Public Property DrillDownHidden As Boolean
	Public Property LinkType As Integer
	Public Property LinkOrder As Integer
	Public Property Text As String
	Public Property Text1 As String
	Public Property Text2 As String
	Public Property Prompt As String
	Public Property ScreenID As Long
	Public Property TableID As Long
	Public Property ViewID As Long
	Public Property PageTitle As String
	Public Property URL As String
	Public Property UtilityType As Integer
	Public Property UtilityID As Long
	Public Property NewWindow As Boolean
	Public Property BaseTable As String
	Public Property LinkToFind As Integer
	Public Property SingleRecord As Integer
	Public Property PrimarySequence As Integer
	Public Property SecondarySequence As Integer
	Public Property FindPage As Boolean
	Public Property EmailAddress As String
	Public Property EmailSubject As String
	Public Property AppFilePath As String
	Public Property AppParameters As String
	Public Property DocumentFilePath As String
	Public Property DisplayDocumentHyperlink As Boolean
	Public Property IsSeparator As Boolean
	Public Property Element_Type As Integer
	Public Property SeparatorOrientation As Integer
	Public Property PictureID As Long
	Public Property Chart_ShowLegend As Boolean
	Public Property Chart_Type As Integer
	Public Property Chart_ShowGrid As Boolean
	Public Property Chart_StackSeries As Boolean
	Public Property Chart_ShowValues As Boolean
	Public Property Chart_ViewID As Long
	Public Property Chart_TableID As Long
	Public Property Chart_ColumnID As Long
	Public Property Chart_FilterID As Long
	Public Property Chart_AggregateType As Long
	Public Property Chart_ColumnName As String
	Public Property Chart_ColumnName_2 As String
	Public Property UseFormatting As Boolean
	Public Property Formatting_DecimalPlaces As Integer
	Public Property Formatting_Use1000Separator As Boolean
	Public Property Formatting_Prefix As String
	Public Property Formatting_Suffix As String
	Public Property UseConditionalFormatting As Boolean
	Public Property ConditionalFormatting_Operator_1 As String
	Public Property ConditionalFormatting_Value_1 As String
	Public Property ConditionalFormatting_Style_1 As String
	Public Property ConditionalFormatting_Colour_1 As String
	Public Property ConditionalFormatting_Operator_2 As String
	Public Property ConditionalFormatting_Value_2 As String
	Public Property ConditionalFormatting_Style_2 As String
	Public Property ConditionalFormatting_Colour_2 As String
	Public Property ConditionalFormatting_Operator_3 As String
	Public Property ConditionalFormatting_Value_3 As String
	Public Property ConditionalFormatting_Style_3 As String
	Public Property ConditionalFormatting_Colour_3 As String
	Public Property SeparatorColour As String
	Public Property InitialDisplayMode As Integer
	Public Property Chart_TableID_2 As Long
	Public Property Chart_ColumnID_2 As Long
	Public Property Chart_TableID_3 As Long
	Public Property Chart_ColumnID_3 As Long
	Public Property Chart_SortOrderID As Long
	Public Property Chart_SortDirection As Integer
	Public Property Chart_ColourID As Long
	Public Property Chart_ShowPercentages As Boolean

End Class
