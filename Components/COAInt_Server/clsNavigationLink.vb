Option Strict Off
Option Explicit On
Public Class clsNavigationLink

  Private miLinkType As Short
  Private miLinkOrder As Short
  Private msText As String
  Private msPrompt As String
  Private mlngScreenID As Integer
  Private msPageTitle As String
  Private msURL As String
  Private msEmailAddress As String
  Private msEmailSubject As String
  Private msAppFilePath As String
  Private msAppParameters As String
  Private msDocumentFilePath As String
  Private mfDisplayDocumentHyperlink As Boolean
  ' Private mfIsSeparator As Boolean
  Private miSeparatorOrientation As Short
  Private mlngPictureID As Integer
  Private mfChart_ShowLegend As Boolean
  Private miChart_Type As Short
  Private mfChart_ShowGrid As Boolean
  Private mfChart_StackSeries As Boolean
  Private mlngChart_ViewID As Integer
  Private mlngChart_TableID As Integer
  Private mlngChart_ColumnID As Integer
  Private mlngChart_FilterID As Integer
  Private mlngChart_AggregateType As Integer
  Private miElement_Type As Short
  Private mfChart_ShowValues As Boolean

  Private mfUseFormatting As Boolean
  Private miFormatting_DecimalPlaces As Short
  Private mfFormatting_Use1000Separator As Boolean
  Private msFormatting_Prefix As String
  Private msFormatting_Suffix As String

  Private mfUseConditionalFormatting As Boolean
  Private msConditionalFormatting_Operator_1 As String
  Private msConditionalFormatting_Value_1 As String
  Private msConditionalFormatting_Style_1 As String
  Private msConditionalFormatting_Colour_1 As String
  Private msConditionalFormatting_Operator_2 As String
  Private msConditionalFormatting_Value_2 As String
  Private msConditionalFormatting_Style_2 As String
  Private msConditionalFormatting_Colour_2 As String
  Private msConditionalFormatting_Operator_3 As String
  Private msConditionalFormatting_Value_3 As String
  Private msConditionalFormatting_Style_3 As String
  Private msConditionalFormatting_Colour_3 As String

  Private msSeparatorColour As String

  Private miInitialDisplayMode As Short
  Private mlngChart_TableID_2 As Integer
  Private mlngChart_ColumnID_2 As Integer
  Private mlngChart_TableID_3 As Integer
  Private mlngChart_ColumnID_3 As Integer
  Private mlngChart_SortOrderID As Integer
  Private miChart_SortDirection As Short
  Private mlngChart_ColourID As Integer
  Private mfChart_ShowPercentages As Boolean

  Private msChart_ColumnName As String
  Private msChart_ColumnName_2 As String

  Private mlngID As Integer
  Private miUtilityType As Short
  Private mlngUtilityID As Integer
  Private mbNewWindow As Boolean
  Private msBaseTable As String
  Private msText1 As String
  Private msText2 As String
  Private miSingleRecord As Short
  Private miLinkToFind As Short
  Private mlngTableID As Integer
  Private mlngViewID As Integer
  Private miPrimarySequence As Short
  Private miSecondarySequence As Short
  Private mbFindPage As Short
  Private mfDrillDownHidden As Boolean


  Public Property ID() As Integer
    Get
      ID = mlngID
    End Get
    Set(ByVal Value As Integer)
      mlngID = Value
    End Set
  End Property


  Public Property DrillDownHidden() As Boolean
    Get
      DrillDownHidden = mfDrillDownHidden
    End Get
    Set(ByVal Value As Boolean)
      mfDrillDownHidden = Value
    End Set
  End Property


  Public Property LinkType() As Short
    Get
      LinkType = miLinkType
    End Get
    Set(ByVal Value As Short)
      miLinkType = Value
    End Set
  End Property


  Public Property LinkOrder() As Short
    Get
      LinkOrder = miLinkOrder
    End Get
    Set(ByVal Value As Short)
      miLinkOrder = Value
    End Set
  End Property


  Public Property Text() As String
    Get
      Text = msText
    End Get
    Set(ByVal Value As String)
      msText = Value
    End Set
  End Property


  Public Property Text1() As String
    Get
      Text1 = msText1
    End Get
    Set(ByVal Value As String)
      msText1 = Value
    End Set
  End Property


  Public Property Text2() As String
    Get
      Text2 = msText2
    End Get
    Set(ByVal Value As String)
      msText2 = Value
    End Set
  End Property


  Public Property Prompt() As String
    Get
      Prompt = msPrompt
    End Get
    Set(ByVal Value As String)
      msPrompt = Value
    End Set
  End Property


  Public Property ScreenID() As Integer
    Get
      ScreenID = mlngScreenID
    End Get
    Set(ByVal Value As Integer)
      mlngScreenID = Value
    End Set
  End Property


  Public Property TableID() As Integer
    Get
      TableID = mlngTableID
    End Get
    Set(ByVal Value As Integer)
      mlngTableID = Value
    End Set
  End Property


  Public Property ViewID() As Integer
    Get
      ViewID = mlngViewID
    End Get
    Set(ByVal Value As Integer)
      mlngViewID = Value
    End Set
  End Property


  Public Property PageTitle() As String
    Get
      PageTitle = msPageTitle
    End Get
    Set(ByVal Value As String)
      msPageTitle = Value
    End Set
  End Property


  Public Property URL() As String
    Get
      URL = msURL
    End Get
    Set(ByVal Value As String)
      msURL = Value
    End Set
  End Property


  Public Property UtilityType() As Short
    Get
      UtilityType = miUtilityType
    End Get
    Set(ByVal Value As Short)
      miUtilityType = Value
    End Set
  End Property


  Public Property UtilityID() As Integer
    Get
      UtilityID = mlngUtilityID
    End Get
    Set(ByVal Value As Integer)
      mlngUtilityID = Value
    End Set
  End Property


  Public Property NewWindow() As Boolean
    Get
      NewWindow = mbNewWindow
    End Get
    Set(ByVal Value As Boolean)
      mbNewWindow = Value
    End Set
  End Property


  Public Property BaseTable() As String
    Get
      BaseTable = msBaseTable
    End Get
    Set(ByVal Value As String)
      msBaseTable = Value
    End Set
  End Property


  Public Property LinkToFind() As Short
    Get
      LinkToFind = miLinkToFind
    End Get
    Set(ByVal Value As Short)
      miLinkToFind = Value
    End Set
  End Property


  Public Property SingleRecord() As Short
    Get
      SingleRecord = miSingleRecord
    End Get
    Set(ByVal Value As Short)
      miSingleRecord = Value
    End Set
  End Property


  Public Property PrimarySequence() As Short
    Get
      PrimarySequence = miPrimarySequence
    End Get
    Set(ByVal Value As Short)
      miPrimarySequence = Value
    End Set
  End Property


  Public Property SecondarySequence() As Short
    Get
      SecondarySequence = miSecondarySequence
    End Get
    Set(ByVal Value As Short)
      miSecondarySequence = Value
    End Set
  End Property


  Public Property FindPage() As Boolean
    Get
      FindPage = mbFindPage
    End Get
    Set(ByVal Value As Boolean)
      mbFindPage = Value
    End Set
  End Property

  Public Property EmailAddress() As String
    Get
      EmailAddress = msEmailAddress
    End Get
    Set(ByVal Value As String)
      msEmailAddress = Value
    End Set
  End Property

  Public Property EmailSubject() As String
    Get
      EmailSubject = msEmailSubject
    End Get
    Set(ByVal Value As String)
      msEmailSubject = Value
    End Set
  End Property

  Public Property AppFilePath() As String
    Get
      AppFilePath = msAppFilePath
    End Get
    Set(ByVal Value As String)
      msAppFilePath = Value
    End Set
  End Property

  Public Property AppParameters() As String
    Get
      AppParameters = msAppParameters
    End Get
    Set(ByVal Value As String)
      msAppParameters = Value
    End Set
  End Property

  Public Property DocumentFilePath() As String
    Get
      DocumentFilePath = msDocumentFilePath
    End Get
    Set(ByVal Value As String)
      msDocumentFilePath = Value
    End Set
  End Property

  Public Property DisplayDocumentHyperlink() As Boolean
    Get
      DisplayDocumentHyperlink = mfDisplayDocumentHyperlink
    End Get
    Set(ByVal Value As Boolean)
      mfDisplayDocumentHyperlink = Value
    End Set
  End Property

  Public Property IsSeparator() As Boolean
    Get
      '  IsSeparator = mfIsSeparator
    End Get
    Set(ByVal Value As Boolean)
      '  mfIsSeparator = pfNewValue
    End Set
  End Property

  Public Property Element_Type() As Short
    Get
      Element_Type = miElement_Type
    End Get
    Set(ByVal Value As Short)
      miElement_Type = Value
    End Set
  End Property


  Public Property SeparatorOrientation() As Short
    Get
      SeparatorOrientation = miSeparatorOrientation
    End Get
    Set(ByVal Value As Short)
      miSeparatorOrientation = Value
    End Set
  End Property

  Public Property PictureID() As Integer
    Get
      PictureID = mlngPictureID
    End Get
    Set(ByVal Value As Integer)
      mlngPictureID = Value
    End Set
  End Property

  Public Property Chart_ShowLegend() As Boolean
    Get
      Chart_ShowLegend = mfChart_ShowLegend
    End Get
    Set(ByVal Value As Boolean)
      mfChart_ShowLegend = Value
    End Set
  End Property

  Public Property Chart_Type() As Short
    Get
      Chart_Type = miChart_Type
    End Get
    Set(ByVal Value As Short)
      miChart_Type = Value
    End Set
  End Property

  Public Property Chart_ShowGrid() As Boolean
    Get
      Chart_ShowGrid = mfChart_ShowGrid
    End Get
    Set(ByVal Value As Boolean)
      mfChart_ShowGrid = Value
    End Set
  End Property

  Public Property Chart_StackSeries() As Boolean
    Get
      Chart_StackSeries = mfChart_StackSeries
    End Get
    Set(ByVal Value As Boolean)
      mfChart_StackSeries = Value
    End Set
  End Property

  Public Property Chart_ShowValues() As Boolean
    Get
      Chart_ShowValues = mfChart_ShowValues
    End Get
    Set(ByVal Value As Boolean)
      mfChart_ShowValues = Value
    End Set
  End Property

  Public Property Chart_ViewID() As Integer
    Get
      Chart_ViewID = mlngChart_ViewID
    End Get
    Set(ByVal Value As Integer)
      mlngChart_ViewID = Value
    End Set
  End Property

  Public Property Chart_TableID() As Integer
    Get
      Chart_TableID = mlngChart_TableID
    End Get
    Set(ByVal Value As Integer)
      mlngChart_TableID = Value
    End Set
  End Property

  Public Property Chart_ColumnID() As Integer
    Get
      Chart_ColumnID = mlngChart_ColumnID
    End Get
    Set(ByVal Value As Integer)
      mlngChart_ColumnID = Value
    End Set
  End Property

  Public Property Chart_FilterID() As Integer
    Get
      Chart_FilterID = mlngChart_FilterID
    End Get
    Set(ByVal Value As Integer)
      mlngChart_FilterID = Value
    End Set
  End Property

  Public Property Chart_AggregateType() As Integer
    Get
      Chart_AggregateType = mlngChart_AggregateType
    End Get
    Set(ByVal Value As Integer)
      mlngChart_AggregateType = Value
    End Set
  End Property

  Public Property Chart_ColumnName() As String
    Get
      Chart_ColumnName = msChart_ColumnName
    End Get
    Set(ByVal Value As String)
      msChart_ColumnName = Value
    End Set
  End Property

  Public Property Chart_ColumnName_2() As String
    Get
      Chart_ColumnName_2 = msChart_ColumnName_2
    End Get
    Set(ByVal Value As String)
      msChart_ColumnName_2 = Value
    End Set
  End Property

  Public Property UseFormatting() As Boolean
    Get
      UseFormatting = mfUseFormatting
    End Get
    Set(ByVal Value As Boolean)
      mfUseFormatting = Value
    End Set
  End Property

  Public Property Formatting_DecimalPlaces() As Short
    Get
      Formatting_DecimalPlaces = miFormatting_DecimalPlaces
    End Get
    Set(ByVal Value As Short)
      miFormatting_DecimalPlaces = Value
    End Set
  End Property

  Public Property Formatting_Use1000Separator() As Boolean
    Get
      Formatting_Use1000Separator = mfFormatting_Use1000Separator
    End Get
    Set(ByVal Value As Boolean)
      mfFormatting_Use1000Separator = Value
    End Set
  End Property

  Public Property Formatting_Prefix() As String
    Get
      Formatting_Prefix = msFormatting_Prefix
    End Get
    Set(ByVal Value As String)
      msFormatting_Prefix = Value
    End Set
  End Property

  Public Property Formatting_Suffix() As String
    Get
      Formatting_Suffix = msFormatting_Suffix
    End Get
    Set(ByVal Value As String)
      msFormatting_Suffix = Value
    End Set
  End Property

  Public Property UseConditionalFormatting() As Boolean
    Get
      UseConditionalFormatting = mfUseConditionalFormatting
    End Get
    Set(ByVal Value As Boolean)
      mfUseConditionalFormatting = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Operator_1() As String
    Get
      ConditionalFormatting_Operator_1 = msConditionalFormatting_Operator_1
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Operator_1 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Value_1() As String
    Get
      ConditionalFormatting_Value_1 = msConditionalFormatting_Value_1
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Value_1 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Style_1() As String
    Get
      ConditionalFormatting_Style_1 = msConditionalFormatting_Style_1
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Style_1 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Colour_1() As String
    Get
      ConditionalFormatting_Colour_1 = msConditionalFormatting_Colour_1
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Colour_1 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Operator_2() As String
    Get
      ConditionalFormatting_Operator_2 = msConditionalFormatting_Operator_2
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Operator_2 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Value_2() As String
    Get
      ConditionalFormatting_Value_2 = msConditionalFormatting_Value_2
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Value_2 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Style_2() As String
    Get
      ConditionalFormatting_Style_2 = msConditionalFormatting_Style_2
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Style_2 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Colour_2() As String
    Get
      ConditionalFormatting_Colour_2 = msConditionalFormatting_Colour_2
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Colour_2 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Operator_3() As String
    Get
      ConditionalFormatting_Operator_3 = msConditionalFormatting_Operator_3
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Operator_3 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Value_3() As String
    Get
      ConditionalFormatting_Value_3 = msConditionalFormatting_Value_3
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Value_3 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Style_3() As String
    Get
      ConditionalFormatting_Style_3 = msConditionalFormatting_Style_3
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Style_3 = Value
    End Set
  End Property

  Public Property ConditionalFormatting_Colour_3() As String
    Get
      ConditionalFormatting_Colour_3 = msConditionalFormatting_Colour_3
    End Get
    Set(ByVal Value As String)
      msConditionalFormatting_Colour_3 = Value
    End Set
  End Property

  Public Property SeparatorColour() As String
    Get
      SeparatorColour = msSeparatorColour
    End Get
    Set(ByVal Value As String)
      msSeparatorColour = Value
    End Set
  End Property

  Public Property InitialDisplayMode() As Short
    Get
      InitialDisplayMode = miInitialDisplayMode
    End Get
    Set(ByVal Value As Short)
      miInitialDisplayMode = Value
    End Set
  End Property

  Public Property Chart_TableID_2() As Integer
    Get
      Chart_TableID_2 = mlngChart_TableID_2
    End Get
    Set(ByVal Value As Integer)
      mlngChart_TableID_2 = Value
    End Set
  End Property

  Public Property Chart_ColumnID_2() As Integer
    Get
      Chart_ColumnID_2 = mlngChart_ColumnID_2
    End Get
    Set(ByVal Value As Integer)
      mlngChart_ColumnID_2 = Value
    End Set
  End Property

  Public Property Chart_TableID_3() As Integer
    Get
      Chart_TableID_3 = mlngChart_TableID_3
    End Get
    Set(ByVal Value As Integer)
      mlngChart_TableID_3 = Value
    End Set
  End Property

  Public Property Chart_ColumnID_3() As Integer
    Get
      Chart_ColumnID_3 = mlngChart_ColumnID_3
    End Get
    Set(ByVal Value As Integer)
      mlngChart_ColumnID_3 = Value
    End Set
  End Property

  Public Property Chart_SortOrderID() As Integer
    Get
      Chart_SortOrderID = mlngChart_SortOrderID
    End Get
    Set(ByVal Value As Integer)
      mlngChart_SortOrderID = Value
    End Set
  End Property

  Public Property Chart_SortDirection() As Short
    Get
      Chart_SortDirection = miChart_SortDirection
    End Get
    Set(ByVal Value As Short)
      miChart_SortDirection = Value
    End Set
  End Property

  Public Property Chart_ColourID() As Integer
    Get
      Chart_ColourID = mlngChart_ColourID
    End Get
    Set(ByVal Value As Integer)
      mlngChart_ColourID = Value
    End Set
  End Property

  Public Property Chart_ShowPercentages() As Boolean
    Get
      Chart_ShowPercentages = mfChart_ShowPercentages
    End Get
    Set(ByVal Value As Boolean)
      mfChart_ShowPercentages = Value
    End Set
  End Property
End Class