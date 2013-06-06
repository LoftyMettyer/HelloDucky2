Imports SystemFramework.Things

Public Module COMInterfaces

  Public Interface ICommitDB
    Function ScriptTables() As Boolean
    Function ScriptTableViews() As Boolean
    Function ScriptObjects() As Boolean
    Function ScriptFunctions() As Boolean
    Function ScriptTriggers() As Boolean
    Function ScriptViews() As Boolean
    Function ScriptIndexes() As Boolean
    Function DropTableViews() As Boolean
    Function DropViews() As Boolean
    Function ApplySecurity() As Boolean
    Function ScriptOvernightStep2() As Boolean
  End Interface

  Public Interface ISystemManager
    Property MetadataDB As Object
    Property CommitDB As Object
    ReadOnly Property ErrorLog As ErrorHandler.Errors
    ReadOnly Property TuningLog As Tuning.Report
    Function GetTable(ByVal id As Integer) As Table
    ReadOnly Property Script As ScriptDB.Script
    ReadOnly Property Options As HCMOptions
    Function Initialise() As Boolean
    Function PopulateObjects() As Boolean
    Function CloseSafely() As Boolean
    ReadOnly Property Version As System.Version
    ReadOnly Property Modifications As Modifications
  End Interface

  Public Interface IErrors
    Sub OutputToFile(ByVal FileName As String)
    Sub Show()
    ReadOnly Property ErrorCount As Integer
    ReadOnly Property IsCatastrophic As Boolean
  End Interface

  Public Interface IForm
    Sub Show()
    Sub ShowDialog()
  End Interface

  Public Interface IOptions
    Property RefreshObjects As Boolean
    Property DevelopmentMode As Boolean
    Property OverflowSafety As Boolean
    Property OptimiseSaveProcess As Boolean
  End Interface

  Public Interface IConnection
    Sub Open()
    Sub Close()
    Function ExecStoredProcedure(ByVal ProcedureName As String, ByVal Parms As Connectivity.Parameters) As System.Data.DataSet
    Function ScriptStatement(ByVal statement As String, ByRef IsCritical As Boolean) As Boolean
    Property Login As Connectivity.Login
  End Interface

  Public Interface IObject
    Property Name As String
    ReadOnly Property PhysicalName As String
  End Interface

  Public Interface ITable
    Inherits IObject
    ' These eventually will be gotten rid of when we port the rest of sysmgr into this framework.
    Property SysMgrInsertTrigger As String
    Property SysMgrUpdateTrigger As String
    Property SysMgrDeleteTrigger As String
  End Interface

  Public Interface IWorkflowElement
    Property WorkflowID As Integer
    Property Caption As String
    Property ConnectionPairID As Integer
    Property LeftCoord As Integer
    Property TopCoord As Integer
    Property DecisionCaptionType As Integer
    Property Identifier As String
    Property TrueFlowIdentifier As String
    Property DataAction As Integer
    Property DataTableID As Integer
    Property DataRecord As Integer
    Property EmailID As Integer
    Property EmailRecord As Integer
    Property WebFormBGColor As Integer
    Property WebFormBGImageID As Integer
    Property WebFormBGImageLocation As Integer
    Property WebFormDefaultFontName As String
    Property WebFormDefaultFontSize As Integer
    Property WebFormDefaultFontBold As Boolean
    Property WebFormDefaultFontItalic As Boolean
    Property WebFormDefaultFontStrikeThru As Boolean
    Property WebFormDefaultFontUnderline As Boolean
    Property WebFormHeight As Integer
    Property WebFormWidth As Integer
    Property RecSelWebFormIdentifier As String
    Property RecSelIdentifier As String
    Property SecondaryDataRecord As Integer
    Property SecondaryRecSelWebFormIdentifier As String
    Property SecondaryRecSelIdentifier As String
    Property EmailSubject As String
    Property TimeoutFrequency As Integer
    Property TimeoutPeriod As Integer
    Property DataRecordTable As Integer
    Property SecondaryDataRecordTable As Integer
    Property TrueFlowType As Integer
    Property TrueFlowExprID As Integer
    Property DescriptionExprID As Integer
    Property WebFormFGColor As Integer
    Property DescHasWorkflowName As Boolean
    Property DescHasElementCaption As Boolean
    Property EmailCCID As Integer
    Property TimeoutExcludeWeekend As Boolean
    Property CompletionMessageType As Integer
    Property CompletionMessage As String
    Property SavedForLaterMessageType As Integer
    Property SavedForLaterMessage As String
    Property FollowOnFormsMessageType As Integer
    Property FollowOnFormsMessage As String
    Property Attachment_Type As String
    Property Attachment_File As String
    Property Attachment_WFElementIdentifier As String
    Property Attachment_WFValueIdentifier As String
    Property Attachment_DBColumnID As Integer
    Property Attachment_DBRecord As Integer
    Property Attachment_DBElement As String
    Property Attachment_DBValue As String
  End Interface

  Public Interface IWorkflowElementItem
    Property ID As Integer
    Property Description As String
    Property ItemType As Integer
    Property Caption As String
    Property DBColumnID As Integer
    Property DBRecord As Integer
    Property InputReturnType As Integer
    Property InputSize As Integer
    Property InputDecimals As Integer
    Property InputIdentifier As String
    Property InputDefault As String
    Property WFFormIdentifier As String
    Property WFValueIdentifier As String
    Property Left As Integer
    Property Top As Integer
    Property Width As Integer
    Property Height As Integer
    Property BackgroundColor As Integer
    Property ForegroundColor As Integer
    Property FontName As String
    Property FontSize As Integer
    Property FontBold As Boolean
    Property FontItalic As Boolean
    Property FontStrikeThru As Boolean
    Property FontUnderline As Boolean
    Property PictureID As Integer
    Property PictureBorder As Integer
    Property Alignment As Integer
    Property ZOrder As Integer
    Property TabIndex As Integer
    Property BackStyle As Integer
    Property BackColorEven As Integer
    Property BackColorOdd As Integer
    Property ColumnHeaders As String
    Property ForeColorEven As Integer
    Property ForeColorOdd As Integer
    Property HeaderBackColor As Integer
    Property HeadFontName As String
    Property HeadFontSize As Integer
    Property HeadFontBold As Integer
    Property HeadFontItalic As Integer
    Property HeadFontStrikeThru As Integer
    Property HeadFontUnderline As Integer
    Property Headlines As String
    Property TableID As Integer
    Property ForeColorHighlight As Integer
    Property BackColorHighlight As Integer
    Property ControlValues As String
    Property LookupTableID As Integer
    Property LookupColumnID As Integer
    Property RecordTableID As Integer
    Property Orientation As Integer
    Property RecordOrderID As Integer
    Property RecordFilterID As Integer
    Property Behaviour As Integer
    Property Mandatory As Boolean
    Property ExpressionID As Integer
    Property CaptionType As Integer
    Property DefaultValueType As Integer
    Property VerticalOffsetBehaviour As Integer
    Property HorizontalOffsetBehaviour As Integer
    Property VerticalOffset As Integer
    Property HorizontalOffset As Integer
    Property HeightBehaviour As Integer
    Property WidthBehaviour As Integer
    Property PasswordType As String
    Property FileExtensions As String
    Property LookupFilterColumnID As Integer
    Property LookupFilterOperator As String
    Property LookupFilterValue As String
  End Interface

  Public Interface IModifications
    Property StructureChanged As Boolean
    Property ExpressionChanged As Boolean
    Property ScreenChanged As Boolean
    Property WorkflowChanged As Boolean
    Property ModuleSetupChanged As Boolean
    Property PlatformChanged As Boolean
  End Interface


  'NOTE: NOT YET USED
  Public Interface IColumn
    Inherits IObject
    Property DataType As ColumnTypes
    Property Size As Integer
    Property Decimals As Integer
    Property Audit As Boolean
    Property Multiline As Boolean
    Property CalculateIfEmpty As Boolean
    Property IsReadOnly As Boolean
    Property CaseType As CaseType
    Property TrimType As TrimType
    Property Alignment As AlignType
    Property Mandatory As Boolean
    Property OLEType As OLEType
    Property DefaultCalcID As Integer
    Property DefaultCalculation As Expression
    Property DefaultValue As String
  End Interface

End Module
