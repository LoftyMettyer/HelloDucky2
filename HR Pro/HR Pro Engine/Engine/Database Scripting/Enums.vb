Namespace ScriptDB

  <HideModuleName()> _
  Public Module Enums

    Public Enum TriggerType
      InsteadOfInsert = 0
      AfterInsert = 1
      InsteadOfUpdate = 2
      AfterUpdate = 3
      InsteadOfDelete = 4
      AfterDelete = 5
    End Enum

    Public Enum ObjectPrefix
      [NameOnly] = 0
      [Table] = 1
      [Column] = 2
      [Default] = 3
      [RecordDescription] = 4
      [Calculation] = 5
      [Validation] = 6
      [View] = 7
      [Index] = 8
      [Variable] = 9
    End Enum

    Public Enum ColumnTypes
      [Logic] = -7
      [Binary] = -4
      [Photograph] = -3
      [Link] = -2
      [WorkingPattern] = -1
      [Numeric] = 2
      [Integer] = 4
      [Date] = 11
      [Text] = 12
    End Enum

    Public Enum OLEType
      LocalPath = 0
      ServerPath = 1
      Linked = 2
      Embedded = 3
      Filestream = 4
    End Enum

    Public Enum RelationshipType
      [Child] = 0
      [Parent] = 1
      [Unknown] = 2
    End Enum

    Public Enum GenerateType
      [InlineCode] = 0
      [SimpleUDF] = 1
      [ComplexUDF] = 2
    End Enum

    Public Enum ExpressionType
      [ColumnDefault] = 0
      [ColumnCalculation] = 1
      [RecordDescription] = 2
      [Mask] = 3
      [RuntimeCode] = 4
      [Validation] = 5
      [DiaryFilter] = 6
      [ColumnFilter] = 7
      [ReferencedColumn] = 8
    End Enum

    Enum OperatorSubType
      [Comparison] = 177
      [Logic] = 178
      [Other] = 179
      [Modifier] = 180
    End Enum

    'Public Enum ColumnSummaryType
    '  [Row] = 0
    '  [Count] = 1
    '  [Total] = 2
    'End Enum

    Public Enum ColumnRowSelection
      '[None] = 0
      [First] = 1
      [Last] = 2
      [Specific] = 3
      [Total] = 4
      [Count] = 5
      ' [Maximum] = 6
      ' [Minimum] = 7
      ' [Average] = 8
    End Enum

    Public Enum ComponentTypes
      Column = 1
      [Function] = 2
      Calculation = 3
      Value = 4
      [Operator] = 5
      TableValue = 6
      PromptedValue = 7
      CustomCalc = 8  ' Not used.
      Expression = 9
      Filter = 10
      WorkflowValue = 11
      WorkflowColumn = 12
      Relation = 13
    End Enum

    Public Enum ComponentValueTypes
      [Unknown] = 0
      [String] = 1
      [Numeric] = 2
      [Logic] = 3
      [Date] = 4
      [SystemVariable] = 5
      [Condition] = 6            ' typically first parameter of a if then else (e.g. field = value)
      '[Another_Unknown] = 100
      '[Component_String] = 101
      '[Component_Numeric] = 102
      '[Component_Logic] = 103
      '[Component_Date] = 104
    End Enum

    'Public Enum ObjectTypes
    '  [All] = -2
    '  '[TablesAndViews] = -1
    '  Unknown = 0
    '  Table = 1
    '  Column = 2
    '  View = 3
    '  Screen = 4
    '  Link = 5
    '  Group = 6
    '  Validation = 7
    '  EmailAddress = 8
    '  Diary = 9
    '  OutlookFolder = 10
    '  CalculatedValue = 11
    '  Filter = 12
    '  ValidationExpression = 13
    '  User = 14
    '  Shortcut = 15
    '  Template = 16
    '  Control = 17
    '  [Function] = 18
    '  [Operator] = 19
    '  Parameter = 20
    '  Value = 21
    '  PromptedValue = 22
    '  PermissionFolder = 23
    '  Permission = 24
    '  [Configuration] = 25
    '  ConfigurationSetting = 26
    '  Orders = 27
    '  Folder = 28
    '  [Structure] = 29
    '  Security = 30
    '  [Tools] = 31
    '  [Resource] = 32
    '  '[Image] = 33
    '  Relationship = 34
    '  FindColumns = 35
    '  Report = 1000
    '  CustomReport = 1001
    '  Import = 1002
    '  Export = 1003
    '  MailMerge = 2001
    '  Workflow = 2002
    '  RuntimeCalculation = 3001
    '  RuntimeFilter = 3002
    '  AuditLog = 4001
    'End Enum

  End Module


End Namespace
