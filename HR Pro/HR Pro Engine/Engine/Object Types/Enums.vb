Namespace Things

  Public Enum TableType
    Parent = 0
    Child = 1
    Lookup = 2
  End Enum

  Public Enum DisplayStyle
    Hierarchy = 0
    List = 1
    Details = 2
    Diagnostic = 3
  End Enum

  Public Enum DateOffsetType
    Day = 0
    Week = 1
    Month = 2
    Year = 3
  End Enum

  Public Enum EmailLinkType
    Column = 0
    Record = 1
    [Date] = 2
  End Enum

  Public Enum Order
    Descending = 0
    Ascending = 1
  End Enum

  Public Enum TrimType
    None = 0
    Both = 1
    Left = 2
    Right = 3
  End Enum

  Public Enum AlignType
    Left = 0
    Right = 1
    Center = 2          ' It pains me to use the American spelling, but probably best for consistency with other controls!
  End Enum

  Public Enum CaseType
    None = 0
    Upper = 1
    Lower = 2
    Proper = 3
  End Enum

  Public Enum ValidationType
    None = 0
    Duplicate = 1
    UniqueInTable = 2
    UniqueInSiblings = 3
    Mandatory = 4
    Offset = 5
  End Enum

  Public Enum ValidationSeverity
    Information = 0
    Warning = 1
    [Error] = 2
    Catastrophic = 3
  End Enum

  Public Enum UniqueCheckScope
    None = 0
    All = -1
    Sibling = -2
  End Enum

  Public Enum SettingType
    ModuleSetting = 1
    CodeItem = 2
    UpdateParameter = 3
    DefaultDataType = 4
  End Enum

  Public Enum AccordPurgeType
    Days = 0
    Weeks = 1
    Months = 2
    Years = 3
  End Enum

End Namespace
