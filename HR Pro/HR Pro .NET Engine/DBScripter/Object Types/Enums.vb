Namespace Things

  <HideModuleName()> _
  Public Module Enums

    Public Enum Type
      All = -1
      Unknown = 0
      Table = 1
      Column = 2
      View = 3
      Expression = 4
      Value = 5
      Component = 6
      Validation = 7
      CodeLibrary = 8
      Relation = 9
      Workflow = 10
      WorkflowElement = 11
      WorkflowElementColumn = 12
      WorkflowElementItem = 13
      Screen = 14
      DiaryLink = 15
      RecordDescription = 16
      Setting = 17
      EmailAddress = 18
      EmailLink = 19
      DataSource = 20
      TableOrder = 21
      TableOrderItem = 22
      GlobalModify = 23
      GlobalModifyItem = 24
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

  End Module

End Namespace
