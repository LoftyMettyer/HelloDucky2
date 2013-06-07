Namespace Enums

  Public Enum ComponentValueTypes
    [Unknown] = 0
    [String] = 1
    [Numeric] = 2
    [Logic] = 3
    [Date] = 4
    [SystemVariable] = 5
    [Condition] = 6            ' typically first parameter of a if then else (e.g. field = value)
    '[Another_Unknown] = 100
    ByRefString = 101
    ByRefNumeric = 102
    ByRefLogic = 103
    ByRefDate = 104
  End Enum

End Namespace