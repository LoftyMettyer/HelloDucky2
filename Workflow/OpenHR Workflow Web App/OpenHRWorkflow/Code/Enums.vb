
Public Enum SqlDataType
  Unknown = 0      ' ?
  Ole = -4         ' OLE columns
  [Boolean] = -7     ' Logic columns
  Numeric = 2      ' Numeric columns
  [Integer] = 4      ' Integer columns
  [Date] = 11        ' Date columns
  VarChar = 12     ' Character columns
  VarBinary = -3   ' Photo columns
  LongVarChar = -1 ' Working Pattern columns
End Enum

Public Enum FilterOperators
  giFILTEROP_UNDEFINED = 0
  giFILTEROP_EQUALS = 1
  giFILTEROP_NOTEQUALTO = 2
  giFILTEROP_ISATMOST = 3
  giFILTEROP_ISATLEAST = 4
  giFILTEROP_ISMORETHAN = 5
  giFILTEROP_ISLESSTHAN = 6
  giFILTEROP_ON = 7
  giFILTEROP_NOTON = 8
  giFILTEROP_AFTER = 9
  giFILTEROP_BEFORE = 10
  giFILTEROP_ONORAFTER = 11
  giFILTEROP_ONORBEFORE = 12
  giFILTEROP_CONTAINS = 13
  giFILTEROP_IS = 14
  giFILTEROP_DOESNOTCONTAIN = 15
  giFILTEROP_ISNOT = 16
End Enum