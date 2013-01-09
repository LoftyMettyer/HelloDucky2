Namespace Connectivity

  Public Enum ConnectionType
    WindowsAuthentication = 0
    SQLAuthentication = 1
    OracleAuthentication = 2
    SomethingElse = 3
  End Enum

  Public Enum Provider
    SQL = 0
    SQLContext = 1
    Oracle = 2
    Access = 3
    FoxPro = 4
    SQLExpress = 5
  End Enum

  Public Enum MetadataProvider
    LegacyDAO = 0
    PhoenixStoredProcs = 1
  End Enum

  Public Enum DBType
    [Integer] = 0
    [Numeric] = 1
    [String] = 2
    [Date] = 3
    [Boolean] = 4
    [Variant] = 5
    [GUID] = 6
  End Enum

End Namespace