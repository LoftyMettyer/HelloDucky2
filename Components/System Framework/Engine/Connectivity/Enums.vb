Namespace Connectivity

  Public Enum ConnectionType
    WindowsAuthentication = 0
    SqlAuthentication = 1
    OracleAuthentication = 2
    SomethingElse = 3
  End Enum

  Public Enum Provider
    Sql = 0
    SqlContext = 1
    Oracle = 2
    Access = 3
    FoxPro = 4
    SqlExpress = 5
  End Enum

  Public Enum MetadataProvider
    LegacyDao = 0
    PhoenixStoredProcs = 1
  End Enum

  Public Enum DbType
    [Integer] = 0
    [Numeric] = 1
    [String] = 2
    [Date] = 3
    [Boolean] = 4
    [Variant] = 5
    Guid = 6
  End Enum

End Namespace