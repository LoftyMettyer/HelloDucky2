
Imports SystemFramework.Enums.Errors

Namespace Structures

  Public Structure [Error]
    Public Id As Guid
    Public Section As Section
    Public ObjectName As String
    Public Severity As Severity
    Public Message As String
    Public Detail As String
    Public DateTime As Date
    Public User As String
    Public ErrorNumber As Long
    Public ErrorArticleId As Long
  End Structure

End Namespace
