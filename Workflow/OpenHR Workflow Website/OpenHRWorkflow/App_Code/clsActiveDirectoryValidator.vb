Imports System
Imports System.DirectoryServices

Public Class ActiveDirectoryValidator
  Private _path As String

  Public Sub New(path As String)
    _path = path
  End Sub

  Public Function IsAuthenticated(domainName As String, userName As String, password As String) As Boolean
    Dim domainAndUsername As String = domainName & "\" & userName
    Dim entry As New DirectoryEntry(_path, domainAndUsername, password)
    Try
      ' Bind to the native AdsObject to force authentication.
      Dim obj As [Object] = entry.NativeObject
      Dim search As New DirectorySearcher(entry)
      search.Filter = "(SAMAccountName=" & userName & ")"
      search.PropertiesToLoad.Add("cn")
      Dim result As SearchResult = search.FindOne()
      If result Is Nothing Then
        Return False
      End If
      ' Update the new path to the user in the directory
      _path = result.Path
      Dim _filterAttribute As String = CStr(result.Properties("cn")(0))
    Catch ex As Exception
      Throw New Exception("Login Error: " & ex.Message)
    End Try
    Return True
  End Function
End Class
