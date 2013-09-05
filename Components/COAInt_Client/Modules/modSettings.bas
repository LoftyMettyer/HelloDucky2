Attribute VB_Name = "modSettings"
'Option Explicit
'
'Private mcolSystemSettings As Collection
'
'
'Private Function AddSystemSetting(strSection As String, strKey As String, strValue As String)
'
'  Dim col As Collection
'
'  On Local Error GoTo LocalErr
'
'  Set col = mcolSystemSettings(strSection)
'  col.Add strValue, strKey
'
'Exit Function
'
'LocalErr:
'  Set col = New Collection
'  mcolSystemSettings.Add col, strSection
'  Resume Next
'
'End Function
'
'
'Private Function GetSystemSetting(strSection As String, strKey As String, strDefault As String) As String
'
'  Dim col As Collection
'
'  On Local Error GoTo LocalErr
'
'  Set col = mcolSystemSettings(strSection)
'  GetSystemSetting = col(strKey)
'
'Exit Function
'
'LocalErr:
'  GetSystemSetting = strDefault
'
'End Function


