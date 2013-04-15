Option Strict Off
Option Explicit On
Public Class clsOutlookLinks
  Private mclsData As New clsDataAccess

  Public Function GetOutlookLinks(ByRef blnOnlyCountFilterMatch As Boolean, Optional ByRef dtStartDate As Date = #12:00:00 AM#, Optional ByRef dtEndDate As Date = #12:00:00 AM#) As Object

    Dim pstrSQL As String
    Dim blnWhere As Boolean
    Dim strWhere As String

    pstrSQL = "SELECT ASRSysTables.TableName, " & vbCrLf & "ASRSysOutlookLinks.Title, " & vbCrLf & "ASRSysOutlookEvents.Subject, " & vbCrLf & "ASRSysOutlookfolders.Name as 'FolderName', " & vbCrLf & "ASRSysOutlookEvents.Folder, " & vbCrLf & "ASRSysOutlookEvents.StartDate, " & vbCrLf & "ASRSysOutlookEvents.EndDate, " & vbCrLf & "ASRSysOutlookEvents.RefreshDate, " & vbCrLf & "case (ASRSysOutlookEvents.Refresh | ASRSysOutlookEvents.Deleted) when 0 then " & vbCrLf & "case isnull(ASRSysOutlookEvents.ErrorMessage,'') when '' then " & vbCrLf & "'Successful' else 'Failed' end else 'Pending' end as 'Status', " & vbCrLf & "isnull(ASRSysOutlookEvents.ErrorMessage,'') " & vbCrLf & "FROM ASRSysOutlookEvents " & vbCrLf & "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysOutlookEvents.TableID " & vbCrLf & "JOIN ASRSysOutlookLinks ON ASRSysOutlookLinks.LinkID = ASRSysOutlookEvents.LinkID " & vbCrLf & "JOIN ASRSysOutlookfolders ON ASRSysOutlookfolders.FolderID = ASRSysOutlookEvents.FolderID " & "WHERE DATEDIFF(d,ASRSysOutlookEvents.startdate, GETDATE()) >= 0 " & "AND (DATEDIFF(d, ASRSysOutlookEvents.enddate, GETDATE()) <= 0 OR ASRSysOutlookEvents.enddate IS NULL) " & "ORDER BY ASRSysOutlookEvents.Subject"

    GetOutlookLinks = mclsData.OpenPersistentRecordset(pstrSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)


  End Function

  Public WriteOnly Property Connection() As Object
    Set(ByVal Value As Object)

      ' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
      If ASRDEVELOPMENT Then
        gADOCon = New ADODB.Connection
        'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gADOCon.Open(Value)
      Else
        gADOCon = Value
      End If

    End Set
  End Property

  Public WriteOnly Property Username() As String
    Set(ByVal Value As String)

      ' Username passed in from the asp page
      gsUsername = Value

    End Set
  End Property
End Class