Option Strict On
Option Explicit On

Imports ADODB

Public Class clsOutlookLinks
	Private mclsData As New clsDataAccess

	Public Function GetOutlookLinks(ByRef blnOnlyCountFilterMatch As Boolean, Optional ByRef dtStartDate As Date = #12:00:00 AM#, Optional ByRef dtEndDate As Date = #12:00:00 AM#) As Recordset

		Dim pstrSQL As String

		pstrSQL = "SELECT ASRSysTables.TableName, ASRSysOutlookLinks.Title, ASRSysOutlookEvents.Subject, ASRSysOutlookfolders.Name as 'FolderName', ASRSysOutlookEvents.Folder, ASRSysOutlookEvents.StartDate, ASRSysOutlookEvents.EndDate, ASRSysOutlookEvents.RefreshDate, " & vbCrLf & "case (ASRSysOutlookEvents.Refresh | ASRSysOutlookEvents.Deleted) when 0 then " & vbCrLf & "case isnull(ASRSysOutlookEvents.ErrorMessage,'') when '' then " & vbCrLf & "'Successful' else 'Failed' end else 'Pending' end as 'Status', " & vbCrLf & "isnull(ASRSysOutlookEvents.ErrorMessage,'') " & vbCrLf & "FROM ASRSysOutlookEvents " & vbCrLf & "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysOutlookEvents.TableID " & vbCrLf & "JOIN ASRSysOutlookLinks ON ASRSysOutlookLinks.LinkID = ASRSysOutlookEvents.LinkID " & vbCrLf & "JOIN ASRSysOutlookfolders ON ASRSysOutlookfolders.FolderID = ASRSysOutlookEvents.FolderID " & "WHERE DATEDIFF(d,ASRSysOutlookEvents.startdate, GETDATE()) >= 0 " & "AND (DATEDIFF(d, ASRSysOutlookEvents.enddate, GETDATE()) <= 0 OR ASRSysOutlookEvents.enddate IS NULL) " & "ORDER BY ASRSysOutlookEvents.Subject"

		GetOutlookLinks = mclsData.OpenRecordset(pstrSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

	End Function

	Public WriteOnly Property Connection() As Connection
		Set(ByVal Value As Connection)
			gADOCon = Value
		End Set
	End Property

	Public WriteOnly Property Username() As String
		Set(ByVal Value As String)

			' Username passed in from the asp page
			gsUsername = Value

		End Set
	End Property
End Class