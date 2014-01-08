Option Strict On
Option Explicit On

Imports HR.Intranet.Server.BaseClasses

Public Class clsOutlookLinks
	Inherits BaseForDMI

	Public Function GetOutlookLinks() As DataTable

		Dim strSQL As String

		strSQL = "SELECT ASRSysTables.TableName, ASRSysOutlookLinks.Title, ASRSysOutlookEvents.Subject, ASRSysOutlookfolders.Name as 'FolderName', ASRSysOutlookEvents.Folder, ASRSysOutlookEvents.StartDate, ASRSysOutlookEvents.EndDate, ASRSysOutlookEvents.RefreshDate, " & vbCrLf & "case (ASRSysOutlookEvents.Refresh | ASRSysOutlookEvents.Deleted) when 0 then case isnull(ASRSysOutlookEvents.ErrorMessage,'') when '' then 'Successful' else 'Failed' end else 'Pending' end as 'Status', " & vbCrLf & "isnull(ASRSysOutlookEvents.ErrorMessage,'') " & vbCrLf & "FROM ASRSysOutlookEvents " & vbCrLf & "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysOutlookEvents.TableID " & vbCrLf & "JOIN ASRSysOutlookLinks ON ASRSysOutlookLinks.LinkID = ASRSysOutlookEvents.LinkID " & vbCrLf & "JOIN ASRSysOutlookfolders ON ASRSysOutlookfolders.FolderID = ASRSysOutlookEvents.FolderID " & "WHERE DATEDIFF(d,ASRSysOutlookEvents.startdate, GETDATE()) >= 0 " & "AND (DATEDIFF(d, ASRSysOutlookEvents.enddate, GETDATE()) <= 0 OR ASRSysOutlookEvents.enddate IS NULL) ORDER BY ASRSysOutlookEvents.Subject"
		Return DB.GetDataTable(strSQL, CommandType.Text)
	End Function

End Class