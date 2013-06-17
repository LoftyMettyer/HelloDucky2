<%@ Application Language="VB" %>

<script runat="server">

	Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs on application startup
		Dim fileInfo As System.IO.FileInfo
        
		Try
			Dim dirInfo As New System.IO.DirectoryInfo(Server.MapPath("pictures"))
			Dim files As System.IO.FileInfo() = dirInfo.GetFiles()
            
			For Each fileInfo In files
				Try
					fileInfo.Delete()
				Catch ex As Exception
				End Try
			Next
		Catch ex As Exception
		End Try
	End Sub
    
	Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs on application shutdown
	End Sub
         
	Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs when an unhandled error occurs
    End Sub
  
	Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs when a new session is started
		Dim fileInfo As System.IO.FileInfo

		Session("TimeoutSecs") = Session.Timeout * 60

		Try
			Dim dirInfo As New System.IO.DirectoryInfo(Server.MapPath("pictures"))
			Dim files As System.IO.FileInfo() = dirInfo.GetFiles()
            
			For Each fileInfo In files
				Try
					If DateDiff(DateInterval.Minute, fileInfo.CreationTime, Now) > Session.Timeout Then
						fileInfo.Delete()
					End If
				Catch ex As Exception
				End Try
			Next
		Catch ex As Exception
		End Try
	End Sub
  
	Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs when a session ends. 
	End Sub
       
</script>