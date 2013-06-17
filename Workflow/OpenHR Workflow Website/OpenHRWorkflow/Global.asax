<%@ Application Language="VB" %>

<script runat="server">
        
    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application startup
        Try
            Dim path = Server.MapPath("~/Pictures")
            
            'Delete the picture files created by default.aspx
            For Each file In System.IO.Directory.GetFiles(path)
                Try
                    System.IO.File.Delete(file)
                Catch ex As Exception
                End Try
            Next
            
            'Delete the picture files created by rest of system
            For Each folder In System.IO.Directory.GetDirectories(path)
                Try
                    System.IO.Directory.Delete(folder, True)
                Catch ex As Exception
                End Try
            Next
            
        Catch ex As Exception
        End Try
    End Sub

	Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application shutdown       
    End Sub

    'TODO catch errors and show message page or setup in config.web
	Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs when an unhandled error occurs
    End Sub
  
	Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a new session is started
    End Sub
  
    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a session ends.
        
        Try
            Dim path = Server.MapPath("~/Pictures")
            
            'Delete the picture files created by default.aspx
            For Each file In System.IO.Directory.GetFiles(path)
                Try
                    System.IO.File.Delete(file)
                Catch ex As Exception
                End Try
            Next
            
        Catch ex As Exception
        End Try
        
    End Sub
    
</script>