<%@ Application Language="VB" %>

<script runat="server">

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application startup
    End Sub
    
	Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs on application shutdown
	End Sub
         
	Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs when an unhandled error occurs
    End Sub
  
	Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs when a new session is started
        
        Session("TimeoutSecs") = Session.Timeout * 60

	End Sub
  
	Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
		' Code that runs when a session ends. 
	End Sub
          
</script>