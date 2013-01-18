<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
    <meta http-equiv="refresh" content="30;URL=poll">
    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css">

    <script type="text/javascript">
        function poll_window_onload() {
            var sMessage = new String("");
            var controlCollection = frmMessages.elements;
            if (controlCollection != null) {
                for (var i = 0; i < controlCollection.length; i++) {
                    if (sMessage.length > 0) {
                        sMessage = sMessage + "\n\n";
                    }
                    sMessage = sMessage + controlCollection.item(i).value;
                }
                if (sMessage.length > 0) {
                    var frmPollMsg = OpenHR.getForm("pollmessageframe", "frmSetMessage");
                    frmPollMsg.txtMessage.value = sMessage;
                    pollmessage_refreshMessage();
                }
            }
        }
    </script>

    <form action="poll" method="post" id="frmHit" name="frmHit">
        <input type="hidden" id="txtDummy" name="txtDummy" value="0">
    </form>

    <form id="frmMessages" name="frmMessages">
        <%
            Dim cmdHit = CreateObject("ADODB.Command")
            cmdHit.CommandText = "sp_ASRIntPoll"
            cmdHit.CommandType = 4 ' Stored Procedure
            cmdHit.ActiveConnection = Session("databaseConnection")

            Err.Clear()
            Dim rstMessages = cmdHit.Execute

            If (Err.Number = 0) Then
                Dim iloop = 1
                Do While Not rstMessages.EOF
                    ' Response.Write("<INPUT type='hidden' id=txtMessage_" & iLoop & " name=txtMessage_" & iLoop & " value=""" & Replace(rstMessages.Fields(0).Value, """", "&quot;") & """>" & vbCrLf)
%>		
	<INPUT type='hidden' 
		id=txtMessage_<%=iLoop%> 
		name=txtMessage_<%=iLoop%> 
		value="<%=replace(rstMessages.Fields(0).Value, """", "&quot;")%>">
<%                    
                    rstMessages.MoveNext()
	
                    iloop = iloop + 1
                Loop

                ' Release the ADO recordset object.
                rstMessages.close()
                ' rstMessages = Nothing
            End If
	
            ' Release the ADO command object.
            cmdHit = Nothing
        %>
    </form>
    
    <script type="text/javascript">poll_window_onload();</script>
