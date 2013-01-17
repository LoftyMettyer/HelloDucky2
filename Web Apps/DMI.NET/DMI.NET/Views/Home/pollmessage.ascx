<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<html>
<head>
    <title></title>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="OpenHR.css">

    <script type="text/javascript">

        function pollmessage_window_onload() {            
            var sMessage;
            var frmGetMessage = document.getElementById("frmGetMessage");
        
            sMessage = new String(frmGetMessage.txtMessage.value);
            if (sMessage.length > 0) {
                window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMessage);
                frmGetMessage.txtMessage.value = "";
            }
        }

    </script>

    <script type="text/javascript">
    function pollmessage_refreshMessage()
    {
        //frmSetMessage.submit();  
        var frmSetMessage = OpenHR.getForm("pollmessageframe", "frmSetMessage");
        OpenHR.submitForm(frmSetMessage);
    }    
    </script>

</head>

<body bgcolor='<%=session("ConvertedDesktopColour")%>'>
    <form action="pollmessage_submit" method="post" id="frmSetMessage" name="frmSetMessage">
        <input type="hidden" id="txtMessage" name="txtMessage">
    </form>

    <form id="frmGetMessage" name="frmGetMessage">
        <%
            Response.Write("<INPUT type='hidden' id=txtMessage name=txtMessage value=""" & Replace(Session("pollMessage"), """", "&quot;") & """>" & vbCrLf)
        %>
    </form>
</body>

    <script type="text/javascript"> pollmessage_window_onload();</script>
    

</html>
