<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Message.aspx.vb" Inherits="Message" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title><%=Session("titleVersion")%></title>
    <link rel="shortcut icon" href="images/logo.ico"/>

	<script type="text/javascript">

	    function window_onload() {

	        var iDefHeight;
	        var iDefWidth;
	        var iResizeByHeight;
	        var iResizeByWidth;
	        var sControlType;

	        try {
	            iDefHeight = 150;
	            iDefWidth = 400;
	            window.focus();
	            if ((iDefHeight > 0) && (iDefWidth > 0)) {
	                iResizeByHeight = iDefHeight - getWindowHeight();
	                iResizeByWidth = iDefWidth - getWindowWidth();
	                window.parent.moveTo((screen.availWidth - iDefWidth) / 2, (screen.availHeight - iDefHeight) / 3);
	                window.parent.resizeBy(iResizeByWidth, iResizeByHeight);	               
	            }				
	        }
	        catch (e) {}

	        //Fault HRPRO-2121
	        try	{
	            window.resizeBy(0,-1);
	            window.resizeBy(0,1);
	        }
	        catch(e) {}

	        document.getElementById('spnClickHere').focus();

	      }

	      function getWindowWidth() {
	        var myWidth = 0;
	        if (typeof (window.innerWidth) == 'number') {
	          //Non-IE
	          myWidth = window.innerWidth;
	        } else if (document.documentElement && (document.documentElement.clientWidth)) {
	          //IE 6+ in 'standards compliant mode'
	          myWidth = document.documentElement.clientWidth;
	        } else if (document.body && (document.body.clientWidth)) {
	          //IE 4 compatible
	          myWidth = document.body.clientWidth;
	        }
	        return myWidth;
	      }

	      function getWindowHeight() {
	        var myHeight = 0;
	        if (typeof (window.innerHeight) == 'number') {
	          //Non-IE
	          myHeight = window.innerHeight;
	        } else if (document.documentElement && (document.documentElement.clientHeight)) {
	          //IE 6+ in 'standards compliant mode'
	          myHeight = document.documentElement.clientHeight;
	        } else if (document.body && (document.body.clientHeight)) {
	          //IE 4 compatible
	          myHeight = document.body.clientHeight;
	        }
	        return myHeight;
	      }
	</script>

    <script type="text/javascript">
        function closeMe() {
            try {
                window.close();

                document.getElementById('lblPrompt1').innerHTML = "For your security please close your browser.";
                document.getElementById('lblClose').innerHTML = "";
                document.getElementById('lblPrompt2').innerHTML = "";
            }

            catch (e) { alert("For your security please close your browser"); }

        }
    </script>
</head>

<body 
	bgcolor="<%=ColourThemeHex()%>" 
	bottommargin="0" rightmargin="0" leftmargin="0" topmargin="0" 
	scroll=auto 
	style="overflow:auto;">
    
	<form name="frmMessage" id="frmMessage" method="post" runat="server" 
		 style="overflow:visible; 
				left: 0px; width: 100%; 
				position: relative; 
				top: 0px; height: 100%;">

		<table height="100%" width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr bgcolor="<%=ColourThemeHex()%>">
				<td colspan="5" height="10"></td>
			</tr>

			<tr height="40">
				<td width="10" bgcolor="<%=ColourThemeHex()%>">&nbsp;&nbsp;</td>
				<td width="40" valign="top"><img src="themes/<%=ColourThemeFolder()%>/CrnrTop.gif" width="40" height="40" alt="" /></td>
				<td width="100%" bgcolor="White"></td>
				<td width="40" valign="top"><img src="themes/<%=ColourThemeFolder()%>/RCrnrTop.gif" width="40" height="40" alt="" /></td>
				<td width="10" bgcolor="<%=ColourThemeHex()%>">&nbsp;&nbsp;</td>
			</tr>

			<tr height="100%">
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
				<td width="40" bgcolor="White"></td>
				<td align="center" bgcolor="White">
					<font face='Verdana' style="color:#333366; font-size:<%=MessageFontSize()%>pt">
						<%=Session("message")%>
					</font>
				</td>
				<td width="40" bgcolor="White"></td>
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
			</tr>

			<tr bgcolor="<%=ColourThemeHex()%>" height="10">
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
				<td colspan="3" bgcolor="White"></td>
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
			</tr>

			<tr height="100%">
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
				<td width="40" bgcolor="White"></td>
				<td align="center" bgcolor="White">
					<font face='Verdana' style="color:#333366; font-size:<%=MessageFontSize()%>pt">
						<asp:Label ID="lblPrompt1" 
								runat="server" Text="Click">
							</asp:Label> 
						<span id="spnClickHere" onclick="closeMe();" tabindex="1"
								onmouseover="try{this.style.color='#ff9608';}catch(e){}" 
								onmouseout="try{this.style.color='#333366';}catch(e){}" 
		            onfocus="try{this.style.color='#ff9608';}catch(e){}" 
		            onblur="try{this.style.color='#333366';}catch(e){}"
		            onkeypress="try{if(window.event.keyCode == 32){spnClickHere.click();}}catch(e){}"
		            >
							<asp:Label ID="lblClose" 
								runat="server" Text="here" 
								Font-Underline="true" 
								style="cursor: pointer;">
							</asp:Label>
						</span>
						<asp:Label ID="lblPrompt2" 
								runat="server" Text="to close this form.">
							</asp:Label> 
					</font>
				</td>
				<td width="40" bgcolor="White"></td>
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
			</tr>

			<tr height=40>
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
				<td width="40" valign="top"><img src="themes/<%=ColourThemeFolder()%>/CrnrBot.gif" width="40" height="40" alt="" /></td>
				<td width="100%" bgcolor="White"></td>
				<td width="40" valign="top"><img src="themes/<%=ColourThemeFolder()%>/RCrnrBot.gif" width="40" height="40" alt="" /></td>
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
			</tr>

			<tr bgcolor="<%=ColourThemeHex()%>">
				<td colspan="5" height="10"></td>
			</tr>
		</table>
	</form>
    <script type="text/javascript">
      window_onload();
    </script>
</body>
</html>
