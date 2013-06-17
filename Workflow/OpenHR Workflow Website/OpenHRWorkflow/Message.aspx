<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Message.aspx.vb" Inherits="Message" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title><%=Session("titleVersion")%></title>
    <link rel="shortcut icon" href="pictures/logo.ico"/>

	<script for="window" event="onload" type="text/javascript">
	<!--
	    var iMINWIDTH = 400;
    
		// Resize the browser.
	    try
	    {
	        window.parent.resizeTo(10, 10);
	    }
	    catch(e) {}

//		iResizeByHeight = frmMessage.offsetParent.scrollHeight - frmMessage.offsetParent.clientHeight;
	    iResizeByHeight = window.document.documentElement.scrollHeight - window.document.documentElement.clientHeight;
		
	    if (document.getElementById("frmMessage").offsetParent.offsetHeight + iResizeByHeight > screen.availHeight) 
	    {
	        try
	        {
	            window.parent.moveTo((screen.width - frmMessage.offsetParent.offsetWidth) / 2, 0);
	            window.parent.resizeTo(document.getElementById("frmMessage").offsetParent.offsetWidth, screen.availHeight);
	        }
	        catch(e) {}
	    }
	    else 
	    {
	        try
	        {
	            window.parent.moveTo((screen.width - document.getElementById("frmMessage").offsetParent.offsetWidth) / 2, (screen.availHeight - (document.getElementById("frmMessage").offsetParent.offsetHeight + iResizeByHeight)) / 3);
	            window.parent.resizeBy(0, iResizeByHeight);
	        }
	        catch(e) {}
	    }

	    if(document.getElementById("frmMessage").offsetParent.scrollWidth < iMINWIDTH)
	    {
	        iResizeByWidth = iMINWIDTH - document.getElementById("frmMessage").offsetParent.clientWidth;
	    }
	    else
	    {
	        //iResizeByWidth = frmMessage.offsetParent.scrollWidth - frmMessage.offsetParent.clientWidth;
	        iResizeByWidth = window.document.documentElement.scrollWidth - window.document.documentElement.clientWidth;
	    }

		//alert(iResizeByWidth);

	    if (document.getElementById("frmMessage").offsetParent.offsetWidth + iResizeByWidth > screen.width) 
	    {
	        try
	        {
	            window.parent.moveTo(0, (screen.availHeight - document.getElementById("frmMessage").offsetParent.offsetHeight) / 3);
	            window.parent.resizeTo(screen.width, document.getElementById("frmMessage").offsetParent.offsetHeight);
	        }
	        catch(e) {}
	    }
	    else 
	    {
	        try
	        {
	            window.parent.moveTo((screen.width - (document.getElementById("frmMessage").offsetParent.offsetWidth + iResizeByWidth)) / 2, (screen.availHeight - document.getElementById("frmMessage").offsetParent.offsetHeight) / 3);
	            window.parent.resizeBy(iResizeByWidth, 0);
	        }
	        catch(e) {}
	    }
		
		// Redo the height calc (it does need to be done again).		
		//iResizeByHeight = frmMessage.offsetParent.scrollHeight - frmMessage.offsetParent.clientHeight;
	    iResizeByHeight = window.document.documentElement.scrollHeight - window.document.documentElement.clientHeight;

	    if (document.getElementById("frmMessage").offsetParent.offsetHeight + iResizeByHeight > screen.availHeight) 
	    {
	        try
	        {
	            window.parent.moveTo((screen.width - document.getElementById("frmMessage").offsetParent.offsetWidth) / 2, 0);
	            window.parent.resizeTo(document.getElementById("frmMessage").offsetParent.offsetWidth, screen.availHeight);
	        }
	        catch(e) {}
	    }
	    else 
	    {
	        try
	        {
	            window.parent.moveTo((screen.width - document.getElementById("frmMessage").offsetParent.offsetWidth) / 2, (screen.availHeight - (document.getElementById("frmMessage").offsetParent.offsetHeight + iResizeByHeight)) / 3);
	            window.parent.resizeBy(0, iResizeByHeight);
	        }
	        catch(e) {}
	    }

	    document.getElementById('spnClickHere').focus();
	-->
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
					<font face='Verdana' style="color:#333366; FONT-SIZE:<%=MessageFontSize()%>pt">
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
					<font face='Verdana' style="color:#333366; FONT-SIZE:<%=MessageFontSize()%>pt">
						<asp:Label ID="lblPrompt1" 
								runat="server" Text="Click">
							</asp:Label> 
						<span id="spnClickHere" name="spnClickHere" onclick="closeMe();" tabindex="1"
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
</body>
</html>
