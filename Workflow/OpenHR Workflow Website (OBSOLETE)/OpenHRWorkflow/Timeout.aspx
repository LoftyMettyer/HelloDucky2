<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Timeout.aspx.vb" Inherits="timeout" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
	<title>Open HR Workflow</title>
    

	<script type="text/javascript">

	  function window_onload() {

	    var iMINWIDTH = 400;

	    // Resize the browser.
	    try {
	      window.parent.resizeTo(iMINWIDTH, iMINWIDTH);
	    }
	    catch (e) { }

	    iResizeByHeight = frmMessage.offsetParent.scrollHeight - frmMessage.offsetParent.clientHeight;
	    if (frmMessage.offsetParent.offsetHeight + iResizeByHeight > screen.availHeight) {
	      try {
	        window.parent.moveTo((screen.width - frmMessage.offsetParent.offsetWidth) / 2, 0);
	        window.parent.resizeTo(frmMessage.offsetParent.offsetWidth, screen.availHeight);
	      }
	      catch (e) { }
	    }
	    else {
	      try {
	        window.parent.moveTo((screen.width - frmMessage.offsetParent.offsetWidth) / 2, (screen.availHeight - (frmMessage.offsetParent.offsetHeight + iResizeByHeight)) / 3);
	        window.parent.resizeBy(0, iResizeByHeight);
	      }
	      catch (e) { }
	    }

	    if (frmMessage.offsetParent.scrollWidth < iMINWIDTH) {
	      iResizeByWidth = iMINWIDTH - frmMessage.offsetParent.clientWidth;
	    }
	    else {
	      iResizeByWidth = frmMessage.offsetParent.scrollWidth - frmMessage.offsetParent.clientWidth;
	    }
	    if (frmMessage.offsetParent.offsetWidth + iResizeByWidth > screen.width) {
	      try {
	        window.parent.moveTo(0, (screen.availHeight - frmMessage.offsetParent.offsetHeight) / 3);
	        window.parent.resizeTo(screen.width, frmMessage.offsetParent.offsetHeight);
	      }
	      catch (e) { }
	    }
	    else {
	      try {
	        window.parent.moveTo((screen.width - (frmMessage.offsetParent.offsetWidth + iResizeByWidth)) / 2, (screen.availHeight - frmMessage.offsetParent.offsetHeight) / 3);
	        window.parent.resizeBy(iResizeByWidth, 0);
	      }
	      catch (e) { }
	    }

	    // Redo the height calc (it does need to be done again).		
	    iResizeByHeight = frmMessage.offsetParent.scrollHeight - frmMessage.offsetParent.clientHeight;
	    if (frmMessage.offsetParent.offsetHeight + iResizeByHeight > screen.availHeight) {
	      try {
	        window.parent.moveTo((screen.width - frmMessage.offsetParent.offsetWidth) / 2, 0);
	        window.parent.resizeTo(frmMessage.offsetParent.offsetWidth, screen.availHeight);
	      }
	      catch (e) { }
	    }
	    else {
	      try {
	        window.parent.moveTo((screen.width - frmMessage.offsetParent.offsetWidth) / 2, (screen.availHeight - (frmMessage.offsetParent.offsetHeight + iResizeByHeight)) / 3);
	        window.parent.resizeBy(0, iResizeByHeight);
	      }
	      catch (e) { }
	    }
	  }


	  function closeMe() {
	    try {
	      window.close();

	      document.getElementById('Label1').innerHTML = "For your security please close your browser.";
	      document.getElementById('Label2').innerHTML = "";
	      document.getElementById('lblBack').innerHTML = "";
	      document.getElementById('Label3').innerHTML = "";
	      document.getElementById('lblClose').innerHTML = "";
	      document.getElementById('Label4').innerHTML = "";
	    }

	    catch (e) { alert("For your security please close your browser"); }
	  }

	</script>

</head>

  
<body 
	bgcolor="<%=ColourThemeHex()%>" onload="return window_onload()"
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
				<td align="center" bgcolor="White" nowrap>
						<font face='Verdana' style="color:#333366; font-size:<%=MessageFontSize()%>pt">
							<asp:Label ID="Label1" 
								runat="server" Text="Session timeout.">
							</asp:Label> 
						</font>
				</td>
				<td width="40" bgcolor="White"></td>
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
			</tr>

			<tr height="100%">
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
				<td colspan="3" align="center" bgcolor="White">
					<font face='Verdana' style="color:#333366; font-size:<%=MessageFontSize()%>pt">
							<asp:Label ID="Label2" 
								runat="server" Text="Click">
							</asp:Label> 
						<span onclick="try{window.history.back();}catch(e){}" tabindex="1"
								onmouseover="try{this.style.color='#ff9608';}catch(e){}" 
								onmouseout="try{this.style.color='#333366';}catch(e){}" 
		            onfocus="try{this.style.color='#ff9608';}catch(e){}" 
		            onblur="try{this.style.color='#333366';}catch(e){}"
						>
							<asp:Label ID="lblBack" 
								runat="server" Text="here" 
								Font-Underline="true"
								style="cursor: hand;">
							</asp:Label>
						</span>
							<asp:Label ID="Label3" 
								runat="server" Text=" to reload this form, or">
							</asp:Label>
						<span onclick="closeMe();" tabindex="1"
								onmouseover="try{this.style.color='#ff9608';}catch(e){}" 
								onmouseout="try{this.style.color='#333366';}catch(e){}" 
		            onfocus="try{this.style.color='#ff9608';}catch(e){}" 
		            onblur="try{this.style.color='#333366';}catch(e){}"
						>
							<asp:Label ID="lblClose" 
								runat="server" Text="here" 
								Font-Underline="true" 
								style="cursor: hand;">
							</asp:Label>
						</span>
							<asp:Label ID="Label4" 
								runat="server" Text=" to close it.">
							</asp:Label>
					</font>
				</td>
				<td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
			</tr>

			<tr height="40">
				<td width="10" bgcolor="<%=ColourThemeHex()%>">&nbsp;&nbsp;</td>
				<td width="40" valign="top"><img src="themes/<%=ColourThemeFolder()%>/CrnrBot.gif" width="40" height="40"></td>
				<td width="100%" bgcolor="White"></td>
				<td width="40" valign="top"><img src="themes/<%=ColourThemeFolder()%>/RCrnrBot.gif" width="40" height="40"></td>
				<td width="10" bgcolor="<%=ColourThemeHex()%>">&nbsp;&nbsp;</td>
			</tr>

			<tr bgcolor="<%=ColourThemeHex()%>">
				<td colspan="5" height="10"></td>
			</tr>
		</table>
	</form>
</body>
</html>
