<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SubmissionMessage.aspx.vb"
    Inherits="SubmissionMessage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    

    <script language="javascript" type="text/javascript">
// <!CDATA[
        function window_onload() {

            document.getElementById('lblSubmissionsMessage_1').innerHTML = window.parent.document.getElementById("frmMain").hdnSubmissionMessage_1.value + "&nbsp;";
            document.getElementById('lblSubmissionsMessage_2').innerHTML = window.parent.document.getElementById("frmMain").hdnSubmissionMessage_2.value;
            document.getElementById('lblSubmissionsMessage_3').innerHTML = "&nbsp;" + window.parent.document.getElementById("frmMain").hdnSubmissionMessage_3.value;
            
            // Resize the fame.
            resizeFrame();

            document.getElementById('spnClickHere').focus();
        }

        function resizeFrame() {
          try {
            window.resizeTo(document.getElementById("frmMessage").offsetParent.scrollWidth, document.getElementById("frmMessage").offsetParent.scrollHeight);
            window.parent.resizeToFit(document.getElementById("frmMessage").offsetParent.scrollWidth, document.getElementById("frmMessage").offsetParent.scrollHeight);
          }
          catch (e) { }
        }

        function doLabelClick() {
            try {
                var sFollowOnForms = window.parent.document.getElementById("frmMain").hdnFollowOnForms.value;
                
                if (sFollowOnForms.length > 0) {
                    window.parent.launchFollowOnForms(sFollowOnForms);
                }
                else {
                  closeMe();
                }
            }
            catch (e) { };
          }

          function closeMe() {
            try {
              window.parent.close();

              document.getElementById('lblSubmissionsMessage_1').innerHTML = 'For your security please close your browser';
              document.getElementById('lblSubmissionsMessage_2').innerHTML = '';
              document.getElementById('lblSubmissionsMessage_3').innerHTML = '';

              resizeFrame();
            } 
             catch (e) { alert("For your security please close your browser"); }
          }
// ]]>

    </script>
</head>

<body onload="return window_onload()" scroll="auto" style="overflow:hidden; padding: 0px; margin: 0px; border: 0px">
    <form id="frmMessage" runat="server">
    <table border="0" cellspacing="0" cellpadding="0" style="top: 0px; left: 0px; width: 100%;
        height: 100%; position: relative; text-align: center; font-size: 10pt; color: black;
        font-family: Verdana; border: black 1px solid;" bgcolor="White">
        <tr style="background-color: <%=ColourThemeHex()%>;">
            <td colspan="5" height="10">
            </td>
        </tr>
        <tr style="height: 40px">
            <td width="10" style="background-color: <%=ColourThemeHex()%>;">
                &nbsp;&nbsp;
            </td>
            <td width="40" valign="top">
                <img src="themes/<%=ColourThemeFolder()%>/CrnrTop.gif" alt="" width="40" height="40" />
            </td>
            <td rowspan="3" style="background-color: White">
                <br />
                <%--NB. Keep all tags all on the same line, otherwise the an extra whitespace appear when viewed in IE8 --%>
                <asp:Label ID="lblSubmissionsMessage_1" runat="server" ForeColor="#333366" Text=""></asp:Label><span id="spnClickHere" style="color:#333366;" tabindex="1" onclick="doLabelClick();" onmouseover="try{this.style.color='#ff9608'}catch(e){}" onmouseout="try{this.style.color='#333366';}catch(e){}" onfocus="try{this.style.color='#ff9608';}catch(e){}" onblur="try{this.style.color='#333366';}catch(e){}" onkeypress="try{if(window.event.keyCode == 32){spnClickHere.click()};}catch(e){}"><asp:Label ID="lblSubmissionsMessage_2" runat="server" Text="" Font-Underline="true" Style="cursor: pointer;"></asp:Label></span><asp:Label ID="lblSubmissionsMessage_3" runat="server" ForeColor="#333366" Text=""></asp:Label><br />
                <br />
            </td>
            <td width="40" valign="top">
                <img src="themes/<%=ColourThemeFolder()%>/RCrnrTop.gif" alt="" width="40" height="40" />
            </td>
            <td width="10" style="background-color: <%=ColourThemeHex()%>;">
                &nbsp;&nbsp;
            </td>
        </tr>
        <tr>
            <td width="10" style="background-color: <%=ColourThemeHex()%>;">
            </td>
            <td>
            </td>
            <td>
            </td>
            <td width="10" style="background-color: <%=ColourThemeHex()%>;">
            </td>
        </tr>
        <%--NB. Keep <TD><IMG></TD> tags all on the same line, otherwise the images do not fully align to bottom--%>
        <tr style="height: 40px">
            <td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
            <td width="40" valign="bottom"><img src="themes/<%=ColourThemeFolder()%>/CrnrBot.gif" width="40" height="40" alt="" /></td>
            <td width="40" valign="bottom"><img src="themes/<%=ColourThemeFolder()%>/RCrnrBot.gif" width="40" height="40" alt="" /></td>
            <td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
        </tr>
        <tr bgcolor="<%=ColourThemeHex()%>">
            <td colspan="5" height="10">
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
