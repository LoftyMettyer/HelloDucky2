<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobileLogin.aspx.vb" Inherits="MobileLogin" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>OpenHR Mobile Login</title>
	<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="Images/Adv_hr&payroll.gif" />  

  <%--<meta name="apple-mobile-web-app-capable" content="yes" />--%>

  <script type="text/javascript">
      // <!CDATA[

      function window_onload() {
        if (typeof (getCookie("Login")) != "object") {
          frmLogin.txtUserName.value = getCookie("Login");
        }


        frmLogin.txtUserName.focus();

      }

      function getCookie(name) {
        var dc = document.cookie;
        var prefix = name + "=";
        var begin = dc.indexOf("; " + prefix);
        if (begin == -1) {
          begin = dc.indexOf(prefix);
          if (begin != 0) return null;
        } else
          begin += 2;
        var end = document.cookie.indexOf(";", begin);
        if (end == -1)
          end = dc.length;
        return unescape(dc.substring(begin + prefix.length, end));
      }

      function closeMsgBox() {
        pnlGreyOut.style.visibility = "hidden";
        pnlMsgBox.style.visibility = "hidden";
      }

      //window.scrollTo(0, 1);

      // ]]>
    </script>

</head>
<body onload="return window_onload()" style="margin:0px;overflow:hidden">
    <form id="frmLogin" runat="server" defaultbutton="btnLogin">

        <div id="pnlContainer" runat="server" style="overflow:hidden;background-color:Red">

          <div id="pnlHeader" runat="server" style="position:absolute;overflow:hidden;left:0px;top:0px;z-index:0;width:100%;height:57px">
          </div>
          <div id="ScrollerFrame" runat="server" style="position:absolute;left:0px;top:57px;z-index:0;height:400px;width:100%">
            <div id="pnlBody" runat="server" style="height:100%;z-index:0; margin: 15px">      
              <table style="position:absolute;width:100%;height:100%" >
                <tr id="space1" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <td colspan="2"><label style="margin:15px" id="lblWelcome" runat="server">lblWelcome</label></td>
                </tr>
                <tr id="space2" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <td style="width:40%" ><label style="margin:15px" id="lblUserName" runat="server">lblUserName</label></td>
                  <td ><input type="text" id="txtUserName" runat="server"/></td>
                </tr>
                <tr id="space3" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <td style="width:50%"><label style="margin:15px" id="lblPassword" runat="server">lblPassword</label></td>
                  <td><input id="txtPassword" runat="server" type="password"/></td>
                </tr>
                <tr id="space4" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <td style="width:50%"><label style="margin:15px" id="lblRememberPwd" runat="server">lblRememberPwd</label></td>
                  <td><input id="chkRememberPwd" type="checkbox" runat="server" /></td>
                </tr>
                <tr id="space5" style="width: 100%;height:80%"><td></td></tr>
               </table>
            </div>
          </div>
          
          <div  style="text-align:center; position:absolute;top:357px;width:100%;z-index-0">
            <p style="font-family: Verdana; font-size: 10px; z-index: 2; color: #333366;">Copyright © Advanced Business Software and Solutions Limited 2011</p>
          </div>

          <div id="pnlFooter" runat="server" style="position:fixed;overflow:hidden;left:0px;bottom:0px;z-index-0;width:100%;height:60px">
            <table id="tblFooter" runat="server" style="height:100%;width:100%">
              <tr style="height:40px">
                <td style="width:33%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnLogin"  runat="server" /></td>
                <td style="width:33%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnForgotPwd" runat="server"/></td>
                <td style="width:33%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnRegister" runat="server" /></td>
              </tr>
              <tr style="height:17px">
                <td style="width:33%;text-align:center;overflow:hidden"><label runat="server" id="btnLogin_label"></label></td>
                <td style="width:33%;text-align:center;overflow:hidden"><label runat="server" id="btnForgotPwd_label"></label></td>
                <td style="width:33%;text-align:center;overflow:hidden"><label runat="server" id="btnRegister_label"></label></td>
              </tr>
            </table>
          </div>        
        
        </div>
        

        <div id="pnlGreyOut" runat="server" style="position: absolute;visibility: hidden;width: 100%;height: 100%;filter:alpha(opacity=50);
                              -moz-opacity:0.5;opacity: 0.5;background-color: #222;margin:0px;z-index:1">
        </div>
          <div id="pnlMsgBox" runat="server" style="background-color: #002248;position: absolute;visibility: hidden;width: 300px;height:100px;
                                            text-align:center;position:absolute;left:50%;top:50%;margin-left:-150px;margin-top:-50px;z-index:2;
                                            border-radius:10px;border:2px solid #001648;vertical-align:middle;opacity:0.8">
            <br/>&nbsp;&nbsp;
            <label id="lblMsgBox" runat="server" style="font-family: Verdana;font-size:large;color:white"></label>
            <br/>&nbsp;&nbsp;<br/>
            <input type="button" value="OK" style="width:100px;height:30px;background-color: ButtonHighlight" onclick="closeMsgBox();"/>
          </div>  


    </form>
</body>
</html>

