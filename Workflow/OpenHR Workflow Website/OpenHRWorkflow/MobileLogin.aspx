﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobileLogin.aspx.vb" Inherits="MobileLogin" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
    <head id="Head1" runat="server">
        <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
        <%--<meta name="apple-mobile-web-app-capable" content="yes" />--%>
        <link rel="apple-touch-icon" href="/Images/Adv_hr&payroll.gif" />
        <link href="CSS/mobile.css" rel="stylesheet" type="text/css" />
        <title>OpenHR Mobile</title>

        <script type="text/javascript">
      // <!CDATA[

            function window_onload() {
                document.getElementById("txtUserName").focus();
            }

            function submitCheck() {

                var header = 'Login Failed';
                
                if (document.getElementById('txtUserName').value.length === 0) {
                    showMsgBox(header, 'Username is required.');
                    return false;
                }
                if (document.getElementById('txtPassword').value.length === 0) {
                    showMsgBox(header, 'Password is required.');
                    return false;
                }
                return true;
            }

            function showMsgBox(header, message) {
                document.getElementById('lblMsgHeader').innerHTML = header;
                document.getElementById('lblMsgBox').innerHTML = message;
                document.getElementById('pnlGreyOut').style.visibility = "visible";
                document.getElementById('pnlMsgBox').style.visibility = "visible";
            }

            function closeMsgBox() {
                document.getElementById('pnlGreyOut').style.visibility = "hidden";
                document.getElementById('pnlMsgBox').style.visibility = "hidden";

                var redirectTo = document.getElementById('hdnRedirectTo').value;

                if (redirectTo.length > 0) {
                    window.location.replace(redirectTo);
                }
            }
        // ]]>
    </script>

    </head>
    <body onload="return window_onload()">
        <form id="frmLogin" runat="server" defaultbutton="btnLogin">

            <div id="pnlContainer" runat="server" style="overflow: hidden;">

                <div id="pnlHeader" runat="server" style="position: absolute; overflow: hidden; left: 0px; top: 0px; z-index: 0; width: 100%; height: 57px"/>
                
                <div id="ScrollerFrame" runat="server" style="position: fixed; left: 0px; top: 57px; z-index: 0; bottom: 60px; width: 100%">
                    
                    <div id="pnlBody" runat="server" style="height: 100%; z-index: 0;">
                              
                        <label id="lblWelcome" style="display: block" runat="server">lblWelcome</label>

                        <table style="width: 100%">
                            <tr>
                                <td><label id="lblUserName" runat="server">lblUserName</label></td>
                                <td ><input type="text" id="txtUserName" runat="server"/></td>
                            </tr>
                            <tr>
                                <td><label id="lblPassword" runat="server">lblPassword</label></td>
                                <td><input id="txtPassword" runat="server" type="password"/></td>
                            </tr>
                            <tr>
                                <td><label  id="lblRememberPwd" runat="server">lblRememberPwd</label></td>
                                <td><input id="chkRememberPwd" type="checkbox" runat="server" /></td>
                            </tr>
                        </table>
                    </div>
                </div>
          
                <div  style="text-align: center; position: fixed; bottom: 75px; width: 100%; z-index: 0">
                    <p style="font-family: Verdana; font-size: 10px; z-index: 2; color: #333366;">Copyright © Advanced Business Software and Solutions Ltd 2012</p>
                </div>

                <div id="pnlFooter" runat="server" style="position: fixed; overflow: hidden; left: 0px; bottom: 0px; z-index: 0; width: 100%; height: 60px">
                    <table id="tblFooter" runat="server" style="height: 100%; width: 100%">
                        <tr style="height: 40px">
                            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnLogin"  runat="server" OnClientClick="return submitCheck();"/></td>
                            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnForgotPwd" runat="server"/></td>
                            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnRegister" runat="server" /></td>
                        </tr>
                        <tr style="height: 17px">
                            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnLogin_label"></label></td>
                            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnForgotPwd_label"></label></td>
                            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnRegister_label"></label></td>
                        </tr>
                    </table>
                </div>        
 
            </div>
            
            <div id="pnlGreyOut" runat="server" style="position: absolute; visibility: hidden; width: 100%; height: 100%; filter: alpha(opacity=50); -moz-opacity: 0.5; opacity: 0.5; background-color: #222; margin: 0px; z-index: 1"/>
            
            <div id="pnlMsgBox" runat="server" style="visibility: hidden; z-index: 2; position: absolute; width: 100%; top: 30%">
                <div id="inner" style="background-color: #002248; border: 2px solid gainsboro; width: 300px; margin: 0px auto; text-align: center; border-radius: 10px; padding: 10px;">
                    <label id="lblMsgHeader" runat="server" style="font-family: Verdana; font-weight: bold; font-size: large; color: white"></label>
                    <br/>
                    <br/>
                    <label id="lblMsgBox" runat="server" style="font-family: Verdana; font-size: large; color: white"></label>
                    <br/>
                    <br/>
                    <input type="hidden" id="hdnRedirectTo" runat="server"/>
                    <input type="button" value="OK" style="width: 100px; height: 30px; background-color: ButtonHighlight" onclick="closeMsgBox(); "/>
                </div>
            </div>
              
        </form>
    </body>
</html>