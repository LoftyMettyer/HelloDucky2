﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobileForgottenLogin.aspx.vb" Inherits="ForgottenLogin" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
    <head id="Head1" runat="server">
        <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
        <link rel="apple-touch-icon" href="/Images/Adv_hr&payroll.gif" />
        <link href="CSS/mobile.css" rel="stylesheet" type="text/css" />
        <title>OpenHR Mobile</title>
    
        <script  type="text/javascript">
// <!CDATA[

            function window_onload() {
                document.getElementById('txtEmail').focus();
            }

            function submitCheck() {

                var header = 'Request Failed';
                
                if (document.getElementById('txtEmail').value.length === 0) {
                    showMsgBox(header, 'Email address is required.');
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
        <form id="frmForgotLogin" runat="server" defaultbutton="btnSubmit">

            <div id="pnlPage" runat="server" style="overflow: hidden;">
                
                <div id="pnlHeader" runat="server" />
                                    
                <div id="pnlBody" runat="server">
                              
                    <label id="lblWelcome" runat="server">lblWelcome</label>

                    <table class="controlgrid">
                        <tr>
                            <td ><label id="lblEmail" runat="server">lblEmail</label></td>
                            <td><input id="txtEmail" runat="server" /></td>
                        </tr>
                
                    </table>
                </div>
          
                <div id="pnlFooter" runat="server">
                    <table id="tblFooter" runat="server" style="height: 100%; width: 100%">
                        <tr style="height: 40px">
                            <td style="width: 50%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnSubmit" runat="server" OnClientClick="return submitCheck();"/></td>                
                            <td style="width: 50%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnCancel" runat="server" /></td>
                        </tr>
                        <tr style="height: 17px">
                            <td style="width: 50%; text-align: center; overflow: hidden"><label runat="server" id="btnSubmit_label"></label></td>
                            <td style="width: 50%; text-align: center; overflow: hidden"><label runat="server" id="btnCancel_label"></label></td>
                        </tr>
                    </table>
                </div>        
            </div>
        </form>
        
        <div id="pnlGreyOut" runat="server"/>
            
        <div id="pnlMsgBox" runat="server" style="visibility: hidden; z-index: 2; position: absolute; width: 100%; top: 30%">
            <div id="inner" style="background-color: #002248; border: 2px solid gainsboro; width: 300px; margin: 0px auto; text-align: center; border-radius: 10px; padding: 10px;">
                <label id="lblMsgHeader" runat="server" style="font-family: Verdana; font-weight: bold; font-size: large; color: white"></label>
                <br/>
                <br/>
                <label id="lblMsgBox" runat="server" style="font-family: Verdana; font-size: large; color: white"></label>
                <br/>
                <br/>
                <input type="hidden" id="hdnRedirectTo" runat="server"/>
                <input type="button" value="OK" style="width: 100px; height: 30px; background-color: ButtonHighlight" onclick=" closeMsgBox(); "/>
            </div>
        </div>
    </body>
</html>