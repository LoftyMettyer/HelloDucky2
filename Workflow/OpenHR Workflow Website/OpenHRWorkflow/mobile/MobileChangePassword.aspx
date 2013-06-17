﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobileChangePassword.aspx.vb" Inherits="ChangePassword" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
    <head id="Head1" runat="server">
        <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
        <link rel="apple-touch-icon" href="/Images/Adv_hr&payroll.gif" />
        <title>OpenHR Mobile</title>
    
        <script type="text/javascript">
// <!CDATA[

            function window_onload() {
                document.getElementById('txtCurrPassword').focus();
            }

            function submitCheck() {

                if (document.getElementById('txtCurrPassword').value.length == 0 && document.getElementById('txtNewPassword').value.length == 0) {
                    showMsgBox("No passwords entered.");
                    return false;
                }

                if (document.getElementById('txtNewPassword').value != document.getElementById('txtConfPassword').value) {
                    showMsgBox("The new passwords do not match.");
                    return false;
                }
                if (document.getElementById('txtNewPassword').value.length == 0) {
                    showMsgBox("A new password must be entered.");
                    return false;
                }
                return true;
            }

            function showMsgBox(strText) {
                document.getElementById('lblMsgBox').innerHTML = strText;
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
    <body onload="return window_onload()" style="margin: 0px; overflow: hidden">
        <form id="frmChangePassword" runat="server" defaultbutton="btnSubmit">
            <div id="pnlContainer" runat="server" style="overflow: hidden; background-color: Red">
                <div id="pnlHeader" runat="server" style="position: absolute; overflow: hidden; left: 0px; top: 0px; z-index: 1; width: 100%; height: 57px">
                </div>
                <div id="ScrollerFrame" runat="server" style="position: fixed; left: 0px; top: 57px; z-index: 1; bottom: 60px; width: 100%">
                    <div id="pnlBody" runat="server" style="height: 100%; z-index: 1">      
                        <table style="position: absolute; width: 100%; height: 100%" >
                            <tr id="space1" style="width: 100%"><td></td></tr>
                            <tr style="width: 100%; height: 21px">
                                <td colspan="2"><label style="margin: 15px" id="lblWelcome" runat="server">lblWelcome</label></td>
                            </tr>
                            <tr id="space2" style="width: 100%"><td></td></tr>
                            <tr style="width: 100%; height: 21px">
                                <td><label style="margin: 15px" id="lblCurrPassword" runat="server">lblCurrPassword</label></td>
                                <td><input id="txtCurrPassword" runat="server" type="password"/></td>
                            </tr>
                            <tr id="space3" style="width: 100%"><td></td></tr>
                            <tr style="width: 100%; height: 21px">
                                <td><label style="margin: 15px" id="lblNewPassword" runat="server">lblNewPassword</label></td>
                                <td><input id="txtNewPassword" runat="server" type="password"/></td>
                            </tr>
                            <tr id="space4" style="width: 100%"><td></td></tr>
                            <tr style="width: 100%; height: 21px">
                                <td><label style="margin: 15px" id="lblConfPassword" runat="server">lblConfPassword</label></td>
                                <td><input id="txtConfPassword" runat="server" type="password"/></td>
                            </tr>
                            <tr id="space6" style="width: 100%; height: 80%"><td></td></tr>
                        </table>
                    </div>
                </div>
          
                <div id="pnlFooter" runat="server" style="position: fixed; overflow: hidden; left: 0px; bottom: 0px; z-index: 1; width: 100%; height: 60px">
                    <table id="tblFooter" runat="server" style="height: 100%; width: 100%">
                        <tr style="height: 40px">
                            <td style="text-align: center; overflow: hidden"><asp:ImageButton ID="btnSubmit" runat="server" OnClientClick="return submitCheck();"/></td>
                            <td style="text-align: center; overflow: hidden"><asp:ImageButton ID="btnCancel" runat="server" /></td>
                        </tr>
                        <tr style="height: 17px">
                            <td style="text-align: center; overflow: hidden"><label runat="server" id="btnSubmit_label"></label></td>
                            <td style="text-align: center; overflow: hidden"><label runat="server" id="btnCancel_label"></label></td>
                        </tr>
                    </table>
                </div>        
            </div>

            <div id="pnlGreyOut" runat="server" style="position: absolute; visibility: hidden; width: 100%; height: 100%; filter: alpha(opacity=50); -moz-opacity: 0.5; opacity: 0.5; background-color: #222; margin: 0px; z-index: 1">
            </div>

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
            
        </form>
    </body>
</html>