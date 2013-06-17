﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobilePendingSteps.aspx.vb" Inherits="PendingSteps" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
    <head runat="server">
        <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
        <link rel="apple-touch-icon" href="/Images/Adv_hr&payroll.gif" />
        <title>OpenHR Mobile</title>
    
        <style type="text/css">
            body { font-family: Verdana; }
        </style>
    </head>
    
    <body style="margin: 0px; overflow: hidden">
        <form id="form1" runat="server">

            <div id="pnlContainer" runat="server" style="overflow: hidden; background-color: Red">
                <div id="pnlHeader" runat="server" style="position: absolute; overflow: hidden; left: 0px; top: 0px; z-index: 1; width: 100%; height: 57px">
                </div>
                <div id="ScrollerFrame" runat="server" style="position: fixed; left: 0px; top: 57px; z-index: 1; bottom: 60px; width: 100%">
                    <div id="pnlBody"  runat="server" style="position: absolute; width: 100%; height: 100%; z-index: 1">      
                  
                        <div runat="server" id="pnlWFList" style="width: 100%; height: 100%; top: 40px; bottom: 0px; overflow: auto">                  
                            <label id="lblNothingTodo" runat="server" style="display: block; margin: 15px;">You have nothing in your 'To Do' list.</label>
                            <label id="lblInstruction" runat="server" style="display: block; margin: 15px;">Click on a 'to do' item to view the details and complete your action.</label>
                        </div>

                    </div>
                </div>
          
                <div id="pnlFooter" runat="server" style="position: fixed; overflow: hidden; left: 0px; bottom: 0px; z-index: 1; width: 100%; height: 60px">
                    <table id="tblFooter" runat="server" style="height: 100%; width: 100%">
                        <tr style="height: 40px">
                            <td style="width: 50%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnRefresh" runat="server" /></td>
                            <td style="width: 50%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnCancel" runat="server" /></td>
                        </tr>
                        <tr style="height: 17px">
                            <td style="width: 50%; text-align: center; overflow: hidden"><label runat="server" id="btnRefresh_label"></label></td>
                            <td style="width: 50%; text-align: center; overflow: hidden"><label runat="server" id="btnCancel_label"></label></td>
                        </tr>
                    </table>
                </div>        
            </div>

        </form>
    </body>

</html>