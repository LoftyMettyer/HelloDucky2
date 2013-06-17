<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobilePendingSteps.aspx.vb" Inherits="PendingSteps" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
    <head runat="server">
        <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
        <link rel="apple-touch-icon" href="/Images/Adv_hr&payroll.gif" />
        <link href="../CSS/mobile.css" rel="stylesheet" type="text/css" />
        <title>OpenHR Mobile</title>
    </head>
    
    <body>
        <form id="form1" runat="server">

            <div id="pnlContainer" runat="server" style="overflow: hidden;">

                <div id="pnlHeader" runat="server"/>
                
                <div id="pnlBody" runat="server">
                    
                    <label id="lblNothingTodo" runat="server">You have nothing in your 'To Do' list.</label>
                    <label id="lblInstruction" runat="server">Click on a 'to do' item to view the details and complete your action.</label>
                    <div runat="server" id="pnlWFList" />                  

                </div>
          
                <div id="pnlFooter" runat="server">
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