<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobileHome.aspx.vb" Inherits="Home" %>

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

            <div id="pnlPage" runat="server" style="overflow: hidden;">
                
                <div id="pnlHeader" runat="server" />
                
                <div id="pnlBody" runat="server">
                                            
                    <label id="lblNothingTodo" runat="server">lblNothingTodo</label>
                    <label id="lblWelcome" runat="server">lblWelcome</label>

                    <div runat="server" id="pnlWFList" />
                </div>
          

                <div id="pnlFooter" runat="server">
                    <table id="tblFooter" runat="server" style="height: 100%; width: 100%">
                        <tr style="height: 40px">
                            <td style="width: 33%; text-align: center; overflow: hidden">
                                <div style="position: relative; width: 40px; height: 100%; margin: auto">
                                    <asp:ImageButton ID="btnToDoList" runat="server"/>
                                    <div id="pnlWFCount" runat="server" style="position: absolute; top: 0px; right: -6px; padding: 1px 2px 1px 2px; background-color: Red; color: White; font-family: verdana; font-weight: bold; font-size: 0.75em; border-radius: 30px; box-shadow: 1px 1px 1px gray;">
                                        <label id="lblWFCount" runat="server"></label>
                                    </div>
                                </div>
                            </td>
                            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnChangePwd" runat="server" /></td>
                            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnLogout" runat="server" /></td>
                        </tr>
                        <tr style="height: 17px">
                            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnToDoList_label"></label></td>
                            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnChangePwd_label"></label></td>
                            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnLogout_label"></label></td>
                        </tr>
                    </table>
                </div>         
            </div>       
        </form>
        
        <div id="pnlGreyOut" runat="server" />
            
        <div id="pnlMsgBox" runat="server" style="visibility: hidden; z-index: 2; position: absolute; width: 100%; top: 30%">
            <div id="inner" style="background-color: #002248; border: 2px solid gainsboro; width: 300px; margin: 0px auto; text-align: center; border-radius: 10px; padding: 10px;">
                <label id="lblMsgHeader" runat="server" style="font-family: Verdana; font-weight: bold; font-size: large; color: white"></label>
                <br/>
                <br/>
                <label id="lblMsgBox" runat="server" style="font-family: Verdana; font-size: large; color: white"></label>
                <br/>
                <br/>
                <input type="button" value="OK" style="width: 100px; height: 30px; background-color: ButtonHighlight" onclick="closeMsgBox(); "/>
            </div>
        </div>  
    </body>

</html>