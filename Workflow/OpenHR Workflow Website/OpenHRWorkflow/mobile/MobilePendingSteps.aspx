<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobilePendingSteps.aspx.vb" Inherits="PendingSteps" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
    <title>OpenHR Mobile </title>
    
<style type="text/css">
body {
    font-family: Verdana;
    }

dh {
    background-color: #FFFFFF;
    border: 1px solid #999999;
    color: #222222;
    display: block;
    font-family: Verdana;
    font-size: 17px;
    font-weight: bold;
    margin-bottom: -1px;
    padding: 12px 10px;
    text-decoration: none;
    text-shadow: 0px 1px 0px #fff;
    background-image: -webkit-gradient(linear, left top, left bottom, from(#ccc), to(#999));
    }    

</style>

    <script type="text/javascript">
// <!CDATA[

        function window_onload() {

            if (document.getElementById('hdnStepCount').value == 0) {
                document.getElementById('lblNothingTodo').style.visibility = "visible";
                document.getElementById('lblNothingTodo').style.display = "block";
            }
            else {
                document.getElementById('lblInstruction').style.visibility = "visible";
                document.getElementById('lblInstruction').style.display = "block";
            }
        }

// ]]>
    </script>
</head>
    
<body onload="return window_onload()" style="margin:0px;overflow:hidden">
    <form id="form1" runat="server">

        <div id="pnlContainer" runat="server" style="overflow:hidden;background-color:Red">
          <div id="pnlHeader" runat="server" style="position:absolute;overflow:hidden;left:0px;top:0px;z-index:1;width:100%;height:57px">
          </div>
          <div id="ScrollerFrame" runat="server" style="position:fixed;left:0px;top:57px;z-index:1;bottom:60px;width:100%">
            <div id="pnlBody"  runat="server" style="position:absolute;width:100%;height:100%;z-index:1">      
                  
              <div runat="server" id="pnlWFList" style="width:100%;height:100%;top:40px;bottom:0px;overflow:auto">                  
                <label id="lblNothingTodo" runat="server" style="visibility:hidden;display:none;margin:15px;">You have nothing in your 'To Do' list.</label>
                <label id="lblInstruction" runat="server" style="visibility:hidden;display:none;margin:15px;">Click on a 'to do' item to view the details and complete your action.</label>
              </div>

            </div>
          </div>
          
          <div  style="text-align:center; position:absolute;top:357px;width:100%;z-index:1">
            <%--<p style="font-family: Verdana; font-size: 10px; z-index: 2; color: #333366;">Copyright © Advanced Business Software and Solutions Ltd 2012</p>--%>
          </div>

         <div id="pnlFooter" runat="server" style="position:fixed;overflow:hidden;left:0px;bottom:0px;z-index:1;width:100%;height:60px">
            <table id="tblFooter" runat="server" style="height:100%;width:100%">
              <tr style="height:40px">
                <td style="width:50%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnRefresh" runat="server" /></td>
                <td style="width:50%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnCancel" runat="server" /></td>
              </tr>
              <tr style="height:17px">
                <td style="width:50%;text-align:center;overflow:hidden"><label runat="server" id="btnRefresh_label"></label></td>
                <td style="width:50%;text-align:center;overflow:hidden"><label runat="server" id="btnCancel_label"></label></td>
              </tr>
            </table>
          </div>        
    </div>

    <asp:HiddenField ID="hdnStepCount" runat="server" Value="0" />

    </form>
</body>

</html>
