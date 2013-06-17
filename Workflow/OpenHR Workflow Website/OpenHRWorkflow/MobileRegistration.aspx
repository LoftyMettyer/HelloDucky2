<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobileRegistration.aspx.vb" Inherits="Registration" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">   
	<meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
    <title>OpenHR Mobile </title>

    <script type="text/javascript">
// <!CDATA[

        function window_onload() {
          window.scrollTo(0, 1);
          document.getElementById("txtEmail").focus();
        }
        
        function submitCheck() {
          if (frmRegistration.txtEmail.value.length == 0) {
               showMsgBox("No email address entered.");
                return false;
            }

//            if (frmRegistration.txtUserName.value.length == 0) {
//              showMsgBox("A Username must be entered.");
//                return false;
//            }
            

//            if (frmRegistration.txtPassword.value != frmRegistration.txtConfPassword.value) {
//              showMsgBox("The passwords do not match.");
//                return false;
//            }
            
//            if (frmRegistration.txtPassword.value.length == 0) {
//              showMsgBox("A password must be entered.");
//                return false;
//            }
            
        }

        function showMsgBox(strText) {
          lblMsgBox.innerHTML = strText;
          pnlGreyOut.style.visibility = "visible";
          pnlMsgBox.style.visibility = "visible";


        }

        function closeMsgBox() {
          pnlGreyOut.style.visibility = "hidden";
          pnlMsgBox.style.visibility = "hidden";

            strNextPage = '<%=session("nextPage") %>';
          if(strNextPage.length > 0) window.location.replace('<%=session("nextPage") %>.aspx');
        }

// ]]>
    </script>
</head>
<body onload="return window_onload()" style="margin:0px;overflow:hidden">
    <form id="frmRegistration" runat="server" defaultbutton="btnRegister">
      <div id="pnlContainer" runat="server" style="overflow:hidden;background-color:Red">
        <div id="pnlHeader" runat="server" style="position:absolute;overflow:hidden;left:0px;top:0px;z-index:1;width:100%;height:57px">
        </div>
          <div id="ScrollerFrame" runat="server" style="position:fixed;left:0px;top:57px;z-index:1;bottom:60px;width:100%">
            <div id="pnlBody" runat="server" style="height:100%;z-index:1">      
              <table style="width:100%;height:100%" >
                <tr id="space1" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <td colspan="2"><label  id="lblWelcome" runat="server">lblWelcome</label></td>
                </tr>
                <tr id="space2" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <td><label id="lblEmail" runat="server">lblEmail</label></td>
                  <td><input id="txtEmail" runat="server" /></td>
                </tr>
               <tr id="space3" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <%--<td><label id="lblUserName" runat="server">lblUserName</label></td>
                  <td><input id="txtUserName" runat="server" type="text"/></td>--%>
                </tr>
                <tr id="space4" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <%--<td><label id="lblPassword" runat="server">lblPassword</label></td>
                  <td><input id="txtPassword" runat="server" type="password"/></td>--%>
                </tr>
                <tr id="space5" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <%--<td><label id="lblConfPassword" runat="server">lblConfPassword</label></td>
                  <td><input id="txtConfPassword" runat="server" type="password"/></td>--%>
                </tr>
                <tr id="space6" style="width: 100%;height:80%"><td></td></tr>
               </table>
            </div>
          </div>
          
          <div  style="text-align:center; position:absolute;top:357px;width:100%;z-index:1">
            <%--<p style="font-family: Verdana; font-size: 10px; z-index: 2; color: #333366;">Copyright © Advanced Business Software and Solutions Ltd 2012</p>--%>
          </div>

          <div id="pnlFooter" runat="server" style="position:fixed;overflow:hidden;left:0px;bottom:0px;z-index:1;width:100%;height:60px">
            <table id="tblFooter" runat="server" style="height:100%;width:100%">
              <tr style="height:40px">
                <td style="width:50%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnRegister" runat="server" OnClientClick="return submitCheck();"/></td>
                <td style="width:50%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnHome" runat="server" /></td>
              </tr>
              <tr style="height:17px">
                <td style="width:50%;text-align:center;overflow:hidden"><label runat="server" id="btnRegister_label"></label></td>
                <td style="width:50%;text-align:center;overflow:hidden"><label runat="server" id="btnHome_label"></label></td>
              </tr>
            </table>
          </div>        
        </div>


        <div id="pnlGreyOut" runat="server" style="position: absolute;visibility: hidden;width: 100%;height: 100%;filter:alpha(opacity=50);
                              -moz-opacity:0.5;opacity: 0.5;background-color: #222;margin:0px;z-index:1">
        </div>

       <div id="pnlMsgBox" runat="server" style="visibility: hidden;z-index:2;position:absolute;width:100%;top:30%">
           <div id="inner" style="background-color: #002248;border:2px solid gainsboro;width:300px;margin:0px auto;text-align: center;border-radius:10px;padding: 10px;">
             <label id="lblMsgHeader" runat="server" style="font-family: Verdana;font-weight: bold;font-size:large;color:white">Registration Failed</label>
             <br/>
             <br/>
             <label id="lblMsgBox" runat="server" style="font-family: Verdana;font-size:large;color:white"></label>
             <br/>
             <br/>
             <input type="button" value="OK" style="width:100px;height:30px;background-color: ButtonHighlight" onclick="closeMsgBox();"/>
           </div>
      </div>


<%--        <div id="pnlMsgBox" runat="server" style="background-color: #002248;position: absolute;visibility: hidden;width: 400px;height:150px;
                                          text-align:center;position:absolute;left:50%;top:50%;margin-left:-200px;margin-top:-75px;z-index:2;
                                          border-radius:10px;border:2px solid #001648;vertical-align:middle;opacity:0.8">
          <br/>&nbsp;&nbsp;
          <label id="lblMsgBox" runat="server" style="font-family: Verdana;font-size:large;color:white"></label>
          <br/>&nbsp;&nbsp;<br/>
          <input type="button" value="OK" style="width:100px;height:30px;background-color: ButtonHighlight" onclick="closeMsgBox();"/>
        </div>  
--%>
    </form>
</body>
</html>
