<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MobileHome.aspx.vb" Inherits="Home" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
    <title>OpenHR Mobile </title>

    <style type="text/css">
    body {
          font-family: Verdana;
        }
    </style>

    <script type="text/javascript">
// <!CDATA[

        function window_onload() {
            var dtExpiry = new Date();
            
            //window.scrollTo(0, 1);

            try{
                fixDate(dtExpiry);
                /*
                cookie expires in one year (actually, 365 days)
                365 days in a year
                24 hours in a day
                60 minutes in an hour
                60 seconds in a minute
                1000 milliseconds in a second
                */
                dtExpiry.setTime(dtExpiry.getTime() + 365 * 24 * 60 * 60 * 1000);
                
                setCookie("Login", '<%=replace(cstr(Session("LoginKey")), "\", "\\") %>', dtExpiry);
            }
            catch(e){
                alert(e);
            }
        }
        
        /*
           name - name of the cookie
           value - value of the cookie
           [expires] - expiration date of the cookie
             (defaults to end of current session)
           [path] - path for which the cookie is valid
             (defaults to path of calling document)
           [domain] - domain for which the cookie is valid
             (defaults to domain of calling document)
           [secure] - Boolean value indicating if the cookie transmission requires
             a secure transmission
           * an argument defaults when it is assigned null as a placeholder
           * a null placeholder is not required for trailing omitted arguments
        */

        function setCookie(name, value, expires, path, domain, secure) {
          
            var curCookie = name + "=" + escape(value) +
              ((expires) ? "; expires=" + expires.toGMTString() : "") +
              ((path) ? "; path=" + path : "") +
              ((domain) ? "; domain=" + domain : "") +
              ((secure) ? "; secure" : "");
            document.cookie = curCookie;
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
        
        function deleteCookie(name, path, domain) {
            if (getCookie(name)) {
                document.cookie = name + "=" +
                ((path) ? "; path=" + path : "") +
                ((domain) ? "; domain=" + domain : "") +
                "; expires=Thu, 01-Jan-70 00:00:01 GMT";
            }
        }
        
        // date - any instance of the Date object
        // * hand all instances of the Date object to this function for "repairs"

        function fixDate(date) {
          var base = new Date(0);
          var skew = base.getTime();
          if (skew > 0)
            date.setTime(date.getTime() - skew);
        }


// ]]>
    </script>
</head>
<body onload="return window_onload()" style="margin:0px;overflow:hidden">
    <form id="form2" runat="server">

        <div id="pnlContainer" runat="server" style="overflow:hidden;background-color:Red">
          <div id="pnlHeader" runat="server" style="position:absolute;overflow:hidden;left:0px;top:0px;z-index:1;width:100%;height:57px">
          </div>
          <div id="ScrollerFrame" runat="server" style="position:fixed;left:0px;top:57px;bottom:60px;z-index:1;width:100%">
            <div id="pnlBody" runat="server" style="height:100%;z-index:1">
            <table>
                 <tr id="space1" style="width: 100%"><td></td></tr>
                <tr style="width: 100%; height:21px">
                  <td colspan="2"><label style="margin:15px" id="lblWelcome" runat="server">lblWelcome</label></td>
                </tr>
                <tr id="space2" style="width: 100%"><td></td></tr>  
                </table>
              <div runat="server" id="pnlWFList" style="margin:15px;width:100%;height:100%;top:40px;bottom:0px;overflow:auto">
              </div>
            </div>
          </div>
          
<%--          <div  style="text-align:center; position:absolute;top:357px;width:100%;z-index:1">
            <p style="font-family: Verdana; font-size: 10px; z-index: 2; color: #333366;">Copyright © Advanced Business Software and Solutions Ltd 2012</p>
          </div>
--%>
         <div id="pnlFooter" runat="server" style="position:fixed;overflow:hidden;left:0px;bottom:0px;z-index:1;width:100%;height:60px">
            <table id="tblFooter" runat="server" style="height:100%;width:100%">
              <tr style="height:40px">
                <td style="width:33%;text-align:center;overflow:hidden">
                  <div style="position:relative;width:40px;height:100%;margin:auto">
                    <asp:ImageButton ID="btnToDoList" runat="server"/>
                    <div id="pnlWFCount" runat="server" style="position:absolute;top: 0px;right: -6px;padding:1px 2px 1px 2px;background-color:Red;color:White;font-family:verdana;font-weight:bold;font-size:0.75em;border-radius:30px;box-shadow:1px 1px 1px gray;">
                      <label id="lblWFCount" runat="server"></label>
                    </div>
                  </div>
                </td>
                <td style="width:33%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnChangePwd" runat="server" /></td>
                <td style="width:33%;text-align:center;overflow:hidden"><asp:ImageButton ID="btnLogout" runat="server" /></td>
              </tr>
              <tr style="height:17px">
                <td style="width:33%;text-align:center;overflow:hidden"><label runat="server" id="btnToDoList_label"></label></td>
                <td style="width:33%;text-align:center;overflow:hidden"><label runat="server" id="btnChangePwd_label"></label></td>
                <td style="width:33%;text-align:center;overflow:hidden"><label runat="server" id="btnLogout_label"></label></td>
              </tr>
            </table>
          </div>        

 
        </div>
        <div id="pnlGreyOut" runat="server" style="position: absolute;visibility: hidden;width: 100%;height: 100%;filter:alpha(opacity=50);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         -moz-opacity:0.5;opacity: 0.5;background-color: #222;margin:0px;z-index:1">
        </div>
        
         <div id="pnlMsgBox" runat="server" style="visibility: hidden;z-index:2;position:absolute;width:100%;top:30%">
             <div id="inner" style="background-color: #002248;border:2px solid gainsboro;width:300px;margin:0px auto;text-align: center;border-radius:10px;padding: 10px;">
               <label id="Label1" runat="server" style="font-family: Verdana;font-weight: bold;font-size:large;color:white">Request Failed</label>
               <br/>
               <br/>
               <label id="lblMsgBox" runat="server" style="font-family: Verdana;font-size:large;color:white"></label>
               <br/>
               <br/>
               <input type="button" value="OK" style="width:100px;height:30px;background-color: ButtonHighlight" onclick="closeMsgBox();"/>
             </div>
        </div>   
                     
    </form>
</body>

</html>
