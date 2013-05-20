<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<!DOCTYPE html>

<html>
<head id="Head1" runat="server">
    <title>Absence Details</title>
    
    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />

    <script type="text/javascript">
        function absenceCalendar_details_window_onload() {

            var iResizeBy;
            var iNewWidth;
            var iNewHeight;

            self.focus();

            // Resize the grid to show all prompted values.
            iResizeBy = frmDetails.offsetParent.scrollWidth - frmDetails.offsetParent.clientWidth;
            if (frmDetails.offsetParent.offsetWidth + iResizeBy > screen.width) {
                window.dialogWidth = new String(screen.width) + "px";
            } else {
                iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
                iNewWidth = iNewWidth + iResizeBy;
                window.dialogWidth = new String(iNewWidth) + "px";
            }

            iResizeBy = frmDetails.offsetParent.scrollHeight - frmDetails.offsetParent.clientHeight;
            if (frmDetails.offsetParent.offsetHeight + iResizeBy > screen.height) {
                window.dialogHeight = new String(screen.height) + "px";
            } else {
                iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
                iNewHeight = iNewHeight + iResizeBy;
                window.dialogHeight = new String(iNewHeight) + "px";
            }
        }

    </script>

</head>
<body>

    <form id="frmDetails" name="frmDetails">
        <table class="outline" cellpadding="0" cellspacing="7" width="100%">
            <tr>
                <td>Start&nbsp;Date :</td>
                <td><%Response.Write(Request("txtStartDate"))%>&nbsp;
		<%Response.Write(Request("txtStartSession"))%>
                </td>
            </tr>
            <tr>
                <td>End Date :</td>
                <td><%Response.Write(Request("txtEndDate"))%>&nbsp;
		<%Response.Write(Request("txtEndSession"))%>
                </td>
            </tr>
            <tr>
                <td>Duration :</td>
                <td><%Response.Write(Request("txtDuration"))%></td>
            </tr>
            <tr>
                <td>Type :</td>
                <td><%Response.Write(Request("txtType"))%></td>
            </tr>
            <tr>
                <td>Type Code :</td>
                <td><%Response.Write(Request("txtTypeCode"))%></td>
            </tr>
            <tr>
                <td>Calendar Code :</td>
                <td><%Response.Write(Request("txtCalCode"))%></td>
            </tr>
            <tr>
                <td>Reason :</td>
                <td><%Response.Write(Request("txtReason"))%></td>
            </tr>

            <% If Request("txtDisableRegions") = "False" Then%>
            <tr>
                <td>Region :</td>
                <td><%Response.Write(Request("txtRegion"))%></td>
            </tr>
            <% End If%>

            <% If Request("txtDisableWPs") = "False" Then%>
            <tr>
                <td>Working Pattern :</td>
                <td>
                    <%
                        Dim objAbsenceCalendar As New HR.Intranet.Server.AbsenceCalendar
                        Response.Write(objAbsenceCalendar.HTML_WorkingPattern(Request("txtWorkingPattern")))
                        objAbsenceCalendar = Nothing
                    %>
                </td>
            </tr>
            <% End If%>

  <tr>
    <td colspan="2">
		<input id="cmdOK" name="cmdOK" class="btn" type="button" value="OK" style="HEIGHT: 24px; WIDTH: 89px" 
		    onclick="self.close()"	
            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
            onmouseout="try{button_onMouseOut(this);}catch(e){}"
            onfocus="try{button_onFocus(this);}catch(e){}"
            onblur="try{button_onBlur(this);}catch(e){}" />
    </td>
  </tr>
  </table>
</form>

    <script type="text/javascript">
        absenceCalendar_details_window_onload();
    </script>

</body>
	

</html>
