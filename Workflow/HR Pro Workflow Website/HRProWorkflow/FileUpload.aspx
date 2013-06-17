<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FileUpload.aspx.vb" Inherits="FileUpload" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <title></title>

  <script language="javascript" type="text/javascript">
// <!CDATA[
    function window_onload() {
      // Resize the frame.
      try {
        if (frmFileUpload.hdnCount_Errors.value > 0) {
          showErrorMessages(true);
        }
        else {
          if (frmFileUpload.hdnExitMode.value > 0) {
            exitFileUpload(frmFileUpload.hdnExitMode.value); // 1 = cleared, 2 = uploaded
            return;
          }
        }
      }
      catch (e) { };

      resizeFrame();
    }

    function resizeFrame() {
      var iRequiredWidth = frmFileUpload.offsetParent.scrollWidth;
      var iMinWidth = 0.9 * window.parent.frmMain.hdnFormWidth.value;

      if (iRequiredWidth < iMinWidth) {
        iRequiredWidth = iMinWidth;
      }

      document.getElementById('FileUpload1').style.width = (iRequiredWidth - 150) + 'px';

      window.resizeTo(iRequiredWidth, frmFileUpload.offsetParent.scrollHeight);
      window.parent.resizeToFit(frmFileUpload.offsetParent.scrollWidth, frmFileUpload.offsetParent.scrollHeight);
    }

    function showErrorMessages(pfDisplay) {
      if (((frmFileUpload.hdnCount_Errors.value > 0)
				|| (frmFileUpload.hdnCount_Warnings.value > 0))
				&& (pfDisplay == false)) {
        imgErrorMessages_Max.style.display = "block";
        imgErrorMessages_Max.style.visibility = "visible";
      }
      else {
        imgErrorMessages_Max.style.display = "none";
        imgErrorMessages_Max.style.visibility = "hidden";
      }

      if (pfDisplay == true) {
        //divErrorMessages_Inner.style.visibility = "visible";
        divErrorMessages_Outer.style.filter = "revealTrans(duration=0.3, transition=4)";
        divErrorMessages_Outer.filters.revealTrans.apply();
        divErrorMessages_Outer.style.display = "block";
        divErrorMessages_Outer.style.visibility = "visible";
        divErrorMessages_Outer.filters.revealTrans.play();
      }
      else {
        divErrorMessages_Outer.style.filter = "revealTrans(duration=0.3, transition=5)";
        divErrorMessages_Outer.filters.revealTrans.apply();
        divErrorMessages_Outer.style.visibility = "hidden";
        //divErrorMessages_Outer.style.display = "none";
        //divErrorMessages_Inner.style.visibility = "hidden";
        divErrorMessages_Outer.filters.revealTrans.play();

      }
    }

    function unblockErrorMessageDIV() {
      try {
        if ((divErrorMessages_Outer.style.visibility == "hidden") &&
					(divErrorMessages_Outer.style.display != "none")) {
          divErrorMessages_Outer.style.display = "none";
        }
      }
      catch (e) { }
    }

    function exitFileUpload(piExitMode) {
      try {
        window.parent.fileUploadDone(frmFileUpload.hdnElementID.value, piExitMode);
      }
      catch (e) { }
    }

    function refreshFileUploadButton(psFileUploadValue) {
      try {
        // Trim leading and trailing spaces.
        psFileUploadValue = psFileUploadValue.replace(/(^\s*)|(\s*$)/g, "");

        var button = ig_getWebControlById("btnFileUpload");

        if (psFileUploadValue.length > 0) {
          button.setEnabled(true);
        }
        else {
          button.setEnabled(false);
        }
      }
      catch (e) { }
    }

// ]]>
  </script>

</head>
<body onload="return window_onload()" scroll="auto" style="overflow: auto; padding: 0px;
  margin: 0px; border: 0px; text-align: center;">
  <img id="imgErrorMessages_Max" src="Images/uparrows_white.gif" alt="Show messages"
    style="position: absolute; right: 1px; bottom: 1px; display: none; visibility: hidden;
    z-index: 1;" onclick="showErrorMessages(true);" />
  <form id="frmFileUpload" runat="server" style="height: 100%; width: 100%; top: 0px;
  left: 0px;">
  <!--
    Web Form Validation Error Messages
    -->
  <div id="divErrorMessages_Outer" onfilterchange="unblockErrorMessageDIV();" style="position: absolute;
    width: 100%; bottom: 0px; left: 0px; right: 0px; display: none; visibility: hidden;
    z-index: 1">
    <div id="divErrorMessages_Inner" style="background-color: white; text-align: left;
      position: relative; margin: 0px; padding: 5px; border: 1px solid; font-size: 8pt;
      color: black; font-family: Verdana;">
      <img id="imgErrorMessages_Min" src="Images/downarrows_white.gif" alt="Hide messages"
        style="right: 1px; position: absolute; top: 0px;" onclick="showErrorMessages(false);" />
      <asp:Label ID="lblErrors" runat="server" Text="Unable to upload the file due to the following error:"></asp:Label>
      <asp:BulletedList ID="bulletErrors" runat="server" Style="margin-top: 0px; margin-bottom: 0px;
        padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" Font-Names="Verdana"
        Font-Size="8pt" BorderStyle="None">
      </asp:BulletedList>
    </div>
  </div>
  <!--
    File Upload Controls
    -->
  <div id="divFileUpload" style="z-index: 0; width: 100%; text-align: center; padding: 0px;
    margin: 0px;">
    <table border="0" cellspacing="0" cellpadding="0" style="top: 0px; left: 0px; width: 100%;
      height: 100%; position: relative; text-align: center; font-size: 10pt; color: black;
      font-family: Verdana; border: black 1px solid;" bgcolor="White">
      <tr style="background-color: <%=ColourThemeHex()%>;">
        <td colspan="5" height="10">
        </td>
      </tr>
      <tr style="height: 40px">
        <td width="10" style="background-color: <%=ColourThemeHex()%>;">
          &nbsp;&nbsp;
        </td>
        <td width="40" valign="top">
          <img src="themes/<%=ColourThemeFolder()%>/CrnrTop.gif" alt="" width="40" height="40" />
        </td>
        <td rowspan="2" style="background-color: White" nowrap="nowrap">
          <br />
          <asp:Label ID="lblFileUploadPrompt" runat="server" Text="Select the file you wish to upload:"
            Font-Names="Verdana"></asp:Label>
        </td>
        <td width="40" valign="top">
          <img src="themes/<%=ColourThemeFolder()%>/RCrnrTop.gif" alt="" width="40" height="40" />
        </td>
        <td width="10" style="background-color: <%=ColourThemeHex()%>;">
          &nbsp;&nbsp;
        </td>
      </tr>
      <tr>
        <td width="10" style="background-color: <%=ColourThemeHex()%>;">
        </td>
        <td>
        </td>
        <td>
        </td>
        <td width="10" style="background-color: <%=ColourThemeHex()%>;">
        </td>
      </tr>
      <tr style="height: 40px">
        <td width="10" style="background-color: <%=ColourThemeHex()%>;">
          &nbsp;&nbsp;
        </td>
        <td>
        </td>
        <td style="background-color: White" valign="middle">
          <asp:FileUpload ID="FileUpload1" runat="server" Width="400" onChange="refreshFileUploadButton(this.value);"
            onKeyUp="refreshFileUploadButton(this.value);" />
        </td>
        <td>
        </td>
        <td width="10" style="background-color: <%=ColourThemeHex()%>;">
          &nbsp;&nbsp;
        </td>
      </tr>
      <tr>
        <td width="10" style="background-color: <%=ColourThemeHex()%>;">
        </td>
        <td>
        </td>
        <td rowspan="2">
          <igtxt:WebImageButton ID="btnFileUpload" runat="server" text="Upload" AccessKey="U"
            UnderlineAccessKey="true" enabled="false">
          </igtxt:WebImageButton>&nbsp;
          <igtxt:WebImageButton ID="btnClear" runat="server" text="Clear" AccessKey="l" UnderlineAccessKey="true">
          </igtxt:WebImageButton>&nbsp;
          <igtxt:WebImageButton ID="btnCancel" runat="server" text="Cancel" AccessKey="C" UnderlineAccessKey="true">
          </igtxt:WebImageButton>
          <br />
          <br />
        </td>
        <td>
        </td>
        <td width="10" style="background-color: <%=ColourThemeHex()%>;">
        </td>
      </tr>
      <%--NB. Keep <TD><IMG></TD> tags all on the same line, otherwise the images do not fully align to bottom--%>
      <tr style="height: 40px">
        <td width="10" bgcolor="<%=ColourThemeHex()%>">
        </td>
        <td width="40" valign="bottom"><img src="themes/<%=ColourThemeFolder()%>/CrnrBot.gif" width="40" height="40" alt="" /></td>
        <td width="40" valign="bottom"><img src="themes/<%=ColourThemeFolder()%>/RCrnrBot.gif" width="40" height="40" alt="" /></td>
        <td width="10" bgcolor="<%=ColourThemeHex()%>"></td>
      </tr>
      <tr bgcolor="<%=ColourThemeHex()%>">
        <td colspan="5" height="10">
        </td>
      </tr>
    </table>
  </div>
  <asp:HiddenField ID="hdnCount_Errors" runat="server" Value="" />
  <asp:HiddenField ID="hdnElementID" runat="server" Value="" />
  <asp:HiddenField ID="hdnExitMode" runat="server" Value="0" />
  </form>
</body>
</html>
