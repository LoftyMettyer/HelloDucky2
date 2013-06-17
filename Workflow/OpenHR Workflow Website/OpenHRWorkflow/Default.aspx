<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" EnableSessionState="True" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" id="htmMain">
<head id="Head1" runat="server">
    <meta name="format-detection" content="telephone=no"/>
    <link rel="shortcut icon" href="images/logo.ico"/>
    <link href="CSS/default.css" rel="stylesheet" type="text/css" />
	<title></title>	
</head>

<body id="bdyMain" style="overflow: auto; text-align: center; margin: 0px; padding: 0px; background-color: <%= ColourThemeHex()%>;">
	
	<form runat="server" hidefocus="true" id="frmMain" onsubmit="return submitForm();">
	
    <script src="Scripts/default.js" type="text/javascript"></script>
    <script src="Scripts/resizable-table.js" type="text/javascript"></script>
    <script src="scripts/WebNumericEditValidation.js" type="text/javascript"></script>

    <ajx:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server" EnablePartialRendering="true" EnablePageMethods="true" CombineScripts="True">
        
    </ajx:ToolkitScriptManager>
	<!--
        Web Form Validation Error Messages
    -->        

    <div id="pleasewaitScreen" style="position:absolute;z-index:5;top:30%;width:150px;height:60px;left:50%;margin-left:-75px;visibility:hidden">
		<table border="0" cellspacing="0" cellpadding="10" style="top: 0px; left: 0px; width: 100%; height: 100%; position: relative; text-align: center; font-size: 10pt; color: black; font-family: Verdana; border: black 1px solid;" bgcolor="White">
				<tr>
					<td style="width:100%;height:100%;background-color:White;text-align:center;vertical-align:middle">
								<label id="pleasewaitText">Processing...<br/><br/>Please wait.<br/></label>
					</td>
				</tr>
		</table>
	</div>
		
    <img id="imgErrorMessages_Max" src="Images/uparrows_white.gif" alt="Show messages" style="position: absolute; right: 1px; bottom: 1px; display: none; visibility: hidden; z-index: 1; width:20px; height:20px;" onclick="showErrorMessages(true);" />

	<div id="divErrorMessages_Outer" onfilterchange="unblockErrorMessageDIV();" style="position: absolute; bottom: 0px; left: 0px; right: 0px; display: none; visibility: hidden; z-index: 1">
		
        <div id="divErrorMessages_Inner" style="background-color: white; text-align: left; position: relative; margin: 0px; padding: 5px; border: 1px solid; font-size: 11px; color: black; font-family: Verdana;">
		    
			<img id="imgErrorMessages_Min" src="Images/downarrows_white.gif" alt="Hide messages" style="right: 1px; position: absolute; top: 0px; width:20px; height:20px;" onclick="showErrorMessages(false);" />
            
			<igmisc:WebAsyncRefreshPanel id="pnlErrorMessages" runat="server" style="position: relative;" width="90%" height="100%">
				<asp:Label ID="lblErrors" runat="server" Text=""></asp:Label>				
				<asp:BulletedList ID="bulletErrors" runat="server" Style="margin-top: 0px; margin-bottom: 0px; padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" Font-Names="Verdana" Font-Size="11pt" BorderStyle="None">
				</asp:BulletedList>
				<asp:Label ID="lblWarnings" runat="server" Text=""></asp:Label>
				<asp:BulletedList ID="bulletWarnings" runat="server" Style="margin-top: 0px; margin-bottom: 0px; padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" Font-Names="Verdana" Font-Size="11px" BorderStyle="None">
				</asp:BulletedList>
				<asp:Label ID="lblWarningsPrompt_1" runat="server" Text="Click"></asp:Label>
				<span id="spnClickHere" name="spnClickHere" tabindex="1" style="color:#333366;" onclick="overrideWarningsAndSubmit();" onmouseover="try{this.style.color='#ff9608';}catch(e){}"
					onmouseout="try{this.style.color='#333366';}catch(e){}" onfocus="try{this.style.color='#ff9608';}catch(e){}"
					onblur="try{this.style.color='#333366';}catch(e){}" onkeypress="try{if(window.event.keyCode == 32)spnClickHere.click();}catch(e){}">
					<asp:Label ID="lblWarningsPrompt_2" runat="server" Text="here" Font-Underline="true" style="cursor: pointer;"></asp:Label>
				</span>
				<asp:Label ID="lblWarningsPrompt_3" runat="server" Text=""></asp:Label>
			</igmisc:WebAsyncRefreshPanel>
		</div>
	</div>
	<!--
    Submission and Exceptional Errors Popup 
    -->
	<div id="divSubmissionMessages" style="position: absolute; left: 0px; top: 15%; width: 100%; display: none; z-index: 3; visibility: hidden; text-align: center;" nowrap="nowrap">
		<iframe id="ifrmMessages" src="" frameborder="0" scrolling="no"></iframe>
	</div>
	<!--
    File Upload Popup
    -->
	<div id="divFileUpload" style="position: absolute; left: 0px; top: 15%; width: 100%; display: none; z-index: 3; visibility: hidden; text-align: center;" nowrap="nowrap" onfilterchange="return unblockFileUploadDIV();">
		<iframe id="ifrmFileUpload" src="" style="width:550px"  frameborder="0" scrolling="no"></iframe>
	</div>
    
    <div id="divOverlay"></div>

	<!--
        Web Form Controls
        -->
	<div id="divInput" style="top:0px; left:0px; z-index: 0; padding: 0px; margin: 0px; text-align: center;float:left" runat="server">
        <asp:UpdatePanel ID="pnlInput" runat="server">
            <ContentTemplate>
                <div id = "pnlInputDiv" runat="server" style="position:relative;padding-right:0px;padding-left:0px;padding-bottom:0px;margin-top:0px;margin-bottom:0px;margin-right:auto;margin-left:auto;padding-top:0px;">
                    
                    <div id="pnlTabsDiv" style="position: absolute;" runat="server">
                        <div id="pnlTabsBorder" style="position: absolute; top: 20px; left: 0; right: 0; bottom: 0; border: 1px solid black;">
                        </div>
                    </div>
                </div>    
                <asp:Button id="btnSubmit" runat="server" style="visibility: hidden; top: 0px; position: absolute; left: 0px; width: 0px; height: 0px;" text=""/>
                <asp:Button id="btnReEnableControls" runat="server" style="visibility: hidden; top: 0px; position: absolute; left: 0px; width: 0px; height: 0px;" text=""/>
                <asp:HiddenField ID="hdnMobileLookupFilter" runat="server" Value="" />
			    <asp:HiddenField ID="hdnCount_Errors" runat="server" Value="" />
			    <asp:HiddenField ID="hdnCount_Warnings" runat="server" Value="" />
			    <asp:HiddenField ID="hdnOverrideWarnings" runat="server" Value="0" />
			    <asp:HiddenField ID="hdnLastButtonClicked" runat="server" Value="" />
			    <asp:HiddenField ID="hdnNoSubmissionMessage" runat="server" Value="0" />
			    <asp:HiddenField ID="hdnFollowOnForms" runat="server" Value="" />
			    <asp:HiddenField ID="hdnErrorMessage" runat="server" Value="" />
			    <asp:HiddenField ID="hdnSiblingForms" runat="server" Value="" />
			    <asp:HiddenField ID="hdnSubmissionMessage_1" runat="server" Value="" />
			    <asp:HiddenField ID="hdnSubmissionMessage_2" runat="server" Value="" />
			    <asp:HiddenField ID="hdnSubmissionMessage_3" runat="server" Value="" />
	        </ContentTemplate>
        </asp:UpdatePanel>			
	</div>
	<!--
    Temporary values from the server
    -->
	<asp:HiddenField ID="hdnFormHeight" runat="server" Value="0" />
	<asp:HiddenField ID="hdnFormWidth" runat="server" Value="0" />
	<asp:HiddenField ID="hdnFirstControl" runat="server" Value="" />
    <asp:HiddenField ID="hdnDefaultPageNo" runat="server" Value="0" />
	</form>
	<!--
    Temporary client-side values
    -->
	<input type="hidden" id="txtPostbackMode" name="txtPostbackMode" value="0" />
	<input type="hidden" id="txtActiveElement" name="txtActiveElement" value="" />
	<input type="hidden" id="txtLastDate_Month" name="txtLastDate_Month" value="" />
	<input type="hidden" id="txtLastDate_Day" name="txtLastDate_Day" value="" />
	<input type="hidden" id="txtLastDate_Year" name="txtLastDate_Year" value="" />	
	<input type="hidden" id="txtActiveDDE" name="txtActiveDDE" value="" />	
    
    <script type="text/javascript">
        window_onload();
    </script>
</body>
</html>
