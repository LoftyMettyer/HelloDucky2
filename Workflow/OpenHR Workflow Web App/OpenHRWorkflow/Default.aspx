<%@ Page Language="VB" AutoEventWireup="false" Inherits="OpenHRWorkflow.Default" EnableSessionState="True" Codebehind="Default.aspx.vb" %>
<%@ Import Namespace="OpenHRWorkflow" %> 

<%@ Register Assembly="ScriptReferenceProfiler" Namespace="ScriptReferenceProfiler" TagPrefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head runat="server">
		<title></title>
		<meta name="format-detection" content="telephone=no"/>
    
		<link rel="shortcut icon" href="images/logo.ico"/>
		<link href="~/Content/default.css" rel="stylesheet" type="text/css" />
		<link href="~/Content/themes/base/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
		  
		<script src="Scripts/jquery-2.1.4.min.js" type="text/javascript"> </script>
		<script src="Scripts/jquery-ui-1.11.4.min.js" type="text/javascript"> </script>
	</head>
	<body id="bdyMain" style="overflow: auto; background-color: <%=App.Config.ColourThemeHex()%>;">

		<form runat="server" id="frmMain" onsubmit="return submitForm();" autocomplete="off">
	    
			<div id="innerMeasurements" style="visibility: hidden; background-color: red; position: fixed; top: 0px; left: 0px; right: 0px; bottom: 0px;"></div>    

			<script type="text/javascript">
                //Fault HRPRO-2269 - includes the 'innerMeasurements' div shown above.
				window.currentHeight = document.getElementById("innerMeasurements").offsetHeight;
				window.currentWidth = document.getElementById("innerMeasurements").offsetWidth;

				window.autoFocusControl = '<%=AutoFocusControl()%>';
				window.localeDateFormat = '<%=LocaleDateFormat()%>';
				window.localeDateFormatjQuery = '<%=LocaleDateFormatjQuery()%>';
				window.localeDecimal = '<%=LocaleDecimal()%>';
				window.isMobile = <%=If(IsMobileBrowser(), "true", "false")%>;
				window.androidLayerBug = <%=If(AndroidLayerBug(), "true", "false")%>;
			</script>
     
			<asp:ScriptManager ID="PSM" runat="server" LoadScriptsBeforeUI="False" ScriptMode="Release">
				<CompositeScript>
					<Scripts>
						<asp:ScriptReference Name="WebForms.js" Assembly="System.Web" />
						<asp:ScriptReference Name="MicrosoftAjax.js" />  
						<asp:ScriptReference Name="MicrosoftAjaxWebForms.js" /> 
						<asp:ScriptReference name="Common.Common.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference name="ExtenderBase.BaseScripts.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference name="HoverExtender.HoverBehavior.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference name="DynamicPopulate.DynamicPopulateBehavior.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference name="Compat.Timer.Timer.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference name="Animation.Animations.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference name="Animation.AnimationBehavior.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference name="PopupExtender.PopupBehavior.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference name="DropDown.DropDownBehavior.js" assembly="AjaxControlToolkit, Version=4.1.51116.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e"/>
						<asp:ScriptReference Path="~/Scripts/jquery.metadata.min.js" />
						<asp:ScriptReference Path="~/Scripts/autoNumeric-1.9.25.min.js" />	
						<asp:ScriptReference Path="~/Scripts/resizable-table.min.js" />
						<asp:ScriptReference Path="~/Scripts/default.js" />
					</Scripts>
				</CompositeScript>
			</asp:ScriptManager>  

			<%--<cc1:ScriptReferenceProfiler ID="ScriptReferenceProfiler1" runat="server" />--%>

			<div id="pleasewaitScreen" style="display: none;">
				<span id="pleasewaitText">Processing...<br/><br/>Please wait.</span>
			</div>
		
			<!-- Submission and Exceptional Errors Popup -->
			<div id="divSubmissionMessages" style="position: absolute; left: 0px; top: 15%; width: 100%; display: none; z-index: 102; visibility: hidden; text-align: center;">
				<iframe id="ifrmMessages" src="" frameborder="0" scrolling="no"></iframe>
			</div>

			<!-- File Upload Popup -->
			<div id="divFileUpload" style="position: absolute; left: 0px; top: 15%; width: 100%; display: none; z-index: 101; text-align: center;">
				<iframe id="ifrmFileUpload" src="" style="width: 550px" frameborder="0" scrolling="no"></iframe>
			</div>
    
			<div id="divOverlay"></div>

			<!-- Web Form Controls -->
			<div id="divInput" style="float: left" runat="server">
				<asp:UpdatePanel ID="pnlInput" runat="server">
					<ContentTemplate>
						<div id = "pnlInputDiv" runat="server" style="position: relative; margin: 0 auto;">
                    
							<!-- Tab Control -->
							<div id="pnlTabsDiv" style="position: absolute;" runat="server">
								<div id="pnlTabsBorder" style="position: absolute; top: 20px; left: 0; right: 0; bottom: 0; border: 1px solid black;">
								</div>
							</div>							
						</div>
						<asp:Button id="btnDoFilter" runat="server" style="display: none;"/>
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
                
						<!--
							Validation Messages
							-->
						<img id="errorMessageMax" src="Images/uparrows_white.gif" alt="Show messages" style="display: none;" onclick=" showErrorMessages('max'); " />

						<asp:Panel id="errorMessagePanel" runat="server" style="display: none;">
                    
							<img id="errorMessageMin" src="Images/downarrows_white.gif" alt="Hide messages" onclick=" showErrorMessages('min'); " />
                            		    
							<asp:Label ID="lblErrors" runat="server" Text=""/>
                            			
							<asp:BulletedList ID="bulletErrors" runat="server" Style="margin-top: 0px; margin-bottom: 0px; padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" BorderStyle="None" />

							<asp:Label ID="lblWarnings" runat="server" Text=""/>

							<asp:BulletedList ID="bulletWarnings" runat="server" Style="margin-top: 0px; margin-bottom: 0px; padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" BorderStyle="None" />
                            
							<span id="overrideWarning" runat="server">Click 
								<span onclick=" overrideWarningsAndSubmit(); " style="cursor: pointer; text-decoration: underline;">here</span> 
								to ignore the warnings and submit the form.</span>

						</asp:Panel>

					</ContentTemplate>
				</asp:UpdatePanel>		
			</div>
			<!--
				Temporary values from the server
				-->
			<asp:HiddenField ID="hdnDefaultPageNo" runat="server" Value="0" />
		</form>
		<!--
			Temporary client-side values
			-->
		<input type="hidden" id="txtPostbackMode" name="txtPostbackMode" value="0" />
		<input type="hidden" id="txtActiveDDE" name="txtActiveDDE" value="" />	
   
		<script type="text/javascript">
		    InitialiseWindow();
		</script>
	</body>
</html>