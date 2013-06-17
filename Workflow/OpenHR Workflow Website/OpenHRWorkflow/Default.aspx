<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" EnableSessionState="True" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" id="htmMain">
<head id="Head1" runat="server">
    <meta name="format-detection" content="telephone=no"/>
    
    <link rel="shortcut icon" href="images/logo.ico"/>
    <link href="Css/default.css" rel="stylesheet" type="text/css" />
    <link href="Content/themes/base/jquery.ui.all.css" rel="stylesheet" type="text/css" />

    <script src="Scripts/jquery-1.5.1.min.js" type="text/javascript"></script>
    <script src="Scripts/jquery-ui-1.8.21.custom.min.js" type="text/javascript"></script>
    <script src="Scripts/autoNumeric-1.7.4.js" type="text/javascript"></script>

	<title></title>
    
    <script type="text/javascript">
        jQuery.noConflict();

        jQuerySetup = function() {
            
            jQuery.datepicker.setDefaults({
                changeYear: true,
                changeMonth: true,
                showOtherMonths: true,
                selectOtherMonths: true,
                dateFormat: window.localeDateFormatjQuery
            });

            jQuery('input.date').datepicker();
            jQuery('input.numeric').autoNumeric({ aSep: '', aDec: window.localeDecimal, wEmpty: 'zero', vMin: '-9999.99', vMax: '9999.99' });

            jQuery('input.date').change(function() {
                var $this = jQuery(this);
                var value = $this.val();
                try {
                    var date = jQuery.datepicker.parseDate(window.localeDateFormatjQuery, value);
                    if (date!= null) {
                        $this.val(jQuery.datepicker.formatDate(window.localeDateFormatjQuery, date));
                        jQuery.datepicker.setDefaults({ defaultDate: date });   
                    }
                } catch(e) {
                    $this.val('');
                }
            });

            jQuery('input.date').keyup(function(e) {
                //F2 should set todays date
                if (e.which == 113) {
                    var date = new Date();
                    jQuery(this).val(jQuery.datepicker.formatDate(window.localeDateFormatjQuery, date));
                    jQuery.datepicker.setDefaults({ defaultDate: date });
                }
            });

        };

        jQuery(jQuerySetup);

    </script>
</head>
<body id="bdyMain" style="overflow: auto; text-align: center; background-color: <%= ColourThemeHex()%>;">
<%-- TODO PG get rid off text-align: center --%>

  <form runat="server" id="frmMain" onsubmit="return submitForm();" autocomplete="off">
	    
  <div id="innerMeasurements" style="visibility:hidden;background-color:red;position:fixed;top:0px;left:0px;right:0px;bottom:0px;">
  </div>    

  <script type="text/javascript">
    //Fault HRPRO-2269 - includes the 'innerMeasurements' div shown above.
      window.currentHeight = document.getElementById("innerMeasurements").offsetHeight;
      window.currentWidth = document.getElementById("innerMeasurements").offsetWidth;

      window.localeDateFormat = "<%= LocaleDateFormat() %>";
      window.localeDateFormatjQuery = "<%= LocaleDateFormatjQuery() %>";
      window.localeDecimal = "<%= LocaleDecimal() %>";
  </script>

    <%--    <script src="Scripts/default.js" type="text/javascript"></script>
    <script src="Scripts/resizable-table.js" type="text/javascript"></script>--%>
     
    <%--TODO remove the file WebNumericEditValidation.js --%>
    <ajx:ToolkitScriptManager ID="tsm" runat="server" 
        EnablePartialRendering="true" EnablePageMethods="true" CombineScripts="True" 
        LoadScriptsBeforeUI="True">
        <CompositeScript>
            <Scripts>
                <asp:ScriptReference Name="MicrosoftAjax.js" />  
                <asp:ScriptReference Name="MicrosoftAjaxWebForms.js" />  
                <asp:ScriptReference Path="~/Scripts/default.js" />
                <asp:ScriptReference Path="~/Scripts/resizable-table.js" />
            </Scripts>
        </CompositeScript>
    </ajx:ToolkitScriptManager>   

    <div id="pleasewaitScreen" style="display:none">
        <span id="pleasewaitText">Processing...<br/><br/>Please wait.</span>
	</div>
		
	<!--
    Submission and Exceptional Errors Popup
    -->
	<div id="divSubmissionMessages" style="position: absolute; left: 0px; top: 15%; width: 100%; display: none; z-index: 102; visibility: hidden; text-align: center;">
		<iframe id="ifrmMessages" src="" frameborder="0" scrolling="no"></iframe>
	</div>
	<!--
    File Upload Popup
    -->
	<div id="divFileUpload" style="position: absolute; left: 0px; top: 15%; width: 100%; display: none; z-index: 101; text-align: center;">
		<iframe id="ifrmFileUpload" src="" style="width:550px" frameborder="0" scrolling="no"></iframe>
	</div>
    
    <div id="divOverlay"></div>

	<!--
        Web Form Controls
        -->
	<div id="divInput" style="z-index: 0; text-align: center; float:left" runat="server">
        <asp:UpdatePanel ID="pnlInput" runat="server">
            <ContentTemplate>
                <div id = "pnlInputDiv" runat="server" style="position:relative; margin: 0 auto;">
                    
                    <div id="pnlTabsDiv" style="position: absolute;" runat="server">
                        <div id="pnlTabsBorder" style="position: absolute; top: 20px; left: 0; right: 0; bottom: 0; border: 1px solid black;">
                        </div>
                    </div>
                </div>    
                <asp:Button id="btnDoFilter" runat="server" style="display:none;"/>
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
                <img id="errorMessageMax" src="Images/uparrows_white.gif" alt="Show messages" style="display: none;" onclick="showErrorMessages('max');" />

                <asp:Panel id="errorMessagePanel" runat="server" style="display: none;">
                    
                        <img id="errorMessageMin" src="Images/downarrows_white.gif" alt="Hide messages" onclick="showErrorMessages('min');" />
                            		    
                        <asp:Label ID="lblErrors" runat="server" Text=""/>
                            			
                        <asp:BulletedList ID="bulletErrors" runat="server" Style="margin-top: 0px; margin-bottom: 0px; padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" BorderStyle="None" />

                        <asp:Label ID="lblWarnings" runat="server" Text=""/>

                        <asp:BulletedList ID="bulletWarnings" runat="server" Style="margin-top: 0px; margin-bottom: 0px; padding-top: 5px; padding-bottom: 5px;" BulletStyle="Disc" BorderStyle="None" />
                            
                        <span id="overrideWarning" runat="server">Click 
                        <span onclick="overrideWarningsAndSubmit();" style="cursor:pointer; text-decoration:underline;">here</span> 
                        to ignore the warnings and submit the form.</span>

                </asp:Panel>

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
    <input type="hidden" id="txtActiveDDE" name="txtActiveDDE" value="" />	
   
    <script type="text/javascript">
        window_onload();
    </script>
</body>
</html>
