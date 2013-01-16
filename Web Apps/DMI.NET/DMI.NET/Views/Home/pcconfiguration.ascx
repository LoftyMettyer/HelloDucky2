<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<h2>pcconfiguration</h2>


<script type="text/javascript">
<!--

        function pcconfiguration_window_onload() {
            window.parent.document.all.item("workframeset").cols = "*, 0";

            var frmMenu = OpenHR.getForm("menuframe", "frmMenuInfo");

            // Get menu to refresh the menu.
            menu_refreshMenu();
	
            // Load the original Network File Location values. 
            sKey = new String("documentspath_");
            sKey = sKey.concat(frmMenu.txtDatabase.value);	
            sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);


            frmConfiguration.txtDocuments.value = sPath;
            frmOriginalConfiguration.txtDocumentsPath.value = sPath;

            sKey = new String("olePath_");
            sKey = sKey.concat(frmMenu.txtDatabase.value);	
            sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            frmConfiguration.txtOLEServer.value = sPath;
            frmOriginalConfiguration.txtOLEServerPath.value = sPath;

            sKey = new String("localolePath_");
            sKey = sKey.concat(frmMenu.txtDatabase.value);	
            sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            frmConfiguration.txtOLELocal.value = sPath;
            frmOriginalConfiguration.txtOLELocalPath.value = sPath;

            sKey = new String("photopath_");
            sKey = sKey.concat(frmMenu.txtDatabase.value);	
            sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            frmConfiguration.txtPhoto.value = sPath;
            frmOriginalConfiguration.txtPhotoPath.value = sPath;

            sKey = new String("imagepath_");
            sKey = sKey.concat(frmMenu.txtDatabase.value);	
            sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            frmConfiguration.txtImage.value = sPath;
            frmOriginalConfiguration.txtImagePath.value = sPath;

            sKey = new String("tempmenufilepath_");
            sKey = sKey.concat(frmMenu.txtDatabase.value);	
            sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            if(sPath == "") {
                sPath = "c:\\";
            }
            if(sPath == "<NONE>") {
                sPath = "";
            }
            frmConfiguration.txtTempMenuFile.value = sPath;
            frmOriginalConfiguration.txtTempMenuFilePath.value = sPath;

            frmConfiguration.txtDocuments.focus();
        }

    function clearPath(psKey)
    {
        if (psKey == "DOCUMENTS") {
            frmConfiguration.txtDocuments.value = "";
        }
        if (psKey == "OLESERVER") {
            frmConfiguration.txtOLEServer.value = "";
        }
        if (psKey == "OLELOCAL") {
            frmConfiguration.txtOLELocal.value = "";
        }
        if (psKey == "PHOTO") {
            frmConfiguration.txtPhoto.value = "";
        }
        if (psKey == "IMAGE") {
            frmConfiguration.txtImage.value = "";
        }
        if (psKey == "TEMPMENUFILE") {
            frmConfiguration.txtTempMenuFile.value = "";
        }
    }

    function selectPath(psKey)
    {

        if (psKey == "DOCUMENTS") 
        {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtDocuments.value,"","Document Default Output Path"));
            frmConfiguration.txtDocuments.value = sPath;
        }

        if (psKey == "OLESERVER") 
        {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtOLEServer.value,"","OLE Path (Server)"));
            frmConfiguration.txtOLEServer.value = sPath;
        }
        if (psKey == "OLELOCAL") 
        {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtOLELocal.value,"","OLE Path (Local)"));
            frmConfiguration.txtOLELocal.value = sPath;
        }
        if (psKey == "PHOTO") 
        {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtPhoto.value,"","Photograph Path (non-linked)"));
            frmConfiguration.txtPhoto.value = sPath;
        }
        if (psKey == "IMAGE") 
        {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtImage.value,"","Image Path"));
            frmConfiguration.txtImage.value = sPath;
        }
        if (psKey == "TEMPMENUFILE") 
        {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtDocuments.value,"","Temporary Menu File Path "));
            frmConfiguration.txtTempMenuFile.value = sPath;
        }
    }

    function saveConfiguration()
    {

        var frmConfiguration = OpenHR.getForm("workframe", "frmConfiguration");

        // Validate the Documents path.
        var sPath = new String(frmConfiguration.txtDocuments.value);
        if (sPath.length > 0) 
        {
            
            if (!OpenHR.ValidateDir(sPath))
            {
                OpenHR.messageBox("The Documents Path is not valid.");
                return false;
            }
        }
	
        // Validate the OLE (server) path.
        var sPath = new String(frmConfiguration.txtOLEServer.value);
        if (sPath.length > 0) 
        {
            if (!OpenHR.ValidateDir(sPath))
            {
                OpenHR.messageBox("The OLE Path (server) is not valid.");
                return false;
            }
        }
	
        // Validate the OLE (local) path.
        sPath = frmConfiguration.txtOLELocal.value;
        if (sPath.length > 0) 
        {
            if (!OpenHR.ValidateDir(sPath))
            {
                OpenHR.messageBox("The OLE Path (local) is not valid.");
                return false;
            }
        }
	
        // Validate the Photo path.
        sPath = frmConfiguration.txtPhoto.value;
        if (sPath.length > 0) 
        {
            if (!OpenHR.ValidateDir(sPath))
            {
                OpenHR.messageBox("The Photo Path is not valid.");
                return false;
            }
        }

        // Validate the Image path.
        sPath = frmConfiguration.txtImage.value;
        if (sPath.length > 0) 
        {
            if (!OpenHR.ValidateDir(sPath))
            {
                OpenHR.messageBox("The Image Path is not valid.");
                return false;
            }
        }

        // Validate the Temp Menu File path.
        sPath = frmConfiguration.txtTempMenuFile.value;
        if (sPath.length > 0) 
        {
            if (!OpenHR.ValidateDir(sPath))
            {			OpenHR.messageBox("The Temporary Menu File Path is not valid.");
                return false;
            }
		
            try 
            {
                sTestPath = sPath;
                if (sTestPath.substr(sTestPath.length - 1, 1) != "\\") 
                {
                    sTestPath = sTestPath.concat("\\");
                }
                sTestPath = sTestPath.concat("testmenu");

                //window.parent.frames("menuframe").abMainMenu.save(sTestPath, "");
            }
            catch(e) 
            {
                OpenHR.messageBox("The Temporary Menu File Path cannot be written to.");
                return false;
            }			
        }

        // Save the registry values.
        var frmMenuInfo = OpenHR.getForm("menuframe","frmMenuInfo")
        sKey = new String("documentspath_");
        sKey = sKey.concat(frmMenuInfo.txtDatabase.value);	
        OpenHR.SaveRegistrySetting("HR Pro", "DataPaths", sKey, frmConfiguration.txtDocuments.value);

        sKey = new String("olePath_");
        sKey = sKey.concat(frmMenuInfo.txtDatabase.value);	
        OpenHR.SaveRegistrySetting("HR Pro", "DataPaths", sKey, frmConfiguration.txtOLEServer.value);

        sKey = new String("localolePath_");
        sKey = sKey.concat(frmMenuInfo.txtDatabase.value);	
        OpenHR.SaveRegistrySetting("HR Pro", "DataPaths", sKey, frmConfiguration.txtOLELocal.value);

        sKey = new String("photoPath_");
        sKey = sKey.concat(frmMenuInfo.txtDatabase.value);	
        OpenHR.SaveRegistrySetting("HR Pro", "DataPaths", sKey, frmConfiguration.txtPhoto.value);

        sKey = new String("imagePath_");
        sKey = sKey.concat(frmMenuInfo.txtDatabase.value);	
        OpenHR.SaveRegistrySetting("HR Pro", "DataPaths", sKey, frmConfiguration.txtImage.value);

        sKey = new String("tempMenuFilePath_");
        sKey = sKey.concat(frmMenuInfo.txtDatabase.value);	
        if (frmConfiguration.txtTempMenuFile.value.length == 0) {
            OpenHR.SaveRegistrySetting("HR Pro", "DataPaths", sKey, "<NONE>");
        }
        else {
            OpenHR.SaveRegistrySetting("HR Pro", "DataPaths", sKey, frmConfiguration.txtTempMenuFile.value);
        }

        // Try to use the height property of the menu.
        // If this fails then the menu has failed to load properly, so we need to define
        // a temporary menu file path.
//        try 
  //      {
    //        a = window.parent.frames("menuframe").abMainMenu.Bands.Item("mnuMainMenu").height;
      //  }
//        catch (e) 
//        {
            // The menu has failed to load properly, so we now need to reload
            // the menu.
//            window.parent.location.replace("main");
//            return;
//        }

        OpenHR.submitForm(frmConfiguration);

    }

    function okClick()
    {
        frmConfiguration.txtReaction.value = "DEFAULT";
        saveConfiguration();
    }

    /* Return to the default page. */
    function cancelClick()
    {
        // Try to use the height property of the menu.
        // If this fails then the menu has failed to load properly, so we need to define
        // a temporary menu file path.
   //     try 
   //     {
   //         a = window.parent.frames("menuframe").abMainMenu.Bands.Item("mnuMainMenu").height;
   //     }
   //     catch (e) 
   //     {
            // The menu has failed to load properly, so we now need to reload
            // the menu.
   //         okClick();
   //         return;
   //     }

        debugger;

        if (definitionChanged() == false) {
            window.location.href = "main";
            return;
        }


        answer = OpenHR.messageBox("You have changed the current configuration. Save changes ?",3,"");
        if (answer == 7) {
            // No
            window.location.href = "main";
            return (false);
        }
        if (answer == 6) {
            // Yes
            frmConfiguration.txtReaction.value = "DEFAULT";
            saveConfiguration();
        }
    }

    function saveChanges(psAction, pfPrompt, pfTBOverride)
    {
        if (definitionChanged() == false) {
            return 7; //No to saving the changes, as none have been made.
        }

        answer = OpenHR.messageBox("You have changed the current definition. Save changes ?",3,"");
        if (answer == 7) {
            // No
            return 7;
        }
        if (answer == 6) {
            // Yes
            frmConfiguration.txtReaction.value = psAction;
            saveConfiguration();
        }

        return 2; //Cancel.
    }

    function definitionChanged()
    {
        // Compare the network file location values with the original values.
        if (frmConfiguration.txtDocuments.value != frmOriginalConfiguration.txtDocumentsPath.value) {
            return true;
        }
		
        if (frmConfiguration.txtOLEServer.value != frmOriginalConfiguration.txtOLEServerPath.value) {
            return true;
        }
		
        if (frmConfiguration.txtOLELocal.value != frmOriginalConfiguration.txtOLELocalPath.value) {
            return true;
        }
		
        if (frmConfiguration.txtPhoto.value != frmOriginalConfiguration.txtPhotoPath.value) {
            return true;
        }		

        if (frmConfiguration.txtImage.value != frmOriginalConfiguration.txtImagePath.value) {
            return true;
        }		

        if (frmConfiguration.txtTempMenuFile.value != frmOriginalConfiguration.txtTempMenuFilePath.value) {
            return true;
        }		

        // If you reach here then nothing has changed.
        return false;
    }
    -->
</script>




    <OBJECT 
	classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0" 
	VIEWASTEXT>
	<PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
</OBJECT>

<OBJECT classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB" 
	id=dlg 
    codebase="cabs/comdlg32.cab#Version=1,0,0,0"
	style="LEFT: 0px; TOP: 0px" 
	VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="847">
	<PARAM NAME="_ExtentY" VALUE="847">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="CancelError" VALUE="0">
	<PARAM NAME="Color" VALUE="0">
	<PARAM NAME="Copies" VALUE="1">
	<PARAM NAME="DefaultExt" VALUE="">
	<PARAM NAME="DialogTitle" VALUE="">
	<PARAM NAME="FileName" VALUE="">
	<PARAM NAME="Filter" VALUE="">
	<PARAM NAME="FilterIndex" VALUE="0">
	<PARAM NAME="Flags" VALUE="0">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="FontName" VALUE="">
	<PARAM NAME="FontSize" VALUE="8">
	<PARAM NAME="FontStrikeThru" VALUE="0">
	<PARAM NAME="FontUnderLine" VALUE="0">
	<PARAM NAME="FromPage" VALUE="0">
	<PARAM NAME="HelpCommand" VALUE="0">
	<PARAM NAME="HelpContext" VALUE="0">
	<PARAM NAME="HelpFile" VALUE="">
	<PARAM NAME="HelpKey" VALUE="">
	<PARAM NAME="InitDir" VALUE="">
	<PARAM NAME="Max" VALUE="0">
	<PARAM NAME="Min" VALUE="0">
	<PARAM NAME="MaxFileSize" VALUE="260">
	<PARAM NAME="PrinterDefault" VALUE="1">
	<PARAM NAME="ToPage" VALUE="0">
	<PARAM NAME="Orientation" VALUE="1"></OBJECT>


    <form action="confirmok" method="post" id="frmConfiguration" name="frmConfiguration">
	<br>
	
	<table align=center class="outline" cellPadding=5 cellSpacing=0>
		<TR>
			<TD>
				<table align=center class="invisible" cellPadding=0 cellSpacing=0>
					<TR>
						<td height=10 colspan=7></td>
					</TR>
					<TR>
						<td align=center colspan=7>
							<STRONG>Network File Locations</STRONG>
						</td>
					</TR>
						
					<TR>
						<td height=10 colspan=7></td>
					</TR>
						
					<TR>
						<td width=20></td>
						<td align=left nowrap>
							Document Default Output Path :
						</td>
						<td width=20></td>
						<td align=left>
							<INPUT id=txtDocuments name=txtDocuments class="text" style="HEIGHT: 22px; WIDTH: 200px" width=200>		
						</td>
						<TD width=20>
							<INPUT type="button" class="btn" style="WIDTH: 30px" value="..." id=btnDocuments name=btnDocuments 
							    onclick="selectPath('DOCUMENTS')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<TD width=20>
							<INPUT type="button" class="btn" value="Clear" id=btnClearDocuments name=btnClearDocuments 
							    onclick="clearPath('DOCUMENTS')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<td width=20></td>
					</TR>

					<TR>
						<td height="5" colspan=7></td>
					</TR>

					<TR>
						<td width=20></td>
						<td align=left nowrap>
							OLE Path (Server) :
						</td>
						<td width=20></td>
						<td align=left>
							<INPUT id=txtOLEServer name=txtOLEServer class="text" style="HEIGHT: 22px; WIDTH: 200px" width=200 
				     >		
						</td>
						<TD width=20>
							<INPUT type="button" class="btn" style="WIDTH: 30px" value="..." id=btnOLEServer name=btnOLEServer 
							    onclick="selectPath('OLESERVER')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<TD width=20>
							<INPUT type="button" class="btn" value="Clear" id=btnClearOLEServer name=btnClearOLEServer 
							    onclick="clearPath('OLESERVER')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<td width=20></td>
					</TR>

					<TR>
						<td height="5" colspan=7></td>
					</TR>

					<TR>
						<td width=20></td>
						<td align=left nowrap>
							OLE Path (Local) :
						</td>
						<td width=20></td>
						<td align=left>
							<INPUT id=txtOLELocal name=txtOLELocal class="text" style="HEIGHT: 22px; WIDTH: 200px" width=200>		
						</td>
						<TD width=20>
							<INPUT type="button" class="btn" style="WIDTH: 30px" value="..." id=btnOLELocal name=btnOLELocal 
							    onclick="selectPath('OLELOCAL')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </TD>
						<TD width=20>
							<INPUT type="button" class="btn" value="Clear" id=btnClearOLELocal name=btnClearOLELocal 
							    onclick="clearPath('OLELOCAL')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<td width=20></td>
					</TR>

					<TR>
						<td height="5" colspan=7>
						</td>
					</TR>

					<TR>
						<td width=20></td>
						<td align=left nowrap>
							Photograph Path (non-linked) :
						</td>
						<td width=20></td>
						<td align=left>
							<INPUT id=txtPhoto name=txtPhoto class="text" style="HEIGHT: 22px; WIDTH: 200px" width=200>		
						</td>
						<TD width=20>
							<INPUT type="button" class="btn" style="WIDTH: 30px" value="..." id=btnPhoto name=btnPhoto 
							    onclick="selectPath('PHOTO')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<TD width=20>
							<INPUT type="button" class="btn" value="Clear" id=btnClearPhoto name=btnClearPhoto 
							    onclick="clearPath('PHOTO')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<td width=20></td>
					</TR>

					<TR>
						<td height="5" colspan=7>
						</td>
					</TR>

					<TR>
						<td width=20></td>
						<td align=left nowrap>
							Image Path :
						</td>
						<td width=20></td>
						<td align=left>
							<INPUT id=txtImage name=txtImage class="text" style="HEIGHT: 22px; WIDTH: 200px" width=200>		
						</td>
						<TD width=20>
							<INPUT type="button" class="btn" style="WIDTH: 30px" value="..." id=btnImage name=btnImage 
							    onclick="selectPath('IMAGE')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<TD width=20>
							<INPUT type="button" class="btn" value="Clear" id=btnClearImage name=btnClearImage 
							    onclick="clearPath('IMAGE')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<td width=20></td>
					</TR>

					<TR>
						<td height="5" colspan=7>
						</td>
					</TR>

					<TR>
						<td width=20></td>
						<td align=left nowrap>
							Temporary Menu File Path :
						</td>
						<td width=20></td>
						<td align=left>
							<INPUT id=txtTempMenuFile name=txtTempMenuFile class="text" style="HEIGHT: 22px; WIDTH: 200px" width=200>		
						</td>
						<TD width=20>
							<INPUT type="button" class="btn" style="WIDTH: 30px" value="..." id=btnTempMenuFile name=btnTempMenuFile 
							    onclick="selectPath('TEMPMENUFILE')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<TD width=20>
							<INPUT type="button" class="btn" value="Clear" id=btnClearTempMenuFile name=btnClearTempMenuFile 
							    onclick="clearPath('TEMPMENUFILE')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
						</TD>
						<td width=20></td>
					</TR>

					<TR>
						<td height="20" colspan=7>
						</td>
					</TR>

					<TR>
						<td height="5" colspan=7>
							<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0 align="center">
								<TD>&nbsp;</TD>
								<TD width=80>
									<input id="btnDiv2OK" name="btnDiv2OK" type="button" class="btn" value="OK" style="WIDTH: 75px" width="75" 
									    onclick="okClick()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=20></TD>
								<TD width=80>
									<input id="btnDiv2Cancel" name="btnDiv2Cancel" type="button" class="btn" value="Cancel" style="WIDTH: 75px" width="75" 
									    onclick="cancelClick()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD>&nbsp;</TD>
							</TABLE>
						</td>
					</TR>
					<TR>
						<td height=10 colspan=7></td>
					</TR>
				</table>
			</TD>
		</TR>
	</table>

	<INPUT type="hidden" id=txtReaction name=txtReaction>
</form>

<FORM id=frmOriginalConfiguration name=frmOriginalConfiguration>	
	<INPUT type="hidden" id=txtDocumentsPath name=txtDocumentsPath>
	<INPUT type="hidden" id=txtOLEServerPath name=txtOLEServerPath>
	<INPUT type="hidden" id=txtOLELocalPath name=txtOLELocalPath>
	<INPUT type="hidden" id=txtPhotoPath name=txtPhotoPath>	
	<INPUT type="hidden" id=txtImagePath name=txtImagePath>	
	<INPUT type="hidden" id=txtTempMenuFilePath name=txtTempMenuFilePath>	
</FORM>

<FORM action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</FORM>

<script type="text/javascript">
    pcconfiguration_window_onload();
</script>
