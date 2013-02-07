<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">

    function pcconfiguration_window_onload() {

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
        if (sPath == "") {
            sPath = "c:\\";
        }
        if (sPath == "<NONE>") {
            sPath = "";
        }
        frmConfiguration.txtTempMenuFile.value = sPath;
        frmOriginalConfiguration.txtTempMenuFilePath.value = sPath;

        frmConfiguration.txtDocuments.focus();
    }

    function clearPath(psKey) {
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

    function selectPath(psKey) {

        if (psKey == "DOCUMENTS") {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtDocuments.value, "", "Document Default Output Path"));
            frmConfiguration.txtDocuments.value = sPath;
        }

        if (psKey == "OLESERVER") {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtOLEServer.value, "", "OLE Path (Server)"));
            frmConfiguration.txtOLEServer.value = sPath;
        }
        if (psKey == "OLELOCAL") {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtOLELocal.value, "", "OLE Path (Local)"));
            frmConfiguration.txtOLELocal.value = sPath;
        }
        if (psKey == "PHOTO") {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtPhoto.value, "", "Photograph Path (non-linked)"));
            frmConfiguration.txtPhoto.value = sPath;
        }
        if (psKey == "IMAGE") {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtImage.value, "", "Image Path"));
            frmConfiguration.txtImage.value = sPath;
        }
        if (psKey == "TEMPMENUFILE") {
            var sPath = new String(menu_selectFolder(frmConfiguration.txtDocuments.value, "", "Temporary Menu File Path "));
            frmConfiguration.txtTempMenuFile.value = sPath;
        }
    }

    function saveConfiguration() {

        var frmConfiguration = OpenHR.getForm("workframe", "frmConfiguration");

        // Validate the Documents path.
        var sPath = new String(frmConfiguration.txtDocuments.value);
        if (sPath.length > 0) {

            if (!OpenHR.ValidateDir(sPath)) {
                OpenHR.messageBox("The Documents Path is not valid.");
                return false;
            }
        }

        // Validate the OLE (server) path.
        var sPath = new String(frmConfiguration.txtOLEServer.value);
        if (sPath.length > 0) {
            if (!OpenHR.ValidateDir(sPath)) {
                OpenHR.messageBox("The OLE Path (server) is not valid.");
                return false;
            }
        }

        // Validate the OLE (local) path.
        sPath = frmConfiguration.txtOLELocal.value;
        if (sPath.length > 0) {
            if (!OpenHR.ValidateDir(sPath)) {
                OpenHR.messageBox("The OLE Path (local) is not valid.");
                return false;
            }
        }

        // Validate the Photo path.
        sPath = frmConfiguration.txtPhoto.value;
        if (sPath.length > 0) {
            if (!OpenHR.ValidateDir(sPath)) {
                OpenHR.messageBox("The Photo Path is not valid.");
                return false;
            }
        }

        // Validate the Image path.
        sPath = frmConfiguration.txtImage.value;
        if (sPath.length > 0) {
            if (!OpenHR.ValidateDir(sPath)) {
                OpenHR.messageBox("The Image Path is not valid.");
                return false;
            }
        }

        // Validate the Temp Menu File path.
        sPath = frmConfiguration.txtTempMenuFile.value;
        if (sPath.length > 0) {
            if (!OpenHR.ValidateDir(sPath)) {
                OpenHR.messageBox("The Temporary Menu File Path is not valid.");
                return false;
            }

            try {
                sTestPath = sPath;
                if (sTestPath.substr(sTestPath.length - 1, 1) != "\\") {
                    sTestPath = sTestPath.concat("\\");
                }
                sTestPath = sTestPath.concat("testmenu");

                //window.parent.frames("menuframe").abMainMenu.save(sTestPath, "");
            }
            catch (e) {
                OpenHR.messageBox("The Temporary Menu File Path cannot be written to.");
                return false;
            }
        }

        // Save the registry values.
        var frmMenuInfo = OpenHR.getForm("menuframe", "frmMenuInfo")
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

    function okClick() {
        frmConfiguration.txtReaction.value = "DEFAULT";
        saveConfiguration();
    }

    /* Return to the default page. */
    function cancelClick() {
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


        answer = OpenHR.messageBox("You have changed the current configuration. Save changes ?", 3, "");
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

    function saveChanges(psAction, pfPrompt, pfTBOverride) {
        if (definitionChanged() == false) {
            return 7; //No to saving the changes, as none have been made.
        }

        answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3, "");
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

    function definitionChanged() {
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

</script>




<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0"
    viewastext>
    <param name="LPKPath" value="lpks/main.lpk">
</object>

<object classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
    id="dlg"
    codebase="cabs/comdlg32.cab#Version=1,0,0,0"
    style="LEFT: 0px; TOP: 0px"
    viewastext>
    <param name="_ExtentX" value="847">
    <param name="_ExtentY" value="847">
    <param name="_Version" value="393216">
    <param name="CancelError" value="0">
    <param name="Color" value="0">
    <param name="Copies" value="1">
    <param name="DefaultExt" value="">
    <param name="DialogTitle" value="">
    <param name="FileName" value="">
    <param name="Filter" value="">
    <param name="FilterIndex" value="0">
    <param name="Flags" value="0">
    <param name="FontBold" value="0">
    <param name="FontItalic" value="0">
    <param name="FontName" value="">
    <param name="FontSize" value="8">
    <param name="FontStrikeThru" value="0">
    <param name="FontUnderLine" value="0">
    <param name="FromPage" value="0">
    <param name="HelpCommand" value="0">
    <param name="HelpContext" value="0">
    <param name="HelpFile" value="">
    <param name="HelpKey" value="">
    <param name="InitDir" value="">
    <param name="Max" value="0">
    <param name="Min" value="0">
    <param name="MaxFileSize" value="260">
    <param name="PrinterDefault" value="1">
    <param name="ToPage" value="0">
    <param name="Orientation" value="1">
</object>

<form action="confirmok" method="post" id="frmConfiguration" name="frmConfiguration">
    <br>

    <table align="center" class="outline" cellpadding="5" cellspacing="0">
        <tr>
            <td>
                <table align="center" class="invisible" cellpadding="0" cellspacing="0">
                    <tr>
                        <td height="10" colspan="7"></td>
                    </tr>
                    <tr>
                        <td align="center" colspan="7">
                            <strong>Network File Locations</strong>
                        </td>
                    </tr>

                    <tr>
                        <td height="10" colspan="7"></td>
                    </tr>

                    <tr>
                        <td width="20"></td>
                        <td align="left" nowrap>Document Default Output Path :
                        </td>
                        <td width="20"></td>
                        <td align="left">
                            <input id="txtDocuments" name="txtDocuments" class="text" style="HEIGHT: 22px; WIDTH: 200px" width="200">
                        </td>
                        <td width="20">
                            <input type="button" class="btn" style="WIDTH: 30px" value="..." id="btnDocuments" name="btnDocuments"
                                onclick="selectPath('DOCUMENTS')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20">
                            <input type="button" class="btn" value="Clear" id="btnClearDocuments" name="btnClearDocuments"
                                onclick="clearPath('DOCUMENTS')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20"></td>
                    </tr>

                    <tr>
                        <td height="5" colspan="7"></td>
                    </tr>

                    <tr>
                        <td width="20"></td>
                        <td align="left" nowrap>OLE Path (Server) :
                        </td>
                        <td width="20"></td>
                        <td align="left">
                            <input id="txtOLEServer" name="txtOLEServer" class="text" style="HEIGHT: 22px; WIDTH: 200px" width="200">
                        </td>
                        <td width="20">
                            <input type="button" class="btn" style="WIDTH: 30px" value="..." id="btnOLEServer" name="btnOLEServer"
                                onclick="selectPath('OLESERVER')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20">
                            <input type="button" class="btn" value="Clear" id="btnClearOLEServer" name="btnClearOLEServer"
                                onclick="clearPath('OLESERVER')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20"></td>
                    </tr>

                    <tr>
                        <td height="5" colspan="7"></td>
                    </tr>

                    <tr>
                        <td width="20"></td>
                        <td align="left" nowrap>OLE Path (Local) :
                        </td>
                        <td width="20"></td>
                        <td align="left">
                            <input id="txtOLELocal" name="txtOLELocal" class="text" style="HEIGHT: 22px; WIDTH: 200px" width="200">
                        </td>
                        <td width="20">
                            <input type="button" class="btn" style="WIDTH: 30px" value="..." id="btnOLELocal" name="btnOLELocal"
                                onclick="selectPath('OLELOCAL')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20">
                            <input type="button" class="btn" value="Clear" id="btnClearOLELocal" name="btnClearOLELocal"
                                onclick="clearPath('OLELOCAL')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20"></td>
                    </tr>

                    <tr>
                        <td height="5" colspan="7"></td>
                    </tr>

                    <tr>
                        <td width="20"></td>
                        <td align="left" nowrap>Photograph Path (non-linked) :
                        </td>
                        <td width="20"></td>
                        <td align="left">
                            <input id="txtPhoto" name="txtPhoto" class="text" style="HEIGHT: 22px; WIDTH: 200px" width="200">
                        </td>
                        <td width="20">
                            <input type="button" class="btn" style="WIDTH: 30px" value="..." id="btnPhoto" name="btnPhoto"
                                onclick="selectPath('PHOTO')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20">
                            <input type="button" class="btn" value="Clear" id="btnClearPhoto" name="btnClearPhoto"
                                onclick="clearPath('PHOTO')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20"></td>
                    </tr>

                    <tr>
                        <td height="5" colspan="7"></td>
                    </tr>

                    <tr>
                        <td width="20"></td>
                        <td align="left" nowrap>Image Path :
                        </td>
                        <td width="20"></td>
                        <td align="left">
                            <input id="txtImage" name="txtImage" class="text" style="HEIGHT: 22px; WIDTH: 200px" width="200">
                        </td>
                        <td width="20">
                            <input type="button" class="btn" style="WIDTH: 30px" value="..." id="btnImage" name="btnImage"
                                onclick="selectPath('IMAGE')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20">
                            <input type="button" class="btn" value="Clear" id="btnClearImage" name="btnClearImage"
                                onclick="clearPath('IMAGE')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20"></td>
                    </tr>

                    <tr>
                        <td height="5" colspan="7"></td>
                    </tr>

                    <tr>
                        <td width="20"></td>
                        <td align="left" nowrap>Temporary Menu File Path :
                        </td>
                        <td width="20"></td>
                        <td align="left">
                            <input id="txtTempMenuFile" name="txtTempMenuFile" class="text" style="HEIGHT: 22px; WIDTH: 200px" width="200">
                        </td>
                        <td width="20">
                            <input type="button" class="btn" style="WIDTH: 30px" value="..." id="btnTempMenuFile" name="btnTempMenuFile"
                                onclick="selectPath('TEMPMENUFILE')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20">
                            <input type="button" class="btn" value="Clear" id="btnClearTempMenuFile" name="btnClearTempMenuFile"
                                onclick="clearPath('TEMPMENUFILE')"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width="20"></td>
                    </tr>

                    <tr>
                        <td height="20" colspan="7"></td>
                    </tr>

                    <tr>
                        <td height="5" colspan="7">
                            <table width="100%" class="invisible" cellspacing="0" cellpadding="0" align="center">
                                <td>&nbsp;</td>
                                <td width="80">
                                    <input id="btnDiv2OK" name="btnDiv2OK" type="button" class="btn" value="OK" style="WIDTH: 75px" width="75"
                                        onclick="okClick()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
                                </td>
                                <td width="20"></td>
                                <td width="80">
                                    <input id="btnDiv2Cancel" name="btnDiv2Cancel" type="button" class="btn" value="Cancel" style="WIDTH: 75px" width="75"
                                        onclick="cancelClick()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
                                </td>
                                <td>&nbsp;</td>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td height="10" colspan="7"></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

    <input type="hidden" id="txtReaction" name="txtReaction">
</form>

<form id="frmOriginalConfiguration" name="frmOriginalConfiguration">
    <input type="hidden" id="txtDocumentsPath" name="txtDocumentsPath">
    <input type="hidden" id="txtOLEServerPath" name="txtOLEServerPath">
    <input type="hidden" id="txtOLELocalPath" name="txtOLELocalPath">
    <input type="hidden" id="txtPhotoPath" name="txtPhotoPath">
    <input type="hidden" id="txtImagePath" name="txtImagePath">
    <input type="hidden" id="txtTempMenuFilePath" name="txtTempMenuFilePath">
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<script type="text/javascript">
    pcconfiguration_window_onload();
</script>
