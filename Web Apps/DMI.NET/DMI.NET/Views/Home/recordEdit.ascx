<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import namespace="DMI.NET" %>

<script type="text/javascript">

    function recordEdit_window_onload() {

        $("#ctlRecordEdit").tabs();

        var frmRecordEditForm = OpenHR.getForm("workframe", "frmRecordEditForm");

        var fOK;
        fOK = true;
        var sErrMsg = frmRecordEditForm.txtErrorDescription.value;
        if (sErrMsg.length > 0) {
            fOK = false;
            OpenHR.messageBox(sErrMsg);
            window.parent.location.replace("login");
        }

        if (fOK == true) {
            // Expand the work frame and hide the option frame.
            //window.parent.document.all.item("workframeset").cols = "*, 0";
            $("#workframe").attr("data-framesource", "RECORDEDIT");
           
            var recEditCtl = document.getElementById("ctlRecordEdit"); // frmRecordEditForm.ctlRecordEdit;

            if (recEditCtl == null) {
                fOK = false;

                // The recEdit control was not loaded properly.
                OpenHR.messageBox("Record Edit control not loaded.");
                window.location = "login";
            }
        }

        if (fOK == true) {
            //TODO:
            //var sKey = new String("photopath_");
            //sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);
            //var sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            //frmRecordEditForm.txtPicturePath.value = sPath;

            //sKey = new String("imagepath_");
            //sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);
            //sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            //frmRecordEditForm.txtImagePath.value = sPath;

            //sKey = new String("olePath_");
            //sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);
            //sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            //frmRecordEditForm.txtOLEServerPath.value = sPath;

            //sKey = new String("localolePath_");
            //sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);
            //sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
            //frmRecordEditForm.txtOLELocalPath.value = sPath;


            // Read and then reset the HR Pro Navigation flag.
            var HRProNavigationFlagValue;
            //var HRProNavigationFlag = window.parent.frames("menuframe").document.forms("frmWorkAreaInfo").txtHRProNavigation;            
            var HRProNavigationFlag = document.getElementById("txtHRProNavigation");
            HRProNavigationFlagValue = HRProNavigationFlag.value;
            HRProNavigationFlag.value = 0;

            if (HRProNavigationFlagValue == 0) {
                var frmGoto = OpenHR.getForm("workframe", "frmGoto");
                frmGoto.txtGotoTableID.value = frmRecordEditForm.txtCurrentTableID.value;
                frmGoto.txtGotoViewID.value = frmRecordEditForm.txtCurrentViewID.value;
                frmGoto.txtGotoScreenID.value = frmRecordEditForm.txtCurrentScreenID.value;
                frmGoto.txtGotoOrderID.value = frmRecordEditForm.txtCurrentOrderID.value;
                frmGoto.txtGotoRecordID.value = frmRecordEditForm.txtCurrentRecordID.value;
                frmGoto.txtGotoParentTableID.value = frmRecordEditForm.txtCurrentParentTableID.value;
                frmGoto.txtGotoParentRecordID.value = frmRecordEditForm.txtCurrentParentRecordID.value;
                frmGoto.txtGotoPage.value = "recordEdit.asp";

                HRProNavigationFlag.value = 1;
                //frmGoto.submit();
                OpenHR.submitForm(frmGoto);
            } else {

                // Set the recEdit control properties.               
                //TODO: initialise clears the recordDMI activeX object and sets the module variables as below.
                //fOK = recEditCtl.initialise(
                //    frmRecordEditForm.txtRecEditTableID.value,
                //    frmRecordEditForm.txtRecEditHeight.value,
                //    frmRecordEditForm.txtRecEditWidth.value + 1,
                //    frmRecordEditForm.txtRecEditTabCount.value,
                //    frmRecordEditForm.txtRecEditTabCaptions.value,
                //    frmRecordEditForm.txtRecEditFontName.value,
                //    frmRecordEditForm.txtRecEditFontSize.value,
                //    frmRecordEditForm.txtRecEditFontBold.value,
                //    frmRecordEditForm.txtRecEditFontItalic.value,
                //    frmRecordEditForm.txtRecEditFontUnderline.value,
                //    frmRecordEditForm.txtRecEditFontStrikethru.value,
                //    frmRecordEditForm.txtRecEditRealSource.value,
                //    frmRecordEditForm.txtPicturePath.value,
                //    frmRecordEditForm.txtRecEditEmpTableID.value,
                //    frmRecordEditForm.txtRecEditCourseTableID.value,
                //    frmRecordEditForm.txtRecEditTBStatusColumnID.value,
                //    frmRecordEditForm.txtRecEditCourseCancelDateColumnID.value
                //);

                if (fOK == true) {
                    // Get the recEdit control to instantiate the required controls.
                    var sControlName;
                    var controlCollection = frmRecordEditForm.elements;
                    if (controlCollection != null) {
                        var txtControls = new Array();
                        var txtControlsCount = 0;


                        //two loops here - the controlCollection was growing as controls were added, which didn't help.
                        for (var i = 0; i < controlCollection.length; i++) {
                            sControlName = controlCollection.item(i).name;
                            sControlName = sControlName.substr(0, 18);
                            if (sControlName == "txtRecEditControl_") {
                                //fOK = recEditCtl.addControl(controlCollection.item(i).value);
                                txtControls[txtControlsCount] = controlCollection.item(i).name;
                                txtControlsCount += 1;
                            }

                            if (fOK == false) {
                                break;
                            }
                        }

                        //Now add the form controls based on the fixed array of txtRecEditControl_ items...
                        for (var i = 0; i < txtControls.length; i++) {
                            var txtControlValue = $("#" + txtControls[i]).val();
                            AddHtmlControl(txtControlValue);
                        }

                    }
                }

                //jQuery Functionality:
                if (fOK == true) {
                    //add datepicker functionality.
                    $(".datepicker").datepicker();
                    //add spinner functionality
                    $(".spinner").spinner();
                }


                if (fOK == true) {
                    // Set the column control values in the recEdit control.
                    var sControlName;
                    var controlCollection = frmRecordEditForm.elements;
                    if (controlCollection != null) {
                        var txtControls = new Array();
                        var txtControlsCount = 0;

                        for (i = 0; i < controlCollection.length; i++) {
                            sControlName = controlCollection.item(i).name;
                            if (sControlName) {
                                sControlName = sControlName.substr(0, 24);
                                if (sControlName == "txtRecEditControlValues_") {
                                    //fOK = recEditCtl.addControlValues(controlCollection.item(i).value);
                                    txtControls[txtControlsCount] = controlCollection.item(i).name;
                                    txtControlsCount += 1;
                                }
                            }
                            if (fOK == false) {
                                break;
                            }
                        }

                        //Now add the form control values based on the fixed array of txtRecEditControl_ items...
                        for (var i = 0; i < txtControls.length; i++) {
                            var txtControlValue = $("#" + txtControls[i]).val();
                            addHTMLControlValues(txtControlValue);
                        }
                    }
                }

                if (fOK == true) {
                    // Get the recEdit control to format itself.
                    //No longer necessary

                    //recEditCtl.formatscreen();

                    //JPD 20021021 - Added picture functionality.
                    //TODO: NPG
                    if (frmRecordEditForm.txtImagePath.value.length > 0) {
                        var controlCollection = frmRecordEditForm.elements;
                        if (controlCollection != null) {
                            for (i = 0; i < controlCollection.length; i++) {
                                sControlName = controlCollection.item(i).name;
                                sControlName = sControlName.substr(0, 18);
                                if (sControlName == "txtRecEditPicture_") {
                                    sControlName = controlCollection.item(i).name;
                                    iPictureID = new Number(sControlName.substr(18));
                                    // recEditCtl.updatePicture(iPictureID, frmRecordEditForm.txtImagePath.value + "/" + controlCollection.item(i).value);
                                }
                            }
                        }
                    }
                }

                if (fOK == true) {
                    // Get the data.asp to get the required data.
                    var action = document.getElementById("txtAction");
                    if (((frmRecordEditForm.txtAction.value == "NEW") ||
                            (frmRecordEditForm.txtAction.value == "COPY")) &&
                        (frmRecordEditForm.txtRecEditInsertGranted.value == "True")) {
                        action.value = frmRecordEditForm.txtAction.value;
                    } else {
                        action.value = "LOAD";
                    }

                    if (frmRecordEditForm.txtCurrentOrderID.value != frmRecordEditForm.txtRecEditOrderID.value) {
                        frmRecordEditForm.txtCurrentOrderID.value = frmRecordEditForm.txtRecEditOrderID.value;
                    }

                    var dataForm = OpenHR.getForm("dataframe", "frmGetData");
                    dataForm.txtCurrentTableID.value = frmRecordEditForm.txtCurrentTableID.value;
                    dataForm.txtCurrentScreenID.value = frmRecordEditForm.txtCurrentScreenID.value;
                    dataForm.txtCurrentViewID.value = frmRecordEditForm.txtCurrentViewID.value;
                    dataForm.txtSelectSQL.value = frmRecordEditForm.txtRecEditSelectSQL.value;
                    dataForm.txtFromDef.value = frmRecordEditForm.txtRecEditFromDef.value;
                    dataForm.txtFilterSQL.value = "";
                    dataForm.txtFilterDef.value = "";
                    dataForm.txtRealSource.value = frmRecordEditForm.txtRecEditRealSource.value;
                    dataForm.txtRecordID.value = frmRecordEditForm.txtCurrentRecordID.value;
                    dataForm.txtParentTableID.value = frmRecordEditForm.txtCurrentParentTableID.value;
                    dataForm.txtParentRecordID.value = frmRecordEditForm.txtCurrentParentRecordID.value;
                    //dataForm.txtDefaultCalcCols.value = recEditCtl.CalculatedDefaultColumns();

                    //this should be in scope by now.
                    //TODO: NPG
                    //data_refreshData(); //window.parent.frames("dataframe").refreshData();
                }

                if (fOK != true) {
                    // The recEdit control was not initialised properly.
                    OpenHR.messageBox("Record Edit control not initialised properly.");
                    window.location = "login";
                }
            }
        }

        try {

            //frmRecordEditForm.ctlRecordEdit.SetWidth(frmRecordEditForm.txtRecEditWidth.value);

            //NPG - recedit not resizing. Do it manually.
            var newHeight = frmRecordEditForm.txtRecEditHeight.value / 15;
            var newWidth = frmRecordEditForm.txtRecEditWidth.value / 15;

            $("#ctlRecordEdit").height(newHeight + "px");
            $("#ctlRecordEdit").width(newWidth + "px");

            //parent.window.resizeBy(-1, -1);
            //parent.window.resizeBy(1, 1);
        } catch (e) {
        }

    }


    function addControl(tabNumber, controlDef) {

        var tabID = "FI_21_" + tabNumber

        if (($("#" + tabID).length <= 0) && (tabNumber > 0)) {
            //tab doesn't exist - create it...
            var tabs = $("#ctlRecordEdit").tabs(),
        tabTemplate = "<li><a href='#{href}'>#{label}</a></li>";

            var label = "Tab" + tabNumber,
                li = $(tabTemplate.replace(/#\{href\}/g, "#" + tabID).replace(/#\{label\}/g, label));

            tabs.find(".ui-tabs-nav").append(li);
            tabs.append("<div style='position: relative;' id='" + tabID + "'></div>");
            tabs.tabs("refresh");
            if (tabNumber == 1) tabs.tabs("option", "active", 0);
        }

        //add control to tab.
        try {
            $("#" + tabID).append(controlDef);
        }
        catch (e) { alert("unable to add control!"); }

    }


    function applyLocation(formItem, controlItemArray) {
        formItem.style.top = (Number(controlItemArray[4]) / 15) + "px";
        formItem.style.left = (Number(controlItemArray[5]) / 15) + "px";
        formItem.style.height = (Number(controlItemArray[6]) / 15) + "px";
        formItem.style.width = (Number(controlItemArray[7]) / 15) + "px";
        formItem.style.position = "absolute";
    }

    // -------------------------------------------------- Add the record Edit controls ------------------------------------------------
    function AddHtmlControl(controlItem) {
        var controlItemArray = controlItem.split("\t");
        var sProperty;
        var fDoingIDColumn = false;
        var iPageNo = 0;
        var iPageName = "";
        var sHTMLTag = ""
        var sHTMLAttributes = ""
        var sHTMLStyles = ""
        var sHTMLContent = ""
        var controlID = ""

        //do nowt for negative id columns.
        if (controlItemArray[0] < 0) return;

        var iPageNo = Number(controlItemArray[0]);
        var controlID = "FI_" + controlItemArray[2];
        //ColumnID used for controlvalues etc, not unique.
        var ColumnID = controlItemArray[2];
        var tabIndex = Number(controlItemArray[18]);

        //TODO: move styling to classes
        //TODO: can't use ID's as they may not be unique....

        var ControlType = Number(controlItemArray[3]);
        var fSelectOK = false;
        var fParentTableControl = false;
        var fControlEnabled = true;
        var fReadOnly = false;

        //Permissions. From activeX recordDMI.formatscreen function.
        if ($("#txtRecEditTableID").val() == controlItemArray[1]) {
            if (controlItemArray[2] > 0) { }
            fSelectOK = (Number(controlItemArray[47]) != 0);
            fParentTableControl = false;

            // Disable control if no permission is granted.
            fControlEnabled = (Number(controlItemArray[40]) == 0);  //database readonly value
            if ((ControlType == 8) || ((ControlType == 64) && (Number(controlItemArray[37]) != 0))) fControlEnabled = true;    //enable all multiline text, or OLEs


            if (fControlEnabled) {
                if ((ControlType == 64) && (Number(controlItemArray[23]) == 11)) {
                    //Date Control
                    fControlEnabled = (Number(controlItemArray[48] != 0));  // UpdateGranted property
                }
                else if (ControlType == 2048) {
                    //CommandButton
                    fControlEnabled = false;
                }
                else {
                    fControlEnabled = (Number(controlItemArray[48] != 0));  // UpdateGranted property

                    if ((ControlType == 64) && (Number(controlItemArray[37]) != 0) && ((Number(controlItemArray[23]) == 12) || (Number(controlItemArray[23]) == -1))) {
                        //if multiline text and (sqlVarchar or sqllongvarchar)
                        if ((!fControlEnabled) || (Number(controlItemArray[61]) != 0)) {
                            //if screen.readonly or disabled
                            fControlEnabled = true;
                            fReadOnly = true;
                        }
                    }

                }
            }
        }
        else {
            //Parent table control.
            fParentTableControl = true;
            if ((ControlType == 256) || (ControlType == 512) || (ControlType == 4) || (ControlType == 2 ^ 13) || (ControlType == 2 ^ 14) || (ControlType == 2 ^ 15)) {
                //label, frame, image, line, navigation or colourpicker
                fControlEnabled = false;
            }

            if ((ControlType == 64) && (Number(controlItemArray[37]) != 0) && ((Number(controlItemArray[23]) == 12) || (Number(controlItemArray[23]) == -1))) {
                //if multiline text and (sqlVarchar or sqllongvarchar)
                if ((!fControlEnabled) || (Number(controlItemArray[61]) != 0)) {
                    //if screen.readonly or disabled
                    fControlEnabled = true;
                    fReadOnly = true;
                }
            }
        }

        if (Number(controlItemArray[61]) != 0) {
            //Screen.Readonly
            fControlEnabled = false;
        }


        //Now add the controls to the form...

        switch (Number(controlItemArray[3])) {
            case 1: //checkbox
                //TODO: right-aligned checkboxes                
                var span = document.createElement('span');
                span.className = "checkbox left";
                applyLocation(span, controlItemArray);
                span.style.margin = "0px";
                span.style.textAlign = "left";
                span.style.display = "inline-block";

                var checkbox = span.appendChild(document.createElement('input'));
                checkbox.type = "checkbox";
                checkbox.id = controlID;
                checkbox.style.fontFamily = controlItemArray[11];
                checkbox.style.fontSize = controlItemArray[12] + 'pt';
                checkbox.style.position = "absolute";
                checkbox.style.top = "50%";
                checkbox.style.left = "0px";
                checkbox.style.padding = "0px";
                checkbox.style.margin = "-7px 0px 0px 0px";
                checkbox.style.textAlign = "left";

                var label = span.appendChild(document.createElement('label'));
                label.htmlFor = checkbox.id;
                label.appendChild(document.createTextNode(controlItemArray[8]));
                label.style.marginLeft = "18px";
                label.style.fontFamily = controlItemArray[11];
                label.style.fontSize = controlItemArray[12] + 'pt';


                checkbox.setAttribute("data-ColumnID", ColumnID);

                if (!fControlEnabled) span.disabled = true;

                //Add control to relevant tab, create if required.                
                addControl(iPageNo, span);

                break;
            case 2: //ctlCombo
                var selector = document.createElement('select');
                selector.id = controlID;
                applyLocation(selector, controlItemArray);
                selector.style.backgroundColor = "White";
                selector.style.color = "Black";
                selector.style.fontFamily = controlItemArray[11];
                selector.style.fontSize = controlItemArray[12] + 'pt';
                selector.style.borderWidth = "1px";
                selector.setAttribute("data-ColumnID", ColumnID);

                if (!fControlEnabled) selector.disabled = true;

                addControl(iPageNo, selector);

                var option = document.createElement('option');
                option.value = '0';
                option.appendChild(document.createTextNode(''));
                selector.appendChild(option);

                break;

            case 4, 1024: //Image/Photo
                var image = document.createElement('img');
                image.id = controlID;
                applyLocation(image, controlItemArray);
                image.style.border = "1px solid gray";
                image.style.padding = "0px";
                image.setAttribute("data-ColumnID", ColumnID);

                if (!fControlEnabled) image.disabled = true;

                //Add control to relevant tab, create if required.                
                addControl(iPageNo, image);

                break;
            case 8: //ctlOle
                var button = document.createElement('input');
                button.type = "button";
                button.value = "OLE";
                applyLocation(button, controlItemArray);
                button.style.padding = "0px";
                button.setAttribute("data-ColumnID", ColumnID);
                //button.disabled = false;    //always enabled
                addControl(iPageNo, button);

                break;
            case 16: //ctlRadio
                //TODO: set 'maxlength=.size' if fselectOK is true and not fparentcontrol
                //TODO: .disabled = (!fControlEnabled);
                break;
            case 32: //ctlSpinner
                var spinnerContainer = document.createElement('div');
                applyLocation(spinnerContainer, controlItemArray);
                spinnerContainer.style.padding = "0px";

                var spinner = spinnerContainer.appendChild(document.createElement("input"));
                spinner.className = "spinner";
                spinner.id = controlID;
                spinner.style.fontFamily = controlItemArray[11];
                spinner.style.fontSize = controlItemArray[12] + 'pt';
                spinner.style.width = (Number((controlItemArray[7]) / 15)) + "px";
                spinner.style.margin = "0px";
                spinner.setAttribute("data-ColumnID", ColumnID);

                if (!fControlEnabled) spinnerContainer.disabled = true;

                //Add control to relevant tab, create if required.                
                addControl(iPageNo, spinnerContainer);
                break;
            case 64: //ctlText

                if (Number(controlItemArray[37]) !== 0) {
                    //Multi-line textbox
                    var textbox = document.createElement('textarea');
                    //textbox.disabled = false;  //always enabled.
                }
                else {
                    var textbox = document.createElement('input');

                    switch (Number(controlItemArray[23])) {
                        case 11: //sqlDate
                            textbox.type = "text";
                            textbox.className = "datepicker";
                            break;
                        case 2, 4: //sqlNumeric, sqlInteger
                            textbox.type = "number";
                            break;
                        default:
                            textbox.type = "text";
                            textbox.isMultiLine = false;

                            if (controlItemArray[35].length > 0) {
                                //TODO: apply mask to control
                            }
                    }

                    if (!fControlEnabled) textbox.disabled = true;

                }

                textbox.id = controlID;
                applyLocation(textbox, controlItemArray);
                textbox.style.fontFamily = controlItemArray[11];
                textbox.style.fontSize = controlItemArray[12] + 'pt';
                textbox.style.padding = "0px";
                textbox.setAttribute("data-ColumnID", ColumnID);

                //Add control to relevant tab, create if required.                
                addControl(iPageNo, textbox);
                break;
            case 128: //ctlTab
                break;
            case 256: //Label
                var span = document.createElement('span');
                applyLocation(span, controlItemArray);
                span.style.backgroundColor = "White";
                span.style.color = "Black";
                span.style.fontFamily = controlItemArray[11];
                span.style.fontSize = controlItemArray[12] + 'pt';
                span.innerText = controlItemArray[8];

                //replaces the SetControlLevel function in recordDMI.ocx.
                span.style.zIndex = 0;

                if (!fControlEnabled) span.disabled = true;

                addControl(iPageNo, span);

                break;
            case 512: //Frame
                var fieldset = document.createElement('fieldset');
                applyLocation(fieldset, controlItemArray);
                fieldset.style.backgroundColor = "transparent";
                fieldset.style.color = "Black";
                fieldset.style.padding = "0px";

                var legend = fieldset.appendChild(document.createElement('legend'));
                legend.appendChild(document.createTextNode(controlItemArray[8]));

                addControl(iPageNo, fieldset);

                break;
                //case 1024: //ctlPhoto - see case 4.            
            case 2048: //ctlCommand
                break;
            case 4096: //ctlWorking Pattern
                //TODO: .disabled = (!fControlEnabled);
                break;
            case 2 ^ 13: //ctlLine
                break;
            case 2 ^ 14: //ctlNavigation
                //TODO: Nav control always .disabled = false.
                break;
            case 2 ^ 15: //ctlColourPicker
                //TODO: .disabled = (!fControlEnabled);
                break;
            default:
                break;
        }
    }

    function addHTMLControlValues(controlValues) {
        var controlValuesArray = controlValues.split("\t");
        var fDoneFirstValue = false;
        var lngColumnID = 0;
        var sValue = "";

        for (var i = 0; i < controlValuesArray.length; i++) {

            sValue = controlValuesArray[i];

            if (lngColumnID > 0) {
                if (sValue.length > 0) {

                    //get the column type, then add this value to it/them.
                    $("#ctlRecordEdit").find("[data-ColumnID='" + lngColumnID + "']").each(function () {
                        //TODO: Option Groups
                        if ($(this).is("select")) {
                            var option = document.createElement('option');
                            option.value = i + 1;
                            option.appendChild(document.createTextNode(sValue));
                            $(this).append(option);
                        }
                    });
                }
            }
            else {
                if (sValue.length > 0) {
                    //set the column ID to apply list to.
                    lngColumnID = Number(sValue);
                }
                else { return false; }
            }
        }
    }



</script>

<script type="text/javascript">
    function addActiveXHandlers() {
        //TODO: NPG
        return false;
        OpenHR.addActiveXHandler("ctlRecordEdit", "dataChanged", ctlRecordEdit_dataChanged);
        OpenHR.addActiveXHandler("ctlRecordEdit", "ToolClickRequest", ctlRecordEdit_ToolClickRequest);
        OpenHR.addActiveXHandler("ctlRecordEdit", "LinkButtonClick", ctlRecordEdit_LinkButtonClick);
        OpenHR.addActiveXHandler("ctlRecordEdit", "LookupClick", ctlRecordEdit_LookupClick);
        OpenHR.addActiveXHandler("ctlRecordEdit", "ImageClick4", ctlRecordEdit_ImageClick4);
        OpenHR.addActiveXHandler("ctlRecordEdit", "OLEClick4", ctlRecordEdit_OLEClick4);
    }
</script>


<SCRIPT type="text/javascript">
    function ctlRecordEdit_dataChanged() {
        // The data in the recEdit control has changed so refresh the menu.
        // Get menu.asp to refresh the menu.
        menu_refreshMenu();
    }

    function ctlRecordEdit_ToolClickRequest(lngIndex, strTool) {
        // The data in the recEdit control has changed so refresh the menu.
        // Get menu.asp to refresh the menu.
        menu_MenuClick(strTool);
    }

    function ctlRecordEdit_LinkButtonClick(plngLinkTableID, plngLinkOrderID, plngLinkViewID, plngLinkRecordID) {
        // A link button has been pressed in the recEdit control,
        // so open the link option page.
        menu_loadLinkPage(plngLinkTableID, plngLinkOrderID, plngLinkViewID, plngLinkRecordID);
    }

    function ctlRecordEdit_LookupClick(plngColumnID, plngLookupColumnID, psLookupValue, pfMandatory, pstrFilterValue) {
        // A lookup button has been pressed in the recEdit control,
        // so open the lookup page.
        menu_loadLookupPage(plngColumnID, plngLookupColumnID, psLookupValue, pfMandatory, pstrFilterValue);
    }

    function ctlRecordEdit_ImageClick4(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize, pbIsReadOnly) {
        // An image has been pressed in the recEdit control,
        // so open the image find page.
        var fOK;

        fOK = true;
        if (frmRecordEditForm.ctlRecordEdit.recordID == 0) {
            OpenHR.messageBox("Unable to edit photo fields until the record has been saved.");
            fOK = false;
        }

        if (fOK == true) {
            //TODO Client DLL stuff
            //    if (plngOLEType < 2) {
            //        fOK = window.parent.frames("menuframe").ASRIntranetFunctions.ValidateDir(frmRecordEditForm.txtPicturePath.value);
            //        if (fOK == true)
            //            window.parent.frames("menuframe").loadImagePage(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize);
            //        else
            //            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit photo fields as the photo path is not valid.");
            //    } else {
            //        window.parent.frames("menuframe").loadImagePage(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize);
            //    }
        }
    }

    function ctlRecordEdit_OLEClick4(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly) {
        // An OLE button has been pressed in the recEdit control,
        // so open the OLE page.	
        var fOK;
        var sKey = new String('');

        fOK = true;
        if (frmRecordEditForm.ctlRecordEdit.recordID == 0) {
            OpenHR.messageBox("Unable to edit OLE fields until the record has been saved.");
            fOK = false;
        }

        //TODO: Client DLL stuff
        //if (fOK == true)
        //{
        //    // Server OLE
        //    if (plngOLEType == 1) {
        //        fOK = window.parent.frames("menuframe").ASRIntranetFunctions.ValidateDir(frmRecordEditForm.txtOLEServerPath.value);
        //        if (fOK == true)
        //            window.parent.frames("menuframe").loadOLEPage(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly);
        //        else
        //            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit server OLE fields as the OLE (Server) path is not valid.");
        //    }

        //        // Local OLE
        //    else if (plngOLEType == 0) {
        //        fOK = window.parent.frames("menuframe").ASRIntranetFunctions.ValidateDir(frmRecordEditForm.txtOLELocalPath.value);
        //        if (fOK == true)
        //            window.parent.frames("menuframe").loadOLEPage(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly);
        //        else
        //            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit local OLE fields as the OLE (Local) path is not valid.");
        //    }

        //        // Embedded OLE
        //    else if (plngOLEType == 2) {
        //        sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
        //        window.parent.frames("menuframe").loadOLEPage(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly);
        //    }

        //        // Linked OLE
        //    else if (plngOLEType == 3) {
        //        sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);
        //        window.parent.frames("menuframe").loadOLEPage(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly);			
        //    }
        //}	        
    }



    function recordEdit_refreshData() {
        // Get the data.asp to get the required data.
        var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
        var frmRecordEditForm = OpenHR.getForm("workframe", "frmRecordEditForm");

        frmGetDataForm.txtAction.value = "LOAD";
        frmGetDataForm.txtReaction.value = "";
        frmGetDataForm.txtCurrentTableID.value = frmRecordEditForm.txtCurrentTableID.value;
        frmGetDataForm.txtCurrentScreenID.value = frmRecordEditForm.txtCurrentScreenID.value;
        frmGetDataForm.txtCurrentViewID.value = frmRecordEditForm.txtCurrentViewID.value;
        frmGetDataForm.txtSelectSQL.value = frmRecordEditForm.txtRecEditSelectSQL.value;
        frmGetDataForm.txtFromDef.value = frmRecordEditForm.txtRecEditFromDef.value;
        frmGetDataForm.txtFilterSQL.value = frmRecordEditForm.txtRecEditFilterSQL.value;
        frmGetDataForm.txtFilterDef.value = frmRecordEditForm.txtRecEditFilterDef.value;
        frmGetDataForm.txtRealSource.value = frmRecordEditForm.txtRecEditRealSource.value;
        frmGetDataForm.txtRecordID.value = OpenHR.getForm("dataframe", "frmData").txtRecordID.value;
        frmGetDataForm.txtParentTableID.value = frmRecordEditForm.txtCurrentParentTableID.value;
        frmGetDataForm.txtParentRecordID.value = frmRecordEditForm.txtCurrentParentRecordID.value;
        //TODO frmGetDataForm.txtDefaultCalcCols.value = frmRecordEditForm.ctlRecordEdit.CalculatedDefaultColumns();
        frmGetDataForm.txtInsertUpdateDef.value = "";
        frmGetDataForm.txtTimestamp.value = "";

        data_refreshData();
    }


    function setRecordID(plngRecordID) {
        frmRecordEditForm.txtCurrentRecordID.value = plngRecordID;
        frmRecordEditForm.ctlRecordEdit.recordID = plngRecordID;
    }

    function setCopiedRecordID(plngRecordID) {
        frmRecordEditForm.ctlRecordEdit.CopiedRecordID = plngRecordID;
    }

    function setParentTableID(plngParentTableID) {
        frmRecordEditForm.txtCurrentParentTableID.value = plngParentTableID;
        frmRecordEditForm.ctlRecordEdit.ParentTableID = plngParentTableID;
    }

    function setParentRecordID(plngParentRecordID) {
        frmRecordEditForm.txtCurrentParentRecordID.value = plngParentRecordID;
        frmRecordEditForm.ctlRecordEdit.ParentRecordID = plngParentRecordID;
    }

</script>

<div <%=session("BodyTag")%>>
<FORM action="" method=post id=frmRecordEditForm name=frmRecordEditForm>


<%
	on error resume next
	
    Dim sErrorDescription As String
	sErrorDescription = ""

	' Get the page title.
    Dim cmdRecEditWindowTitle = CreateObject("ADODB.Command")
	cmdRecEditWindowTitle.CommandText = "sp_ASRIntGetRecordEditInfo"
	cmdRecEditWindowTitle.CommandType = 4 ' Stored Procedure
    cmdRecEditWindowTitle.ActiveConnection = Session("databaseConnection")

    Dim prmTitle = cmdRecEditWindowTitle.CreateParameter("title", 200, 2, 100)
    cmdRecEditWindowTitle.Parameters.Append(prmTitle)

    Dim prmQuickEntry = cmdRecEditWindowTitle.CreateParameter("quickEntry", 11, 2) ' 11=bit, 2=output
    cmdRecEditWindowTitle.Parameters.Append(prmQuickEntry)

    Dim prmScreenID = cmdRecEditWindowTitle.CreateParameter("screenID", 3, 1)
    cmdRecEditWindowTitle.Parameters.Append(prmScreenID)
	prmScreenID.value = cleanNumeric(session("screenID"))

    Dim prmViewID = cmdRecEditWindowTitle.CreateParameter("viewID", 3, 1)
    cmdRecEditWindowTitle.Parameters.Append(prmViewID)
	prmViewID.value = cleanNumeric(session("viewID"))

    Err.Clear()
    cmdRecEditWindowTitle.Execute
  
    If (Err.Number <> 0) Then
        sErrorDescription = "The page title could not be created." & vbCrLf & FormatError(Err.Description)
    End If

	if len(sErrorDescription) = 0 then		  
        'Response.Write(Replace(cmdRecEditWindowTitle.Parameters("title").Value, "_", " ") & " - No activeX" & vbCrLf)        
        Response.Write("<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & cmdRecEditWindowTitle.Parameters("quickEntry").Value & ">" & vbCrLf)
    End If
		

%>

<h3 class="pageTitle"><%Response.Write(Replace(cmdRecEditWindowTitle.Parameters("title").Value, "_", " ") & " - No activeX" & vbCrLf)
    ' Release the ADO command object.
    cmdRecEditWindowTitle = Nothing%></h3>
    
<div id="ctlRecordEdit" style="margin:0 auto;">
    <ul id="tabHeaders">        
    </ul>
</div>

<%
    Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtCurrentTableID name=txtCurrentTableID value=" & Session("tableID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtCurrentViewID name=txtCurrentViewID value=" & Session("viewID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtCurrentScreenID name=txtCurrentScreenID value=" & Session("screenID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtCurrentOrderID name=txtCurrentOrderID value=" & Session("orderID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtCurrentRecordID name=txtCurrentRecordID value=" & Session("recordID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtCurrentParentTableID name=txtCurrentParentTableID value=" & Session("parentTableID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtCurrentParentRecordID name=txtCurrentParentRecordID value=" & Session("parentRecordID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtLineage name=txtLineage value=" & Session("lineage") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtCurrentRecPos name=txtCurrentRecPos value=" & Session("parentRecordID") & ">" & vbCrLf)
	
	if len(sErrorDescription) = 0 then
		' Read the screen definition from the database into 'hidden' controls.
        Dim cmdRecEditDefinition = CreateObject("ADODB.Command")
		cmdRecEditDefinition.CommandText = "sp_ASRIntGetScreenDefinition"
		cmdRecEditDefinition.CommandType = 4 ' Stored Procedure
        cmdRecEditDefinition.ActiveConnection = Session("databaseConnection")

        prmScreenID = cmdRecEditDefinition.CreateParameter("screenID", 3, 1) ' 3=integer, 1=input
        cmdRecEditDefinition.Parameters.Append(prmScreenID)
		prmScreenID.value = cleanNumeric(session("screenID"))

        prmViewID = cmdRecEditDefinition.CreateParameter("viewID", 3, 1) ' 3=integer, 1=input
        cmdRecEditDefinition.Parameters.Append(prmViewID)
		prmViewID.value = cleanNumeric(session("viewID"))

        Err.Clear()
        Dim rstScreenDefinition = cmdRecEditDefinition.Execute
	  
        If (Err.Number <> 0) Then
            sErrorDescription = "The screen definition could not be read." & vbCrLf & FormatError(Err.Description)
        End If

		if len(sErrorDescription) = 0 then		  
            Response.Write("<INPUT type='hidden' id=txtRecEditTableID name=txtRecEditTableID value=" & Session("tableID") & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditViewID name=txtRecEditViewID value=" & Session("viewID") & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditHeight name=txtRecEditHeight value=" & rstScreenDefinition.Fields("height").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditWidth name=txtRecEditWidth value=" & rstScreenDefinition.Fields("width").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditTabCount name=txtRecEditTabCount value=" & rstScreenDefinition.Fields("tabCount").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditTabCaptions name=txtRecEditTabCaptions value=""" & Replace(Replace(rstScreenDefinition.Fields("tabCaptions").Value, "&", "&&"), """", "&quot;") & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditFontName name=txtRecEditFontName value=""" & Replace(rstScreenDefinition.Fields("fontName").Value, """", "&quot;") & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditFontSize name=txtRecEditFontSize value=" & rstScreenDefinition.Fields("fontSize").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditFontBold name=txtRecEditFontBold value=" & rstScreenDefinition.Fields("fontBold").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditFontItalic name=txtRecEditFontItalic value=" & rstScreenDefinition.Fields("fontItalic").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditFontUnderline name=txtRecEditFontUnderline value=" & rstScreenDefinition.Fields("fontUnderline").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditFontStrikethru name=txtRecEditFontStrikethru value=" & rstScreenDefinition.Fields("fontStrikethru").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditRealSource name=txtRecEditRealSource value=""" & Replace(rstScreenDefinition.Fields("realSource").Value, """", "&quot;") & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditInsertGranted name=txtRecEditInsertGranted value=" & rstScreenDefinition.Fields("insertGranted").Value & ">" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditDeleteGranted name=txtRecEditDeleteGranted value=" & rstScreenDefinition.Fields("deleteGranted").Value & ">" & vbCrLf)
        End If
		
		rstScreenDefinition.close
        rstScreenDefinition = Nothing
		
		' Release the ADO command object.
        cmdRecEditDefinition = Nothing
	end if
	
    Response.Write("<INPUT type='hidden' id=txtRecEditEmpTableID name=txtRecEditEmpTableID value=" & Session("TB_EmpTableID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtRecEditCourseTableID name=txtRecEditCourseTableID value=" & Session("TB_CourseTableID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtRecEditTBTableID name=txtRecEditTBTableID value=" & Session("TB_TBTableID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtRecEditTBStatusColumnID name=txtRecEditTBStatusColumnID value=" & Session("TB_TBStatusColumnID") & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtRecEditCourseCancelDateColumnID name=txtRecEditCourseCancelDateColumnID value=" & Session("TB_CourseCancelDateColumnID") & ">" & vbCrLf)
    'ND commented out for now - Response.Write "<INPUT type='hidden' id=txtWaitListOverRideColumnID name=txtWaitListOverRideColumnID value=" & session("TB_WaitListOverRideColumnID") & ">" & vbcrlf

	if len(sErrorDescription) = 0 then
		' Get the screen controls
        Dim cmdRecEditControls = CreateObject("ADODB.Command")
		cmdRecEditControls.CommandText = "sp_ASRIntGetScreenControlsString2"
		cmdRecEditControls.CommandType = 4 ' Stored Procedure
        cmdRecEditControls.ActiveConnection = Session("databaseConnection")

        prmScreenID = cmdRecEditControls.CreateParameter("screenID", 3, 1) ' 3=integer, 1=input
        cmdRecEditControls.Parameters.Append(prmScreenID)
		prmScreenID.value = cleanNumeric(session("screenID"))

        prmViewID = cmdRecEditControls.CreateParameter("viewID", 3, 1) ' 3=integer, 1=input
        cmdRecEditControls.Parameters.Append(prmViewID)
		prmViewID.value = cleanNumeric(session("viewID"))

        Dim prmSelectSQL = cmdRecEditControls.CreateParameter("selectSQL", 200, 2, 2147483646) ' 200=varchar, 2=output
        cmdRecEditControls.Parameters.Append(prmSelectSQL)

        Dim prmFromDef = cmdRecEditControls.CreateParameter("fromDef", 200, 2, 255) ' 200=varchar, 2=output
        cmdRecEditControls.Parameters.Append(prmFromDef)

        Dim prmOrderID = cmdRecEditControls.CreateParameter("orderID", 3, 3) ' 3=integer,  3=input/output
        cmdRecEditControls.Parameters.Append(prmOrderID)
		prmOrderID.value = cleanNumeric(session("orderID"))

        Err.Clear()
        Dim rstScreenControls = cmdRecEditControls.Execute
	  
        If (Err.Number <> 0) Then
            sErrorDescription = "The screen control definitions could not be read." & vbCrLf & FormatError(Err.Description)
        End If

		if len(sErrorDescription) = 0 then		  
            Dim iloop = 1
			do while not rstScreenControls.EOF
                Response.Write("<INPUT type='hidden' id=txtRecEditControl_" & iloop & " name=txtRecEditControl_" & iloop & " value=""" & Replace(rstScreenControls.Fields("controlDefinition").Value, """", "&quot;") & """>" & vbCrLf)
                rstScreenControls.MoveNext()
	
				iloop = iloop + 1
			loop

			' Release the ADO recordset object.
			rstScreenControls.close
            rstScreenControls = Nothing
		
			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
            Response.Write("<INPUT type='hidden' id=txtRecEditSelectSQL name=txtRecEditSelectSQL value=""" & Replace(Replace(cmdRecEditControls.Parameters("selectSQL").Value, "'", "'''"), """", "&quot;") & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditFromDef name=txtRecEditFromDef value=""" & Replace(Replace(cmdRecEditControls.Parameters("fromDef").Value, "'", "'''"), """", "&quot;") & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtRecEditOrderID name=txtRecEditOrderID value=" & cmdRecEditControls.Parameters("orderID").Value & ">" & vbCrLf)
        End If
		
        cmdRecEditControls = Nothing
	end if
	
	if len(sErrorDescription) = 0 then
		' Get the screen column control values
        Dim cmdRecEditControlValues = CreateObject("ADODB.Command")
		cmdRecEditControlValues.CommandText = "sp_ASRIntGetScreenControlValuesString"
		cmdRecEditControlValues.CommandType = 4 ' Stored Procedure
        cmdRecEditControlValues.ActiveConnection = Session("databaseConnection")

        prmScreenID = cmdRecEditControlValues.CreateParameter("screenID", 3, 1)
        cmdRecEditControlValues.Parameters.Append(prmScreenID)
		prmScreenID.value = cleanNumeric(session("screenID"))

        Err.Clear()
        Dim rstScreenControlValues = cmdRecEditControlValues.Execute
		
        If (Err.Number <> 0) Then
            sErrorDescription = "The screen control values could not be read." & vbCrLf & FormatError(Err.Description)
        End If

		if len(sErrorDescription) = 0 then		  
            Dim iloop = 1
			do while not rstScreenControlValues.EOF
                Response.Write("<INPUT type='hidden' id=txtRecEditControlValues_" & iloop & " name=txtRecEditControlValues_" & iloop & " value=""" & Replace(rstScreenControlValues.Fields("valueDefinition").Value, """", "&quot;") & """>" & vbCrLf)
                rstScreenControlValues.MoveNext()
		
				iloop = iloop + 1
			loop

			' Release the ADO recordset object.
			rstScreenControlValues.close
            rstScreenControlValues = Nothing
		end if
	
        cmdRecEditControlValues = Nothing
	end if

    Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
    Response.Write("<INPUT type='hidden' id=txtRecEditFilterDef name=txtRecEditFilterDef value=""" & Replace(Session("filterDef"), """", "&quot;") & """>" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtRecEditFilterSQL name=txtRecEditFilterSQL value=""" & Replace(Session("filterSQL"), """", "&quot;") & """>" & vbCrLf)

	' JPD 20021021 - Added pictures functionlity.
	' JPD 20021127 - Moved Utilities object into session variable.
	'Set objUtilities = CreateObject("COAIntServer.Utilities")
	'objUtilities.Connection = session("databaseConnection")
    Dim objUtilities = Session("UtilitiesObject")
    Dim sTempPath = Server.MapPath("pictures")
    Dim picturesArray = objUtilities.GetPictures(Session("screenID"), CStr(sTempPath))

	for iCount = 1 to UBound(picturesArray,2)
        Response.Write("<INPUT type='hidden' id=txtRecEditPicture_" & picturesArray(1, iCount) & " name=txtRecEditPicture_" & picturesArray(1, iCount) & " value=""" & picturesArray(2, iCount) & """>" & vbCrLf)
    Next
    objUtilities = Nothing

	'sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	'iIndex = inStrRev(sReferringPage, "/")
	'if iIndex > 0 then
	'	sReferringPage = left(sReferringPage, iIndex - 1)
	'	if left(sReferringPage, 5) = "http:" then
	'		sReferringPage = mid(sReferringPage, 6)
	'	end if
	'end if
	'Response.Write "<INPUT type='hidden' id=txtImagePath name=txtImagePath value=""" & sReferringPage & """>" & vbcrlf
%>

	<INPUT type='hidden' id=txtPicturePath name=txtPicturePath>
	<INPUT type='hidden' id=txtImagePath name=txtImagePath>
	<INPUT type='hidden' id=txtOLEServerPath name=txtOLEServerPath>
	<INPUT type='hidden' id=txtOLELocalPath name=txtOLELocalPath>
</FORM>

<FORM action="default_Submit" method=post id=frmGoto name=frmGoto>
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</FORM>

</div>


<script type="text/javascript">
    addActiveXHandlers();
    recordEdit_window_onload();
</script>

<% 
    'function formatError(psErrMsg)
    '  Dim iStart 
    '  dim iFound 
  
    '  iFound = 0
    '  Do
    '    iStart = iFound
    '    iFound = InStr(iStart + 1, psErrMsg, "]")
    '  Loop While iFound > 0
  
    '  If (iStart > 0) And (iStart < Len(Trim(psErrMsg))) Then
    '    formatError = Trim(Mid(psErrMsg, iStart + 1))
    '  Else
    '    formatError = psErrMsg
    '  End If
    'end function
%>
