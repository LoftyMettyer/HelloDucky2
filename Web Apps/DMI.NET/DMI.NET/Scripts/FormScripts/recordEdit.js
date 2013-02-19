
//functions that replicate COAIntRecordDMI.ocx

function addControl(tabNumber, controlDef) {

    var tabID = "FI_21_" + tabNumber;

    if (($("#" + tabID).length <= 0) && (tabNumber > 0)) {
        //tab doesn't exist - create it...
        var tabFontName = $("#txtRecEditFontName").val();
        var tabFontSize = $("#txtRecEditFontSize ").val();

        var tabCss = "style='font-family: " + tabFontName + " ; font-size: " + tabFontSize + "pt'";
        
        var tabs = $("#ctlRecordEdit").tabs(),
    tabTemplate = "<li><a " + tabCss + " href='#{href}'>#{label}</a></li>";

        var label = getTabCaption(tabNumber),
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


function applyLocation(formItem, controlItemArray, bordered) {
    //reduce all sizes by 2 for border.
    var borderwidth;
    borderwidth = (bordered ? 2 : -2);
    formItem.style.top = (Number(controlItemArray[4]) / 15) + "px";
    formItem.style.left = (Number(controlItemArray[5]) / 15) + "px";
    formItem.style.height = (Number((controlItemArray[6]) / 15) - borderwidth) + "px";
    formItem.style.width = (Number((controlItemArray[7]) / 15) - borderwidth) + "px";
    formItem.style.position = "absolute";
}

// -------------------------------------------------- Add the record Edit controls ------------------------------------------------
function AddHtmlControl(controlItem, txtcontrolID, key) {
    var controlItemArray = controlItem.split("\t");
    var iPageNo = 0;
    var controlID = "";
    var tmpNum = txtcontrolID.indexOf("_");
    txtcontrolID = txtcontrolID.substring(tmpNum);

    if (controlItemArray[0] < 0) {
        //The definition is for an id column.            
        var nextAvail;
        if (this.mavIDColumns.length <= 0) {
            nextAvail = 0;
        }
        else {
            nextAvail = this.mavIDColumns.length / 3;
        }

        this.mavIDColumns[nextAvail] = new Array(3);

        this.mavIDColumns[nextAvail][1] = Number(controlItemArray[1]);   //ColumnID
        this.mavIDColumns[nextAvail][2] = controlItemArray[2];   //Column Name
        this.mavIDColumns[nextAvail][3] = 0;   //Value

    }
    
    //-------------------------------------------------Get permissions for this control first -----------------------------------------------------------------
    var controlType = Number(controlItemArray[3]);
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
        if ((controlType == 8) || ((controlType == 64) && (Number(controlItemArray[37]) != 0))) fControlEnabled = true;    //enable all multiline text, or OLEs


        if (fControlEnabled) {
            if ((controlType == 64) && (Number(controlItemArray[23]) == 11)) {
                //Date Control
                fControlEnabled = (Number(controlItemArray[48] != 0));  // UpdateGranted property
            }
            else if (controlType == 2048) {
                //CommandButton
                fControlEnabled = false;
            }
            else {
                fControlEnabled = (Number(controlItemArray[48] != 0));  // UpdateGranted property

                if ((controlType == 64) && (Number(controlItemArray[37]) != 0) && ((Number(controlItemArray[23]) == 12) || (Number(controlItemArray[23]) == -1))) {
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
        if ((controlType == 256) || (controlType == 512) || (controlType == 4) || (controlType == 2 ^ 13) || (controlType == 2 ^ 14) || (controlType == 2 ^ 15)) {
            //label, frame, image, line, navigation or colourpicker
            fControlEnabled = false;
        }

        if ((controlType == 64) && (Number(controlItemArray[37]) != 0) && ((Number(controlItemArray[23]) == 12) || (Number(controlItemArray[23]) == -1))) {
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


    //----------------------------------------------------------------------- Now add the control to the form... ---------------------------------------------------------------
    
    iPageNo = Number(controlItemArray[0]);
    controlID = "FI_" + controlItemArray[2] + txtcontrolID; //ColumnID used for controlvalues etc, not unique.
    var columnID = controlItemArray[2];
    var tabIndex = Number(controlItemArray[18]);

    //TODO: move styling to classes?    
    //TODO: move duplicated property setting blocks to separate functions
    
    var span;
    var top;
    var left;
    var height;
    var width;
    var borderCss;
    var radioTop;
    

    switch (Number(controlItemArray[3])) {
        case 1: //checkbox
            span = document.createElement('span');
            
            applyLocation(span, controlItemArray, true);
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
            
            checkbox.style.padding = "0px";
            checkbox.style.margin = "-7px 0px 0px 0px";
            checkbox.style.textAlign = "left";

            var label = span.appendChild(document.createElement('label'));
            label.htmlFor = checkbox.id;
            label.appendChild(document.createTextNode(controlItemArray[8]));
            
            label.style.fontFamily = controlItemArray[11];
            label.style.fontSize = controlItemArray[12] + 'pt';

            //align left or right...
            if (controlItemArray[20] != "0") {            
                //right align
                span.className = "checkbox right";
                checkbox.style.right = "0px";
            } else {
                //left align
                span.className = "checkbox left";
                checkbox.style.left = "0px";
                label.style.marginLeft = "18px";
            }

            if(tabIndex > 0) checkbox.tabindex = tabIndex;

            checkbox.setAttribute("data-columnID", columnID);
            checkbox.setAttribute("data-control-tag", key);

            if (!fControlEnabled) span.disabled = true;

            //Add control to relevant tab, create if required.                
            addControl(iPageNo, span);

            break;
        case 2: //ctlCombo
            var selector = document.createElement('select');
            selector.id = controlID;
            applyLocation(selector, controlItemArray, true);
            selector.style.backgroundColor = "White";
            selector.style.color = "Black";
            selector.style.fontFamily = controlItemArray[11];
            selector.style.fontSize = controlItemArray[12] + 'pt';
            selector.style.borderWidth = "1px";
            selector.setAttribute("data-columnID", columnID);
            selector.setAttribute("data-control-key", key);
            
            if (!fControlEnabled) selector.disabled = true;

            if (tabIndex > 0) selector.tabindex = tabIndex;

            addControl(iPageNo, selector);

            //var option = document.createElement('option');
            //option.value = '0';
            //option.appendChild(document.createTextNode(''));
            //selector.appendChild(option);

            break;

        case 4: //Image
            var image = document.createElement('img');
            image.id = controlID;
            applyLocation(image, controlItemArray, true);
            image.style.border = "1px solid gray";
            image.style.padding = "0px";
            image.setAttribute("data-columnID", columnID);
            image.setAttribute("data-control-key", key);
            
            if (!fControlEnabled) image.disabled = true;

            var path = window.ROOT + 'Home/ShowImageFromDb?imageID=' + controlItemArray[50];
            
            image.setAttribute('src', path);            

            //Add control to relevant tab, create if required.                
            addControl(iPageNo, image);

            break;
        case 8: //ctlOle
            var button = document.createElement('input');
            button.type = "button";
            button.value = "OLE";
            applyLocation(button, controlItemArray, true);
            button.style.padding = "0px";
            button.setAttribute("data-columnID", columnID);
            button.setAttribute("data-control-key", key);
            
            if (tabIndex > 0) button.tabindex = tabIndex;

            //button.disabled = false;    //always enabled
            addControl(iPageNo, button);

            break;
        case 16: //ctlRadio
            //TODO: set 'maxlength=.size' if fselectOK is true and not fparentcontrol
            //TODO: .disabled = (!fControlEnabled);
            top = (Number(controlItemArray[4]) / 15);
            left = (Number(controlItemArray[5]) / 15);
            height = (Number((controlItemArray[6]) / 15) - 2);
            width = (Number((controlItemArray[7]) / 15) - 2);

            var cssBorderStyle = new Object();

            if (controlItemArray[19] == "0") {
                //pictureborder?
                cssBorderStyle.width = "0px";
                cssBorderStyle.style = "none";
                cssBorderStyle.color = "transparent";
                radioTop = 2;
                width += 2;
            } else {
                cssBorderStyle.width = "1px";
                cssBorderStyle.style = "solid";
                cssBorderStyle.color = "#999";
                width -= 2;
                height -= 2;
                
                //TODO ??  fontadjustment?

                radioTop = 19 + Number((controlItemArray[12] - 8) * 1.375);
                
                //TODO - android browser/tablet adjustment
            }
            
            fieldset = document.createElement("fieldset");
            fieldset.style.position = "absolute";
            fieldset.style.top = top + "px";
            fieldset.style.left = left + "px";
            fieldset.style.width = width + "px";
            fieldset.style.height = height + "px";
            fieldset.style.padding = "0px";
            fieldset.style.borderWidth = cssBorderStyle.width;
            fieldset.style.borderStyle = cssBorderStyle.style;
            fieldset.style.borderColor = cssBorderStyle.color;
            //appply font at fieldset level; it cascades.
            fieldset.style.fontFamily = controlItemArray[11];
            fieldset.style.fontSize = controlItemArray[12] + 'pt';
            fieldset.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
            fieldset.id = controlID;
            fieldset.setAttribute("data-datatype", "Option Group");
            fieldset.setAttribute("data-columnID", columnID);
            fieldset.setAttribute("data-alignment", controlItemArray[20]);
            fieldset.setAttribute("data-control-key", key);
            
            if ((controlItemArray[19] != "0") && (controlItemArray[8].length > 0)) {
                //has a border and a caption
                legend = fieldset.appendChild(document.createElement('legend'));
                legend.style.fontFamily = controlItemArray[11];
                legend.style.fontSize = controlItemArray[12] + 'pt';
                legend.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
                legend.appendChild(document.createTextNode(controlItemArray[8]));
            }

            
            //No Option Group buttons - these are added as values next.

            addControl(iPageNo, fieldset);           


            break;
        case 32: //ctlSpinner
            var spinnerContainer = document.createElement('div');
            applyLocation(spinnerContainer, controlItemArray, true);
            spinnerContainer.style.padding = "0px";

            var spinner = spinnerContainer.appendChild(document.createElement("input"));
            spinner.className = "spinner";
            spinner.id = controlID;
            spinner.style.fontFamily = controlItemArray[11];
            spinner.style.fontSize = controlItemArray[12] + 'pt';
            spinner.style.width = (Number((controlItemArray[7]) / 15)) + "px";
            spinner.style.margin = "0px";
            spinner.setAttribute("data-columnID", columnID);
            spinner.setAttribute("data-control-key", key);
            
            if (tabIndex > 0) spinner.tabindex = tabIndex;
            if (!fControlEnabled) spinnerContainer.disabled = true;

            //Add control to relevant tab, create if required.                
            addControl(iPageNo, spinnerContainer);
            break;
        case 64: //ctlText
            var textbox;
            if (Number(controlItemArray[37]) !== 0) {
                //Multi-line textbox
                textbox = document.createElement('textarea'); //textbox.disabled = false;  //always enabled.
            } else {
                textbox = document.createElement('input');
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
            applyLocation(textbox, controlItemArray, true);
            textbox.style.fontFamily = controlItemArray[11];
            textbox.style.fontSize = controlItemArray[12] + 'pt';
            textbox.style.padding = "0px";
            textbox.setAttribute("data-columnID", columnID);
            textbox.setAttribute("data-control-key", key);
            
            if (tabIndex > 0) textbox.tabindex = tabIndex;
            
            //Add control to relevant tab, create if required.                
            addControl(iPageNo, textbox);
            break;
        case 128: //ctlTab
            break;
        case 256: //Label
            span = document.createElement('span');
            applyLocation(span, controlItemArray, false);
            span.style.backgroundColor = "White";
            span.style.color = "Black";
            span.style.fontFamily = controlItemArray[11];
            span.style.fontSize = controlItemArray[12] + 'pt';
            span.innerText = controlItemArray[8];

            span.setAttribute("data-control-key", key);
            
            //replaces the SetControlLevel function in recordDMI.ocx.
            span.style.zIndex = 0;

            if (!fControlEnabled) span.disabled = true;

            addControl(iPageNo, span);

            break;
        case 512: //Frame
            var fieldset = document.createElement('fieldset');
            applyLocation(fieldset, controlItemArray, true);
            fieldset.style.backgroundColor = "transparent";
            fieldset.style.color = "Black";
            fieldset.style.padding = "0px";
            
            var legend = fieldset.appendChild(document.createElement('legend'));
            legend.style.fontFamily = controlItemArray[11];
            legend.style.fontSize = controlItemArray[12] + 'pt';
            legend.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
            legend.style.textDecoration = (Number(controlItemArray[16]) != 0) ? "underline" : "none";
            legend.appendChild(document.createTextNode(controlItemArray[8]));

            fieldset.setAttribute("data-control-key", key);
            
            addControl(iPageNo, fieldset);

            break;
        case 1024: //ctlPhoto
            


        case 2048: //ctlCommand
            break;
        case 4096: //ctlWorking Pattern
            //TODO: Font size change - this control is fixed in size.
            top = (Number(controlItemArray[4]) / 15);
            left = (Number(controlItemArray[5]) / 15);
            height = 58; //(Number((controlItemArray[6]) / 15) - 2);
            width = 125; //(Number((controlItemArray[7]) / 15) - 2);
            if (controlItemArray[19] == "0") {
                //pictureborder?
                borderCss = "border-style: none;";
            } else {
                borderCss = "border: 1px solid #999;";
                width -= 2;
                height -= 2;

                //TODO ??  fontadjustment?

                //TODO - android browser/tablet adjustment
            }

            fieldset = document.createElement("fieldset");
            fieldset.id = controlID;
            fieldset.setAttribute("data-columnID", columnID);
            fieldset.setAttribute("data-datatype", "Working Pattern");
            fieldset.setAttribute("data-control-key", key);
            
            fieldset.style.position = "absolute";
            fieldset.style.top = top + "px";
            fieldset.style.left = left + "px";
            fieldset.style.width = width + "px";
            fieldset.style.height = height + "px";
            fieldset.style.padding = "0px";
            fieldset.style.border = borderCss;

            for (var i = 0; i < 7; i++) {
                var offsetLeft = 26 + (i * 13);
                var dayLabel = fieldset.appendChild(document.createElement("span"));
                switch (i) {
                    case 0:
                        dayLabel.innerText = "S";
                        break;
                    case 1:
                        dayLabel.innerText = "M";
                        break;
                    case 2:
                        dayLabel.innerText = "T";
                        break;
                    case 3:
                        dayLabel.innerText = "W";
                        break;
                    case 4:
                        dayLabel.innerText = "T";
                        break;
                    case 5:
                        dayLabel.innerText = "F";
                        break;
                    case 6:
                        dayLabel.innerText = "S";
                        break;
                }
                
                //Day labels
                dayLabel.style.fontFamily = controlItemArray[11];
                dayLabel.style.fontSize = controlItemArray[12] + 'pt';
                dayLabel.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
                dayLabel.style.position = "absolute";
                dayLabel.style.top = "6px";
                dayLabel.style.left = offsetLeft + 3 + "px";

                //AM Boxes
                var amCheckbox = fieldset.appendChild(document.createElement("input"));
                amCheckbox.type = "checkbox";
                amCheckbox.id = controlID + "_" + ((i * 2) + 1);
                amCheckbox.style.padding = "0px";
                amCheckbox.style.position = "absolute";
                amCheckbox.style.top = "22px";
                amCheckbox.style.left = offsetLeft + "px";
                if (!fControlEnabled) amCheckbox.disabled = true;
                
                //PM Boxes
                var pmCheckbox = fieldset.appendChild(document.createElement("input"));
                pmCheckbox.type = "checkbox";
                pmCheckbox.id = controlID + "_" + ((i * 2) + 2);
                pmCheckbox.style.padding = "0px";
                pmCheckbox.style.position = "absolute";
                pmCheckbox.style.top ="36px";
                pmCheckbox.style.left = offsetLeft + "px";
                if (!fControlEnabled) pmCheckbox.disabled = true;
            }

            //AM/PM Labels
            label = document.createElement("span");
            label.innerText = "AM";
            label.style.fontFamily = controlItemArray[11];
            label.style.fontSize = controlItemArray[12] + 'pt';
            label.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
            label.style.position = "absolute";
            label.style.top = top + 22 + "px";
            label.style.left = left + 4 + "px";
            addControl(iPageNo, label);

            label = document.createElement("span");
            label.innerText = "PM";
            label.style.fontFamily = controlItemArray[11];
            label.style.fontSize = controlItemArray[12] + 'pt';
            label.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
            label.style.position = "absolute";
            label.style.top = top + 36 + "px";
            label.style.left = left + 4 + "px";
            addControl(iPageNo, label);

            //ADD FIELDSET AND ITS CONTENTS.
            addControl(iPageNo, fieldset);
            

            break;
        case 8192: //2 ^ 13: //ctlLine
            var line = document.createElement('div');            
            applyLocation(line, controlItemArray, true);
            if (controlItemArray[20] != 0) {
                //Vertical line
                line.style.height = "1px";
            } else {
                line.style.width = "1px";
            }

            line.style.backgroundColor = "gray";
            line.style.padding = "0px";
            line.setAttribute("data-control-key", key);
            //.visible = true
            //.container = tabnumber
            //.alignment
            //.border
            //.top
            //.left
            //.height
            //.width
            //.caption
            //tabIndex
            //.backColor
            //.oletype
            //font, fontsize, fontbold, fontitalic, fontstrikethrough, fontunderline
            //forecolor
            //enabled
            //columnid
            //displaytype
            //navto
            //navin
            //navonsave
            //radio options and border
            //spinner min max increment spinnerposition
            //numeric separator, alignment
            //date/number max size
            //format
            //mask
            //showliterals
            //allow space
            //screenreadonly

            addControl(iPageNo, line);

            break;
        case 2 ^ 14: //ctlNavigation
            //TODO: Nav control always .disabled = false.
            //if (tabIndex > 0) checkbox.tabindex = tabIndex;
            break;
        case 2 ^ 15: //ctlColourPicker
            //TODO: .disabled = (!fControlEnabled);
            //if(tabIndex > 0) checkbox.tabindex = tabIndex;
            break;
        default:
            break;
    }
}

function addHTMLControlValues(controlValues) {
    var controlValuesArray = controlValues.split("\t");
    var lngColumnID = 0;
    var sValue;

    //NB This function is only valid for radio buttons (option groups) and dropdown lists (not lookups).

    for (var i = 0; i < controlValuesArray.length; i++) {

        sValue = controlValuesArray[i];

        if (lngColumnID > 0) {
            if (sValue.length > 0) {
                
                //get the column type, then add this value to it/them.
                //we use a .each function, as there may be more than one column with this ID on the screen.
                $("#ctlRecordEdit").find("[data-columnID='" + lngColumnID + "']").each(function () {


                    //Option Groups
                    if (($(this).is("fieldset")) && ($(this).attr("data-datatype") === "Option Group")) {
                        //unique groupname for the radio buttons.
                        var uniqueID = $(this).attr("id");
                        var alignment = $(this).attr("data-alignment");
                        
                        var fieldset = document.getElementById($(this).attr("id"));                        

                        var radio = fieldset.appendChild(document.createElement("input"));
                        var label = fieldset.appendChild(document.createElement("label"));
                        
                        radio.type = "radio";
                        radio.className = "radio";
                        radio.name = uniqueID;  //used to tie separate radio buttons together.
                        radio.value = sValue;
                        radio.id = uniqueID + "_" + i;
                        
                        if (alignment == 0) {
                            //Vertical alignment
                            radio.style.position = "absolute";
                            radio.style.top = (i * 16) + "px";
                            radio.style.left = "12px";
                            radio.style.padding = "0px";

                            //add text to radio button
                            label.style.position = "absolute";
                            label.style.top = (i * 16) + "px";
                            label.style.left = "28px";
                            label.style.padding = "0px";
                            label.htmlFor = uniqueID + "_" + i;
                            label.appendChild(document.createTextNode(sValue));
                        }
                        if (alignment == 1) {
                            $(this).css("padding-left", "17px");
                            //Horizontal alignment
                            radio.style.padding = "0px";

                            //add text to radio button
                            label.style.marginLeft = "3px";
                            label.style.marginRight = "32px";
                            label.htmlFor = uniqueID + "_" + i;
                            
                            label.appendChild(document.createTextNode(sValue));
                        }
                    }

                    //Dropdown Lists
                    if ($(this).is("select")) {
                        var option = document.createElement('option');
                        option.value = sValue;
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
    return false;
}

function recEdit_setData(columnID, value) {
    //Set the given column's value
    //copied from recordDMI.ocx        
    var fIsIDColumn = false;

    if (columnID.toUpperCase() == "TIMESTAMP") {
        // The column is the timestamp column.
        // mlngTimestamp = CDbl(pvValue)
    }
    else {
        var tmp = this.mavIDColumns.indexOf(Number(columnID));
        if (tmp > 0) {
            this.mavIDColumns[tmp][3] = Number(value);
            fIsIDColumn = true;
        }

        if (!fIsIDColumn) {
            updateControl(Number(columnID), value);
        }
    }
}

function updateControl(lngColumnID, value) {
   
    //get the column type, then add this value to it/them.
    $("#ctlRecordEdit").find("[data-columnID='" + lngColumnID + "']").each(function () {
        

        if ($(this).is("textarea")) {
            $(this).val(value);
        }

        //TODO if coa_image.....

        //TODO if mask

        if ($(this).is("input")) {
                        
            switch ($(this).attr("type")) {
                case "text":
                    $(this).val(value);
                    break;
                case "number":
                    $(this).val(Number(value));
                    break;
                case "checkbox":                    
                    $(this).prop("checked", value == "True" ? true : false);
                    break;

                default:
                    $(this).val(value);

            }
        }

        //Working pattern & Option group
        if ($(this).is("fieldset")) {
            if ($(this).attr("data-datatype") === "Working Pattern")
            {                
                //ensure the value is 14 characters long.
                if (value.length < 14) value = value.concat("              ").substring(0, 14);
                //tick relevant boxes.
                for (var i = 1; i <= 14; i++) {
                    $("#FI_" + lngColumnID + "_" + i).prop("checked", value.substring(i - 1 , i) != " " ? true : false);
                }
            }

            if ($(this).attr("data-datatype") === "Option Group") {
            //TODO
  
            }
        }



        //ComboBox
        if (($(this).is("select")) && (this.length > 0)) {
            $(this).val(value);
        }

        //Lookup
        if (($(this).is("select")) && (this.length == 0)) {
            var option = document.createElement('option');
            option.value = value;
            option.appendChild(document.createTextNode(value));
            $(this).append(option);
            $(this).val(value);
        }

        //Option Group

        //OLE

        //Spinner - done nwith number above..




    });


}

function getTabCaption(tabNumber) {
    
    var psNewValues = $("#txtRecEditTabCaptions").val();
    var arr = psNewValues.split("\t");

    var tabCaption = arr[tabNumber - 1];

    return tabCaption;

}