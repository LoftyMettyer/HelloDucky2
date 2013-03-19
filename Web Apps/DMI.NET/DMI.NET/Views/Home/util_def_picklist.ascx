<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0"
    viewastext>
    <param name="LPKPath" value="lpks/main.lpk">
</object>

<script type="text/javascript">

    function util_def_picklist_onload() {

        var fOK;
        fOK = true;

        $("#workframe").attr("data-framesource", "UTIL_DEF_PICKLIST");

        var sErrMsg = frmUseful.txtErrorDescription.value;

        setGridFont(frmDefinition.ssOleDBGrid);

        if (fOK == true) {
            // Expand the work frame and hide the option frame.
//            window.parent.document.all.item("workframeset").cols = "*, 0";

            if (frmUseful.txtAction.value.toUpperCase() == "NEW") {
                frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
                frmDefinition.txtDescription.value = "";
            } else {
                loadDefinition();
            }

            if (frmUseful.txtAction.value.toUpperCase() != "EDIT") {
                frmUseful.txtUtilID.value = 0;
            }

            if (frmUseful.txtAction.value.toUpperCase() == "COPY") {
                frmUseful.txtChanged.value = 1;
            }

            try {
                frmDefinition.txtName.focus();
            } catch(e) {
            }

            refreshControls();
            frmUseful.txtLoading.value = 'N';
            try {
                frmDefinition.txtName.focus();
            } catch(e) {
            }

            // Get menu.asp to refresh the menu.
            menu_refreshMenu();            
        }
    }

    function refreshControls() {

        var frmUseful = OpenHR.getForm("workframe", "frmUseful");

        fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
        fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

        radio_disable(frmDefinition.optAccessRW, ((fIsNotOwner) || (fViewing)));
        radio_disable(frmDefinition.optAccessRO, ((fIsNotOwner) || (fViewing)));
        radio_disable(frmDefinition.optAccessHD, ((fIsNotOwner) || (fViewing)));
	
        fAddDisabled = fViewing;
        fAddAllDisabled = fViewing;
        fFilteredAddDisabled = fViewing;
        fRemoveDisabled = ((frmDefinition.ssOleDBGrid.SelBookmarks.Count == 0) 
            || (fViewing == true));
        fRemoveAllDisabled = ((frmDefinition.ssOleDBGrid.Rows == 0) 
            || (fViewing == true));

        button_disable(frmDefinition.cmdAdd, fAddDisabled);
        button_disable(frmDefinition.cmdAddAll, fAddAllDisabled);
        button_disable(frmDefinition.cmdFilteredAdd, true);
        button_disable(frmDefinition.cmdRemove, fRemoveDisabled);
        button_disable(frmDefinition.cmdRemoveAll, fRemoveAllDisabled);
    
//        button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
//            (fViewing == true) ||
//            (frmDefinition.ssOleDBGrid.Rows == 0)));

        // Get menu.asp to refresh the menu.
        menu_refreshMenu();		  
    }

    function submitDefinition()
    {
        var i;
        var iIndex;
        var sColumnID;
        var sType;
        var sURL;
	
        if (validate() == false) {menu_refreshMenu(); return;}
        if (populateSendForm() == false) {menu_refreshMenu(); return;}

        // first populate the validate fields
        frmValidate.validatePass.value = 1;
        frmValidate.validateName.value = frmDefinition.txtName.value;
        frmValidate.validateAccess.value = frmSend.txtSend_access.value;

        if(frmUseful.txtAction.value.toUpperCase() == "EDIT"){
            frmValidate.validateTimestamp.value = frmOriginalDefinition.txtDefn_Timestamp.value;
            frmValidate.validateUtilID.value = frmUseful.txtUtilID.value;
        }
        else {
            frmValidate.validateTimestamp.value = 0;
            frmValidate.validateUtilID.value = 0;
        }

//        sURL = "util_validate_picklist" +
//            "?validatePass=" + frmValidate.validatePass.value +
//            "&validateName=" + escape(frmValidate.validateName.value) + 
//            "&validateTimestamp=" + frmValidate.validateTimestamp.value +
//            "&validateUtilID=" + frmValidate.validateUtilID.value +
//            "&validateAccess=" + frmValidate.validateAccess.value +
//            "&validateBaseTableID=" + frmValidate.validateBaseTableID.value;
//        openDialog(sURL, (screen.width) / 2, (screen.height) / 3);
        OpenHR.showInReportFrame(frmValidate);
        //OpenHR.submitForm(frmValidate);
    }

    function addClick() {

        var sURL;
        var vBM;
	
        /* Get the current selected delegate IDs. */
        var sSelectedIDs1 = new String("0");
	
        frmDefinition.ssOleDBGrid.redraw = false;
        if (frmDefinition.ssOleDBGrid.rows > 0) 
        {
            frmDefinition.ssOleDBGrid.MoveFirst();
        }
	
        for (var iIndex = 1; iIndex <= frmDefinition.ssOleDBGrid.rows; iIndex++) 
        {	
            vBM = frmDefinition.ssOleDBGrid.AddItemBookmark(iIndex);
		
            var sRecordID = new String(frmDefinition.ssOleDBGrid.Columns("ID").CellValue(vBM));
		
            sSelectedIDs1 = sSelectedIDs1 + "," + sRecordID;

        }
        frmDefinition.ssOleDBGrid.redraw = true;

        frmPicklistSelection.selectionType.value = "ALL";
        frmPicklistSelection.selectedIDs1.value = sSelectedIDs1;

//        sURL = "util_dialog_picklist" + "?action=add";
  //      openDialog(sURL, (screen.width) / 3, (screen.height) / 2);

        var frmSend = document.getElementById("frmAddSelection");
        frmSend.selectionAction = "add";

        $("#workframeset").hide();
        OpenHR.showInReportFrame(frmSend);

    }

    function openWindow(mypage, myname, w, h, scroll)
    {
        var winl = (screen.width - w) / 2;
        var wint = (screen.height - h) / 2;

        if (scroll == 'no')	
        {
            winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',scrollbars='+scroll+',resize=no';
        }
        else 
        {
            winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',scrollbars='+scroll+',resizable';
        }

        win = window.open(mypage, myname, winprops);
        if (win.opener == null) win.opener = self;
        if (parseInt(navigator.appVersion) >= 4) win.window.focus();
    }

    function addAllClick()
    {	
        frmUseful.txtChanged.value = 1;
        picklistdef_makeSelection("ALLRECORDS", 0, "");
    }

    function filteredAddClick()
    {	
        var sURL;
        var vBM;
	
        /* Get the current selected delegate IDs. */
        var sSelectedIDs1 = new String("0");
	
        frmDefinition.ssOleDBGrid.redraw = false;
        if (frmDefinition.ssOleDBGrid.rows > 0) 
        {
            frmDefinition.ssOleDBGrid.MoveFirst();
        }
	
        for (var iIndex = 1; iIndex <= frmDefinition.ssOleDBGrid.rows; iIndex++) 
        {	
            vBM = frmDefinition.ssOleDBGrid.AddItemBookmark(iIndex);	
            var sRecordID = new String(frmDefinition.ssOleDBGrid.Columns("ID").CellValue(vBM));
            sSelectedIDs1 = sSelectedIDs1 + "," + sRecordID;
        }
        frmDefinition.ssOleDBGrid.redraw = true;

        frmPicklistSelection.selectionType.value = "FILTER";
        frmPicklistSelection.selectedIDs1.value = sSelectedIDs1;

        var frmSend = document.getElementById("frmAddSelection");
        frmSend.selectionAction = "add";

        OpenHR.showInReportFrame(frmSend);

    }

    function removeClick()
    {
        // Do nothing of the Add button is disabled (read-only mode).
        if (frmUseful.txtAction.value.toUpperCase() == "VIEW") return;
	
        iCount = frmDefinition.ssOleDBGrid.selbookmarks.Count();		
        for (i=iCount-1; i >= 0; i--) {
            frmDefinition.ssOleDBGrid.bookmark = frmDefinition.ssOleDBGrid.selbookmarks(i);
            iRowIndex = frmDefinition.ssOleDBGrid.AddItemRowIndex(frmDefinition.ssOleDBGrid.Bookmark);
				
            if ((frmDefinition.ssOleDBGrid.Rows == 1) && (iRowIndex == 0)) {
                frmDefinition.ssOleDBGrid.RemoveAll();
            }
            else {
                frmDefinition.ssOleDBGrid.RemoveItem(iRowIndex);
            }
        }
				
        frmUseful.txtChanged.value = 1;
        refreshControls();
    }

    function removeAllClick()
    {
        iAnswer =OpenHR.messageBox("Remove all records from the picklist. \n Are you sure ?",36,"Confirmation");
        if (iAnswer == 7)	{
            // cancel 
            return;
        }
	
        frmDefinition.ssOleDBGrid.redraw = false;
        frmDefinition.ssOleDBGrid.RemoveAll();
        frmDefinition.ssOleDBGrid.redraw = true;
	
        frmUseful.txtChanged.value = 1;
	
        refreshControls();
    }

    function cancelClick() {
        if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
            (definitionChanged() == false)) {
            menu_loadDefSelPage(10, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, false);
            return (false);
        }

        answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3);
        if (answer == 7) {
            // No
            menu_loadDefSelPage(10, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, false);
            return (false);
        }
        if (answer == 6) {
            // Yes
            okClick();
        }
    }

    function okClick()
    {
        menu_refreshMenu();
	
        frmSend.txtSend_reaction.value = "PICKLISTS";
        submitDefinition();
    }

    function picklistdef_makeSelection(psType, piID, psPrompts)
    {
        
        $(".popup").dialog("close");
        $("#workframeset").show();

        /* Get the current selected delegate IDs. */
        sSelectedIDs = "0";

        if (psType != "ALLRECORDS") 
        {
            frmDefinition.ssOleDBGrid.redraw = false;
            if (frmDefinition.ssOleDBGrid.rows > 0) 
            {
                frmDefinition.ssOleDBGrid.MoveFirst();
            }
            for (iIndex = 1; iIndex <= frmDefinition.ssOleDBGrid.rows; iIndex++) 
            {	
                sRecordID = new String(frmDefinition.ssOleDBGrid.Columns("ID").Value);

                sSelectedIDs = sSelectedIDs + "," + sRecordID;
				
                if (iIndex < frmDefinition.ssOleDBGrid.rows) 
                {
                    frmDefinition.ssOleDBGrid.MoveNext();
                }
                else 
                {
                    break;
                }
            }
            frmDefinition.ssOleDBGrid.redraw = true;
        }
	
        if ((psType == "ALL") && (psPrompts.length > 0)) {
            sSelectedIDs = sSelectedIDs + "," + psPrompts;
        }
	
        // Get the optionData.asp to get the required records.
        var optionDataForm = OpenHR.getForm ("optiondataframe","frmGetOptionData");
        optionDataForm.txtOptionAction.value = "GETPICKLISTSELECTION";
        optionDataForm.txtOptionPageAction.value = psType;
        optionDataForm.txtOptionRecordID.value = piID;
        optionDataForm.txtOptionValue.value = sSelectedIDs;
        optionDataForm.txtOptionPromptSQL.value = psPrompts;
        optionDataForm.txtOptionTableID.value = frmUseful.txtTableID.value;
        optionDataForm.txtOption1000SepCols.value = frmDefinition.txt1000SepCols.value;
	
        refreshOptionData();
    }

    function saveChanges(psAction, pfPrompt, pfTBOverride)
    {
        if ((frmUseful.txtAction.value.toUpperCase() == "VIEW") ||
            (definitionChanged() == false)) {
            return 7; //No to saving the changes, as none have been made.
        }

        answer = OpenHR.messageBox("You have changed the current definition. Save changes ?",3);
        if (answer == 7) {
            // No
            return 7;
        }
        if (answer == 6) {
            // Yes
            okClick();
        }

        return 2; //Cancel.
    }

    function definitionChanged()
    {
        if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
            return false;
        }
	
        if (frmUseful.txtChanged.value == 1) {
            return true;
        }
        else {
            if (frmUseful.txtAction.value.toUpperCase() != "NEW") {
                // Compare the controls with the original values.
                if (frmDefinition.txtName.value != frmOriginalDefinition.txtDefn_Name.value) {
                    return true;
                }
			
                if (frmDefinition.txtDescription.value != frmOriginalDefinition.txtDefn_Description.value) {
                    return true;
                }
			
                if (frmOriginalDefinition.txtDefn_Access.value == "RW") {
                    if (frmDefinition.optAccessRW.checked == false) {
                        return true;
                    }
                }
                else {
                    if (frmOriginalDefinition.txtDefn_Access.value == "RO") {
                        if (frmDefinition.optAccessRO.checked == false) {
                            return true;
                        }
                    }		
                    else {
                        if (frmDefinition.optAccessHD.checked == false) {
                            return true;
                        }
                    }		
                }
            }
        }
	
        return false;
    }

    function openDialog(pDestination, pWidth, pHeight)
    {
        dlgwinprops = "center:yes;" +
            "dialogHeight:" + pHeight + "px;" +
            "dialogWidth:" + pWidth + "px;" +
            "help:no;" +
            "resizable:yes;" +
            "scroll:yes;" +
            "status:no;";
        window.showModalDialog(pDestination, self, dlgwinprops);
        //window.open(pDestination);

    }

    function validate()
    {
        // Check name has been entered.
        if (frmDefinition.txtName.value == '') {
            OpenHR.messageBox("You must enter a name for this definition.");
            return (false);
        }

        // Check thet picklist list does have some records.      
        if (frmDefinition.ssOleDBGrid.rows == 0) {
            OpenHR.messageBox("Picklists must contain at least one record.");
            return (false);
        }
      
        return (true);
    }

    function createNew(pPopup)
    {
        pPopup.close();
	
        frmUseful.txtUtilID.value = 0;
        frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
        frmUseful.txtAction.value = "new";
	
        submitDefinition();
    }

    function populateSendForm() {
        var i;
        var iIndex;
        var sControlName;
        var iNum;

        // Copy all the header information to frmSend
        frmSend.txtSend_ID.value = frmUseful.txtUtilID.value;
        frmSend.txtSend_name.value = frmDefinition.txtName.value;
        frmSend.txtSend_description.value = frmDefinition.txtDescription.value;
        frmSend.txtSend_userName.value = frmDefinition.txtOwner.value;
        if (frmDefinition.optAccessRW.checked == true) {
            frmSend.txtSend_access.value = "RW";
        }
        if (frmDefinition.optAccessRO.checked == true) {
            frmSend.txtSend_access.value = "RO";
        }
        if (frmDefinition.optAccessHD.checked == true) {
            frmSend.txtSend_access.value = "HD";
        }

        // Now go through the records grid
        var sColumns = '';

        frmDefinition.ssOleDBGrid.Redraw = false;
        frmDefinition.ssOleDBGrid.movefirst();

        for (i = 0; i < frmDefinition.ssOleDBGrid.rows; i++) {
            sColumns = sColumns + frmDefinition.ssOleDBGrid.columns("ID").text + ',';

            frmDefinition.ssOleDBGrid.movenext();
        }
        frmDefinition.ssOleDBGrid.Redraw = true;

        frmSend.txtSend_columns.value = sColumns.substr(0, 8000);
        frmSend.txtSend_columns2.value = sColumns.substr(8000, 8000);

        if (sColumns.length > 16000) {
            OpenHR.messageBox("Too many records selected.");
            return false;
        }
        else {
            return true;
        }
    }

    function loadDefinition() {

        frmDefinition.txtName.value = frmOriginalDefinition.txtDefn_Name.value;

        if((frmUseful.txtAction.value.toUpperCase() == "EDIT") ||
            (frmUseful.txtAction.value.toUpperCase() == "VIEW")) {
            frmDefinition.txtOwner.value = frmOriginalDefinition.txtDefn_Owner.value;
        }
        else {
            frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
        }

        frmDefinition.txtDescription.value= frmOriginalDefinition.txtDefn_Description.value;

        if (frmOriginalDefinition.txtDefn_Access.value == "RW") {
            frmDefinition.optAccessRW.checked = true;
        }
        else {
            if (frmOriginalDefinition.txtDefn_Access.value == "RO") {
                frmDefinition.optAccessRO.checked = true;
            }		
            else {
                frmDefinition.optAccessHD.checked = true;
            }		
        }
	
        // Load the selected records into the grid.
        //makeSelection("ALL", 0, frmOriginalDefinition.txtSelectedRecords.value);
        picklistdef_makeSelection("PICKLIST", frmUseful.txtUtilID.value, '');
	
        frmDefinition.ssOleDBGrid.MoveFirst();
        frmDefinition.ssOleDBGrid.FirstRow = frmDefinition.ssOleDBGrid.Bookmark;

        // If its read only, disable everything.
        if(frmUseful.txtAction.value.toUpperCase() == "VIEW"){
            disableAll();
        }
    }

    function disableAll()
    {
        var i;
	
        var dataCollection = frmDefinition.elements;
        if (dataCollection!=null) {
            for (i=0; i<dataCollection.length; i++)  {
                var eElem = frmDefinition.elements[i];

                if (("text" == eElem.type) || ("TEXTAREA" == eElem.tagName)) 
                {
                    textarea_disable(eElem, true);
                }
                else if ("checkbox" == eElem.type)  
                {
                    checkbox_disable(eElem, true);
                }
                else if ("radio" == eElem.type) 
                {
                    radio_disable(eElem, true);
                }
                else if ("button" == eElem.type) 
                {
                    if (eElem.value != "Cancel") 
                    {
                        button_disable(eElem, true);
                    }
                }
                else if ("SELECT" == eElem.tagName) {
                    combo_disable(eElem, true);
                }
                else {
                    treeView_disable(eElem, true);
                }
            }
        }	
    }

    function locateRecord(psSearchFor)
    {  
        var fFound

        fFound = false;
	
        frmDefinition.ssOleDBGrid.redraw = false;

        frmDefinition.ssOleDBGrid.MoveLast();
        frmDefinition.ssOleDBGrid.MoveFirst();

        frmDefinition.ssOleDBGrid.SelBookmarks.removeall();
	
        for (iIndex = 1; iIndex <= frmDefinition.ssOleDBGrid.rows; iIndex++) {	
            var sGridValue = new String(frmDefinition.ssOleDBGrid.Columns(0).value);
            sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
            if (sGridValue == psSearchFor.toUpperCase()) {
                frmDefinition.ssOleDBGrid.SelBookmarks.Add(frmDefinition.ssOleDBGrid.Bookmark);
                fFound = true;
                break;
            }

            if (iIndex < frmDefinition.ssOleDBGrid.rows) {
                frmDefinition.ssOleDBGrid.MoveNext();
            }
            else {
                break;
            }
        }

        if ((fFound == false) && (frmDefinition.ssOleDBGrid.rows > 0)) {
            // Select the top row.
            frmDefinition.ssOleDBGrid.MoveFirst();
            frmDefinition.ssOleDBGrid.SelBookmarks.Add(frmDefinition.ssOleDBGrid.Bookmark);
        }

        frmDefinition.ssOleDBGrid.redraw = true;
    }

    function changeName() {
        frmUseful.txtChanged.value = 1;
        refreshControls();
    }

    function changeDescription() {
        frmUseful.txtChanged.value = 1;
        refreshControls();
    }

    function changeAccess() {
        frmUseful.txtChanged.value = 1;
        refreshControls();
    }

</script>

<script type="text/javascript">

    function util_def_addhandlers() {
        //OpenHR.addActiveXHandler("ssOleDBGrid", "rowColChange", ssOleDBGrid_rowColChange);
        OpenHR.addActiveXHandler("ssOleDBGrid", "KeyPress", ssOleDBGrid_KeyPress);
        OpenHR.addActiveXHandler("ssOleDBGrid", "SelChange", ssOleDBGrid_SelChange);
    }

    function ssOleDBGrid_rowColChange() {
        refreshControls();        
    }

    function ssOleDBGrid_KeyPress(iKeyAscii) {

        if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {	
            var dtTicker = new Date();
            var iThisTick = new Number(dtTicker.getTime());
            if (txtLastKeyFind.value.length > 0) {
                var iLastTick = new Number(txtTicker.value);
            }
            else {
                var iLastTick = new Number("0");
            }
		
            if (iThisTick > (iLastTick + 1500)) {
                var sFind = String.fromCharCode(iKeyAscii);
            }
            else {
                var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
            }
		
            txtTicker.value = iThisTick;
            txtLastKeyFind.value = sFind;

            locateRecord(sFind);
        }
    }
        
    function ssOleDBGrid_SelChange() {
        refreshControls();        
    }

</script>


<form id=frmDefinition>
<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
				<tr> 
					<TD width=10></td>
					<td>
						<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=5>
							<tr valign=top> 
								<td>
									<TABLE WIDTH="100%" height="100%" class="invisible"CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD colspan=9 height=5></TD>
										</TR>

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=10>Name :</TD>
											<TD width=5>&nbsp;</TD>
											<TD>
                                                <input id="txtName" name="txtName" class="text" maxlength="50" style="WIDTH: 100%" onchange="changeName()">
											</TD>
											<TD width=20>&nbsp;</TD>
											<TD width=10>Owner :</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%">
												<INPUT id=txtOwner name=txtOwner class="text textdisabled" style="WIDTH: 100%"  disabled="disabled" tabindex="-1">
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>
											
										<TR>
											<TD colspan=9 height=5></TD>
										</TR>
											
										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=10 nowrap>Description :</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%" rowspan="5">
												<TEXTAREA id=txtDescription name=txtDescription class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap=VIRTUAL height="0" maxlength="255" 
												    onkeyup="changeDescription()" 
												    onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}" 
												    onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
												</TEXTAREA>
											</TD>
											<TD width=20 nowrap>&nbsp;</TD>
											<TD width=10>Access :</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%">
												<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 width="100%">
													<TR>
														<TD width=5>
															<INPUT CHECKED id=optAccessRW name=optAccess type=radio 
															    onclick="changeAccess()"
		                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
														</TD>
														<TD width=5>&nbsp;</TD>
														<TD width=30>
                                                            <label 
                                                                tabindex="-1"
	                                                            for="optAccessRW"
	                                                            class="radio"
		                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                    />
    															Read/Write
                                    	    		        </label>
														</TD>
														<TD>&nbsp;</TD>
													</TR>
												</TABLE>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>
											
										<TR>
											<TD colspan=8 height=5></TD>
										</TR>					

										<TR height=10>
											<TD width=5>&nbsp;</TD>

											<TD width=10>&nbsp;</TD>
											<TD width=5>&nbsp;</TD>

											<TD width=20 nowrap>&nbsp;</TD>

											<TD width=10>&nbsp;</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%">
												<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
													<TR>
														<TD width=5>
															<input id=optAccessRO name=optAccess type=radio 
															    onclick="changeAccess()"
		                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
														</TD>
														<TD width=5>&nbsp;</TD>
														<TD width=80 nowrap>
                                                            <label 
                                                                tabindex="-1"
	                                                            for="optAccessRO"
	                                                            class="radio"
		                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                    />
    															Read Only
                                                            </label>
														</TD>
														<TD>&nbsp;</TD>
													</TR>
												</TABLE>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>
											
										<TR>
											<TD colspan=8 height=5></TD>
										</TR>					

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=10>&nbsp;</TD>
											<TD width=5>&nbsp;</TD>
											<TD width=20 nowrap>&nbsp;</TD>
											<TD width=10>&nbsp;</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%">
												<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
													<TR>
														<TD width=5>
															<input id=optAccessHD name=optAccess type=radio 
															    onclick="changeAccess()"
		                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
														</TD>
														<TD width=5>&nbsp;</TD>
														<TD width=60 nowrap>
                                                            <label 
                                                                tabindex="-1"
	                                                            for="optAccessHD"
	                                                            class="radio"
		                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                    />
															    Hidden
															</label>
														</TD>
														<TD>&nbsp;</TD>
													</TR>
												</TABLE>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>
											
										<TR>
											<TD colspan=9>
												<TABLE WIDTH=100% HEIGHT=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD colspan=3 height=30><hr></TD>
													</TR>
													<TR height=10>
														<TD rowspan=14>
<%
	' Get the employee find columns.
    Dim cmdFindRecords
    Dim prmTableID
    Dim prmErrorMsg
    Dim prm1000SepCols
    Dim rstFindRecords
    Dim sErrorDescription As String
    Dim lngColCount As Long

    cmdFindRecords = Server.CreateObject("ADODB.Command")
	cmdFindRecords.CommandText = "sp_ASRIntGetDefaultOrderColumns"
	cmdFindRecords.CommandType = 4 ' Stored Procedure
    cmdFindRecords.ActiveConnection = Session("databaseConnection")
	cmdFindRecords.CommandTimeout = 180

    prmTableID = cmdFindRecords.CreateParameter("tableID", 3, 1) ' 3=integer, 1 = input
    cmdFindRecords.Parameters.Append(prmTableID)
	prmTableID.value = cleanNumeric(session("utiltableid"))

    prmErrorMsg = cmdFindRecords.CreateParameter("errorMsg", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdFindRecords.Parameters.Append(prmErrorMsg)

    prm1000SepCols = cmdFindRecords.CreateParameter("1000SepCols", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdFindRecords.Parameters.Append(prm1000SepCols)

    Err.Clear()
    rstFindRecords = cmdFindRecords.Execute

    If (Err.Number <> 0) Then
        sErrorDescription = "The find columns could not be retrieved." & vbCrLf & formatError(Err.Description)
    End If

    If Len(sErrorDescription) = 0 Then
        ' Instantiate and initialise the grid. 
        Response.Write("<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGrid name=ssOleDBGrid  codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
        Response.Write("	<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""SelectTypeRow"" VALUE=""3"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
        Response.Write("	<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
        Response.Write("	<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
        Response.Write("	<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)

        lngColCount = 0
        Do While Not rstFindRecords.EOF
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Width"" VALUE=""3200"">" & vbCrLf)
	
            If rstFindRecords.fields("columnName").value = "ID" Then
                Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Visible"" VALUE=""0"">" & vbCrLf)
            Else
                Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Visible"" VALUE=""-1"">" & vbCrLf)
            End If
	
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Caption"" VALUE=""" & Replace(rstFindRecords.fields("columnName").value, "_", " ") & """>" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Name"" VALUE=""" & rstFindRecords.fields("columnName").value & """>" & vbCrLf)
				
            If (rstFindRecords.fields("dataType").value = 131) Or (rstFindRecords.fields("dataType").value = 3) Then
                Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Alignment"" VALUE=""1"">" & vbCrLf)
            Else
                Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Alignment"" VALUE=""0"">" & vbCrLf)
            End If
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Bound"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").DataField"" VALUE=""Column " & lngColCount & """>" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").DataType"" VALUE=""8"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Level"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").NumberFormat"" VALUE="""">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Case"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").FieldLen"" VALUE=""4096"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Locked"" VALUE=""0"">" & vbCrLf)
				
            If rstFindRecords.fields("dataType").value = -7 Then
                ' Find column is a logic column.
                Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Style"" VALUE=""2"">" & vbCrLf)
            Else
                Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Style"" VALUE=""0"">" & vbCrLf)
            End If

            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").RowCount"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").ColCount"" VALUE=""1"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").ForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").BackColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").StyleSet"" VALUE="""">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Nullable"" VALUE=""1"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Mask"" VALUE="""">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").ClipMode"" VALUE=""0"">" & vbCrLf)
            Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").PromptChar"" VALUE=""95"">" & vbCrLf)

            lngColCount = lngColCount + 1
            rstFindRecords.MoveNext()
        Loop
		
        Response.Write("	<PARAM NAME=""Columns.Count"" VALUE=""" & lngColCount & """>" & vbCrLf)
        Response.Write("	<PARAM NAME=""Col.Count"" VALUE=""" & lngColCount & """>" & vbCrLf)

        Response.Write("	<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
        Response.Write("	<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
        Response.Write("	<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

        Response.Write("</OBJECT>" & vbCrLf)

        ' Release the ADO recordset object.
		rstFindRecords.close
        rstFindRecords = Nothing

		' NB. IMPORTANT ADO NOTE.
		' When calling a stored procedure which returns a recordset AND has output parameters
		' you need to close the recordset and set it to nothing before using the output parameters. 
		If Len(cmdFindRecords.Parameters("errorMsg").Value) > 0 Then
			Session("ErrorTitle") = "Picklist Definition Page"
			Session("ErrorText") = cmdFindRecords.Parameters("errorMsg").Value
			Response.Clear()
			
			'Response.Redirect("error.asp")
			Response.Redirect("FormError")
			
        Else
            Response.Write("<INPUT type='hidden' id=txt1000SepCols name=txt1000SepCols value=""" & cmdFindRecords.Parameters("1000SepCols").Value & """>" & vbCrLf)
        End If
    End If
	
	' Release the ADO command object.
    cmdFindRecords = Nothing
%>
														</TD>
														<TD rowspan=14 width=10>&nbsp;</TD>
														<TD width=100>
															<input type=button id=cmdAdd name=cmdAdd class="btn" value=Add style="WIDTH: 100%"  
															    onclick="addClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD></TD>
													</TR>
													<TR height=10>
														<TD width=100>
															<input type=button id=cmdAddAll name=cmdAddAll class="btn" value="Add All" style="WIDTH: 100%"  
															    onclick="addAllClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD></TD>
													</TR>
													<TR height=10>
														<TD width=100>
															<input type=button id=cmdFilteredAdd name=cmdFilteredAdd class="btn" value="Filtered Add" style="WIDTH: 100%"  
															    onclick="filteredAddClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD></TD>
													</TR>
													<TR height=10>
														<TD width=100>
															<input type=button id=cmdRemove name=cmdRemove class="btn" value="Remove" style="WIDTH: 100%"  
															    onclick="removeClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD></TD>
													</TR>
													<TR height=10>
														<TD width=100>
															<input type=button id=cmdRemoveAll name=cmdRemoveAll class="btn" value="Remove All" style="WIDTH: 100%"  
															    onclick="removeAllClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD></TD>
													</TR>
													<TR>
														<TD></TD>
													</TR>
													<TR height=10>
														<TD width=100>
															<input type=button id=cmdOK name=cmdOK class="btn" value=OK style="WIDTH: 100%"
															    onclick="okClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD></TD>
													</TR>
													<TR height=10>
														<TD width=100>
															<input type=button id=cmdCancel name=cmdCancel class="btn" value=Cancel style="WIDTH: 100%"  
															    onclick="cancelClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
											
										<TR height=5>
											<TD colspan=9 height=5></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
						</TABLE>
					</td>
					<TD width=10></td>
				</tr> 

				<tr height=5> 
					<td colspan=3></td>
				</tr> 
			</TABLE>
		</td>
	</tr> 
</TABLE>
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<form id=frmOriginalDefinition style="visibility:hidden;display:none">
<%
    Dim sErrMsg As String
    Dim cmdDefn
    Dim prmUtilID
    Dim prmAction
    Dim prmErrMsg
    Dim prmName
    Dim prmOwner
    
    Dim prmDescription
    Dim prmAccess
    Dim prmTimestamp
    Dim rstDefinition
    Dim sSelectedRecords
    
	sErrMsg = ""

	if session("action") <> "new"	then
        cmdDefn = Server.CreateObject("ADODB.Command")
		cmdDefn.CommandText = "sp_ASRIntGetPicklistDefinition"
		cmdDefn.CommandType = 4 ' Stored Procedure
        cmdDefn.ActiveConnection = Session("databaseConnection")

        prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1) ' 3=integer, 1=input
        cmdDefn.Parameters.Append(prmUtilID)
        prmUtilID.value = CLng(CleanNumeric(Session("utilid")))

        prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefn.Parameters.Append(prmAction)
		prmAction.value = session("action")

        prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000) '200=varchar, 2=output, 8000=size
        cmdDefn.Parameters.Append(prmErrMsg)

        prmName = cmdDefn.CreateParameter("name", 200, 2, 8000) '200=varchar, 2=output, 8000=size
        cmdDefn.Parameters.Append(prmName)

        prmOwner = cmdDefn.CreateParameter("owner", 200, 2, 8000) '200=varchar, 2=output, 8000=size
        cmdDefn.Parameters.Append(prmOwner)

        prmDescription = cmdDefn.CreateParameter("description", 200, 2, 8000) '200=varchar, 2=output, 8000=size
        cmdDefn.Parameters.Append(prmDescription)

        prmAccess = cmdDefn.CreateParameter("access", 200, 2, 8000) '200=varchar, 2=output, 8000=size
        cmdDefn.Parameters.Append(prmAccess)

        prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2) ' 3=integer, 2=output
        cmdDefn.Parameters.Append(prmTimestamp)

        Err.Clear()
        '       rstDefinition = cmdDefn.Execute

        cmdDefn.Execute()

        If (Err.Number <> 0) Then
            sErrMsg = "'" & Session("utilname") & "' picklist definition could not be read." & vbCrLf & formatError(Err.Description)
        Else
            '			if rstDefinition.state <> 0 then
            '				' Read recordset values.
            sSelectedRecords = "0"
            '				do while not rstDefinition.EOF
            '					sSelectedRecords = sSelectedRecords & "," & cstr(rstDefinition.fields("recordID").value)
            '
            '					rstDefinition.MoveNext
            '				loop

            Response.Write("<INPUT type='hidden' id=txtSelectedRecords name=txtSelectedRecords value=""" & sSelectedRecords & """>" & vbCrLf)
	
            ' Release the ADO recordset object.
            '            rstDefinition.close()
            '			end if
            rstDefinition = Nothing
			
            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If Len(cmdDefn.Parameters("errMsg").Value) > 0 Then
                sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").Value
            Else
				
                Response.Write("<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").Value, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").Value, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").Value, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtDefn_Access name=txtDefn_Access value=""" & cmdDefn.Parameters("access").Value & """>" & vbCrLf)
                Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").Value & ">" & vbCrLf)
            End If
        End If

        ' Release the ADO command object.
        cmdDefn = Nothing
	End If
%>
</form>

<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
    <input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
    <input type="hidden" id="txtLoading" name="txtLoading" value="Y">
    <input type="hidden" id="txtChanged" name="txtChanged" value="0">
    <input type="hidden" id="txtUtilID" name="txtUtilID" value='<% =session("utilid")%>'>
    <input type="hidden" id="txtTableID" name="txtTableID" value='<% =session("utiltableid")%>'>
    <input type="hidden" id="txtAction" name="txtAction" value='<% =session("action")%>'>
    <%
        Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
    %>
</form>

<form id="frmValidate" name="frmValidate" method="post" action="util_validate_picklist" style="visibility: hidden; display: none">
    <input type="hidden" id="validatePass" name="validatePass" value="0">
    <input type="hidden" id="validateName" name="validateName" value=''>
    <input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
    <input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
    <input type="hidden" id="validateAccess" name="validateAccess" value=''>
    <input type="hidden" id="validateBaseTableID" name="validateBaseTableID" value='<%=session("utiltableid")%>'>
</form>

<form id="frmAddSelection" name="frmAddSelection" target="validate" method="post" action="util_dialog_picklist" style="visibility: hidden; display: none">
    <input type="hidden" id="selectionAction" name="selectionAction" value="0">
</form>

<form id="frmSend" name="frmSend" method="post" action="util_def_picklist_Submit" style="visibility: hidden; display: none">
    <input type="hidden" id="txtSend_ID" name="txtSend_ID">
    <input type="hidden" id="txtSend_name" name="txtSend_name">
    <input type="hidden" id="txtSend_description" name="txtSend_description">
    <input type="hidden" id="txtSend_access" name="txtSend_access">
    <input type="hidden" id="txtSend_userName" name="txtSend_userName">
    <input type="hidden" id="txtSend_columns" name="txtSend_columns">
    <input type="hidden" id="txtSend_columns2" name="txtSend_columns2">
    <input type="hidden" id="txtSend_reaction" name="txtSend_reaction">
    <input type="hidden" id="txtSend_tableID" name="txtSend_tableID" value='<% =session("utiltableid")%>'>
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
    <input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<form id="frmPicklistSelection" name="frmPicklistSelection" action="picklistSelectionMain" method="post" style="visibility: hidden; display: none">
    <input type="hidden" id="selectionType" name="selectionType">
    <input type="hidden" id="Hidden1" name="txtTableID" value='<% =session("utiltableid")%>'>
    <input type="hidden" id="selectedIDs1" name="selectedIDs1">
</form>

<script runat="server" language="vb">

    Function formatError(psErrMsg)
        Dim iStart
        Dim iFound
  
        iFound = 0
        Do
            iStart = iFound
            iFound = InStr(iStart + 1, psErrMsg, "]")
        Loop While iFound > 0
  
        If (iStart > 0) And (iStart < Len(Trim(psErrMsg))) Then
            formatError = Trim(Mid(psErrMsg, iStart + 1))
        Else
            formatError = psErrMsg
        End If
    End Function

</script>


<script type="text/javascript">
    util_def_addhandlers();
    util_def_picklist_onload();
</script>


