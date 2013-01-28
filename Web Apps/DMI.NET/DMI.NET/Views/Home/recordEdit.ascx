<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/svrCleanup.asp" -->
<%
	Dim sReferringPage

	' Only open the form if there was a referring page.
	' If it wasn't then redirect to the login page.
	sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	if inStrRev(sReferringPage, "/") > 0 then
		sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	end if

	if len(sReferringPage) = 0 then
		Response.Redirect("login.asp")
	end if
%>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="OpenHR.css" rel=stylesheet type=text/css>
<TITLE>OpenHR Intranet</TITLE>

<OBJECT classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0" 
	VIEWASTEXT>
	<PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
</OBJECT>

<SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>
<!--
    var fOK
    fOK = true;	
    var sErrMsg = frmRecordEditForm.txtErrorDescription.value;
    if (sErrMsg.length > 0) {
        fOK = false;
        window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg);
        window.parent.location.replace("login.asp");
    }

    if (fOK == true) {
        // Expand the work frame and hide the option frame.
        window.parent.document.all.item("workframeset").cols = "*, 0";	

        var recEditCtl = frmRecordEditForm.ctlRecordEdit;

        if (recEditCtl==null){
            fOK = false;

            // The recEdit control was not loaded properly.
            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Record Edit control not loaded.");
            window.parent.location.replace("login.asp");
        }
    }

    if (fOK == true) {
        sKey = new String("photopath_");
        sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
        sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
        frmRecordEditForm.txtPicturePath.value = sPath;

        sKey = new String("imagepath_");
        sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
        sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
        frmRecordEditForm.txtImagePath.value = sPath;

        sKey = new String("olePath_");
        sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
        sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
        frmRecordEditForm.txtOLEServerPath.value = sPath;

        sKey = new String("localolePath_");
        sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
        sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
        frmRecordEditForm.txtOLELocalPath.value = sPath;


        // Read and then reset the HR Pro Navigation flag.
        var HRProNavigationFlagValue;
        var HRProNavigationFlag = window.parent.frames("menuframe").document.forms("frmWorkAreaInfo").txtHRProNavigation;
        HRProNavigationFlagValue = HRProNavigationFlag.value;
        HRProNavigationFlag.value = 0;
	
        if (HRProNavigationFlagValue==0) {
            frmGoto.txtGotoTableID.value = frmRecordEditForm.txtCurrentTableID.value;
            frmGoto.txtGotoViewID.value = frmRecordEditForm.txtCurrentViewID.value;
            frmGoto.txtGotoScreenID.value = frmRecordEditForm.txtCurrentScreenID.value;
            frmGoto.txtGotoOrderID.value = frmRecordEditForm.txtCurrentOrderID.value;
            frmGoto.txtGotoRecordID.value = frmRecordEditForm.txtCurrentRecordID.value;
            frmGoto.txtGotoParentTableID.value = frmRecordEditForm.txtCurrentParentTableID.value;
            frmGoto.txtGotoParentRecordID.value = frmRecordEditForm.txtCurrentParentRecordID.value;
            frmGoto.txtGotoPage.value = "recordEdit.asp";

            HRProNavigationFlag.value = 1;
            frmGoto.submit();	
        }
        else {
            // Set the recEdit control properties.
            fOK = recEditCtl.initialise (
				frmRecordEditForm.txtRecEditTableID.value,
				frmRecordEditForm.txtRecEditHeight.value,	
				frmRecordEditForm.txtRecEditWidth.value + 1,
				frmRecordEditForm.txtRecEditTabCount.value,
				frmRecordEditForm.txtRecEditTabCaptions.value,
				frmRecordEditForm.txtRecEditFontName.value,
				frmRecordEditForm.txtRecEditFontSize.value,
				frmRecordEditForm.txtRecEditFontBold.value,
				frmRecordEditForm.txtRecEditFontItalic.value,
				frmRecordEditForm.txtRecEditFontUnderline.value,
				frmRecordEditForm.txtRecEditFontStrikethru.value,
				frmRecordEditForm.txtRecEditRealSource.value,
				frmRecordEditForm.txtPicturePath.value,
				frmRecordEditForm.txtRecEditEmpTableID.value,
				frmRecordEditForm.txtRecEditCourseTableID.value,
				frmRecordEditForm.txtRecEditTBStatusColumnID.value,
				frmRecordEditForm.txtRecEditCourseCancelDateColumnID.value
				);

            if (fOK == true) {		
                // Get the recEdit control to instantiate the required controls.
                var sControlName;
                var controlCollection = frmRecordEditForm.elements;
                if (controlCollection!=null) {
                    for (i=0; i<controlCollection.length; i++)  {
                        sControlName = controlCollection.item(i).name;
                        sControlName = sControlName.substr(0, 18);
                        if (sControlName=="txtRecEditControl_") {
                            fOK = recEditCtl.addControl(controlCollection.item(i).value);
                        }
						
                        if (fOK == false) {
                            break;
                        }
                    }
                }	
            }
			
            if (fOK == true) {
                // Set the column control values in the recEdit control.
                var sControlName;
                var controlCollection = frmRecordEditForm.elements;
                if (controlCollection!=null) {
                    for (i=0; i<controlCollection.length; i++)  {
                        sControlName = controlCollection.item(i).name;
                        sControlName = sControlName.substr(0, 24);
                        if (sControlName=="txtRecEditControlValues_") {
                            fOK = recEditCtl.addControlValues(controlCollection.item(i).value);
                        }

                        if (fOK == false) {
                            break;
                        }
                    }
                }
            }
			
            if (fOK == true) {			
                // Get the recEdit control to format itself.
                recEditCtl.formatscreen();
			
                //JPD 20021021 - Added picture functionality.
                if (frmRecordEditForm.txtImagePath.value.length > 0) {
                    var controlCollection = frmRecordEditForm.elements;
                    if (controlCollection!=null) {
                        for (i=0; i<controlCollection.length; i++)  {
                            sControlName = controlCollection.item(i).name;
                            sControlName = sControlName.substr(0, 18);
                            if (sControlName=="txtRecEditPicture_") {
                                sControlName = controlCollection.item(i).name;
                                iPictureID = new Number(sControlName.substr(18));
                                recEditCtl.updatePicture(iPictureID, frmRecordEditForm.txtImagePath.value + "/" + controlCollection.item(i).value);
                            }
                        }
                    }
                }
            }
						
            if (fOK == true) {			
                // Get the data.asp to get the required data.
                var action = window.parent.frames("dataframe").document.forms("frmGetData").txtAction;
                if (((frmRecordEditForm.txtAction.value == "NEW") ||
						(frmRecordEditForm.txtAction.value == "COPY")) && 
						(frmRecordEditForm.txtRecEditInsertGranted.value == "True")){
                    action.value = frmRecordEditForm.txtAction.value;
                }
                else {
                    action.value = "LOAD";
                }

                if (frmRecordEditForm.txtCurrentOrderID.value != frmRecordEditForm.txtRecEditOrderID.value) {
                    frmRecordEditForm.txtCurrentOrderID.value = frmRecordEditForm.txtRecEditOrderID.value;
                }
				
                var dataForm = window.parent.frames("dataframe").document.forms("frmGetData");
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
                dataForm.txtDefaultCalcCols.value = recEditCtl.CalculatedDefaultColumns();

                window.parent.frames("dataframe").refreshData();
            }

            if (fOK != true) {			
                // The recEdit control was not initialised properly.
                window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Record Edit control not initialised properly.");
                window.parent.location.replace("login.asp");
            }
        }	
    }
    try 
    {
        frmRecordEditForm.ctlRecordEdit.SetWidth(frmRecordEditForm.txtRecEditWidth.value);
        parent.window.resizeBy(-1,-1);
        parent.window.resizeBy(1,1);
    }
    catch(e) {}
    -->
</SCRIPT>

<SCRIPT FOR=ctlRecordEdit EVENT=dataChanged LANGUAGE=JavaScript>
<!--
    // The data in the recEdit control has changed so refresh the menu.
    // Get menu.asp to refresh the menu.
    window.parent.frames("menuframe").refreshMenu();
    -->
</script>

<SCRIPT FOR=ctlRecordEdit EVENT="ToolClickRequest(lngIndex, strTool)" LANGUAGE=JavaScript>
<!--
    // The data in the recEdit control has changed so refresh the menu.
    // Get menu.asp to refresh the menu.
    window.parent.frames("menuframe").MenuClick(strTool);
    -->
</script>

<SCRIPT FOR=ctlRecordEdit EVENT="LinkButtonClick(plngLinkTableID, plngLinkOrderID, plngLinkViewID, plngLinkRecordID)" LANGUAGE=JavaScript>
<!--
    // A link button has been pressed in the recEdit control,
    // so open the link option page.
    window.parent.frames("menuframe").loadLinkPage(plngLinkTableID, plngLinkOrderID, plngLinkViewID, plngLinkRecordID);
    -->
</script>

<SCRIPT FOR=ctlRecordEdit EVENT="LookupClick(plngColumnID, plngLookupColumnID, psLookupValue, pfMandatory, pstrFilterValue)" LANGUAGE=JavaScript>
<!--
    // A lookup button has been pressed in the recEdit control,
    // so open the lookup page.
    window.parent.frames("menuframe").loadLookupPage(plngColumnID, plngLookupColumnID, psLookupValue, pfMandatory, pstrFilterValue);
    -->
</script>

<SCRIPT FOR=ctlRecordEdit EVENT="ImageClick4(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize, pbIsReadOnly)" LANGUAGE=JavaScript>
<!--
    // An image has been pressed in the recEdit control,
    // so open the image find page.
    var fOK;

    fOK = true;
    if (frmRecordEditForm.ctlRecordEdit.recordID == 0)
    {
        window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit photo fields until the record has been saved.");
        fOK = false;
    }

    if (fOK == true) {
        if (plngOLEType < 2) {
            fOK = window.parent.frames("menuframe").ASRIntranetFunctions.ValidateDir(frmRecordEditForm.txtPicturePath.value);
            if (fOK == true)
                window.parent.frames("menuframe").loadImagePage(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize);
            else
                window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit photo fields as the photo path is not valid.");
        }

        else {
            window.parent.frames("menuframe").loadImagePage(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize);
        }
    }
	
    -->
</script>

<SCRIPT FOR=ctlRecordEdit EVENT="OLEClick4(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly)" LANGUAGE=JavaScript>
<!--
    // An OLE button has been pressed in the recEdit control,
    // so open the OLE page.	
    var fOK;
    var sKey = new String('');
  
    fOK = true;
    if (frmRecordEditForm.ctlRecordEdit.recordID == 0)
    {
        window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit OLE fields until the record has been saved.");
        fOK = false;
    }

    if (fOK == true)
    {
        // Server OLE
        if (plngOLEType == 1) {
            fOK = window.parent.frames("menuframe").ASRIntranetFunctions.ValidateDir(frmRecordEditForm.txtOLEServerPath.value);
            if (fOK == true)
                window.parent.frames("menuframe").loadOLEPage(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly);
            else
                window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit server OLE fields as the OLE (Server) path is not valid.");
        }

            // Local OLE
        else if (plngOLEType == 0) {
            fOK = window.parent.frames("menuframe").ASRIntranetFunctions.ValidateDir(frmRecordEditForm.txtOLELocalPath.value);
            if (fOK == true)
                window.parent.frames("menuframe").loadOLEPage(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly);
            else
                window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit local OLE fields as the OLE (Local) path is not valid.");
        }

            // Embedded OLE
        else if (plngOLEType == 2) {
            sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
            window.parent.frames("menuframe").loadOLEPage(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly);
        }

            // Linked OLE
        else if (plngOLEType == 3) {
            sKey = sKey.concat(window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);
            window.parent.frames("menuframe").loadOLEPage(plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly);			
        }
    }
	
    -->
</script>

<script LANGUAGE="JavaScript">
<!--
    function refreshData()
    {
        // Get the data.asp to get the required data.
        var frmGetDataForm = window.parent.frames("dataframe").document.forms("frmGetData");

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
        frmGetDataForm.txtRecordID.value = window.parent.frames("dataframe").document.forms("frmData").txtRecordID.value;
        frmGetDataForm.txtParentTableID.value = frmRecordEditForm.txtCurrentParentTableID.value;
        frmGetDataForm.txtParentRecordID.value = frmRecordEditForm.txtCurrentParentRecordID.value;
        frmGetDataForm.txtDefaultCalcCols.value = frmRecordEditForm.ctlRecordEdit.CalculatedDefaultColumns();
        frmGetDataForm.txtInsertUpdateDef.value = "";
        frmGetDataForm.txtTimestamp.value = "";

        window.parent.frames("dataframe").refreshData();
    }

    function setRecordID(plngRecordID)
    {
        frmRecordEditForm.txtCurrentRecordID.value = plngRecordID;
        frmRecordEditForm.ctlRecordEdit.recordID = plngRecordID;
    }

    function setCopiedRecordID(plngRecordID)
    {
        frmRecordEditForm.ctlRecordEdit.CopiedRecordID = plngRecordID;
    }

    function setParentTableID(plngParentTableID)
    {
        frmRecordEditForm.txtCurrentParentTableID.value = plngParentTableID;
        frmRecordEditForm.ctlRecordEdit.ParentTableID = plngParentTableID;
    }

    function setParentRecordID(plngParentRecordID)
    {
        frmRecordEditForm.txtCurrentParentRecordID.value = plngParentRecordID;
        frmRecordEditForm.ctlRecordEdit.ParentRecordID = plngParentRecordID;
    }

    -->
</script>

<!--The following objects are included to ensure that some of the controls 
that are used in the ASRIntRecEdit control are downloaded and installed properly.
-->
<OBJECT 
	classid=clsid:66A90C04-346D-11D2-9BC0-00A024695830 
	codebase="cabs/timask6.cab#version=6,0,1,1" 
	id=TDBMask1 
	VIEWASTEXT>
</OBJECT>
 
<OBJECT 
	classid=clsid:49CBFCC2-1337-11D2-9BBF-00A024695830 
	codebase="cabs/tinumb6.cab#version=6,0,1,1" 
	id=TDBNumber1 
	VIEWASTEXT>
</OBJECT>

<OBJECT 
	id=ASRUserImage1 
	CLASSID="CLSID:8FF15C8D-49D5-4B79-8419-C36C26654283"
	CODEBASE="cabs/COA_Image.cab#version=1,0,0,7" 
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="2619">
		<PARAM NAME="_ExtentY" VALUE="2619">
		<PARAM NAME="ForeColor" VALUE="0">
		<PARAM NAME="Enabled" VALUE="-1">
		<PARAM NAME="BorderStyle" VALUE="0">
		<PARAM NAME="ASRDataField" VALUE="0">
</OBJECT>

<OBJECT 
	CLASSID="CLSID:C25C3704-2AA7-44E5-943A-B40B14E2348F"
	CODEBASE="cabs/COA_Spinner.cab#version=1,0,0,3"
	id=ASRSpinner1 
	style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
	VIEWASTEXT>
</OBJECT>

<OBJECT 
	classid="clsid:A49CE0E4-C0F9-11D2-B0EA-00A024695830" 
	codebase="cabs/tidate6.cab#version=6,0,1,1" 
	id=TDBDate1 
	style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
	VIEWASTEXT>
</OBJECT>

<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
</HEAD>

<BODY <%=session("BodyTag")%>>
<FORM action="" method=post id=frmRecordEditForm name=frmRecordEditForm>

<table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
		<TD>
			<TABLE WIDTH="100%" ALIGN=center BORDER=0 CELLSPACING=0 CELLPADDING=0>
				<TR>
					<td width=20></td>
					<TD height=10>
						<H3 align=center>
<%
	on error resume next
	
	Dim sErrorDescription
	sErrorDescription = ""

	' Get the page title.
	Set cmdRecEditWindowTitle = Server.CreateObject("ADODB.Command")
	cmdRecEditWindowTitle.CommandText = "sp_ASRIntGetRecordEditInfo"
	cmdRecEditWindowTitle.CommandType = 4 ' Stored Procedure
	Set cmdRecEditWindowTitle.ActiveConnection = session("databaseConnection")

	Set prmTitle = cmdRecEditWindowTitle.CreateParameter("title",200,2,100)
	cmdRecEditWindowTitle.Parameters.Append prmTitle

	Set prmQuickEntry = cmdRecEditWindowTitle.CreateParameter("quickEntry",11,2) ' 11=bit, 2=output
	cmdRecEditWindowTitle.Parameters.Append prmQuickEntry

	Set prmScreenID = cmdRecEditWindowTitle.CreateParameter("screenID",3,1)
	cmdRecEditWindowTitle.Parameters.Append prmScreenID
	prmScreenID.value = cleanNumeric(session("screenID"))

	Set prmViewID = cmdRecEditWindowTitle.CreateParameter("viewID",3,1)
	cmdRecEditWindowTitle.Parameters.Append prmViewID
	prmViewID.value = cleanNumeric(session("viewID"))

	err = 0
    cmdRecEditWindowTitle.Execute
  
	if (err <> 0) then
		sErrorDescription = "The page title could not be created." & vbcrlf & formatError(Err.Description)
	end if

	if len(sErrorDescription) = 0 then		  
		Response.Write replace(cmdRecEditWindowTitle.Parameters("title").Value, "_", " ") & vbcrlf
		Response.Write "<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & cmdRecEditWindowTitle.Parameters("quickEntry").Value & ">" & vbcrlf
	end if
		
	' Release the ADO command object.
	Set cmdRecEditWindowTitle = nothing
%>
						</H3>
					</TD>
					<td width=20></td>
				</TR>
				<TR align=middle>
					<td width=20></td>
					<TD align=middle class="bordered"> 
						<OBJECT 
							CLASSID="CLSID:2D0A5ED7-6669-481F-9A5D-19BA14E92364"
							CODEBASE="cabs/COAInt_RecordDMI.cab#version=1,0,0,21"
							id=ctlRecordEdit 
							VIEWASTEXT>
								<PARAM NAME="_ExtentX" VALUE="16007">
								<PARAM NAME="_ExtentY" VALUE="6403">
								<PARAM NAME="TabCount" VALUE="0">
								<PARAM NAME="TabCaptions" VALUE="">
								<PARAM NAME="BorderStyle" VALUE="0">
						</OBJECT>
					</TD>
					<td width=20></td>
				</TR>
				<TR>
				 <TD colSpan=3 height=10></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</table>

<%
	Response.Write "<INPUT type='hidden' id=txtAction name=txtAction value=" & session("action") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtCurrentTableID name=txtCurrentTableID value=" & session("tableID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtCurrentViewID name=txtCurrentViewID value=" & session("viewID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtCurrentScreenID name=txtCurrentScreenID value=" & session("screenID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtCurrentOrderID name=txtCurrentOrderID value=" & session("orderID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtCurrentRecordID name=txtCurrentRecordID value=" & session("recordID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtCurrentParentTableID name=txtCurrentParentTableID value=" & session("parentTableID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtCurrentParentRecordID name=txtCurrentParentRecordID value=" & session("parentRecordID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtLineage name=txtLineage value=" & session("lineage") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtCurrentRecPos name=txtCurrentRecPos value=" & session("parentRecordID") & ">" & vbcrlf
	
	if len(sErrorDescription) = 0 then
		' Read the screen definition from the database into 'hidden' controls.
		Set cmdRecEditDefinition = Server.CreateObject("ADODB.Command")
		cmdRecEditDefinition.CommandText = "sp_ASRIntGetScreenDefinition"
		cmdRecEditDefinition.CommandType = 4 ' Stored Procedure
		Set cmdRecEditDefinition.ActiveConnection = session("databaseConnection")

		Set prmScreenID = cmdRecEditDefinition.CreateParameter("screenID", 3, 1) ' 3=integer, 1=input
		cmdRecEditDefinition.Parameters.Append prmScreenID
		prmScreenID.value = cleanNumeric(session("screenID"))

		Set prmViewID = cmdRecEditDefinition.CreateParameter("viewID", 3, 1) ' 3=integer, 1=input
		cmdRecEditDefinition.Parameters.Append prmViewID
		prmViewID.value = cleanNumeric(session("viewID"))

		err = 0
	  Set rstScreenDefinition = cmdRecEditDefinition.Execute
	  
		if (err <> 0) then
			sErrorDescription = "The screen definition could not be read." & vbcrlf & formatError(Err.Description)
		end if

		if len(sErrorDescription) = 0 then		  
			Response.Write "<INPUT type='hidden' id=txtRecEditTableID name=txtRecEditTableID value=" & session("tableID") & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditViewID name=txtRecEditViewID value=" & session("viewID") & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditHeight name=txtRecEditHeight value=" & rstScreenDefinition.Fields("height").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditWidth name=txtRecEditWidth value=" & rstScreenDefinition.Fields("width").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditTabCount name=txtRecEditTabCount value=" & rstScreenDefinition.Fields("tabCount").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditTabCaptions name=txtRecEditTabCaptions value=""" & replace(replace(rstScreenDefinition.Fields("tabCaptions").Value, "&", "&&"), """", "&quot;") & """>" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditFontName name=txtRecEditFontName value=""" & replace(rstScreenDefinition.Fields("fontName").Value, """", "&quot;") & """>" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditFontSize name=txtRecEditFontSize value=" & rstScreenDefinition.Fields("fontSize").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditFontBold name=txtRecEditFontBold value=" & rstScreenDefinition.Fields("fontBold").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditFontItalic name=txtRecEditFontItalic value=" & rstScreenDefinition.Fields("fontItalic").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditFontUnderline name=txtRecEditFontUnderline value=" & rstScreenDefinition.Fields("fontUnderline").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditFontStrikethru name=txtRecEditFontStrikethru value=" & rstScreenDefinition.Fields("fontStrikethru").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditRealSource name=txtRecEditRealSource value=""" & replace(rstScreenDefinition.Fields("realSource").Value, """", "&quot;") & """>" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditInsertGranted name=txtRecEditInsertGranted value=" & rstScreenDefinition.Fields("insertGranted").Value & ">" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditDeleteGranted name=txtRecEditDeleteGranted value=" & rstScreenDefinition.Fields("deleteGranted").Value & ">" & vbcrlf
		end if
		
		rstScreenDefinition.close
		Set rstScreenDefinition = nothing
		
		' Release the ADO command object.
		Set cmdRecEditDefinition = nothing
	end if
	
	Response.Write "<INPUT type='hidden' id=txtRecEditEmpTableID name=txtRecEditEmpTableID value=" & session("TB_EmpTableID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtRecEditCourseTableID name=txtRecEditCourseTableID value=" & session("TB_CourseTableID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtRecEditTBTableID name=txtRecEditTBTableID value=" & session("TB_TBTableID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtRecEditTBStatusColumnID name=txtRecEditTBStatusColumnID value=" & session("TB_TBStatusColumnID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtRecEditCourseCancelDateColumnID name=txtRecEditCourseCancelDateColumnID value=" & session("TB_CourseCancelDateColumnID") & ">" & vbcrlf
	'ND commented out for now - Response.Write "<INPUT type='hidden' id=txtWaitListOverRideColumnID name=txtWaitListOverRideColumnID value=" & session("TB_WaitListOverRideColumnID") & ">" & vbcrlf

	if len(sErrorDescription) = 0 then
		' Get the screen controls
		Set cmdRecEditControls = Server.CreateObject("ADODB.Command")
		cmdRecEditControls.CommandText = "sp_ASRIntGetScreenControlsString2"
		cmdRecEditControls.CommandType = 4 ' Stored Procedure
	    Set cmdRecEditControls.ActiveConnection = session("databaseConnection")

		Set prmScreenID = cmdRecEditControls.CreateParameter("screenID",3,1) ' 3=integer, 1=input
		cmdRecEditControls.Parameters.Append prmScreenID
		prmScreenID.value = cleanNumeric(session("screenID"))

		Set prmViewID = cmdRecEditControls.CreateParameter("viewID",3,1) ' 3=integer, 1=input
		cmdRecEditControls.Parameters.Append prmViewID
		prmViewID.value = cleanNumeric(session("viewID"))

		Set prmSelectSQL = cmdRecEditControls.CreateParameter("selectSQL",200,2,2147483646) ' 200=varchar, 2=output
		cmdRecEditControls.Parameters.Append prmSelectSQL

		Set prmFromDef = cmdRecEditControls.CreateParameter("fromDef",200,2,255) ' 200=varchar, 2=output
		cmdRecEditControls.Parameters.Append prmFromDef

		Set prmOrderID = cmdRecEditControls.CreateParameter("orderID", 3, 3) ' 3=integer,  3=input/output
		cmdRecEditControls.Parameters.Append prmOrderID
		prmOrderID.value = cleanNumeric(session("orderID"))

		err = 0
	  Set rstScreenControls = cmdRecEditControls.Execute
	  
		if (err <> 0) then
			sErrorDescription = "The screen control definitions could not be read." & vbcrlf & formatError(Err.Description)
		end if

		if len(sErrorDescription) = 0 then		  
			iloop = 1
			do while not rstScreenControls.EOF
				Response.Write "<INPUT type='hidden' id=txtRecEditControl_" & iLoop & " name=txtRecEditControl_" & iLoop & " value=""" & replace(rstScreenControls.Fields("controlDefinition").Value, """", "&quot;") & """>" & vbcrlf
				rstScreenControls.MoveNext
	
				iloop = iloop + 1
			loop

			' Release the ADO recordset object.
			rstScreenControls.close
			Set rstScreenControls = nothing
		
			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			Response.Write "<INPUT type='hidden' id=txtRecEditSelectSQL name=txtRecEditSelectSQL value=""" & replace(replace(cmdRecEditControls.Parameters("selectSQL").Value, "'", "'''"), """", "&quot;") & """>" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditFromDef name=txtRecEditFromDef value=""" & replace(replace(cmdRecEditControls.Parameters("fromDef").Value, "'", "'''"), """", "&quot;") & """>" & vbcrlf
			Response.Write "<INPUT type='hidden' id=txtRecEditOrderID name=txtRecEditOrderID value=" & cmdRecEditControls.Parameters("orderID").Value & ">" & vbcrlf
		end if
		
		Set cmdRecEditControls = nothing
	end if
	
	if len(sErrorDescription) = 0 then
		' Get the screen column control values
		Set cmdRecEditControlValues = Server.CreateObject("ADODB.Command")
		cmdRecEditControlValues.CommandText = "sp_ASRIntGetScreenControlValuesString"
		cmdRecEditControlValues.CommandType = 4 ' Stored Procedure
		Set cmdRecEditControlValues.ActiveConnection = session("databaseConnection")

		Set prmScreenID = cmdRecEditControlValues.CreateParameter("screenID",3,1)
		cmdRecEditControlValues.Parameters.Append prmScreenID
		prmScreenID.value = cleanNumeric(session("screenID"))

		err = 0
		Set rstScreenControlValues = cmdRecEditControlValues.Execute
		
		if (err <> 0) then
			sErrorDescription = "The screen control values could not be read." & vbcrlf & formatError(Err.Description)
		end if

		if len(sErrorDescription) = 0 then		  
			iloop = 1
			do while not rstScreenControlValues.EOF
				Response.Write "<INPUT type='hidden' id=txtRecEditControlValues_" & iLoop & " name=txtRecEditControlValues_" & iLoop & " value=""" & replace(rstScreenControlValues.Fields("valueDefinition").Value, """", "&quot;") & """>" & vbcrlf
				rstScreenControlValues.MoveNext
		
				iloop = iloop + 1
			loop

			' Release the ADO recordset object.
			rstScreenControlValues.close
			Set rstScreenControlValues = nothing
		end if
	
		Set cmdRecEditControlValues = nothing
	end if

	Response.Write "<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>"
	Response.Write "<INPUT type='hidden' id=txtRecEditFilterDef name=txtRecEditFilterDef value=""" & replace(session("filterDef"), """", "&quot;") & """>" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtRecEditFilterSQL name=txtRecEditFilterSQL value=""" & replace(session("filterSQL"), """", "&quot;") & """>" & vbcrlf

	' JPD 20021021 - Added pictures functionlity.
	' JPD 20021127 - Moved Utilities object into session variable.
	'Set objUtilities = server.CreateObject("COAIntServer.Utilities")
	'objUtilities.Connection = session("databaseConnection")
	Set objUtilities = session("UtilitiesObject")
	sTempPath = server.MapPath("pictures")
	picturesArray = objUtilities.GetPictures(session("screenID"), cstr(sTempPath))

	for iCount = 1 to UBound(picturesArray,2)
		Response.Write "<INPUT type='hidden' id=txtRecEditPicture_" & picturesArray(1,icount) & " name=txtRecEditPicture_" & picturesArray(1,icount) & " value=""" & picturesArray(2,icount) & """>" & vbcrlf
	next 
	Set objUtilities = nothing

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

<FORM action="default_Submit.asp" method=post id=frmGoto name=frmGoto><!--#include file="include\gotoWork.txt"-->
</FORM>

</BODY>
</HTML>

<!-- Embeds createActiveX.js script reference -->
<!--#include file="include\ctl_CreateControl.txt"-->

<% 
function formatError(psErrMsg)
  Dim iStart 
  dim iFound 
  
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
end function
%>
