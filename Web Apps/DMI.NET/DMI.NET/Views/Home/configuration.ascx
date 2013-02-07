<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%
    Dim sReferringPage
    Dim sTemp
    Dim cmdDisplayDefault
    Dim prmSection
    Dim prmKey
    Dim prmDefault
    Dim prmUserSetting
    Dim prmResult
    Dim cmdDefSelOnlyMine
    Dim cmdUtilWarning


	
    'Primary Start Mode.
    cmdDisplayDefault = CreateObject("ADODB.Command")
    cmdDisplayDefault.CommandText = "sp_ASRIntGetSetting"
    cmdDisplayDefault.CommandType = 4 ' Stored procedure.
    cmdDisplayDefault.ActiveConnection = Session("databaseConnection")

    prmSection = cmdDisplayDefault.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmSection)
    prmSection.value = "recordediting"

    prmKey = cmdDisplayDefault.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmKey)
    prmKey.value = "primary"

    prmDefault = cmdDisplayDefault.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmDefault)
    prmDefault.value = "3"

    prmUserSetting = cmdDisplayDefault.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
    cmdDisplayDefault.Parameters.Append(prmUserSetting)
    prmUserSetting.value = 1

    prmResult = cmdDisplayDefault.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdDisplayDefault.Parameters.Append(prmResult)

    Err.Clear()
    cmdDisplayDefault.Execute()
    Session("PrimaryStartMode") = CLng(cmdDisplayDefault.Parameters("result").Value)
    cmdDisplayDefault = Nothing
	
    'History Start Mode.
    cmdDisplayDefault = CreateObject("ADODB.Command")
    cmdDisplayDefault.CommandText = "sp_ASRIntGetSetting"
    cmdDisplayDefault.CommandType = 4 ' Stored procedure.
    cmdDisplayDefault.ActiveConnection = Session("databaseConnection")

    prmSection = cmdDisplayDefault.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmSection)
    prmSection.value = "recordediting"

    prmKey = cmdDisplayDefault.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmKey)
    prmKey.value = "history"

    prmDefault = cmdDisplayDefault.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmDefault)
    prmDefault.value = "3"

    prmUserSetting = cmdDisplayDefault.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
    cmdDisplayDefault.Parameters.Append(prmUserSetting)
    prmUserSetting.value = 1

    prmResult = cmdDisplayDefault.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdDisplayDefault.Parameters.Append(prmResult)

    Err.Clear()
    cmdDisplayDefault.Execute()
    Session("HistoryStartMode") = CLng(cmdDisplayDefault.Parameters("result").Value)
    cmdDisplayDefault = Nothing
	
    'Lookup Start Mode.
    cmdDisplayDefault = CreateObject("ADODB.Command")
    cmdDisplayDefault.CommandText = "sp_ASRIntGetSetting"
    cmdDisplayDefault.CommandType = 4 ' Stored procedure.
    cmdDisplayDefault.ActiveConnection = Session("databaseConnection")

    prmSection = cmdDisplayDefault.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmSection)
    prmSection.value = "recordediting"

    prmKey = cmdDisplayDefault.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmKey)
    prmKey.value = "lookup"

    prmDefault = cmdDisplayDefault.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmDefault)
    prmDefault.value = "3"

    prmUserSetting = cmdDisplayDefault.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
    cmdDisplayDefault.Parameters.Append(prmUserSetting)
    prmUserSetting.value = 1

    prmResult = cmdDisplayDefault.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdDisplayDefault.Parameters.Append(prmResult)

    Err.Clear()
    cmdDisplayDefault.Execute()
    Session("LookupStartMode") = CLng(cmdDisplayDefault.Parameters("result").Value)
    cmdDisplayDefault = Nothing
	
    'Quick Access Start Mode.
    cmdDisplayDefault = CreateObject("ADODB.Command")
    cmdDisplayDefault.CommandText = "sp_ASRIntGetSetting"
    cmdDisplayDefault.CommandType = 4 ' Stored procedure.
    cmdDisplayDefault.ActiveConnection = Session("databaseConnection")

    prmSection = cmdDisplayDefault.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmSection)
    prmSection.value = "recordediting"

    prmKey = cmdDisplayDefault.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmKey)
    prmKey.value = "quickaccess"

    prmDefault = cmdDisplayDefault.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmDefault)
    prmDefault.value = "3"

    prmUserSetting = cmdDisplayDefault.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
    cmdDisplayDefault.Parameters.Append(prmUserSetting)
    prmUserSetting.value = 1

    prmResult = cmdDisplayDefault.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdDisplayDefault.Parameters.Append(prmResult)

    Err.Clear()
    cmdDisplayDefault.Execute()
    Session("QuickAccessStartMode") = CLng(cmdDisplayDefault.Parameters("result").Value)
    cmdDisplayDefault = Nothing
	
    'Expression colour mode.
    cmdDisplayDefault = CreateObject("ADODB.Command")
    cmdDisplayDefault.CommandText = "sp_ASRIntGetSetting"
    cmdDisplayDefault.CommandType = 4 ' Stored procedure.
    cmdDisplayDefault.ActiveConnection = Session("databaseConnection")

    prmSection = cmdDisplayDefault.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmSection)
    prmSection.value = "expressionbuilder"

    prmKey = cmdDisplayDefault.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmKey)
    prmKey.value = "viewcolours"

    prmDefault = cmdDisplayDefault.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmDefault)
    prmDefault.value = "1"

    prmUserSetting = cmdDisplayDefault.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
    cmdDisplayDefault.Parameters.Append(prmUserSetting)
    prmUserSetting.value = 1

    prmResult = cmdDisplayDefault.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdDisplayDefault.Parameters.Append(prmResult)

    Err.Clear()
    cmdDisplayDefault.Execute()
    Session("ExprColourMode") = CLng(cmdDisplayDefault.Parameters("result").Value)
    cmdDisplayDefault = Nothing

    'Expression expand mode.
    cmdDisplayDefault = CreateObject("ADODB.Command")
    cmdDisplayDefault.CommandText = "sp_ASRIntGetSetting"
    cmdDisplayDefault.CommandType = 4 ' Stored procedure.
    cmdDisplayDefault.ActiveConnection = Session("databaseConnection")

    prmSection = cmdDisplayDefault.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmSection)
    prmSection.value = "expressionbuilder"

    prmKey = cmdDisplayDefault.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmKey)
    prmKey.value = "nodesize"

    prmDefault = cmdDisplayDefault.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmDefault)
    prmDefault.value = "1"

    prmUserSetting = cmdDisplayDefault.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
    cmdDisplayDefault.Parameters.Append(prmUserSetting)
    prmUserSetting.value = 1

    prmResult = cmdDisplayDefault.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdDisplayDefault.Parameters.Append(prmResult)

    Err.Clear()
    cmdDisplayDefault.Execute()
    Session("ExprNodeMode") = CLng(cmdDisplayDefault.Parameters("result").Value)
    cmdDisplayDefault = Nothing
	
    'Find window records.
    cmdDisplayDefault = CreateObject("ADODB.Command")
    cmdDisplayDefault.CommandText = "sp_ASRIntGetSetting"
    cmdDisplayDefault.CommandType = 4 ' Stored procedure.
    cmdDisplayDefault.ActiveConnection = Session("databaseConnection")

    prmSection = cmdDisplayDefault.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmSection)
    prmSection.value = "IntranetFindWindow"

    prmKey = cmdDisplayDefault.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmKey)
    prmKey.value = "BlockSize"

    prmDefault = cmdDisplayDefault.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDisplayDefault.Parameters.Append(prmDefault)
    prmDefault.value = "1000"

    prmUserSetting = cmdDisplayDefault.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
    cmdDisplayDefault.Parameters.Append(prmUserSetting)
    prmUserSetting.value = 1

    prmResult = cmdDisplayDefault.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdDisplayDefault.Parameters.Append(prmResult)

    Err.Clear()
    cmdDisplayDefault.Execute()
    Session("FindRecords") = CLng(cmdDisplayDefault.Parameters("result").Value)
    cmdDisplayDefault = Nothing

	
    ' Get the DefSel 'only mine' settings.
    For i = 0 To 20
        sTemp = "onlymine "

        Select Case i
            Case 0
                sTemp = sTemp & "BatchJobs"
            Case 1
                sTemp = sTemp & "Calculations"
            Case 2
                sTemp = sTemp & "CrossTabs"
            Case 3
                sTemp = sTemp & "CustomReports"
            Case 4
                sTemp = sTemp & "DataTransfer"
            Case 5
                sTemp = sTemp & "Export"
            Case 6
                sTemp = sTemp & "Filters"
            Case 7
                sTemp = sTemp & "GlobalAdd"
            Case 8
                sTemp = sTemp & "GlobalUpdate"
            Case 9
                sTemp = sTemp & "GlobalDelete"
            Case 10
                sTemp = sTemp & "Import"
            Case 11
                sTemp = sTemp & "MailMerge"
            Case 12
                sTemp = sTemp & "Picklists"
            Case 13
                sTemp = sTemp & "CalendarReports"
            Case 14
                sTemp = sTemp & "Labels"
            Case 15
                sTemp = sTemp & "LabelDefinition"
            Case 16
                sTemp = sTemp & "MatchReports"
            Case 17
                sTemp = sTemp & "CareerProgression"
            Case 18
                sTemp = sTemp & "EmailGroups"
            Case 19
                sTemp = sTemp & "RecordProfile"
            Case 20
                sTemp = sTemp & "SuccessionPlanning"
        End Select
			
        cmdDefSelOnlyMine = CreateObject("ADODB.Command")
        cmdDefSelOnlyMine.CommandText = "sp_ASRIntGetSetting"
        cmdDefSelOnlyMine.CommandType = 4 ' Stored procedure.
        cmdDefSelOnlyMine.ActiveConnection = Session("databaseConnection")

        prmSection = cmdDefSelOnlyMine.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmSection)
        prmSection.value = "defsel"

        prmKey = cmdDefSelOnlyMine.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmKey)
        prmKey.value = sTemp

        prmDefault = cmdDefSelOnlyMine.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmDefault)
        prmDefault.value = "0"

        prmUserSetting = cmdDefSelOnlyMine.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
        cmdDefSelOnlyMine.Parameters.Append(prmUserSetting)
        prmUserSetting.value = 1

        prmResult = cmdDefSelOnlyMine.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmResult)

        Err.Clear()
        cmdDefSelOnlyMine.Execute()
        Session(sTemp) = CLng(cmdDefSelOnlyMine.Parameters("result").Value)
        cmdDefSelOnlyMine = Nothing
    Next

    ' Get the Utility Warning settings.
    For i = 0 To 4
        sTemp = "warning "

        Select Case i
            Case 0
                sTemp = sTemp & "DataTransfer"
            Case 1
                sTemp = sTemp & "GlobalAdd"
            Case 2
                sTemp = sTemp & "GlobalUpdate"
            Case 3
                sTemp = sTemp & "GlobalDelete"
            Case 4
                sTemp = sTemp & "Import"
        End Select
			
            
        cmdUtilWarning = CreateObject("ADODB.Command")
        cmdUtilWarning.CommandText = "sp_ASRIntGetSetting"
        cmdUtilWarning.CommandType = 4 ' Stored procedure.
        cmdUtilWarning.ActiveConnection = Session("databaseConnection")

        prmSection = cmdUtilWarning.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdUtilWarning.Parameters.Append(prmSection)
        prmSection.value = "warningmsg"

        prmKey = cmdUtilWarning.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdUtilWarning.Parameters.Append(prmKey)
        prmKey.value = sTemp

        prmDefault = cmdUtilWarning.CreateParameter("default", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdUtilWarning.Parameters.Append(prmDefault)
        prmDefault.value = "1"

        prmUserSetting = cmdUtilWarning.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
        cmdUtilWarning.Parameters.Append(prmUserSetting)
        prmUserSetting.value = 1

        prmResult = cmdUtilWarning.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
        cmdUtilWarning.Parameters.Append(prmResult)

        Err.Clear()
        cmdUtilWarning.Execute()
        Session(sTemp) = CLng(cmdUtilWarning.Parameters("result").Value)
        cmdUtilWarning = Nothing
    Next
%>


<script type="text/javascript">
    function configuration_window_onload() {

        $("#workframe").attr("data-framesource", "CONFIGURATION");

//        var frmOriginalConfiguration = OpenHR.getForm("workframe", "frmOriginalConfiguration");

        // Load the original values into tab 1.
        setComboValue("PARENT", frmOriginalConfiguration.txtPrimaryStartMode.value);
        setComboValue("HISTORY", frmOriginalConfiguration.txtHistoryStartMode.value);
        setComboValue("LOOKUP", frmOriginalConfiguration.txtLookupStartMode.value);
        setComboValue("QUICKACCESS", frmOriginalConfiguration.txtQuickAccessStartMode.value);
        setComboValue("EXPRCOLOURMODE", frmOriginalConfiguration.txtExprColourMode.value);
        setComboValue("EXPRNODEMODE", frmOriginalConfiguration.txtExprNodeMode.value);
        frmConfiguration.txtFindSize.value = frmOriginalConfiguration.txtFindSize.value;

        // Load the original values into tab 2. 
        //frmConfiguration.chkOwner_BatchJobs.checked = (frmOriginalConfiguration.txtOnlyMineBatchJobs.value == 1);
        frmConfiguration.chkOwner_Calculations.checked = (frmOriginalConfiguration.txtOnlyMineCalculations.value == 1);
        frmConfiguration.chkOwner_CrossTabs.checked = (frmOriginalConfiguration.txtOnlyMineCrossTabs.value == 1);
        frmConfiguration.chkOwner_CustomReports.checked = (frmOriginalConfiguration.txtOnlyMineCustomReports.value == 1);
        //frmConfiguration.chkOwner_DataTransfer.checked = (frmOriginalConfiguration.txtOnlyMineDataTransfer.value == 1);
        //frmConfiguration.chkOwner_Export.checked = (frmOriginalConfiguration.txtOnlyMineExport.value == 1);
        frmConfiguration.chkOwner_Filters.checked = (frmOriginalConfiguration.txtOnlyMineFilters.value == 1);
        //frmConfiguration.chkOwner_GlobalAdd.checked = (frmOriginalConfiguration.txtOnlyMineGlobalAdd.value == 1);
        //frmConfiguration.chkOwner_GlobalUpdate.checked = (frmOriginalConfiguration.txtOnlyMineGlobalUpdate.value == 1);
        //frmConfiguration.chkOwner_GlobalDelete.checked = (frmOriginalConfiguration.txtOnlyMineGlobalDelete.value == 1);
        //frmConfiguration.chkOwner_Import.checked = (frmOriginalConfiguration.txtOnlyMineImport.value == 1);
        frmConfiguration.chkOwner_MailMerge.checked = (frmOriginalConfiguration.txtOnlyMineMailMerge.value == 1);
        frmConfiguration.chkOwner_Picklists.checked = (frmOriginalConfiguration.txtOnlyMinePicklists.value == 1);
        frmConfiguration.chkOwner_CalendarReports.checked = (frmOriginalConfiguration.txtOnlyMineCalendarReports.value == 1);
        //frmConfiguration.chkOwner_CareerProgression.checked = (frmOriginalConfiguration.txtOnlyMineCareerProgression.value == 1);
        //frmConfiguration.chkOwner_EmailGroups.checked = (frmOriginalConfiguration.txtOnlyMineEmailGroups.value == 1);
        //frmConfiguration.chkOwner_Labels.checked = (frmOriginalConfiguration.txtOnlyMineLabels.value == 1);
        //frmConfiguration.chkOwner_LabelDefinition.checked = (frmOriginalConfiguration.txtOnlyMineLabelDefinition.value == 1);
        //frmConfiguration.chkOwner_MatchReports.checked = (frmOriginalConfiguration.txtOnlyMineMatchReports.value == 1);
        //frmConfiguration.chkOwner_RecordProfile.checked = (frmOriginalConfiguration.txtOnlyMineRecordProfile.value == 1);
        //frmConfiguration.chkOwner_SuccessionPlanning.checked = (frmOriginalConfiguration.txtOnlyMineSuccessionPlanning.value == 1);

        //frmConfiguration.chkWarn_DataTransfer.checked = (frmOriginalConfiguration.txtUtilWarnDataTransfer.value == 1);
        //frmConfiguration.chkWarn_GlobalAdd.checked = (frmOriginalConfiguration.txtUtilWarnGlobalAdd.value == 1);
        //frmConfiguration.chkWarn_GlobalUpdate.checked = (frmOriginalConfiguration.txtUtilWarnGlobalUpdate.value == 1);
        //frmConfiguration.chkWarn_GlobalDelete.checked = (frmOriginalConfiguration.txtUtilWarnGlobalDelete.value == 1);
        //frmConfiguration.chkWarn_Import.checked = (frmOriginalConfiguration.txtUtilWarnImport.value == 1);

        displayPage(1);

    }
</script>


<script type="text/javascript">

    function displayPage(piPageNumber) {
        var iLoop;
        var frmDisplay;

        //TODO: Is this necessary?
        //   frmDisplay = OpenHR.getForm("refreshframe","frmRefresh");
        //     frmDisplay.submit();

        if (piPageNumber == 1) {
            div1.style.visibility = "visible";
            div1.style.display = "block";
            div2.style.visibility = "hidden";
            div2.style.display = "none";

            frmConfiguration.cboPrimaryTableDisplay.focus();
        }

        if (piPageNumber == 2) {
            div1.style.visibility = "hidden";
            div1.style.display = "none";
            div2.style.visibility = "visible";
            div2.style.display = "block";
        }
    }

    function setComboValue(psCombo, piValue) {
        var i;
        var cboCombo;

        if (psCombo == "PARENT") {
            cboCombo = frmConfiguration.cboPrimaryTableDisplay;
        }
        if (psCombo == "HISTORY") {
            cboCombo = frmConfiguration.cboHistoryTableDisplay;
        }
        if (psCombo == "LOOKUP") {
            cboCombo = frmConfiguration.cboLookupTableDisplay;
        }
        if (psCombo == "QUICKACCESS") {
            cboCombo = frmConfiguration.cboQuickAccessDisplay;
        }
        if (psCombo == "EXPRCOLOURMODE") {
            cboCombo = frmConfiguration.cboViewInColour;
        }
        if (psCombo == "EXPRNODEMODE") {
            cboCombo = frmConfiguration.cboExpandNodes;
        }

        for (i = 0; i < cboCombo.options.length; i++) {
            if (cboCombo.options(i).value == piValue) {
                cboCombo.selectedIndex = i;
                return;
            }
        }

        cboCombo.selectedIndex = 0;
    }

    function saveConfiguration() {
        var chkControl;
        var txtControl;
        var sType;
        var frmConfiguration = OpenHR.getForm("workframe", "frmConfiguration");
        // Validate the find window block size.
        if (validateFindBlockSize == false) {
            return (false);
        }

        frmConfiguration.txtPrimaryStartMode.value = frmConfiguration.cboPrimaryTableDisplay.options(frmConfiguration.cboPrimaryTableDisplay.options.selectedIndex).value;
        frmConfiguration.txtHistoryStartMode.value = frmConfiguration.cboHistoryTableDisplay.options(frmConfiguration.cboHistoryTableDisplay.options.selectedIndex).value;
        frmConfiguration.txtLookupStartMode.value = frmConfiguration.cboLookupTableDisplay.options(frmConfiguration.cboLookupTableDisplay.options.selectedIndex).value;
        frmConfiguration.txtQuickAccessStartMode.value = frmConfiguration.cboQuickAccessDisplay.options(frmConfiguration.cboQuickAccessDisplay.options.selectedIndex).value;
        frmConfiguration.txtExprColourMode.value = frmConfiguration.cboViewInColour.options(frmConfiguration.cboViewInColour.options.selectedIndex).value;
        frmConfiguration.txtExprNodeMode.value = frmConfiguration.cboExpandNodes.options(frmConfiguration.cboExpandNodes.options.selectedIndex).value;

        var menuForm = OpenHR.getForm("menuframe", "frmMenuInfo");
        menuForm.txtPrimaryStartMode.value = frmConfiguration.txtPrimaryStartMode.value;
        menuForm.txtHistoryStartMode.value = frmConfiguration.txtHistoryStartMode.value;
        menuForm.txtLookupStartMode.value = frmConfiguration.txtLookupStartMode.value;
        menuForm.txtQuickAccessStartMode.value = frmConfiguration.txtQuickAccessStartMode.value;

        //if (frmConfiguration.chkOwner_BatchJobs.checked == true) frmConfiguration.txtOwner_BatchJobs.value = 1;
        if (frmConfiguration.chkOwner_Calculations.checked == true) frmConfiguration.txtOwner_Calculations.value = 1;
        if (frmConfiguration.chkOwner_CrossTabs.checked == true) frmConfiguration.txtOwner_CrossTabs.value = 1;
        if (frmConfiguration.chkOwner_CustomReports.checked == true) frmConfiguration.txtOwner_CustomReports.value = 1;
        //if (frmConfiguration.chkOwner_DataTransfer.checked == true) frmConfiguration.txtOwner_DataTransfer.value = 1;
        //if (frmConfiguration.chkOwner_Export.checked == true) frmConfiguration.txtOwner_Export.value = 1;
        if (frmConfiguration.chkOwner_Filters.checked == true) frmConfiguration.txtOwner_Filters.value = 1;
        //if (frmConfiguration.chkOwner_GlobalAdd.checked == true) frmConfiguration.txtOwner_GlobalAdd.value = 1;
        //if (frmConfiguration.chkOwner_GlobalDelete.checked == true) frmConfiguration.txtOwner_GlobalDelete.value = 1;
        //if (frmConfiguration.chkOwner_GlobalUpdate.checked == true) frmConfiguration.txtOwner_GlobalUpdate.value = 1;
        //if (frmConfiguration.chkOwner_Import.checked == true) frmConfiguration.txtOwner_Import.value = 1;
        if (frmConfiguration.chkOwner_MailMerge.checked == true) frmConfiguration.txtOwner_MailMerge.value = 1;
        if (frmConfiguration.chkOwner_Picklists.checked == true) frmConfiguration.txtOwner_Picklists.value = 1;
        if (frmConfiguration.chkOwner_CalendarReports.checked == true) frmConfiguration.txtOwner_CalendarReports.value = 1;
        //if (frmConfiguration.chkOwner_CareerProgression.checked == true) frmConfiguration.txtOwner_CareerProgression.value = 1;
        //if (frmConfiguration.chkOwner_EmailGroups.checked == true) frmConfiguration.txtOwner_EmailGroups.value = 1;
        //if (frmConfiguration.chkOwner_Labels.checked == true) frmConfiguration.txtOwner_Labels.value = 1;
        //if (frmConfiguration.chkOwner_LabelDefinition.checked == true) frmConfiguration.txtOwner_LabelDefinition.value = 1;
        //if (frmConfiguration.chkOwner_MatchReports.checked == true) frmConfiguration.txtOwner_MatchReports.value = 1;
        //if (frmConfiguration.chkOwner_RecordProfile.checked == true) frmConfiguration.txtOwner_RecordProfile.value = 1;
        //if (frmConfiguration.chkOwner_SuccessionPlanning.checked == true) frmConfiguration.txtOwner_SuccessionPlanning.value = 1;

        //if (frmConfiguration.chkWarn_DataTransfer.checked == true) frmConfiguration.txtWarn_DataTransfer.value = 1;
        //if (frmConfiguration.chkWarn_GlobalAdd.checked == true) frmConfiguration.txtWarn_GlobalAdd.value = 1;
        //if (frmConfiguration.chkWarn_GlobalDelete.checked == true) frmConfiguration.txtWarn_GlobalDelete.value = 1;
        //if (frmConfiguration.chkWarn_GlobalUpdate.checked == true) frmConfiguration.txtWarn_GlobalUpdate.value = 1;
        //if (frmConfiguration.chkWarn_Import.checked == true) frmConfiguration.txtWarn_Import.value = 1;

        //frmConfiguration.submit();
        OpenHR.submitForm(frmConfiguration);

    }

    function validateFindBlockSize() {
        var sConvertedFindSize;
        var sDecimalSeparator;
        var sThousandSeparator;
        var sPoint;
        var iValue;

        sDecimalSeparator = "\\";
        sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator());
        var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

        sThousandSeparator = "\\";
        sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator());
        var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

        sPoint = "\\.";
        var rePoint = new RegExp(sPoint, "gi");

        if (frmConfiguration.txtFindSize.value == '') {
            frmConfiguration.txtFindSize.value = 0;
        }

        // Convert the find size value from locale to UK settings for use with the isNaN funtion.
        sConvertedFindSize = new String(frmConfiguration.txtFindSize.value);
        // Remove any thousand separators.
        sConvertedFindSize = sConvertedFindSize.replace(reThousandSeparator, "");
        frmConfiguration.txtFindSize.value = sConvertedFindSize;

        // Convert any decimal separators to '.'.
        if (OpenHR.LocaleDecimalSeparator() != ".") {
            // Remove decimal points.
            sConvertedFindSize = sConvertedFindSize.replace(rePoint, "A");
            // replace the locale decimal marker with the decimal point.
            sConvertedFindSize = sConvertedFindSize.replace(reDecimalSeparator, ".");
        }

        if (isNaN(sConvertedFindSize) == true) {
            OpenHR.messageBox("Find window block size must be numeric.");
            frmConfiguration.txtFindSize.value = frmOriginalConfiguration.txtLastFindSize.value;
            displayPage(1);
            frmConfiguration.txtFindSize.focus();
            return false;
        }

        if (frmConfiguration.txtFindSize.value <= 0) {
            OpenHR.messageBox("Find window block size must be greater than 0.");
            frmConfiguration.txtFindSize.value = frmOriginalConfiguration.txtLastFindSize.value;
            displayPage(1);
            frmConfiguration.txtFindSize.focus();
            return false;
        }

        // Find size must be integer.		
        if (sConvertedFindSize.indexOf(".") >= 0) {
            OpenHR.messageBox("Find window block size must be an integer value.");
            frmConfiguration.txtFindSize.value = frmOriginalConfiguration.txtLastFindSize.value;
            displayPage(1);
            frmConfiguration.txtFindSize.focus();
            return false;
        }

        iValue = new Number(frmConfiguration.txtFindSize.value);
        if (iValue > 100000) {
            OpenHR.messageBox("Find window block size cannot be greater than 100000.");
            frmConfiguration.txtFindSize.value = "100000";
            displayPage(1);
            frmConfiguration.txtFindSize.focus();
            return false;
        }

        frmOriginalConfiguration.txtLastFindSize.value = frmConfiguration.txtFindSize.value;

        return true;
    }

    function okClick() {
        frmConfiguration.txtReaction.value = "DEFAULT";
        saveConfiguration();
    }

    /* Return to the default page. */
    function cancelClick() {
        if (definitionChanged() == false) {
            window.location.href = "main";
            return;
        }

        answer = OpenHR.messageBox("You have changed the current configuration. Save changes ?", 3);
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

        answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3);
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
        // Compare the tab 1 controls with the original values.
        if (frmConfiguration.cboPrimaryTableDisplay.options[frmConfiguration.cboPrimaryTableDisplay.selectedIndex].value != frmOriginalConfiguration.txtPrimaryStartMode.value) {
            return true;
        }

        if (frmConfiguration.cboHistoryTableDisplay.options[frmConfiguration.cboHistoryTableDisplay.selectedIndex].value != frmOriginalConfiguration.txtHistoryStartMode.value) {
            return true;
        }

        if (frmConfiguration.cboLookupTableDisplay.options[frmConfiguration.cboLookupTableDisplay.selectedIndex].value != frmOriginalConfiguration.txtLookupStartMode.value) {
            return true;
        }
        if (frmConfiguration.cboQuickAccessDisplay.options[frmConfiguration.cboQuickAccessDisplay.selectedIndex].value != frmOriginalConfiguration.txtQuickAccessStartMode.value) {
            return true;
        }
        if (frmConfiguration.cboViewInColour.options[frmConfiguration.cboViewInColour.selectedIndex].value != frmOriginalConfiguration.txtExprColourMode.value) {
            return true;
        }
        if (frmConfiguration.cboExpandNodes.options[frmConfiguration.cboExpandNodes.selectedIndex].value != frmOriginalConfiguration.txtExprNodeMode.value) {
            return true;
        }

        if (frmConfiguration.txtFindSize.value != frmOriginalConfiguration.txtFindSize.value) {
            return true;
        }

        // Compare the tab 2 controls with the original values.
        /*if ((frmConfiguration.chkOwner_BatchJobs.checked != (frmOriginalConfiguration.txtOnlyMineBatchJobs.value == 1)) ||
            (frmConfiguration.chkOwner_Calculations.checked != (frmOriginalConfiguration.txtOnlyMineCalculations.value == 1)) ||
            (frmConfiguration.chkOwner_CrossTabs.checked != (frmOriginalConfiguration.txtOnlyMineCrossTabs.value == 1)) ||
            (frmConfiguration.chkOwner_CustomReports.checked != (frmOriginalConfiguration.txtOnlyMineCustomReports.value == 1)) ||
            (frmConfiguration.chkOwner_DataTransfer.checked != (frmOriginalConfiguration.txtOnlyMineDataTransfer.value == 1)) ||
            (frmConfiguration.chkOwner_Export.checked != (frmOriginalConfiguration.txtOnlyMineExport.value == 1)) ||
            (frmConfiguration.chkOwner_Filters.checked != (frmOriginalConfiguration.txtOnlyMineFilters.value == 1)) ||
            (frmConfiguration.chkOwner_GlobalAdd.checked != (frmOriginalConfiguration.txtOnlyMineGlobalAdd.value == 1)) ||
            (frmConfiguration.chkOwner_GlobalUpdate.checked != (frmOriginalConfiguration.txtOnlyMineGlobalUpdate.value == 1)) ||
            (frmConfiguration.chkOwner_GlobalDelete.checked != (frmOriginalConfiguration.txtOnlyMineGlobalDelete.value == 1)) ||
            (frmConfiguration.chkOwner_Import.checked != (frmOriginalConfiguration.txtOnlyMineImport.value == 1)) ||
            (frmConfiguration.chkOwner_MailMerge.checked != (frmOriginalConfiguration.txtOnlyMineMailMerge.value == 1)) ||
            (frmConfiguration.chkOwner_Picklists.checked != (frmOriginalConfiguration.txtOnlyMinePicklists.value == 1)) ||
            (frmConfiguration.chkOwner_CalendarReports.checked != (frmOriginalConfiguration.txtOnlyMineCalendarReports.value == 1)) ||
            (frmConfiguration.chkOwner_CareerProgression.checked != (frmOriginalConfiguration.txtOnlyMineCareerProgression.value == 1)) ||
            (frmConfiguration.chkOwner_EmailGroups.checked != (frmOriginalConfiguration.txtOnlyMineEmailGroups.value == 1)) ||
            (frmConfiguration.chkOwner_Labels.checked != (frmOriginalConfiguration.txtOnlyMineLabels.value == 1)) ||
            (frmConfiguration.chkOwner_LabelDefinition.checked != (frmOriginalConfiguration.txtOnlyMineLabelDefinition.value == 1)) ||
            (frmConfiguration.chkOwner_MatchReports.checked != (frmOriginalConfiguration.txtOnlyMineMatchReports.value == 1)) ||
            (frmConfiguration.chkOwner_RecordProfile.checked != (frmOriginalConfiguration.txtOnlyMineRecordProfile.value == 1)) ||
            (frmConfiguration.chkOwner_SuccessionPlanning.checked != (frmOriginalConfiguration.txtOnlyMineSuccessionPlanning.value == 1))) 
            {*/
        if ((frmConfiguration.chkOwner_Calculations.checked != (frmOriginalConfiguration.txtOnlyMineCalculations.value == 1)) ||
            (frmConfiguration.chkOwner_CrossTabs.checked != (frmOriginalConfiguration.txtOnlyMineCrossTabs.value == 1)) ||
            (frmConfiguration.chkOwner_CustomReports.checked != (frmOriginalConfiguration.txtOnlyMineCustomReports.value == 1)) ||
            (frmConfiguration.chkOwner_Filters.checked != (frmOriginalConfiguration.txtOnlyMineFilters.value == 1)) ||
            (frmConfiguration.chkOwner_MailMerge.checked != (frmOriginalConfiguration.txtOnlyMineMailMerge.value == 1)) ||
            (frmConfiguration.chkOwner_Picklists.checked != (frmOriginalConfiguration.txtOnlyMinePicklists.value == 1)) ||
            (frmConfiguration.chkOwner_CalendarReports.checked != (frmOriginalConfiguration.txtOnlyMineCalendarReports.value == 1))) {
            return true;
        }

        /*if ((frmConfiguration.chkWarn_DataTransfer.checked != (frmOriginalConfiguration.txtUtilWarnDataTransfer.value == 1)) ||
            (frmConfiguration.chkWarn_GlobalAdd.checked != (frmOriginalConfiguration.txtUtilWarnGlobalAdd.value == 1)) ||
            (frmConfiguration.chkWarn_GlobalDelete.checked != (frmOriginalConfiguration.txtUtilWarnGlobalDelete.value == 1)) ||
            (frmConfiguration.chkWarn_GlobalUpdate.checked != (frmOriginalConfiguration.txtUtilWarnGlobalUpdate.value == 1)) ||
            (frmConfiguration.chkWarn_Import.checked != (frmOriginalConfiguration.txtUtilWarnImport.value == 1))) {
            return true;
        }*/

        // If you reach here then nothing has changed.
        return false;
    }

    function restoreDefaults() {
        var answer;

        answer = OpenHR.messageBox("Are you sure you want to restore all default settings?", 36);
        if (answer == 6) {
            setComboValue("PARENT", 3);
            setComboValue("HISTORY", 3);
            setComboValue("LOOKUP", 3);
            setComboValue("QUICKACCESS", 1);

            setComboValue("EXPRCOLOURMODE", 1);
            setComboValue("EXPRNODEMODE", 1);

            frmConfiguration.txtFindSize.value = 1000;

            //frmConfiguration.chkOwner_BatchJobs.checked = false;
            frmConfiguration.chkOwner_Calculations.checked = false;
            frmConfiguration.chkOwner_CrossTabs.checked = false;
            frmConfiguration.chkOwner_CustomReports.checked = false;
            //frmConfiguration.chkOwner_DataTransfer.checked = false;
            //frmConfiguration.chkOwner_Export.checked = false;
            frmConfiguration.chkOwner_Filters.checked = false;
            //frmConfiguration.chkOwner_GlobalAdd.checked = false;
            //frmConfiguration.chkOwner_GlobalDelete.checked = false;
            //frmConfiguration.chkOwner_GlobalUpdate.checked = false;
            //frmConfiguration.chkOwner_Import.checked = false;
            frmConfiguration.chkOwner_MailMerge.checked = false;
            frmConfiguration.chkOwner_Picklists.checked = false;
            frmConfiguration.chkOwner_CalendarReports.checked = false;
            //frmConfiguration.chkOwner_CareerProgression.checked = false;
            //frmConfiguration.chkOwner_EmailGroups.checked = false;
            //frmConfiguration.chkOwner_Labels.checked = false;
            //frmConfiguration.chkOwner_LabelDefinition.checked = false;
            //frmConfiguration.chkOwner_MatchReports.checked = false;
            //frmConfiguration.chkOwner_RecordProfile.checked = false;
            //frmConfiguration.chkOwner_SuccessionPlanning.checked = false;

            //frmConfiguration.chkWarn_DataTransfer.checked = true;
            //frmConfiguration.chkWarn_GlobalAdd.checked = true;
            //frmConfiguration.chkWarn_GlobalDelete.checked = true;
            //frmConfiguration.chkWarn_GlobalUpdate.checked = true;
            //frmConfiguration.chkWarn_Import.checked = true;
        }
    }

</script>

<form action="configuration_Submit" method="post" id="frmConfiguration" name="frmConfiguration">
	<br><!-- First tab -->
	<DIV id=div1>
		<table align=center class="outline" cellPadding=5 cellSpacing=0>
			<TR>
				<TD>
					<table align=center class="invisible" cellPadding=0 cellSpacing=0>
                        <TR>
						    <td height=10 colspan=5></td>
						</TR>
						<TR>
							<td height=10 colspan=5>
								<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 align="center">
									<TR>
										<TD width=10>
										    <INPUT type="button" value="Display Defaults" id=btnDummyTab1 name=btnDummyTab1 class="btn btndisabled" disabled=true>
										</TD>
										<td width=10></td>
										<TD width=10>
										    <INPUT type="button" value="Reports/Utilities & Tools" id=btnTab2 name=btnTab2 class="btn" 
										        onclick="displayPage(2)"
			                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                    onfocus="try{button_onFocus(this);}catch(e){}"
			                                    onblur="try{button_onBlur(this);}catch(e){}" />
										</TD>
									</TR>
								</TABLE>
							</td>
						</TR>
						<TR>
							<td height=20 colspan=5></td>
						</TR>
						<TR>
							<td align=center colspan=5>
								<STRONG>Record Editing Start Mode</STRONG>
							</td>
						</TR>
						<TR>
							<td height=10 colspan=5></td>
						</TR>
						<TR>
							<td width=20></td>
							<td align=left nowrap>
								Parent Tables :
							</td>
							<td width=20></td>
							<td align=left>
								<select id=cboPrimaryTableDisplay name=cboPrimaryTableDisplay class="combo" style="HEIGHT: 22px; WIDTH: 200px;" width=200> 
									<option value="3" selected>Find Window</option>
									<option value="2">First Record</option>
									<option value="1">New Record</option>
					 			</select>
							</td>
							<td width=20></td>
						</TR>
						<TR>
							<td height="5" colspan=5>
							</td>
						</TR>
						<TR>
							<td width=20></td>
							<td align=left nowrap>
								Child Tables :
							</td>
							<td width=20>
							</td>
							<td align=left>
								<select id=cboHistoryTableDisplay name=cboHistoryTableDisplay class="combo" style="HEIGHT: 22px; WIDTH: 200px" width=200> 
									<option value="3" selected>Find Window</option>
									<option value="2">First Record</option>
									<option value="1">New Record</option>
					 			</select>
							</td>
							<td width=20>
							</td>
						</TR>
						<TR>
							<td height="5" colspan=5>
							</td>
						</TR>									
						<TR>
							<td width=20></td>
							<td align=left nowrap>
								Lookup Tables :
							</td>
							<td width=20></td>
							<td align=left>
								<select id=cboLookupTableDisplay name=cboLookupTableDisplay class="combo" style="HEIGHT: 22px; WIDTH: 200px" width=200> 
									<option value="3" selected>Find Window</option>
									<option value="2">First Record</option>
									<option value="1">New Record</option>
					 			</select>
							</td>
							<td width=20></td>
						</TR>
						<TR>
							<td height="5" colspan=5>
							</td>
						</TR>
						<TR>
							<td width=20>
							</td>
							<td align=left nowrap>
								Quick Access :
							</td>
							<td width=20>
							</td>
							<td align=left>
								<select id=cboQuickAccessDisplay name=cboQuickAccessDisplay class="combo" style="HEIGHT: 22px; WIDTH: 200px" width=200> 
									<option value="3" selected>Find Window</option>
									<option value="2">First Record</option>
									<option value="1">New Record</option>
					 			</select>
							</td>
							<td width=20>
							</td>
						</TR>
						<TR>
							<td height="20" colspan=5>
							</td>
						</TR>
						<TR>
							<td align=center colspan=5>
								<STRONG>Filters / Calculations</STRONG>
							</td>
						</TR>
						<TR>
							<td height=10 colspan=5>
							</td>
						</TR>
						<TR>
							<td width=20>
							</td>
							<td align=left nowrap>
								View in Colour :
							</td>
							<td width=20>
							</td>
							<td align=left>
								<select id=cboViewInColour name=cboViewInColour class="combo" style="HEIGHT: 22px; WIDTH: 200px" width=200> 
									<option value="1" selected>Monochrome</option>
									<option value="2">Colour Levels</option>
					 			</select>
							</td>
							<td width=20>
							</td>
						</TR>
						<TR>
							<td height="5" colspan=5>
							</td>
						</TR>
						<TR>
							<td width=20>
							</td>
							<td align=left nowrap>
								Expand Nodes :
							</td>
							<td width=20>
							</td>
							<td align=left>
								<select id=cboExpandNodes name=cboExpandNodes class="combo" style="HEIGHT: 22px; WIDTH: 200px" width=200> 
									<option value="1" selected>Minimized</option>
									<option value="2">Expand All</option>
									<option value="4">Expand Top Level</option>
					 			</select>
							</td>
							<td width=20>
							</td>
						</TR>
						<TR>
							<td height="20" colspan=5>
							</td>
						</TR>
						<TR>
							<td align=center colspan=5>
								<STRONG>Find Window / Event Log</STRONG>
							</td>
						</TR>
						<TR>
							<td height=10 colspan=5>
							</td>
						</TR>
						<TR>
							<td width=20>
							</td>
							<td align=left nowrap>
								Block Size :
							</td>
							<td width=20>
							</td>
							<td align=left>
								<INPUT id=txtFindSize name=txtFindSize class="text" style="HEIGHT: 22px; WIDTH: 200px" width=200 
								    onkeyup="validateFindBlockSize()" 
								    onchange="validateFindBlockSize()" />		
							</td>
							<td width=20>
							</td>
						</TR>
						<TR>
							<td height=20 colspan=5></td>
						</TR>

						<TR>
							<td height="5" colspan=5>
								<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0 align="center">
									<td width=20>
									</td>
									<TD width=150>
										<input id="btnDiv1Restore" name="btnDiv1Restore" type="button" value="Restore Defaults" class="btn" style="WIDTH: 150px" width="150" 
										    onclick="restoreDefaults()"
			                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                onfocus="try{button_onFocus(this);}catch(e){}"
			                                onblur="try{button_onBlur(this);}catch(e){}" />
									</TD>
									<TD>&nbsp;</TD>
									<TD width=75>
										<input id="btnDiv1OK" name="btnDiv1OK" type="button" value="OK" class="btn" style="WIDTH: 75px" width="75" 
										    onclick="okClick()"
			                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                onfocus="try{button_onFocus(this);}catch(e){}"
			                                onblur="try{button_onBlur(this);}catch(e){}" />
									</TD>
									<TD width=20></TD>
									<TD width=75>
										<input id="btnDiv1Cancel" name="btnDiv1Cancel" type="button" value="Cancel" class="btn" style="WIDTH: 75px" width="75" 
										    onclick="cancelClick()"
			                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                onfocus="try{button_onFocus(this);}catch(e){}"
			                                onblur="try{button_onBlur(this);}catch(e){}" />
									</TD>
									<td width=20>
									</td>
								</TABLE>
							</td>
						</TR>
						<TR>
							<td height=10 colspan=5></td>
						</TR>

					</TABLE>

				</TD>
			</TR>
		</table>
	</DIV>

	<DIV id=div2 style="visibility:hidden;display:none">
		<table align=center class="outline" cellPadding=5 cellSpacing=0>
			<TR>
				<TD>
					<table align=center class="invisible" cellPadding=0 cellSpacing=0>
						<TR>
							<td height=10 colspan=7></td>
						</TR>
						<TR>
							<td height=10 colspan=7>
								<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 align="center">
									<TR>
										<TD width=10>
										    <INPUT type="button" value="Display Defaults" id=btnTab1 name=btnTab1 class="btn"
										        onclick="displayPage(1)" 
			                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                    onfocus="try{button_onFocus(this);}catch(e){}"
			                                    onblur="try{button_onBlur(this);}catch(e){}" />
										</TD>
										<td width=10></td>
										<TD width=10>
										    <INPUT type="button" value="Reports/Utilities & Tools" id=btnDummyTab2 name=btnDummyTab2 class="btn btndisabled" disabled=true />
										</TD>
									</TR>
								</TABLE>
							</td>
						</TR>
						<TR>
							<td height=20 colspan=7></td>
						</TR>
						<TR>
							<td align=center colspan=7>
								<STRONG>Reports/Utilities & Tools Selection</STRONG>
							</td>
						</TR>
						<TR>
							<td height=10 colspan=7></td>
						</TR>
						<TR>
							<td width=20></td>
							<td align=center colspan=5>
								Only show definitions where owner is '<%=session("username")%>' for the following :
							</td>
							<td width=20></td>
						</TR>
						<TR>
							<td height=10 colspan=7></td>
						</TR>
						<TR>
							<td width=20></td>
							<td align=left nowrap>
								<INPUT type="checkbox" id=chkOwner_Calculations name=chkOwner_Calculations tabindex=-1 
		                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                <label 
				                    for="chkOwner_Calculations"
				                    class="checkbox"
				                    tabindex=0 
				                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                Calculations
    		    		        </label>
							</td>
							<td width=20></td>
							<td align=left nowrap>
								<INPUT type="checkbox" id=chkOwner_Filters name=chkOwner_Filters tabindex=-1 
		                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                <label 
				                    for="chkOwner_Filters"
				                    class="checkbox"
				                    tabindex=0 
				                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
								Filters
    		    		        </label>
							</td>
							<td width=20></td>
						</TR>

						<TR>
							<td width=20></td>
							<td align=left nowrap>
								<INPUT type="checkbox" id=chkOwner_CalendarReports name=chkOwner_CalendarReports tabindex=-1 
		                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                <label 
				                    for="chkOwner_CalendarReports"
				                    class="checkbox"
				                    tabindex=0 
				                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
								Calendar Reports
    		    		        </label>
							</td>
							<td width=20></td>
							<td align=left nowrap>
								<INPUT type="checkbox" id=chkOwner_MailMerge name=chkOwner_MailMerge tabindex=-1 
		                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                <label 
				                    for="chkOwner_MailMerge"
				                    class="checkbox"
				                    tabindex=0 
				                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
								Mail Merge
    		    		        </label>
							</td>
							<td width=20></td>
						</TR>

						<TR>
							<td width=20></td>
							<td align=left nowrap>
								<INPUT type="checkbox" id=chkOwner_CrossTabs name=chkOwner_CrossTabs tabindex=-1 
		                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                <label 
				                    for="chkOwner_CrossTabs"
				                    class="checkbox"
				                    tabindex=0 
				                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
								Cross Tabs
    		    		        </label>
							</td>
							<td width=20></td>
							<td align=left nowrap>
								<INPUT type="checkbox" id=chkOwner_Picklists name=chkOwner_Picklists tabindex=-1 
		                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                <label 
				                    for="chkOwner_Picklists"
				                    class="checkbox"
				                    tabindex=0 
				                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
								Picklists
    		    		        </label>
							</td>
							<td width=20></td>
						</TR>

						<TR>
							<td width=20></td>
							<td align=left nowrap>
								<INPUT type="checkbox" id=chkOwner_CustomReports name=chkOwner_CustomReports tabindex=-1 
		                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                <label 
				                    for="chkOwner_CustomReports"
				                    class="checkbox"
				                    tabindex=0 
				                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
								Custom Reports
    		    		        </label>
							</td>
							<td width=20></td>
							<td align=left nowrap>
								
							</td>
							<td width=20></td>
						</TR>

						<TR>
							<td height="20" colspan=7>
							</td>
						</TR>

						<TR>
							<td height="5" colspan=7>
								<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0 align="center">
									<td width=20></td>
									<TD width=150>
										<input id="btnDiv2Restore" name="btnDiv2Restore" type="button" class="btn" value="Restore Defaults" style="WIDTH: 150px" width="150" 
										    onclick="restoreDefaults()"
			                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                onfocus="try{button_onFocus(this);}catch(e){}"
			                                onblur="try{button_onBlur(this);}catch(e){}" />
									</TD>
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
									<td width=20>
									</td>
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
	</DIV>

    <input type="hidden" id="txtReaction" name="txtReaction">

    <input type="hidden" id="txtPrimaryStartMode" name="txtPrimaryStartMode">
    <input type="hidden" id="txtHistoryStartMode" name="txtHistoryStartMode">
    <input type="hidden" id="txtLookupStartMode" name="txtLookupStartMode">
    <input type="hidden" id="txtQuickAccessStartMode" name="txtQuickAccessStartMode">
    <input type="hidden" id="txtExprColourMode" name="txtExprColourMode">
    <input type="hidden" id="txtExprNodeMode" name="txtExprNodeMode">

    <input type="hidden" id="txtOwner_BatchJobs" name="txtOwner_BatchJobs" value="0">
    <input type="hidden" id="txtOwner_Calculations" name="txtOwner_Calculations" value="0">
    <input type="hidden" id="txtOwner_CrossTabs" name="txtOwner_CrossTabs" value="0">
    <input type="hidden" id="txtOwner_CustomReports" name="txtOwner_CustomReports" value="0">
    <input type="hidden" id="txtOwner_DataTransfer" name="txtOwner_DataTransfer" value="0">
    <input type="hidden" id="txtOwner_Export" name="txtOwner_Export" value="0">
    <input type="hidden" id="txtOwner_Filters" name="txtOwner_Filters" value="0">
    <input type="hidden" id="txtOwner_GlobalAdd" name="txtOwner_GlobalAdd" value="0">
    <input type="hidden" id="txtOwner_GlobalUpdate" name="txtOwner_GlobalUpdate" value="0">
    <input type="hidden" id="txtOwner_GlobalDelete" name="txtOwner_GlobalDelete" value="0">
    <input type="hidden" id="txtOwner_Import" name="txtOwner_Import" value="0">
    <input type="hidden" id="txtOwner_MailMerge" name="txtOwner_MailMerge" value="0">
    <input type="hidden" id="txtOwner_Picklists" name="txtOwner_Picklists" value="0">
    <input type="hidden" id="txtOwner_CalendarReports" name="txtOwner_CalendarReports" value="0">
    <input type="hidden" id="txtOwner_CareerProgression" name="txtOwner_CareerProgression" value="0">
    <input type="hidden" id="txtOwner_EmailGroups" name="txtOwner_EmailGroups" value="0">
    <input type="hidden" id="txtOwner_Labels" name="txtOwner_Labels" value="0">
    <input type="hidden" id="txtOwner_LabelDefinition" name="txtOwner_LabelDefinition" value="0">
    <input type="hidden" id="txtOwner_MatchReports" name="txtOwner_MatchReports" value="0">
    <input type="hidden" id="txtOwner_RecordProfile" name="txtOwner_RecordProfile" value="0">
    <input type="hidden" id="txtOwner_SuccessionPlanning" name="txtOwner_SuccessionPlanning" value="0">

    <input type="hidden" id="txtWarn_DataTransfer" name="txtWarn_DataTransfer" value="0">
    <input type="hidden" id="txtWarn_GlobalAdd" name="txtWarn_GlobalAdd" value="0">
    <input type="hidden" id="txtWarn_GlobalUpdate" name="txtWarn_GlobalUpdate" value="0">
    <input type="hidden" id="txtWarn_GlobalDelete" name="txtWarn_GlobalDelete" value="0">
    <input type="hidden" id="txtWarn_Import" name="txtWarn_Import" value="0">
</form>

<form id="frmOriginalConfiguration" name="frmOriginalConfiguration">
    <input type="hidden" id="Hidden1" name="txtPrimaryStartMode" value='<%=session("PrimaryStartMode")%>'>
    <input type="hidden" id="Hidden2" name="txtHistoryStartMode" value='<%=session("HistoryStartMode")%>'>
    <input type="hidden" id="Hidden3" name="txtLookupStartMode" value='<%=session("LookupStartMode")%>'>
    <input type="hidden" id="Hidden4" name="txtQuickAccessStartMode" value='<%=session("QuickAccessStartMode")%>'>
    <input type="hidden" id="Hidden5" name="txtExprColourMode" value='<%=session("ExprColourMode")%>'>
    <input type="hidden" id="Hidden6" name="txtExprNodeMode" value='<%=session("ExprNodeMode")%>'>
    <input type="hidden" id="Hidden7" name="txtFindSize" value='<%=session("FindRecords")%>'>
    <input type="hidden" id="txtLastFindSize" name="txtLastFindSize" value='<%=session("FindRecords")%>'>

    <input type="hidden" id="txtOnlyMineBatchJobs" name="txtOnlyMineBatchJobs" value='<%=session("onlyMine BatchJobs")%>'>
    <input type="hidden" id="txtOnlyMineCalculations" name="txtOnlyMineCalculations" value='<%=session("onlyMine Calculations")%>'>
    <input type="hidden" id="txtOnlyMineCrossTabs" name="txtOnlyMineCrossTabs" value='<%=session("onlyMine CrossTabs")%>'>
    <input type="hidden" id="txtOnlyMineCustomReports" name="txtOnlyMineCustomReports" value='<%=session("onlyMine CustomReports")%>'>
    <input type="hidden" id="txtOnlyMineDataTransfer" name="txtOnlyMineDataTransfer" value='<%=session("onlyMine DataTransfer")%>'>
    <input type="hidden" id="txtOnlyMineExport" name="txtOnlyMineExport" value='<%=session("onlyMine Export")%>'>
    <input type="hidden" id="txtOnlyMineFilters" name="txtOnlyMineFilters" value='<%=session("onlyMine Filters")%>'>
    <input type="hidden" id="txtOnlyMineGlobalAdd" name="txtOnlyMineGlobalAdd" value='<%=session("onlyMine GlobalAdd")%>'>
    <input type="hidden" id="txtOnlyMineGlobalUpdate" name="txtOnlyMineGlobalUpdate" value='<%=session("onlyMine GlobalUpdate")%>'>
    <input type="hidden" id="txtOnlyMineGlobalDelete" name="txtOnlyMineGlobalDelete" value='<%=session("onlyMine GlobalDelete")%>'>
    <input type="hidden" id="txtOnlyMineImport" name="txtOnlyMineImport" value='<%=session("onlyMine Import")%>'>
    <input type="hidden" id="txtOnlyMineMailMerge" name="txtOnlyMineMailMerge" value='<%=session("onlyMine MailMerge")%>'>
    <input type="hidden" id="txtOnlyMinePicklists" name="txtOnlyMinePicklists" value='<%=session("onlyMine Picklists")%>'>
    <input type="hidden" id="txtOnlyMineCalendarReports" name="txtOnlyMineCalendarReports" value='<%=session("onlyMine CalendarReports")%>'>
    <input type="hidden" id="txtOnlyMineCareerProgression" name="txtOnlyMineCareerProgression" value='<%=session("onlyMine CareerProgression")%>'>
    <input type="hidden" id="txtOnlyMineEmailGroups" name="txtOnlyMineEmailGroups" value='<%=session("onlyMine EmailGroups")%>'>
    <input type="hidden" id="txtOnlyMineLabels" name="txtOnlyMineLabels" value='<%=session("onlyMine Labels")%>'>
    <input type="hidden" id="txtOnlyMineLabelDefinition" name="txtOnlyMineLabelDefinition" value='<%=session("onlyMine LabelDefinition")%>'>
    <input type="hidden" id="txtOnlyMineMatchReports" name="txtOnlyMineMatchReports" value='<%=session("onlyMine MatchReports")%>'>
    <input type="hidden" id="txtOnlyMineRecordProfile" name="txtOnlyMineRecordProfile" value='<%=session("onlyMine RecordProfile")%>'>
    <input type="hidden" id="txtOnlyMineSuccessionPlanning" name="txtOnlyMineSuccessionPlanning" value='<%=session("onlyMine SuccessionPlanning")%>'>

    <input type="hidden" id="txtUtilWarnDataTransfer" name="txtUtilWarnDataTransfer" value='<%=session("warning DataTransfer")%>'>
    <input type="hidden" id="txtUtilWarnGlobalAdd" name="txtUtilWarnGlobalAdd" value='<%=session("warning GlobalAdd")%>'>
    <input type="hidden" id="txtUtilWarnGlobalUpdate" name="txtUtilWarnGlobalUpdate" value='<%=session("warning GlobalUpdate")%>'>
    <input type="hidden" id="txtUtilWarnGlobalDelete" name="txtUtilWarnGlobalDelete" value='<%=session("warning GlobalDelete")%>'>
    <input type="hidden" id="txtUtilWarnImport" name="txtUtilWarnImport" value='<%=session("warning Import")%>'>
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<script type="text/javascript">
    configuration_window_onload();
</script>
