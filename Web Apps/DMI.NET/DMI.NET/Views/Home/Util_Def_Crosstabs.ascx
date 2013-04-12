<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/FormScripts/crosstabdef.js")%>" type="text/javascript"></script>


<%--<script FOR=grdAccess EVENT=ComboCloseUp LANGUAGE=JavaScript>
<!--
    frmUseful.txtChanged.value = 1;
    if((grdAccess.AddItemRowIndex(grdAccess.Bookmark) == 0) &&
        (grdAccess.Columns("Access").Text.length > 0)) {
        ForceAccess(grdAccess, AccessCode(grdAccess.Columns("Access").Text));
    
        grdAccess.MoveFirst();
        grdAccess.Col = 1;
    }
    refreshTab1Controls();
-->
</script>

<script FOR=grdAccess EVENT=GotFocus LANGUAGE=JavaScript>
<!--
    grdAccess.Col = 1
-->
</script>

<script FORM=grdAccess EVENT=RowColChange(LastRow, LastCol) LANGUAGE=JavaScript>
<!--
    var fViewing;
    var fIsNotOwner;
    var varBkmk;
		
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

    if (grdAccess.AddItemRowIndex(grdAccess.Bookmark) == 0) {
        grdAccess.Columns("Access").Text = "";
    }

    varBkmk = grdAccess.SelBookmarks(0);

    if ((fIsNotOwner == true) ||
        (fViewing == true) ||
        (frmSelectionAccess.forcedHidden.value == "Y") ||
        (grdAccess.Columns("SysSecMgr").CellText(varBkmk) == "1")) {
        grdAccess.Columns("Access").Style = 0; // 0 = Edit
    }
    else {
        grdAccess.Columns("Access").Style = 3; // 3 = Combo box
        grdAccess.Columns("Access").RemoveAll();
        grdAccess.Columns("Access").AddItem(AccessDescription("RW"));
        grdAccess.Columns("Access").AddItem(AccessDescription("RO"));
        grdAccess.Columns("Access").AddItem(AccessDescription("HD"));
    }

    grdAccess.Col = 1;
</script>

<script FOR=grdAccess EVENT=RowLoaded(Bookmark) LANGUAGE=JavaScript>
    var fViewing;
    var fIsNotOwner;
		
    fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
    fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());

    if ((fIsNotOwner == true) ||
        (fViewing == true) ||
        (frmSelectionAccess.forcedHidden.value == "Y")) {
        grdAccess.Columns("GroupName").CellStyleSet("ReadOnly");
        grdAccess.Columns("Access").CellStyleSet("ReadOnly");
        grdAccess.ForeColor = "-2147483631";
    }  
    else {
        if (grdAccess.Columns("SysSecMgr").CellText(Bookmark) == "1") {
            grdAccess.Columns("GroupName").CellStyleSet("SysSecMgr");
            grdAccess.Columns("Access").CellStyleSet("SysSecMgr");
            grdAccess.ForeColor = "0";
        }
        else {
            grdAccess.ForeColor = "0";
        }
    }
</script>--%>

<object classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
    id="dialog"
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

<DIV <%=session("BodyTag")%>>
    <form id=frmTables style="visibility:hidden;display:none">
        <%
            Dim sErrorDescription = ""

            ' Get the table records.
            Dim cmdTables = CreateObject("ADODB.Command")
            cmdTables.CommandText = "sp_ASRIntGetCrossTabTablesInfo"
            cmdTables.CommandType = 4 ' Stored Procedure
            cmdTables.ActiveConnection = Session("databaseConnection")
	
            Response.Write("<B>Set Connection</B>")
	
            Err.Clear()
            Dim rstTablesInfo = cmdTables.Execute
	
            Response.Write("<B>Executed SP</B>")
	
            If (Err.Number <> 0) Then
                sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                ' Dim iCount = 0
                Do While Not rstTablesInfo.EOF
                    Response.Write("<INPUT type='hidden' id=txtTableName_" & rstTablesInfo.fields("tableID").value & " name=txtTableName_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("tableName").value & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtTableType_" & rstTablesInfo.fields("tableID").value & " name=txtTableType_" & rstTablesInfo.fields("tableID").value & " value=" & rstTablesInfo.fields("tableType").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenString").value & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenNames").value & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtTableParents_" & rstTablesInfo.fields("tableID").value & " name=txtTableParents_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("parentsString").value & """>" & vbCrLf)

                    rstTablesInfo.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstTablesInfo.close()
                rstTablesInfo = Nothing
            End If
	
            ' Release the ADO command object.
            cmdTables = Nothing
%>
    </form>
    <form id=frmOriginalDefinition name=frmOriginalDefinition style="visibility:hidden;display:none">
        <%
            Dim sErrMsg = ""
            Dim lngHStart = 0
            Dim lngHStop = 0
            Dim lngHStep = 0
            Dim lngVStart = 0
            Dim lngVStop = 0
            Dim lngVStep = 0
            Dim lngPStart = 0
            Dim lngPStop = 0
            Dim lngPStep = 0

            If Session("action") <> "new" Then
                Dim cmdDefn = CreateObject("ADODB.Command")
                cmdDefn.CommandText = "sp_ASRIntGetCrossTabDefinition"
                cmdDefn.CommandType = 4 ' Stored Procedure
                cmdDefn.ActiveConnection = Session("databaseConnection")
                
                Dim prmUtilDefnID = cmdDefn.CreateParameter("utilid", 3, 1) ' 3=integer, 1=input
                cmdDefn.Parameters.Append(prmUtilDefnID)
                prmUtilDefnID.value = CleanNumeric(Session("utilid"))
                
                Dim prmUser = cmdDefn.CreateParameter("user", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
                cmdDefn.Parameters.Append(prmUser)
                prmUser.value = Session("username")

                Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
                cmdDefn.Parameters.Append(prmAction)
                prmAction.value = Session("action")

                Dim prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmErrMsg)

                Dim prmName = cmdDefn.CreateParameter("name", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmName)

                Dim prmOwner = cmdDefn.CreateParameter("owner", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOwner)

                Dim prmDescription = cmdDefn.CreateParameter("description", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmDescription)

                Dim prmBaseTableID = cmdDefn.CreateParameter("baseTableID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmBaseTableID)

                Dim prmAllRecords = cmdDefn.CreateParameter("allRecords", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmAllRecords)

                Dim prmPicklistID = cmdDefn.CreateParameter("picklistID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmPicklistID)

                Dim prmPicklistName = cmdDefn.CreateParameter("picklistName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmPicklistName)

                Dim prmPicklistHidden = cmdDefn.CreateParameter("picklistHidden", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmPicklistHidden)

                Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmFilterID)

                Dim prmFilterName = cmdDefn.CreateParameter("filterName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmFilterName)

                Dim prmFilterHidden = cmdDefn.CreateParameter("filterHidden", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmFilterHidden)
		
                Dim prmPrintFilter = cmdDefn.CreateParameter("PrintFilter", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmPrintFilter)

                Dim prmHColID = cmdDefn.CreateParameter("HColID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmHColID)

                Dim prmHStart = cmdDefn.CreateParameter("HStart", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmHStart)

                Dim prmHStop = cmdDefn.CreateParameter("HStop", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmHStop)

                Dim prmHStep = cmdDefn.CreateParameter("HStep", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmHStep)

                Dim prmVColID = cmdDefn.CreateParameter("VColID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmVColID)

                Dim prmVStart = cmdDefn.CreateParameter("VStart", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmVStart)

                Dim prmVStop = cmdDefn.CreateParameter("VStop", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmVStop)

                Dim prmVStep = cmdDefn.CreateParameter("VStep", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmVStep)

                Dim prmPColID = cmdDefn.CreateParameter("PColID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmPColID)

                Dim prmPStart = cmdDefn.CreateParameter("PStart", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmPStart)

                Dim prmPStop = cmdDefn.CreateParameter("PStop", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmPStop)

                Dim prmPStep = cmdDefn.CreateParameter("PStep", 200, 2, 20) '3=integer, 2=output, 20=size
                cmdDefn.Parameters.Append(prmPStep)

                Dim prmIType = cmdDefn.CreateParameter("IType", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmIType)

                Dim prmIColID = cmdDefn.CreateParameter("IColID", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmIColID)

                Dim prmPercentage = cmdDefn.CreateParameter("Percentage", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmPercentage)

                Dim prmPerPage = cmdDefn.CreateParameter("PerPage", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmPerPage)

                Dim prmSuppress = cmdDefn.CreateParameter("Suppress", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmSuppress)

                Dim prmThousand = cmdDefn.CreateParameter("Thousand", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmThousand)

                Dim prmOutputPreview = cmdDefn.CreateParameter("outputPreview", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputPreview)
		
                Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmOutputFormat)
		
                Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputScreen)
		
                Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputPrinter)
		
                Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputPrinterName)
		
                Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputSave)
		
                Dim prmOutputSaveExisting = cmdDefn.CreateParameter("outputSaveExisting", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmOutputSaveExisting)
		
                Dim prmOutputEmail = cmdDefn.CreateParameter("outputEmail", 11, 2) '11=bit, 2=output
                cmdDefn.Parameters.Append(prmOutputEmail)
		
                Dim prmOutputEmailAddr = cmdDefn.CreateParameter("outputEmailAddr", 3, 2) '3=integer, 2=output
                cmdDefn.Parameters.Append(prmOutputEmailAddr)

                Dim prmOutputEmailAddrName = cmdDefn.CreateParameter("outputEmailAddrName", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputEmailAddrName)

                Dim prmOutputEmailSubject = cmdDefn.CreateParameter("outputEmailSubject", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputEmailSubject)

                Dim prmOutputEmailAttachAs = cmdDefn.CreateParameter("outputEmailAttachAs", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputEmailAttachAs)

                Dim prmOutputFilename = cmdDefn.CreateParameter("outputFilename", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdDefn.Parameters.Append(prmOutputFilename)

                Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2) ' 3=integer, 2=output
                cmdDefn.Parameters.Append(prmTimestamp)

                Err.Clear()
                cmdDefn.Execute()

                Dim iHiddenCalcCount As Integer = 0
                If (Err.Number <> 0) Then
                    sErrMsg = "'" & Session("utilname") & "' cross tab definition could not be read." & vbCrLf & FormatError(Err.Description)
                Else

                    'rstDefinition.close
                    'set rstDefinition = nothing

                    ' NB. IMPORTANT ADO NOTE.
                    ' When calling a stored procedure which returns a recordset AND has output parameters
                    ' you need to close the recordset and set it to nothing before using the output parameters. 
                    If Len(cmdDefn.Parameters("errMsg").value) > 0 Then
                        sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").value
                    End If

                    lngHStart = cmdDefn.Parameters("HStart").value
                    lngHStop = cmdDefn.Parameters("HStop").value
                    lngHStep = cmdDefn.Parameters("HStep").value
                    lngVStart = cmdDefn.Parameters("VStart").value
                    lngVStop = cmdDefn.Parameters("VStop").value
                    lngVStep = cmdDefn.Parameters("VStep").value
                    lngPStart = cmdDefn.Parameters("PStart").value
                    lngPStop = cmdDefn.Parameters("PStop").value
                    lngPStep = cmdDefn.Parameters("PStep").value

                    Response.Write("<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & cmdDefn.Parameters("baseTableID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_AllRecords name=txtDefn_AllRecords value=" & cmdDefn.Parameters("allRecords").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & cmdDefn.Parameters("picklistID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(cmdDefn.Parameters("picklistName").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & cmdDefn.Parameters("picklistHidden").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & cmdDefn.Parameters("filterID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(cmdDefn.Parameters("filterName").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & cmdDefn.Parameters("filterHidden").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_FilterHeader name=txtDefn_FilterHeader value=" & cmdDefn.Parameters("PrintFilter").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PrintFilter name=txtDefn_PrintFilter value=" & cmdDefn.Parameters("PrintFilter").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HColID name=txtDefn_HColID value=" & cmdDefn.Parameters("HColID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HStart name=txtDefn_HStart value=" & cmdDefn.Parameters("HStart").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HStop name=txtDefn_HStop value=" & cmdDefn.Parameters("HStop").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HStep name=txtDefn_HStep value=" & cmdDefn.Parameters("HStep").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_VColID name=txtDefn_VColID value=" & cmdDefn.Parameters("VColID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_VStart name=txtDefn_VStart value=" & cmdDefn.Parameters("VStart").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_VStop name=txtDefn_VStop value=" & cmdDefn.Parameters("VStop").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_VStep name=txtDefn_VStep value=" & cmdDefn.Parameters("VStep").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PColID name=txtDefn_PColID value=" & cmdDefn.Parameters("PColID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PStart name=txtDefn_PStart value=" & cmdDefn.Parameters("PStart").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PStop name=txtDefn_PStop value=" & cmdDefn.Parameters("PStop").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PStep name=txtDefn_PStep value=" & cmdDefn.Parameters("PStep").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_IType name=txtDefn_IType value=" & cmdDefn.Parameters("IType").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_IColID name=txtDefn_IColID value=" & cmdDefn.Parameters("IColID").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Percentage name=txtDefn_Percentage value=" & cmdDefn.Parameters("Percentage").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_PerPage name=txtDefn_PerPage value=" & cmdDefn.Parameters("PerPage").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Suppress name=txtDefn_Suppress value=" & cmdDefn.Parameters("Suppress").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_Use1000 name=txtDefn_Use1000 value=" & cmdDefn.Parameters("Thousand").value & ">" & vbCrLf)

                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputPreview name=txtDefn_OutputPreview value=" & cmdDefn.Parameters("OutputPreview").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & cmdDefn.Parameters("OutputFormat").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & cmdDefn.Parameters("OutputScreen").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & cmdDefn.Parameters("OutputPrinter").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & cmdDefn.Parameters("OutputPrinterName").value & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & cmdDefn.Parameters("OutputSave").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputSaveExisting name=txtDefn_OutputSaveExisting value=" & cmdDefn.Parameters("OutputSaveExisting").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmail name=txtDefn_OutputEmail value=" & cmdDefn.Parameters("OutputEmail").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddr name=txtDefn_OutputEmailAddr value=" & cmdDefn.Parameters("OutputEmailAddr").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddrName name=txtDefn_OutputEmailName value=""" & Replace(cmdDefn.Parameters("OutputEmailAddrName").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailSubject name=txtDefn_OutputEmailSubject value=""" & Replace(cmdDefn.Parameters("OutputEmailSubject").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAttachAs name=txtDefn_OutputEmailAttachAs value=""" & Replace(cmdDefn.Parameters("OutputEmailAttachAs").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_OutputFilename name=txtDefn_OutputFilename value=""" & cmdDefn.Parameters("OutputFilename").value & """>" & vbCrLf)

                    Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").value & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=session_action name=session_action value=" & Session("action") & ">" & vbCrLf)
                    Response.Write("</form>" & vbCrLf)

                End If

                ' Release the ADO command object.
                cmdDefn = Nothing

                If Len(sErrMsg) > 0 Then
                    Session("confirmtext") = sErrMsg
                    Session("confirmtitle") = "OpenHR Intranet"
                    Session("followpage") = "defsel"
                    Session("reaction") = "CROSSTABS"
                    Response.Clear()
                    Response.Redirect("confirmok")
                End If
	
            Else
                Session("childcount") = 0
                Session("hiddenfiltercount") = 0
            End If
%>
    </form>

    <form id=frmDefinition name=frmDefinition>
        <table valign=top align=center class="outline" cellPadding=5 width=100% height=100% cellSpacing=0>
            <tr>
                <td>
                    <TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
                        <tr height=5> 
                            <td colspan=3></td>
                        </tr> 

                        <tr height=10>
                            <td width=10></td>
                            <td>
                                <INPUT type="button" value="Definition" id=btnTab1 name=btnTab1 class="btn btndisabled" disabled="disabled"
                                       onclick="displayPage(1)" 
                                       onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                       onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                       onfocus="try{button_onFocus(this);}catch(e){}"
                                       onblur="try{button_onBlur(this);}catch(e){}" />
                                <INPUT type="button" value="Columns" id=btnTab2 name=btnTab2 class="btn btndisabled" disabled="disabled"
                                       onclick="displayPage(2)" 
                                       onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                       onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                       onfocus="try{button_onFocus(this);}catch(e){}"
                                       onblur="try{button_onBlur(this);}catch(e){}" />
                                <INPUT type="button" value="Output" id=btnTab3 name=btnTab3 class="btn btndisabled" disabled="disabled"
                                       onclick="displayPage(3)" 
                                       onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                       onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                       onfocus="try{button_onFocus(this);}catch(e){}"
                                       onblur="try{button_onBlur(this);}catch(e){}" />
                            </td>
                            <td width=10></td>
                        </tr>

                        <tr height=10> 
                            <td colspan=3></td>
                        </tr> 

                        <tr>
                            <td width=10></td>
                            <td>
                                <!-- First tab -->
                                <DIV id=div1>
                                    <TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
                                        <tr valign=top> 
                                            <td>
                                                <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                    <tr height=10>
                                                        <td width=5>&nbsp;</td>
                                                        <td width=10>Name :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td>
                                                            <INPUT id=txtName name=txtName maxlength="50" style="WIDTH: 100%" class="text"
                                                                   onkeyup="changeTab1Control()">
                                                        </td>
                                                        <td width=20>&nbsp;</td>
                                                        <td width=10>Owner :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%">
                                                            <INPUT id=txtOwner name=txtOwner class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
                                                        </td>
                                                        <td width=5>&nbsp;</td>
                                                    </tr>

                                                    <tr>
                                                        <td colspan=9 height=5></td>
                                                    </tr>

                                                    <tr height=60>
                                                        <td width=5>&nbsp;</td>
                                                        <td width=10 nowrap valign=top>Description :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%" rowspan="3">
                                                            <TEXTAREA id=txtDescription name=txtDescription class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap=VIRTUAL height="0" maxlength="255" 
                                                                      onkeyup="changeTab1Control()" 
                                                                      onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}" 
                                                                      onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
													</TEXTAREA>
                                                        </td>
                                                        <td width=20 nowrap>&nbsp;</td>
                                                        <td width=10 valign=top>Access :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%" rowspan="3" valign=top>
                                                        </td>
                                                        <td width=5>&nbsp;</td>
                                                    </tr>

                                                    <tr height=10>
                                                        <td colspan=7>&nbsp;</td>
                                                    </tr>

                                                    <tr height=10>
                                                        <td colspan=7>&nbsp;</td>
                                                    </tr>

                                                    <tr>
                                                        <td colspan=9><hr></td>
                                                    </tr>

                                                    <tr height=10>
                                                        <td width=5>&nbsp;</td>
                                                        <td width=100 nowrap vAlign=top>Base Table :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%" vAlign=top>
                                                            <select id=cboBaseTable name=cboBaseTable style="WIDTH: 100%" class="combo combodisabled"
                                                                    onchange="changeBaseTable()" disabled="disabled"> 
                                                            </select>
                                                        </td>
                                                        <td width=20 nowrap>&nbsp;</td>
                                                        <td width=10 vAlign=top>Records :</td>
                                                        <td width=5>&nbsp;</td>
                                                        <td width="40%"> 
                                                            <TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
                                                                <tr>
                                                                    <td width=5>
                                                                        <input CHECKED id=optRecordSelection1 name=optRecordSelection type=radio 
                                                                               onclick="changeBaseTableRecordOptions()"
                                                                               onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                               onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td width=30>
                                                                        <label 
                                                                            tabindex="-1"
                                                                            for="optRecordSelection1"
                                                                            class="radio"
                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
                                                                        All
                                                                    </label>
                                                                    </td>
                                                                    <td>&nbsp;</td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan=6 height=5></td>
                                                                </tr>
                                                                <tr>
                                                                    <td width=5>
                                                                        <input id=optRecordSelection2 name=optRecordSelection type=radio 
                                                                               onclick="changeBaseTableRecordOptions()"
                                                                               onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                               onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td width=20>
                                                                        <label 
                                                                            tabindex="-1"
                                                                            for="optRecordSelection2"
                                                                            class="radio"
                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
                                                                        Picklist
                                                                    </label>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td>
                                                                        <INPUT id=txtBasePicklist name=txtBasePicklist class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
                                                                    </td>
                                                                    <td width=30>
                                                                        <INPUT id=cmdBasePicklist name=cmdBasePicklist style="WIDTH: 100%" type=button disabled="disabled" value="..." class="btn btndisabled"
                                                                               onclick="selectRecordOption('base', 'picklist')"
                                                                               onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{button_onFocus(this);}catch(e){}"
                                                                               onblur="try{button_onBlur(this);}catch(e){}" />
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan=6 height=5></td>
                                                                </tr>
                                                                <tr>
                                                                    <td width=5>
                                                                        <input id=optRecordSelection3 name=optRecordSelection type=radio
                                                                               onclick=changeBaseTableRecordOptions() 
                                                                               onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                               onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td width=20>
                                                                        <label 
                                                                            tabindex="-1"
                                                                            for="optRecordSelection3"
                                                                            class="radio"
                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
                                                                        Filter
                                                                    </label>
                                                                    </td>
                                                                    <td width=5>&nbsp;</td>
                                                                    <td>
                                                                        <INPUT id=txtBaseFilter name=txtBaseFilter disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
                                                                    </td>
                                                                    <td width=30>
                                                                        <INPUT id=cmdBaseFilter name=cmdBaseFilter style="WIDTH: 100%" type=button class="btn btndisabled" disabled="disabled" value="..."
                                                                               onclick="selectRecordOption('base', 'filter')" 
                                                                               onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                               onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                               onfocus="try{button_onFocus(this);}catch(e){}"
                                                                               onblur="try{button_onBlur(this);}catch(e){}" />
                                                                    </td>
                                                                </tr>
                                                            </TABLE>
                                                        </td>
                                                        <td width=5>&nbsp;</td>
                                                    </tr>
											
                                                    <tr>
                                                        <td colspan=9 height=5>&nbsp;</td>
                                                    </tr>
											

                                                    <tr>
                                                        <td colspan=5>&nbsp;</td>
                                                        <td colspan=3>
                                                            <input name=chkPrintFilter id=chkPrintFilter type=checkbox disabled="disabled" tabindex=-1 
                                                                   onclick="changeTab1Control()"
                                                                   onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                   onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
                                                                for="chkPrintFilter"
                                                                class="checkbox checkboxdisabled"
                                                                tabindex=0 
                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                Display filter or picklist title in the report header
                                                            </label> 
                                                        </td>
                                                        <td width=5>&nbsp;</td>
                                                    </tr>
											
                                                    <tr>
                                                        <td colspan=9 height=5>&nbsp;</td>
                                                    </tr>
                                                </TABLE>
                                            </td>
                                        </tr>
                                    </TABLE>
                                </DIV>
                                <DIV id=div2 style="visibility:hidden;display:none">
                                    <TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>

                                        <tr valign=top> 
                                            <td>
                                                <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>

                                                <tr height=10>
                                                    <td width=5>&nbsp;</td>
                                                    <td colspan=4 vAlign=top><U>Headings & Breaks</U></td>
                                                    <td width="15%" align=Center>Start</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%" align=Center>Stop</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%" align=Center>Increment</td>
                                                    <td>&nbsp;</td>
                                                </tr>

                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>
                                                <tr height=23>
                                                    <td width=5>&nbsp;</td>
                                                    <td width=80 nowrap vAlign=top>Horizontal :</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="40%" vAlign=top>
                                                        <select id=cboHor name=cboHor style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                onchange="cboHor_Change();changeTab2Control(); "> 
                                                        </select>
                                                    </td>
                                                    <td width=15>&nbsp;</td>
                                                    <td>
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtHorStart name=txtHorStart width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngHStart%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtHorStop name=txtHorStop width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngHStop%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtHorStep name=txtHorStep width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngHStep%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>

                                                    <td>&nbsp;</td>
                                                </tr>

                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>
                                                <tr height=23>
                                                    <td width=5>&nbsp;</td>
                                                    <td width=80 nowrap vAlign=top>Vertical :</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="40%" vAlign=top>
                                                        <select id=cboVer name=cboVer style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                onchange="cboVer_Change();changeTab2Control(); "> 
                                                        </select>
                                                    </td>
                                                    <td width=15>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtVerStart name=txtVerStart width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngVStart%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtVerStop name=txtVerStop width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngVStop%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtVerStep name=txtVerStep width="100%" height="100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngVStep%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>

                                                    <td>&nbsp;</td>
                                                </tr>
                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>
                                                <tr height=23>
                                                    <td width=5>&nbsp;</td>
                                                    <td width=100 nowrap vAlign=top>Page Break :</td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="40%" vAlign=top>
                                                        <select id=cboPgb name=cboPgb style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                onchange="cboPgb_Change();changeTab2Control(); " > 
                                                        </select>
                                                    </td>
                                                    <td width=15>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtPgbStart name=txtPgbStart style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngPStart%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtPgbStop name=txtPgbStop style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngPStop%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td width=5>&nbsp;</td>
                                                    <td width="15%">
                                                        <OBJECT CLASSID="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"  codebase="cabs/tinumb6.cab#version=6,0,1,1" id=txtPgbStep name=txtPgbStep style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" 
                                                                onkeyup="changeTab2Control()">
                                                            <PARAM NAME="DisplayFormat" VALUE="##########0.0000">
                                                            <PARAM NAME="Format" VALUE="##########0.0000">
                                                            <PARAM NAME="MaxValue" VALUE="99999999999.9999">
                                                            <PARAM NAME="Value" VALUE="<%=lngPStep%>">
                                                            <PARAM NAME="Appearance" VALUE="0">
                                                            <PARAM NAME="BackColor" VALUE="15988214">
                                                            <PARAM NAME="ForeColor" VALUE="6697779">
                                                        </OBJECT>
                                                    </td>
                                                    <td>&nbsp;</td>
                                                </tr>

                                                <tr height=40>
                                                    <td colspan=11><hr></td>
                                                </tr>

                                                <tr height=10>
                                                    <td width=5>&nbsp;</td>
                                                    <td width=80 colspan=4 nowrap vAlign=top><U>Intersection</U></td>
                                                </tr>
                                                <tr>
                                                    <td colspan=9 height=5></td>
                                                </tr>

                                                <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                    <tr height=0>
                                                        <td width=90></td>
                                                        <td width="40%"></td>
                                                    </tr>

                                                    <td colspan=2>
                                                        <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                            <tr height=10>
                                                                <td width=5>&nbsp;</td>
                                                                <td width=80 nowrap vAlign=top>Column :</td>
                                                                <td width=5>&nbsp;</td>
                                                                <td width="100%" vAlign=top>
                                                                    <select id=cboInt name=cboInt style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                            onchange="cboInt_Change();changeTab2Control(); " > 
                                                                    </select>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td colspan=9 height=5></td>
                                                            </tr>
                                                            <tr height=10>
                                                                <td width=5>&nbsp;</td>
                                                                <td width=80 nowrap vAlign=top>Type :</td>
                                                                <td width=5>&nbsp;</td>
                                                                <td width="100%" vAlign=top>
                                                                    <select id=cboIntType name=cboIntType style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
                                                                            onchange="changeTab2Control()" >
                                                                        <option value="1">Average</option>
                                                                        <option value="0" selected>Count</option>
                                                                        <option value="2">Maximum</option>
                                                                        <option value="3">Minimum</option>
                                                                        <option value="4">Total</option>
                                                                    </select>
                                                                </td>
                                                            </tr>
                                                        </TABLE>
                                                    </td>
                                                    <td width=15>&nbsp;</td>
                                                    <td>
                                                        <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                                                            <tr>
                                                                <td>
                                                                    <INPUT type="checkbox" id=chkPercentage name=chkPercentage tabindex=-1
                                                                           onclick="changeTab2Control()"
                                                                           onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                    <label 
                                                                        for="chkPercentage"
                                                                        class="checkbox"
                                                                        tabindex=0 
                                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    
                                                                        Percentage of Type
                                                                    </label> 
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td height=5></td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <INPUT type="checkbox" id=chkPerPage name=chkPerPage tabindex=-1
                                                                           onclick="changeTab2Control()"                                                                     
                                                                           onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                    <label 
                                                                        for="chkPerPage"
                                                                        class="checkbox"
                                                                        tabindex=0 
                                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    
                                                                        Percentage of Page
                                                                    </label> 
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td height=5></td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <INPUT type="checkbox" id=chkSuppress name=chkSuppress tabindex=-1
                                                                           onclick="changeTab2Control()"
                                                                           onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                    <label 
                                                                        for="chkSuppress"
                                                                        class="checkbox"
                                                                        tabindex=0 
                                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    
                                                                        Suppress Zeros
                                                                    </label> 
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td height=5></td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <INPUT type="checkbox" id=chkUse1000 name=chkUse1000 tabindex=-1
                                                                           onclick="changeTab2Control()" 																    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                    <label 
                                                                        for="chkUse1000"
                                                                        class="checkbox"
                                                                        tabindex=0 
                                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																    
                                                                        Use 1000 Separators (,)
                                                                    </label> 
                                                                </td>
                                                            </tr>													
                                                        </TABLE>
                                                    </td>
                                                    <tr>
                                                        <td colspan=9 height=5></td>
                                                    </tr>
                                                </TABLE>
                                            </td>
                                        </tr>
                                    </TABLE>
                                </DIV>

                                <!-- OUTPUT OPTIONS -->
                                <DIV id=div3 style="visibility:hidden;display:none">
                                <TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
                                    <tr valign=top> 
                                        <td>
                                            <TABLE WIDTH="100%" class="invisible" CELLSPACING=10 CELLPADDING=0>
                                                <tr>						
                                                    <td valign=top rowspan=2 width=25% height="100%">
                                                        <table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
                                                            <tr height=10> 
                                                                <td height=10 align=left valign=top>
                                                                    Output Format : <BR><BR>
                                                                                        <TABLE class="invisible" cellspacing="0" cellpadding="0" width="100%">
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat0 value=0
                                                                                                           onClick="formatClick(0);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat0"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Data Only
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat1 value=1
                                                                                                           onClick="formatClick(1);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat1"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    CSV File
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat2 value=2
                                                                                                           onClick="formatClick(2);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat2"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    HTML Document
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat3 value=3
                                                                                                           onClick="formatClick(3);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat3"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Word Document
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=20>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat4 value=4
                                                                                                           onClick="formatClick(4);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td align=left nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat4"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Excel Worksheet
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=5>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat5 value=5
                                                                                                           onClick="formatClick(5);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat5"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Excel Chart
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=10> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                            <tr height=5>
                                                                                                <td width=5>&nbsp</td>
                                                                                                <td align=left width=15>
                                                                                                    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat6 value=6
                                                                                                           onClick="formatClick(6);" 
                                                                                                           onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                                           onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                           onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                           onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                </td>
                                                                                                <td nowrap>
                                                                                                    <label 
                                                                                                        tabindex="-1"
                                                                                                        for="optOutputFormat6"
                                                                                                        class="radio"
                                                                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                                        />
                                                                                                    Excel Pivot Table
                                                                                                </label>
                                                                                                </td>
                                                                                                <td width=5>&nbsp</td>
                                                                                            </tr>
                                                                                            <tr height=5> 
                                                                                                <td colspan=4></td>
                                                                                            </tr>
                                                                                        </TABLE>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                    <td valign=top width="75%">
                                                        <table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
                                                            <tr height=10> 
                                                                <td height=10 align=left valign=top>
                                                                    Output Destination(s) : <BR><BR>
                                                                                                <TABLE class="invisible" cellspacing="0" cellpadding="0" width="100%">
                                                                                                    <tr height=20>
                                                                                                        <td width=5>&nbsp</td>
                                                                                                        <td align=left colspan=6 nowrap>
                                                                                                            <input name=chkPreview id=chkPreview type=checkbox disabled="disabled" tabindex=-1 
                                                                                                                   onClick="changeTab3Control();"
                                                                                                                   onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                                                                   onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                                            <label 
                                                                                                                for="chkPreview"
                                                                                                                class="checkbox checkboxdisabled"
                                                                                                                tabindex=0 
                                                                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                                                                Preview on screen
                                                                                                            </label>
                                                                                                        </td>
                                                                                                        <td width=5>&nbsp</td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=10> 
                                                                                                        <td colspan=8></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=20>
                                                                                                        <td></td>
                                                                                                        <td align=left colspan=6 nowrap>
                                                                                                            <input name=chkDestination0 id=chkDestination0 type=checkbox disabled="disabled" tabindex=-1 
                                                                                                                   onClick="changeTab3Control();"
                                                                                                                   onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                                                                   onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                                            <label 
                                                                                                                for="chkDestination0"
                                                                                                                class="checkbox checkboxdisabled"
                                                                                                                tabindex=0 
                                                                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                                                                Display output on screen 
                                                                                                            </label>
                                                                                                        </td>
                                                                                                        <td></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=10> 
                                                                                                        <td colspan=8></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=20>
                                                                                                        <td></td>
                                                                                                        <td align=left nowrap>
                                                                                                            <input name=chkDestination1 id=chkDestination1 type=checkbox disabled="disabled" tabindex=-1  
                                                                                                                   onClick="changeTab3Control();"
                                                                                                                   onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                                                                   onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                                            <label 
                                                                                                                for="chkDestination1"
                                                                                                                class="checkbox checkboxdisabled"
                                                                                                                tabindex=0 
                                                                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                                                                Send to printer 
                                                                                                            </label>
                                                                                                        </td>
                                                                                                        <td width=30 nowrap>&nbsp</td>
                                                                                                        <td align=left nowrap>
                                                                                                            Printer location : 
                                                                                                        </td>
                                                                                                        <td width=15>&nbsp</td>
                                                                                                        <td colspan=2>
                                                                                                            <select id=cboPrinterName name=cboPrinterName width=100% style="WIDTH: 400px" class="combo"
                                                                                                                    onchange="changeTab3Control()">	
                                                                                                            </select>								
                                                                                                        </td>
                                                                                                        <td></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=10> 
                                                                                                        <td colspan=8></td>
                                                                                                    </tr>
																	
                                                                                                    <tr height=20>
                                                                                                    <td></td>
                                                                                                    <td align=left nowrap>
                                                                                                        <input name=chkDestination2 id=chkDestination2 type=checkbox disabled="disabled" tabindex=-1 
                                                                                                               onClick="changeTab3Control();"
                                                                                                               onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                                                               onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                                        <label 
                                                                                                            for="chkDestination2"
                                                                                                            class="checkbox checkboxdisabled"
                                                                                                            tabindex=0 
                                                                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                                                            Save to file
                                                                                                        </label>
                                                                                                    </td>
                                                                                                    <td></td>
                                                                                                    <td align=left nowrap>
                                                                                                        File name :   
                                                                                                    </td>
                                                                                                    <td></td>
                                                                                                    <td colspan=2>
                                                                                                        <TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 style="WIDTH: 400px">
                                                                                                        <tr>
                                                                                                        <td>
                                                                                                            <INPUT id=txtFilename name=txtFilename class="text textdisabled" disabled="disabled" style="WIDTH: 375px">
                                                                                                        </td>
                                                                                                        <td width=25>
                                                                                                            <INPUT id=cmdFilename name=cmdFilename style="WIDTH: 100%" type=button class="btn" value="..."
                                                                                                                   onClick="saveFile();changeTab3Control();"  			                                
                                                                                                                   onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                                                                   onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                                                                   onfocus="try{button_onFocus(this);}catch(e){}"
                                                                                                                   onblur="try{button_onBlur(this);}catch(e){}" />
                                                                                                        </td>
                                                                                                    </td>
                                                                                                </TABLE>
                                                                </td>
                                                                <td></td>
                                                            </tr>
																	
                                                            <tr height=10> 
                                                                <td colspan=8></td>
                                                            </tr>
																	
                                                            <tr height=20>
                                                                <td colspan=3></td>
                                                                <td align=left nowrap>
                                                                    If existing file :
                                                                </td>
                                                                <td></td>
                                                                <td colspan=2 width=100% nowrap>
                                                                    <select id=cboSaveExisting name=cboSaveExisting width=100% style="WIDTH: 400px" class="combo"
                                                                            onchange="changeTab3Control()">
                                                                    </select>							
                                                                </td>
                                                                <td></td>
                                                            </tr>
																	
                                                            <tr height=10> 
                                                                <td colspan=8></td>
                                                            </tr>
																	
                                                            <tr height=20>
                                                            <td></td>
                                                            <td align=left nowrap>
                                                                <input name=chkDestination3 id=chkDestination3 type=checkbox disabled="disabled" tabindex=-1 
                                                                       onClick="changeTab3Control();"
                                                                       onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                       onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                <label 
                                                                    for="chkDestination3"
                                                                    class="checkbox checkboxdisabled"
                                                                    tabindex=0 
                                                                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                                    onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                    onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                                    Send as email 
                                                                </label>
                                                            </td>
                                                            <td></td>
                                                            <td align=left nowrap>
                                                                Email group :   
                                                            </td>
                                                            <td></td>
                                                            <td colspan=2>
                                                                <TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 style="WIDTH: 400px">
                                                                <tr>
                                                                <td>
                                                                    <INPUT id=txtEmailGroup name=txtEmailGroup class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
                                                                    <INPUT id=txtEmailGroupID name=txtEmailGroupID type=hidden>
                                                                </td>
                                                                <td width=25>
                                                                    <INPUT id=cmdEmailGroup name=cmdEmailGroup style="WIDTH: 100%" type=button class="btn" value="..."
                                                                           onClick="selectEmailGroup();changeTab3Control();" 
                                                                           onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                                           onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                                           onfocus="try{button_onFocus(this);}catch(e){}"
                                                                           onblur="try{button_onBlur(this);}catch(e){}" />
                                                                </td>
                                                            </td>
                                                        </TABLE>
                                                    </td>
                                                    <td></td>
                                                </tr>
																	
                                                <tr height=10> 
                                                    <td colspan=8></td>
                                                </tr>
																	
                                                <tr height=20>
                                                    <td colspan=3></td>
                                                    <td align=left nowrap>
                                                        Email subject :   
                                                    </td>
                                                    <td></td>
                                                    <td colspan=2 width=100% nowrap>
                                                        <INPUT id=txtEmailSubject disabled="disabled" class="text textdisabled" maxlength=255 name=txtEmailSubject style="WIDTH: 400px" 
                                                               onchange="frmUseful.txtChanged.value = 1;" 
                                                               onkeydown="frmUseful.txtChanged.value = 1;">
                                                    </td>
                                                    <td></td>
                                                </tr>
																	
                                                <tr height=10>
                                                    <td colspan=8></td>
                                                </tr>
																	
                                                <tr height=20>
                                                    <td colspan=3></td>
                                                    <td align=left nowrap>
                                                        Attach as :   
                                                    </td>
                                                    <td></td>
                                                    <td colspan=2 width=100% nowrap>
                                                        <INPUT id=txtEmailAttachAs disabled="disabled" maxlength=255 class="text textdisabled" name=txtEmailAttachAs style="WIDTH: 400px" 
                                                               onchange="frmUseful.txtChanged.value = 1;" 
                                                               onkeydown="frmUseful.txtChanged.value = 1;">
                                                    </td>
                                                    <td></td>
                                                </tr>
																	
                                                <tr height=10>
                                                    <td colspan=8></td>
                                                </tr>
                                            </TABLE>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </TABLE>
                </td>
            </tr>
        </TABLE></form>
</DIV>
    </td>
        <td width=10></td>
    </tr> 

        <tr height=10> 
            <td colspan=3></td>
        </tr> 

        <tr height=10>
            <td width=10></td>
            <td>
                <TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
                    <tr>
                        <td>&nbsp;</td>
                        <td width=80>
                            <input type=button id=cmdOK name=cmdOK value=OK style="WIDTH: 100%" class="btn"
                                   onclick="okClick()"
                                   onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                   onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                   onfocus="try{button_onFocus(this);}catch(e){}"
                                   onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                        <td width=10></td>
                        <td width=80>
                            <input type=button id=cmdCancel name=cmdCancel value=Cancel style="WIDTH: 100%"  class="btn"
                                   onclick="cancelClick()"
                                   onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                   onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                   onfocus="try{button_onFocus(this);}catch(e){}"
                                   onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                    </tr>
                </TABLE>
            </td>
            <td width=10></td>
        </tr> 

        <tr height=5> 
            <td colspan=3></td>
        </tr> 
    </table>
    </td>
    </tr> 
    </table>

        <input type='hidden' id=txtBasePicklistID name=txtBasePicklistID>
        <input type='hidden' id=txtBaseFilterID name=txtBaseFilterID>
        <input type='hidden' id=txtDatabase name=txtDatabase value="<%=session("Database")%>">

        <input type='hidden' id=txtWordVer name=txtWordVer value="<%=Session("WordVer")%>">
        <input type='hidden' id=txtExcelVer name=txtExcelVer value="<%=Session("ExcelVer")%>">
        <input type='hidden' id=txtWordFormats name=txtWordFormats value="<%=Session("WordFormats")%>">
        <input type='hidden' id=txtExcelFormats name=txtExcelFormats value="<%=Session("ExcelFormats")%>">
        <input type='hidden' id=txtWordFormatDefaultIndex name=txtWordFormatDefaultIndex value="<%=Session("WordFormatDefaultIndex")%>">
        <input type='hidden' id=txtExcelFormatDefaultIndex name=txtExcelFormatDefaultIndex value="<%=Session("ExcelFormatDefaultIndex")%>">

    </form>

    <form id=frmAccess>
        <%
            sErrorDescription = ""
	
            ' Get the table records.
            Dim cmdAccess = CreateObject("ADODB.Command")
            cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
            cmdAccess.CommandType = 4 ' Stored Procedure
            cmdAccess.ActiveConnection = Session("databaseConnection")

            Dim prmUtilType = cmdAccess.CreateParameter("utilType", 3, 1) ' 3=integer, 1=input
            cmdAccess.Parameters.Append(prmUtilType)
            prmUtilType.value = 1 ' 1 = cross tabs

            Dim prmUtilID = cmdAccess.CreateParameter("utilID", 3, 1) ' 3=integer, 1=input
            cmdAccess.Parameters.Append(prmUtilID)
            If UCase(Session("action")) = "NEW" Then
                prmUtilID.value = 0
            Else
                prmUtilID.value = CleanNumeric(Session("utilid"))
            End If

            Dim prmFromCopy = cmdAccess.CreateParameter("fromCopy", 3, 1) ' 3=integer, 1=input
            cmdAccess.Parameters.Append(prmFromCopy)
            If UCase(Session("action")) = "COPY" Then
                prmFromCopy.value = 1
            Else
                prmFromCopy.value = 0
            End If

            Err.Clear()
            Dim rstAccessInfo = cmdAccess.Execute
            If (Err.Number <> 0) Then
                sErrorDescription = "The access information could not be retrieved." & vbCrLf & FormatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
                Dim iCount = 0
                Do While Not rstAccessInfo.EOF
                    Response.Write("<INPUT type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & rstAccessInfo.fields("accessDefinition").value & """>" & vbCrLf)

                    iCount = iCount + 1
                    rstAccessInfo.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstAccessInfo.close()
                rstAccessInfo = Nothing
            End If
	
            ' Release the ADO command object.
            cmdAccess = Nothing
%>
    </form>

    <FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
        <INPUT type="hidden" id=txtUserName name=txtUserName value="<%=session("username")%>">
        <INPUT type="hidden" id=txtLoading name=txtLoading value="Y">
        <INPUT type="hidden" id=txtCurrentBaseTableID name=txtCurrentBaseTableID>
        <INPUT type="hidden" id=txtCurrentHColID name=txtCurrentHColID>
        <INPUT type="hidden" id=txtCurrentVColID name=txtCurrentVColID>
        <INPUT type="hidden" id=txtCurrentPColID name=txtCurrentPColID>
        <INPUT type="hidden" id=txtCurrentIColID name=txtCurrentIColID>
        <INPUT type="hidden" id=txtTablesChanged name=txtTablesChanged>
        <INPUT type="hidden" id=txtSelectedColumnsLoaded name=txtSelectedColumnsLoaded value=0>
        <INPUT type="hidden" id=txtSortLoaded name=txtSortLoaded value=0>
        <INPUT type="hidden" id=txtSecondTabShown name=txtSecondTabShown value=0>
        <INPUT type="hidden" id=txtRepetitionLoaded name=txtRepetitionLoaded value=0>
        <INPUT type="hidden" id=txtChanged name=txtChanged value=0>
        <INPUT type="hidden" id=txtUtilID name=txtUtilID value=<%=session("utilid")%>>
        <%
            Dim cmdDefinition = CreateObject("ADODB.Command")
            cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
            cmdDefinition.CommandType = 4 ' Stored procedure.
            cmdDefinition.ActiveConnection = Session("databaseConnection")

            Dim prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
            cmdDefinition.Parameters.Append(prmModuleKey)
            prmModuleKey.value = "MODULE_PERSONNEL"

            Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
            cmdDefinition.Parameters.Append(prmParameterKey)
            prmParameterKey.value = "Param_TablePersonnel"

            Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000) '200=varchar, 2=output, 8000=size
            cmdDefinition.Parameters.Append(prmParameterValue)

            Err.Clear()
            cmdDefinition.Execute()

            Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
            cmdDefinition = Nothing

            Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
            Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
%>
    </FORM>

    <FORM id=frmValidate name=frmValidate target=validate method=post action=util_validate_crosstab style="visibility:hidden;display:none">
        <INPUT type=hidden id="validateBaseFilter" name=validateBaseFilter value=0>
        <INPUT type=hidden id="validateBasePicklist" name=validateBasePicklist value=0>
        <INPUT type=hidden id="validateEmailGroup" name=validateEmailGroup value=0>
        <INPUT type=hidden id="validateCalcs" name=validateCalcs value = ''>
        <INPUT type=hidden id="validateHiddenGroups" name=validateHiddenGroups value = ''>
        <INPUT type=hidden id="validateName" name=validateName value=''>
        <INPUT type=hidden id="validateTimestamp" name=validateTimestamp value=''>
        <INPUT type=hidden id="validateUtilID" name=validateUtilID value=''>
    </FORM>

    <FORM id=frmSend name=frmSend method=post action=util_def_crosstabs_Submit style="visibility:hidden;display:none">
        <INPUT type="hidden" id=txtSend_ID name=txtSend_ID value=0>
        <INPUT type="hidden" id=txtSend_name name=txtSend_name value=''>
        <INPUT type="hidden" id=txtSend_description name=txtSend_description value=''>
        <INPUT type="hidden" id=txtSend_baseTable name=txtSend_baseTable value=0>
        <INPUT type="hidden" id=txtSend_allRecords name=txtSend_allRecords value=0>
        <INPUT type="hidden" id=txtSend_picklist name=txtSend_picklist value=0>
        <INPUT type="hidden" id=txtSend_filter name=txtSend_filter value=0>
        <INPUT type="hidden" id=txtSend_PrintFilter name=txtSend_PrintFilter value=0>
        <INPUT type="hidden" id=txtSend_access name=txtSend_access value=''>
        <INPUT type="hidden" id=txtSend_userName name=txtSend_userName value=''>

        <INPUT type="hidden" id=txtSend_HColID name=txtSend_HColID value=0>
        <INPUT type="hidden" id=txtSend_HStart name=txtSend_HStart value=''>
        <INPUT type="hidden" id=txtSend_HStop name=txtSend_HStop value=''>
        <INPUT type="hidden" id=txtSend_HStep name=txtSend_HStep value=''>
        <INPUT type="hidden" id=txtSend_VColID name=txtSend_VColID value=0>
        <INPUT type="hidden" id=txtSend_VStart name=txtSend_VStart value=''>
        <INPUT type="hidden" id=txtSend_VStop name=txtSend_VStop value=''>
        <INPUT type="hidden" id=txtSend_VStep name=txtSend_VStep value=''>
        <INPUT type="hidden" id=txtSend_PColID name=txtSend_PColID value=0>
        <INPUT type="hidden" id=txtSend_PStart name=txtSend_PStart value=''>
        <INPUT type="hidden" id=txtSend_PStop name=txtSend_PStop value=''>
        <INPUT type="hidden" id=txtSend_PStep name=txtSend_PStep value=''>
        <INPUT type="hidden" id=txtSend_IType name=txtSend_IType value=0>
        <INPUT type="hidden" id=txtSend_IColID name=txtSend_IColID value=0>
        <INPUT type="hidden" id=txtSend_Percentage name=txtSend_Percentage value=0>
        <INPUT type="hidden" id=txtSend_PerPage name=txtSend_PerPage value=0>
        <INPUT type="hidden" id=txtSend_Suppress name=txtSend_Suppress value=0>
        <INPUT type="hidden" id=txtSend_Use1000Separator name=txtSend_Use1000Separator value=0>

        <INPUT type="hidden" id=txtSend_OutputPreview name=txtSend_OutputPreview>
        <INPUT type="hidden" id=txtSend_OutputFormat name=txtSend_OutputFormat>
        <INPUT type="hidden" id=txtSend_OutputScreen name=txtSend_OutputScreen>
        <INPUT type="hidden" id=txtSend_OutputPrinter name=txtSend_OutputPrinter>
        <INPUT type="hidden" id=txtSend_OutputPrinterName name=txtSend_OutputPrinterName>
        <INPUT type="hidden" id=txtSend_OutputSave name=txtSend_OutputSave>
        <INPUT type="hidden" id=txtSend_OutputSaveExisting name=txtSend_OutputSaveExisting>
        <INPUT type="hidden" id=txtSend_OutputEmail name=txtSend_OutputEmail>
        <INPUT type="hidden" id=txtSend_OutputEmailAddr name=txtSend_OutputEmailAddr>
        <INPUT type="hidden" id=txtSend_OutputEmailSubject name=txtSend_OutputEmailSubject>
        <INPUT type="hidden" id=txtSend_OutputEmailAttachAs name=txtSend_OutputEmailAttachAs>
        <INPUT type="hidden" id=txtSend_OutputFilename name=txtSend_OutputFilename>
	
        <INPUT type="hidden" id=txtSend_reaction name=txtSend_reaction>

        <INPUT type="hidden" id=txtSend_jobsToHide name=txtSend_jobsToHide>
        <INPUT type="hidden" id=txtSend_jobsToHideGroups name=txtSend_jobsToHideGroups>
    </FORM>

    <FORM id=frmRecordSelection name=frmRecordSelection target="recordSelection" action="util_recordSelection" method=post style="visibility:hidden;display:none">
        <INPUT type="hidden" id=recSelType name=recSelType>
        <INPUT type="hidden" id=recSelTableID name=recSelTableID>
        <INPUT type="hidden" id=recSelCurrentID name=recSelCurrentID>
        <INPUT type="hidden" id=recSelTable name=recSelTable>
        <INPUT type="hidden" id=recSelDefOwner name=recSelDefOwner>
        <INPUT type="hidden" id=recSelDefType name=recSelDefType>
    </FORM>

    <FORM id=frmEmailSelection name=frmEmailSelection target="emailSelection" action="util_emailSelection" method=post style="visibility:hidden;display:none">
        <INPUT type="hidden" id=EmailSelCurrentID name=EmailSelCurrentID>
    </FORM>

    <FORM id=frmSelectionAccess name=frmSelectionAccess style="visibility:hidden;display:none">
        <INPUT type="hidden" id=forcedHidden name=forcedHidden value="N">
        <INPUT type="hidden" id=baseHidden name=baseHidden value="N">

        <!-- need the count of hidden child filter access info -->
        <INPUT type="hidden" id=childHidden name=childHidden value="N">
        <INPUT type="hidden" id=calcsHiddenCount name=calcsHiddenCount value=0>
    </FORM>

    <FORM action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
        <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
    </FORM>

    <INPUT type='hidden' id=txtTicker name=txtTicker value=0>
    <INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<script type="text/javascript">
    util_def_crosstabs_window_onload();
    util_def_crosstabs_addhandlers();
</script>
