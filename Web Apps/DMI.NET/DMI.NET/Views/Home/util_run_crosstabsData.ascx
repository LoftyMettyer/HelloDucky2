﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%		
    Dim objCrossTab As Object
    Dim intCount As Integer
    Dim strCrossTabName As String
    Dim lngCount
    Dim objUser
    Dim lngLoopMin As Long
    Dim lngLoopMax As Long
    Dim strEmailAddresses As String
    Dim cmdReportsCols
    Dim prmEmailGroupID
    Dim rstReportColumns
    Dim iLoop As Integer
    Dim sErrorDescription As String
    
    
    Response.Write("<script type=""text/javascript"">" & vbCrLf)
    Response.Write("  //" & Session("CT_Mode") & vbCrLf)
    Response.Write("  function util_run_crosstabs_data_window_onload() {" & vbCrLf & vbCrLf)
    Response.Write("    $(""#reportdataframe"").attr(""data-framesource"", ""UTIL_RUN_CROSSTABSDATA"");" & vbCrLf & vbCrLf)

    Response.Write("    $(""#reportdataframe"").hide();" & vbCrLf & vbCrLf)
	Response.Write("    $(""#reportbreakdownframe"").hide();" & vbCrLf & vbCrLf)
    Response.Write("    $(""#reportworkframe"").show();" & vbCrLf & vbCrLf)
    
    Response.Write("    frmOriginalDefinition = OpenHR.getForm(""reportworkframe"",""frmOriginalDefinition"");" & vbCrLf)
    Response.Write("    frmExportData = OpenHR.getForm(""reportworkframe"",""frmExportData"");" & vbCrLf)
    Response.Write("    var ssOutputGrid;" & vbCrLf)
       
    If Session("CT_Mode") = "OUTPUTRUN" Or _
       Session("CT_Mode") = "OUTPUTRUNTHENCLOSE" Then
        Response.Write("    ssOutputGrid = document.getElementById(""ssHiddenGrid"");" & vbCrLf & vbCrLf)
    Else
        Response.Write("    ssOutputGrid = document.getElementById(""ssOutputGrid"");" & vbCrLf & vbCrLf)
    End If


    '**************************************
    ' LOAD
    '**************************************

    If Session("CT_Mode") = "LOAD" Then
        'Populate Controls

        objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))

        Response.Write("  AddToIntTypeCombo(""Count"",""0"");" & vbCrLf)
        Response.Write("  AddToIntTypeCombo(""Average"",""1"");" & vbCrLf)
        Response.Write("  AddToIntTypeCombo(""Maximum"",""2"");" & vbCrLf)
        Response.Write("  AddToIntTypeCombo(""Minimum"",""3"");" & vbCrLf)
        Response.Write("  AddToIntTypeCombo(""Total"",""4"");" & vbCrLf)
       
        For intCount = 0 To objCrossTab.ColumnHeadingUbound(2)
            Response.Write("  AddToPgbCombo(""" & CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(2, intCount), 255)) & """,""" & CStr(intCount) & """);" & vbCrLf)
        Next

        Response.Write("  txtPageColumn.value = """ & CleanStringForJavaScript(objCrossTab.PageBreakColumnName) & """;" & vbCrLf)
        Response.Write("  chkPercentType.checked = " & LCase(CStr(objCrossTab.ShowPercentage)) & ";" & vbCrLf)
        Response.Write("  chkPercentPage.checked = " & LCase(CStr(objCrossTab.PercentageOfPage)) & ";" & vbCrLf)
        Response.Write("  chkSuppressZeros.checked = " & LCase(CStr(objCrossTab.SuppressZeros)) & ";" & vbCrLf)
        Response.Write("  chkUse1000.checked = " & LCase(CStr(objCrossTab.Use1000Separator)) & ";" & vbCrLf)
		
        Response.Write("  txtIntersectionColumn.value = """ & CleanStringForJavaScript(objCrossTab.IntersectionColumnName) & """;" & vbCrLf)
        Response.Write("  cboIntersectionType.selectedIndex = " & CStr(objCrossTab.IntersectionType) & ";" & vbCrLf)
        'Response.Write "  frmWorkFrame.cboFileFormat.selectedIndex = " & cstr(objCrossTab.DefaultExportTo) & ";" & vbcrlf
        'Response.Write "  frmWorkFrame.txtFile.value = """ & Replace(objCrossTab.DefaultSaveAs,"\","\\") & """;" & vbcrlf
        'Response.Write "  frmWorkFrame.chkCloseApp.checked = " & lcase(cstr(objCrossTab.DefaultCloseApp)) & ";" & vbcrlf & vbcrlf
        'Response.Write "  frmWorkFrame.refreshScreen();" & vbcrlf

        If objCrossTab.PageBreakColumn = True Then
            Response.Write("  cboPage.selectedIndex = 0;" & vbCrLf)
        Else
            Response.Write("  control_disable(cboPage, true);" & vbCrLf)
        End If

        Response.Write("  frmExportData.txtPreview.value = """ & objCrossTab.OutputPreview & """;" & vbCrLf)
        Response.Write("  frmExportData.txtFormat.value = """ & objCrossTab.OutputFormat & """;" & vbCrLf)
        Response.Write("  frmExportData.txtScreen.value = """ & objCrossTab.OutputScreen & """;" & vbCrLf)
        Response.Write("  frmExportData.txtPrinter.value = """ & objCrossTab.OutputPrinter & """;" & vbCrLf)
        Response.Write("  frmExportData.txtPrinterName.value = """ & CleanStringForJavaScript(objCrossTab.OutputPrinterName) & """;" & vbCrLf)
        Response.Write("  frmExportData.txtSave.value = """ & objCrossTab.OutputSave & """;" & vbCrLf)
        Response.Write("  frmExportData.txtSaveExisting.value = """ & objCrossTab.OutputSaveExisting & """;" & vbCrLf)
        Response.Write("  frmExportData.txtEmail.value = """ & objCrossTab.OutputEmail & """;" & vbCrLf)
        Response.Write("  frmExportData.txtEmailAddr.value = " & objCrossTab.OutputEmailID & ";" & vbCrLf)
        Response.Write("  frmExportData.txtEmailAddrName.value = """ & CleanStringForJavaScript(objCrossTab.OutputEmailGroupName) & """;" & vbCrLf)
        Response.Write("  frmExportData.txtEmailSubject.value = """ & CleanStringForJavaScript(objCrossTab.OutputEmailSubject) & """;" & vbCrLf)
        Response.Write("  frmExportData.txtEmailAttachAs.value = """ & CleanStringForJavaScript(objCrossTab.OutputEmailAttachAs) & """;" & vbCrLf)
        Response.Write("  frmExportData.txtFileName.value = """ & CleanStringForJavaScript(objCrossTab.OutputFileName) & """;" & vbCrLf)
    End If

    '**************************************
    ' LOAD / PRINT
    '**************************************

    If Session("CT_Mode") = "LOAD" Or _
       Session("CT_Mode") = "OUTPUTRUN" Or _
       Session("CT_Mode") = "OUTPUTRUNTHENCLOSE" Then
        'Initalise Grid

        objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
        strCrossTabName = CleanStringForJavaScript(Replace(objCrossTab.CrossTabName, "&", "&&"))

        'Response.Write "  if (ssOutputGrid.Caption = """") {" & vbcrlf
        'Response.Write "    ssOutputGrid.Redraw = false;" & vbcrlf
        Response.Write("    ssOutputGrid.Caption = """ & strCrossTabName & """;" & vbCrLf)
        Response.Write("    ssOutputGrid.focus();" & vbCrLf)
		
        Response.Write("    ssOutputGrid.Columns.RemoveAll();" & vbCrLf & vbCrLf)
        Response.Write("    ssOutputGrid.Columns.Add(0);" & vbCrLf)
        Response.Write("    ssOutputGrid.Columns(0).Caption = """";" & vbCrLf)
        Response.Write("    ssOutputGrid.Columns(0).Locked = true;" & vbCrLf)
        Response.Write("    ssOutputGrid.Columns(0).Visible = true;" & vbCrLf)
        Response.Write("    ssOutputGrid.Columns(0).Alignment = 1;" & vbCrLf)
        Response.Write("    ssOutputGrid.Columns(0).Style = 4;" & vbCrLf)
        Response.Write("    ssOutputGrid.Columns(0).ButtonsAlways = 1;" & vbCrLf)
        Response.Write("    ssOutputGrid.Columns(0).BackColor = -2147483633;" & vbCrLf & vbCrLf)

        For lngCount = 0 To objCrossTab.ColumnHeadingUbound(0)
            Response.Write("    ssOutputGrid.Columns.Add(" & CStr(lngCount + 1) & ");" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").Caption = """ & CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(CLng(0), lngCount), 255)) & """;" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").Locked = true;" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").Visible = true;" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").Alignment = 1;" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").CaptionAlignment = 2;" & vbCrLf)

            If objCrossTab.CrossTabType = 3 Then
                Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").Width = 80;" & vbCrLf)
            End If
			
        Next

        If objCrossTab.CrossTabType <> 3 Then
            lngCount = objCrossTab.ColumnHeadingUbound(0) + 1
            Response.Write("    ssOutputGrid.Columns.Add(" & CStr(lngCount + 1) & ");" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").Locked = true;" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").Visible = true;" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").Alignment = 1;" & vbCrLf)
            Response.Write("    ssOutputGrid.Columns(" & CStr(lngCount + 1) & ").CaptionAlignment = 2;" & vbCrLf & vbCrLf)

            Response.Write("    ssOutputGrid.SplitterPos = 1;" & vbCrLf)
            Response.Write("    ssOutputGrid.SplitterVisible = false;" & vbCrLf & vbCrLf)
            'Response.Write "  }" & vbcrlf

            Session("CT_ShowPercentage") = objCrossTab.ShowPercentage
            Session("CT_PercentageOfPage") = objCrossTab.PercentageOfPage
            Session("CT_SuppressZeros") = objCrossTab.SuppressZeros
            Session("CT_IntersectionType") = objCrossTab.IntersectionType
            Session("CT_Use1000") = objCrossTab.Use1000Separator

            If Session("CT_Mode") = "LOAD" Then
                Session("CT_PageNumber") = 0
            End If
        End If

    End If


    '**************************************
    ' LOAD / REFRESH
    '**************************************

    If Session("CT_Mode") = "LOAD" Or _
       Session("CT_Mode") = "REFRESH" Then
        'PopulateScreen

        objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
        objCrossTab.IntersectionType = Session("CT_IntersectionType")
        objCrossTab.ShowPercentage = Session("CT_ShowPercentage")
        objCrossTab.PercentageOfPage = Session("CT_PercentageOfPage")
        objCrossTab.SuppressZeros = Session("CT_SuppressZeros")
        objCrossTab.Use1000Separator = Session("CT_Use1000")
        objCrossTab.BuildOutputStrings(CLng(Session("CT_PageNumber")))

        If objCrossTab.IntersectionColumn = True Then
            Response.Write("  control_disable(cboIntersectionType, false);" & vbCrLf)
        Else
            Response.Write("  cboIntersectionType.style.backgroundColor = 'threedface';" & vbCrLf)
        End If

        If objCrossTab.PageBreakColumn = True Then
            Response.Write("  control_disable(cboPage, false);" & vbCrLf)
            If objCrossTab.ShowPercentage = True Then
                Response.Write("  control_disable(chkPercentPage, false);" & vbCrLf)
            End If
        End If

        Response.Write("  control_disable(chkPercentType, false);" & vbCrLf)
        Response.Write("  control_disable(chkSuppressZeros, false);" & vbCrLf)
        Response.Write("  control_disable(chkUse1000, false);" & vbCrLf)

        'Response.Write("  ssOutputGrid.Columns.RemoveAll();" & vbCrLf & vbCrLf)
        
        Response.Write("  ssOutputGrid.Redraw = false;" & vbCrLf & vbCrLf)
        Response.Write("  lngCol = ssOutputGrid.LeftCol;" & vbCrLf)
        Response.Write("  lngRow = ssOutputGrid.FirstRow;" & vbCrLf)

        Response.Write("  ssOutputGrid.Columns(ssOutputGrid.Columns.Count-1).Caption = cboIntersectionType.options[cboIntersectionType.selectedIndex].text;" & vbCrLf)
		
        '		Response.Write "  while (ssOutputGrid.Rows > 0) {"
        '		Response.Write "    ssOutputGrid.RemoveAll(); }"  & vbcrlf & vbcrlf
        '		Response.Write "  ssOutputGrid.Redraw = true;" & vbcrlf & vbcrlf
        
        For intCount = 1 To objCrossTab.OutputArrayDataUBound
            Response.Write("  ssOutputGrid.Additem(""" & CleanStringForJavaScript(Left(objCrossTab.OutputArrayData(intCount), 255)) & """);" & vbCrLf)
        Next

        Response.Write("  ssOutputGrid.RowHeight = 20;" & vbCrLf)
        Response.Write("  ssOutputGrid.LeftCol = lngCol;" & vbCrLf)
        Response.Write("  ssOutputGrid.FirstRow = lngRow;" & vbCrLf)
        Response.Write("  ssOutputGrid.Redraw = true;" & vbCrLf & vbCrLf)

        ' JDM - Fault 4849 - Disable these controls in absence breakdown mode
        If objCrossTab.CrossTabType = 3 Then
            Response.Write("  CrossTabPage.style.visibility = ""hidden"" ;" & vbCrLf)
            Response.Write("  chkPercentPage.style.visibility = ""hidden"" ;" & vbCrLf)
            Response.Write("  chkPercentType.style.visibility = ""hidden"" ;" & vbCrLf)
            Response.Write("  chkUse1000.style.visibility = ""hidden"" ;" & vbCrLf)
        End If

    End If


    If Session("CT_Mode") = "LOAD" Then
        If objCrossTab.OutputPreview = False Then
            'Response.Write "    if (frmExportData.txtFormat.value == ""0"") {" & vbcrlf
            'Response.Write "      if (frmExportData.txtPrinter.value == ""1"") {" & vbcrlf
            'Response.Write "        ExportData(""OUTPUTRUNTHENCLOSE"", frmExportData.txtEmailAddr.value);" & vbcrlf
            Response.Write("        frmGetData.txtEmailGroupID.value = frmExportData.txtEmailAddr.value;" & vbCrLf)
            Response.Write("        ExportData(""EMAILGROUPTHENCLOSE"");" & vbCrLf)
            'Response.Write "      }" & vbcrlf
            'Response.Write "    }" & vbcrlf
        End If
    End If


'**************************************
' BREAKDOWN
'**************************************

If Session("CT_Mode") = "BREAKDOWN" Then

        objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
        objCrossTab.BuildBreakdownStrings(CLng(Session("CT_Hor")), CLng(Session("CT_Ver")), CLng(Session("CT_Pgb")))

        'Look up the Int Type Text from the Int Type Number...
        Response.Write("  var frmBreakdown = OpenHR.getForm(""dataframe"", ""frmBreakdown"");" & vbCrLf)
		' Response.Write("  frmBreakdown.txtIntersectionType.value = cboIntersectionType.options[frmBreakdown.txtIntersectionType.value].innerText;" & vbCrLf)
		Response.Write("  document.getElementById('txtDataIntersectionType').value = cboIntersectionType.options[document.getElementById('txtDataIntersectionType').value].innerText;" & vbCrLf)
		
		Response.Write("  OpenHR.submitForm(frmBreakdown, null, false);" & vbCrLf)


'**************************************
' OUTPUTPROMPT
'**************************************

ElseIf Session("CT_Mode") = "OUTPUTPROMPT" Then

        Response.Write("			sURL = ""util_run_outputoptions"" +" & vbCrLf & _
            """?txtUtilType="" + escape(frmExportData.txtUtilType.value) +" & vbCrLf & _
            """&txtPreview="" + escape(frmExportData.txtPreview.value) +" & vbCrLf & _
            """&txtFormat="" + escape(frmExportData.txtFormat.value) + " & vbCrLf & _
            """&txtScreen="" + escape(frmExportData.txtScreen.value) +" & vbCrLf & _
            """&txtPrinter="" + escape(frmExportData.txtPrinter.value) +" & vbCrLf & _
            """&txtPrinterName="" + escape(frmExportData.txtPrinterName.value) +" & vbCrLf & _
            """&txtSave="" + escape(frmExportData.txtSave.value) +" & vbCrLf & _
            """&txtSaveExisting="" + escape(frmExportData.txtSaveExisting.value) +" & vbCrLf & _
            """&txtEmail="" + escape(frmExportData.txtEmail.value) +" & vbCrLf & _
            """&txtEmailAddr="" + escape(frmExportData.txtEmailAddr.value) +" & vbCrLf & _
            """&txtEmailAddrName="" + escape(frmExportData.txtEmailAddrName.value) +" & vbCrLf & _
            """&txtEmailSubject="" + escape(frmExportData.txtEmailSubject.value) +" & vbCrLf & _
            """&txtEmailAttachAs="" + escape(frmExportData.txtEmailAttachAs.value) +" & vbCrLf & _
            """&txtFileName="" + escape(frmExportData.txtFileName.value);" & vbCrLf & _
            "  ShowOutputOptionsFrame(sURL);" & vbCrLf)

'**************************************
' OUTPUTRUN
'**************************************

ElseIf Session("CT_Mode") = "OUTPUTRUN" Or _
       Session("CT_Mode") = "OUTPUTRUNTHENCLOSE" Then

objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
objUser = Server.CreateObject("COAIntServer.clsSettings")

        Response.Write("  frmMenuFrame = OpenHR.getFrame(""menuframe"");" & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.UserName = """ & CleanStringForJavaScript(Session("Username")) & """;" & vbCrLf)
Response.Write("  window.parent.parent.ASRIntranetOutput.SaveAsValues = """ & CleanStringForJavaScript(Session("OfficeSaveAsValues")) & """;" & vbCrLf)
Response.Write("  window.parent.parent.ASRIntranetOutput.SettingOptions(")
Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "WordTemplate", "")) & """, ")
Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "ExcelTemplate", "")) & """, ")

If (objUser.GetUserSetting("Output", "ExcelGridlines", "0") = "1") Then
    Response.Write("true, ")
Else
    Response.Write("false, ")
End If

If (objUser.GetUserSetting("Output", "ExcelHeaders", "0") = "1") Then
    Response.Write("true, ")
Else
    Response.Write("false, ")
End If

If (objUser.GetUserSetting("Output", "AutoFitCols", "1") = "1") Then
    Response.Write("true, ")
Else
    Response.Write("false, ")
End If

If (objUser.GetUserSetting("Output", "Landscape", "1") = "1") Then
    Response.Write("true, " & vbCrLf)
Else
    Response.Write("false, " & vbCrLf)
End If

Response.Write("frmMenuFrame.document.all.item(""txtSysPerm_EMAILGROUPS_VIEW"").value);" & vbCrLf)


Response.Write("  window.parent.parent.ASRIntranetOutput.SettingLocations(")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleCol", "3")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleRow", "2")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataCol", "2")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataRow", "4")) & ");" & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.SettingTitle(")
If (objUser.GetUserSetting("Output", "TitleGridLines", "0") = "1") Then
    Response.Write("true, ")
Else
    Response.Write("false, ")
End If

If (objUser.GetUserSetting("Output", "TitleBold", "1") = "1") Then
    Response.Write("true, ")
Else
    Response.Write("false, ")
End If

If (objUser.GetUserSetting("Output", "TitleUnderline", "0") = "1") Then
    Response.Write("true, ")
Else
    Response.Write("false, ")
End If

Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215"))) & ", ")
Response.Write(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) & ");" & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.SettingHeading(")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingGridLines", "1")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBold", "1")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingUnderline", "0")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")) & ", ")
Response.Write(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")) & ", ")
Response.Write(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")) & ");" & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.SettingData(")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataGridLines", "1")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBold", "0")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataUnderline", "0")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")) & ", ")
Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataForecolour", "6697779")) & ", ")
Response.Write(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")) & ", ")
Response.Write(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataForecolour", "6697779")) & ");" & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
Response.Write("  window.parent.parent.ASRIntranetOutput.HeaderCols = 1;" & vbCrLf)
Response.Write("  fok = window.parent.parent.ASRIntranetOutput.SetOptions(false, " & _
    "frmExportData.txtFormat.value, frmExportData.txtScreen.value, " & _
    "frmExportData.txtPrinter.value, frmExportData.txtPrinterName.value, " & _
    "frmExportData.txtSave.value, frmExportData.txtSaveExisting.value, " & _
    "frmExportData.txtEmail.value, frmExportData.txtEmailGroupAddr.value, " & _
    "frmExportData.txtEmailSubject.value, frmExportData.txtEmailAttachAs.value, frmExportData.txtFileName.value);" & vbCrLf)

Response.Write("  if (fok == true) {" & vbCrLf)
Response.Write("  if (window.parent.parent.ASRIntranetOutput.GetFile() == true) {" & vbCrLf)

'DATA ONLY
Response.Write("  if (frmExportData.txtFormat.value == ""0"") {" & vbCrLf)

objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))

'if clng(Session("CT_PageNumber")) = -1 then
'All pages
lngLoopMin = 0
lngLoopMax = objCrossTab.ColumnHeadingUbound(2)
'else
'	'Current page
'	lngLoopMin = clng(Session("CT_PageNumber"))
'	lngLoopMax = clng(Session("CT_PageNumber"))
'end if

Response.Write("  frmOriginalDefinition.txtOptionsDone.value = 0;" & vbCrLf)
Response.Write("  frmOriginalDefinition.txtCancelPrint.value = 0;" & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.SetOptions(false, " & _
    "frmExportData.txtFormat.value, frmExportData.txtScreen.value, " & _
    "frmExportData.txtPrinter.value, frmExportData.txtPrinterName.value, " & _
    "frmExportData.txtSave.value, frmExportData.txtSaveExisting.value, " & _
    "frmExportData.txtEmail.value, """", " & _
    "frmExportData.txtEmailSubject.value, frmExportData.txtEmailAttachAs.value, frmExportData.txtFileName.value);")
Response.Write("  window.parent.parent.ASRIntranetOutput.SetPrinter();" & vbCrLf)

For lngCount = lngLoopMin To lngLoopMax

    If objCrossTab.PageBreakColumn = True Then
        Response.Write("	frmOriginalDefinition.txtCurrentPrintPage.value = "" (" & objCrossTab.PageBreakColumnName & _
            " : " & CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(2, lngCount), 255)) & ")"";" & vbCrLf)
    End If

    objCrossTab.BuildOutputStrings(lngCount)
    Response.Write("  ssOutputGrid.Redraw = false;" & vbCrLf & vbCrLf)
    Response.Write("  ssOutputGrid.Columns(ssOutputGrid.Columns.Count-1).Caption = frmWorkFrame.cboIntersectionType.options[frmWorkFrame.cboIntersectionType.selectedIndex].text;" & vbCrLf)
    Response.Write("  ssOutputGrid.RemoveAll();" & vbCrLf & vbCrLf)
    For intCount = 1 To objCrossTab.OutputArrayDataUBound
        Response.Write("  ssOutputGrid.Additem(""" & CleanStringForJavaScript(Left(objCrossTab.OutputArrayData(intCount), 255)) & """);" & vbCrLf)
    Next
    Response.Write("  ssOutputGrid.Redraw = true;" & vbCrLf & vbCrLf)

    Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
    Response.Write("    ssOutputGrid.redraw = true;" & vbCrLf)
    'Response.Write "    return;" & vbcrlf
    Response.Write("  }" & vbCrLf)
    Response.Write("  else if (frmOriginalDefinition.txtOptionsDone.value == 0) {" & vbCrLf)
    Response.Write("    button_disable(window.parent.parent.frames(""top"").frmPopup.Cancel, true);" & vbCrLf)
    Response.Write("    ssOutputGrid.PrintData(23,false,true);	" & vbCrLf)
    Response.Write("    ssOutputGrid.value = 1;" & vbCrLf)
    Response.Write("    try {" & vbCrLf)
    Response.Write("      button_disable(window.parent.parent.frames(""top"").frmPopup.Cancel, false);" & vbCrLf)
    Response.Write("      frmOriginalDefinition.txtOptionsDone.value = 1;" & vbCrLf)
    Response.Write("    }" & vbCrLf)
    Response.Write("    catch(e) {" & vbCrLf)
    Response.Write("    }" & vbCrLf)
    Response.Write("  }" & vbCrLf)
    Response.Write("  else {" & vbCrLf)
    Response.Write("    ssOutputGrid.PrintData(23,false,false);" & vbCrLf)
    Response.Write("  }" & vbCrLf)
    Response.Write("  ssOutputGrid.RemoveAll();" & vbCrLf)
Next

objCrossTab.BuildOutputStrings(CLng(Session("CT_PageNumber")))
For intCount = 1 To objCrossTab.OutputArrayDataUBound
    Response.Write("  ssOutputGrid.Additem(""" & CleanStringForJavaScript(Left(objCrossTab.OutputArrayData(intCount), 255)) & """);" & vbCrLf)
Next
Response.Write("  ssOutputGrid.Redraw = true;" & vbCrLf & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.ResetDefaultPrinter();" & vbCrLf)
Response.Write("  window.parent.parent.ASRIntranetOutput.Complete();" & vbCrLf)
        Response.Write("  ShowDataFrame();" & vbCrLf)
Response.Write("  }" & vbCrLf)

'PIVOT TABLE
Response.Write("  else if (frmExportData.txtFormat.value == ""6"") {" & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.PivotSuppressBlanks = (frmWorkFrame.chkSuppressZeros.checked == true);" & vbCrLf)
Response.Write("  window.parent.parent.ASRIntranetOutput.PivotDataFunction = frmWorkFrame.cboIntersectionType.options[frmWorkFrame.cboIntersectionType.selectedIndex].text;" & vbCrLf)

Response.Write("  window.parent.parent.ASRIntranetOutput.AddColumn("" "", 12, 0,false);" & vbCrLf)
For intCount = 0 To objCrossTab.ColumnHeadingUbound(0)
    Response.Write("  window.parent.parent.ASRIntranetOutput.AddColumn(""" & _
          CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(0, intCount), 255)) & """, 2, " & _
          objCrossTab.IntersectionDecimals & "," & LCase(objCrossTab.Use1000Separator) & ");" & vbCrLf)
Next
Response.Write("  window.parent.parent.ASRIntranetOutput.AddColumn(frmWorkFrame.cboIntersectionType.options[frmWorkFrame.cboIntersectionType.selectedIndex].text, 2, " & objCrossTab.IntersectionDecimals & "," & LCase(objCrossTab.Use1000Separator) & ");" & vbCrLf)

objCrossTab.GetPivotRecordset()
For intCount = 1 To objCrossTab.OutputPivotArrayDataUBound
    Response.Write(CleanStringForJavaScript_NotDoubleQuotes(objCrossTab.OutputPivotArrayData(intCount)))
Next

Response.Write("  window.parent.parent.ASRIntranetOutput.Complete();" & vbCrLf)
        Response.Write("  ShowDataFrame();" & vbCrLf)
Response.Write("  }" & vbCrLf)


'OTHER
Response.Write("  else {" & vbCrLf)

'MH20040219
Response.Write("  if (frmWorkFrame.chkPercentType.checked == true) {" & vbCrLf)
Response.Write("    lngExcelDataType = 0;" & vbCrLf)     'sqlNumeric
Response.Write("  }" & vbCrLf)
Response.Write("  else {" & vbCrLf)
Response.Write("    lngExcelDataType = 2;" & vbCrLf)     'sqlUnknown
Response.Write("  }" & vbCrLf)



Response.Write("    window.parent.parent.ASRIntranetOutput.AddColumn("" "", 12, 0,false);" & vbCrLf)
For intCount = 0 To objCrossTab.ColumnHeadingUbound(0)
    Response.Write("  window.parent.parent.ASRIntranetOutput.AddColumn(""" & CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(0, intCount), 255)) & """, lngExcelDataType, " & objCrossTab.IntersectionDecimals & "," & LCase(objCrossTab.Use1000Separator) & ");" & vbCrLf)
Next
Response.Write("  window.parent.parent.ASRIntranetOutput.AddColumn(frmWorkFrame.cboIntersectionType.options[frmWorkFrame.cboIntersectionType.selectedIndex].text, lngExcelDataType, " & objCrossTab.IntersectionDecimals & "," & LCase(objCrossTab.Use1000Separator) & ");" & vbCrLf)


If objCrossTab.PageBreakColumn = True Then
    lngLoopMin = 0
    lngLoopMax = objCrossTab.ColumnHeadingUbound(2)
Else
    lngLoopMin = 0
    lngLoopMax = 0
End If


For lngCount = lngLoopMin To lngLoopMax
    If objCrossTab.PageBreakColumn = True Then
        Response.Write("  window.parent.parent.ASRIntranetOutput.AddPage(ssOutputGrid.Caption, """ & _
             CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(2, lngCount), 255)) & """);" & vbCrLf)
    Else
        If objCrossTab.CrossTabType = 3 Then
            Response.Write("  window.parent.parent.ASRIntranetOutput.AddPage(ssOutputGrid.Caption, """ & "Absence Breakdown" & """);" & vbCrLf)
        Else
            Response.Write("  window.parent.parent.ASRIntranetOutput.AddPage(ssOutputGrid.Caption, """ & CleanStringForJavaScript(objCrossTab.BaseTableName) & """);" & vbCrLf)
        End If
    End If
    objCrossTab.BuildOutputStrings(lngCount)

    Response.Write("  window.parent.parent.ASRIntranetOutput.ArrayDim(" & CStr(objCrossTab.DataArrayCols) & ", " & CStr(objCrossTab.DataArrayRows) & ");" & vbCrLf)
    For intCol = 0 To objCrossTab.DataArrayCols
        For intRow = 0 To objCrossTab.DataArrayRows
            Response.Write("  window.parent.parent.ASRIntranetOutput.ArrayAddTo(" & CStr(intCol) & ", " & CStr(intRow) & ", """ & CleanStringForJavaScript(Left(objCrossTab.DataArray(CLng(intCol), CLng(intRow)), 255)) & """);" & vbCrLf)
        Next
    Next

    Response.Write("  window.parent.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
Next

Response.Write("  window.parent.parent.ASRIntranetOutput.Complete();" & vbCrLf)
        Response.Write("  ShowDataFrame();" & vbCrLf)
Response.Write("  }" & vbCrLf)
Response.Write("}" & vbCrLf)
Response.Write("}" & vbCrLf)

If Session("CT_Mode") = "OUTPUTRUNTHENCLOSE" Then
    Response.Write("  try {" & vbCrLf)
    Response.Write("    if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
    Response.Write("      window.parent.parent.raiseError('',false,true);" & vbCrLf)
    Response.Write("    }" & vbCrLf)
    Response.Write("    else if (window.parent.parent.ASRIntranetOutput.ErrorMessage != """") {" & vbCrLf)
    Response.Write("      window.parent.parent.raiseError(window.parent.parent.ASRIntranetOutput.ErrorMessage,false,false);" & vbCrLf)
    Response.Write("    }" & vbCrLf)
    Response.Write("    else {" & vbCrLf)
    Response.Write("      window.parent.parent.raiseError('',true,false);" & vbCrLf)
    Response.Write("    }" & vbCrLf)
    Response.Write("  }" & vbCrLf)
    Response.Write("  catch (e) {" & vbCrLf)
    Response.Write("  }" & vbCrLf)
Else
    Response.Write("  sUtilTypeDesc = window.parent.parent.frames(""top"").frmPopup.txtUtilTypeDesc.value;" & vbCrLf)
    Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
            Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output failed.\n\nCancelled by user."",64,sUtilTypeDesc);" & vbCrLf)
            Response.Write("  }" & vbCrLf)
            Response.Write("  else if (window.parent.parent.ASRIntranetOutput.ErrorMessage != """") {" & vbCrLf)
            Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output failed.\n\n""+window.parent.parent.ASRIntranetOutput.ErrorMessage,48,sUtilTypeDesc);" & vbCrLf)
            Response.Write("	  var fs = window.parent.parent.document.all.item(""myframeset"");" & vbCrLf)
            Response.Write("	  if (fs) " & vbCrLf)
            Response.Write("	    {" & vbCrLf)
            Response.Write("	  	  fs.rows = ""0,0,*"";" & vbCrLf)
            Response.Write("	    }" & vbCrLf)
            Response.Write("  }" & vbCrLf)
            Response.Write("  else {" & vbCrLf)
            Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output complete."",64,sUtilTypeDesc);" & vbCrLf)
    Response.Write("  }" & vbCrLf)
End If


ElseIf Session("CT_Mode") = "EMAILGROUP" Or _
       Session("CT_Mode") = "EMAILGROUPTHENCLOSE" Then
strEmailAddresses = ""
If Session("CT_EmailGroupID") > 0 Then
    cmdReportsCols = Server.CreateObject("ADODB.Command")
    cmdReportsCols.CommandText = "spASRIntGetEmailGroupAddresses"
    cmdReportsCols.CommandType = 4 ' Stored procedure
    cmdReportsCols.ActiveConnection = Session("databaseConnection")

    prmEmailGroupID = cmdReportsCols.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
    cmdReportsCols.Parameters.Append(prmEmailGroupID)
    prmEmailGroupID.value = CleanNumeric(Session("CT_EmailGroupID"))

    Err.Clear()
    rstReportColumns = cmdReportsCols.Execute

    If (Err.Number <> 0) Then
        sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(Err.Description)
    End If

    If Len(sErrorDescription) = 0 Then
        iLoop = 1
        Do While Not rstReportColumns.EOF
            If iLoop > 1 Then
                strEmailAddresses = strEmailAddresses & ";"
            End If
            strEmailAddresses = strEmailAddresses & rstReportColumns.Fields("Fixed").Value
            rstReportColumns.MoveNext()
            iLoop = iLoop + 1
        Loop

        ' Release the ADO recordset object.
        rstReportColumns.close()
    End If
					
    rstReportColumns = Nothing
    cmdReportsCols = Nothing

End If
'Session("CT_EmailGroupAddr") = strEmailAddresses
Response.Write("  frmExportData.txtEmailGroupAddr.value = """ & CleanStringForJavaScript(strEmailAddresses) & """;" & vbCrLf)

If Session("CT_Mode") = "EMAILGROUPTHENCLOSE" Then
    Response.Write("  ExportData(""OUTPUTRUNTHENCLOSE"");" & vbCrLf)
Else
    Response.Write("  ExportData(""OUTPUTRUN"");" & vbCrLf)
End If


ElseIf Session("CT_Mode") = "" Then
'Must be the first time this asp is called...
        Response.Write(" loadAddRecords();" & vbCrLf)

End If

    Response.Write("  try {" & vbCrLf)
    Response.Write("    refreshCombo(""INTERSECTIONTYPE"");" & vbCrLf)
    Response.Write("    refreshCombo(""PAGE"");" & vbCrLf)
    Response.Write("    refreshCombo(""FILEFORMAT"");" & vbCrLf)
    Response.Write("  }" & vbCrLf)
    Response.Write("  catch(e) {" & vbCrLf)
    Response.Write("  }" & vbCrLf)

    Response.Write("}" & vbCrLf)
    Response.Write("</script>" & vbCrLf & vbCrLf)
%>

<script type="text/javascript">	
	
    function getData(strMode, lngPageNumber, lngIntType, blnShowPer, blnPerPage, blnSupZeros, blnThousand) {

        var frmWorkFrame = OpenHR.getFrame("reportworkframe");
        control_disable(frmWorkFrame.cboIntersectionType, true);
        control_disable(frmWorkFrame.chkPercentPage, true);
        control_disable(frmWorkFrame.chkPercentType, true);
        control_disable(frmWorkFrame.chkSuppressZeros, true);
        control_disable(frmWorkFrame.chkUse1000, true);
        control_disable(frmWorkFrame.cboPage, true);

        var frmGetData = OpenHR.getForm("reportdataframe", "frmGetCrossTabData");
        frmGetData.txtMode.value = strMode;
        frmGetData.txtPageNumber.value = lngPageNumber;
        frmGetData.txtIntersectionType.value = lngIntType;
        frmGetData.txtShowPercentage.value = blnShowPer;
        frmGetData.txtPercentageOfPage.value = blnPerPage;
        frmGetData.txtSuppressZeros.value = blnSupZeros;
        frmGetData.txtUse1000.value = blnThousand;
        OpenHR.submitForm(frmGetData, null, false);
    }

    function getBreakdown(lngHor, lngVer, lngPgb, txtIntType, txtCellValue) {

        var frmGetData = OpenHR.getForm("reportbreakdownframe", "frmGetCrossTabData");
        frmGetData.txtMode.value = "BREAKDOWN";
        frmGetData.txtHor.value = lngHor;
        frmGetData.txtVer.value = lngVer;
        frmGetData.txtPgb.value = lngPgb;
        frmGetData.txtIntersectionType.value = txtIntType;
        frmGetData.txtCellValue.value = txtCellValue;
        OpenHR.submitForm(frmGetData, null, false);
    }

    function ExportData(strMode) {
        
        var frmGetData = OpenHR.getForm("reportdataframe", "frmGetCrossTabData");
        frmGetData.txtMode.value = strMode;
        OpenHR.submitForm(frmGetData, null, false);
    }

</script>

<form action="util_run_crosstabsDataSubmit" method="post" id="frmGetCrossTabData" name="frmGetCrossTabData">
    <input type="hidden" id="txtMode" name="txtMode" value="<%=Session("CT_Mode")%>">
    <input type="hidden" id="txtPageNumber" name="txtPageNumber" value="<%=Session("CT_PageNumber")%>">
    <input type="hidden" id="txtShowPercentage" name="txtShowPercentage" value="<%=Session("CT_ShowPercentage")%>">
    <input type="hidden" id="txtPercentageOfPage" name="txtPercentageOfPage" value="<%=Session("CT_PercentageOfPage")%>">
    <input type="hidden" id="txtSuppressZeros" name="txtSuppressZeros" value="<%=Session("CT_SuppressZeros")%>">
    <input type="hidden" id="txtUse1000" name="txtUse1000" value="<%=Session("CT_Use1000")%>">
    <input type="hidden" id="txtHor" name="txtHor" value="<%=Session("CT_Hor")%>">
    <input type="hidden" id="txtVer" name="txtVer" value="<%=Session("CT_Ver")%>">
    <input type="hidden" id="txtPgb" name="txtPgb" value="<%=Session("CT_Pgb")%>">
    <input type="hidden" id="txtIntersectionType" name="txtIntersectionType" value="<%=Session("CT_IntersectionType")%>">
	  <input type="hidden" id="txtDataIntersectionType" name="txtDataIntersectionType" value="<%=Session("CT_IntersectionType")%>">
    <input type="hidden" id="txtCellValue" name="txtCellValue" value="<%=Session("CT_CellValue")%>">
    <input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("CT_UtilID")%>">
    <input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=Session("CT_EmailGroupID")%>">
</form>

<textarea id="holdtext" style="display: none;"></textarea>


<script type="text/javascript">
    // Generated by the response.writes above
    util_run_crosstabs_data_window_onload();
</script>
