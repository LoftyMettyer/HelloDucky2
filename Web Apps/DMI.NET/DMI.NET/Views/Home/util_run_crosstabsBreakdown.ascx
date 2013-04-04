<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<form action="util_run_CrossTabsBreakdown" method="post" id="frmBreakdown" name="frmBreakdown">
    <input type="hidden" id="txtMode" name="txtMode" value="<%=Session("CT_Mode")%>">
    <input type="hidden" id="txtHor" name="txtHor" value="<%=Session("CT_Hor")%>">
    <input type="hidden" id="txtVer" name="txtVer" value="<%=Session("CT_Ver")%>">
    <input type="hidden" id="txtPgb" name="txtPgb" value="<%=Session("CT_Pgb")%>">
    <input type="hidden" id="txtIntersectionType" name="txtIntersectionType" value=0>
    <input type="hidden" id="txtCellValue" name="txtCellValue" value="<%=Session("CT_CellValue")%>">
</form>

<%
    Dim objCrossTab As Object

    If Session("CT_Mode") = "BREAKDOWN" Then
     
        objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))

        If objCrossTab.CrossTabType = 3 Then
            Response.Write("Absence Breakdown Cell Breakdown")
        Else
            Response.Write("Cross Tabs Cell Breakdown")
        End If
    
%>

    <object
        classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
        id="Microsoft_Licensed_Class_Manager_1_0">
        <param name="LPKPath" value="lpks/main.lpk">
    </object>


    <table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="400px">
        <tr height="1">
            <td>
                <table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">

                    <%
                        objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
			
                        If objCrossTab.PageBreakColumnName <> "<None>" Then
                            Response.Write("<TR HEIGHT=5>" & vbCrLf)
                            Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
                            Response.Write("  <TD WIDTH=""50%"">" & objCrossTab.PageBreakColumnName & " :</TD>" & vbCrLf)
                            Response.Write("  <TD WIDTH=""50%""><INPUT id=txtPgb class=""text textdisabled"" name=txtPgb value="" " & _
                                            objCrossTab.ColumnHeading(2, Session("txtPgb")) & _
                                            """ style=""WIDTH: 100%"" disabled=""disabled""></TD>" & vbCrLf)
                            Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
                            Response.Write("</TR>" & vbCrLf)
                            Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)
                        End If

                        Response.Write("<TR HEIGHT=5>" & vbCrLf)
                        Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
                        Response.Write("  <TD WIDTH=""50%"">" & objCrossTab.HorizontalColumnName & " :</TD>" & vbCrLf)
                        Response.Write("  <TD WIDTH=""50%""><INPUT id=txtHor name=txtHor value="" ")

                        If CLng(Session("txtHor")) > CLng(objCrossTab.ColumnHeadingUbound(0)) Then
                            Response.Write("<All>")
                        Else
                            Response.Write(objCrossTab.ColumnHeading(0, Session("txtHor")))
                        End If
                        Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
                        Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
                        Response.Write("</TR>" & vbCrLf)
                        Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)

                        Response.Write("<TR HEIGHT=5>" & vbCrLf)
                        Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
                        Response.Write("  <TD>" & objCrossTab.VerticalColumnName & " :</TD>" & vbCrLf)
                        Response.Write("  <TD><INPUT id=txtVer name=txtVer value="" ")
                        If CLng(Session("txtVer")) > CLng(objCrossTab.ColumnHeadingUbound(1)) Then
                            Response.Write("<All>")
                        Else
                            Response.Write(objCrossTab.ColumnHeading(1, Session("txtVer")))
                        End If
                        Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
                        Response.Write("</TR>" & vbCrLf)
                        Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)


                        Response.Write("<TR HEIGHT=5>" & vbCrLf)
                        Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
                    	Response.Write("  <TD>" & Session("txtDataIntersectionType") & " :</TD>" & vbCrLf)
                        Response.Write("  <TD><INPUT id=txtCellValue name=txtCellValue value="" ")
				
                        If objCrossTab.CrossTabType = 3 Then
                            Response.Write(objCrossTab.OutputArrayDataUBound)
                        Else
                            Response.Write(Session("txtCellValue"))
                        End If

                        Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
                        Response.Write("</TR>" & vbCrLf)
                        Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)

                    %>
                    <tr>
                        <td></td>
                        <td colspan="2">
                            <object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
                                codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
                                id="ssOutputBreakdown" name="ssOutputBreakdown"
                                style="HEIGHT: 300px; LEFT: 0px; TOP: 0px; WIDTH: 400px">
                                <param name="ScrollBars" value="4">
                                <param name="_Version" value="196617">
                                <param name="DataMode" value="2">
                                <param name="Cols" value="0">
                                <param name="Rows" value="0">
                                <param name="BorderStyle" value="1">
                                <param name="RecordSelectors" value="0">
                                <param name="GroupHeaders" value="1">
                                <param name="ColumnHeaders" value="1">
                                <param name="GroupHeadLines" value="1">
                                <param name="HeadLines" value="1">
                                <param name="FieldDelimiter" value="(None)">
                                <param name="FieldSeparator" value="(Tab)">
                                <param name="Row.Count" value="0">
                                <param name="Col.Count" value="1">
                                <param name="stylesets.count" value="1">
                                <param name="TagVariant" value="EMPTY">
                                <param name="UseGroups" value="0">
                                <param name="HeadFont3D" value="0">
                                <param name="Font3D" value="0">
                                <param name="DividerType" value="3">
                                <param name="DividerStyle" value="1">
                                <param name="DefColWidth" value="3528">
                                <param name="BeveColorScheme" value="2">
                                <param name="BevelColorFrame" value="-2147483642">
                                <param name="BevelColorHighlight" value="-2147483643">
                                <param name="BevelColorShadow" value="-2147483632">
                                <param name="BevelColorFace" value="-2147483633">
                                <param name="CheckBox3D" value="1">
                                <param name="AllowAddNew" value="0">
                                <param name="AllowDelete" value="0">
                                <param name="AllowUpdate" value="1">
                                <param name="MultiLine" value="0">
                                <param name="ActiveCellStyleSet" value="Highlight">
                                <param name="RowSelectionStyle" value="0">
                                <param name="AllowRowSizing" value="1">
                                <param name="AllowGroupSizing" value="1">
                                <param name="AllowColumnSizing" value="1">
                                <param name="AllowGroupMoving" value="0">
                                <param name="AllowColumnMoving" value="0">
                                <param name="AllowGroupSwapping" value="0">
                                <param name="AllowColumnSwapping" value="0">
                                <param name="AllowGroupShrinking" value="0">
                                <param name="AllowColumnShrinking" value="0">
                                <param name="AllowDragDrop" value="0">
                                <param name="UseExactRowCount" value="1">
                                <param name="SelectTypeCol" value="0">
                                <param name="SelectTypeRow" value="0">
                                <param name="SelectByCell" value="1">
                                <param name="BalloonHelp" value="0">
                                <param name="RowNavigation" value="0">
                                <param name="CellNavigation" value="0">
                                <param name="MaxSelectedRows" value="1">
                                <param name="HeadStyleSet" value="">
                                <param name="StyleSet" value="">
                                <param name="ForeColorEven" value="0">
                                <param name="ForeColorOdd" value="0">
                                <param name="BackColorEven" value="-2147483643">
                                <param name="BackColorOdd" value="-2147483643">
                                <param name="Levels" value="1">
                                <param name="RowHeight" value="239">
                                <param name="ExtraHeight" value="239">
                                <param name="ActiveRowStyleSet" value="">
                                <param name="CaptionAlignment" value="2">
                                <param name="SplitterPos" value="0">
                                <param name="SplitterVisible" value="0">
                                <param name="Columns.Count" value="1">
                                <param name="Columns(0).Width" value="3528">
                                <param name="Columns(0).Visible" value="-1">
                                <param name="Columns(0).Columns.Count" value="1">
                                <param name="Columns(0).Caption" value="  ">
                                <param name="Columns(0).Name" value="">
                                <param name="Columns(0).Alignment" value="0">
                                <param name="Columns(0).CaptionAlignment" value="3">
                                <param name="Columns(0).Bound" value="0">
                                <param name="Columns(0).AllowSizing" value="1">
                                <param name="Columns(0).DataField" value="">
                                <param name="Columns(0).DataType" value="8">
                                <param name="Columns(0).Level" value="0">
                                <param name="Columns(0).NumberFormat" value="">
                                <param name="Columns(0).Case" value="0">
                                <param name="Columns(0).FieldLen" value="4096">
                                <param name="Columns(0).VertScrollBar" value="0">
                                <param name="Columns(0).Locked" value="0">
                                <param name="Columns(0).Style" value="0">
                                <param name="Columns(0).ButtonsAlways" value="0">
                                <param name="Columns(0).RowCount" value="0">
                                <param name="Columns(0).ColCount" value="1">
                                <param name="Columns(0).HasHeadForeColor" value="0">
                                <param name="Columns(0).HasHeadBackColor" value="0">
                                <param name="Columns(0).HasForeColor" value="0">
                                <param name="Columns(0).HasBackColor" value="0">
                                <param name="Columns(0).HeadForeColor" value="0">
                                <param name="Columns(0).HeadBackColor" value="0">
                                <param name="Columns(0).ForeColor" value="0">
                                <param name="Columns(0).BackColor" value="0">
                                <param name="Columns(0).HeadStyleSet" value="">
                                <param name="Columns(0).StyleSet" value="">
                                <param name="Columns(0).Nullable" value="1">
                                <param name="Columns(0).Mask" value="">
                                <param name="Columns(0).PromptInclude" value="0">
                                <param name="Columns(0).ClipMode" value="0">
                                <param name="Columns(0).PromptChar" value="95">
                                <param name="UseDefaults" value="-1">
                                <param name="TabNavigation" value="1">
                                <param name="BatchUpdate" value="0">
                                <param name="_ExtentX" value="2646">
                                <param name="_ExtentY" value="1323">
                                <param name="_StockProps" value="79">
                                <param name="Caption" value="">
                                <param name="ForeColor" value="0">
                                <param name="BackColor" value="16777215">
                                <param name="Enabled" value="-1">
                                <param name="DataMember" value="">
                            </object>
                        </td>
                    </tr>
                    <tr height="2">
                        <td>&nbsp;</td>
                    </tr>
                    <tr height="5">
                        <td colspan="3" align="RIGHT">
                            <input type="button" id="cmdClose" name="cmdClose" value="OK" style="WIDTH: 80px" class="btn"
                                onclick="ShowDataFrame();"
                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
                    </tr>
                </table>

            </td>
        </tr>
    </table>

    <%

        Dim iInterSectionColumnCount As Integer

        objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
        iInterSectionColumnCount = 1
 
        Response.Write("<script type=""text/javascript"">" & vbCrLf)
        Response.Write("function util_run_crosstabsBreakdown_window_onload()" & vbCrLf)
    	Response.Write("{" & vbCrLf & vbCrLf)

    %>
	
	setGridFont(ssOutputBreakdown);
	
	<%	    
		Response.Write("  ssOutputBreakdown.Columns.RemoveAll();" & vbCrLf)
		Response.Write("  ssOutputBreakdown.Columns.Add(0);" & vbCrLf)
		Response.Write("  ssOutputBreakdown.Columns(0).Caption = """ & CleanStringForJavaScript(objCrossTab.BaseTableName) & """;" & vbCrLf)
		Response.Write("  ssOutputBreakdown.Columns(0).Locked = true;" & vbCrLf)
		Response.Write("  ssOutputBreakdown.Columns(0).Visible = true;" & vbCrLf)
		Response.Write("  ssOutputBreakdown.Columns(0).Width = 300;" & vbCrLf)

	    ' Absence Breakdown
	    If objCrossTab.CrossTabType = 3 Then
	        Response.Write("  ssOutputBreakdown.Columns.Add(1);" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(1).Caption = ""Start Date"";" & vbCrLf)
			Response.Write("  ssOutputBreakdown.Columns(1).Locked = true;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(1).Visible = true;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(1).Width = 150;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(1).Alignment = 1;" & vbCrLf)
		
	        Response.Write("  ssOutputBreakdown.Columns.Add(2);" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(2).Caption = ""End Date"";" & vbCrLf)
			Response.Write("  ssOutputBreakdown.Columns(2).Locked = true;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(2).Visible = true;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(2).Width = 150;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(2).Alignment = 1;" & vbCrLf)
				
	        Response.Write("  ssOutputBreakdown.Columns.Add(3);" & vbCrLf)
		
	        Response.Write("  ssOutputBreakdown.Columns(3).Caption = """ & CleanStringForJavaScript(objCrossTab.ColumnHeading(0, Session("txtHor"))) & "'s taken" & """;" & vbCrLf)
			Response.Write("  ssOutputBreakdown.Columns(3).Locked = true;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(3).Visible = true;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(3).Width = 150;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(3).Alignment = 1;" & vbCrLf)
				
	        iInterSectionColumnCount = 4
	    End If


		Response.Write("ssOutputBreakdown.Redraw = false;" & vbCrLf & vbCrLf)
	    
	    If objCrossTab.IntersectionColumn = True Then
	        Response.Write("  ssOutputBreakdown.Columns.Add(" & iInterSectionColumnCount & ");" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(" & iInterSectionColumnCount & ").Caption = """ & CleanStringForJavaScript(objCrossTab.IntersectionColumnName) & """;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(" & iInterSectionColumnCount & ").Locked = true;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(" & iInterSectionColumnCount & ").Visible = true;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(" & iInterSectionColumnCount & ").Width = 150;" & vbCrLf)
	        Response.Write("  ssOutputBreakdown.Columns(" & iInterSectionColumnCount & ").Alignment = 1;" & vbCrLf)
	    End If

		For intCount = 1 To CLng(objCrossTab.OutputArrayDataUBound)
			Response.Write("  ssOutputBreakdown.AddItem(""" & CleanStringForJavaScript(objCrossTab.OutputArrayData(CLng(intCount))) & """);" & vbCrLf)
		Next
	    Response.Write("  ssOutputBreakdown.RowHeight = 10;" & vbCrLf)

	    Response.Write("  ssOutputBreakdown.VisibleCols = 2;" & vbCrLf)
	    Response.Write("  ssOutputBreakdown.VisibleRows = 10;" & vbCrLf)
	    
	    Response.Write("  ssOutputBreakdown.Redraw = true;" & vbCrLf)

	    
	    
		Response.Write("}" & vbCrLf)
	    Response.Write("</script>" & vbCrLf & vbCrLf)

	    objCrossTab = Nothing

    %>


<script type="text/javascript">

    $("#reportbreakdownframe").attr("data-framesource", "UTIL_RUN_CROSSTABSBREAKDOWN");

    $("#reportframe").show();
    $("#reportdataframe").hide();
    $("#reportworkframe").hide();
    $("#reportbreakdownframe").show();

    setTimeout("util_run_crosstabsBreakdown_window_onload()", 100);
    
</script>

<%
    End If
    %>
