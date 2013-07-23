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
	Dim objCrossTab As Object	' HR.Intranet.Server.CrossTab

		If Session("CT_Mode") = "BREAKDOWN" Then
		 
				objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))

				If objCrossTab.CrossTabType = 3 Then
						Response.Write("Absence Breakdown Cell Breakdown")
				Else
						Response.Write("Cross Tabs Cell Breakdown")
				End If
		
%>

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
																						objCrossTab.ColumnHeading(2, Session("CT_Pgb")) & _
																						""" style=""WIDTH: 100%"" disabled=""disabled""></TD>" & vbCrLf)
														Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
														Response.Write("</TR>" & vbCrLf)
														Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)
												End If

												Response.Write("<TR HEIGHT=5>" & vbCrLf)
												Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
												Response.Write("  <TD WIDTH=""50%"">" & objCrossTab.HorizontalColumnName & " :</TD>" & vbCrLf)
												Response.Write("  <TD WIDTH=""50%""><INPUT id=txtHor name=txtHor value="" ")

												If CLng(Session("CT_Hor")) > CLng(objCrossTab.ColumnHeadingUbound(0)) Then
														Response.Write("<All>")
												Else
														Response.Write(objCrossTab.ColumnHeading(0, Session("CT_Hor")))
												End If
												Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
												Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
												Response.Write("</TR>" & vbCrLf)
												Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)

												Response.Write("<TR HEIGHT=5>" & vbCrLf)
												Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
												Response.Write("  <TD>" & objCrossTab.VerticalColumnName & " :</TD>" & vbCrLf)
												Response.Write("  <TD><INPUT id=txtVer name=txtVer value="" ")
												If CLng(Session("CT_Ver")) > CLng(objCrossTab.ColumnHeadingUbound(1)) Then
														Response.Write("<All>")
												Else
														Response.Write(objCrossTab.ColumnHeading(1, Session("CT_Ver")))
												End If
												Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
												Response.Write("</TR>" & vbCrLf)
												Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)


												Response.Write("<TR HEIGHT=5>" & vbCrLf)
												Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
											Response.Write("  <TD>" & objCrossTab.IntersectionTypeValue(Session("CT_IntersectionType")) & " :</TD>" & vbCrLf)
												Response.Write("  <TD><INPUT id=txtCellValue name=txtCellValue value="" ")
				
												If objCrossTab.CrossTabType = 3 Then
														Response.Write(objCrossTab.OutputArrayDataUBound)
												Else
														Response.Write(Session("CT_CellValue"))
												End If

												Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
												Response.Write("</TR>" & vbCrLf)
												Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)

										%>
										<tr>
												<td></td>
												<td colspan="2">
													<table id="ssOutputBreakdown"></table>
													
												</td>
										</tr>
										<tr height="2">
												<td>&nbsp;</td>
										</tr>
										<tr height="5">
												<td colspan="3" align="RIGHT">
														<input type="button" id="cmdClose" name="cmdClose" value="OK" style="WIDTH: 80px" class="btn" onclick="ShowDataFrame();" />
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

			Response.Write("    var colNames = [];" & vbCrLf)
			Response.Write("    var colData = [];" & vbCrLf)
			Response.Write("    var colMode = [];" & vbCrLf)
			Response.Write("    var value;" & vbCrLf)
			Response.Write("    var i;" & vbCrLf)
			Response.Write("    var sColumnName;" & vbCrLf)
			Response.Write("    var iCount2;" & vbCrLf)
			Response.Write("    var obj;" & vbCrLf)
		
		Response.Write("  colNames.push('" & CleanStringForJavaScript(objCrossTab.BaseTableName) & "');" & vbCrLf)
		Response.Write("	colMode.push({ name: '" & CleanStringForJavaScript(objCrossTab.BaseTableName) & "' });" & vbCrLf)	
		
		' Absence Breakdown
		If objCrossTab.CrossTabType = 3 Then

			Response.Write("  colNames.push('Start Date');" & vbCrLf)
			Response.Write("	colMode.push({ name: 'Start Date' });" & vbCrLf)

			Response.Write("  colNames.push('End Date');" & vbCrLf)
			Response.Write("	colMode.push({ name: 'End Date' });" & vbCrLf)

			Response.Write("  colNames.push(""" & CleanStringForJavaScript(objCrossTab.ColumnHeading(0, Session("CT_Hor"))) & "'s taken" & """);" & vbCrLf)
			Response.Write("	colMode.push({ name: """ & CleanStringForJavaScript(objCrossTab.ColumnHeading(0, Session("CT_Hor"))) & "'s taken" & """ });" & vbCrLf)
				
			iInterSectionColumnCount = 4
		End If
		
		If objCrossTab.IntersectionColumn = True Then
			Response.Write("  colNames.push('" & CleanStringForJavaScript(objCrossTab.IntersectionColumnName) & "');" & vbCrLf)
			Response.Write("	colMode.push({ name: '" & CleanStringForJavaScript(objCrossTab.IntersectionColumnName) & "' });" & vbCrLf)
		End If

		
		Dim objData As String()
			For intCount = 1 To CLng(objCrossTab.OutputArrayDataUBound)
								
				Response.Write("  obj = {};" & vbCrLf)
				objData = Split(objCrossTab.OutputArrayData(intCount), vbTab)
				For intCount2 = 0 To UBound(objData)
					Response.Write("  obj[colNames[" & intCount2 & "]] = '" & objData(intCount2) & "';" & vbCrLf)
				Next
				Response.Write("  colData.push(obj);")
			Next
		
		
		'For intCount = 1 To CLng(objCrossTab.OutputArrayDataUBound)
		'	Response.Write("  ssOutputBreakdown.AddItem(""" & CleanStringForJavaScript(objCrossTab.OutputArrayData(CLng(intCount))) & """);" & vbCrLf)
		'Next
		
		'Response.Write("  ssOutputBreakdown.RowHeight = 10;" & vbCrLf)

		'Response.Write("  ssOutputBreakdown.VisibleCols = 2;" & vbCrLf)
		'Response.Write("  ssOutputBreakdown.VisibleRows = 10;" & vbCrLf)
			
		'Response.Write("  ssOutputBreakdown.Redraw = true;" & vbCrLf)

		Response.Write("	$('#ssOutputBreakdown').jqGrid({data: colData, datatype: 'local', colNames: colNames, colModel: colMode, autowidth: true" & vbCrLf)
		Response.Write("	, cmTemplate: { editable: true }});")
		
			
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
