<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<script src="<%: Url.Content("~/bundles/recordedit")%>" type="text/javascript"></script>

<script type="text/javascript">
	$(document).ready(function () {
		if ('<%=session("linktype")%>' == 'multifind') {
			//for multifind (SSI views) show relevant buttons with applicable functions
			menu_setVisibletoolbarGroupById("mnuSectionRecordFindEdit", false);
			menu_setVisibleMenuItem("mnutoolAccessLinksFind", true);
			menu_setVisibleMenuItem("mnutoolCancelLinksFind", false);

			//redo the doubleclick function			
			setTimeout('$("#findGridTable").jqGrid("setGridParam", { ondblClickRow: function (rowID) {doEdit();}});', 200);

		} else {
			menu_setVisibletoolbarGroupById("mnuSectionRecordFindEdit", true);
			menu_setVisibleMenuItem("mnutoolAccessLinksFind", false);
			menu_setVisibleMenuItem("mnutoolCancelLinksFind", false);
		}
	});

	function doEdit() {
		var sRecordID = selectedRecordID();

		if ("<%=session("linkType")%>" == "multifind") {
			var sParams = "<%=session("TableID")%>!<%=session("ViewID")%>_";
			sParams = sParams.concat(sRecordID);
			loadPartialView('linksMain', 'Home', 'workframe', sParams);
		}
	}
</script>

<div id="divFindForm" <%=session("BodyTag")%>>
	<form action="" class="absolutefull" method="POST" id="frmFindForm" name="frmFindForm">
		<div class="absolutefull">
			<div id="row1" style="margin-left: 20px; margin-right: 20px">
				<%
					On Error Resume Next
	
					Dim sErrorDescription As String = ""
					If ViewBag.pageTitle.ToString().Length = 0 Then
						' DMI View.
						' Display the appropriate page title.
						Dim cmdFindWindowTitle = New ADODB.Command
						cmdFindWindowTitle.CommandText = "sp_ASRIntGetFindWindowInfo"
						cmdFindWindowTitle.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
						cmdFindWindowTitle.ActiveConnection = Session("databaseConnection")

						Dim prmTitle = cmdFindWindowTitle.CreateParameter("title", 200, 2, 100)
						cmdFindWindowTitle.Parameters.Append(prmTitle)

						Dim prmQuickEntry = cmdFindWindowTitle.CreateParameter("quickEntry", 11, 2)	' 11=bit, 2=output
						cmdFindWindowTitle.Parameters.Append(prmQuickEntry)

						Dim prmScreenID = cmdFindWindowTitle.CreateParameter("screenID", 3, 1)
						cmdFindWindowTitle.Parameters.Append(prmScreenID)
						prmScreenID.Value = CleanNumeric(Session("screenID"))

						Dim prmViewID = cmdFindWindowTitle.CreateParameter("viewID", 3, 1)
						cmdFindWindowTitle.Parameters.Append(prmViewID)
						prmViewID.Value = CleanNumeric(Session("viewID"))

						Err.Clear()
						cmdFindWindowTitle.Execute()
						If (Err.Number <> 0) Then
							sErrorDescription = "The page title could not be created." & vbCrLf & FormatError(Err.Description)
						End If

						If Len(sErrorDescription) = 0 Then
							Dim homelinkURL = "javascript:loadPartialView(""linksMain"", ""Home"", ""workframe"", null);"
							Response.Write(String.Format("<div class='pageTitleDiv'><a href='{0}' title='Home'><i class='pageTitleIcon icon-circle-arrow-left'></i></a><span class='pageTitle'>Find - " & _
									Replace(cmdFindWindowTitle.Parameters("title").Value, "_", " ") & "</span>" & vbCrLf, homelinkURL))
							Response.Write("<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & cmdFindWindowTitle.Parameters("quickEntry").Value & "></div>" & vbCrLf)
						End If
								
						' Release the ADO command object.
						cmdFindWindowTitle = Nothing
					Else
						' SSI View.
						Dim homelinkURL = "javascript:loadPartialView(""linksMain"", ""Home"", ""workframe"", null);"
						Response.Write(String.Format("<div class='pageTitleDiv'><a href='{0}' title='Home'><i class='pageTitleIcon icon-circle-arrow-left'></i></a><span class='pageTitle'>" & _
								ViewBag.pageTitle & "</span>" & vbCrLf, homelinkURL))
						Response.Write("<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & ViewBag.pageTitle & "></div>" & vbCrLf)
					End If
				%>
			</div>
			<div id="findGridRow" style="height: <%If Session("parentTableID") > 0 Then%>65%<%Else%>85%<%End If%>; margin-right: 20px; margin-left: 20px;">
				<%
					Dim sTemp As String
					Dim sThousandColumns As String
					Dim sBlankIfZeroColumns As String
					Dim sColDef As String
					Dim iCount As Integer
					Dim sAddString As String
								
					Const AD_STATE_OPEN = 1
								
					Dim fCancelDateColumn = True
					If (Len(sErrorDescription) = 0) And (Session("TB_CourseTableID") > 0) And Len(Session("lineage").ToString()) > 0 Then
						Dim sSubString As String = Session("lineage").ToString()
						Dim iIndex = InStr(sSubString, "_")
						sSubString = Mid(sSubString, iIndex + 1)
						iIndex = InStr(sSubString, "_")
						sSubString = Mid(sSubString, iIndex + 1)
						iIndex = InStr(sSubString, "_")
						sSubString = Mid(sSubString, iIndex + 1)
						iIndex = InStr(sSubString, "_")
						sSubString = Mid(sSubString, iIndex + 1)
						iIndex = InStr(sSubString, "_")
						Dim lngRecordID = Left(sSubString, iIndex - 1)

						' Get the Course Date
						Dim cmdGetCancelDateColumn = New Command
						cmdGetCancelDateColumn.CommandText = "spASRIntGetCancelCourseDate"
						cmdGetCancelDateColumn.CommandType = CommandTypeEnum.adCmdStoredProc
						cmdGetCancelDateColumn.ActiveConnection = Session("databaseConnection")
						cmdGetCancelDateColumn.CommandTimeout = 180
				
						Dim prmError = cmdGetCancelDateColumn.CreateParameter("error", 11, 2)	' 11=bit, 2=output
						cmdGetCancelDateColumn.Parameters.Append(prmError)

						Dim prmRecID = cmdGetCancelDateColumn.CreateParameter("recordID", 3, 1)	' 3=integer, 1=input
						cmdGetCancelDateColumn.Parameters.Append(prmRecID)
						prmRecID.Value = CleanNumeric(lngRecordID)

						Dim prmCancelDateColumn = cmdGetCancelDateColumn.CreateParameter("CancelDateColumn", 11, 2)	' 11=bit, 2=output
						cmdGetCancelDateColumn.Parameters.Append(prmCancelDateColumn)
			
						Err.Clear()
						cmdGetCancelDateColumn.Execute()

						If (Err.Number <> 0) Then
							sErrorDescription = "Unable to check for a Cancelled Course Date." & vbCrLf & FormatError(Err.Description)
						End If

						If Len(sErrorDescription) = 0 Then
							fCancelDateColumn = cmdGetCancelDateColumn.Parameters("CancelDateColumn").Value
						End If
			
						' Release the ADO command object.
						cmdGetCancelDateColumn = Nothing
					End If

					If Len(sErrorDescription) = 0 Then
						' Get the find records.
						Dim cmdFindRecords = New Command
						cmdFindRecords.CommandText = "sp_ASRIntGetFindRecords3"
						cmdFindRecords.CommandType = CommandTypeEnum.adCmdStoredProc
						cmdFindRecords.ActiveConnection = Session("databaseConnection")
						cmdFindRecords.CommandTimeout = 180

						Dim prmError = cmdFindRecords.CreateParameter("error", 11, 2)	' 11=bit, 2=output
						cmdFindRecords.Parameters.Append(prmError)

						Dim prmSomeSelectable = cmdFindRecords.CreateParameter("someSelectable", 11, 2)	' 11=bit, 2=output
						cmdFindRecords.Parameters.Append(prmSomeSelectable)

						Dim prmSomeNotSelectable = cmdFindRecords.CreateParameter("someNotSelectable", 11, 2)	' 11=bit, 2=output
						cmdFindRecords.Parameters.Append(prmSomeNotSelectable)

						Dim prmRealSource = cmdFindRecords.CreateParameter("realSource", 200, 2, 255)	' 200=varchar, 2=output, 255=size
						cmdFindRecords.Parameters.Append(prmRealSource)

						Dim prmInsertGranted = cmdFindRecords.CreateParameter("insertGranted", 11, 2)	' 11=bit, 2=output
						cmdFindRecords.Parameters.Append(prmInsertGranted)

						Dim prmDeleteGranted = cmdFindRecords.CreateParameter("deleteGranted", 11, 2)	' 11=bit, 2=output
						cmdFindRecords.Parameters.Append(prmDeleteGranted)

						Dim prmTableID = cmdFindRecords.CreateParameter("tableID", 3, 1)
						cmdFindRecords.Parameters.Append(prmTableID)
						prmTableID.Value = CleanNumeric(Session("tableID"))

						Dim prmViewID = cmdFindRecords.CreateParameter("viewID", 3, 1)
						cmdFindRecords.Parameters.Append(prmViewID)
						prmViewID.Value = CleanNumeric(Session("viewID"))

						Dim prmOrderID = cmdFindRecords.CreateParameter("orderID", 3, 1)
						cmdFindRecords.Parameters.Append(prmOrderID)
						prmOrderID.Value = CleanNumeric(Session("orderID"))

						Dim prmParentTableID = cmdFindRecords.CreateParameter("parentTableID", 3, 1)
						cmdFindRecords.Parameters.Append(prmParentTableID)
						prmParentTableID.Value = CleanNumeric(Session("parentTableID"))

						Dim prmParentRecordID = cmdFindRecords.CreateParameter("parentRecordID", 3, 1)
						cmdFindRecords.Parameters.Append(prmParentRecordID)
						prmParentRecordID.Value = CleanNumeric(Session("parentRecordID"))

						Dim prmFilterDef = cmdFindRecords.CreateParameter("filterDef", 200, 1, 2147483646)
						cmdFindRecords.Parameters.Append(prmFilterDef)
						prmFilterDef.Value = Session("filterDef")

						Dim prmReqRecs = cmdFindRecords.CreateParameter("reqRecs", 3, 1)
						cmdFindRecords.Parameters.Append(prmReqRecs)
						prmReqRecs.Value = CleanNumeric(Session("FindRecords"))

						Dim prmIsFirstPage = cmdFindRecords.CreateParameter("isFirstPage", 11, 2)	' 11=bit, 2=output
						cmdFindRecords.Parameters.Append(prmIsFirstPage)

						Dim prmIsLastPage = cmdFindRecords.CreateParameter("isLastPage", 11, 2)	' 11=bit, 2=output
						cmdFindRecords.Parameters.Append(prmIsLastPage)

						Dim prmLocateValue = cmdFindRecords.CreateParameter("locateValue", 200, 1, 2147483646)
						cmdFindRecords.Parameters.Append(prmLocateValue)
						prmLocateValue.Value = Session("locateValue")

						Dim prmColumnType = cmdFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
						cmdFindRecords.Parameters.Append(prmColumnType)

						Dim prmColumnSize = cmdFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
						cmdFindRecords.Parameters.Append(prmColumnSize)

						Dim prmColumnDecimals = cmdFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
						cmdFindRecords.Parameters.Append(prmColumnDecimals)

						Dim prmAction = cmdFindRecords.CreateParameter("action", 200, 1, 255)
						cmdFindRecords.Parameters.Append(prmAction)
						prmAction.Value = Session("action")

						Dim prmTotalRecCount = cmdFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
						cmdFindRecords.Parameters.Append(prmTotalRecCount)

						Dim prmFirstRecPos = cmdFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
						cmdFindRecords.Parameters.Append(prmFirstRecPos)
						prmFirstRecPos.Value = CleanNumeric(Session("firstRecPos"))

						Dim prmCurrentRecCount = cmdFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
						cmdFindRecords.Parameters.Append(prmCurrentRecCount)
						prmCurrentRecCount.Value = CleanNumeric(Session("currentRecCount"))

						Dim prmDecSeparator = cmdFindRecords.CreateParameter("decSeparator", 200, 1, 255)	' 200=varchar, 1=input, 255=size
						cmdFindRecords.Parameters.Append(prmDecSeparator)
						prmDecSeparator.Value = Session("LocaleDecimalSeparator")

						Dim prmDateFormat = cmdFindRecords.CreateParameter("dateFormat", 200, 1, 255)	' 200=varchar, 1=input, 255=size
						cmdFindRecords.Parameters.Append(prmDateFormat)
						prmDateFormat.Value = Session("LocaleDateFormat")
							
						Err.Clear()

						' Get the recordset parameters
						Dim rstParameters = cmdFindRecords.Execute
								
						sThousandColumns = rstParameters.Fields("ThousandColumns").Value.ToString()
						sBlankIfZeroColumns = rstParameters.Fields("BlankIfZeroColumns").Value.ToString()
									
						' Get the actual data
						Dim rstFindRecords = rstParameters.NextRecordset()

						If (Err.Number <> 0) Then
							sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(Err.Description)
						End If

						If Len(sErrorDescription) = 0 Then
							' Instantiate and initialise the grid. 
							Response.Write("<table class='outline' style='width : 100%; ' id='findGridTable'>" & vbCrLf)
							Response.Write("<div id='pager-coldata'></div>" & vbCrLf)
										
							If Len(sErrorDescription) = 0 Then
								If rstFindRecords.State = AD_STATE_OPEN Then
									iCount = 0
									Do While Not rstFindRecords.EOF
										sAddString = ""
						
										For iloop = 0 To (rstFindRecords.Fields.Count - 1)
											If iloop > 0 Then
												sAddString = sAddString & "	"
											End If
							
											If iCount = 0 Then
												sColDef = Replace(rstFindRecords.Fields(iloop).Name, "_", " ") & "	" & rstFindRecords.Fields(iloop).Type
												Response.Write("<INPUT type='hidden' id=txtFindColDef_" & iloop & " name=txtFindColDef_" & iloop & " value=""" & sColDef & """>" & vbCrLf)
											End If
							
											If rstFindRecords.Fields(iloop).Type = 135 Then
												' Field is a date so format as such.
												sAddString = sAddString & ConvertSQLDateToLocale(rstFindRecords.Fields(iloop).Value)
											ElseIf rstFindRecords.Fields(iloop).Type = 131 Then
												' Field is a numeric so format as such.
												If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
													If Mid(sThousandColumns, iloop + 1, 1) = "1" Then
														sTemp = ""
														sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).NumericScale, True, False, True)
													Else
														sTemp = ""
														sTemp = FormatNumber(rstFindRecords.Fields(iloop).Value, rstFindRecords.Fields(iloop).NumericScale, True, False, False)
													End If
													sTemp = Replace(sTemp, ".", "x")
													sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
													sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
													sAddString = sAddString & sTemp
												End If
											Else
												If Not IsDBNull(rstFindRecords.Fields(iloop).Value) Then
													sAddString = sAddString & Replace(rstFindRecords.Fields(iloop).Value, """", "&quot;")
												End If
											End If
										Next

										Response.Write("<INPUT type='hidden' id=txtAddString_" & iCount & " name=txtAddString_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
					
										iCount = iCount + 1
										rstFindRecords.MoveNext()
									Loop
	
									' Release the ADO recordset object.
									rstFindRecords.Close()
								End If
							End If
							rstFindRecords = Nothing
							Response.Write("</table>")
	
							' NB. IMPORTANT ADO NOTE.
							' When calling a stored procedure which returns a recorddim AND has output parameters
							' you need to close the recorddim and dim it to nothing before using the output parameters. 
							If cmdFindRecords.Parameters("error").Value <> 0 Then
								sErrorDescription = "Error reading order definition."
							Else
								If cmdFindRecords.Parameters("someSelectable").Value = 0 Then
									sErrorDescription = "You do not have permission to read any of the selected order's find columns."
								End If
							End If
			
							Response.Write("<INPUT type='hidden' id=txtInsertGranted name=txtInsertGranted value=" & cmdFindRecords.Parameters("insertGranted").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtDeleteGranted name=txtDeleteGranted value=" & cmdFindRecords.Parameters("deleteGranted").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=" & iCount & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtFindRecords name=txtFindRecords value=" & Session("FindRecords") & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtCurrentRecCount name=txtCurrentRecCount value=" & iCount & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtCancelDateColumn name=txtCancelDateColumn value=" & fCancelDateColumn & ">" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtGotoAction name=txtGotoAction value=" & Session("action") & ">" & vbCrLf)
			
							Session("realSource") = cmdFindRecords.Parameters("realSource").Value
						End If
	
						' Release the ADO command object.
						cmdFindRecords = Nothing
					End If
				%>
			</div>
			<%
				If Len(sErrorDescription) = 0 Then
					' Get the summary fields (if required).
					If Session("parentTableID") > 0 Then
						Dim cmdSummaryFields = New ADODB.Command
						cmdSummaryFields.CommandText = "sp_ASRIntGetSummaryFields"
						cmdSummaryFields.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
						cmdSummaryFields.ActiveConnection = Session("databaseConnection")

						Dim prmHistoryTableID = cmdSummaryFields.CreateParameter("historyTableID", 3, 1)	'Type 3 = integer, Direction 1 = Input
						cmdSummaryFields.Parameters.Append(prmHistoryTableID)
						prmHistoryTableID.Value = CleanNumeric(Session("tableID"))

						Dim prmParentTableID = cmdSummaryFields.CreateParameter("parentTableID", 3, 1) 'Type 3 = integer, Direction 1 = Input
						cmdSummaryFields.Parameters.Append(prmParentTableID)
						prmParentTableID.Value = CleanNumeric(Session("parentTableID"))

						Dim prmParentRecordID = cmdSummaryFields.CreateParameter("parentRecordID", 3, 1)	'Type 3 = integer, Direction 1 = Input
						cmdSummaryFields.Parameters.Append(prmParentRecordID)
						prmParentRecordID.Value = CleanNumeric(Session("parentRecordID"))

						Dim prmCanSelect = cmdSummaryFields.CreateParameter("canSelect", 11, 2)	'Type 11 = bit, Direction 2 = Output
						cmdSummaryFields.Parameters.Append(prmCanSelect)
	
						Err.Clear()
						Dim rstSummaryFields = cmdSummaryFields.Execute

						If (Err.Number <> 0) Then
							sErrorDescription = "The summary field definition could not be retrieved." & vbCrLf & FormatError(Err.Description)
						End If

						Dim sThousSepSummaryFields As String
						Dim aSummaryFields(0, 0) As String
						Dim iTotalCount As Integer
								
						If Len(sErrorDescription) = 0 Then
							sThousSepSummaryFields = ","
							' Read the summary field definitions into an array.
							' We do this as we may be doing a lot of jumping around
							' the definitions and its easy to jump around an array than
							' a recordset.
							ReDim aSummaryFields(9, 0)
							Do While Not rstSummaryFields.EOF
								iTotalCount = UBound(aSummaryFields, 2) + 1
								ReDim Preserve aSummaryFields(9, iTotalCount)

								aSummaryFields(1, iTotalCount) = rstSummaryFields.Fields(1).Value
								aSummaryFields(2, iTotalCount) = rstSummaryFields.Fields(2).Value
								aSummaryFields(3, iTotalCount) = rstSummaryFields.Fields(3).Value
								aSummaryFields(4, iTotalCount) = rstSummaryFields.Fields(4).Value
								aSummaryFields(5, iTotalCount) = rstSummaryFields.Fields(5).Value
								aSummaryFields(6, iTotalCount) = rstSummaryFields.Fields(6).Value
								aSummaryFields(7, iTotalCount) = rstSummaryFields.Fields(7).Value
								aSummaryFields(8, iTotalCount) = rstSummaryFields.Fields(8).Value
								aSummaryFields(9, iTotalCount) = rstSummaryFields.Fields(9).Value
	
								If rstSummaryFields.Fields(9).Value Then
									sThousSepSummaryFields = sThousSepSummaryFields & CStr(rstSummaryFields.Fields(3).Value) & ","
								End If
					
								rstSummaryFields.MoveNext()
							Loop

							' Release the ADO recorddim object.
							rstSummaryFields.Close()
							rstSummaryFields = Nothing

							Dim iRowCount = CLng((iTotalCount + 1) / 2)

							If iTotalCount > 0 Then
								Response.Write("			<div id='row3' style='margin-top: 25px;'>" & vbCrLf)
								Response.Write("<table>" & vbCrLf)
								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("				  <TD colspan=5 height=10></TD>" & vbCrLf)
								Response.Write("				</TR>" & vbCrLf)
								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("				  <TD colspan=5 align=center height=10>" & vbCrLf)
								Response.Write("    				<STRONG>History Summary</STRONG>" & vbCrLf)
								Response.Write("  				</TD>" & vbCrLf)
								Response.Write("				</TR>" & vbCrLf)
								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("				  <TD colspan=5 height=10></TD>" & vbCrLf)
								Response.Write("				</TR>" & vbCrLf)

								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("  				<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)

								Response.Write("  				<TD width=""48%"" height=10>" & vbCrLf)
								Response.Write("      			<TABLE WIDTH=100% class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
							End If

							For iLoop = 1 To iRowCount
								Response.Write("   						<TR>" & vbCrLf)
								Response.Write("   							<TD nowrap=true>" & Replace(aSummaryFields(2, iLoop), "_", " ") & " :</TD>" & vbCrLf)
								Response.Write("								<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
								Response.Write("								<TD width=""100%"">" & vbCrLf)

								If aSummaryFields(7, iLoop) = 1 Then
									' The summary control is a checkbox.
			%>

			<input type="checkbox" id="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
				name="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
				disabled="disabled">
			<%Else%>
			<%--' The summary control is not a checkbox. Use a textbox for everything else.--%>
			<input type="text" id="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
				name="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
				class="text textdisabled" disabled="disabled"
				<%If aSummaryFields(8, iLoop) = 1 Then%>
				style="width: 100%; text-align: right" />
			<%ElseIf aSummaryFields(8, iLoop) = 2 Then%>
					style="width: 100%;text-align: center" />
				<% End If%>
			<%
			End If
			Response.Write("								</TD>" & vbCrLf)
			Response.Write("							</TR>" & vbCrLf)
		Next
		
		If iTotalCount > 0 Then
			Response.Write("      			</TABLE>" & vbCrLf)
			Response.Write("      		</TD>" & vbCrLf)
			Response.Write("  				<TD width=100 height=10>&nbsp;&nbsp;&nbsp;&nbsp;</TD>" & vbCrLf)
			' Do the second column now.
			Response.Write("  				<TD width=""48%"" height=10>" & vbCrLf)
			Response.Write("      			<TABLE WIDTH=100% class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
		End If
				
		Dim iColumn2Index As Integer
				
		For iLoop = 1 To iRowCount
			iColumn2Index = iLoop + iRowCount
						
			If iColumn2Index <= iTotalCount Then
				Response.Write("   						<TR>" & vbCrLf)
				Response.Write("								<TD nowrap=true>" & Replace(aSummaryFields(2, iColumn2Index), "_", " ") & " :</TD>" & vbCrLf)
				Response.Write("								<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
				Response.Write("								<TD width=""100%"">" & vbCrLf)

				If aSummaryFields(7, iColumn2Index) = 1 Then%>
			<%--The summary control is a checkbox.--%>
			<input type="checkbox" id="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				name="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				disabled="disabled">
			<%Else%>
			<%--The summary control is not a checkbox. Use a textbox for everything else.--%>
			<input type="text" id="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				name="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
				disabled="disabled" class="text textdisabled"
				<%If aSummaryFields(8, iColumn2Index) = 1 Then%>
				style="width: 100%; text-align: right" />
			<%ElseIf aSummaryFields(8, iColumn2Index) = 2 Then%>
					style="width: 100%;text-align: center" />
				<%End If%>
			<%	
			End If
		End If

		Response.Write("								</TD>" & vbCrLf)
		Response.Write("							</TR>" & vbCrLf)
	Next

	If iTotalCount > 0 Then
		Response.Write("      			</TABLE>" & vbCrLf)
		Response.Write("      		</TD>" & vbCrLf)
		Response.Write("  				<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
		Response.Write("				</TR>" & vbCrLf)
	End If
End If
Response.Write("</table>" & vbCrLf)
Response.Write("</div>" & vbCrLf)
					
' NB. IMPORTANT ADO NOTE.
' When calling a stored procedure which returns a recorddim AND has output parameters
' you need to close the recorddim and dim it to nothing before using the output parameters. 
Dim fCanSelect = cmdSummaryFields.Parameters("canSelect").Value

' Release the ADO command object.
cmdSummaryFields = Nothing

If fCanSelect Then
	Dim cmdSummaryValues = New ADODB.Command
	cmdSummaryValues.CommandText = "spASRIntGetSummaryValues"
	cmdSummaryValues.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
	cmdSummaryValues.ActiveConnection = Session("databaseConnection")

	Dim prmHistoryTableID2 = cmdSummaryValues.CreateParameter("historyTableID", 3, 1)	'Type 3 = integer, Direction 1 = Input
	cmdSummaryValues.Parameters.Append(prmHistoryTableID2)
	prmHistoryTableID2.Value = CleanNumeric(Session("tableID"))

	Dim prmParentTableID2 = cmdSummaryValues.CreateParameter("parentTableID", 3, 1)	'Type 3 = integer, Direction 1 = Input
	cmdSummaryValues.Parameters.Append(prmParentTableID2)
	prmParentTableID2.Value = CleanNumeric(Session("parentTableID"))

	Dim prmParentRecordID2 = cmdSummaryValues.CreateParameter("parentRecordID", 3, 1)	'Type 3 = integer, Direction 1 = Input
	cmdSummaryValues.Parameters.Append(prmParentRecordID2)
	prmParentRecordID2.Value = CleanNumeric(Session("parentRecordID"))

	Err.Clear()
	Dim rstSummaryValues = cmdSummaryValues.Execute

	If (Err.Number <> 0) Then
		sErrorDescription = "The summary field values could not be retrieved." & vbCrLf & FormatError(Err.Description)
	End If
	Dim sTempValue As String
					
	If Len(sErrorDescription) = 0 Then
		If Not (rstSummaryValues.EOF And rstSummaryValues.BOF) Then
			For iLoop = 0 To (rstSummaryValues.Fields.Count - 1)
				If rstSummaryValues.Fields(iLoop).Type = 131 Then
					sTemp = "," & rstSummaryValues.Fields(iLoop).Name & ","

					If IsDBNull(rstSummaryValues.Fields(iLoop).Value) Then
						sTempValue = "0"
					Else
						sTempValue = rstSummaryValues.Fields(iLoop).Value
					End If

					If InStr(sThousSepSummaryFields, sTemp) > 0 Then
						sTemp = ""
						sTemp = FormatNumber(sTempValue, rstSummaryValues.Fields(iLoop).NumericScale, True, False, True)
					Else
						sTemp = ""
						sTemp = FormatNumber(sTempValue, rstSummaryValues.Fields(iLoop).NumericScale, True, False, False)
					End If
					sTemp = Replace(sTemp, ".", "x")
					sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
					sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
							
					Response.Write("			<INPUT type='hidden' id=txtSummaryData_" & rstSummaryValues.Fields(iLoop).Name & " name=txtSummaryData_" & rstSummaryValues.Fields(iLoop).Name & " value=""" & sTemp & """>" & vbCrLf)
				Else
					Response.Write("			<INPUT type='hidden' id=txtSummaryData_" & rstSummaryValues.Fields(iLoop).Name & " name=txtSummaryData_" & rstSummaryValues.Fields(iLoop).Name & " value=""" & rstSummaryValues.Fields(iLoop).Value & """>" & vbCrLf)
				End If
			Next
		End If

		rstSummaryValues.Close()
	End If

	rstSummaryValues = Nothing
	cmdSummaryValues = Nothing
End If
End If
End If
	
If Len(sErrorDescription) = 0 Then
Response.Write("				<INPUT type='hidden' id=txtCurrentTableID name=txtCurrentTableID value=" & Session("tableID") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtCurrentViewID name=txtCurrentViewID value=" & Session("viewID") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtCurrentScreenID name=txtCurrentScreenID value=" & Session("screenID") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtCurrentOrderID name=txtCurrentOrderID value=" & Session("orderID") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtCurrentRecordID name=txtCurrentRecordID value=" & Session("recordID") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtCurrentParentTableID name=txtCurrentParentTableID value=" & Session("parentTableID") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtCurrentParentRecordID name=txtCurrentParentRecordID value=" & Session("parentRecordID") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtRealSource name=txtRealSource value=" & Session("realSource") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtLineage name=txtLineage value=" & Session("lineage") & ">" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtFilterDef name=txtFilterDef value=""" & Replace(Session("filterDef"), """", "&quot;") & """>" & vbCrLf)
Response.Write("				<INPUT type='hidden' id=txtFilterSQL name=txtFilterSQL value=""" & Replace(Session("filterSQL"), """", "&quot;") & """>" & vbCrLf)
End If

Response.Write("				<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
			%>
		</div>
	</form>

	<form id="frmTBData" name="frmTBData">
		<%
			If CLng(Session("tableID")) = CLng(Session("TB_TBTableID")) Then
				Response.Write("				<INPUT type='hidden' id=txtTBCancelCourseDate name=txtTBCancelCourseDate value=""" & Session("lineage") & """>")
			End If
		%>
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>

	<script type="text/javascript"> find_window_onload();</script>
</div>
