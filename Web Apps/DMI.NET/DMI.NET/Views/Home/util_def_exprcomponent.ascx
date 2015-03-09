<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET.Classes" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_expressions")%>" type="text/javascript"></script>

<form action="" method="POST" id="frmMainForm" name="frmMainForm">
	<%
		Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
		Dim iPassBy As Integer
		Dim sErrMsg As String
		
		iPassBy = 1
		If Session("optionFunctionID") > 0 Then
		
			Try

				Dim prmPassByType As New SqlParameter("piResult", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				objDataAccess.ExecuteSP("spASRIntGetParameterPassByType" _
					, New SqlParameter("piFunctionID", SqlDbType.Int) With {.Value = CleanNumeric(CInt(Session("optionFunctionID")))} _
					, New SqlParameter("piParamIndex", SqlDbType.Int) With {.Value = CleanNumeric(CInt(Session("optionParameterIndex")))} _
					, prmPassByType)

				iPassBy = CInt(prmPassByType.Value)
				
			Catch ex As Exception
				sErrMsg = "Error checking parameter pass-by type." & vbCrLf & FormatError(ex.Message)
				
			End Try
			
		End If
		Response.Write("<input type='hidden' id=txtPassByType name=txtPassByType value=" & iPassBy & ">" & vbCrLf)
	%>
	<div class="gridwithbuttons">
		<div class="divExpressionTypes stretchyfixed200">
			<h3>Type</h3>
			<input id="optType_Field" name="optType" type="radio" selected onclick="changeType(1)" />
			<label tabindex="-1" for="optType_Field">Field</label>
			<br />
			<input id="optType_Operator" name="optType" type="radio" <%If iPassBy = 2 Then Response.Write("disabled class='disabled'")%> onclick="changeType(5)" />
			<label <%If iPassBy = 2 Then Response.Write("class='disabled'")%> tabindex="-1" for="optType_Operator">Operator</label>
			<br />
			<input id="optType_Function" name="optType" type="radio" <%If iPassBy = 2 Then Response.Write("disabled class='disabled'")%> onclick="changeType(2)" />
			<label <%If iPassBy = 2 Then Response.Write("class='disabled'")%> tabindex="-1" for="optType_Function">Function</label>
			<br />
			<input id="optType_Value" name="optType" type="radio" <%If iPassBy = 2 Then Response.Write("disabled class='disabled'")%> onclick="changeType(4)" />
			<label <%If iPassBy = 2 Then Response.Write("class='disabled'")%> tabindex="-1" for="optType_Value">Value</label>
			<br />
			<input id="optType_LookupTableValue" name="optType" type="radio" <%If iPassBy = 2 Then Response.Write("disabled class='disabled'")%> onclick="changeType(6)" />
			<label <%If iPassBy = 2 Then Response.Write("class='disabled'")%> tabindex="-1" for="optType_LookupTableValue">Lookup Table Value</label>
			<br />
			<div id="trType_PVal">
				<input id="optType_PromptedValue" name="optType" type="radio" <%If iPassBy = 2 Then Response.Write("disabled class='disabled'")%> onclick="changeType(7)" />
				<label <%If iPassBy = 2 Then Response.Write("class='disabled'")%> tabindex="-1" for="optType_PromptedValue">Prompted Value</label>
			</div>
			<br />
			<div id="trType_Calc">
				<input id="optType_Calculation" name="optType" type="radio" <%	 If iPassBy = 2 Then Response.Write("disabled class='disabled'")%> onclick="changeType(3)" />
				<label <%If iPassBy = 2 Then Response.Write("class='disabled'")%> tabindex="-1" for="optType_Calculation">Calculation</label>
			</div>
			<br />
			<div id="trType_Filter">
				<input id="optType_Filter" name="optType" type="radio" <%If iPassBy = 2 Then Response.Write("disabled class='disabled'")%> onclick="changeType(10)" />
				<label <%If iPassBy = 2 Then Response.Write("class='disabled'")%> tabindex="-1" for="optType_Filter">Filter</label>
			</div>
		</div>

		<div class="divExpressionOptions stretchyfill">
			<div id="divField">
				<h3>Field</h3>
				<%If iPassBy = 1 Then%>
				<input id="optField_Field" name="optField" type="radio" selected onclick="field_refreshTable()" />
				<label tabindex="-1" for="optField_Field">Field</label>
				<input id="optField_Count" name="optField" type="radio" selected onclick="field_refreshTable()" />
				<label tabindex="-1" for="optField_Count">Count</label>
				<input id="optField_Total" name="optField" type="radio" selected onclick="field_refreshTable()" />
				<label tabindex="-1" for="optField_Total">Total</label>
				<%End If%>
				<div class="formField">
					<label>Table :</label>
					<select id="cboFieldTable" name="cboFieldTable" onchange="field_changeTable()"></select>
				</div>
				<div class="formField">
					<label>Column :</label>
					<select id="cboFieldColumn" name="cboFieldColumn"></select>
					<select id="cboFieldDummyColumn" name="cboFieldDummyColumn" style="display: none" disabled=""></select>
				</div>
				<%If iPassBy = 1 Then%>
				<br/>
				<h3>Child Field Options</h3>
				<span style="display: inline-block">
					<input id="optFieldRecSel_First" name="optFieldRecSel" type="radio" onclick="field_refreshChildFrame()" />
					<label tabindex="-1" for="optFieldRecSel_First">First</label>
					<input id="optFieldRecSel_Last" name="optFieldRecSel" type="radio" onclick="field_refreshChildFrame()" />
					<label tabindex="-1" for="optFieldRecSel_Last">Last</label>
					<input id="optFieldRecSel_Specific" name="optFieldRecSel" type="radio" onclick="field_refreshChildFrame()" />
					<div id="divFieldRecSel_Specific" style="display: none">
						<label tabindex="-1" for="optFieldRecSel_Specific">Specific</label>
					</div>
					<input id="txtFieldRecSel_Specific" name="txtFieldRecSel_Specific">
				</span>
				<div class="formField">
					<label>Order :</label>
					<input type="text" id="txtFieldRecOrder" name="txtFieldRecOrder" disabled="disabled">
					<input id="btnFieldRecOrder" name="btnFieldRecOrder" class="btn" type="button" value="..." onclick="field_selectRecOrder()" />
				</div>
				<div class="formField">
					<label>Filter :</label>
					<input type="text" id="txtFieldRecFilter" name="txtFieldRecFilter" disabled="disabled">
					<input id="btnFieldRecFilter" name="btnFieldRecFilter" class="btn" type="button" value="..." onclick="field_selectRecFilter()" />
				</div>
				<%End If%>
			</div>
			

			<div id="divFunction" style="display: none">
				<h3>Function</h3>
				<div id="SSFunctionTree" style="left: 0; top: 0; width: 100%;">
					<ul>
						<li id="FUNCTION_ROOT" class="root"><a href='#'>Functions</a></li>
					</ul>
				</div>
			</div>

			<div id="divOperator" style="display: none">
				<h3>Operator</h3>
				<div id="SSOperatorTree" style="left: 0; top: 0; width: 100%;">
					<ul>
						<li id="OPERATOR_ROOT" class="root"><a href='#'>Operators</a></li>
					</ul>
				</div>
			</div>

			<div id="divValue" style="display: none">
				<h3>Value</h3>
				<div class="formField">
					<label>Type :</label>
					<select id="cboValueType" name="cboValueType" onchange="value_changeType()">
						<option value="1">Character</option>
						<option value="2">Numeric</option>
						<option value="3">Logic</option>
						<option value="4">Date</option>
					</select>
				</div>
				<div class="formField">
					<label>Value :</label>
					<select id="selectValue" name='selectValue'>
						<option value="1">True</option>
						<option value="0">False</option>
					</select>
					<input id="txtValue" name="txtValue">
				</div>
			</div>

			<div id="divLookupValue" style="display: none">
				<h3>Lookup Table Value</h3>
				<div class="formField">
					<label>Table :</label>
					<select id="cboLookupValueTable" name="cboLookupValueTable" onchange="lookupValue_changeTable()"></select>
				</div>
				<div class="formField">
					<label>Column :</label>
					<select id="cboLookupValueColumn" name="cboLookupValueColumn" onchange="lookupValue_changeColumn()"></select>
				</div>
				<div class="formField">
					<label>Value :</label>
					<select id="cboLookupValueValue" name="cboLookupValueValue"></select>
					<input type="text" class="textwarning" id="txtValueNotInLookup" name="txtValueNotInLookup" value="<value> does not appear in <table>.<column>" style="text-align: left; display: none" readonly>
				</div>
			</div>

			<div id="divCalculation" style="display: none">
				<h3>Calculation</h3>
				<table id="ssOleDBGridCalculations"></table>
				<br />
				<textarea id="txtCalcDescription" name="txtCalcDescription" tabindex="-1" style="width: 100%; white-space: normal !important;" wrap="VIRTUAL" disabled="disabled"></textarea>
				<input <%	If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" name="chkOwnersCalcs" id="chkOwnersCalcs" value="chkOwnersCalcs" tabindex="0"
					onclick="calculationAndFilter_refresh();" />
				<label for="chkOwnersCalcs" class="checkbox" tabindex="-1" onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}">
					Only show calculations where owner is '<%:Session("Username")%>'
				</label>
			</div>

			<div id="divFilter" style="display: none">
				<h3>Filter</h3>
				<table id="ssOleDBGridFilters"></table>
				<br />
				<textarea id="txtFilterDescription" name="txtFilterDescription" tabindex="-1" style="width: 100%; white-space: normal !important;" wrap="VIRTUAL" disabled="disabled"></textarea>
				<input <%	If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" name="chkOwnersFilters" id="chkOwnersFilters" value="chkOwnersFilters" tabindex="0"
					onclick="calculationAndFilter_refresh();" />
				<label for="chkOwnersFilters" class="checkbox" tabindex="-1">
					Only show filters where owner is '<%:Session("Username") %>'
				</label>
			</div>

			<div id="divPromptedValue" style="display: none">
				<h3>Prompted Value</h3>
				<div class="formField">
					<label>Prompt :</label>
					<input id="Text1" name="txtPrompt" onkeyup="pVal_changePrompt()">
				</div>
				<h3>Type</h3>
				<select id="cboPValType" name="cboPValType" onchange="pVal_changeType()">
					<option value="1">Character</option>
					<option value="2">Numeric</option>
					<option value="3">Logic</option>
					<option value="4">Date</option>
					<option value="5">Lookup Value</option>
				</select>
				<label id="tdPValSizePrompt" name="tdPValSizePrompt">Size :</label>
				<input id="txtPValSize" name="txtPValSize">
				<label id="tdPValDecimalsPrompt" name="tdPValDecimalsPrompt">Decimals :</label>
				<input id="txtPValDecimals" name="txtPValDecimals">
				<div id="trPValFormat">
					<h3>Mask</h3>
					<input id="txtPValFormat" name="txtPValFormat">
					<table>
						<tr>
							<td>A - Uppercase</td>
							<td>9 - Numbers (0-9)</td>
							<td>B - Binary (0 or 1)</td>
						</tr>
						<tr>
							<td>a - Lowercase</td>
							<td># - Numbers, Symbols</td>
							<td>\ - Follow by any literal</td>
						</tr>
					</table>
				</div>

				<div id="trPValFormat2">
				</div>

				<div id="trPValLookup" style="display: none">
					<h3>Lookup Table Value</h3>
					<div class="formField">
						<label>Table :</label>
						<select id="cboPValTable" name="cboPValTable" onchange="pVal_changeTable()">
						</select>
					</div>
					<div class="formField">
						<label>Column :</label>
						<select id="cboPValColumn" name="cboPValColumn" onchange="pVal_changeColumn()">
						</select>
					</div>
				</div>

				<div id="trPValLookup2">
				</div>

				<h3>Default Value</h3>
				<div id="trPValDateOptions">
					<table>
						<tr>
							<td>
								<input id="optPValDate_Explicit" name="optPValDate" type="radio" selected onclick="pVal_changeDateOption(0)" />
								<label tabindex="-1" for="optPValDate_Explicit">Explicit</label>
							</td>
							<td>
								<input id="optPValDate_MonthStart" name="optPValDate" type="radio"
									onclick="pVal_changeDateOption(2)" />
								<label tabindex="-1" for="optPValDate_MonthStart">Month Start</label>
							</td>
							<td>
								<input id="optPValDate_YearStart" name="optPValDate" type="radio" onclick="pVal_changeDateOption(4)" />
								<label tabindex="-1" for="optPValDate_YearStart">Year Start</label>
							</td>
						</tr>
						<tr>
							<td>
								<input id="optPValDate_Current" name="optPValDate" type="radio" onclick="pVal_changeDateOption(1)" />
								<label tabindex="-1" for="optPValDate_Current">Current</label>
							</td>
							<td>
								<input id="optPValDate_MonthEnd" name="optPValDate" type="radio" onclick="pVal_changeDateOption(3)" />
								<label tabindex="-1" for="optPValDate_MonthEnd">Month End</label>
							</td>
							<td>
								<input id="optPValDate_YearEnd" name="optPValDate" type="radio" onclick="pVal_changeDateOption(5)" />
								<label tabindex="-1" for="optPValDate_YearEnd">Year End</label>
							</td>
						</tr>
					</table>
				</div>

				<div id="trPValDateOptions2">
				</div>

				<div id="trPValTextDefault">
					<input id="txtPValDefault" name="txtPValDefault">
				</div>

				<div id="trPValComboDefault" style="display: none">
					<select id="cboPValDefault" name="cboPValDefault">
					</select>
				</div>

			</div>
		</div>
	</div>
</form>

<div id="util_def_exprcomponent_frmUseful">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtExprType" name="txtExprType" value='<%=session("optionExprType")%>'>
	<input type="hidden" id="txtExprID" name="txtExprID" value='<%=session("optionExprID")%>'>
	<input type="hidden" id="txtAction" name="txtAction" value='<%=session("optionAction")%>'>
	<input type="hidden" id="txtLinkRecordID" name="txtLinkRecordID" value='<%=session("optionLinkRecordID")%>'>
	<input type="hidden" id="txtTableID" name="txtTableID" value='<%=session("optionTableID")%>'>
	<input type="hidden" id="txtInitialising" name="txtInitialising" value="0">
	<input type="hidden" id="txtChildFieldOrderID" name="txtChildFieldOrderID" value="0">
	<input type="hidden" id="txtChildFieldFilterID" name="txtChildFieldFilterID" value="0">
	<input type="hidden" id="txtChildFieldFilterHidden" name="txtChildFieldFilterHidden" value="0">
	<input type="hidden" id="txtFunctionsLoaded" name="txtFunctionsLoaded" value="0">
	<input type="hidden" id="txtOperatorsLoaded" name="txtOperatorsLoaded" value="0">
	<input type="hidden" id="txtLookupTablesLoaded" name="txtLookupTablesLoaded" value="0">
	<input type="hidden" id="txtPValLookupTablesLoaded" name="txtPValLookupTablesLoaded" value="0">
</div>

<form id="util_def_exprcomponent_frmOriginalDefinition" name="util_def_exprcomponent_frmOriginalDefinition">
	<%
		Dim sDefnString As String
		Dim sFieldTableID As String
		Dim sFieldColumnID As String
		Dim sLookupTableID As String
		Dim sLookupColumnID As String
	
		If Session("optionAction") = OptionActionType.EDITEXPRCOMPONENT Then
			sDefnString = Session("optionExtension").ToString()

			Response.Write("<INPUT type='hidden' id=txtComponentID name=txtComponentID value=" & ComponentParameter(sDefnString, "COMPONENTID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtType name=txtType value=" & ComponentParameter(sDefnString, "TYPE") & ">" & vbCrLf)
			sFieldTableID = ComponentParameter(sDefnString, "FIELDTABLEID")
			sFieldColumnID = ComponentParameter(sDefnString, "FIELDCOLUMNID")
			Response.Write("<INPUT type='hidden' id=txtFieldPassBy name=txtFieldPassBy value=" & ComponentParameter(sDefnString, "FIELDPASSBY") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionTableID name=txtFieldSelectionTableID value=" & ComponentParameter(sDefnString, "FIELDSELECTIONTABLEID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionRecord name=txtFieldSelectionRecord value=" & ComponentParameter(sDefnString, "FIELDSELECTIONRECORD") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionLine name=txtFieldSelectionLine value=" & ComponentParameter(sDefnString, "FIELDSELECTIONLINE") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionOrderID name=txtFieldSelectionOrderID value=" & ComponentParameter(sDefnString, "FIELDSELECTIONORDERID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionFilter name=txtFieldSelectionFilter value=" & ComponentParameter(sDefnString, "FIELDSELECTIONFILTER") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFunctionID name=txtFunctionID value=" & ComponentParameter(sDefnString, "FUNCTIONID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtCalculationID name=txtCalculationID value=" & ComponentParameter(sDefnString, "CALCULATIONID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOperatorID name=txtOperatorID value=" & ComponentParameter(sDefnString, "OPERATORID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueType name=txtValueType value=" & ComponentParameter(sDefnString, "VALUETYPE") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueCharacter name=txtValueCharacter value=""" & Replace(ComponentParameter(sDefnString, "VALUECHARACTER"), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueNumeric name=txtValueNumeric value=" & ComponentParameter(sDefnString, "VALUENUMERIC") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueLogic name=txtValueLogic value=" & ComponentParameter(sDefnString, "VALUELOGIC") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueDate name=txtValueDate value=" & ComponentParameter(sDefnString, "VALUEDATE") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptDescription name=txtPromptDescription value=""" & Replace(ComponentParameter(sDefnString, "PROMPTDESCRIPTION"), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptMask name=txtPromptMask value=""" & Replace(ComponentParameter(sDefnString, "PROMPTMASK"), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptSize name=txtPromptSize value=" & ComponentParameter(sDefnString, "PROMPTSIZE") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptDecimals name=txtPromptDecimals value=" & ComponentParameter(sDefnString, "PROMPTDECIMALS") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFunctionReturnType name=txtFunctionReturnType value=" & ComponentParameter(sDefnString, "FUNCTIONRETURNTYPE") & ">" & vbCrLf)
			sLookupTableID = ComponentParameter(sDefnString, "LOOKUPTABLEID")
			sLookupColumnID = ComponentParameter(sDefnString, "LOOKUPCOLUMNID")
			Response.Write("<INPUT type='hidden' id=txtFilterID name=txtFilterID value=" & ComponentParameter(sDefnString, "FILTERID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldOrderName name=txtFieldOrderName value=""" & ComponentParameter(sDefnString, "FIELDSELECTIONORDERNAME") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldFilterName name=txtFieldFilterName value=""" & ComponentParameter(sDefnString, "FIELDSELECTIONFILTERNAME") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptDateType name=txtPromptDateType value=" & ComponentParameter(sDefnString, "PROMPTDATETYPE") & ">" & vbCrLf)
		Else
			Response.Write("<INPUT type='hidden' id=txtComponentID name=txtComponentID value=0>" & vbCrLf)
			sFieldTableID = Session("optionTableID").ToString()
			sFieldColumnID = "0"
			sLookupTableID = "0"
			sLookupColumnID = "0"
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionRecord name=txtFieldSelectionRecord value=1>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueCharacter name=txtValueCharacter value="""">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueNumeric name=txtValueNumeric value=0>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueLogic name=txtValueLogic value=""False"">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueDate name=txtValueDate value="""">" & vbCrLf)
		End If

		Response.Write("<INPUT type='hidden' id=txtFieldTableID name=txtFieldTableID value=" & sFieldTableID & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtFieldColumnID name=txtFieldColumnID value=" & sFieldColumnID & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtLookupTableID name=txtLookupTableID value=" & sLookupTableID & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtLookupColumnID name=txtLookupColumnID value=" & sLookupColumnID & ">" & vbCrLf)
	%>
</form>

<form id="frmExprTables" name="frmExprTables">
	<%

		Dim iCount As Integer
		
		If Len(sErrMsg) = 0 Then
			
			Try
				Dim rstTables = objDataAccess.GetFromSP("sp_ASRIntGetExprTables" _
					, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))})

				iCount = 0
				For Each objRow As DataRow In rstTables.Rows
					iCount += 1
					Response.Write("<input type='hidden' id=txtTable_" & iCount & " name=txtTable_" & iCount & " value=""" & objRow("definitionString").ToString & """>" & vbCrLf)
				Next
				
			Catch ex As Exception
				sErrMsg = "Error reading component tables." & vbCrLf & FormatError(ex.Message)

			End Try
		End If
	%>
</form>

<form id="frmFunctions" name="frmFunctions">
	<%
		
		If Len(sErrMsg) = 0 Then
			
			Try
				Dim rstFunctions = objDataAccess.GetFromSP("spASRIntGetExprFunctions" _
						, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
						, New SqlParameter("@pbAbsenceEnabled", SqlDbType.Bit) With {.Value = Licence.IsModuleLicenced(SoftwareModule.Absence)})

				iCount = 0
				For Each objRow As DataRow In rstFunctions.Rows
					iCount += 1
					Response.Write("<input type='hidden' id=txtFunction_" & iCount & " name=txtFunction_" & iCount & " value=""" & objRow("definitionString").ToString() & """>" & vbCrLf)
				Next

			Catch ex As Exception
				sErrMsg = "Error reading component functions." & vbCrLf & FormatError(ex.Message)

			End Try
			
		End If
	%>
</form>

<form id="frmFunctionParameters" name="frmFunctionParameters">
	<%
		
		If Len(sErrMsg) = 0 Then
			
			Try
				Dim rstFunctionParameters = objDataAccess.GetFromSP("sp_ASRIntGetExprFunctionParameters")

				iCount = 0
				For Each objRow As DataRow In rstFunctionParameters.Rows
					Response.Write("<input type='hidden' id=txtFunctionParameters_" & objRow("functionID").ToString() & "_" & iCount & " name=txtFunctionParameters_" & objRow("functionID").ToString() & "_" & iCount & " value=""" & objRow("parameterName").ToString() & """>" & vbCrLf)
					iCount += 1
				Next

			Catch ex As Exception
				sErrMsg = "Error reading component functions." & vbCrLf & FormatError(ex.Message)

			End Try
			
		End If
	%>
</form>

<form id="frmOperators" name="frmOperators">
	<%
		
		If Len(sErrMsg) = 0 Then
			
			Try
				Dim rstOperators = objDataAccess.GetFromSP("sp_ASRIntGetExprOperators")

				iCount = 0
				For Each objRow As DataRow In rstOperators.Rows
					iCount += 1
					Response.Write("<input type='hidden' id=txtOperator_" & iCount & " name=txtOperator_" & iCount & " value=""" & objRow("definitionString").ToString & """>" & vbCrLf)
				Next

			Catch ex As Exception
				sErrMsg = "Error reading component operators." & vbCrLf & FormatError(ex.Message)

			End Try
			
		End If
	%>
</form>

<form id="frmCalcs" name="frmCalcs">
	<%
	
		If Len(sErrMsg) = 0 Then

			Try
				Dim rstCalcs = objDataAccess.GetFromSP("sp_ASRIntGetExprCalcs" _
					, New SqlParameter("piCurrentExprID", SqlDbType.Int) With {.Value = CleanNumeric(CInt(Session("optionExprID")))} _
					, New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = CleanNumeric(CInt(Session("optionTableID")))})

				iCount = 0
				
				For Each objRow As DataRow In rstCalcs.Rows
					iCount += 1
					Response.Write("<input type='hidden' id=txtCalc_" & iCount & " name=txtCalc_" & iCount & " value=""" & Replace(objRow("definitionString").ToString(), """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtCalcDesc_" & iCount & " name=txtCalcDesc_" & iCount & " value=""" & Replace(objRow("description").ToString(), """", "&quot;") & """>" & vbCrLf)
				Next
				
			Catch ex As Exception
				sErrMsg = "Error reading component calculations." & vbCrLf & FormatError(ex.Message)
				
			End Try

		End If
	%>
</form>

<form id="frmFilters" name="frmFilters">
	<%
		
		If Len(sErrMsg) = 0 Then
			
			Try
				
				Dim rstFilters = objDataAccess.GetFromSP("sp_ASRIntGetExprFilters" _
					, New SqlParameter("piCurrentExprID", SqlDbType.Int) With {.Value = CleanNumeric(CInt(Session("optionExprID")))} _
					, New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = CleanNumeric(CInt(Session("optionTableID")))})

				iCount = 0
				
				For Each objRow As DataRow In rstFilters.Rows
					iCount += 1
					Response.Write("<input type='hidden' id=txtFilter_" & iCount & " name=txtFilter_" & iCount & " value=""" & Replace(objRow("definitionString").ToString(), """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtFilterDesc_" & iCount & " name=txtFilterDesc_" & iCount & " value=""" & Replace(objRow("description").ToString(), """", "&quot;") & """>" & vbCrLf)
				Next
								
			Catch ex As Exception
				sErrMsg = "Error reading component filters." & vbCrLf & FormatError(ex.Message)

			End Try

		End If
	%>
</form>

<div id="frmFieldRec" name="frmFieldRec" style="visibility: hidden; display: none">
	<input type="hidden" id="selectionType" name="selectionType">
	<input type="hidden" id="Hidden1" name="txtTableID">
	<input type="hidden" id="selectedID" name="selectedID">
</div>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<%
	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrMsg & """>" & vbCrLf)
%>

<script runat="server" language="vb">

	Function ComponentParameter(psDefnString As String, psParameter As String) As String
		Dim iCharIndex As Integer
		Dim sDefn As String
	
		sDefn = psDefnString
	
		iCharIndex = InStr(sDefn, "	")
		If iCharIndex >= 0 Then
			If psParameter = "COMPONENTID" Then
				ComponentParameter = Left(sDefn, iCharIndex - 1)
				Exit Function
			End If
		
			sDefn = Mid(sDefn, iCharIndex + 1)
			iCharIndex = InStr(sDefn, "	")
			If iCharIndex >= 0 Then
				If psParameter = "EXPRID" Then
					ComponentParameter = Left(sDefn, iCharIndex - 1)
					Exit Function
				End If
			
				sDefn = Mid(sDefn, iCharIndex + 1)
				iCharIndex = InStr(sDefn, "	")
				If iCharIndex >= 0 Then
					If psParameter = "TYPE" Then
						ComponentParameter = Left(sDefn, iCharIndex - 1)
						Exit Function
					End If
				
					sDefn = Mid(sDefn, iCharIndex + 1)
					iCharIndex = InStr(sDefn, "	")
					If iCharIndex >= 0 Then
						If psParameter = "FIELDCOLUMNID" Then
							ComponentParameter = Left(sDefn, iCharIndex - 1)
							Exit Function
						End If
					
						sDefn = Mid(sDefn, iCharIndex + 1)
						iCharIndex = InStr(sDefn, "	")
						If iCharIndex >= 0 Then
							If psParameter = "FIELDPASSBY" Then
								ComponentParameter = Left(sDefn, iCharIndex - 1)
								Exit Function
							End If
						
							sDefn = Mid(sDefn, iCharIndex + 1)
							iCharIndex = InStr(sDefn, "	")
							If iCharIndex >= 0 Then
								If psParameter = "FIELDSELECTIONTABLEID" Then
									ComponentParameter = Left(sDefn, iCharIndex - 1)
									Exit Function
								End If
							
								sDefn = Mid(sDefn, iCharIndex + 1)
								iCharIndex = InStr(sDefn, "	")
								If iCharIndex >= 0 Then
									If psParameter = "FIELDSELECTIONRECORD" Then
										ComponentParameter = Left(sDefn, iCharIndex - 1)
										Exit Function
									End If
								
									sDefn = Mid(sDefn, iCharIndex + 1)
									iCharIndex = InStr(sDefn, "	")
									If iCharIndex >= 0 Then
										If psParameter = "FIELDSELECTIONLINE" Then
											ComponentParameter = Left(sDefn, iCharIndex - 1)
											Exit Function
										End If
									
										sDefn = Mid(sDefn, iCharIndex + 1)
										iCharIndex = InStr(sDefn, "	")
										If iCharIndex >= 0 Then
											If psParameter = "FIELDSELECTIONORDERID" Then
												ComponentParameter = Left(sDefn, iCharIndex - 1)
												Exit Function
											End If
										
											sDefn = Mid(sDefn, iCharIndex + 1)
											iCharIndex = InStr(sDefn, "	")
											If iCharIndex >= 0 Then
												If psParameter = "FIELDSELECTIONFILTER" Then
													ComponentParameter = Left(sDefn, iCharIndex - 1)
													Exit Function
												End If
											
												sDefn = Mid(sDefn, iCharIndex + 1)
												iCharIndex = InStr(sDefn, "	")
												If iCharIndex >= 0 Then
													If psParameter = "FUNCTIONID" Then
														ComponentParameter = Left(sDefn, iCharIndex - 1)
														Exit Function
													End If
												
													sDefn = Mid(sDefn, iCharIndex + 1)
													iCharIndex = InStr(sDefn, "	")
													If iCharIndex >= 0 Then
														If psParameter = "CALCULATIONID" Then
															ComponentParameter = Left(sDefn, iCharIndex - 1)
															Exit Function
														End If
													
														sDefn = Mid(sDefn, iCharIndex + 1)
														iCharIndex = InStr(sDefn, "	")
														If iCharIndex >= 0 Then
															If psParameter = "OPERATORID" Then
																ComponentParameter = Left(sDefn, iCharIndex - 1)
																Exit Function
															End If
														
															sDefn = Mid(sDefn, iCharIndex + 1)
															iCharIndex = InStr(sDefn, "	")
															If iCharIndex >= 0 Then
																If psParameter = "VALUETYPE" Then
																	ComponentParameter = Left(sDefn, iCharIndex - 1)
																	Exit Function
																End If
															
																sDefn = Mid(sDefn, iCharIndex + 1)
																iCharIndex = InStr(sDefn, "	")
																If iCharIndex >= 0 Then
																	If psParameter = "VALUECHARACTER" Then
																		ComponentParameter = Left(sDefn, iCharIndex - 1)
																		Exit Function
																	End If
																
																	sDefn = Mid(sDefn, iCharIndex + 1)
																	iCharIndex = InStr(sDefn, "	")
																	If iCharIndex >= 0 Then
																		If psParameter = "VALUENUMERIC" Then
																			ComponentParameter = Left(sDefn, iCharIndex - 1)
																			Exit Function
																		End If
																	
																		sDefn = Mid(sDefn, iCharIndex + 1)
																		iCharIndex = InStr(sDefn, "	")
																		If iCharIndex >= 0 Then
																			If psParameter = "VALUELOGIC" Then
																				ComponentParameter = Left(sDefn, iCharIndex - 1)
																				Exit Function
																			End If
																		
																			sDefn = Mid(sDefn, iCharIndex + 1)
																			iCharIndex = InStr(sDefn, "	")
																			If iCharIndex >= 0 Then
																				If psParameter = "VALUEDATE" Then
																					ComponentParameter = Left(sDefn, iCharIndex - 1)
																					Exit Function
																				End If
																			
																				sDefn = Mid(sDefn, iCharIndex + 1)
																				iCharIndex = InStr(sDefn, "	")
																				If iCharIndex >= 0 Then
																					If psParameter = "PROMPTDESCRIPTION" Then
																						ComponentParameter = Left(sDefn, iCharIndex - 1)
																						Exit Function
																					End If
																				
																					sDefn = Mid(sDefn, iCharIndex + 1)
																					iCharIndex = InStr(sDefn, "	")
																					If iCharIndex >= 0 Then
																						If psParameter = "PROMPTMASK" Then
																							ComponentParameter = Left(sDefn, iCharIndex - 1)
																							Exit Function
																						End If
																					
																						sDefn = Mid(sDefn, iCharIndex + 1)
																						iCharIndex = InStr(sDefn, "	")
																						If iCharIndex >= 0 Then
																							If psParameter = "PROMPTSIZE" Then
																								ComponentParameter = Left(sDefn, iCharIndex - 1)
																								Exit Function
																							End If
																						
																							sDefn = Mid(sDefn, iCharIndex + 1)
																							iCharIndex = InStr(sDefn, "	")
																							If iCharIndex >= 0 Then
																								If psParameter = "PROMPTDECIMALS" Then
																									ComponentParameter = Left(sDefn, iCharIndex - 1)
																									Exit Function
																								End If
																							
																								sDefn = Mid(sDefn, iCharIndex + 1)
																								iCharIndex = InStr(sDefn, "	")
																								If iCharIndex >= 0 Then
																									If psParameter = "FUNCTIONRETURNTYPE" Then
																										ComponentParameter = Left(sDefn, iCharIndex - 1)
																										Exit Function
																									End If
																								
																									sDefn = Mid(sDefn, iCharIndex + 1)
																									iCharIndex = InStr(sDefn, "	")
																									If iCharIndex >= 0 Then
																										If psParameter = "LOOKUPTABLEID" Then
																											ComponentParameter = Left(sDefn, iCharIndex - 1)
																											Exit Function
																										End If
																									
																										sDefn = Mid(sDefn, iCharIndex + 1)
																										iCharIndex = InStr(sDefn, "	")
																										If iCharIndex >= 0 Then
																											If psParameter = "LOOKUPCOLUMNID" Then
																												ComponentParameter = Left(sDefn, iCharIndex - 1)
																												Exit Function
																											End If
																										
																											sDefn = Mid(sDefn, iCharIndex + 1)
																											iCharIndex = InStr(sDefn, "	")
																											If iCharIndex >= 0 Then
																												If psParameter = "FILTERID" Then
																													ComponentParameter = Left(sDefn, iCharIndex - 1)
																													Exit Function
																												End If
																											
																												sDefn = Mid(sDefn, iCharIndex + 1)
																												iCharIndex = InStr(sDefn, "	")
																												If iCharIndex >= 0 Then
																													If psParameter = "EXPANDEDNODE" Then
																														ComponentParameter = Left(sDefn, iCharIndex - 1)
																														Exit Function
																													End If
																												
																													sDefn = Mid(sDefn, iCharIndex + 1)
																													iCharIndex = InStr(sDefn, "	")
																													If iCharIndex >= 0 Then
																														If psParameter = "PROMPTDATETYPE" Then
																															ComponentParameter = Left(sDefn, iCharIndex - 1)
																															Exit Function
																														End If
																													
																														sDefn = Mid(sDefn, iCharIndex + 1)
																														iCharIndex = InStr(sDefn, "	")
																														If iCharIndex >= 0 Then
																															If psParameter = "DESCRIPTION" Then
																																ComponentParameter = Left(sDefn, iCharIndex - 1)
																																Exit Function
																															End If
																														
																															sDefn = Mid(sDefn, iCharIndex + 1)
																															iCharIndex = InStr(sDefn, "	")
																															If iCharIndex >= 0 Then
																																If psParameter = "FIELDTABLEID" Then
																																	ComponentParameter = Left(sDefn, iCharIndex - 1)
																																	Exit Function
																																End If
																															
																																sDefn = Mid(sDefn, iCharIndex + 1)
																																iCharIndex = InStr(sDefn, "	")
																																If iCharIndex >= 0 Then
																																	If psParameter = "FIELDSELECTIONORDERNAME" Then
																																		ComponentParameter = Left(sDefn, iCharIndex - 1)
																																		Exit Function
																																	End If
																																
																																	sDefn = Mid(sDefn, iCharIndex + 1)
																																	If psParameter = "FIELDSELECTIONFILTERNAME" Then
																																		ComponentParameter = sDefn
																																		Exit Function
																																	End If
																																End If
																															End If
																														End If
																													End If
																												End If
																											End If
																										End If
																									End If
																								End If
																							End If
																						End If
																					End If
																				End If
																			End If
																		End If
																	End If
																End If
															End If
														End If
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	
		ComponentParameter = ""
	End Function
</script>

<script type="text/javascript">
	util_def_exprcomponent_onload();
</script>


