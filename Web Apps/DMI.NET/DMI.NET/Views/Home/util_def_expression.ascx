<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_expressions")%>" type="text/javascript"></script>

<div id ="divDefExpression">

<form id="frmDefinition">

	<div class="absolutefull" style="margin-top:10px;">
		<div class="nowrap" id="nav">
			<div class="tablerow">
				<label>Name :</label>
				<input id="txtName" name="txtName" maxlength="50" onkeyup="changeName()" style="width:90%;">				
				<label>Owner :</label>
				<input id="txtOwner" style=" margin-left: 4px; width: 90%;" name="txtOwner" disabled="disabled" tabindex="-1">
			</div>
			<br/>
			<div class="tablerow">
				<label>Description :</label>
				<textarea id="txtDescription" name="txtDescription" wrap="VIRTUAL" style="width:90%; height: 60px; white-space: normal !important;" maxlength="255" onkeyup="changeDescription()"></textarea>
				<label>Access :</label>
				<div>
					<input class="inline-block" id="optAccessRW" name="optAccess" type="radio" onclick="changeAccess()" checked />
					<label class="inline-block" for="optAccessRW">Read/Write</label><br/>
					<input class="inline-block" id="optAccessRO" name="optAccess" type="radio" onclick="changeAccess()"/>
					<label class="inline-block" for="optAccessRO">Read Only</label><br/>
					<input class="inline-block" id="optAccessHD" name="optAccess" type="radio" onclick="changeAccess()" />
					<label class="inline-block" for="optAccessHD">Hidden</label>
				</div>
			</div>
		</div>


		<div class="clearboth">
			<hr />
		</div>

		<div class="gridwithbuttons clearboth">

			<div class="stretchyfill">
				<div id="SSTree1" style="overflow: auto;"></div>
			</div>

			<div class="stretchyfixed">
				<input type="button" id="cmdAdd" name="cmdAdd" class="btn" value="Add" onclick="addClick()" />
				<br />
				<input type="button" id="cmdInsert" name="cmdInsert" class="btn" value="Insert" onclick="insertClick()" />
				<br />
				<input type="button" id="cmdEdit" name="cmdEdit" class="btn" value="Edit" onclick="editClick()" />
				<br />
				<input type="button" id="cmdDelete" name="cmdDelete" class="btn" value="Delete" onclick="deleteClick()" />
				<br />
				<input type="button" id="cmdPrint" name="cmdPrint" class="btn" value="Print" onclick="printClick()" />
				<br />
				<%If Session("utiltype") = 11 Then%>
				<input type="button" id="cmdTest" name="cmdTest" class="btn" value="Test" onclick="testClick()" />
				<br />				
				<%End If%>				
				<input type="button" id="cmdCancel" name="cmdCancel" class="btn" value="Cancel" onclick="cancelClick()" />
			</div>
		</div>
	</div>
	<%=Html.AntiForgeryToken()%>
	</form>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		
		Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
		Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

		Dim sReaction As String = ""
		Dim sUtilTypeName As String
		Dim sErrMsg As String = ""
		Dim iCount As Integer
		
		sUtilTypeName = "expression"
		If Session("utiltype") = 11 Then
			sUtilTypeName = "filter"
			sReaction = "FILTERS"
		Else
			If Session("utiltype") = 12 Then
				sUtilTypeName = "calculation"
				sReaction = "CALCULATIONS"
			End If
		End If

		If Session("action") <> "new" Then

			Try

				Dim prmErrMsg As New SqlParameter("psErrMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmTimestamp As New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				Dim rstDefinition = objDataAccess.GetFromSP("sp_ASRIntGetExpressionDefinition" _
					, New SqlParameter("piExprID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))} _
					, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = Session("action")} _
					, prmErrMsg _
					, prmTimestamp)

				If Len(prmErrMsg.Value.ToString()) > 0 Then
					sErrMsg = "'" & Session("utilname").ToString() & "' " & prmErrMsg.Value.ToString()
				Else
					iCount = 0
					For Each objRow As DataRow In rstDefinition.Rows
						Response.Write("<input type='hidden' id=txtDefn_" & objRow("type").ToString() & "_" & iCount & " name=txtDefn_" & objRow("type").ToString() & "_" & iCount & " value=""" & Replace(objRow("definition").ToString(), """", "&quot;") & """>" & vbCrLf)
						iCount += 1
					Next
					Response.Write("<input type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & prmTimestamp.Value.ToString() & ">" & vbCrLf)
				End If
			Catch ex As Exception
				sErrMsg = "'" & Session("utilname").ToString() & "' " & sUtilTypeName & " definition could not be read." & vbCrLf & FormatError(ex.Message)
			End Try	
	
			If Len(sErrMsg) > 0 Then
				Session("confirmtext") = sErrMsg
				Session("confirmtitle") = "OpenHR"
				Session("followpage") = "defsel"
				Session("reaction") = sReaction
				Response.Clear()
				Response.Redirect("DefSel")

			End If
		End If			
	%>
	<input type="hidden" id="txtOriginalAccess" name="txtOriginalAccess" value="RW">
</form>

<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%:session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<%:session("utilid")%>'>
	<input type="hidden" id="txtTableID" name="txtTableID" value='<%:session("utiltableid")%>'>
	<input type="hidden" id="txtAction" name="txtAction" value='<%:session("action")%>'>
	<input type="hidden" id="txtUtilType" name="txtUtilType" value='<%:CInt(Session("utiltype"))%>'>
	<input type="hidden" id="txtLocaleDecimal" name="txtLocaleDecimal" value='<%:session("LocaleDecimalSeparator")%>'>
	<input type="hidden" id="txtExprColourMode" name="txtExprColourMode" value='<%:session("ExprColourMode")%>'>
	<input type="hidden" id="txtExprNodeMode" name="txtExprNodeMode" value='<%:session("ExprNodeMode")%>'>

	<%
		Dim sErrorDescription As String
				
		Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
	
		Dim sTableName = objDatabase.GetTableName(CInt(Session("utiltableid")))	
		Response.Write("<input type='hidden' id='txtTableName' name='txtTableName' value=""" & sTableName & """>" & vbCrLf)
	
	%>
	<input type="hidden" id="txtCanDelete" name="txtCanDelete" value="0">
	<input type="hidden" id="txtCanInsert" name="txtCanInsert" value="0">
	<input type="hidden" id="txtCanCut" name="txtCanCut" value="0">
	<input type="hidden" id="txtCanCopy" name="txtCanCopy" value="0">
	<input type="hidden" id="txtCanPaste" name="txtCanPaste" value="0">
	<input type="hidden" id="txtCanMoveUp" name="txtCanMoveUp" value="0">
	<input type="hidden" id="txtCanMoveDown" name="txtCanMoveDown" value="0">
	<input type="hidden" id="txtUndoType" name="txtUndoType" value="">
	<input type="hidden" id="txtCutCopyType" name="txtCutCopyType" value="">
	<input type="hidden" id="txtOldText" name="txtOldText" value="">
</form>


<form id="frmSend" method="POST" name="frmSend" style="visibility: hidden; display: none">
	<input type="hidden" id="txtSend_ID" name="txtSend_ID">
	<input type="hidden" id="txtSend_type" name="txtSend_type">
	<input type="hidden" id="txtSend_name" name="txtSend_name">
	<input type="hidden" id="txtSend_description" name="txtSend_description">
	<input type="hidden" id="txtSend_access" name="txtSend_access">
	<input type="hidden" id="txtSend_userName" name="txtSend_userName">
	<input type="hidden" id="txtSend_components1" name="txtSend_components1">
	<input type="hidden" id="txtSend_reaction" name="txtSend_reaction">
	<input type="hidden" id="txtSend_tableID" name="txtSend_tableID" value='<%:session("utiltableid")%>'>
	<input type="hidden" id="txtSend_names" name="txtSend_names" value="">
	<%=Html.AntiForgeryToken()%>
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<form id="frmShortcutKeys" name="frmShortcutKeys" style="visibility: hidden; display: none">
	<%

		Dim sShortcutKeys As String = ""

		Try
			Dim rstShortcutKeys = objDataAccess.GetFromSP("spASRIntGetOpFuncShortcuts")

			iCount = 0
			
			For Each objRow As DataRow In rstShortcutKeys.Rows

				sShortcutKeys = sShortcutKeys & objRow("shortcutKeys").ToString()
				Response.Write("<input type='hidden' id=txtShortcutKeys_" & iCount & " name=txtShortcutKeys_" & iCount & " value=""" & Replace(objRow("shortcutKeys").ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtShortcutType_" & iCount & " name=txtShortcutType_" & iCount & " value=""" & Replace(objRow("componentType").ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtShortcutID_" & iCount & " name=txtShortcutID_" & iCount & " value=""" & Replace(objRow("ID").ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtShortcutParams_" & iCount & " name=txtShortcutParams_" & iCount & " value=""" & Replace(objRow("params").ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtShortcutName_" & iCount & " name=txtShortcutName_" & iCount & " value=""" & Replace(objRow("name").ToString(), """", "&quot;") & """>" & vbCrLf)

				iCount += 1
			Next

			Response.Write("<input type='hidden' id=txtShortcutKeys name=txtShortcutKeys value=""" & Replace(sShortcutKeys, """", "&quot;") & """>" & vbCrLf)

		Catch ex As Exception
			' sErrMsg = "'" & Session("utilname").ToString() & "' " & sUtilTypeName & " definition could not be read." & vbCrLf & FormatError(ex.Message)

		End Try

	
	%>
</form>
	
</div>

<script type="text/javascript">	
	util_def_expression_onload();
</script>
