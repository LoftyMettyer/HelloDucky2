﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_picklists")%>" type="text/javascript"></script>

<form id="frmDefinition">
	<div class="absolutefull">

		<div style="display: block;">
			<div class="formField floatleft formInput">
				<label>Name :</label>
				<input id="txtName" name="txtName" class="text" maxlength="50" onkeyup="changeName()">
			</div>
			<div class="formField floatright">
				<label>Owner :</label>
				<input id="txtOwner" name="txtOwner" class="text textdisabled" disabled="disabled" tabindex="-1">
			</div>

			<div class="formTextArea clearboth floatleft">
				<label>Description :</label>
				<textarea id="txtDescription" name="txtDescription" class="textarea" wrap="VIRTUAL" maxlength="255"	onkeyup="changeDescription()"></textarea>
			</div>

			<div class="formOptionGroup floatright">
				<label>Access :</label>
				<div>
					<label>
						<input id="optAccessRW" name="optAccess" type="radio" onclick="changeAccess()" checked />
						Read/Write</label>
					<label>
						<input id="optAccessRO" name="optAccess" type="radio" onclick="changeAccess()" />
						Read Only</label>
					<label>
						<input id="optAccessHD" name="optAccess" type="radio" onclick="changeAccess()" />
						Hidden</label>
				</div>
			</div>
		</div>

		<div class="clearboth"><hr /></div>

		<div class="gridwithbuttons clearboth">

			<div class="stretchyfill">
				<%																																	
					' Get the employee find columns.
					Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

					Dim sErrorDescription As String

					Try
						Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prm1000SepCols = New SqlParameter("ps1000SeparatorCols", SqlDbType.VarChar, 8000) With {.Direction = ParameterDirection.Output}

						Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetDefaultOrderColumns" _
								, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utiltableid"))} _
								, prmErrMsg, prm1000SepCols)
																		
						If Len(prmErrMsg.Value) > 0 Then
							Session("ErrorTitle") = "Picklist Definition Page"
							Session("ErrorText") = prmErrMsg.Value
							Response.Clear()
			
							Response.Redirect("FormError")
			
						Else
							Response.Write("<INPUT type='hidden' id=txt1000SepCols name=txt1000SepCols value=""" & prm1000SepCols.Value & """>" & vbCrLf)
						End If

					Catch ex As Exception
						sErrorDescription = "The find columns could not be retrieved." & vbCrLf & FormatError(ex.Message)

					End Try

				%>
				<div id="PickListGrid" style="height: 400px;">
					<table id="ssOleDBGrid"></table>
				</div>
			</div>

			<div class="stretchyfixed">
				<input type="button" id="cmdAdd" name="cmdAdd" class="btn" value="Add" onclick="addClick()" />
				<br />
				<input type="button" id="cmdAddAll" name="cmdAddAll" class="btn" value="Add All" onclick="addAllClick()" />
				<br />
				<input type="button" id="cmdFilteredAdd" name="cmdFilteredAdd" class="btn" value="Filtered Add" onclick="filteredAddClick()" />
				<br />
				<input type="button" id="cmdRemove" name="cmdRemove" class="btn" value="Remove" onclick="removeClick()" />
				<br />
				<input type="button" id="cmdRemoveAll" name="cmdRemoveAll" class="btn" value="Remove All" onclick="removeAllClick()" />
				<br />
				<input type="button" id="cmdOK" name="cmdOK" class="btn" value="OK" onclick="okClick()" />
				<br />
				<input type="button" id="cmdCancel" name="cmdCancel" class="btn" value="Cancel" onclick="cancelClick()" />
			</div>

		</div>

		<div id="RecordCountDIV" style="margin-top: 40px; position: relative;"></div>


		</div>
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Dim sErrMsg As String
		Dim sSelectedRecords As String
		
		sErrMsg = ""

		If Session("action") <> "new" Then
			
			Try

				Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmName = New SqlParameter("psPicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmOwner = New SqlParameter("psPicklistOwner", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmDescription = New SqlParameter("psPicklistDesc", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmAccess = New SqlParameter("psAccess", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmTimestamp = New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				Dim rstDefinition = objDataAccess.GetFromSP("sp_ASRIntGetPicklistDefinition" _
					, New SqlParameter("piPicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))} _
					, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Session("action")} _
					, prmErrMsg, prmName, prmOwner, prmDescription, prmAccess, prmTimestamp)
		
			
				sSelectedRecords = "0"
				Response.Write("<input type='hidden' id='txtSelectedRecords' name='txtSelectedRecords' value='" & sSelectedRecords & "'>" & vbCrLf)
				
				If Len(prmErrMsg.Value) > 0 Then
					sErrMsg = "'" & Session("utilname") & "' " & prmErrMsg.Value
				Else
				
					'Response.Write("<input type='hidden' id='txtDefn_Name' name='txtDefn_Name' value='" & Replace(prmName.Value.ToString(), """", "&quot;") & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(prmName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Owner' name='txtDefn_Owner' value=""" & Replace(prmOwner.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Description' name='txtDefn_Description' value=""" & Replace(prmDescription.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Access' name='txtDefn_Access' value='" & prmAccess.Value & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Timestamp' name='txtDefn_Timestamp' value='" & prmTimestamp.Value & "'>" & vbCrLf)
				End If

			Catch ex As Exception
				sErrMsg = "'" & Session("utilname") & "' picklist definition could not be read." & vbCrLf & FormatError(ex.Message)

				
			End Try

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
	<input type="hidden" id="txtTableID" name="txtTableID" value='<% =session("utiltableid")%>'>
	<input type="hidden" id="selectedIDs1" name="selectedIDs1">
</form>

<script type="text/javascript">
	util_def_picklist_onload();
</script>


