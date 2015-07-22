<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_picklists")%>" type="text/javascript"></script>

<form id="frmDefinition" name="frmDefinition">
	<div class="absolutefull" style="padding-top: 15px">
		<div class="nowrap">
			<div class="tablerow">
				<label>Name :</label>
				<input id="txtName" name="txtName" maxlength="50" onkeyup="changeName()" style="width: 90%;">
				<label>Owner :</label>
				<input id="txtOwner" style="margin-left: 4px; width: 90%;" name="txtOwner" disabled="disabled" tabindex="-1">
			</div>
			<br />
			<div class="tablerow">
				<label>Description :</label>
				<textarea id="txtDescription" name="txtDescription" style="width: 90%; height: 60px; white-space: normal !important;" maxlength="255" onkeyup="changeDescription()"></textarea>
				<label>Access :</label>
				<div>
					<input class="inline-block" id="optAccessRW" name="optAccess" type="radio" onclick="changeAccess()" checked />
					<label class="inline-block" for="optAccessRW">Read/Write</label><br />
					<input class="inline-block" id="optAccessRO" name="optAccess" type="radio" onclick="changeAccess()" />
					<label class="inline-block" for="optAccessRO">Read Only</label><br />
					<input class="inline-block" id="optAccessHD" name="optAccess" type="radio" onclick="changeAccess()" />
					<label class="inline-block" for="optAccessHD">Hidden</label>
				</div>
			</div>
		</div>

		<div class="clearboth">
			<hr />
		</div>

		<div class="gridwithbuttons clearboth">
			<div id="PickListGrid" style="height: 400px;" class="stretchyfill">
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
							
							Dim iCount As Integer
							Dim sAddString As String
							If Session("action") = "new" Then
								sAddString = ""
								If rstFindRecords.Rows.Count > 0 Then
									For Each objRow As DataRow In rstFindRecords.Rows

										sAddString = ""
									
										For iloop = 0 To (rstFindRecords.Columns.Count - 1)
											If iloop > 0 Then
												sAddString = sAddString & "	"
											End If
										
											If Not IsDBNull(objRow(iloop)) Then
												sAddString = sAddString & Replace(objRow(iloop).ToString(), """", "&quot;")
											End If
										Next
										Response.Write("<INPUT type='hidden' id=txtOptionColDef_" & iCount & " name=txtOptionColDef_" & iCount & " value=""" & sAddString & """>" & vbCrLf)
										
										iCount += 1
									Next
								End If
							End If
							
						End If

					Catch ex As Exception
						sErrorDescription = "The find columns could not be retrieved." & vbCrLf & FormatError(ex.Message)

					End Try
					
				%>
				<table id="ssOleDBGrid"></table>
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

	<div id="RecordCountDIV" style="margin-top: 30px; position: relative;"></div>
	</div>
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
	<input type="hidden" id="txtFlag_To_Identify_Page_Source" name="txtFlag_To_Identify_Page_Source" value='<%: Session("IsLoadedFromReportDefinition")%>'>
	<%
		Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
	%>
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
	<input type="hidden" id="txtSend_tableID" name="txtSend_tableID" value='<%: Session("utiltableid")%>'>
	<%=Html.AntiForgeryToken()%>
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">


<script type="text/javascript">
	util_def_picklist_onload();
	BindDefaultGridOnNewDefinition();
</script>


