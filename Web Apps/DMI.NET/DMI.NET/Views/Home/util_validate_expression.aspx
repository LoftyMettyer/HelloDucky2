<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage(of DMI.NET.Models.ObjectRequests.ValidateExpressionModel)" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %> 
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="DMI.NET.Helpers" %>
<%@ Import Namespace="HR.Intranet.Server.Expressions" %>

	<script type="text/javascript">

		function util_validate_window_onload() {

			if ($('#util_validate_expression #txtDisplay').val() != "False") {
				$('#PleaseWaitDiv').hide();
				var dialogWidth = screen.width / 3;
				$('#divValidateExpression').dialog("option", "width", dialogWidth);
			}
			else {
				nextPass();
			}
		}
		
		function overwrite(){
			nextPass();
		}

		function removeComponents(piIndex) {

			var sKeys;

			if (piIndex == 1) {
				sKeys = $('#util_validate_expression #txtDeletedKeys').val();
			}
			else {
				if (piIndex == 2) {
					sKeys = $('#util_validate_expression #txtHiddenNotOwnerKeys').val();
				}
				else {
					sKeys = $('#util_validate_expression #txtHiddenOwnerKeys').val();
				}
			}
			tree_NodesRemove(sKeys);
			uve_cancelClick();
		}

		function returnToDefSel() {
			OpenHR.returnToDefSel();		  
			uve_cancelClick();
		}

		function makeHidden() {
			ude_makeHidden();
		}

		function nextPass() {
			
			var iNextPass = parseInt($("#validatePass").val());
			iNextPass += 1;

			if (iNextPass <= 3) {
				$("#validatePass").val(iNextPass);

				var postData = {
					validatePass: iNextPass,
					validateUtilID: <%:Model.validateUtilID%>,
					validateUtilType: <%:CInt(Model.validateUtilType)%>,
					validateAccess: '<%:Model.validateAccess%>',
					validateOriginalAccess: '<%:Model.validateOriginalAccess%>',
					validateOwner: '<%:Model.validateOwner%>',
					components1: '<%:Model.components1%>',
				  validateBaseTableID: <%:Model.validateBaseTableID%>, 
				  validateReturnType: $("#txtExpressionReturnType").val(), 
				  validateExpressionType: $("#txtExpressionType").val(),
					<%:Html.AntiForgeryTokenForAjaxPost() %>
				}

				OpenHR.submitForm(null, 'divValidateExpression', null, postData, "util_validate_expression");
			}
			else {

				var frmSend = OpenHR.getForm("divDefExpression", "frmSend");	
				frmSend.txtSend_ReturnType.value = <%:CInt(Model.validateReturnType)%>; 
				frmSend.txtSend_ExpressionType.value = <%:CInt(Model.validateExpressionType)%>; 
				
				// Check if the user has loaded tools screen (picklist/filter/cal) from the report definition.
				var displayDiv = (IsToolsScreenLoadedFromReportDefinition() == true ? "ToolsFrame" : "workframe");

				OpenHR.submitForm(frmSend, displayDiv, null, null, "util_def_expression_Submit", uve_cancelClick);
			}
		}

		function uve_cancelClick() {			
			var iIndex;
			var sCurrentPage = OpenHR.currentWorkPage();

			try {
				sCurrentPage = sCurrentPage.toString();
				iIndex = sCurrentPage.lastIndexOf("/");

				if (iIndex >= 0) {
					sCurrentPage = sCurrentPage.substr(iIndex + 1);
				}
	
				iIndex = sCurrentPage.indexOf(".");
				if (iIndex >= 0) {
					sCurrentPage = sCurrentPage.substr(0, iIndex);
				}

				try {
					sCurrentPage = sCurrentPage.toUpperCase();
				} catch(e) {}
					
				if (sCurrentPage == "UTIL_DEF_EXPRESSION") {
					OpenHR.reEnableControls();
				}
			}
			catch(e) {
			}

			// Close the popup
			if ($('#divValidateExpression').dialog('isOpen') == true) {
				$('#divValidateExpression').dialog('close');
				$('#divValidateExpression').html();
			}

		}
	</script>

<div id="util_validate_expression" data-framesource="util_validate_expression">
	<div id="PleaseWaitDiv">
		<h3>
			<%
			  
				If Model.validateUtilType = UtilityType.utlFilter Then
					Response.Write("Validating Filter")
				ElseIf Model.validateUtilType = UtilityType.utlCalculation Then
					Response.Write("Validating Calculation")
				Else
					Response.Write("Validating Expression")
				End If
			%>				
		</h3>
		Please wait...
		<br />
		<br />
		<input type="button" value="Cancel" class="btn" name="Cancel" style="float: right; width: 80px" id="Cancel" onclick="uve_cancelClick()" />
	</div>
<%
  Dim fOK As Boolean = True
  Dim fDisplay As Boolean = False
  Dim sUtilType As String
  Dim iExprType As ExpressionTypes

  Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
  Dim objSessionInfo As SessionInfo = CType(Session("SessionContext"), SessionInfo)
  Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

  Dim iErrorCode As Integer
  Dim sDeletedKeys As String
  Dim sHiddenOwnerKeys As String
  Dim sHiddenNotOwnerKeys As String
  Dim sDeletedDescs As String
  Dim sHiddenOwnerDescs As String
  Dim sHiddenNotOwnerDescs As String

  Dim objExpression As Expression
  Dim iReturnType As ExpressionValueTypes

  Dim iValidityCode As ExprValidationCodes
  Dim sValidityMessage As String
  Dim iOriginalReturnType As Integer
  Dim sDescription As String

  Dim sHiddenErrorMsg As String

  If Model.validateUtilType = UtilityType.utlFilter Then
    sUtilType = "Filter"
    iExprType = ExpressionTypes.giEXPR_RUNTIMEFILTER
    iReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC
  Else
    sUtilType = "Calculation"
    iExprType = CType(iif(Model.validateBaseTableID = 0, ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC, ExpressionTypes.giEXPR_RUNTIMECALCULATION), ExpressionTypes)
    iReturnType = ExpressionValueTypes.giEXPRVALUE_UNDEFINED
  End If
 
  If Model.validatePass = 1 Then

    Dim prmDeletedKeys = New SqlParameter("psDeletedKeys", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
    Dim prmHiddenOwnerKeys = New SqlParameter("psHiddenOwnerKeys", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
    Dim prmHiddenNotOwnerKeys = New SqlParameter("psHiddenNotOwnerKeys", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
    Dim prmDeletedDescs = New SqlParameter("psDeletedDescs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
    Dim prmHiddenOwnerDescs = New SqlParameter("psHiddenOwnerDescs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
    Dim prmHiddenNotOwnerDescs = New SqlParameter("psHiddenNotOwnerDescs", SqlDbType.VarChar - 1) With {.Direction = ParameterDirection.Output}
    Dim prmErrorCode = New SqlParameter("piErrorCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

    Try

      objDataAccess.ExecuteSP("sp_ASRIntValidateExpression" _
        , New SqlParameter("psUtilName", SqlDbType.VarChar, 255) With {.Value = Model.validateName} _
        , New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = Model.validateUtilID} _
        , New SqlParameter("piUtilType", SqlDbType.Int) With {.Value = iExprType} _
        , New SqlParameter("psUtilOwner", SqlDbType.VarChar, 128) With {.Value = Model.validateOwner} _
        , New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = Model.validateBaseTableID} _
        , New SqlParameter("psComponentDefn", SqlDbType.VarChar, -1) With {.Value = HttpUtility.HtmlDecode(Model.components1)} _
        , New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = Model.validateTimestamp} _
        , prmDeletedKeys, prmHiddenOwnerKeys, prmHiddenNotOwnerKeys, prmDeletedDescs _
        , prmHiddenOwnerDescs, prmHiddenNotOwnerDescs, prmErrorCode)

    Catch ex As Exception
      Throw ex
    End Try

    Response.Write("<input type='hidden' id='txtErrorCode' name='txtErrorCode' value='" & prmErrorCode.Value & "'>" & vbCrLf)
    Response.Write("<input type='hidden' id='txtDeletedKeys' name='txtDeletedKeys' value='" & prmDeletedKeys.Value & "'>" & vbCrLf)
    Response.Write("<input type='hidden' id='txtHiddenOwnerKeys' name='txtHiddenOwnerKeys' value='" & prmHiddenOwnerKeys.Value & "'>" & vbCrLf)
    Response.Write("<input type='hidden' id='txtHiddenNotOwnerKeys' name='txtHiddenNotOwnerKeys' value='" & prmHiddenNotOwnerKeys.Value & "'>" & vbCrLf)

    iErrorCode = CInt(prmErrorCode.Value)
    sDeletedKeys = prmDeletedKeys.Value.ToString()
    sHiddenOwnerKeys = prmHiddenOwnerKeys.Value.ToString()
    sHiddenNotOwnerKeys = prmHiddenNotOwnerKeys.Value.ToString()
    sDeletedDescs = prmDeletedDescs.Value.ToString()
    sHiddenOwnerDescs = prmHiddenOwnerDescs.Value.ToString()
    sHiddenNotOwnerDescs = prmHiddenNotOwnerDescs.Value.ToString()

    If (iErrorCode = 1) Or (iErrorCode = 2) Then
      ' 1 = Expression deleted by another user. Save as new ? 
      ' 2 = Made hidden/read-only by another user. Save as new ? 
      fDisplay = True

      Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
      If (iErrorCode = 1) Then
        Response.Write("The " & sUtilType.ToLower & " has been deleted by another user. Save as a new definition ?" & vbCrLf)
      Else
        Response.Write("The " & sUtilType.ToLower & " has been amended by another user and is now Read Only. Save as a new definition ?" & vbCrLf)
      End If
      Response.Write("<br/><br/>" & vbCrLf)

      Response.Write("<input type='button' value='No' class='btn' name='btnNo' style='float: right; width: 80px;' id='btnNo' OnClick=""uve_cancelClick()""/>" & vbCrLf)
      Response.Write("<input type='button' value='Yes' class='btn' name='btnYes' style='float: right; width: 80px; margin-right: 10px;' id='btnYes' OnClick=""ude_createNew()""/>" & vbCrLf)

    End If

    If iErrorCode = 4 Then
      ' 4 = Non-unique name. Save fails */
      fDisplay = True

      Response.Write("<div>")
      Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
      Response.Write("A " & sUtilType.ToLower & " called '" & Model.validateName & "' already exists.<br/><br/>")
      Response.Write("<input type='button' value='Close' class='btn' name='Cancel' style='float: right; width: 80px' id='Cancel' OnClick='uve_cancelClick()'/>" & vbCrLf)
      Response.Write("</div>")
    End If

    If (iErrorCode = 0) Or (iErrorCode = 3) Then
      ' 0 = No error (but must check the strings of keys)
      If Len(sDeletedKeys) > 0 Then
        fDisplay = True

        Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
        Response.Write("1The following calculations and filters have been deleted, and will be removed from the " & sUtilType.ToLower & " definition." & vbCrLf)

        Response.Write("<ul>" & vbCrLf)
        Dim arrDeletedDescs = sDeletedDescs.Split(CChar("	"))
        For Each sDeletedDesc As String In arrDeletedDescs
          Response.Write("<li>" & sDeletedDesc & "</li>" & vbCrLf)
        Next

        Response.Write("</ul>" & vbCrLf)
        Response.Write("<br/><br/>" & vbCrLf)
        Response.Write("<input type='button' value='Close' class='btn' name='RemoveComponents' style='float: right; width: 80px' id='RemoveComponents' OnClick=""removeComponents(1)""/>" & vbCrLf)
      Else
        If (Len(sHiddenOwnerKeys) > 0) Or (Len(sHiddenNotOwnerKeys) > 0) Then

          If (UCase(Model.validateOwner) = UCase(Session("Username"))) Then
            ' Current user IS the owner of the filter/calc.
            If (Len(sHiddenNotOwnerKeys) > 0) Then
              ' There are hidden components in the expression that are NOT owned by the current user.
              ' Need to remove the hidden components.	
              fDisplay = True

              Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
              Response.Write("2The following calculations and filters have been made hidden, and will be removed from the " & sUtilType.ToLower & " definition." & vbCrLf)

              Response.Write("<ul>" & vbCrLf)
              Dim arrHiddenNotOwners = sHiddenNotOwnerDescs.Split(CChar("	"))
              For Each sHiddenNotOwnerDesc As String In arrHiddenNotOwners
                Response.Write("<li>" & sHiddenNotOwnerDesc & "</li>" & vbCrLf)
              Next
              Response.Write("</ul>" & vbCrLf)
              Response.Write("<br/><br/>" & vbCrLf)
              Response.Write("<input type='button' value='Close' class='btn' name='RemoveComponents' style='float: right; width: 80px' id='RemoveComponents' OnClick=""removeComponents(2)""/>" & vbCrLf)
            Else
              ' There are hidden components in the expression that ARE owned by the current user.
              ' Need to make the expression hidden too.
              If (Model.validateAccess <> "HD") Then
                fDisplay = True

                Response.Write("<div>")
                Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
                Response.Write("The following calculations and filters have been made hidden. <br/>The " & sUtilType.ToLower & " will now be made hidden.<br/><br/>")


                Response.Write("<ul>" & vbCrLf)
                Dim arrHiddenOwners = sHiddenOwnerDescs.Split(CChar("	"))
                For Each sHiddenOwnerDesc As String In arrHiddenOwners
                  Response.Write("<li>" & sHiddenOwnerDesc & "</li>" & vbCrLf)
                Next
                Response.Write("</ul>" & vbCrLf)
                Response.Write("<br/><br/>")
                Response.Write("<input type='button' value='Close' class='btn' name='makeHidden' style='float: right; width: 80px' id='makeHidden' OnClick='makeHidden()'/>" & vbCrLf)
                Response.Write("</div>")

              End If
            End If
          Else
            ' Current user is NOT the owner of the filter/calc.
            fDisplay = True

            If (Len(sHiddenNotOwnerKeys) > 0) Then
              ' There are hidden components in the expression that are NOT owned by the current user.
              ' Cannot edit the expression as it too must now be hidden.
              Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
              Response.Write("The following calculations and filters have been made hidden by another user. Cannot make any modifications to this " & sUtilType.ToLower & " definition." & vbCrLf)

              Response.Write("<ul>" & vbCrLf)
              Dim arrHiddenNotOwners = sHiddenNotOwnerDescs.Split(CChar("	"))
              For Each sHiddenNotOwnerDesc As String In arrHiddenNotOwners
                Response.Write("<li>" & sHiddenNotOwnerDesc & "</li>" & vbCrLf)
              Next
              Response.Write("</ul>" & vbCrLf)
              Response.Write("<br/><br/>" & vbCrLf)
              Response.Write("<input type='button' value='Close' class='btn' name='ReturnToDefSel' style='float: right; width: 80px' id='ReturnToDefSel' OnClick=""returnToDefSel()""/>" & vbCrLf)
            Else
              ' There are hidden components in the expression that ARE owned by the current user.
              ' Need to remove the hidden components.	
              Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
              Response.Write("The following calculations and filters have been made hidden, and will be removed from the " & sUtilType.ToLower & " definition." & vbCrLf)
              Response.Write("<ul>" & vbCrLf)
              Dim arrHiddenOwners = sHiddenOwnerDescs.Split(CChar("	"))
              For Each sHiddenOwnerDesc As String In arrHiddenOwners
                Response.Write("<li>" & sHiddenOwnerDesc & "</li>" & vbCrLf)
              Next
              Response.Write("</ul>" & vbCrLf)

              Response.Write("<br/><br/>" & vbCrLf)
              Response.Write("<input type='button' value='Close' class='btn' name='RemoveComponents' style='float: right; width: 80px' id='RemoveComponents' OnClick=""removeComponents(3)""/>" & vbCrLf)
            End If
          End If
        End If
      End If
    End If

    If (iErrorCode = 3) And (fDisplay = False) Then
      ' 3 = Modified by another user (still writable). Overwrite ? 
      fDisplay = True

      Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
      Response.Write("The " & sUtilType.ToLower & " has been amended by another user. Would you like to overwrite this definition ?" & vbCrLf)
      Response.Write("<br/><br/>" & vbCrLf)
      Response.Write("<input type='button' value='No' class='btn' name='btnNo' style='float: right; width: 80px' id='btnNo' OnClick=""uve_cancelClick()""/>" & vbCrLf)
      Response.Write("<input type='button' value='Yes' class='btn' name='btnYes' style='float: right; width: 80px; margin-right: 10px;' id='btnYes' OnClick=""overwrite()""/>" & vbCrLf)
    End If
  End If


  If Model.validatePass = 2 Then
    ' Get the server DLL to validate the expression definition
    objExpression = New Expression(objSessionInfo)

    fOK = objExpression.Initialise(Model.validateBaseTableID, Model.validateUtilID, iExprType, iReturnType)

    If fOK Then
      fOK = objExpression.SetExpressionDefinition(HttpUtility.HtmlDecode(Model.components1), "", "", "", "", "")
    End If

    If fOK Then
      iValidityCode = objExpression.ValidateExpression
      If iValidityCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
        iReturnType = objExpression.ReturnType
      Else
        fDisplay = True
        Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
        sValidityMessage = objExpression.ValidityMessage(iValidityCode)
        sValidityMessage = Replace(sValidityMessage, vbCr, "<BR>")
        Response.Write(sValidityMessage & vbCrLf)
        Response.Write("<br/><br/>" & vbCrLf)
        Response.Write("<input type='button' value='Close' class='btn' name='Cancel' style='float: right; width: 80px' id='Cancel' OnClick=""uve_cancelClick()""/>" & vbCrLf)
      End If
    End If

    Model.validateExpressionType = iExprType
    Model.validateReturnType = iReturnType

    objExpression = Nothing

    If Not fDisplay Then
      ' Check if the expression return type has changed. 
      ' If so, check if it can be.

      If Model.validateUtilID > 0 Then
        objExpression = New Expression(objSessionInfo)

        iOriginalReturnType = objExpression.ExistingExpressionReturnType(Model.validateUtilID)
        objExpression = Nothing

        If iReturnType <> iOriginalReturnType Then

          Dim rsUsage = objDatabase.GetUtilityUsage(Model.validateUtilType, Model.validateUtilID)

          If rsUsage.Rows.Count > 0 Then
            fDisplay = True
            Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
            Response.Write("The return type cannot be changed, as the " & sUtilType.ToLower & " is currently being used in the following definitions:" & vbCrLf)
            Response.Write("<br/><ul>" & vbCrLf)
            For Each objRow As DataRow In rsUsage.Rows
              sDescription = objRow("description").ToString()
              sDescription = Replace(sDescription, "<", "&lt;")
              sDescription = Replace(sDescription, ">", "&gt;")

              Response.Write("<li>" & sDescription & "</li>" & vbCrLf)
            Next

            Response.Write("</ul>" & vbCrLf)

            Response.Write("</br><br/>" & vbCrLf)
            Response.Write("<input type='button' value='Close' class='btn' name='Cancel' style='float: right;width: 80px' id='Cancel' OnClick=""uve_cancelClick()""/>" & vbCrLf)
          End If

        End If
      End If
    End If
  End If

  If Model.validatePass = 3 Then
    If (Model.validateUtilID > 0) And _
        (UCase(Model.validateOwner) = UCase(Session("username"))) And _
        (Model.validateAccess = "HD") And _
        (Model.validateOriginalAccess <> "HD") Then
      ' Check if the expression can be made hidden.

      Dim prmResult = New SqlParameter("piResult", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
      Dim prmMsg = New SqlParameter("psMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

      objDataAccess.ExecuteSP("sp_ASRIntCheckCanMakeHidden" _
              , New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = Model.validateUtilType} _
              , New SqlParameter("piUtilityID", SqlDbType.VarChar, 255) With {.Value = Model.validateUtilID} _
              , prmResult, prmMsg)

      If prmResult.Value = 1 Then
        ' calc/filter used only in utilities owned by the current user - we then need to prompt the user if they want to make these utilities hidden too.
        fDisplay = True
        sHiddenErrorMsg = "Making this " & sUtilType.ToLower & " hidden will automatically make the following definition(s), of which you are the owner, hidden also :" & _
          "<BR><BR>" & _
          prmMsg.Value.ToString() & _
          "<BR><BR>" & _
          "Do you wish to continue ?"

        Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
        Response.Write(sHiddenErrorMsg & vbCrLf)
        Response.Write("<br/><br/>" & vbCrLf)
        Response.Write("<input type='button' value='No' class='btn' name='btnNo' style='float: right; width: 80px' id='btnNo' OnClick=""uve_cancelClick()""/>" & vbCrLf)
        Response.Write("<input type='button' value='Yes' class='btn' name='btnYes' style='float: right; width: 80px; margin-right: 10px;' id='btnYes' OnClick=""overwrite()""/>" & vbCrLf)
      End If

      If (prmResult.Value = 2) Or _
        (prmResult.Value = 3) Or _
        (prmResult.Value = 4) Then
        ' calc/filter used in utilities which are in batch jobs not owned by the current user - Cannot therefore make the calc/filter hidden.
        If (prmResult.Value = 2) Then
          sHiddenErrorMsg = "This " & sUtilType.ToLower & " cannot be made hidden as it is used in definition(s) which are included in the following batch jobs of which you are not the owner :" & _
            "<BR><BR>" & _
            prmMsg.Value.ToString()
        Else
          If (prmResult.Value = 3) Then
            sHiddenErrorMsg = "This " & sUtilType.ToLower & " cannot be made hidden as it is used in definition(s), of which you are not the owner :" & _
              "<BR><BR>" & _
              prmMsg.Value.ToString()
          Else
            sHiddenErrorMsg = "This " & sUtilType.ToLower & " cannot be made hidden as it is used in definition(s) which are included in the following batch jobs which are scheduled to be run by other user groups :" & _
              "<BR><BR>" & _
              prmMsg.Value.ToString()
          End If
        End If
        fDisplay = True

        Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
        Response.Write(sHiddenErrorMsg & vbCrLf)
        Response.Write("<br/><br/>" & vbCrLf)
        Response.Write("<input type='button' value='Close' class='btn' name='Cancel' style='float: right; width: 80px' id='Cancel' OnClick=""uve_cancelClick()""/>" & vbCrLf)
      End If
    End If
  End If

  Response.Write("<input type='hidden' id='txtDisplay' name='txtDisplay' value='" & fDisplay & "'>" & vbCrLf)

%>

  <input type='hidden' id='txtExpressionReturnType' name='txtExpressionReturnType' value='<%:CInt(Model.validateReturnType)%>'>
  <input type='hidden' id='txtExpressionType' name='txtExpressionType' value='<%:CInt(Model.validateExpressionType)%>'>
	<input type="hidden" id="validatePass" name="validatePass" value='<%:Model.validatePass%>'>
</div>

<script type="text/javascript">
	util_validate_window_onload();
</script>
