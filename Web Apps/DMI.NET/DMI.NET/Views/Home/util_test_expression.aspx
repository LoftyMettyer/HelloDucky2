<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage(of DMI.NET.Models.ObjectRequests.TestPromptedValuesModel)" %>

<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="HR.Intranet.Server.Expressions" %>

<script type="text/javascript">
	
	function util_test_expression_close() {

		// Close the popup
		if ($('#divValidateExpression').dialog('isOpen') == true) {
			$('#divValidateExpression').dialog('close');
		}

	}

	function util_test_expression_onload() {

		if ($('#util_test_expression #txtDisplay').val() != "False") {
			// Hide the 'please wait' message.
			$('#PleaseWaitDiv').hide();
			var dialogWidth = screen.width / 3;
			$('#divValidateExpression').dialog("option", "width", dialogWidth);

		}
		else {
			util_test_expression_close();
		}
	}

</script>


<div id="bdyMain">
	<div data-framesource="util_test_expression">

		<div id="PleaseWaitDiv">
			<%
				If Model.UtilType = UtilityType.utlFilter Then
					Response.Write("<h3>Testing Filter</h3>")
				ElseIf Model.UtilType = UtilityType.utlCalculation Then
					Response.Write("<h3>Testing Calculation</h3>")
				Else
					Response.Write("<h3>Testing Expression</h3>")
				End If
			%>						
			Please wait...
			<br />
			<br />
			<input id="Cancel" name="Cancel" class="btn" type="button" value="OK" style="width: 80px; float: right;" onclick="util_test_expression_close();" />
		</div>

		<%
			Dim fOK As Boolean
			Dim fDisplay As Boolean
			Dim sUtilType As String
			Dim objExpression As Expression
			Dim iExprType As Integer
			Dim iReturnType As Integer
			Dim iValidityCode As Integer
			Dim sValidityMessage As String
			Dim sFilterCode As String
			Dim sMsg1 As String
			Dim sMsg As String

			Dim objSessionInfo As SessionInfo = CType(Session("SessionContext"), SessionInfo)
	
			fOK = True
			fDisplay = False
	
			If Model.UtilType = UtilityType.utlFilter Then
				sUtilType = "Filter"
			Else
				sUtilType = "Calculation"
			End If
			
			Dim aPrompts = Session("TestPrompts")
		
			' Get the server DLL to test the expression definition
			objExpression = New Expression(objSessionInfo)
	
			If fOK Then
				If Model.UtilType = UtilityType.utlFilter Then
					iExprType = 11
					iReturnType = 3
				Else
					iExprType = 10
					iReturnType = 0
				End If
				
				fOK = objExpression.Initialise(Model.TableID, 0, CShort(iExprType), CType(iReturnType, ExpressionValueTypes))
			End If

			If fOK Then
				fOK = objExpression.SetExpressionDefinition(HttpUtility.HtmlDecode(Model.components1), "", "", "", "", "")
			End If

			If fOK Then
				iValidityCode = objExpression.ValidateExpression

				If iValidityCode > 0 Then
					fDisplay = True
					Response.Write("<h3>Error Testing " & sUtilType & "</h3>" & vbCrLf)

					sValidityMessage = objExpression.ValidityMessage(CType(iValidityCode, ExprValidationCodes))
					sValidityMessage = Replace(sValidityMessage, vbCr, "<BR>")
					Response.Write(sValidityMessage & vbCrLf)
		%>
		<br />
		<br />
		<input id="Button1" name="Cancel" type="button" class="btn" value="OK" style="width: 80px; float: right;" onclick="util_test_expression_close();" />
		<%
		End If
	End If

	If fOK And (fDisplay = False) Then
		objExpression.SetPromptedValues(aPrompts)

		sFilterCode = objExpression.RuntimeFilterCode

		' Create dynamic User defined functions
		objExpression.UDFFunctions(True)
		
		Dim iRecCount As Integer
		
		fDisplay = True
		If Len(sFilterCode) = 0 Then
			sMsg1 = "Testing " & sUtilType
			sMsg = "Your " & sUtilType.ToLower & " is defined correctly."
		Else
			iRecCount = objExpression.TestFilterCode(CStr(sFilterCode))
			
			If iRecCount < 0 Then
				sMsg1 = "Error Testing " & sUtilType
				sMsg = "Error running the test " & sUtilType.ToLower & " SQL code."
			Else
				sMsg1 = "Testing " & sUtilType
				sMsg = "Your " & sUtilType.ToLower & " is defined correctly.<BR><BR>" & _
							"You have permission to view " & iRecCount & " record"
					
				If (iRecCount <> 1) Then
					sMsg = sMsg & "s"
				End If
				sMsg = sMsg & " using this filter." & vbCrLf
			End If
		End If

		' Remove dynamic User defined functions
		objExpression.UDFFunctions(False)
				
		Response.Write("<h3>" & sMsg1 & "</h3>" & vbCrLf)
		Response.Write(sMsg & vbCrLf)
		%>
		<br />
		<br />
		<input id="Button2" name="Cancel" type="button" class="btn" value="OK" style="width:80px; float: right;" onclick="util_test_expression_close();" />
		<%
		End If		
	
		Response.Write("<input type=hidden id=txtDisplay name=txtDisplay value=" & fDisplay & ">" & vbCrLf)
		%>
	</div>
</div>

<script type="text/javascript">
	util_test_expression_onload();
</script>
