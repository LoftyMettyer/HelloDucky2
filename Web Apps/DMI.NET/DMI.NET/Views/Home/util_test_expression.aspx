<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="HR.Intranet.Server.Expressions" %>

<%
	' Write the prompted values from the calling form into a session variable.
	Dim j As Integer
	Dim aPrompts(1, 0) As String
	Dim sKey As String
	j = 0
	ReDim Preserve aPrompts(1, 0)
	For i = 0 To (Request.Form.Count) - 1
		sKey = Request.Form.Keys(i)
		If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
				(UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
			ReDim Preserve aPrompts(1, j)
			If (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
				aPrompts(0, j) = "prompt_3_" & Mid(sKey, 11)
				aPrompts(1, j) = UCase(Request.Form.Item(i))
			Else
				aPrompts(0, j) = sKey
				Select Case Mid(sKey, 8, 1)
					Case "2"
						' Numeric. Replace locale decimal point with '.'
						aPrompts(1, j) = Replace(Request.Form.Item(i), Session("LocaleDecimalSeparator").ToString(), ".")
					Case "4"
						' Date. Reformat to match SQL's mm/dd/yyyy format.
						aPrompts(1, j) = ConvertLocaleDateToSQL(Request.Form.Item(i))
					Case Else
						aPrompts(1, j) = Request.Form.Item(i)
				End Select
			End If
			j = j + 1
		End If
	Next
	Session("TestPrompts") = aPrompts
%>

<script type="text/javascript">
	function util_test_expression_onload() {

		if ($('#util_test_expression #txtDisplay').val() != "False") {
			// Hide the 'please wait' message.
			$('#PleaseWaitDiv').hide();
			var dialogWidth = screen.width / 3;
			$('#tmpDialog').dialog("option", "width", dialogWidth);

		}
		else {
			OpenHR.clearTmpDialog();
		}
	}

</script>


<div id="bdyMain">
	<div data-framesource="util_test_expression">

		<div id="PleaseWaitDiv">
			<%
				If Request.Form("type") = 11 Then
					Response.Write("<h3>Testing Filter</h3>")
				Else
					If Request.Form("type") = 12 Then
						Response.Write("<h3>Testing Calculation</h3>")
					Else
						Response.Write("<h3>Testing Expression</h3>")
					End If
				End If
			%>						
			Loading...
			<br />
			<br />
			<input id="Cancel" name="Cancel" class="btn" type="button" value="OK" style="width: 80px; float: right;" onclick="OpenHR.clearTmpDialog();" />
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
	
			If Request.Form("type") = "11" Then
				sUtilType = "Filter"
			Else
				sUtilType = "Calculation"
			End If
		
			' Get the server DLL to test the expression definition
			objExpression = New Expression(objSessionInfo)
	
			If fOK Then
				If Request.Form("type") = 11 Then
					iExprType = 11
					iReturnType = 3
				Else
					iExprType = 10
					iReturnType = 0
				End If
				
				fOK = objExpression.Initialise(CInt(Request.Form("tableID")), 0, CShort(iExprType), CType(iReturnType, ExpressionValueTypes))
			End If

			If fOK Then
				fOK = objExpression.SetExpressionDefinition(CStr(Request.Form("components1")), "", "", "", "", "")
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
		<input id="Button1" name="Cancel" type="button" class="btn" value="OK" style="width: 80px; float: right;" onclick="OpenHR.clearTmpDialog();" />
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
		<input id="Button2" name="Cancel" type="button" class="btn" value="OK" style="width:80px; float: right;" onclick="OpenHR.clearTmpDialog();" />
		<%
		End If		
	
		Response.Write("<input type=hidden id=txtDisplay name=txtDisplay value=" & fDisplay & ">" & vbCrLf)
		%>
	</div>
</div>

<script type="text/javascript">
	util_test_expression_onload();
</script>
