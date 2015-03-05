<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<script runat="server">
	Private Function GetEmailSelection() As String
		Dim emailSelectionHtmlTable As New StringBuilder 'Used to construct the (temporary) HTML table that will be transformed into a jQuery grid table
		Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)		'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
		
		'Get the records.
		Dim rstDefSelRecords As DataTable = objDataAccess.GetDataTable("spASRIntGetRecordSelection", CommandType.StoredProcedure _
			, New SqlParameter("@psType", SqlDbType.VarChar, 255) With {.Value = "EMAIL"} _
			, New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = 0})

		'Create an HTML table
		With emailSelectionHtmlTable
			.Append("<table id=""EmailSelectionTable"">")
			.Append("<tr>")
			.Append("<th id=""EmailGroupIDHeader"">EmailGroupID</th>")
			.Append("<th id=""NameHeader"">Name</th>")
			.Append("<th id=""FullNameHeader"">FullName</th>")
			.Append("</tr>")
		End With

        'Populate the table
		Dim i As Integer = 1
		For Each r As DataRow In rstDefSelRecords.Rows
			With emailSelectionHtmlTable
				.Append("<tr>")
				.Append("<td id='Row" & i & "'>" & r("ID").ToString & "</td>")
				.Append("<td>" & r("name").ToString.Replace("_", " ").Replace("""", "&quot;") & "</td>")
				'Loop around and add on fullemailaddresses to the grid
				If r("ID") < 1 Then
				Else
					Try
						Dim rstEmailAddr = objDataAccess.GetDataTable("spASRIntGetEmailGroupAddresses", CommandType.StoredProcedure _
						, New SqlParameter("EmailGroupID", SqlDbType.Int) With {.Value = r("ID")})
						Dim x As Integer = 1
						If Not rstEmailAddr Is Nothing Then
							If rstEmailAddr.Rows.Count < 2 Then
								For Each objRow In rstEmailAddr.Rows
									If x > 0 Then
										.Append("<td>" & objRow(0).ToString.Replace("_", " ").Replace("""", "&quot;") & "</td>")
									End If
									x += 1
								Next
							Else
								'Append two email addresses
								Dim buildMultipleEmailString As String = ""
								For Each objRow In rstEmailAddr.Rows
									If buildMultipleEmailString.Length > 0 Then buildMultipleEmailString = buildMultipleEmailString + ";"
									buildMultipleEmailString = buildMultipleEmailString + objRow(0).ToString.Replace("_", " ").Replace("""", "&quot;")
									x += 1
								Next
								If x > 0 Then
									.Append("<td>" & buildMultipleEmailString.ToString.Replace("_", " ").Replace("""", "&quot;") & "</td>")
								End If
							End If
						End If
                            
					Catch ex As Exception
						Dim sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(ex.Message)
					End Try
				End If
				.Append("</tr>")
				i += 1
			End With
		Next

		emailSelectionHtmlTable.Append("</table>")

		Return emailSelectionHtmlTable.ToString
	End Function
</script>

<script type="text/javascript">

	function emailSelection_window_onload() {		
		$("#EventLogEmailSelect .button").button();
		//Hide the EmailGroup table header and its column
		$('.ui-jqgrid-htable tr th:nth-child(1)').hide();
		$('#EmailSelectionTable tr td:nth-child(1)').hide();
	};

	function setForm() {
		var frmPopup = document.getElementById("frmPopup");
		
		if (document.getElementById('txtAbsenceEmailGroup') != null) {
			document.getElementById('txtAbsenceEmailGroup').value = frmPopup.txtSelectedName.value;
			document.getElementById('txtAbsenceEmailGroupID').value = frmPopup.txtSelectedID.value;
		}

		if (document.getElementById('txtEmailGroup') != null) {
			document.getElementById('txtEmailGroup').value = frmPopup.txtSelectedName.value;
			document.getElementById('txtEmailGroupID').value = frmPopup.txtSelectedID.value;
		}

		closeEmailSelect();
		return false;
	};

</script>

<div>
	<form class="displaynone" id="frmPopup" name="frmPopup" onsubmit="return setForm();">
		<input id="txtSelectedID" name="txtSelectedID" type="hidden">
		<input id="txtSelectedName" name="txtSelectedName" type="hidden">
		<input id="txtSelectedAccess" name="txtSelectedAccess" type="hidden">
		<input id="txtSelectedUserName" name="txtSelectedUserName" type="hidden">
	</form>

	<div class="absolutefull">
		<div class="pageTitleDiv padbot15 margeTop10">
			<span class="pageTitle" id="EventLogEmailTitle">Email Selection</span>
		</div>
		<div>
		<%=GetEmailSelection()%>
		<div id="divEmailSelectionButtons">
			<input class="button" id="cmdok" name="cmdok" onclick="emailSelectionEvent()" type="button" value="OK" />
			<input class="button" id="cmdcancel" name="cmdcancel" onclick="closeEmailSelect()" type="button" value="Cancel" />
		</div>
		</div>
	</div>

	<form class="displaynone" id="frmEmailDetails" name="frmEmailDetails">
		<%
			'Get the required Email information
			Dim sErrorDescription As String
			Dim sEmailInfo As String = vbNullString
			Dim iLastEventID As Integer = -1
			Dim iDetailCount As Integer = 0
			Dim eventCounter As Integer = 0
			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Try
				Dim prmSubject = New SqlParameter("psSubject", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim rsEmailDetails = objDataAccess.GetFromSP("spASRIntGetEventLogEmailInfo" _
				, New SqlParameter("psSelectedIDs", SqlDbType.VarChar, -1) With {.Value = Request("txtSelectedEventIDs")} _
				, prmSubject _
				, New SqlParameter("psOrderColumn", SqlDbType.VarChar, -1) With {.Value = CStr(Request("txtEmailOrderColumn"))} _
				, New SqlParameter("psOrderOrder", SqlDbType.VarChar, -1) With {.Value = CStr(Request("txtEmailOrderOrder"))})
			
				If rsEmailDetails.Rows.Count > 0 Then
					For Each objRow As DataRow In rsEmailDetails.Rows
						If iLastEventID <> CInt(objRow("ID")) Then
							eventCounter = eventCounter + 1
							Response.Write(CStr(eventCounter))
							sEmailInfo = sEmailInfo & StrDup(Len(objRow("Name").ToString()) + 30, "-") & vbCrLf
							sEmailInfo = sEmailInfo & "Event Name : " & objRow("Name").ToString() & vbCrLf
							sEmailInfo = sEmailInfo & StrDup(Len(objRow("Name").ToString()) + 30, "-") & vbCrLf
							sEmailInfo = sEmailInfo & "Mode :		" & objRow("Mode").ToString() & vbCrLf & vbCrLf
							sEmailInfo = sEmailInfo & "Start Time :	" & ConvertSQLDateToLocale(objRow("DateTime")) & " " & ConvertSqlDateToTime(objRow("DateTime")) & vbCrLf
							If IsDBNull(objRow("EndTime")) Then
								sEmailInfo = sEmailInfo & "End Time :	" & vbCrLf
							Else
								sEmailInfo = sEmailInfo & "End Time :	" & ConvertSQLDateToLocale(objRow("DateTime")) & " " & ConvertSqlDateToTime(objRow("EndTime")) & vbCrLf
							End If
							sEmailInfo = sEmailInfo & "Duration :	" & FormatEventDuration(CInt(objRow("Duration"))) & vbCrLf
							sEmailInfo = sEmailInfo & "Type :		" & objRow("Type").ToString() & vbCrLf
							sEmailInfo = sEmailInfo & "Status :		" & objRow("Status").ToString() & vbCrLf
							sEmailInfo = sEmailInfo & "User name :	" & objRow("Username").ToString() & vbCrLf & vbCrLf
							If Request("txtFromMain") = 0 Then
								If Request("txtBatchy") Then
									sEmailInfo = sEmailInfo & Request("txtBatchInfo") & vbCrLf
								End If
							Else
								If (Not IsDBNull(objRow("BatchName"))) And (Len(objRow("BatchName").ToString()) > 0) Then
									sEmailInfo = sEmailInfo & "Batch Job Name	: " & objRow("BatchName").ToString() & vbCrLf & vbCrLf
								End If
							End If
										
							sEmailInfo = sEmailInfo & "Records Successful :	" & objRow("SuccessCount").ToString() & vbCrLf
							sEmailInfo = sEmailInfo & "Records Failed :		" & objRow("FailCount").ToString() & vbCrLf & vbCrLf
							sEmailInfo = sEmailInfo & "Details : " & vbCrLf & vbCrLf
							iLastEventID = CInt(objRow("ID"))
							iDetailCount = 0
						End If
				
						iDetailCount += 1
				
						If objRow("count") > 0 Then
							If (Not IsDBNull(objRow("Notes"))) And (Len(objRow("Notes")) > 0) Then
								sEmailInfo = sEmailInfo & "*** Log Entry " & CStr(iDetailCount) & " of " & CStr(objRow("count")) & " ***" & vbCrLf
								sEmailInfo = sEmailInfo & objRow("Notes").ToString()
							End If
						Else
							sEmailInfo = sEmailInfo & "There are no details for this event log entry" & vbCrLf
						End If
						sEmailInfo = sEmailInfo & vbCrLf & vbCrLf & vbCrLf
					Next
					Response.Write("<input  name=txtEventDeleted id=txtEventDeleted value=0>" & vbCrLf)
				Else
					Response.Write("<input  name=txtEventDeleted id=txtEventDeleted value=1>" & vbCrLf)
				End If

				Response.Write("<input  name=txtBody id=txtBody value=""" & Replace(sEmailInfo, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input  name=txtSubject id=txtSubject value=""" & Replace(prmSubject.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
			Catch ex As Exception
				sErrorDescription = "Error getting the event log records." & vbCrLf & FormatError(ex.Message)
			End Try
		%>
	</form>

	<form class="displaynone" id="frmFromOpener" name="frmFromOpener">
		<input id="calcEmailCurrentID" name="calcEmailCurrentID" type="hidden" value='<%= Request("emailSelCurrentID") %>'>
	</form>

	<input id="txtTicker" name="txtTicker" type="hidden" value="0">
	<input id="txtLastKeyFind" name="txtLastKeyFind" type="hidden" value="">
</div>

<script type="text/javascript">
	// Table to jQuery grid
	tableToGrid("#EmailSelectionTable", {
		multiselect: true,
		ondblClickRow: function () { },
		colNames: ['EmailGroupIDHeader', 'FullNameHeader', 'To', 'Cc', 'Bcc', 'Recipient'],
		colModel: [
			{ name: 'EmailGroupIDHeader', hidden: true },
			{ name: 'FullNameHeader', sortable: false, hidden: true },
			{ name: 'to', edittype: 'checkbox', index: 'to', editoptions: { value: "True:False" }, formatter: 'checkbox', formatoptions: { disabled: false }, align: 'center', width: '10' },
			{ name: 'cc', edittype: 'checkbox', index: 'cc', editoptions: { value: "True:False" }, formatter: 'checkbox', formatoptions: { disabled: false }, align: 'center', width: '10' },
			{ name: 'bcc', edittype: 'checkbox', index: 'bcc', editoptions: { value: "True:False" }, formatter: 'checkbox', formatoptions: { disabled: false }, align: 'center', width: '10' },
			{ name: 'NameHeader', sortable: false, width: '90%' }
		],
		cmTemplate: { sortable: false },
		rowNum: 1000,
		height: ((screen.height) / 3.5) + 25,
		autowidth: true,
		beforeSelectRow: function () { return false; }
	});

	function emailSelectionEvent() {
	
		var sTo = getEmails(4);
		var SCc = getEmails(5);
		var SBcc = getEmails(6);
		var sSubject = getSubject();
		var sBody = getBody();
		
		$.ajax({
			type: "POST",
			url: "SendEmail",
			data: { 'to': sTo, 'cc': SCc, 'bcc': SBcc, 'subject': sSubject, 'body': sBody, __RequestVerificationToken: $('input[name="__RequestVerificationToken"]').val() },
			dataType: "text",
			success: function (a, b, c) {
				OpenHR.modalPrompt(c.statusText, 0, "Event Log");
				closeEmailSelect();
			},
			error: function (req, status, errorObj) {
				if (!(errorObj == "" || req.responseText == "")) {
					OpenHR.modalPrompt(errorObj, 0, "Event Log");
				}
			}
		});

		return true;
	}

	function getEmails(typeIndex) {
		var localList = "";
		$('#EmailSelectionTable').find('td:nth-child(' + typeIndex + ')').each(function () {
			$(this).find('input:checked').each(function () {
				localList += $(this).parent().siblings()[2].innerHTML + ";";
			});
		});
		return localList;
	}

	function getSubject() {
		return $('#txtSubject').val();
	}

	function getBody() {
		return $('#txtBody').val();
	}

	function closeEmailSelect() {
		$('#EventLogEmailSelect').dialog("close");
	}

	</script>

<script type="text/javascript">
	emailSelection_window_onload();
</script>


